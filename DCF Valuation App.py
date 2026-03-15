import FreeSimpleGUI as sg
from google.oauth2.service_account import Credentials
import gspread

# ── Google Sheets setup ────────────────────────────────────────────────────────

CREDENTIALS_PATH = 'credentials.json' #Google servise account credentials
SPREADSHEET_NAME = 'DCF DB' #Sheet file name (change to whatever fork for you)
SHEET_NAME = 'DB' #Sheet name

scope = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=scope)
gc = gspread.authorize(creds)


def get_worksheet():
    try:
        return gc.open(SPREADSHEET_NAME).worksheet(SHEET_NAME)
    except gspread.SpreadsheetNotFound:
        sg.popup_error(f"Spreadsheet '{SPREADSHEET_NAME}' not found!")
    except gspread.WorksheetNotFound:
        sg.popup_error(f"Sheet '{SHEET_NAME}' not found in '{SPREADSHEET_NAME}'!")
    except Exception as e:
        sg.popup_error(f"Error accessing spreadsheet: {e}")
    return None


# ── Core DCF logic ─────────────────────────────────────────────────────────────

def dcf_calculate(last_fcf, debt, cash, shares_outstanding, years,
                  growth_rate, wacc, terminal_growth_rate):
    """Returns intrinsic stock price and intermediate values."""
    if wacc <= terminal_growth_rate:
        raise ValueError("WACC must be greater than terminal growth rate.")
    if shares_outstanding <= 0:
        raise ValueError("Shares outstanding must be greater than zero.")

    projected_fcfs = [last_fcf * (1 + growth_rate) ** i for i in range(1, years + 1)]
    discounted_fcfs = [fcf / (1 + wacc) ** i for i, fcf in enumerate(projected_fcfs, 1)]
    terminal_value = projected_fcfs[-1] * (1 + terminal_growth_rate) / (wacc - terminal_growth_rate)
    discounted_tv = terminal_value / (1 + wacc) ** years
    enterprise_value = sum(discounted_fcfs) + discounted_tv
    intrinsic_value = enterprise_value - debt + cash
    intrinsic_price = intrinsic_value / shares_outstanding

    return {
        "price": intrinsic_price,
        "enterprise_value": enterprise_value,
        "intrinsic_value": intrinsic_value,
        "discounted_tv": discounted_tv,
        "sum_dcf": sum(discounted_fcfs),
    }


def margin_of_safety(intrinsic_price, market_price):
    """MoS % — positive means undervalued, negative means overvalued."""
    if intrinsic_price <= 0:
        return None
    return (intrinsic_price - market_price) / intrinsic_price * 100



def build_sensitivity_table(last_fcf, debt, cash, shares, years,
                             base_growth, base_wacc, terminal_growth):
    """
    5x5 sensitivity table: rows = WACC ± 2 steps, cols = growth rate ± 2 steps.
    Returns (row_labels, col_labels, 2D list of prices).
    """
    wacc_steps   = [base_wacc   + delta for delta in [-0.02, -0.01, 0, 0.01, 0.02]]
    growth_steps = [base_growth + delta for delta in [-0.02, -0.01, 0, 0.01, 0.02]]

    table = []
    for w in wacc_steps:
        row = []
        for g in growth_steps:
            try:
                result = dcf_calculate(last_fcf, debt, cash, shares, years,
                                       g, w, terminal_growth)
                row.append(f"${result['price']:.2f}")
            except ValueError:
                row.append("N/A")
        table.append(row)

    row_labels = [f"WACC {w*100:.1f}%" for w in wacc_steps]
    col_labels = [f"g {g*100:.1f}%" for g in growth_steps]
    return row_labels, col_labels, table


# ── Database helpers ───────────────────────────────────────────────────────────

HEADER = [
    "analysis_name", "last_fcf", "debt", "cash", "shares", "years",
    "market_price",
    "pessimistic_growth", "pessimistic_wacc", "pessimistic_terminal",
    "middle_growth",      "middle_wacc",      "middle_terminal",
    "optimistic_growth",  "optimistic_wacc",  "optimistic_terminal",
    "notes",
]


def load_database():
    worksheet = get_worksheet()
    if not worksheet:
        return [], []
    try:
        data = worksheet.get_all_values()
        if not data or len(data) < 2:
            return [], []
        header = data[0]
        database = [dict(zip(header, row)) for row in data[1:]]
        names = [a.get("analysis_name", "N/A") for a in database]
        return database, names
    except Exception as e:
        sg.popup_error(f"Error loading database: {e}")
        return [], []


def save_analysis(analysis_name, values):
    worksheet = get_worksheet()
    if not worksheet:
        return
    try:
        # Get all existing data to find the true next empty row.
        # append_row() can overwrite the last row when the sheet has
        # trailing empty rows, so we write explicitly by row index instead.
        all_rows = worksheet.get_all_values()

        if not all_rows:
            worksheet.update(values=[HEADER], range_name="A1")
            next_row = 2
        else:
            if all_rows[0] != HEADER:
                worksheet.update(values=[HEADER], range_name="A1")
            next_row = len(all_rows) + 1

        row = [
            analysis_name,
            values.get("-LAST_FCF-", ""),
            values.get("-DEBT-", ""),
            values.get("-CASH-", ""),
            values.get("-SHARES-", ""),
            values.get("-YEARS-", ""),
            values.get("-MARKET_PRICE-", ""),
            values.get("-PESSIMISTIC_GROWTH-", ""),
            values.get("-PESSIMISTIC_WACC-", ""),
            values.get("-PESSIMISTIC_TERMINAL-", ""),
            values.get("-MIDDLE_GROWTH-", ""),
            values.get("-MIDDLE_WACC-", ""),
            values.get("-MIDDLE_TERMINAL-", ""),
            values.get("-OPTIMISTIC_GROWTH-", ""),
            values.get("-OPTIMISTIC_WACC-", ""),
            values.get("-OPTIMISTIC_TERMINAL-", ""),
            values.get("-NOTES-", "").strip(),
        ]
        worksheet.add_rows(1)  # Expand grid before writing to avoid exceeding row limits
        worksheet.update(values=[row], range_name=f"A{next_row}")
        sg.popup("Analysis saved successfully!")
    except Exception as e:
        sg.popup_error(f"Error saving: {e}")


def delete_analysis(analysis_name):
    worksheet = get_worksheet()
    if not worksheet:
        return False
    try:
        cell = worksheet.find(analysis_name)
        if cell:
            worksheet.delete_rows(cell.row)
            return True
        sg.popup_error(f"Analysis '{analysis_name}' not found.")
        return False
    except Exception as e:
        sg.popup_error(f"Error deleting analysis: {e}")
        return False



def refresh_sensitivity_table(window, params, values):
    """Build sensitivity table anchored to the selected scenario and update the widget."""
    if not params:
        return
    if values.get("-SENS_BEST-"):
        scenario = "OPTIMISTIC"
        label = "Best"
    elif values.get("-SENS_BASE-"):
        scenario = "MIDDLE"
        label = "Base"
    else:
        scenario = "PESSIMISTIC"
        label = "Worst"

    base_growth, base_wacc, base_terminal = params[scenario]
    row_labels, col_labels, table = build_sensitivity_table(
        params["last_fcf"], params["debt"], params["cash"],
        params["shares"],   params["years"],
        base_growth, base_wacc, base_terminal,
    )
    table_data = [[row_labels[i]] + table[i] for i in range(len(row_labels))]

    market_price = params.get("market_price")
    row_colors = []
    if market_price:
        for i, row in enumerate(table_data):
            prices = []
            for v in row[1:]:
                try:
                    prices.append(float(v.replace("$", "")))
                except ValueError:
                    pass
            if prices and max(prices) > market_price:
                row_colors.append((i, "#d4edda"))

    window["-SENS_TABLE-"].update(
        values=table_data,
        row_colors=row_colors if market_price else [],
    )
    window["-SENS_NOTE-"].update(f"Anchored to: {label} scenario")
    if market_price:
        window["-SENS_MKT_NOTE-"].update(
            f"Market price: ${market_price:.2f}  —  green rows contain at least one undervalued estimate"
        )
    else:
        window["-SENS_MKT_NOTE-"].update("")


# ── GUI layout ─────────────────────────────────────────────────────────────────

sg.theme("Reddit")

scenario_rows = [
    [sg.Text("", size=(28, 1)),
     sg.Text("Worst", size=(8, 1), font=("Helvetica", 10, "bold"), text_color="#C0392B"),
     sg.Text("Base",  size=(8, 1), font=("Helvetica", 10, "bold"), text_color="#7F8C8D"),
     sg.Text("Best",  size=(8, 1), font=("Helvetica", 10, "bold"), text_color="#27AE60")],
    *[
        [sg.Text(label, size=(28, 1))] +
        [sg.InputText(default, key=key, size=(8, 1))
         for default, key in zip(defaults, keys)]
        for label, defaults, keys in [
            ("Growth Rate (%)",    [3, 6, 9],
             ["-PESSIMISTIC_GROWTH-", "-MIDDLE_GROWTH-", "-OPTIMISTIC_GROWTH-"]),
            ("WACC (%)",           [6.5, 7.5, 8.5],
             ["-PESSIMISTIC_WACC-", "-MIDDLE_WACC-", "-OPTIMISTIC_WACC-"]),
            ("Terminal Growth (%)",[2, 3, 4],
             ["-PESSIMISTIC_TERMINAL-", "-MIDDLE_TERMINAL-", "-OPTIMISTIC_TERMINAL-"]),
        ]
    ],
]

# Blank 5x5 sensitivity table — populated after Calculate
_blank_row = [""] * 6
SENS_HEADINGS = ["WACC / Growth", "g -2%", "g -1%", "g Base", "g +1%", "g +2%"]

left_col = [
    [sg.Text("Discounted Cash Flow Analysis", font=("Helvetica", 16, "bold"), text_color="#0079d3")],
    [sg.Text("Analysis Name:", size=(28, 1)), sg.InputText(key="-ANALYSIS_NAME-", size=(28, 1))],
    [sg.HorizontalSeparator()],

    # Parameters
    [sg.Text("Parameters", font=("Helvetica", 11, "bold"))],
    [sg.Text("Last Free Cash Flow (millions):", size=(28, 1)), sg.InputText("6770",   key="-LAST_FCF-", size=(14, 1))],
    [sg.Text("Total Debt (millions):",          size=(28, 1)), sg.InputText("11860",  key="-DEBT-",    size=(14, 1))],
    [sg.Text("Cash Equivalents (millions):",    size=(28, 1)), sg.InputText("10820",  key="-CASH-",    size=(14, 1))],
    [sg.Text("Shares Outstanding (millions):",  size=(28, 1)), sg.InputText("989.24", key="-SHARES-",  size=(14, 1))],
    [sg.Text("Years to Project:",               size=(28, 1)), sg.InputText("5",      key="-YEARS-",   size=(14, 1))],
    [sg.Text("Current Market Price:",       size=(28, 1)), sg.InputText("",       key="-MARKET_PRICE-", size=(14, 1))],
    [sg.HorizontalSeparator()],

    # Scenarios
    [sg.Text("Scenario Modeling", font=("Helvetica", 11, "bold"))],
    *scenario_rows,

    # Notes
    [sg.Text("Notes", font=("Helvetica", 11, "bold"))],
    [sg.Multiline(key="-NOTES-", size=(55, 10), no_scrollbar=False)],

    # Buttons
    [
        sg.Button("Calculate"),
        sg.Button("Save Analysis"),
    ],
]

right_col = [
    # Results
    [sg.Text("Results", font=("Helvetica", 11, "bold"))],
    [sg.Text("",          size=(12, 1)),
     sg.Text("Price",     size=(10, 1), font=("Helvetica", 9, "bold")),
     sg.Text("MoS",       size=(8, 1),  font=("Helvetica", 9, "bold")),
     sg.Text("Upside",    size=(9, 1),  font=("Helvetica", 9, "bold"))],
    [sg.Text("Worst case:", size=(12, 1)),
     sg.Text("—", key="-PESSIMISTIC_RESULT-",  size=(10, 1)),
     sg.Text("—", key="-PESSIMISTIC_MOS-",     size=(8, 1)),
     sg.Text("—", key="-PESSIMISTIC_UPSIDE-",  size=(9, 1))],
    [sg.Text("Base case:",  size=(12, 1)),
     sg.Text("—", key="-MIDDLE_RESULT-",       size=(10, 1)),
     sg.Text("—", key="-MIDDLE_MOS-",          size=(8, 1)),
     sg.Text("—", key="-MIDDLE_UPSIDE-",       size=(9, 1))],
    [sg.Text("Best case:",  size=(12, 1)),
     sg.Text("—", key="-OPTIMISTIC_RESULT-",   size=(10, 1)),
     sg.Text("—", key="-OPTIMISTIC_MOS-",      size=(8, 1)),
     sg.Text("—", key="-OPTIMISTIC_UPSIDE-",   size=(9, 1))],
    [sg.HorizontalSeparator()],

    # Sensitivity table
    [sg.Text("Sensitivity Table — Intrinsic Stock Price",
             font=("Helvetica", 11, "bold"))],
    [sg.Text("Rows: WACC ±2%  |  Columns: Growth rate ±2%  |  Anchored to selected scenario",
             font=("Helvetica", 9))],
    [sg.Text("Anchor scenario:", font=("Helvetica", 9)),
     sg.Radio("Worst", "SENS_SCENARIO", key="-SENS_WORST-",  default=True,  enable_events=True, font=("Helvetica", 9)),
     sg.Radio("Base",  "SENS_SCENARIO", key="-SENS_BASE-",   default=False, enable_events=True, font=("Helvetica", 9)),
     sg.Radio("Best",  "SENS_SCENARIO", key="-SENS_BEST-",   default=False, enable_events=True, font=("Helvetica", 9))],
    [sg.Text("Run Calculate to populate.", key="-SENS_NOTE-",
             font=("Helvetica", 9), text_color="grey")],
    [sg.Table(
        values=[_blank_row] * 5,
        headings=SENS_HEADINGS,
        key="-SENS_TABLE-",
        col_widths=[12, 9, 9, 9, 9, 9],
        auto_size_columns=False,
        justification="right",
        num_rows=5,
        hide_vertical_scroll=True,
        enable_events=False,
        font=("Courier", 10),
        header_font=("Helvetica", 9, "bold"),
        row_colors=None,
    )],
    [sg.Text("", key="-SENS_MKT_NOTE-", font=("Helvetica", 9), text_color="green")],
    [sg.HorizontalSeparator()],

    # Saved Analyses
    [sg.Text("Saved Analyses", font=("Helvetica", 11, "bold"))],
    [sg.Listbox(values=[], key="-ANALYSIS_LIST-", size=(70, 10), enable_events=True)],
    [
        sg.Button("Load Selected",   disabled=True, key="-LOAD_SELECTED-"),
        sg.Button("Delete Selected", disabled=True, key="-DELETE_SELECTED-"),
        sg.Button("Reload Database"),
    ],
]

layout = [
    [
        sg.Column(left_col,  vertical_alignment="top"),
        sg.VerticalSeparator(),
        sg.Column(right_col, vertical_alignment="top", pad=((12, 0), 0)),
    ]
]

window = sg.Window("DCF Calculation", layout, size=(980, 600), resizable=True, finalize=True)

loaded_database, analysis_names = load_database()
window["-ANALYSIS_LIST-"].update(values=analysis_names)
has_items = bool(analysis_names)
window["-LOAD_SELECTED-"].update(disabled=not has_items)
window["-DELETE_SELECTED-"].update(disabled=not has_items)

# Cache last calculated results for sensitivity table
last_params = {}

# ── Event loop ─────────────────────────────────────────────────────────────────

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break

    # ── Calculate ──────────────────────────────────────────────────────────────
    if event == "Calculate":
        try:
            def fp(key): return float(values[key])
            def ip(key): return int(float(values[key]))

            last_fcf = fp("-LAST_FCF-")
            debt     = fp("-DEBT-")
            cash     = fp("-CASH-")
            shares   = fp("-SHARES-")
            years    = ip("-YEARS-")
            mp_raw   = values["-MARKET_PRICE-"].strip()
            market_price = float(mp_raw) if mp_raw else None

            scenarios = {
                "PESSIMISTIC": ("-PESSIMISTIC_GROWTH-", "-PESSIMISTIC_WACC-", "-PESSIMISTIC_TERMINAL-"),
                "MIDDLE":      ("-MIDDLE_GROWTH-",      "-MIDDLE_WACC-",      "-MIDDLE_TERMINAL-"),
                "OPTIMISTIC":  ("-OPTIMISTIC_GROWTH-",  "-OPTIMISTIC_WACC-",  "-OPTIMISTIC_TERMINAL-"),
            }

            results = {}
            for name, (gk, wk, tk) in scenarios.items():
                results[name] = dcf_calculate(
                    last_fcf, debt, cash, shares, years,
                    fp(gk) / 100, fp(wk) / 100, fp(tk) / 100
                )

            label_map = {
                "PESSIMISTIC": ("-PESSIMISTIC_RESULT-", "-PESSIMISTIC_MOS-", "-PESSIMISTIC_UPSIDE-"),
                "MIDDLE":      ("-MIDDLE_RESULT-",      "-MIDDLE_MOS-",      "-MIDDLE_UPSIDE-"),
                "OPTIMISTIC":  ("-OPTIMISTIC_RESULT-",  "-OPTIMISTIC_MOS-",  "-OPTIMISTIC_UPSIDE-"),
            }

            for name, (res_key, mos_key, upside_key) in label_map.items():
                price = results[name]["price"]
                window[res_key].update(f"${price:.2f}")
                if market_price:
                    mos = margin_of_safety(price, market_price)
                    upside = (price - market_price) / market_price * 100
                    color = "green" if upside > 0 else "red"
                    window[mos_key].update(f"{mos:+.1f}%",    text_color=color)
                    window[upside_key].update(f"{upside:+.1f}%", text_color=color)
                else:
                    window[mos_key].update("—")
                    window[upside_key].update("—")

            # Cache params for sensitivity table (re-used when radio changes)
            last_params = {
                "last_fcf": last_fcf, "debt": debt, "cash": cash,
                "shares": shares, "years": years, "market_price": market_price,
                "PESSIMISTIC": (fp("-PESSIMISTIC_GROWTH-") / 100,
                                fp("-PESSIMISTIC_WACC-")   / 100,
                                fp("-PESSIMISTIC_TERMINAL-") / 100),
                "MIDDLE":      (fp("-MIDDLE_GROWTH-") / 100,
                                fp("-MIDDLE_WACC-")   / 100,
                                fp("-MIDDLE_TERMINAL-") / 100),
                "OPTIMISTIC":  (fp("-OPTIMISTIC_GROWTH-") / 100,
                                fp("-OPTIMISTIC_WACC-")   / 100,
                                fp("-OPTIMISTIC_TERMINAL-") / 100),
            }
            refresh_sensitivity_table(window, last_params, values)

        except ValueError as e:
            sg.popup_error(f"Input error: {e}")

    # ── Sensitivity scenario selector ─────────────────────────────────────────
    elif event in ("-SENS_WORST-", "-SENS_BASE-", "-SENS_BEST-"):
        refresh_sensitivity_table(window, last_params, values)

    # ── Save ───────────────────────────────────────────────────────────────────
    elif event == "Save Analysis":
        name = values["-ANALYSIS_NAME-"].strip()
        if not name:
            sg.popup_error("Please enter an analysis name.")
        else:
            exists = any(a.get("analysis_name") == name for a in loaded_database)
            if exists:
                new_name = sg.popup_get_text(
                    f"'{name}' already exists. Enter a new name:", title="Rename")
                if new_name and new_name != name:
                    values["-ANALYSIS_NAME-"] = new_name
                    save_analysis(new_name, values)
                    loaded_database, analysis_names = load_database()
                    window["-ANALYSIS_LIST-"].update(values=analysis_names)
            else:
                save_analysis(name, values)
                loaded_database, analysis_names = load_database()
                window["-ANALYSIS_LIST-"].update(values=analysis_names)

    # ── Reload ─────────────────────────────────────────────────────────────────
    elif event == "Reload Database":
        loaded_database, analysis_names = load_database()
        window["-ANALYSIS_LIST-"].update(values=analysis_names)
        has_items = bool(analysis_names)
        window["-LOAD_SELECTED-"].update(disabled=not has_items)
        window["-DELETE_SELECTED-"].update(disabled=not has_items)

    # ── List selection ─────────────────────────────────────────────────────────
    elif event == "-ANALYSIS_LIST-":
        selected = bool(values["-ANALYSIS_LIST-"])
        window["-LOAD_SELECTED-"].update(disabled=not selected)
        window["-DELETE_SELECTED-"].update(disabled=not selected)

    # ── Load selected ──────────────────────────────────────────────────────────
    elif event == "-LOAD_SELECTED-":
        if values["-ANALYSIS_LIST-"]:
            sel_name = values["-ANALYSIS_LIST-"][0]
            for a in loaded_database:
                if a.get("analysis_name") == sel_name:
                    field_map = {
                        "-ANALYSIS_NAME-":        "analysis_name",
                        "-LAST_FCF-":             "last_fcf",
                        "-DEBT-":                 "debt",
                        "-CASH-":                 "cash",
                        "-SHARES-":               "shares",
                        "-YEARS-":                "years",
                        "-MARKET_PRICE-":         "market_price",
                        "-PESSIMISTIC_GROWTH-":   "pessimistic_growth",
                        "-MIDDLE_GROWTH-":        "middle_growth",
                        "-OPTIMISTIC_GROWTH-":    "optimistic_growth",
                        "-PESSIMISTIC_WACC-":     "pessimistic_wacc",
                        "-MIDDLE_WACC-":          "middle_wacc",
                        "-OPTIMISTIC_WACC-":      "optimistic_wacc",
                        "-PESSIMISTIC_TERMINAL-": "pessimistic_terminal",
                        "-MIDDLE_TERMINAL-":      "middle_terminal",
                        "-OPTIMISTIC_TERMINAL-":  "optimistic_terminal",
                        "-NOTES-":                "notes",
                    }
                    for gui_key, db_key in field_map.items():
                        window[gui_key].update(a.get(db_key, ""))
                    sg.popup(f"'{sel_name}' loaded.")
                    break

    # ── Delete selected ────────────────────────────────────────────────────────
    elif event == "-DELETE_SELECTED-":
        if values["-ANALYSIS_LIST-"]:
            sel_name = values["-ANALYSIS_LIST-"][0]
            confirm = sg.popup_yes_no(
                f"Delete analysis '{sel_name}'?", title="Confirm Delete")
            if confirm == "Yes":
                if delete_analysis(sel_name):
                    loaded_database, analysis_names = load_database()
                    window["-ANALYSIS_LIST-"].update(values=analysis_names)
                    has_items = bool(analysis_names)
                    window["-LOAD_SELECTED-"].update(disabled=not has_items)
                    window["-DELETE_SELECTED-"].update(disabled=not has_items)
                    sg.popup(f"'{sel_name}' deleted.")

window.close()