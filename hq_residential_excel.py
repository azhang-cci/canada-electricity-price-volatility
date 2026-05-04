"""
Hydro-Québec Comparison of Electricity Prices
Residential Customers – Average Prices (¢/kWh)
Years: 2012–2024  |  Consumption levels: 625, 750, 1000, 2000, 3000 kWh/month
Source: Hydro-Québec annual reports (April 1 rates each year)

2025 data is Monthly Bills (C$), not average prices — kept on a separate sheet.
2014 image showed Monthly Bills (C$), not average prices — also on separate sheet.

Run:  pip install openpyxl pandas
      python this_script.py
Output: HydroQuebec_Residential.xlsx
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── DATA ──────────────────────────────────────────────────────────────────────
CITIES = [
    "Montréal, QC",
    "Calgary, AB",
    "Charlottetown, PE",
    "Edmonton, AB",
    "Halifax, NS",
    "Moncton, NB",
    "Ottawa, ON",
    "Regina, SK",
    "St. John's, NL",
    "Toronto, ON",
    "Vancouver, BC",
    "Winnipeg, MB",
]

CONS = [625, 750, 1000, 2000, 3000]

# Average Prices (¢/kWh) — keyed by year, then city, then [625,750,1000,2000,3000]
avg_prices = {
    2012: {
        "Montréal, QC":       [7.27, 6.95, 6.76, 7.13, 7.26],
        "Calgary, AB":        [15.01,14.51,13.89,12.95,12.64],
        "Charlottetown, PE":  [15.98,15.33,14.51,13.28,11.92],
        "Edmonton, AB":       [14.21,13.63,12.90,11.82,11.46],
        "Halifax, NS":        [15.66,15.37,15.01,14.46,14.28],
        "Moncton, NB":        [13.01,12.48,11.82,10.84,10.51],
        "Ottawa, ON":         [13.40,13.16,13.14,13.17,13.18],
        "Regina, SK":         [13.69,13.18,12.54,11.57,11.25],
        "St. John's, NL":     [12.73,12.31,11.80,11.02,10.77],
        "Toronto, ON":        [14.51,14.08,13.57,13.07,12.91],
        "Vancouver, BC":      [7.91, 8.14, 8.78, 9.74,10.06],
        "Winnipeg, MB":       [7.87, 7.68, 7.46, 7.11, 7.00],
    },
    2013: {
        "Montréal, QC":       [7.36, 7.04, 6.87, 7.32, 7.48],
        "Calgary, AB":        [15.93,15.44,14.81,13.88,13.56],
        "Charlottetown, PE":  [16.34,15.69,14.87,13.64,12.23],
        "Edmonton, AB":       [15.14,14.59,13.90,12.87,12.52],
        "Halifax, NS":        [16.10,15.81,15.45,14.90,14.72],
        "Moncton, NB":        [13.01,12.48,11.82,10.84,10.51],
        "Ottawa, ON":         [12.90,12.67,12.39,11.97,11.83],
        "Regina, SK":         [14.37,13.83,13.15,12.14,11.80],
        "St. John's, NL":     [13.47,13.06,12.55,11.78,11.52],
        "Toronto, ON":        [13.28,12.91,12.48,12.08,11.94],
        "Vancouver, BC":      [8.03, 8.26, 8.91, 9.88,10.21],
        "Winnipeg, MB":       [8.04, 7.85, 7.63, 7.28, 7.17],
    },
    # 2014: only monthly bills image provided — see separate sheet
    2015: {
        "Montréal, QC":       [7.63, 7.31, 7.19, 7.90, 8.13],
        "Calgary, AB":        [12.94,12.37,11.66,10.58,10.23],
        "Charlottetown, PE":  [17.09,16.44,15.62,14.39,13.05],
        "Edmonton, AB":       [12.91,12.30,11.55,10.41,10.04],
        "Halifax, NS":        [16.68,16.39,16.03,15.49,15.31],
        "Moncton, NB":        [13.53,12.98,12.30,11.27,10.93],
        "Ottawa, ON":         [14.78,14.52,14.20,13.72,13.55],
        "Regina, SK":         [15.59,15.05,14.37,13.36,13.02],
        "St. John's, NL":     [12.41,12.03,11.55,10.84,10.60],
        "Toronto, ON":        [15.06,14.71,14.31,13.99,13.89],
        "Vancouver, BC":      [9.27, 9.54,10.29,11.42,11.80],
        "Winnipeg, MB":       [8.55, 8.35, 8.11, 7.75, 7.62],
    },
    2016: {
        "Montréal, QC":       [7.66, 7.34, 7.23, 7.95, 8.20],
        "Calgary, AB":        [11.68,11.11,10.40, 9.33, 8.97],
        "Charlottetown, PE":  [17.49,16.84,16.02,14.79,13.46],
        "Edmonton, AB":       [11.70,11.11,10.37, 9.26, 8.89],
        "Halifax, NS":        [16.53,16.24,15.88,15.34,15.16],
        "Moncton, NB":        [13.75,13.19,12.50,11.46,11.11],
        "Ottawa, ON":         [17.01,16.63,16.15,15.44,15.20],
        "Regina, SK":         [15.86,15.32,14.65,13.63,13.30],
        "St. John's, NL":     [12.89,12.48,11.96,11.19,10.93],
        "Toronto, ON":        [19.23,18.60,17.81,16.62,16.23],
        "Vancouver, BC":      [9.64, 9.92,10.70,11.88,12.27],
        "Winnipeg, MB":       [8.88, 8.68, 8.43, 8.05, 7.92],
    },
    2017: {
        "Montréal, QC":       [7.77, 7.45, 7.07, 8.00, 8.30],
        "Calgary, AB":        [11.65,11.12,10.45, 9.45, 9.12],
        "Charlottetown, PE":  [17.89,17.24,16.42,15.19,13.82],
        "Edmonton, AB":       [11.73,11.11,10.34, 9.18, 8.80],
        "Halifax, NS":        [16.80,16.51,16.15,15.60,15.42],
        "Moncton, NB":        [14.27,13.69,12.97,11.89,11.53],
        "Ottawa, ON":         [16.19,15.75,15.21,14.40,14.13],
        "Regina, SK":         [17.26,16.67,15.94,14.84,14.47],
        "St. John's, NL":     [12.10,11.68,11.15,10.36,10.10],
        "Toronto, ON":        [17.91,17.20,16.32,14.99,14.55],
        "Vancouver, BC":      [9.98,10.27,11.08,12.30,12.70],
        "Winnipeg, MB":       [9.18, 8.97, 8.71, 8.32, 8.19],
    },
    2018: {
        "Montréal, QC":       [7.86, 7.54, 7.13, 8.00, 8.37],
        "Calgary, AB":        [17.30,16.63,15.79,14.53,14.10],
        "Charlottetown, PE":  [18.30,17.65,16.83,15.60,14.21],
        "Edmonton, AB":       [15.79,15.15,14.35,13.15,12.75],
        "Halifax, NS":        [17.06,16.77,16.41,15.87,15.69],
        "Moncton, NB":        [14.27,13.69,12.97,11.89,11.53],
        "Ottawa, ON":         [13.33,12.81,12.16,11.18,10.85],
        "Regina, SK":         [17.87,17.27,16.51,15.37,14.99],
        "St. John's, NL":     [12.98,12.56,12.03,11.24,10.97],
        "Toronto, ON":        [15.09,14.26,13.24,11.69,11.18],
        "Vancouver, BC":      [10.28,10.58,11.42,12.67,13.09],
        "Winnipeg, MB":       [9.49, 9.27, 9.00, 8.60, 8.47],
    },
    2019: {
        "Montréal, QC":       [8.03, 7.71, 7.30, 8.01, 8.47],
        "Calgary, AB":        [17.20,16.55,15.74,14.53,14.13],
        "Charlottetown, PE":  [18.30,17.65,16.83,15.60,14.21],
        "Edmonton, AB":       [16.18,15.51,14.68,13.44,13.02],
        "Halifax, NS":        [17.34,17.05,16.69,16.14,15.96],
        "Moncton, NB":        [14.41,13.82,13.10,12.00,11.64],
        "Ottawa, ON":         [13.43,12.81,12.04,10.88,10.50],
        "Regina, SK":         [17.87,17.27,16.51,15.37,14.99],
        "St. John's, NL":     [13.75,13.33,12.80,12.01,11.75],
        "Toronto, ON":        [16.12,15.13,13.89,12.04,11.42],
        "Vancouver, BC":      [10.47,10.77,11.62,12.90,13.32],
        "Winnipeg, MB":       [9.87, 9.65, 9.37, 8.95, 8.81],
    },
    2020: {
        "Montréal, QC":       [8.03, 7.71, 7.30, 8.01, 8.47],
        "Calgary, AB":        [16.28,15.64,14.83,13.62,13.22],
        "Charlottetown, PE":  [18.30,17.65,16.83,15.60,14.21],
        "Edmonton, AB":       [15.89,15.18,14.29,12.96,12.52],
        "Halifax, NS":        [17.54,17.25,16.89,16.35,16.17],
        "Moncton, NB":        [14.76,14.17,13.42,12.30,11.93],
        "Ottawa, ON":         [11.49,10.96,10.29, 9.28, 8.95],
        "Regina, SK":         [17.87,17.27,16.51,15.37,14.99],
        "St. John's, NL":     [14.55,14.13,13.60,12.81,12.54],
        "Toronto, ON":        [12.63,11.95,11.10, 9.82, 9.39],
        "Vancouver, BC":      [10.38,10.67,11.51,12.77,13.19],
        "Winnipeg, MB":       [10.12, 9.89, 9.60, 9.17, 9.03],
    },
    2021: {
        "Montréal, QC":       [8.13, 7.81, 7.39, 8.11, 8.58],
        "Calgary, AB":        [18.88,18.16,17.26,15.90,15.45],
        "Charlottetown, PE":  [18.85,18.20,17.38,16.15,14.73],
        "Edmonton, AB":       [18.63,17.91,16.99,15.62,15.16],
        "Halifax, NS":        [17.74,17.45,17.09,16.55,16.37],
        "Moncton, NB":        [15.03,14.42,13.66,12.52,12.14],
        "Ottawa, ON":         [13.87,13.24,12.45,11.28,10.88],
        "Regina, SK":         [17.87,17.27,16.51,15.37,14.99],
        "St. John's, NL":     [14.55,14.13,13.60,12.81,12.54],
        "Toronto, ON":        [15.25,14.44,13.43,11.91,11.41],
        "Vancouver, BC":      [10.44,10.74,11.58,12.84,13.26],
        "Winnipeg, MB":       [10.40,10.16, 9.87, 9.43, 9.28],
    },
    2022: {
        "Montréal, QC":       [8.35, 8.01, 7.59, 8.32, 8.80],
        "Calgary, AB":        [21.63,20.88,19.94,18.53,18.06],
        "Charlottetown, PE":  [19.25,18.60,17.78,16.55,15.13],
        "Edmonton, AB":       [21.16,20.42,19.48,18.07,17.61],
        "Halifax, NS":        [17.95,17.66,17.30,16.76,16.58],
        "Moncton, NB":        [15.33,14.71,13.94,12.77,12.39],
        "Ottawa, ON":         [14.50,13.81,12.94,11.64,11.21],
        "Regina, SK":         [17.87,17.27,16.51,15.37,14.99],
        "St. John's, NL":     [14.70,14.28,13.76,12.98,12.72],
        "Toronto, ON":        [15.87,14.99,13.88,12.23,11.68],
        "Vancouver, BC":      [10.31,10.59,11.39,12.60,13.00],
        "Winnipeg, MB":       [10.80,10.55,10.24, 9.78, 9.63],
    },
    2023: {
        "Montréal, QC":       [8.60, 8.25, 7.81, 8.57, 9.06],
        "Calgary, AB":        [31.55,30.78,29.80,28.34,27.86],
        "Charlottetown, PE":  [19.25,18.60,17.78,16.55,15.13],
        "Edmonton, AB":       [29.36,28.66,27.78,26.47,26.03],
        "Halifax, NS":        [19.42,18.91,18.27,17.31,16.99],
        "Moncton, NB":        [16.08,15.43,14.61,13.38,12.97],
        "Ottawa, ON":         [15.23,14.45,13.48,12.02,11.53],
        "Regina, SK":         [19.69,18.89,17.89,16.39,15.89],
        "St. John's, NL":     [14.67,14.25,13.73,12.94,12.68],
        "Toronto, ON":        [16.03,15.07,13.88,12.08,11.49],
        "Vancouver, BC":      [10.51,10.80,11.62,12.85,13.26],
        "Winnipeg, MB":       [10.80,10.55,10.24, 9.78, 9.63],
    },
    2025: {
        "Montréal, QC":       [9.12, 8.75, 8.29, 9.10, 9.61],
        "Calgary, AB":        [25.05,24.10,22.90,21.11,20.51],
        "Charlottetown, PE":  [21.16,20.51,19.69,18.46,16.89],
        "Edmonton, AB":       [23.67,22.93,22.00,20.62,20.15],
        "Halifax, NS":        [21.63,21.12,20.48,19.52,19.20],
        "Moncton, NB":        [18.38,17.59,16.61,15.13,14.64],
        "Ottawa, ON":         [16.18,15.37,14.35,12.82,12.33],
        "Regina, SK":         [19.69,18.89,17.89,16.39,15.89],
        "St. John's, NL":     [16.52,16.11,15.59,14.80,14.54],
        "Toronto, ON":        [17.98,16.91,15.57,13.56,12.89],
        "Vancouver, BC":      [12.28,12.32,12.60,13.02,13.17],
        "Winnipeg, MB":       [11.10,10.85,10.53,10.06, 9.90],
    },
    2024: {
        "Montréal, QC":       [8.85, 8.50, 8.05, 8.83, 9.34],
        "Calgary, AB":        [25.28,24.34,23.17,21.41,20.82],
        "Charlottetown, PE":  [20.60,19.95,19.13,17.90,16.37],
        "Edmonton, AB":       [25.68,24.92,23.98,22.58,22.11],
        "Halifax, NS":        [20.61,20.10,19.46,18.51,18.19],
        "Moncton, NB":        [17.99,17.24,16.30,14.88,14.41],
        "Ottawa, ON":         [15.97,15.22,14.28,12.87,12.22],
        "Regina, SK":         [19.69,18.89,17.89,16.39,15.89],
        "St. John's, NL":     [15.56,15.14,14.62,13.84,13.58],
        "Toronto, ON":        [17.15,16.22,15.06,13.31,12.72],
        "Vancouver, BC":      [11.49,11.61,12.06,12.73,12.96],
        "Winnipeg, MB":       [11.10,10.85,10.53,10.06, 9.90],
    },
}

# 2014 & 2025 monthly bills (C$) — separate sheet
monthly_bills = {
    2014: {
        "Montréal, QC":       [47.00,  53.97,  70.58, 153.18, 235.78],
        "Calgary, AB":        [92.39, 106.28, 134.06, 245.19, 356.32],
        "Charlottetown, PE":  [104.45,120.42, 152.37, 280.17, 378.67],
        "Edmonton, AB":       [82.05,  94.30, 118.79, 216.78, 314.77],
        "Halifax, NS":        [104.25,122.93, 160.30, 309.77, 459.24],
        "Moncton, NB":        [82.89,  95.46, 120.58, 221.08, 321.58],
        "Ottawa, ON":         [87.63, 103.25, 134.49, 259.45, 384.41],
        "Regina, SK":         [94.79, 109.70, 139.53, 258.84, 378.15],
        "St. John's, NL":     [76.20,  88.60, 113.39, 212.58, 311.76],
        "Toronto, ON":        [93.11, 108.02, 137.84, 257.10, 376.37],
        "Vancouver, BC":      [54.66,  67.49,  97.07, 215.41, 333.74],
        "Winnipeg, MB":       [51.98,  60.96,  78.92, 150.75, 222.58],
    },
    2025: {
        "Montréal, QC":       [57.00,  65.63,  82.90, 181.92, 288.44],
        "Calgary, AB":        [156.59,180.74, 229.04, 422.23, 615.42],
        "Charlottetown, PE":  [132.26,153.80, 196.87, 369.17, 506.67],
        "Edmonton, AB":       [147.91,171.95, 220.02, 412.33, 604.64],
        "Halifax, NS":        [135.18,158.38, 204.78, 390.39, 576.00],
        "Moncton, NB":        [114.88,131.95, 166.08, 302.61, 439.14],
        "Ottawa, ON":         [101.15,115.27, 143.51, 256.44, 369.91],
        "Regina, SK":         [123.08,141.70, 178.94, 327.89, 476.84],
        "St. John's, NL":     [103.28,120.81, 155.86, 296.08, 436.30],
        "Toronto, ON":        [112.36,126.80, 155.69, 271.26, 386.82],
        "Vancouver, BC":      [76.72,  92.40, 126.02, 260.48, 394.95],
        "Winnipeg, MB":       [69.38,  81.36, 105.33, 201.20, 297.07],
    },
}

# ── BUILD DATAFRAMES ───────────────────────────────────────────────────────────
def build_panel(data_dict, value_label):
    rows = []
    for year, city_data in sorted(data_dict.items()):
        for city in CITIES:
            vals = city_data.get(city, [None]*5)
            for i, kwh in enumerate(CONS):
                rows.append({
                    "Year": year,
                    "City": city,
                    "Consumption_kWh": kwh,
                    value_label: vals[i],
                })
    return pd.DataFrame(rows)

def build_wide(data_dict):
    """City × consumption columns, years stacked."""
    rows = []
    for year, city_data in sorted(data_dict.items()):
        for city in CITIES:
            vals = city_data.get(city, [None]*5)
            row = {"Year": year, "City": city}
            for i, kwh in enumerate(CONS):
                row[f"{kwh} kWh"] = vals[i]
            rows.append(row)
    return pd.DataFrame(rows)

# ── WRITE EXCEL ────────────────────────────────────────────────────────────────
fname = "HydroQuebec_Residential.xlsx"

with pd.ExcelWriter(fname, engine="openpyxl") as writer:

    # Sheet 1: Panel (long format) — avg prices
    df_panel = build_panel(avg_prices, "Avg_Price_cents_per_kWh")
    df_panel.to_excel(writer, sheet_name="Panel_AvgPrice", index=False)

    # Sheet 2: Wide format — avg prices (year × city, one col per consumption)
    df_wide = build_wide(avg_prices)
    df_wide.to_excel(writer, sheet_name="Wide_AvgPrice", index=False)

    # Sheets 3-15: one sheet per year, avg prices
    for year, city_data in sorted(avg_prices.items()):
        rows = [{"City": c, **{f"{k} kWh": city_data.get(c,[None]*5)[i]
                               for i,k in enumerate(CONS)}}
                for c in CITIES]
        pd.DataFrame(rows).to_excel(writer, sheet_name=str(year), index=False)

    # Sheet: Monthly Bills (2014 only — no avg price image available)
    df_bills = build_wide({2014: monthly_bills[2014]})
    df_bills.to_excel(writer, sheet_name="MonthlyBills_2014_only", index=False)

# ── STYLE ─────────────────────────────────────────────────────────────────────
wb = load_workbook(fname)

HEADER_FILL  = PatternFill("solid", fgColor="003366")   # dark blue
HEADER_FONT  = Font(color="FFFFFF", bold=True)
MTL_FILL     = PatternFill("solid", fgColor="D9E1F2")   # light blue for Montréal
ALT_FILL     = PatternFill("solid", fgColor="F2F2F2")
BORDER = Border(
    bottom=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin",  color="CCCCCC"),
)

def style_sheet(ws):
    # Header row
    for cell in ws[1]:
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    # Data rows
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        # Check if this is Montréal row
        is_mtl = any("Montréal" in str(c.value) for c in row)
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center")
            if is_mtl:
                cell.fill = MTL_FILL
                cell.font = Font(bold=True)
            elif row_idx % 2 == 0:
                cell.fill = ALT_FILL
    # Auto-width
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 28)

for sh in wb.sheetnames:
    style_sheet(wb[sh])

# Add notes to bills sheet
bills_ws = wb["MonthlyBills_2014_only"]
bills_ws.insert_rows(1)
bills_ws["A1"] = "NOTE: Values are Monthly Bills in C$ (not ¢/kWh). No avg-price image was available for 2014."
bills_ws["A1"].font = Font(italic=True, color="CC0000")
bills_ws.merge_cells("A1:G1")

wb.save(fname)
print(f"✓ Saved {fname}")
print(f"  Sheets: {wb.sheetnames}")
