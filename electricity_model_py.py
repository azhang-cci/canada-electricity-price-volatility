"""
Canadian Residential Electricity Rate Model
============================================
Panel OLS with province + year fixed effects (within-group demeaning).
Outputs: regression results, driver decomposition, and all inputs → Excel.

Dependencies:
    pip install pandas numpy openpyxl scipy

Run:
    python electricity_model.py
    → produces: electricity_rate_model_outputs.xlsx
"""

import numpy as np
import pandas as pd
from scipy import stats
from openpyxl import load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers as xl_numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
import warnings
warnings.filterwarnings("ignore")


# ── 1. RAW DATA ───────────────────────────────────────────────────────────────

YEARS = [2012, 2013, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025]

# Residential average price ¢/kWh at 1,000 kWh/month
# Source: Hydro-Québec Comparison of Electricity Prices in Major North American Cities (2012–2025)
# Rates effective April 1 each year; exclude taxes and non-utility levies
RATES = {
    "Alberta":     [13.89,14.81,11.66,10.40,10.45,15.79,15.74,14.83,17.26,19.94,29.80,23.17,22.90],
    "BC":          [ 8.78, 8.91,10.29,10.70,11.08,11.42,11.62,11.51,11.58,11.39,11.62,12.06,12.60],
    "Manitoba":    [ 7.46, 7.63, 8.11, 8.43, 8.71, 9.00, 9.37, 9.60, 9.87,10.24,10.24,10.53,10.53],
    "NB":          [11.82,11.82,12.30,12.50,12.97,12.97,13.10,13.42,13.66,13.94,14.61,16.30,16.61],
    "NL":          [11.80,12.55,11.55,11.96,11.15,12.03,12.80,13.60,13.60,13.76,13.73,14.62,15.59],
    "NS":          [15.01,15.45,16.03,15.88,16.15,16.41,16.69,16.89,17.09,17.30,18.27,19.46,20.48],
    "Ontario":     [13.36,12.44,14.26,16.48,15.77,12.70,12.97,10.70,12.94,13.41,13.68,14.67,14.35],
    "PEI":         [14.51,14.87,15.62,16.02,16.42,16.83,16.83,16.83,17.38,17.78,17.78,19.13,19.69],
    "Quebec":      [ 6.76, 6.87, 7.19, 7.23, 7.07, 7.13, 7.30, 7.30, 7.39, 7.59, 7.81, 8.05, 8.29],
    "Saskatchewan":[12.54,13.15,14.37,14.65,15.94,16.51,16.51,16.51,16.51,16.51,17.89,17.89,17.89],
}

# Consumed fossil share (generation mix adjusted for net electricity trade)
# Source: CER Provincial Energy Profiles; StatCan 25-10-0016-01
# PEI: ~68% imports from NB; BC: ~10 TWh imports from US western grid (~60% fossil)
FOSSIL = {
    "Alberta":     [0.89,0.89,0.84,0.82,0.79,0.76,0.71,0.69,0.66,0.61,0.57,0.57,0.56],
    "BC":          [0.20,0.20,0.19,0.18,0.17,0.17,0.16,0.15,0.14,0.13,0.13,0.12,0.12],
    "Manitoba":    [0.03,0.03,0.02,0.02,0.02,0.02,0.01,0.02,0.03,0.02,0.02,0.02,0.02],
    "NB":          [0.33,0.30,0.30,0.30,0.27,0.27,0.28,0.28,0.28,0.36,0.34,0.33,0.32],
    "NL":          [0.07,0.07,0.07,0.07,0.07,0.06,0.06,0.06,0.04,0.03,0.03,0.03,0.03],
    "NS":          [0.66,0.66,0.64,0.63,0.62,0.62,0.61,0.60,0.60,0.61,0.60,0.59,0.58],
    "Ontario":     [0.17,0.11,0.08,0.08,0.08,0.09,0.09,0.09,0.09,0.10,0.14,0.13,0.12],
    "PEI":         [0.19,0.19,0.20,0.21,0.21,0.21,0.21,0.21,0.20,0.21,0.21,0.21,0.21],
    "Quebec":      [0.01,0.01,0.01,0.01,0.01,0.01,0.01,0.01,0.01,0.00,0.01,0.01,0.01],
    "Saskatchewan":[0.77,0.76,0.73,0.73,0.73,0.73,0.72,0.72,0.72,0.73,0.74,0.74,0.74],
}

# Market structure: Dereg / Hybrid / Reg
MARKET_STRUCTURE = {
    "Alberta":      "Deregulated",
    "BC":           "Regulated",
    "Manitoba":     "Regulated",
    "NB":           "Hybrid",
    "NL":           "Regulated",
    "NS":           "Regulated",
    "Ontario":      "Hybrid",
    "PEI":          "Regulated",
    "Quebec":       "Regulated",
    "Saskatchewan": "Regulated",
}

# Gas price benchmark per province
# AECO: western Canada (AB, BC, MB, SK)
# HH:   Henry Hub → Atlantic (NB, NS, NL, PEI)
# MIX:  average of AECO and HH (Ontario — Dawn Hub tracks HH)
# NONE: Quebec — hydro, no fuel cost exposure
GAS_BENCHMARK = {
    "Alberta":      "AECO",
    "BC":           "AECO",
    "Manitoba":     "AECO",
    "NB":           "HH",
    "NL":           "HH",
    "NS":           "HH",
    "Ontario":      "MIX",
    "PEI":          "HH",
    "Quebec":       "NONE",
    "Saskatchewan": "AECO",
}

# AECO annual average CAD/GJ
# Source: Alberta Energy & Minerals monthly reference price archive
# 2025: average of Jan–Dec 2025 monthly values (1.62+1.75+1.64+1.79+1.65+1.05+0.90+0.61+0.50+0.95+2.04+2.71)/12
AECO = {
    2012: (2.56+2.13+1.83+1.63+1.58+1.75+1.92+2.05+1.94+2.36+2.90+2.98)/12,
    2013: (2.76+2.69+2.85+3.14+3.24+3.17+2.78+2.37+2.14+2.49+3.06+3.22)/12,
    2015: (2.83+2.51+2.53+2.30+2.33+2.36+2.38+2.56+2.55+2.40+2.21+2.12)/12,
    2016: (2.11+1.82+1.36+1.11+0.94+1.23+1.82+1.82+2.12+2.41+2.48+2.75)/12,
    2017: (2.86+2.39+2.20+2.34+2.46+2.39+1.83+1.72+1.20+1.11+1.92+1.82)/12,
    2018: (1.74+1.76+1.54+1.26+0.78+0.75+1.14+0.94+1.03+1.21+1.59+1.69)/12,
    2019: (1.55+2.10+1.99+0.91+1.22+0.55+0.87+0.82+0.76+1.63+2.19+2.22)/12,
    2020: (2.06+1.79+1.60+1.56+1.66+1.65+1.62+1.85+2.00+1.99+2.58+2.41)/12,
    2021: (2.32+3.00+2.54+2.33+2.56+2.78+3.17+2.78+3.15+4.01+4.57+3.99)/12,
    2022: (3.88+4.26+4.26+5.22+5.97+6.53+5.44+3.55+4.00+3.53+5.12+5.65)/12,
    2023: (4.55+3.24+2.72+2.24+2.00+1.94+1.93+2.24+2.25+2.07+2.30+2.04)/12,
    2024: (2.63+1.73+1.48+1.25+1.01+0.78+0.64+0.53+0.43+0.66+1.33+1.62)/12,
    2025: (1.62+1.75+1.64+1.79+1.65+1.05+0.90+0.61+0.50+0.95+2.04+2.71)/12,
}

# Henry Hub annual average USD/MMBtu → converted to CAD/GJ
# Source: EIA AHHNGSP series; Bank of Canada annual average CAD/USD exchange rates
# Conversion: CAD/GJ = USD/MMBtu × FX_CAD_per_USD / 1.05506 (GJ per MMBtu)
HH_USD = {
    2012: 2.75, 2013: 3.73, 2015: 2.62, 2016: 2.62, 2017: 3.02,
    2018: 3.15, 2019: 2.57, 2020: 2.03, 2021: 3.89, 2022: 6.45,
    2023: 2.57, 2024: 2.21, 2025: 3.52,
}

FX_CAD_PER_USD = {
    2012: 0.9996, 2013: 1.0299, 2015: 1.2787, 2016: 1.3248, 2017: 1.2986,
    2018: 1.2957, 2019: 1.3269, 2020: 1.3415, 2021: 1.2535, 2022: 1.3013,
    2023: 1.3497, 2024: 1.3601, 2025: 1.3800,
}

HH_CAD_GJ = {y: HH_USD[y] * FX_CAD_PER_USD[y] / 1.05506 for y in HH_USD}


# ── 2. BUILD PANEL DATAFRAME ──────────────────────────────────────────────────

def gas_price(prov: str, year: int) -> float:
    """Return the relevant gas benchmark price (CAD/GJ) for a province-year."""
    bench = GAS_BENCHMARK[prov]
    if bench == "AECO": return AECO.get(year, np.nan)
    if bench == "HH":   return HH_CAD_GJ.get(year, np.nan)
    if bench == "MIX":  return (AECO.get(year, np.nan) + HH_CAD_GJ.get(year, np.nan)) / 2
    return 0.0  # Quebec — no fuel cost exposure


rows = []
for prov in RATES:
    for i, year in enumerate(YEARS):
        rows.append({
            "Province":       prov,
            "Year":           year,
            "Rate_c_kWh":     RATES[prov][i],
            "FossilShare":    FOSSIL[prov][i],
            "GasPrice_CAD_GJ": gas_price(prov, year),
            "MarketStructure": MARKET_STRUCTURE[prov],
            "GasBenchmark":   GAS_BENCHMARK[prov],
            "Dereg":          1 if MARKET_STRUCTURE[prov] == "Deregulated" else 0,
            "Hybrid":         1 if MARKET_STRUCTURE[prov] == "Hybrid" else 0,
            "AECO_CAD_GJ":    AECO.get(year, np.nan),
            "HH_CAD_GJ":      HH_CAD_GJ.get(year, np.nan),
            "FX_CAD_USD":     FX_CAD_PER_USD.get(year, np.nan),
        })

df = pd.DataFrame(rows)


# ── 3. WITHIN-GROUP DEMEANING (province + year fixed effects) ─────────────────

def demean(df: pd.DataFrame) -> pd.DataFrame:
    """
    Two-way within-transformation absorbing province and year fixed effects.
    For each variable v: v_w = v - mean_prov(v) - mean_year(v) + grand_mean(v)
    """
    d = df.copy()
    for col in ["Rate_c_kWh", "FossilShare", "GasPrice_CAD_GJ"]:
        gm = d[col].mean()
        pm = d.groupby("Province")[col].transform("mean")
        ym = d.groupby("Year")[col].transform("mean")
        d[col + "_w"] = d[col] - pm - ym + gm
    # Store province and year means for decomposition
    d["ProvMean_Rate"] = d.groupby("Province")["Rate_c_kWh"].transform("mean")
    d["YearMean_Rate"] = d.groupby("Year")["Rate_c_kWh"].transform("mean")
    d["GrandMean_Rate"] = d["Rate_c_kWh"].mean()
    d["ProvMean_Fossil"] = d.groupby("Province")["FossilShare"].transform("mean")
    d["YearMean_Fossil"] = d.groupby("Year")["FossilShare"].transform("mean")
    d["GrandMean_Fossil"] = d["FossilShare"].mean()
    d["ProvMean_Gas"] = d.groupby("Province")["GasPrice_CAD_GJ"].transform("mean")
    d["YearMean_Gas"] = d.groupby("Year")["GasPrice_CAD_GJ"].transform("mean")
    d["GrandMean_Gas"] = d["GasPrice_CAD_GJ"].mean()
    return d

df = demean(df)


# ── 4. OLS REGRESSION ─────────────────────────────────────────────────────────

def run_ols(df: pd.DataFrame) -> dict:
    """
    Panel OLS on within-transformed variables.

    Model:
        rate_w ~ β1·fossil_w + β2·(fossil_w × Dereg) + β3·(fossil_w × Hybrid)
               + β4·Dereg + β5·Hybrid
               + β6·gas_w + β7·(gas_w × FossilShare) + β8·(gas_w × Dereg)
               + ε

    Where _w denotes within-transformed (province + year FE absorbed).
    Coefficients in ¢/kWh per unit change in respective regressor.
    """
    d = df.dropna(subset=["Rate_c_kWh_w", "FossilShare_w", "GasPrice_CAD_GJ_w"])

    y = d["Rate_c_kWh_w"].values
    fW = d["FossilShare_w"].values
    gW = d["GasPrice_CAD_GJ_w"].values
    f  = d["FossilShare"].values
    D  = d["Dereg"].values
    H  = d["Hybrid"].values

    X = np.column_stack([
        fW,          # β1: fossil share (baseline — regulated province)
        fW * D,      # β2: fossil × deregulated (amplification via mkt structure)
        fW * H,      # β3: fossil × hybrid
        D,           # β4: deregulated fixed effect (level difference)
        H,           # β5: hybrid fixed effect
        gW,          # β6: gas price (baseline pass-through)
        gW * f,      # β7: gas × fossil share (fuel mix amplification)
        gW * D,      # β8: gas × deregulated (market structure amplification)
    ])

    varnames = [
        "Fossil share",
        "Fossil × Deregulated",
        "Fossil × Hybrid",
        "Deregulated (FE)",
        "Hybrid (FE)",
        "Gas price (CAD/GJ)",
        "Gas price × Fossil share",
        "Gas price × Deregulated",
    ]

    n, k = X.shape
    beta, ssr, _, _ = np.linalg.lstsq(X, y, rcond=None)
    yhat  = X @ beta
    resid = y - yhat
    sse   = resid @ resid
    sst   = ((y - y.mean()) ** 2).sum()
    r2    = 1 - sse / sst
    adj_r2 = 1 - (1 - r2) * (n - 1) / (n - k - 1)
    rmse  = np.sqrt(sse / n)

    # Heteroskedasticity-robust (HC1) variance-covariance matrix
    # V = (X'X)^{-1} · X' diag(e²·n/(n-k)) X · (X'X)^{-1}
    XtX_inv = np.linalg.inv(X.T @ X)
    meat = sum(((e**2) * np.outer(xi, xi)) for xi, e in zip(X, resid)) * n / (n - k)
    vcov = XtX_inv @ meat @ XtX_inv
    se_robust = np.sqrt(np.diag(vcov))

    t_stats = beta / se_robust
    p_vals  = 2 * stats.t.sf(np.abs(t_stats), df=n - k)
    ci_lo   = beta - 1.96 * se_robust
    ci_hi   = beta + 1.96 * se_robust

    def stars(p):
        if p < 0.01: return "***"
        if p < 0.05: return "**"
        if p < 0.10: return "*"
        return ""

    coef_df = pd.DataFrame({
        "Variable":    varnames,
        "Coefficient": beta,
        "Robust_SE":   se_robust,
        "t_stat":      t_stats,
        "p_value":     p_vals,
        "CI_lower_95": ci_lo,
        "CI_upper_95": ci_hi,
        "Significance": [stars(p) for p in p_vals],
    })

    return {
        "coef_df":  coef_df,
        "beta":     beta,
        "yhat":     yhat,
        "resid":    resid,
        "r2":       r2,
        "adj_r2":   adj_r2,
        "rmse":     rmse,
        "n":        n,
        "k":        k,
        "vcov":     vcov,
        "df_used":  d.reset_index(drop=True),
    }

results = run_ols(df)
b = results["beta"]
print(f"\n{'='*60}")
print("REGRESSION RESULTS")
print(f"{'='*60}")
print(f"R² (within):  {results['r2']:.4f}")
print(f"Adj. R²:      {results['adj_r2']:.4f}")
print(f"RMSE:         {results['rmse']:.3f}¢/kWh")
print(f"Observations: {results['n']}")
print()
print(results["coef_df"].to_string(index=False))


# ── 5. DRIVER DECOMPOSITION ───────────────────────────────────────────────────

def decompose(df: pd.DataFrame, beta: np.ndarray) -> pd.DataFrame:
    """
    Decompose predicted rate into four labelled driver contributions.

    Components (sum to predicted rate):
        Baseline      = province FE + year FE
                      = ProvMean_Rate + YearMean_Rate - GrandMean_Rate
        Fossil mix    = β1·fW + β2·fW·Dereg + β3·fW·Hybrid
        Gas price     = β6·gW + β7·gW·FossilShare
        Market struct = β4·Dereg + β5·Hybrid + β8·gW·Dereg
        Residual      = Actual - Predicted
    """
    b1,b2,b3,b4,b5,b6,b7,b8 = beta
    d = df.copy()

    fW = d["FossilShare_w"].values
    gW = d["GasPrice_CAD_GJ_w"].values
    f  = d["FossilShare"].values
    D  = d["Dereg"].values
    H  = d["Hybrid"].values

    d["Comp_Baseline"]      = d["ProvMean_Rate"] + d["YearMean_Rate"] - d["GrandMean_Rate"]
    d["Comp_FossilMix"]     = b1*fW + b2*fW*D + b3*fW*H
    d["Comp_GasPassThrough"]= b6*gW + b7*gW*f
    d["Comp_MarketStruct"]  = b4*D  + b5*H  + b8*gW*D
    d["Rate_Predicted"]     = (d["Comp_Baseline"] + d["Comp_FossilMix"] +
                               d["Comp_GasPassThrough"] + d["Comp_MarketStruct"])
    d["Residual"]           = d["Rate_c_kWh"] - d["Rate_Predicted"]

    cols = ["Province","Year","Rate_c_kWh","Rate_Predicted","Residual",
            "Comp_Baseline","Comp_FossilMix","Comp_GasPassThrough","Comp_MarketStruct",
            "FossilShare","GasPrice_CAD_GJ","MarketStructure","GasBenchmark",
            "Dereg","Hybrid"]
    return d[cols].rename(columns={
        "Rate_c_kWh":         "Actual_Rate_c_kWh",
        "Comp_Baseline":      "Driver_Baseline",
        "Comp_FossilMix":     "Driver_FossilMix",
        "Comp_GasPassThrough":"Driver_GasPassThrough",
        "Comp_MarketStruct":  "Driver_MarketStructure",
    })

decomp_df = decompose(df, results["beta"])


# ── 6. GAS PRICE REFERENCE TABLE ─────────────────────────────────────────────

gas_ref = pd.DataFrame({
    "Year":           list(AECO.keys()),
    "AECO_CAD_GJ":    [round(AECO[y], 3) for y in AECO],
    "HH_USD_MMBtu":   [HH_USD[y] for y in AECO],
    "FX_CAD_per_USD": [FX_CAD_PER_USD[y] for y in AECO],
    "HH_CAD_GJ":      [round(HH_CAD_GJ[y], 3) for y in AECO],
}).sort_values("Year").reset_index(drop=True)

# ── 7. RATE DATA WIDE TABLE ───────────────────────────────────────────────────

rate_wide = pd.DataFrame({"Year": YEARS})
for prov in RATES:
    rate_wide[prov] = RATES[prov]

fossil_wide = pd.DataFrame({"Year": YEARS})
for prov in FOSSIL:
    fossil_wide[prov] = [round(v*100, 1) for v in FOSSIL[prov]]


# ── 8. WRITE TO EXCEL ─────────────────────────────────────────────────────────

FNAME = "electricity_rate_model_outputs.xlsx"

with pd.ExcelWriter(FNAME, engine="openpyxl") as writer:
    # Sheet 1: Regression results
    coef_out = results["coef_df"].copy()
    for col in ["Coefficient","Robust_SE","t_stat","CI_lower_95","CI_upper_95"]:
        coef_out[col] = coef_out[col].round(4)
    coef_out["p_value"] = coef_out["p_value"].round(4)
    coef_out.to_excel(writer, sheet_name="1_Regression", index=False, startrow=4)

    # Sheet 2: Full decomposition panel
    decomp_df.round(3).to_excel(writer, sheet_name="2_Decomposition", index=False)

    # Sheet 3: Pivot — driver contributions by province × year
    for driver, label in [
        ("Driver_FossilMix",       "3a_Driver_FossilMix"),
        ("Driver_GasPassThrough",  "3b_Driver_GasPassThrough"),
        ("Driver_MarketStructure", "3c_Driver_MarketStructure"),
        ("Residual",               "3d_Residual"),
    ]:
        piv = decomp_df.pivot(index="Province", columns="Year", values=driver).round(3)
        piv.to_excel(writer, sheet_name=label)

    # Sheet 4: Rate data (wide)
    rate_wide.to_excel(writer, sheet_name="4_Rates_Wide", index=False)

    # Sheet 5: Consumed fossil share (wide, %)
    fossil_wide.to_excel(writer, sheet_name="5_FossilShare_Wide_pct", index=False)

    # Sheet 6: Gas price reference
    gas_ref.to_excel(writer, sheet_name="6_GasPrices", index=False)

    # Sheet 7: Full panel (long format)
    panel_cols = ["Province","Year","Rate_c_kWh","FossilShare","GasPrice_CAD_GJ",
                  "MarketStructure","GasBenchmark","Dereg","Hybrid","AECO_CAD_GJ",
                  "HH_CAD_GJ","FX_CAD_USD"]
    df[panel_cols].round(4).to_excel(writer, sheet_name="7_Panel_Long", index=False)


# ── 9. STYLE THE WORKBOOK ─────────────────────────────────────────────────────

HDR  = PatternFill("solid", fgColor="1F3864")
HDR_F = Font(color="FFFFFF", bold=True, size=10)
ALT  = PatternFill("solid", fgColor="EEF2F7")
BRD  = Border(bottom=Side(style="thin", color="CCCCCC"),
               right=Side(style="thin",  color="CCCCCC"))
POS_F = Font(color="1D6F42", bold=True)   # positive coef — green
NEG_F = Font(color="9B1D20", bold=True)   # negative coef — red
SIG_F = Font(color="1D6F42", bold=True)   # significant — green

def style_table(ws, min_row=1):
    """Apply header + alternating row styling to a sheet."""
    for cell in ws[min_row]:
        cell.fill = HDR; cell.font = HDR_F
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for ri, row in enumerate(ws.iter_rows(min_row=min_row + 1), start=2):
        for cell in row:
            cell.border = BRD
            cell.alignment = Alignment(horizontal="center")
            if ri % 2 == 0:
                cell.fill = ALT
    for col in ws.columns:
        mx = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 3, 24)

wb = load_workbook(FNAME)

# ── Sheet 1: Regression ───────────────────────────────────────────────────────
ws1 = wb["1_Regression"]
ws1["A1"] = "Canadian residential electricity rate model — panel OLS"
ws1["A2"] = (f"Province + year fixed effects (within-group demeaning) · "
             f"N={results['n']} obs · 10 provinces × 13 years (2012–2025 excl. 2014)")
ws1["A3"] = (f"R²={results['r2']:.4f}  Adj.R²={results['adj_r2']:.4f}  "
             f"RMSE={results['rmse']:.3f}¢/kWh  "
             f"SE: HC1 robust")
for r in [1, 2, 3]:
    ws1.cell(r, 1).font = Font(bold=True, size=11 if r==1 else 10)
    ws1.merge_cells(f"A{r}:H{r}")

style_table(ws1, min_row=5)

# Colour coefficient column
coef_col = 2
for row in ws1.iter_rows(min_row=6, max_col=8):
    cell_coef = row[1]  # "Coefficient" column
    cell_sig  = row[7]  # "Significance" column
    try:
        v = float(cell_coef.value)
        cell_coef.font = POS_F if v > 0 else NEG_F
    except (TypeError, ValueError):
        pass
    if cell_sig.value and cell_sig.value != "":
        cell_sig.font = SIG_F

# ── Sheets 2–7: standard styling ─────────────────────────────────────────────
for sh in ["2_Decomposition","3a_Driver_FossilMix","3b_Driver_GasPassThrough",
           "3c_Driver_MarketStructure","3d_Residual",
           "4_Rates_Wide","5_FossilShare_Wide_pct","6_GasPrices","7_Panel_Long"]:
    if sh in wb.sheetnames:
        style_table(wb[sh])

# Add colour-scale conditional formatting to driver pivot sheets
def colorscale(ws, reverse=False):
    mr, mc = ws.max_row, ws.max_column
    if mr < 2 or mc < 2: return
    rng = f"B2:{get_column_letter(mc)}{mr}"
    lo, hi = ("#63BE7B","#F8696B") if not reverse else ("#F8696B","#63BE7B")
    ws.conditional_formatting.add(rng, ColorScaleRule(
        start_type="min", start_color=lo,
        end_type="max",   end_color=hi))

for sh, rev in [("3a_Driver_FossilMix",False),("3b_Driver_GasPassThrough",False),
                ("3c_Driver_MarketStructure",False),("3d_Residual",True)]:
    if sh in wb.sheetnames:
        colorscale(wb[sh], rev)

# Freeze panes
for sh in wb.sheetnames:
    wb[sh].freeze_panes = "B2"

wb.save(FNAME)
print(f"\n✓  Saved → {FNAME}")
print(f"   Sheets: {wb.sheetnames}")
