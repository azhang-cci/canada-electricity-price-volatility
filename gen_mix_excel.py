"""
Canadian Provincial Electricity Generation Mix (2012–2023)
Sources: Canada Energy Regulator (CER) Provincial Energy Profiles,
         Statistics Canada Table 25-10-0020-01, Wikipedia Electricity sector in Canada,
         CER Renewable Energy in Canada provincial pages

Data represents annual generation shares (%) by source for the 10 provinces
matched to the Hydro-Québec rate comparison cities.

City → Province mapping used:
  Calgary / Edmonton → Alberta
  Vancouver          → British Columbia
  Winnipeg           → Manitoba
  Moncton            → New Brunswick
  St. John's         → Newfoundland and Labrador
  Halifax            → Nova Scotia
  Ottawa / Toronto   → Ontario
  Charlottetown      → Prince Edward Island
  Montréal           → Quebec
  Regina             → Saskatchewan

Generation sources tracked:
  Hydro, Nuclear, Natural Gas, Coal, Wind, Solar, Oil/Other, Biomass

Run: pip install openpyxl pandas
     python this_script.py
Output: GenMix_Rates_Panel.xlsx
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule

# ── GENERATION MIX DATA (% of total annual generation) ────────────────────────
# Sources: CER provincial profiles, StatCan 25-10-0020-01, CER Renewable Energy pages
# Format: {year: {province: [hydro, nuclear, nat_gas, coal, wind, solar, oil_other, biomass]}}
# Shares sum to ~100%; minor rounding differences are normal.

SOURCES = ["Hydro_%", "Nuclear_%", "NatGas_%", "Coal_%", "Wind_%", "Solar_%", "OilOther_%", "Biomass_%"]

gen_mix = {
    # ── ALBERTA (AB) ─────────────────────────────────────────────────────────
    # CER: 2021 = hydro 3%, nuclear 0%, gas 63%, coal 22%, wind 9%, bio 2%, solar 1%
    # 2012: coal ~55%, gas ~37%, wind ~5%, hydro ~3%  (pre-major coal phaseout)
    # 2016: coal ~43%, gas ~42%, wind ~8%, hydro ~3%
    # 2019: coal ~30%, gas ~55%, wind ~10%, hydro ~3%
    # 2023: coal ~0% (fully phased Jun 2024, ~2% in 2023), gas ~80%, wind ~12%, hydro ~3%, solar ~3%
    "Alberta": {
        2012: [3,  0, 38, 52,  5, 0,  1, 1],
        2013: [3,  0, 39, 51,  5, 0,  1, 1],
        2014: [3,  0, 40, 48,  6, 0,  1, 2],
        2015: [3,  0, 42, 45,  7, 0,  1, 2],
        2016: [3,  0, 43, 43,  8, 0,  1, 2],
        2017: [3,  0, 48, 40,  7, 0,  1, 1],
        2018: [3,  0, 53, 35,  7, 0,  1, 1],
        2019: [3,  0, 55, 30, 10, 0,  1, 1],
        2020: [3,  0, 59, 27, 10, 0,  0, 1],
        2021: [3,  0, 63, 22,  9, 1,  0, 2],
        2022: [3,  0, 68, 15, 11, 1,  0, 2],
        2023: [3,  0, 79,  2, 13, 2,  0, 1],
    },
    # ── BRITISH COLUMBIA (BC) ─────────────────────────────────────────────────
    # CER: 2023 = hydro 87%, biomass 7%, wind 3%, gas 2%, other 1%
    # CER: 2010 = 96.4% renewable; 2023 = 97.2% renewable, total 56.1 TWh
    "British Columbia": {
        2012: [88, 0,  6, 0, 1, 0, 2, 3],
        2013: [88, 0,  6, 0, 1, 0, 2, 3],
        2014: [88, 0,  5, 0, 2, 0, 2, 3],
        2015: [88, 0,  5, 0, 2, 0, 2, 3],
        2016: [88, 0,  5, 0, 2, 0, 2, 3],
        2017: [88, 0,  5, 0, 2, 0, 2, 3],
        2018: [88, 0,  4, 0, 2, 0, 2, 4],
        2019: [88, 0,  4, 0, 2, 0, 2, 4],
        2020: [88, 0,  4, 0, 2, 0, 2, 4],
        2021: [88, 0,  3, 0, 2, 0, 3, 4],
        2022: [88, 0,  3, 0, 3, 0, 2, 4],
        2023: [87, 0,  2, 0, 3, 0, 1, 7],
    },
    # ── MANITOBA (MB) ─────────────────────────────────────────────────────────
    # CER: 97% hydro in 2019; historically 95–98% hydro
    # Small wind ~1-2%, some oil in remote communities
    "Manitoba": {
        2012: [96, 0, 1, 0, 1, 0, 2, 0],
        2013: [96, 0, 1, 0, 1, 0, 2, 0],
        2014: [96, 0, 1, 0, 1, 0, 2, 0],
        2015: [96, 0, 1, 0, 2, 0, 1, 0],
        2016: [96, 0, 0, 0, 2, 0, 2, 0],
        2017: [96, 0, 0, 0, 2, 0, 2, 0],
        2018: [96, 0, 0, 0, 2, 0, 2, 0],
        2019: [97, 0, 0, 0, 2, 0, 1, 0],
        2020: [96, 0, 0, 0, 2, 0, 2, 0],
        2021: [95, 0, 1, 0, 2, 0, 2, 0],
        2022: [96, 0, 1, 0, 2, 0, 1, 0],
        2023: [95, 0, 1, 0, 3, 0, 1, 0],
    },
    # ── NEW BRUNSWICK (NB) ────────────────────────────────────────────────────
    # CER 2021: nuclear 40%, fossil 27%, hydro 23%, wind+biomass ~10%
    # CER 2023: nuclear 34%, hydro 26%, gas 15%, coal 13%, wind 5%, biomass 5%, oil 2%
    # Point Lepreau refurbishment completed 2012, restoring ~40% nuclear share
    "New Brunswick": {
        2012: [23, 37, 10, 15,  3, 0, 8, 4],
        2013: [23, 39, 10, 13,  4, 0, 7, 4],
        2014: [24, 39,  9, 13,  4, 0, 7, 4],
        2015: [24, 38,  9, 14,  4, 0, 7, 4],
        2016: [24, 38, 10, 13,  4, 0, 7, 4],
        2017: [24, 39, 11, 12,  4, 0, 6, 4],
        2018: [24, 40, 11, 12,  4, 0, 5, 4],
        2019: [24, 40, 12, 11,  4, 0, 5, 4],
        2020: [24, 40, 13, 11,  4, 0, 4, 4],
        2021: [23, 40, 14, 11,  4, 0, 4, 4],
        2022: [25, 30, 15, 18,  4, 0, 3, 5],
        2023: [26, 34, 15, 13,  5, 0, 2, 5],
    },
    # ── NEWFOUNDLAND AND LABRADOR (NL) ────────────────────────────────────────
    # CER: 2023 = hydro 97%, petroleum 2.1%, wind 0.4%, gas 0.6%
    # Muskrat Falls added 824 MW in 2021, further boosting hydro share
    "Newfoundland and Labrador": {
        2012: [93, 0, 0, 0, 0, 0, 7, 0],
        2013: [93, 0, 0, 0, 0, 0, 7, 0],
        2014: [93, 0, 0, 0, 0, 0, 7, 0],
        2015: [93, 0, 0, 0, 0, 0, 7, 0],
        2016: [93, 0, 0, 0, 0, 0, 7, 0],
        2017: [93, 0, 0, 0, 0, 0, 7, 0],
        2018: [94, 0, 0, 0, 0, 0, 6, 0],
        2019: [94, 0, 0, 0, 0, 0, 6, 0],
        2020: [94, 0, 0, 0, 0, 0, 6, 0],
        2021: [96, 0, 0, 0, 0, 0, 4, 0],
        2022: [97, 0, 0, 0, 0, 0, 3, 0],
        2023: [97, 0, 1, 0, 0, 0, 2, 0],
    },
    # ── NOVA SCOTIA (NS) ──────────────────────────────────────────────────────
    # CER 2023: coal 47%, gas 22%, wind 15%, hydro 11%, biomass 4%, solar 1%
    # CER 2010: coal ~63%  |  steady coal decline, wind/gas growth
    "Nova Scotia": {
        2012: [10,  0, 13, 63, 10, 0, 2, 2],
        2013: [10,  0, 13, 62, 11, 0, 2, 2],
        2014: [10,  0, 13, 62, 11, 0, 2, 2],
        2015: [11,  0, 13, 60, 12, 0, 2, 2],
        2016: [11,  0, 14, 58, 13, 0, 2, 2],
        2017: [11,  0, 15, 56, 14, 0, 2, 2],
        2018: [11,  0, 17, 54, 14, 0, 2, 2],
        2019: [11,  0, 18, 52, 15, 0, 2, 2],
        2020: [11,  0, 19, 51, 15, 0, 2, 2],
        2021: [11,  0, 21, 50, 14, 0, 1, 3],
        2022: [11,  0, 21, 51, 13, 0, 1, 3],
        2023: [11,  0, 22, 47, 15, 1, 1, 4],
    },
    # ── ONTARIO (ON) ──────────────────────────────────────────────────────────
    # CER 2021: nuclear 55%, hydro 24%, wind 8%, solar 4%, gas 7%, bio 2%
    # CER 2023: nuclear 51%, hydro 25%, gas 14%, wind 8%, solar 1%, bio 1%
    # Coal phased out April 2014; gradual gas growth for peaking
    "Ontario": {
        2012: [23, 54,  8, 10,  4, 0, 0, 1],
        2013: [24, 57,  7,  5,  5, 1, 0, 1],
        2014: [23, 60,  8,  1,  6, 1, 0, 1],  # coal phased out mid-2014
        2015: [23, 59,  9,  0,  7, 1, 0, 1],
        2016: [23, 59,  9,  0,  7, 1, 0, 1],
        2017: [24, 59,  9,  0,  8, 1, 0, 1],  # Wikipedia: 2017 total 132.1 TWh
        2018: [24, 58, 10,  0,  7, 1, 0, 0],
        2019: [24, 58, 10,  0,  7, 1, 0, 0],
        2020: [24, 58, 10,  0,  7, 1, 0, 0],
        2021: [24, 55, 10,  0,  8, 2, 0, 1],
        2022: [25, 54, 11,  0,  8, 1, 0, 1],
        2023: [25, 51, 14,  0,  8, 1, 0, 1],
    },
    # ── PRINCE EDWARD ISLAND (PE) ────────────────────────────────────────────
    # PEI generates ~99% from wind (CER/Wikipedia); imports balance from NB
    # Wind grew from ~25% in 2010 to ~99% locally generated by ~2016
    # PEI generates locally mostly wind; imports cover ~60% of consumption from NB
    # "Generation mix" here refers to on-island generation only
    "Prince Edward Island": {
        2012: [0,  0,  5, 0, 90, 0, 5, 0],
        2013: [0,  0,  3, 0, 93, 0, 4, 0],
        2014: [0,  0,  3, 0, 94, 0, 3, 0],
        2015: [0,  0,  2, 0, 95, 0, 3, 0],
        2016: [0,  0,  1, 0, 97, 0, 2, 0],
        2017: [0,  0,  1, 0, 97, 0, 2, 0],
        2018: [0,  0,  1, 0, 97, 0, 2, 0],
        2019: [0,  0,  1, 0, 98, 0, 1, 0],
        2020: [0,  0,  1, 0, 98, 0, 1, 0],
        2021: [0,  0,  1, 0, 98, 0, 1, 0],
        2022: [0,  0,  1, 0, 98, 0, 1, 0],
        2023: [0,  0,  1, 0, 98, 0, 1, 0],
    },
    # ── QUEBEC (QC) ───────────────────────────────────────────────────────────
    # Overwhelmingly hydro (~97%); small wind, some oil in remote areas, biomass
    "Quebec": {
        2012: [97, 0, 0, 0, 2, 0, 1, 0],
        2013: [97, 0, 0, 0, 2, 0, 1, 0],
        2014: [97, 0, 0, 0, 2, 0, 1, 0],
        2015: [97, 0, 0, 0, 2, 0, 1, 0],
        2016: [96, 0, 0, 0, 3, 0, 1, 0],
        2017: [96, 0, 0, 0, 3, 0, 1, 0],
        2018: [96, 0, 0, 0, 3, 0, 1, 0],
        2019: [96, 0, 0, 0, 3, 0, 1, 0],
        2020: [96, 0, 0, 0, 3, 0, 1, 0],
        2021: [96, 0, 0, 0, 3, 0, 1, 0],
        2022: [97, 0, 0, 0, 3, 0, 0, 0],
        2023: [96, 0, 0, 0, 3, 0, 1, 0],
    },
    # ── SASKATCHEWAN (SK) ─────────────────────────────────────────────────────
    # Wikipedia 2013: fossil ~86%; CER: coal dominant, growing gas & wind
    # SaskPower coal phase-down: coal ~40% by 2022, gas growing
    "Saskatchewan": {
        2012: [20, 0, 20, 55,  3, 0, 2, 0],
        2013: [20, 0, 20, 54,  4, 0, 2, 0],
        2014: [20, 0, 20, 53,  5, 0, 2, 0],
        2015: [19, 0, 21, 52,  6, 0, 2, 0],
        2016: [19, 0, 22, 51,  6, 0, 2, 0],
        2017: [18, 0, 24, 49,  7, 0, 2, 0],
        2018: [18, 0, 26, 47,  7, 0, 2, 0],
        2019: [17, 0, 28, 46,  7, 0, 2, 0],
        2020: [17, 0, 30, 44,  7, 0, 2, 0],
        2021: [17, 0, 33, 40,  8, 0, 2, 0],
        2022: [17, 0, 35, 38,  9, 0, 1, 0],
        2023: [16, 0, 37, 37,  9, 0, 1, 0],
    },
}

# ── AVERAGE PRICES (¢/kWh) FROM HYDRO-QUÉBEC REPORTS ─────────────────────────
# 1,000 kWh/month consumption point (most representative for residential)
# Source: Hydro-Québec Comparison of Electricity Prices 2012–2025
# Cities mapped to provinces for merge

city_province = {
    "Calgary, AB":        "Alberta",
    "Edmonton, AB":       "Alberta",
    "Vancouver, BC":      "British Columbia",
    "Winnipeg, MB":       "Manitoba",
    "Moncton, NB":        "New Brunswick",
    "St. John's, NL":     "Newfoundland and Labrador",
    "Halifax, NS":        "Nova Scotia",
    "Ottawa, ON":         "Ontario",
    "Toronto, ON":        "Ontario",
    "Charlottetown, PE":  "Prince Edward Island",
    "Montréal, QC":       "Quebec",
    "Regina, SK":         "Saskatchewan",
}

# [625kWh, 750kWh, 1000kWh, 2000kWh, 3000kWh]
avg_prices = {
    2012: {
        "Calgary, AB":       [15.01,14.51,13.89,12.95,12.64],
        "Edmonton, AB":      [14.21,13.63,12.90,11.82,11.46],
        "Vancouver, BC":     [7.91, 8.14, 8.78, 9.74,10.06],
        "Winnipeg, MB":      [7.87, 7.68, 7.46, 7.11, 7.00],
        "Moncton, NB":       [13.01,12.48,11.82,10.84,10.51],
        "St. John's, NL":    [12.73,12.31,11.80,11.02,10.77],
        "Halifax, NS":       [15.66,15.37,15.01,14.46,14.28],
        "Ottawa, ON":        [13.40,13.16,13.14,13.17,13.18],
        "Toronto, ON":       [14.51,14.08,13.57,13.07,12.91],
        "Charlottetown, PE": [15.98,15.33,14.51,13.28,11.92],
        "Montréal, QC":      [7.27, 6.95, 6.76, 7.13, 7.26],
        "Regina, SK":        [13.69,13.18,12.54,11.57,11.25],
    },
    2013: {
        "Calgary, AB":       [15.93,15.44,14.81,13.88,13.56],
        "Edmonton, AB":      [15.14,14.59,13.90,12.87,12.52],
        "Vancouver, BC":     [8.03, 8.26, 8.91, 9.88,10.21],
        "Winnipeg, MB":      [8.04, 7.85, 7.63, 7.28, 7.17],
        "Moncton, NB":       [13.01,12.48,11.82,10.84,10.51],
        "St. John's, NL":    [13.47,13.06,12.55,11.78,11.52],
        "Halifax, NS":       [16.10,15.81,15.45,14.90,14.72],
        "Ottawa, ON":        [12.90,12.67,12.39,11.97,11.83],
        "Toronto, ON":       [13.28,12.91,12.48,12.08,11.94],
        "Charlottetown, PE": [16.34,15.69,14.87,13.64,12.23],
        "Montréal, QC":      [7.36, 7.04, 6.87, 7.32, 7.48],
        "Regina, SK":        [14.37,13.83,13.15,12.14,11.80],
    },
    2015: {
        "Calgary, AB":       [12.94,12.37,11.66,10.58,10.23],
        "Edmonton, AB":      [12.91,12.30,11.55,10.41,10.04],
        "Vancouver, BC":     [9.27, 9.54,10.29,11.42,11.80],
        "Winnipeg, MB":      [8.55, 8.35, 8.11, 7.75, 7.62],
        "Moncton, NB":       [13.53,12.98,12.30,11.27,10.93],
        "St. John's, NL":    [12.41,12.03,11.55,10.84,10.60],
        "Halifax, NS":       [16.68,16.39,16.03,15.49,15.31],
        "Ottawa, ON":        [14.78,14.52,14.20,13.72,13.55],
        "Toronto, ON":       [15.06,14.71,14.31,13.99,13.89],
        "Charlottetown, PE": [17.09,16.44,15.62,14.39,13.05],
        "Montréal, QC":      [7.63, 7.31, 7.19, 7.90, 8.13],
        "Regina, SK":        [15.59,15.05,14.37,13.36,13.02],
    },
    2016: {
        "Calgary, AB":       [11.68,11.11,10.40, 9.33, 8.97],
        "Edmonton, AB":      [11.70,11.11,10.37, 9.26, 8.89],
        "Vancouver, BC":     [9.64, 9.92,10.70,11.88,12.27],
        "Winnipeg, MB":      [8.88, 8.68, 8.43, 8.05, 7.92],
        "Moncton, NB":       [13.75,13.19,12.50,11.46,11.11],
        "St. John's, NL":    [12.89,12.48,11.96,11.19,10.93],
        "Halifax, NS":       [16.53,16.24,15.88,15.34,15.16],
        "Ottawa, ON":        [17.01,16.63,16.15,15.44,15.20],
        "Toronto, ON":       [19.23,18.60,17.81,16.62,16.23],
        "Charlottetown, PE": [17.49,16.84,16.02,14.79,13.46],
        "Montréal, QC":      [7.66, 7.34, 7.23, 7.95, 8.20],
        "Regina, SK":        [15.86,15.32,14.65,13.63,13.30],
    },
    2017: {
        "Calgary, AB":       [11.65,11.12,10.45, 9.45, 9.12],
        "Edmonton, AB":      [11.73,11.11,10.34, 9.18, 8.80],
        "Vancouver, BC":     [9.98,10.27,11.08,12.30,12.70],
        "Winnipeg, MB":      [9.18, 8.97, 8.71, 8.32, 8.19],
        "Moncton, NB":       [14.27,13.69,12.97,11.89,11.53],
        "St. John's, NL":    [12.10,11.68,11.15,10.36,10.10],
        "Halifax, NS":       [16.80,16.51,16.15,15.60,15.42],
        "Ottawa, ON":        [16.19,15.75,15.21,14.40,14.13],
        "Toronto, ON":       [17.91,17.20,16.32,14.99,14.55],
        "Charlottetown, PE": [17.89,17.24,16.42,15.19,13.82],
        "Montréal, QC":      [7.77, 7.45, 7.07, 8.00, 8.30],
        "Regina, SK":        [17.26,16.67,15.94,14.84,14.47],
    },
    2018: {
        "Calgary, AB":       [17.30,16.63,15.79,14.53,14.10],
        "Edmonton, AB":      [15.79,15.15,14.35,13.15,12.75],
        "Vancouver, BC":     [10.28,10.58,11.42,12.67,13.09],
        "Winnipeg, MB":      [9.49, 9.27, 9.00, 8.60, 8.47],
        "Moncton, NB":       [14.27,13.69,12.97,11.89,11.53],
        "St. John's, NL":    [12.98,12.56,12.03,11.24,10.97],
        "Halifax, NS":       [17.06,16.77,16.41,15.87,15.69],
        "Ottawa, ON":        [13.33,12.81,12.16,11.18,10.85],
        "Toronto, ON":       [15.09,14.26,13.24,11.69,11.18],
        "Charlottetown, PE": [18.30,17.65,16.83,15.60,14.21],
        "Montréal, QC":      [7.86, 7.54, 7.13, 8.00, 8.37],
        "Regina, SK":        [17.87,17.27,16.51,15.37,14.99],
    },
    2019: {
        "Calgary, AB":       [17.20,16.55,15.74,14.53,14.13],
        "Edmonton, AB":      [16.18,15.51,14.68,13.44,13.02],
        "Vancouver, BC":     [10.47,10.77,11.62,12.90,13.32],
        "Winnipeg, MB":      [9.87, 9.65, 9.37, 8.95, 8.81],
        "Moncton, NB":       [14.41,13.82,13.10,12.00,11.64],
        "St. John's, NL":    [13.75,13.33,12.80,12.01,11.75],
        "Halifax, NS":       [17.34,17.05,16.69,16.14,15.96],
        "Ottawa, ON":        [13.43,12.81,12.04,10.88,10.50],
        "Toronto, ON":       [16.12,15.13,13.89,12.04,11.42],
        "Charlottetown, PE": [18.30,17.65,16.83,15.60,14.21],
        "Montréal, QC":      [8.03, 7.71, 7.30, 8.01, 8.47],
        "Regina, SK":        [17.87,17.27,16.51,15.37,14.99],
    },
    2020: {
        "Calgary, AB":       [16.28,15.64,14.83,13.62,13.22],
        "Edmonton, AB":      [15.89,15.18,14.29,12.96,12.52],
        "Vancouver, BC":     [10.38,10.67,11.51,12.77,13.19],
        "Winnipeg, MB":      [10.12, 9.89, 9.60, 9.17, 9.03],
        "Moncton, NB":       [14.76,14.17,13.42,12.30,11.93],
        "St. John's, NL":    [14.55,14.13,13.60,12.81,12.54],
        "Halifax, NS":       [17.54,17.25,16.89,16.35,16.17],
        "Ottawa, ON":        [11.49,10.96,10.29, 9.28, 8.95],
        "Toronto, ON":       [12.63,11.95,11.10, 9.82, 9.39],
        "Charlottetown, PE": [18.30,17.65,16.83,15.60,14.21],
        "Montréal, QC":      [8.03, 7.71, 7.30, 8.01, 8.47],
        "Regina, SK":        [17.87,17.27,16.51,15.37,14.99],
    },
    2021: {
        "Calgary, AB":       [18.88,18.16,17.26,15.90,15.45],
        "Edmonton, AB":      [18.63,17.91,16.99,15.62,15.16],
        "Vancouver, BC":     [10.44,10.74,11.58,12.84,13.26],
        "Winnipeg, MB":      [10.40,10.16, 9.87, 9.43, 9.28],
        "Moncton, NB":       [15.03,14.42,13.66,12.52,12.14],
        "St. John's, NL":    [14.55,14.13,13.60,12.81,12.54],
        "Halifax, NS":       [17.74,17.45,17.09,16.55,16.37],
        "Ottawa, ON":        [13.87,13.24,12.45,11.28,10.88],
        "Toronto, ON":       [15.25,14.44,13.43,11.91,11.41],
        "Charlottetown, PE": [18.85,18.20,17.38,16.15,14.73],
        "Montréal, QC":      [8.13, 7.81, 7.39, 8.11, 8.58],
        "Regina, SK":        [17.87,17.27,16.51,15.37,14.99],
    },
    2022: {
        "Calgary, AB":       [21.63,20.88,19.94,18.53,18.06],
        "Edmonton, AB":      [21.16,20.42,19.48,18.07,17.61],
        "Vancouver, BC":     [10.31,10.59,11.39,12.60,13.00],
        "Winnipeg, MB":      [10.80,10.55,10.24, 9.78, 9.63],
        "Moncton, NB":       [15.33,14.71,13.94,12.77,12.39],
        "St. John's, NL":    [14.70,14.28,13.76,12.98,12.72],
        "Halifax, NS":       [17.95,17.66,17.30,16.76,16.58],
        "Ottawa, ON":        [14.50,13.81,12.94,11.64,11.21],
        "Toronto, ON":       [15.87,14.99,13.88,12.23,11.68],
        "Charlottetown, PE": [19.25,18.60,17.78,16.55,15.13],
        "Montréal, QC":      [8.35, 8.01, 7.59, 8.32, 8.80],
        "Regina, SK":        [17.87,17.27,16.51,15.37,14.99],
    },
    2023: {
        "Calgary, AB":       [31.55,30.78,29.80,28.34,27.86],
        "Edmonton, AB":      [29.36,28.66,27.78,26.47,26.03],
        "Vancouver, BC":     [10.51,10.80,11.62,12.85,13.26],
        "Winnipeg, MB":      [10.80,10.55,10.24, 9.78, 9.63],
        "Moncton, NB":       [16.08,15.43,14.61,13.38,12.97],
        "St. John's, NL":    [14.67,14.25,13.73,12.94,12.68],
        "Halifax, NS":       [19.42,18.91,18.27,17.31,16.99],
        "Ottawa, ON":        [15.23,14.45,13.48,12.02,11.53],
        "Toronto, ON":       [16.03,15.07,13.88,12.08,11.49],
        "Charlottetown, PE": [19.25,18.60,17.78,16.55,15.13],
        "Montréal, QC":      [8.60, 8.25, 7.81, 8.57, 9.06],
        "Regina, SK":        [19.69,18.89,17.89,16.39,15.89],
    },
    2024: {
        "Calgary, AB":       [25.28,24.34,23.17,21.41,20.82],
        "Edmonton, AB":      [25.68,24.92,23.98,22.58,22.11],
        "Vancouver, BC":     [11.49,11.61,12.06,12.73,12.96],
        "Winnipeg, MB":      [11.10,10.85,10.53,10.06, 9.90],
        "Moncton, NB":       [17.99,17.24,16.30,14.88,14.41],
        "St. John's, NL":    [15.56,15.14,14.62,13.84,13.58],
        "Halifax, NS":       [20.61,20.10,19.46,18.51,18.19],
        "Ottawa, ON":        [15.97,15.22,14.28,12.87,12.22],
        "Toronto, ON":       [17.15,16.22,15.06,13.31,12.72],
        "Charlottetown, PE": [20.60,19.95,19.13,17.90,16.37],
        "Montréal, QC":      [8.85, 8.50, 8.05, 8.83, 9.34],
        "Regina, SK":        [19.69,18.89,17.89,16.39,15.89],
    },
    2025: {
        "Calgary, AB":       [25.05,24.10,22.90,21.11,20.51],
        "Edmonton, AB":      [23.67,22.93,22.00,20.62,20.15],
        "Vancouver, BC":     [12.28,12.32,12.60,13.02,13.17],
        "Winnipeg, MB":      [11.10,10.85,10.53,10.06, 9.90],
        "Moncton, NB":       [18.38,17.59,16.61,15.13,14.64],
        "St. John's, NL":    [16.52,16.11,15.59,14.80,14.54],
        "Halifax, NS":       [21.63,21.12,20.48,19.52,19.20],
        "Ottawa, ON":        [16.18,15.37,14.35,12.82,12.33],
        "Toronto, ON":       [17.98,16.91,15.57,13.56,12.89],
        "Charlottetown, PE": [21.16,20.51,19.69,18.46,16.89],
        "Montréal, QC":      [9.12, 8.75, 8.29, 9.10, 9.61],
        "Regina, SK":        [19.69,18.89,17.89,16.39,15.89],
    },
}

CONS_LABELS = ["625kWh", "750kWh", "1000kWh", "2000kWh", "3000kWh"]

# ── BUILD PANEL DATAFRAME ──────────────────────────────────────────────────────
rows = []
for year, city_data in sorted(avg_prices.items()):
    for city, prices in city_data.items():
        prov = city_province[city]
        mix  = gen_mix.get(prov, {}).get(year, [None]*8)
        # Fossil share = gas + coal + oil/other
        fossil = None if None in mix else mix[2] + mix[3] + mix[6]
        lowcarbon = None if None in mix else mix[0] + mix[1] + mix[4] + mix[5] + mix[7]
        row = {
            "Year":        year,
            "City":        city,
            "Province":    prov,
            "Rate_625kWh":   prices[0],
            "Rate_750kWh":   prices[1],
            "Rate_1000kWh":  prices[2],
            "Rate_2000kWh":  prices[3],
            "Rate_3000kWh":  prices[4],
        }
        for i, src in enumerate(SOURCES):
            row[src] = mix[i]
        row["FossilShare_%"]   = fossil
        row["LowCarbonShare_%"] = lowcarbon
        rows.append(row)

df_panel = pd.DataFrame(rows)

# ── PROVINCE-LEVEL AGGREGATED (avg of cities within same province) ─────────────
df_prov = (df_panel.groupby(["Year", "Province"])
           .agg({
               "Rate_625kWh":   "mean",
               "Rate_750kWh":   "mean",
               "Rate_1000kWh":  "mean",
               "Rate_2000kWh":  "mean",
               "Rate_3000kWh":  "mean",
               **{s: "first" for s in SOURCES},
               "FossilShare_%":   "first",
               "LowCarbonShare_%": "first",
           })
           .reset_index())

# ── GEN MIX WIDE (provinces × years) ─────────────────────────────────────────
mix_rows = []
for prov, year_data in gen_mix.items():
    for year, vals in sorted(year_data.items()):
        r = {"Province": prov, "Year": year}
        for i, s in enumerate(SOURCES):
            r[s] = vals[i]
        fossil = vals[2] + vals[3] + vals[6]
        r["FossilShare_%"]    = fossil
        r["LowCarbonShare_%"] = 100 - fossil
        mix_rows.append(r)
df_mix = pd.DataFrame(mix_rows)

# ── WRITE EXCEL ────────────────────────────────────────────────────────────────
fname = "GenMix_Rates_Panel.xlsx"

with pd.ExcelWriter(fname, engine="openpyxl") as writer:
    df_panel.to_excel(writer, sheet_name="Full_Panel",         index=False)
    df_prov.to_excel( writer, sheet_name="Province_Panel",     index=False)
    df_mix.to_excel(  writer, sheet_name="GenMix_AllProvinces",index=False)
    # Pivot: fossil share by province × year
    fossil_pivot = df_mix.pivot(index="Province", columns="Year", values="FossilShare_%")
    fossil_pivot.to_excel(writer, sheet_name="FossilShare_Pivot")
    # Pivot: 1000kWh rate by province × year
    rate_pivot = df_prov.pivot(index="Province", columns="Year", values="Rate_1000kWh")
    rate_pivot.to_excel(writer, sheet_name="Rate1000kWh_Pivot")

# ── STYLE ─────────────────────────────────────────────────────────────────────
wb = load_workbook(fname)

HDR_FILL  = PatternFill("solid", fgColor="1F3864")
HDR_FONT  = Font(color="FFFFFF", bold=True, size=10)
ALT_FILL  = PatternFill("solid", fgColor="EEF2F7")
BORDER    = Border(bottom=Side(style="thin", color="CCCCCC"),
                   right=Side(style="thin",  color="CCCCCC"))

def style_ws(ws):
    for cell in ws[1]:
        cell.fill = HDR_FILL; cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center")
            if ri % 2 == 0:
                cell.fill = ALT_FILL
    for col in ws.columns:
        mx = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 3, 22)

for sh in ["Full_Panel", "Province_Panel", "GenMix_AllProvinces"]:
    style_ws(wb[sh])

# Color-scale conditional formatting for pivot sheets
def add_colorscale(ws, min_row, max_row, min_col, max_col, reverse=False):
    rng = (f"{get_column_letter(min_col)}{min_row}:"
           f"{get_column_letter(max_col)}{max_row}")
    lo, hi = ("63BE7B", "F8696B") if not reverse else ("F8696B", "63BE7B")
    ws.conditional_formatting.add(rng, ColorScaleRule(
        start_type="min", start_color=lo,
        end_type="max",   end_color=hi))

for sh_name, reverse in [("FossilShare_Pivot", False), ("Rate1000kWh_Pivot", False)]:
    ws = wb[sh_name]
    for cell in ws[1]:
        cell.fill = HDR_FILL; cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center")
    for col in ws.columns:
        mx = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 3, 14)
    max_row = ws.max_row; max_col = ws.max_column
    add_colorscale(ws, 2, max_row, 2, max_col)

# Add source notes to gen mix sheet
ws_mix = wb["GenMix_AllProvinces"]
ws_mix.insert_rows(1)
ws_mix["A1"] = ("Sources: CER Provincial Energy Profiles; StatCan Table 25-10-0020-01; "
                "CER Renewable Energy in Canada pages (2025). "
                "Generation mix covers 2012–2023 (rate data extends to 2025). "
                "Shares are % of total annual provincial generation.")
ws_mix["A1"].font = Font(italic=True, color="444444", size=9)
ws_mix.merge_cells(f"A1:{get_column_letter(ws_mix.max_column)}1")

wb.save(fname)
print(f"✓ Saved {fname}")
print(f"  Sheets: {wb.sheetnames}")
print(f"  Full panel: {len(df_panel)} rows | Province panel: {len(df_prov)} rows")
