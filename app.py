"""
Georgia Trade Report Web Application.
Generates interactive trade reports between Georgia and selected countries
using data from the National Statistics Office of Georgia (GeoStat).
"""

import json
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template

app = Flask(__name__)

DATA_DIR = Path(__file__).parent / "geostat_data"

# ---------------------------------------------------------------------------
# Data loading helpers
# ---------------------------------------------------------------------------

_cache = {}


def _load_sheet(filename, sheet):
    key = (filename, sheet)
    if key not in _cache:
        path = DATA_DIR / filename
        _cache[key] = pd.read_excel(path, sheet_name=sheet, header=None)
    return _cache[key]


def get_countries():
    """Return list of {code, name} dicts for all countries in the dataset."""
    df = _load_sheet("Export-Country_1995-2026.xlsx", "1995-2025-years")
    rows = df.iloc[4:, [0, 1]].copy()
    rows.columns = ["code", "name"]
    # Keep only rows with numeric country codes
    rows = rows[pd.to_numeric(rows["code"], errors="coerce").notna()]
    rows["code"] = rows["code"].astype(int)
    rows = rows.dropna(subset=["name"])
    return rows.to_dict("records")


def _get_country_row(df, country_code):
    """Find the row index for a given country code."""
    for i in range(4, len(df)):
        val = df.iloc[i, 0]
        try:
            if int(float(val)) == int(country_code):
                return i
        except (ValueError, TypeError):
            continue
    return None


def get_yearly_trade(country_code):
    """
    Return yearly export/import data for a country.
    Returns dict with years as keys, values as {export, import, turnover, balance}.
    """
    exp_df = _load_sheet("Export-Country_1995-2026.xlsx", "1995-2025-years")
    imp_df = _load_sheet("Import-Country-1995-2026.xlsx", "1995-2025-years")

    exp_row = _get_country_row(exp_df, country_code)
    imp_row = _get_country_row(imp_df, country_code)

    if exp_row is None or imp_row is None:
        return {}

    years_raw = exp_df.iloc[3, 2:].tolist()
    exp_vals = exp_df.iloc[exp_row, 2:].tolist()
    imp_vals = imp_df.iloc[imp_row, 2:].tolist()

    result = {}
    for y, e, i in zip(years_raw, exp_vals, imp_vals):
        year_str = str(y).replace("*", "").strip()
        try:
            year_int = int(float(year_str))
        except (ValueError, TypeError):
            continue
        try:
            exp_v = float(e) if pd.notna(e) else 0
        except (ValueError, TypeError):
            exp_v = 0
        try:
            imp_v = float(i) if pd.notna(i) else 0
        except (ValueError, TypeError):
            imp_v = 0

        result[year_int] = {
            "export": round(exp_v, 2),
            "import": round(imp_v, 2),
            "turnover": round(exp_v + imp_v, 2),
            "balance": round(exp_v - imp_v, 2),
        }

    return result


def get_current_year_trade(country_code):
    """Get the current partial-year (2026) trade data."""
    exp_df = _load_sheet("Export-Country_1995-2026.xlsx", "2026")
    imp_df = _load_sheet("Import-Country-1995-2026.xlsx", "2026")

    exp_row = _get_country_row(exp_df, country_code)
    imp_row = _get_country_row(imp_df, country_code)

    if exp_row is None or imp_row is None:
        return None

    # Column 2 is the cumulative (e.g. Jan-Feb), col 3+ are monthly
    months_header = exp_df.iloc[4, 2:].tolist()
    exp_vals = exp_df.iloc[exp_row, 2:].tolist()
    imp_vals = imp_df.iloc[imp_row, 2:].tolist()

    monthly = []
    for m, e, i in zip(months_header, exp_vals, imp_vals):
        if pd.isna(m):
            continue
        exp_v = float(e) if pd.notna(e) else 0
        imp_v = float(i) if pd.notna(i) else 0
        monthly.append({
            "period": str(m),
            "export": round(exp_v, 2),
            "import": round(imp_v, 2),
            "turnover": round(exp_v + imp_v, 2),
            "balance": round(exp_v - imp_v, 2),
        })

    return monthly


def get_hs4_data(trade_type="export"):
    """
    Get HS4 product data (total Georgia, not per country).
    Returns list of {code, name, yearly_values}.
    """
    if trade_type == "export":
        filename = "Export-Product-by-4-digit-2015-2026.xlsx"
    else:
        filename = "Import-Product-by-4-digit-2015-2026.xlsx"

    df = _load_sheet(filename, "2020-2025-years")
    years = df.iloc[3, 2:].tolist()
    year_labels = []
    for y in years:
        s = str(y).replace("*", "").strip()
        try:
            year_labels.append(int(float(s)))
        except (ValueError, TypeError):
            year_labels.append(s)

    products = []
    for i in range(4, len(df)):
        code = df.iloc[i, 0]
        name = df.iloc[i, 1]

        if pd.isna(code) or pd.isna(name):
            continue
        if str(name).strip() in ("Total Exports", "Total Imports", "of which:"):
            continue

        try:
            code_int = int(float(code))
        except (ValueError, TypeError):
            continue

        vals = {}
        for y_label, v in zip(year_labels, df.iloc[i, 2:].tolist()):
            try:
                vals[str(y_label)] = round(float(v), 2) if pd.notna(v) else 0
            except (ValueError, TypeError):
                vals[str(y_label)] = 0

        products.append({
            "code": f"{code_int:04d}",
            "name": name.strip(),
            "values": vals,
        })

    return products, [str(y) for y in year_labels]


def get_monthly_trade(country_code, year):
    """Get monthly trade data for a country in a specific year."""
    exp_df = _load_sheet("Export-Country_1995-2026.xlsx", "1995-2025-months")
    imp_df = _load_sheet("Import-Country-1995-2026.xlsx", "1995-2025-months")

    exp_row = _get_country_row(exp_df, country_code)
    imp_row = _get_country_row(imp_df, country_code)

    if exp_row is None or imp_row is None:
        return []

    # Find columns for the requested year
    # Row 3 has year headers, row 4 has month names
    year_row = exp_df.iloc[3, 2:].tolist()
    month_row = exp_df.iloc[4, 2:].tolist()

    monthly = []
    month_names = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ]

    for col_offset, (yr, mn) in enumerate(zip(year_row, month_row)):
        yr_str = str(yr).replace("*", "").strip() if pd.notna(yr) else ""
        # Year headers span across months - find the right year group
        try:
            yr_int = int(float(yr_str))
            current_year = yr_int
        except (ValueError, TypeError):
            pass

        if pd.isna(mn):
            continue

        mn_str = str(mn).strip()
        if mn_str not in month_names:
            continue

        # Check if we've reached the right year section
        # We need to track the current year from the year row
        col_idx = col_offset + 2
        # Walk backwards to find the year
        found_year = None
        for back in range(col_offset, -1, -1):
            y = year_row[back]
            if pd.notna(y):
                try:
                    found_year = int(float(str(y).replace("*", "").strip()))
                    break
                except (ValueError, TypeError):
                    continue

        if found_year != int(year):
            continue

        exp_v = exp_df.iloc[exp_row, col_idx]
        imp_v = imp_df.iloc[imp_row, col_idx]
        exp_v = float(exp_v) if pd.notna(exp_v) else 0
        imp_v = float(imp_v) if pd.notna(imp_v) else 0

        monthly.append({
            "month": mn_str,
            "export": round(exp_v, 2),
            "import": round(imp_v, 2),
            "turnover": round(exp_v + imp_v, 2),
            "balance": round(exp_v - imp_v, 2),
        })

    return monthly


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/countries")
def api_countries():
    countries = get_countries()
    return jsonify(countries)


@app.route("/api/report/<int:country_code>")
def api_report(country_code):
    yearly = get_yearly_trade(country_code)
    current_year = get_current_year_trade(country_code)

    # Get country name
    countries = get_countries()
    country_name = next(
        (c["name"] for c in countries if c["code"] == country_code),
        f"Country {country_code}",
    )

    # HS4 top products (total Georgia)
    exp_products, exp_years = get_hs4_data("export")
    imp_products, imp_years = get_hs4_data("import")

    # Sort by most recent year value, take top 15
    if exp_years:
        latest = exp_years[-1]
        exp_products.sort(key=lambda p: p["values"].get(latest, 0), reverse=True)
        imp_products.sort(key=lambda p: p["values"].get(latest, 0), reverse=True)

    return jsonify({
        "country_code": country_code,
        "country_name": country_name,
        "yearly": yearly,
        "current_year_2026": current_year,
        "top_export_products": exp_products[:15],
        "top_import_products": imp_products[:15],
        "hs4_years": exp_years,
    })


@app.route("/api/monthly/<int:country_code>/<int:year>")
def api_monthly(country_code, year):
    data = get_monthly_trade(country_code, year)
    return jsonify(data)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
