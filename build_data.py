"""
Convert GeoStat XLSX files into compact JSON for the static site.
Run: python build_data.py
Outputs: data/trade.json
"""

import json
from pathlib import Path

import pandas as pd

DATA_DIR = Path("geostat_data")
OUT_DIR = Path("data")


def load(filename, sheet):
    return pd.read_excel(DATA_DIR / filename, sheet_name=sheet, header=None)


def extract_country_yearly(filename, sheet="1995-2025-years"):
    df = load(filename, sheet)
    years_raw = df.iloc[3, 2:].tolist()
    years = []
    for y in years_raw:
        try:
            years.append(int(float(str(y).replace("*", "").strip())))
        except (ValueError, TypeError):
            years.append(str(y))

    countries = {}
    for i in range(4, len(df)):
        code = df.iloc[i, 0]
        name = df.iloc[i, 1]
        if pd.isna(code) or pd.isna(name):
            continue
        try:
            code_int = int(float(code))
        except (ValueError, TypeError):
            continue

        vals = []
        for v in df.iloc[i, 2:].tolist():
            try:
                vals.append(round(float(v), 2) if pd.notna(v) else 0)
            except (ValueError, TypeError):
                vals.append(0)

        countries[str(code_int)] = {"name": str(name).strip(), "values": vals}

    return years, countries


def extract_country_monthly_current_year(filename):
    """Read the current-year sheet (named with a 4-digit year, e.g. '2026').

    Structure:
      row 3: Code | Countries | <year>* | NaN ...
      row 4: NaN  | NaN       | Jan-Feb (YTD, skip) | January | February | ...
      row 5+: country data
    """
    import openpyxl

    wb = openpyxl.load_workbook(DATA_DIR / filename, read_only=True)
    year_sheets = [s for s in wb.sheetnames if s.rstrip("*").strip().isdigit()]
    wb.close()

    if not year_sheets:
        return {}

    month_names = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ]
    countries = {}

    for sheet in year_sheets:
        df = load(filename, sheet)
        year_str = str(df.iloc[3, 2]).replace("*", "").strip()
        try:
            year = int(float(year_str))
        except (ValueError, TypeError):
            continue

        # Row 4 (index 4): identify columns that hold individual month names
        month_row = df.iloc[4, 2:].tolist()
        col_map = {}  # offset from col-index 2 -> month name
        for j, cell in enumerate(month_row):
            m = str(cell).strip() if pd.notna(cell) else ""
            if m in month_names:
                col_map[j] = m

        for i in range(5, len(df)):
            code = df.iloc[i, 0]
            if pd.isna(code):
                continue
            try:
                code_int = int(float(code))
            except (ValueError, TypeError):
                continue

            for j, month in col_map.items():
                v = df.iloc[i, j + 2]
                try:
                    val = round(float(v), 2) if pd.notna(v) else 0
                except (ValueError, TypeError):
                    val = 0
                if val == 0:
                    continue
                countries.setdefault(str(code_int), {}).setdefault(year, {})[month] = val

    return countries


def extract_country_monthly(filename, sheet="1995-2025-months"):
    df = load(filename, sheet)
    year_row = df.iloc[3, 2:].tolist()
    month_row = df.iloc[4, 2:].tolist()

    # Build (year, month_name) index
    month_names = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ]
    col_info = []
    current_year = None
    for yr, mn in zip(year_row, month_row):
        if pd.notna(yr):
            try:
                current_year = int(float(str(yr).replace("*", "").strip()))
            except (ValueError, TypeError):
                pass
        mn_str = str(mn).strip() if pd.notna(mn) else ""
        if mn_str in month_names and current_year:
            col_info.append((current_year, mn_str))
        else:
            col_info.append(None)

    countries = {}
    for i in range(4, len(df)):
        code = df.iloc[i, 0]
        if pd.isna(code):
            continue
        try:
            code_int = int(float(code))
        except (ValueError, TypeError):
            continue

        monthly = {}
        for j, info in enumerate(col_info):
            if info is None:
                continue
            yr, mn = info
            if yr < 2019:  # only keep recent years
                continue
            v = df.iloc[i, j + 2]
            val = round(float(v), 2) if pd.notna(v) else 0
            if val == 0:
                continue
            monthly.setdefault(yr, {})[mn] = val

        countries[str(code_int)] = monthly

    return countries


def extract_hs4(filename, sheet="2020-2025-years"):
    df = load(filename, sheet)
    years_raw = df.iloc[3, 2:].tolist()
    years = []
    for y in years_raw:
        try:
            years.append(int(float(str(y).replace("*", "").strip())))
        except (ValueError, TypeError):
            years.append(str(y))

    products = []
    for i in range(4, len(df)):
        code = df.iloc[i, 0]
        name = df.iloc[i, 1]
        if pd.isna(code) or pd.isna(name):
            continue
        name_str = str(name).strip()
        if name_str in ("Total Exports", "Total Imports", "of which:"):
            continue
        try:
            code_int = int(float(code))
        except (ValueError, TypeError):
            continue

        vals = []
        for v in df.iloc[i, 2:].tolist():
            try:
                vals.append(round(float(v), 2) if pd.notna(v) else 0)
            except (ValueError, TypeError):
                vals.append(0)

        products.append({
            "c": f"{code_int:04d}",
            "n": name_str,
            "v": vals,
        })

    return years, products


def compute_hs4_ytd(filename):
    """Compute YTD export/import values per HS4 product for the latest
    period and the same period of the previous year.

    Returns dict with year, month, and a list of product dicts, or None.
    """
    import openpyxl

    month_names = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ]
    skip_names = {"Total Exports", "Total Imports", "of which:"}

    wb = openpyxl.load_workbook(DATA_DIR / filename, read_only=True)
    year_sheets = [s for s in wb.sheetnames if s.rstrip("*").strip().isdigit()]
    wb.close()

    if not year_sheets:
        return None

    # --- current-year sheet (e.g. '2026') ---
    df_cur = load(filename, year_sheets[0])
    year_str = str(df_cur.iloc[3, 2]).replace("*", "").strip()
    current_year = int(float(year_str))

    month_row = df_cur.iloc[4, 2:].tolist()
    cur_col_map = {}
    for j, cell in enumerate(month_row):
        m = str(cell).strip() if pd.notna(cell) else ""
        if m in month_names:
            cur_col_map[j] = m

    latest_month_idx = max(month_names.index(m) for m in cur_col_map.values())
    ytd_months = set(month_names[: latest_month_idx + 1])

    products = {}
    for i in range(5, len(df_cur)):
        code = df_cur.iloc[i, 0]
        name = df_cur.iloc[i, 1]
        if pd.isna(code) or pd.isna(name):
            continue
        name_str = str(name).strip()
        if name_str in skip_names:
            continue
        try:
            code_int = int(float(code))
        except (ValueError, TypeError):
            continue

        cur_sum = 0.0
        for j, month in cur_col_map.items():
            v = df_cur.iloc[i, j + 2]
            try:
                val = float(v) if pd.notna(v) else 0
            except (ValueError, TypeError):
                val = 0
            cur_sum += val

        products[f"{code_int:04d}"] = {
            "n": name_str,
            "cur": round(cur_sum, 2),
            "prev": 0.0,
        }

    # --- previous-year same-period from monthly sheet ---
    prev_year = current_year - 1
    monthly_sheet = None
    wb2 = openpyxl.load_workbook(DATA_DIR / filename, read_only=True)
    for s in wb2.sheetnames:
        if "months" in s.lower() and str(prev_year) in s:
            monthly_sheet = s
            break
    wb2.close()

    if monthly_sheet:
        df_m = load(filename, monthly_sheet)
        yr_row = df_m.iloc[3, 2:].tolist()
        mn_row = df_m.iloc[4, 2:].tolist()

        prev_cols = []
        cur_yr = None
        for j_idx, (yr, mn) in enumerate(zip(yr_row, mn_row)):
            if pd.notna(yr):
                try:
                    cur_yr = int(float(str(yr).replace("*", "").strip()))
                except (ValueError, TypeError):
                    pass
            mn_str = str(mn).strip() if pd.notna(mn) else ""
            if cur_yr == prev_year and mn_str in ytd_months:
                prev_cols.append(j_idx)

        for i in range(5, len(df_m)):
            code = df_m.iloc[i, 0]
            name = df_m.iloc[i, 1]
            if pd.isna(code) or pd.isna(name):
                continue
            name_str = str(name).strip()
            if name_str in skip_names:
                continue
            try:
                code_int = int(float(code))
            except (ValueError, TypeError):
                continue

            code_str = f"{code_int:04d}"
            if code_str not in products:
                products[code_str] = {"n": name_str, "cur": 0.0, "prev": 0.0}

            prev_sum = 0.0
            for j_idx in prev_cols:
                v = df_m.iloc[i, j_idx + 2]
                try:
                    val = float(v) if pd.notna(v) else 0
                except (ValueError, TypeError):
                    val = 0
                prev_sum += val

            products[code_str]["prev"] = round(prev_sum, 2)

    return {
        "year": current_year,
        "month": latest_month_idx + 1,
        "products": [
            {"c": k, "n": v["n"], "cur": v["cur"], "prev": v["prev"]}
            for k, v in products.items()
        ],
    }


def extract_fdi_countries(filename, sheet="FDI (annual)"):
    df = pd.read_excel(DATA_DIR / "fdi" / filename, sheet_name=sheet, header=None)
    years_raw = df.iloc[3, 2:].tolist()
    years = []
    for y in years_raw:
        try:
            years.append(int(float(str(y).replace("*", "").strip())))
        except (ValueError, TypeError):
            years.append(str(y))

    countries = {}
    for i in range(4, len(df)):
        code = df.iloc[i, 0]
        name = df.iloc[i, 1]
        if pd.isna(code) or pd.isna(name):
            continue
        name_str = str(name).strip()
        if name_str in ("Total", "of which:", "EU counties (27)"):
            continue
        try:
            code_int = int(float(code))
        except (ValueError, TypeError):
            continue

        vals = []
        for v in df.iloc[i, 2:].tolist():
            try:
                vals.append(round(float(v), 2) if pd.notna(v) else 0)
            except (ValueError, TypeError):
                vals.append(0)

        countries[str(code_int)] = {"name": name_str, "values": vals}

    # Total row
    total_vals = []
    for v in df.iloc[4, 2:].tolist():
        try:
            total_vals.append(round(float(v), 2) if pd.notna(v) else 0)
        except (ValueError, TypeError):
            total_vals.append(0)

    return years, countries, total_vals


def extract_fdi_sectors(filename, sheet="ENG (annual)"):
    df = pd.read_excel(DATA_DIR / "fdi" / filename, sheet_name=sheet, header=None)
    years_raw = df.iloc[3, 1:].tolist()
    years = []
    for y in years_raw:
        try:
            years.append(int(float(str(y).replace("*", "").strip())))
        except (ValueError, TypeError):
            years.append(str(y))

    sectors = []
    for i in range(4, len(df)):
        name = df.iloc[i, 0]
        if pd.isna(name):
            continue
        name_str = str(name).strip()
        if name_str in ("Total", "of which:", "Other"):
            continue

        vals = []
        for v in df.iloc[i, 1:].tolist():
            try:
                vals.append(round(float(v), 2) if pd.notna(v) else 0)
            except (ValueError, TypeError):
                vals.append(0)

        sectors.append({"n": name_str, "v": vals})

    return years, sectors


def extract_fdi_quarterly(filename):
    df = pd.read_excel(DATA_DIR / "fdi" / filename, header=None)
    rows = []
    for i in range(4, len(df)):
        year = df.iloc[i, 0]
        if pd.isna(year):
            continue
        try:
            yr = int(float(year))
        except (ValueError, TypeError):
            continue
        total = round(float(df.iloc[i, 1]), 2) if pd.notna(df.iloc[i, 1]) else 0
        quarters = []
        for c in range(2, 6):
            v = df.iloc[i, c]
            try:
                quarters.append(round(float(v), 2) if pd.notna(v) else 0)
            except (ValueError, TypeError):
                quarters.append(0)
        rows.append({"y": yr, "t": total, "q": quarters})
    return rows


def main():
    OUT_DIR.mkdir(exist_ok=True)

    print("Processing exports by country (yearly)...")
    exp_years, exp_countries = extract_country_yearly("Export-Country_1995-2026.xlsx")

    print("Processing imports by country (yearly)...")
    imp_years, imp_countries = extract_country_yearly("Import-Country-1995-2026.xlsx")

    print("Processing exports by country (monthly)...")
    exp_monthly = extract_country_monthly("Export-Country_1995-2026.xlsx")
    for code, ydata in extract_country_monthly_current_year("Export-Country_1995-2026.xlsx").items():
        for yr, mdata in ydata.items():
            exp_monthly.setdefault(code, {}).setdefault(yr, {}).update(mdata)

    print("Processing imports by country (monthly)...")
    imp_monthly = extract_country_monthly("Import-Country-1995-2026.xlsx")
    for code, ydata in extract_country_monthly_current_year("Import-Country-1995-2026.xlsx").items():
        for yr, mdata in ydata.items():
            imp_monthly.setdefault(code, {}).setdefault(yr, {}).update(mdata)

    print("Processing HS4 exports...")
    hs4_exp_years, hs4_exp = extract_hs4("Export-Product-by-4-digit-2015-2026.xlsx")

    print("Processing HS4 imports...")
    hs4_imp_years, hs4_imp = extract_hs4("Import-Product-by-4-digit-2015-2026.xlsx")

    print("Computing HS4 export YTD...")
    hs4_exp_ytd = compute_hs4_ytd("Export-Product-by-4-digit-2015-2026.xlsx")

    print("Processing FDI by country...")
    fdi_years, fdi_countries, fdi_total = extract_fdi_countries("FDI_Eng-countries.xlsx")

    print("Processing FDI by sector...")
    fdi_sec_years, fdi_sectors = extract_fdi_sectors("FDI_ENG-sectors-NACE-2.xlsx")

    print("Processing FDI quarterly...")
    fdi_quarterly = extract_fdi_quarterly("FDI_by_Quarters_Eng.xlsx")

    # Build country list from export data
    country_list = []
    for code, info in exp_countries.items():
        country_list.append({"c": int(code), "n": info["name"]})
    country_list.sort(key=lambda x: x["n"])

    # Compute aggregate totals across all countries
    n_years = len(exp_years)
    total_exp_yearly = [0.0] * n_years
    for v in exp_countries.values():
        for i, val in enumerate(v["values"]):
            total_exp_yearly[i] += val
    total_exp_yearly = [round(v, 2) for v in total_exp_yearly]

    total_imp_yearly = [0.0] * n_years
    for v in imp_countries.values():
        for i, val in enumerate(v["values"]):
            total_imp_yearly[i] += val
    total_imp_yearly = [round(v, 2) for v in total_imp_yearly]

    # Aggregate monthly totals
    total_exp_monthly = {}
    for cdata in exp_monthly.values():
        for yr, mdata in cdata.items():
            total_exp_monthly.setdefault(yr, {})
            for m, val in mdata.items():
                total_exp_monthly[yr][m] = round(
                    total_exp_monthly[yr].get(m, 0) + val, 2
                )

    total_imp_monthly = {}
    for cdata in imp_monthly.values():
        for yr, mdata in cdata.items():
            total_imp_monthly.setdefault(yr, {})
            for m, val in mdata.items():
                total_imp_monthly[yr][m] = round(
                    total_imp_monthly[yr].get(m, 0) + val, 2
                )

    trade = {
        "countries": country_list,
        "years": exp_years,
        "total_exp_yearly": total_exp_yearly,
        "total_imp_yearly": total_imp_yearly,
        "total_exp_monthly": total_exp_monthly,
        "total_imp_monthly": total_imp_monthly,
        "export": {k: v["values"] for k, v in exp_countries.items()},
        "import": {k: v["values"] for k, v in imp_countries.items()},
        "exp_monthly": exp_monthly,
        "imp_monthly": imp_monthly,
        "hs4_years": hs4_exp_years,
        "hs4_export": hs4_exp,
        "hs4_import": hs4_imp,
        "hs4_exp_ytd": hs4_exp_ytd,
        "fdi_years": fdi_years,
        "fdi_countries": {k: v["values"] for k, v in fdi_countries.items()},
        "fdi_country_names": {k: v["name"] for k, v in fdi_countries.items()},
        "fdi_total": fdi_total,
        "fdi_sec_years": fdi_sec_years,
        "fdi_sectors": fdi_sectors,
        "fdi_quarterly": fdi_quarterly,
    }

    out_path = OUT_DIR / "trade.json"
    with open(out_path, "w") as f:
        json.dump(trade, f, separators=(",", ":"))

    size_mb = out_path.stat().st_size / (1024 * 1024)
    print(f"Done! {out_path} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
