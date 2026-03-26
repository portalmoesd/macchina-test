#!/usr/bin/env python3
"""
Fetch country×product export data from the GeoStat External Trade API.

API base: https://ex-trade-api.geostat.ge/api/trade
Endpoints used:
  GET  /classificatory?lang=en  → countries, HS4 codes, latest period
  POST /get_data                → trade data by filters

Outputs: data/country_products.json
"""

import csv
import json
import sys
import time
import logging
from pathlib import Path

import requests

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

API_BASE = "https://ex-trade-api.geostat.ge/api/trade"
OUT_DIR = Path("data")
MAX_RETRIES = 4

SESSION = requests.Session()
SESSION.headers.update({
    "Content-Type": "application/json",
    "User-Agent": "GeoStat-Trade-Scraper/1.0 (research)",
})

# Load HS4 short titles
HS4_SHORT = {}
_csv_path = Path(__file__).parent / "hs4_short_titles.csv"
if _csv_path.exists():
    with open(_csv_path, encoding="utf-8") as f:
        for row in csv.reader(f):
            if row and row[0].strip().isdigit():
                HS4_SHORT[row[0].strip().zfill(4)] = row[1].strip()


def _retry_request(method, url, retries=MAX_RETRIES, **kwargs):
    """Request with exponential backoff."""
    for attempt in range(1, retries + 1):
        try:
            resp = SESSION.request(method, url, timeout=60, **kwargs)
            resp.raise_for_status()
            return resp.json()
        except requests.RequestException as exc:
            if attempt == retries:
                raise
            wait = 2 ** attempt
            log.warning("Attempt %d failed (%s). Retrying in %ds...", attempt, exc, wait)
            time.sleep(wait)


def get_classificatory():
    """Fetch filter options (countries, HS4, latest period)."""
    log.info("Fetching classificatory data...")
    data = _retry_request("GET", f"{API_BASE}/classificatory?lang=en")
    return data.get("data", data)


def get_trade_data(filters):
    """Fetch trade data with given filters."""
    filters["locale"] = "en"
    return _retry_request("POST", f"{API_BASE}/get_data", json=filters)


def fetch_country_products():
    """Fetch export data by country and HS4 product for the latest YTD period.

    Strategy:
      1. Get classificatory data to find latest year/month and all country codes
      2. For each country, request export (tradeFlow=10) and domestic export
         (tradeFlow=12) data grouped by HS4, with sum=true for YTD
      3. Compute re-export share = (total - domestic) / total * 100
    """
    classif = get_classificatory()

    latest_year = classif["selected"]["year"]
    latest_month = classif["selected"]["month"]
    countries = classif["countries"]
    all_months = [m["value"] for m in classif["month"] if m["value"] <= latest_month]

    log.info("Latest period: %d months 1-%d", latest_year, latest_month)
    log.info("Countries: %d", len(countries))

    # tradeFlow IDs from the API:
    # Look at types to identify export and domestic export
    types = classif.get("types", [])
    log.info("Trade flow types: %s", [(t["value"], t["label"]) for t in types])

    # We need:
    # - Total export (typically tradeFlow=10)
    # - Domestic export (typically tradeFlow=12) for re-export calculation
    # The exact IDs depend on the API response; 10 is the default

    by_country = {}
    total = len(countries)

    for idx, country in enumerate(countries):
        cc = country["value"]
        cc_str = str(cc)
        name = country.get("label", cc_str)

        if cc_str == "global":
            continue

        log.info("[%d/%d] Fetching exports for %s (%s)...", idx + 1, total, name, cc_str)

        # Fetch total export by HS4 for this country, YTD
        try:
            filters_total = {
                "tradeFlow": 10,  # Total export
                "measurementUnits": [1],  # Value in USD
                "years": [latest_year],
                "months": all_months,
                "countries": [cc],
                "hs4": [],  # all products
                "grouping": ["hs"],
                "sum": True,
                "page": 1,
                "pageSize": 5000,
            }
            result_total = get_trade_data(filters_total)
        except requests.RequestException as exc:
            log.warning("Failed for %s (total): %s", name, exc)
            continue

        # Fetch domestic export by HS4 for this country, YTD
        try:
            filters_domestic = {
                "tradeFlow": 12,  # Domestic export
                "measurementUnits": [1],
                "years": [latest_year],
                "months": all_months,
                "countries": [cc],
                "hs4": [],
                "grouping": ["hs"],
                "sum": True,
                "page": 1,
                "pageSize": 5000,
            }
            result_domestic = get_trade_data(filters_domestic)
        except requests.RequestException:
            result_domestic = {"data": []}

        # Also fetch previous year same period for change calculation
        try:
            filters_prev = {
                "tradeFlow": 10,
                "measurementUnits": [1],
                "years": [latest_year - 1],
                "months": all_months,
                "countries": [cc],
                "hs4": [],
                "grouping": ["hs"],
                "sum": True,
                "page": 1,
                "pageSize": 5000,
            }
            result_prev = get_trade_data(filters_prev)
        except requests.RequestException:
            result_prev = {"data": []}

        # Parse results into product dicts
        # The API response structure may vary; adapt based on actual data
        total_data = result_total.get("data", [])
        domestic_data = result_domestic.get("data", [])
        prev_data = result_prev.get("data", [])

        if not total_data:
            continue

        # Build lookup: hs4_code -> value
        def extract_products(data_rows):
            products = {}
            for row in data_rows:
                # Row might be a dict or list depending on API response
                if isinstance(row, dict):
                    hs_code = str(row.get("hs4", row.get("code", ""))).zfill(4)
                    value = row.get("value", row.get("val", 0))
                elif isinstance(row, list):
                    # Try to find HS code and value in list
                    hs_code = str(row[0]).zfill(4) if len(row) > 0 else ""
                    value = row[-1] if len(row) > 1 else 0
                else:
                    continue
                try:
                    products[hs_code] = float(value) if value else 0
                except (ValueError, TypeError):
                    products[hs_code] = 0
            return products

        total_prods = extract_products(total_data)
        domestic_prods = extract_products(domestic_data)
        prev_prods = extract_products(prev_data)

        if not total_prods:
            continue

        products = []
        for hs_code, cur_val in total_prods.items():
            if cur_val == 0:
                continue
            prev_val = prev_prods.get(hs_code, 0)
            dom_val = domestic_prods.get(hs_code, 0)
            re_val = cur_val - dom_val
            re_share = round(re_val / cur_val * 100, 1) if cur_val > 0 else 0
            re_share = max(0, re_share)

            name = HS4_SHORT.get(hs_code, hs_code)
            products.append({
                "c": hs_code,
                "n": name,
                "cur": round(cur_val, 2),
                "prev": round(prev_val, 2),
                "re": re_share,
            })

        products.sort(key=lambda p: -p["cur"])
        by_country[cc_str] = products

        # Be polite — small delay between countries
        time.sleep(0.3)

    return by_country


def main():
    OUT_DIR.mkdir(exist_ok=True)

    log.info("Starting country×product export data fetch...")
    data = fetch_country_products()

    out_path = OUT_DIR / "country_products.json"
    with open(out_path, "w") as f:
        json.dump(data, f, separators=(",", ":"))

    n_countries = len(data)
    n_products = sum(len(v) for v in data.values())
    size_kb = out_path.stat().st_size / 1024
    log.info("Done! %s (%d countries, %d products, %.1f KB)", out_path, n_countries, n_products, size_kb)


if __name__ == "__main__":
    main()
