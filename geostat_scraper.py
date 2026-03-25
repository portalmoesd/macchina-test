#!/usr/bin/env python3
"""
GeoStat External Trade Data Scraper

Automatically downloads trade statistics (export/import) from the
National Statistics Office of Georgia (geostat.ge).

Data source: https://ex-trade.geostat.ge/
"""

import os
import sys
import time
import argparse
import logging
from pathlib import Path
from urllib.parse import quote

import requests
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Dataset catalogue – direct XLSX URLs from geostat.ge
# ---------------------------------------------------------------------------

DATASETS = {
    # ── Exports ────────────────────────────────────────────────────────────
    "export_by_country": {
        "url": "https://geostat.ge/media/77619/Export-Country_1995-2026.xlsx",
        "description": "Exports by Countries (1995-2026)",
    },
    "export_by_country_group": {
        "url": "https://geostat.ge/media/77620/Export-_Country_Group-1995-2026.xlsx",
        "description": "Exports by Country Groups (1995-2026)",
    },
    "export_hs4_2015": {
        "url": "https://geostat.ge/media/77621/Export-Product-by-4-digit-2015-2026.xlsx",
        "description": "Exports by HS 4-digit (2015-2026)",
    },
    "domestic_export_hs4": {
        "url": "https://geostat.ge/media/77607/Domestic-Exports_by-4-digit-2014-2026.xlsx",
        "description": "Domestic exports by HS 4-digit (2014-2026)",
    },
    "export_hs4_2000": {
        "url": "https://geostat.ge/media/51582/Export-Product-by-4-digit-2000-2014.xlsx",
        "description": "Exports by HS 4-digit (2000-2014)",
    },
    "export_hs4_1995": {
        "url": "https://geostat.ge/media/42913/Export-Product-by-4-digit-1995-1999.xlsx",
        "description": "Exports by HS 4-digit (1995-1999)",
    },
    "export_hs6_2020": {
        "url": "https://geostat.ge/media/77622/Export-Product-by-6-digit-2020-2026.xlsx",
        "description": "Exports by HS 6-digit (2020-2026)",
    },
    "export_hs6_2015": {
        "url": "https://geostat.ge/media/76620/Export-Product-by-6-digit-2015-2019.xlsx",
        "description": "Exports by HS 6-digit (2015-2019)",
    },
    "export_hs6_2000": {
        "url": "https://geostat.ge/media/69436/Export-Product-by-6-digit-2000-2014.xlsx",
        "description": "Exports by HS 6-digit (2000-2014)",
    },
    "export_bec1": {
        "url": "https://geostat.ge/media/77623/Export-BEC_2000-2026.xlsx",
        "description": "Exports by BEC 1-digit (2000-2026)",
    },
    "export_bec3": {
        "url": "https://geostat.ge/media/77624/Export-BEC-3-digit_2000-2026.xlsx",
        "description": "Exports by BEC 3-digit (2000-2026)",
    },
    "export_sitc": {
        "url": "https://geostat.ge/media/77625/Export-SITC_2000-2026.xlsx",
        "description": "Exports by SITC Section (2000-2026)",
    },
    "export_sitc5_2009": {
        "url": "https://geostat.ge/media/77626/Export-SITC-5-digit_2009-2026.xlsx",
        "description": "Exports by SITC 5-digit (2009-2026)",
    },
    "export_sitc5_2000": {
        "url": "https://geostat.ge/media/66431/Export-SITC-5digit-2000-2008.xlsx",
        "description": "Exports by SITC 5-digit (2000-2008)",
    },
    "export_by_region": {
        "url": "https://geostat.ge/media/77084/export-by-regions-and-SITC-4-digital.xlsx",
        "description": "Exports by Regions and SITC Subgroups",
    },
    "export_transport": {
        "url": "https://geostat.ge/media/77627/Export-Transports-2016_2026.xlsx",
        "description": "Exports by Mode of Transport (2016-2026)",
    },
    "export_nace": {
        "url": "https://geostat.ge/media/77628/Export-NACE-2015_2026.xlsx",
        "description": "Exports by Economic Activities NACE Rev.2 (2015-2026)",
    },
    "export_size": {
        "url": "https://geostat.ge/media/77629/Export-Size-2015_2026.xlsx",
        "description": "Export by Size Classes of Traders (2015-2026)",
    },
    # ── Imports ────────────────────────────────────────────────────────────
    "import_by_country": {
        "url": "https://geostat.ge/media/77641/Import-Country-1995-2026.xlsx",
        "description": "Imports by Countries (1995-2026)",
    },
    "import_by_country_group": {
        "url": "https://geostat.ge/media/77642/Import_Country_Group-1995-2026.xlsx",
        "description": "Imports by Country Groups (1995-2026)",
    },
    "import_hs4_2015": {
        "url": "https://geostat.ge/media/77643/Import-Product-by-4-digit-2015-2026.xlsx",
        "description": "Imports by HS 4-digit (2015-2026)",
    },
    "import_hs4_2000": {
        "url": "https://geostat.ge/media/76630/Import-Product-by-4-digit-2000-2014.xlsx",
        "description": "Imports by HS 4-digit (2000-2014)",
    },
    "import_hs4_1995": {
        "url": "https://geostat.ge/media/72200/Import-products--1995-1999_eng.xlsx",
        "description": "Imports by HS 4-digit (1995-1999)",
    },
    "import_hs6_2020": {
        "url": "https://geostat.ge/media/77644/Import-Product-by-6-digit-2020-2026.xlsx",
        "description": "Imports by HS 6-digit (2020-2026)",
    },
    "import_hs6_2015": {
        "url": "https://geostat.ge/media/76632/Import-Product-by-6-digit-2015-2019.xlsx",
        "description": "Imports by HS 6-digit (2015-2019)",
    },
    "import_hs6_2000": {
        "url": "https://geostat.ge/media/76633/Import-Product-by-6-digit-2000-2014.xlsx",
        "description": "Imports by HS 6-digit (2000-2014)",
    },
    "import_bec1": {
        "url": "https://geostat.ge/media/77645/Import-BEC-2000-2026.xlsx",
        "description": "Imports by BEC 1-digit (2000-2026)",
    },
    "import_bec3": {
        "url": "https://geostat.ge/media/77646/Import-BEC-3-digit_2000-2026.xlsx",
        "description": "Imports by BEC 3-digit (2000-2026)",
    },
    "import_sitc": {
        "url": "https://geostat.ge/media/77647/Import-SITC_2000-2026.xlsx",
        "description": "Imports by SITC Section (2000-2026)",
    },
    "import_sitc5_2009": {
        "url": "https://geostat.ge/media/77648/Import-SITC-5-digit_2009-2026.xlsx",
        "description": "Imports by SITC 5-digit (2009-2026)",
    },
    "import_sitc5_2000": {
        "url": "https://geostat.ge/media/49867/Import-SITC-5digit-2000-2008.xlsx",
        "description": "Imports by SITC 5-digit (2000-2008)",
    },
    "import_by_region": {
        "url": "https://geostat.ge/media/77086/import-by-regions-and-SITC-4-digital.xlsx",
        "description": "Imports by Regions and SITC Subgroups",
    },
    "import_transport": {
        "url": "https://geostat.ge/media/77649/Import-Transports-2016_2026.xlsx",
        "description": "Imports by Mode of Transport (2016-2026)",
    },
    "import_nace": {
        "url": "https://geostat.ge/media/77650/Import-NACE-2015_2026.xlsx",
        "description": "Imports by Economic Activities NACE Rev.2 (2015-2026)",
    },
    "import_size": {
        "url": "https://geostat.ge/media/77651/Import-Size-2015_2026.xlsx",
        "description": "Import by Size Classes of Traders (2015-2026)",
    },
    # ── FDI (Foreign Direct Investment) ────────────────────────────────────
    "fdi_by_country": {
        "url": "https://geostat.ge/media/77518/FDI_Eng-countries.xlsx",
        "description": "FDI by Countries (English)",
        "subdir": "fdi",
    },
    "fdi_by_sector": {
        "url": "https://geostat.ge/media/77519/FDI_ENG-sectors-NACE-2.xlsx",
        "description": "FDI by Economic Sectors NACE Rev.2 (English)",
        "subdir": "fdi",
    },
    "fdi_quarterly": {
        "url": "https://geostat.ge/media/77525/FDI_by_Quarters_Eng.xlsx",
        "description": "FDI by Quarters (English)",
        "subdir": "fdi",
    },
}

# GeoStat also exposes an XLSX→CSV conversion endpoint
CSV_API = "https://autoapi.geostat.ge/api/v1/xlsx-to-csv/download-all"

DEFAULT_OUTPUT_DIR = "geostat_data"
MAX_RETRIES = 4
SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "GeoStat-Scraper/1.0 (research)"})


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _retry_download(url: str, retries: int = MAX_RETRIES) -> requests.Response:
    """Download with exponential back-off on transient errors."""
    for attempt in range(1, retries + 1):
        try:
            resp = SESSION.get(url, timeout=120)
            resp.raise_for_status()
            return resp
        except requests.RequestException as exc:
            if attempt == retries:
                raise
            wait = 2 ** attempt
            log.warning("Attempt %d failed (%s). Retrying in %ds…", attempt, exc, wait)
            time.sleep(wait)


def download_xlsx(key: str, out_dir: Path, force: bool = False) -> Path:
    """Download a single XLSX file and return the local path."""
    meta = DATASETS[key]
    url = meta["url"]
    filename = url.rsplit("/", 1)[-1]
    target_dir = out_dir / meta.get("subdir", "")
    target_dir.mkdir(parents=True, exist_ok=True)
    dest = target_dir / filename

    if dest.exists() and not force:
        log.info("Already downloaded: %s", dest)
        return dest

    log.info("Downloading %s → %s", meta["description"], dest)
    resp = _retry_download(url)
    dest.write_bytes(resp.content)
    log.info("  saved %s (%.1f KB)", dest.name, len(resp.content) / 1024)
    return dest


def download_csv(key: str, out_dir: Path) -> Path:
    """Download CSV via GeoStat autoapi conversion endpoint."""
    meta = DATASETS[key]
    xlsx_url = meta["url"]
    csv_url = f"{CSV_API}?url={quote(xlsx_url, safe='')}"
    filename = xlsx_url.rsplit("/", 1)[-1].replace(".xlsx", ".csv")
    dest = out_dir / filename

    if dest.exists():
        log.info("Already downloaded: %s", dest)
        return dest

    log.info("Downloading CSV for %s → %s", meta["description"], dest)
    resp = _retry_download(csv_url)
    dest.write_bytes(resp.content)
    log.info("  saved %s (%.1f KB)", dest.name, len(resp.content) / 1024)
    return dest


def xlsx_to_dataframe(path: Path, sheet: int | str = 0) -> pd.DataFrame:
    """Read a downloaded XLSX into a pandas DataFrame."""
    log.info("Reading %s (sheet=%s)", path.name, sheet)
    return pd.read_excel(path, sheet_name=sheet)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def list_datasets():
    """Print the catalogue of available datasets."""
    print(f"\n{'Key':<30} {'Description'}")
    print("-" * 80)
    for key, meta in DATASETS.items():
        print(f"{key:<30} {meta['description']}")
    print()


def main():
    parser = argparse.ArgumentParser(
        description="Download trade statistics from GeoStat Georgia (geostat.ge)."
    )
    sub = parser.add_subparsers(dest="command")

    # list
    sub.add_parser("list", help="List available datasets")

    # download
    dl = sub.add_parser("download", help="Download one or more datasets")
    dl.add_argument(
        "datasets",
        nargs="*",
        default=["all"],
        help="Dataset key(s) to download, or 'all' (default: all)",
    )
    dl.add_argument(
        "--format",
        choices=["xlsx", "csv", "both"],
        default="xlsx",
        help="Download format (default: xlsx)",
    )
    dl.add_argument(
        "-o", "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help=f"Output directory (default: {DEFAULT_OUTPUT_DIR})",
    )
    dl.add_argument(
        "--exports-only",
        action="store_true",
        help="Only download export datasets",
    )
    dl.add_argument(
        "--imports-only",
        action="store_true",
        help="Only download import datasets",
    )
    dl.add_argument(
        "--force",
        action="store_true",
        help="Re-download even if files already exist",
    )

    # preview
    pv = sub.add_parser("preview", help="Download and preview a dataset")
    pv.add_argument("dataset", help="Dataset key to preview")
    pv.add_argument("--rows", type=int, default=10, help="Number of rows (default: 10)")

    args = parser.parse_args()

    if args.command is None or args.command == "list":
        list_datasets()
        return

    if args.command == "download":
        out_dir = Path(args.output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        # resolve which keys to download
        if "all" in args.datasets:
            keys = list(DATASETS.keys())
        else:
            keys = args.datasets
            for k in keys:
                if k not in DATASETS:
                    log.error("Unknown dataset: %s (use 'list' to see options)", k)
                    sys.exit(1)

        if args.exports_only:
            keys = [k for k in keys if k.startswith("export_")]
        elif args.imports_only:
            keys = [k for k in keys if k.startswith("import_")]

        for key in keys:
            try:
                if args.format in ("xlsx", "both"):
                    download_xlsx(key, out_dir, force=args.force)
                if args.format in ("csv", "both"):
                    download_csv(key, out_dir)
            except requests.RequestException as exc:
                log.error("Failed to download %s: %s", key, exc)

        log.info("Done. Files saved to %s/", out_dir)

    elif args.command == "preview":
        if args.dataset not in DATASETS:
            log.error("Unknown dataset: %s", args.dataset)
            sys.exit(1)

        out_dir = Path(DEFAULT_OUTPUT_DIR)
        out_dir.mkdir(parents=True, exist_ok=True)
        path = download_xlsx(args.dataset, out_dir)
        df = xlsx_to_dataframe(path)
        print(f"\n── {DATASETS[args.dataset]['description']} ──")
        print(f"Shape: {df.shape[0]} rows × {df.shape[1]} columns")
        print(f"Columns: {list(df.columns)}\n")
        print(df.head(args.rows).to_string(index=False))
        print()


if __name__ == "__main__":
    main()
