#!/bin/bash
# Monthly data refresh script for GeoStat trade & FDI statistics.
# Add to crontab:  0 6 5 * * /path/to/refresh.sh
#
# The script downloads fresh XLSX files, rebuilds data/trade.json,
# and the static site picks up the changes automatically.

set -e
cd "$(dirname "$0")"

echo "[$(date)] Starting GeoStat data refresh..."

python3 geostat_scraper.py download \
  export_by_country import_by_country \
  export_hs4_2015 import_hs4_2015 \
  fdi_by_country fdi_by_sector fdi_quarterly \
  --force

python3 build_data.py

echo "[$(date)] Done."
