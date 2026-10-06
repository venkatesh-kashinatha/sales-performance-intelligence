#!/usr/bin/env bash
# Weekly refresh for macOS/Linux. Builds output/Sales_Pack_<date>.xlsx and the Tableau feed.
# Schedule it with cron (see README, step 6). Optional first argument: path to the data file.
set -euo pipefail
cd "$(dirname "$0")/.."
INPUT="${1:-data/raw/superstore.csv}"
source .venv/bin/activate
python -m sales_pack --input "$INPUT" --out output
