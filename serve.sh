#!/usr/bin/env bash
# serve.sh — static server for the Category Review Explorer.
# DuckDB-WASM fetches category_db.parquet, which is blocked under file:// .
set -euo pipefail
PORT="${1:-8000}"
cd "$(dirname "$0")"
echo "Serving $(pwd) at http://localhost:${PORT}  (Ctrl-C to stop)"
python3 -m http.server "${PORT}"
