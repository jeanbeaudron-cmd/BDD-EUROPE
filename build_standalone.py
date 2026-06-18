#!/usr/bin/env python3
"""
build_standalone.py — bundle index.html + category_db.parquet into ONE
self-contained .html file that opens by double-click (no server, no install).

The parquet is embedded as base64 and read in-browser via registerFileBuffer,
so there is no local fetch (which file:// blocks). DuckDB-WASM itself still
loads from jsDelivr, so the first open needs an internet connection.

Usage:
  python3 build_standalone.py [index.html] [category_db.parquet] [output.html]
Default output: category_explorer_standalone.html
"""
import base64, sys
from pathlib import Path

html_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("index.html")
pq_path   = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("category_db.parquet")
out_path  = Path(sys.argv[3]) if len(sys.argv) > 3 else Path("category_explorer_standalone.html")

if not html_path.exists(): sys.exit(f"missing {html_path}")
if not pq_path.exists():   sys.exit(f"missing {pq_path} — run build_category_db.py first")

html = html_path.read_text(encoding="utf-8")
b64 = base64.b64encode(pq_path.read_bytes()).decode("ascii")

inject = f'<script>window.EMBEDDED_PARQUET_B64="{b64}";</script>\n'
# place the data just before the app module so it is defined when boot() runs
marker = '<script type="module">'
if marker not in html:
    sys.exit("could not find the module script tag in index.html")
html = html.replace(marker, inject + marker, 1)

out_path.write_text(html, encoding="utf-8")
mb = out_path.stat().st_size / 1e6
print(f"OK -> {out_path}  ({mb:.1f} MB)  — double-click to open.")
