# Category Review Explorer — Yoghurts Europe

Single-page, backend-free explorer for multi-country Nielsen category data
(FR / UK / DE / ES / IT × YALG / Kefir / Bifidus × Total / Hyper). Built on
**DuckDB-WASM** querying a canonical Parquet in the browser (Architecture
Option B). See [`PRD.md`](PRD.md) and [`TECHNICAL.md`](TECHNICAL.md) for the spec.

## Pipeline

```
Stage 1 (Python, offline)            Stage 2 (browser, single page)
build_category_db.py  ───────────▶   category_db.parquet
   (reads 5 Excel,                        │
    normalises to long)                   ▼
                                     index.html  (DuckDB-WASM + query builder
                                                  + dashboards + coverage badge
                                                  + volume guardrail)
```

## Quick start

```bash
# 1. Stage 1 — build the canonical Parquet from the 5 Nielsen Excel files
python3 build_category_db.py /path/to/uploads category_db.parquet

# 2. Stage 2 — serve the explorer (file:// is blocked for the parquet fetch)
./serve.sh 8000          # then open http://localhost:8000
```

`build_category_db.py` needs `pandas`, `openpyxl`, `pyarrow`. The browser app
needs only a static server and network access to jsDelivr for the pinned
`@duckdb/duckdb-wasm@1.29.0` bundle (vendor it locally for offline/prod use).

## Files

| File | Role |
|---|---|
| `build_category_db.py` | Stage 1 — Excel → canonical long-format Parquet |
| `index.html`           | Stage 2 — UI + DuckDB-WASM + query builder + dashboards |
| `harmonize_flavors.py` | Stage 2c — LLM-assisted, human-validated flavor harmonization |
| `serve.sh`             | `python3 -m http.server` wrapper |
| `category_db.parquet`  | Stage 1 output (not committed — regenerate from source) |

## Engine rules (non-negotiable — PRD §4 / TECHNICAL §3-6)

- **Two aggregation regimes.** Named nodes (family/category/manufacturer/brand/
  segment) are *read* as pre-calculated lines, never summed over children
  (MDD-safe). Attribute cross-cuts (flavor/weight/multipack/bio) are *summed at
  the attribute's own level*.
- **Coverage badge, per country.** Every cross-cut shows coverage =
  `cross-cut sum / manufacturer-level market total`, computed per country and at
  the cross-cut's own level. Never aggregated to a single Europe figure. Below
  100 % → *"partial — excludes MDD / undetailed"*.
- **Additivity by variable.** `additive` (value EUR, units, volume) → summed;
  `derived` (shares, prices) → recomputed from additives, never averaged;
  `node_only` (DN/DV/ROS, YoY, deltas) → shown only for a single node.
- **Volume guardrail.** The four volume bases (EQ/KGS/GESAMT/ALL) are
  incompatible. Volume is restricted to countries sharing a basis (KGS: UK+ES)
  and **blocked** in any Europe / mixed-basis scope; value or units offered
  instead.

`buildQuery(selections) → { sql, needsCoverage, volumeBlocked, coverageSql }`
is a pure function (exposed on `window.buildQuery` for console testing).

## Out of scope (v1)

Read-only (no data editing); no common Europe volume base (4 incompatible bases
— a supplier decision); EAN detail not 100 % for FR/ES/UK (the coverage badge
flags it).
