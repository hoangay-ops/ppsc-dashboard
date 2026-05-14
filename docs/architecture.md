# Architecture Summary

## Overview

This is a **build-time static site generator**, not a web server. There is no runtime backend, no Flask, and no database. All data processing happens once when `automate.py` runs; the output is plain HTML files served by GitHub Pages.

## Data Flow

```
data/*.xlsx  ──┐
data/*.docx  ──┤
data/*hours* ──┤──▶ automate.py ──▶ index.html + archive/*.html
data/*points*──┘         │
                         └──▶ output/dashboard_data.json
                         └──▶ archive/manifest.json
```

## Key Components

### `automate.py` — Build Pipeline

Runs top-to-bottom as a single `process()` function. Logical stages:

1. **File discovery** — scans `data/` for `.xlsx`, separates main workbooks from auxiliary ones using filename heuristics (`is_aux()`)
2. **Workbook parsing** — `read_main_workbook()` → `parse_milestones()` extracts one `MilestoneRow` per "Total Milestone" block; column positions are hardcoded constants (`COL_*`, `MONTH_TO_COL`)
3. **Status classification** — `derive_status()` assigns On Track / Watch / At Risk / Complete based on estimated balance, burn rate, and override dicts
4. **Auxiliary parsing** — `parse_hours_file()`, `parse_stories_file()`, `parse_narrative_doc()` each run independently and enrich `MilestoneRow` objects in place
5. **Chart data assembly** — builds `monthly_chart` list combining actuals + forecast values per month
6. **HTML generation** — `generate_html()` runs **twice**: first pass writes `index.html`, second pass re-runs after archive is written so nav prev/next links can be baked in
7. **Archive management** — copies `index.html` to `archive/{slug}.html`, rebuilds `archive/index.html` and `manifest.json`
8. **Git push** — stages specific paths and force-pushes to `main`

### `dashboard_template.html` — Frontend Shell

A single self-contained HTML file using `__PLACEHOLDER__` tokens replaced by `automate.py` at build time. Contains:

- All CSS (custom variables, grid layout, responsive breakpoints, print styles)
- Inline JS for: milestone table rendering, bar chart (falls back to a div-based chart if Chart.js is unavailable), popover tooltips, column sorting, and the archive site-nav fetch
- Chart.js and html2pdf.js are **fetched from cdnjs at build time** and inlined into the output — the published HTML is fully offline-capable

### `archive/manifest.json` — Navigation State

An ordered JSON array of `{label, file}` objects. The site-nav script in the template `fetch()`es this at page-load time to build the ← Prev / Next → links dynamically. **This file must exist and be correctly ordered for navigation to work.**

## Critical Dependencies and Coupling

| Dependency | Risk if broken |
|---|---|
| Column constants (`COL_TEXT=1`, `COL_BUDGET=3`, etc.) | Wrong column → silent data corruption across all milestones |
| `MONTH_TO_COL` mapping | Month actuals read from wrong columns |
| `__PLACEHOLDER__` token names | Unreplaced tokens appear verbatim in published HTML |
| cdnjs CDN availability at build time | Chart and PDF export silently degrade (warnings printed, empty strings used) |
| `archive/manifest.json` sort order | Nav links point to wrong months |
| Two-pass `generate_html()` | First-pass archive snapshot lacks nav links (fixed by second overwrite) |

## What This Is Not

- Not Flask / Django / any web framework
- Not Bootstrap (all styles are custom CSS variables)
- Not a database-backed application
- Not a real-time dashboard — data is frozen at build time