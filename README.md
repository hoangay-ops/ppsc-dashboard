# PPSC Executive Dashboard

A static-site generator that produces a self-contained financial dashboard from Excel workbooks and optional auxiliary files. The output is a single `index.html` (plus versioned archive snapshots) deployed to GitHub Pages.

## What It Tracks

- Monthly spend vs. forecast (actuals bar chart with dotted reference line)
- Milestone budget health: On Track / Watch / At Risk classification
- Labor hours and LCAT breakdown (from auxiliary hours workbook)
- Story points and velocity (from auxiliary points workbook)
- Narrative highlights and credit/adjustment callouts (from `.docx`)

## Stack

| Layer | Technology |
|---|---|
| Data pipeline | Python 3 (`automate.py`) |
| Spreadsheet parsing | `pandas` + `openpyxl` |
| Word doc parsing | `python-docx` (optional) |
| Frontend | Vanilla JS + custom CSS (no framework) |
| Charts | Chart.js 4 (fetched from cdnjs at build time) |
| PDF export | html2pdf.js (fetched from cdnjs at build time) |
| Hosting | GitHub Pages (static) |

## File Layout

```
/
├── automate.py              # Single entry point — run this to build
├── dashboard_template.html  # HTML/CSS/JS shell with __PLACEHOLDER__ tokens
├── data/
│   ├── *.xlsx               # Main financial workbook(s) — one per reporting month
│   ├── *hours*.xlsx         # (Optional) Hours by LCAT auxiliary workbook
│   ├── *points*.xlsx        # (Optional) Story points by milestone
│   └── *.docx               # (Optional) Narrative/cost justification document
├── archive/
│   ├── *.html               # One snapshot per published month
│   ├── index.html           # Archive browser page (auto-generated)
│   └── manifest.json        # Ordered month list consumed by site nav JS
├── output/
│   └── dashboard_data.json  # Machine-readable summary of the latest build
└── index.html               # Latest dashboard (overwritten each build)
```

## How to Run

```bash
python3 automate.py
```

The script will:
1. Parse all `.xlsx` files in `data/`
2. Classify milestones and compute budget health
3. Fetch Chart.js and html2pdf.js from cdnjs
4. Write `index.html` (two passes — second pass bakes in nav links)
5. Write an archive snapshot under `archive/`
6. Rebuild `archive/index.html` and `archive/manifest.json`
7. Git-add, commit, and push to `main` (if `AUTO_PUSH_GITHUB = True`)

## Configuration (top of `automate.py`)

| Variable | Purpose |
|---|---|
| `SITE_BASE_URL` | Public GitHub Pages root URL |
| `ARCHIVE_BASE_URL` | URL prefix for archive snapshots |
| `AUTO_PUSH_GITHUB` | Set `False` to skip the git push |
| `FORECAST_DEFAULTS` | Fallback monthly forecast values when no workbook covers a month |
| `FRONT_LOADED_MILESTONES` | Milestone IDs exempt from At Risk classification |
| `COMPLETED_MILESTONES` | Milestone IDs forced to Complete status |

## Award Year Reference

Award Year 2 runs **Feb 1 2026 – Jan 31 2027**. Month position is defined in `AWARD_MONTH_POSITION` (Feb = 1 … Jan = 12). All burn-rate and pace calculations are relative to this 12-month window.