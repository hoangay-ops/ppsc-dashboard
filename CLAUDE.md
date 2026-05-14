# CLAUDE.md — Rules for AI Assistance on This Repository

## What This Project Actually Is

**Static site generator, not a web app.**
- Entry point: `automate.py` (run directly with `python3 automate.py`)
- No Flask, no Django, no `app.py`, no `/templates` directory, no Bootstrap
- Output: plain `.html` files deployed to GitHub Pages
- All data processing is offline, build-time only

Do not suggest Flask routes, Jinja2 templates, Bootstrap classes, or server-side session handling. They do not exist here.

---

## Before Writing Any Code

1. **Identify which file(s) are affected** — usually `automate.py`, `dashboard_template.html`, or both
2. **State the risk level** of the change (see Risky Areas below)
3. **Explain the implementation plan** before generating code
4. **Confirm column constants** — if the change touches Excel parsing, restate which `COL_*` constants are in play

---

## Risky Areas — Extra Caution Required

### Column Constants (highest risk)
`COL_TEXT`, `COL_PCT`, `COL_BUDGET`, `COL_TOTAL_ACT`, `COL_TOTAL_FCST`, `COL_EST_SPEND`, `COL_EST_BAL`, and `MONTH_TO_COL` are all 0-based hardcoded column indices. **If the client's workbook format changes even one column, every downstream value is silently wrong.** Never assume these are correct without the user confirming the current workbook layout.

### `__PLACEHOLDER__` Tokens
All template substitutions use double-underscore tokens (e.g. `__MONTH__`, `__MILESTONES_JSON__`). If you add a new placeholder in the template, you **must** add the corresponding replacement in `generate_html()`'s `replacements` dict. An unreplaced token appears verbatim in the published page.

### Two-Pass HTML Generation
`generate_html()` is called **twice** in `process()`. The first pass writes the archive snapshot without nav links; the second pass bakes them in and overwrites both `index.html` and the archive snapshot. Any change to `generate_html()` must work correctly in both passes.

### `archive/manifest.json` Order
The site-nav JavaScript fetches this file at page load and uses array index to build ← Prev / Next → links. Entries must remain chronologically ordered. Do not sort alphabetically or by filename.

### Status Classification Logic
`derive_status()` has specific override precedence: `COMPLETED_MILESTONES` → `FRONT_LOADED_MILESTONES` → negative balance (At Risk) → tight balance with Ahead burn (Watch) → On Track. Changing this order changes which milestones appear in the watchlist and KPI counts.

### CDN Fetches at Build Time
Chart.js and html2pdf.js are fetched from cdnjs and inlined. If you add another CDN dependency, add a `fetch_js()` call and handle the empty-string failure case explicitly — do not assume the fetch succeeds.

---

## Rules

- **Do not rename `process()`, `generate_html()`, `parse_milestones()`, or any other top-level function** — they are referenced by name in `__main__` and within each other
- **Do not add duplicate `id=` attributes** to the template — the milestone table, filter dropdown, and chart all rely on unique IDs (`milestoneRows`, `statusFilter`, `spendChart`, `bars`, etc.)
- **Do not bypass `build_chart_section()`** for chart rendering — it assembles both the Chart.js block and the fallback div-based bars
- **Reuse `fmt_usd()`, `fmt_usd_short()`, `bullet_html()`, `short_text()`** — do not inline equivalent logic elsewhere
- **Do not add runtime Python dependencies** without noting them — the only optional dependency is `python-docx` (gracefully skipped if absent)

---

## Safe Change Patterns

| Task | Safe approach |
|---|---|
| Add a new KPI card | Add token pair to template + `replacements` dict in `generate_html()` |
| Add a new aux file type | Add a new `parse_*` function + call it in `process()` after existing aux parsing |
| Change status thresholds | Modify only `derive_status()` — do not touch `compute_burn()` |
| Change forecast source | Modify only the `prior_forecast` dict-building block in `process()` |
| Add a new milestone override | Add to `FRONT_LOADED_MILESTONES` or `COMPLETED_MILESTONES` at the top of the file |
| Style changes | CSS custom properties are in `:root` in the template — prefer changing variables over adding new rules |