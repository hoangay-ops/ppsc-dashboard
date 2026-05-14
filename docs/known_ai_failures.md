# Known AI Failure Patterns

## 1. Wrong Framework Assumptions

**Symptom:** AI suggests `app.py`, Flask routes, Jinja2 `{{ }}` syntax, Bootstrap classes, or `session[]` variables.

**Cause:** The project description superficially resembles a Flask app. The actual stack is a Python script that writes static HTML — no web server, no templates directory, no Bootstrap.

**Fix:** Remind the AI: *"This is a static site generator. No Flask. No Bootstrap. Output is plain HTML files."*

---

## 2. Silent Column Drift

**Symptom:** Milestone budgets, actuals, or balances are all wrong by a consistent offset, or all zero.

**Cause:** AI modified or "cleaned up" the column index constants (`COL_TEXT`, `COL_BUDGET`, `COL_TOTAL_ACT`, etc.) without knowing the actual workbook layout. These are 0-based hardcoded positions — if the client's Excel file adds or removes a column, every value silently shifts.

**Fix:** Never let AI change `COL_*` or `MONTH_TO_COL` constants without the user explicitly confirming the new column positions from the actual workbook. Always print a sample row during debugging to verify alignment.

---

## 3. Unreplaced Placeholder Tokens

**Symptom:** Published dashboard shows literal text like `__MILESTONES_JSON__` or `__WATCHLIST_HTML__` on the page.

**Cause:** AI added a new `__TOKEN__` to `dashboard_template.html` but forgot to add the matching entry in the `replacements` dict inside `generate_html()`, or vice versa.

**Fix:** Any `__TOKEN__` added to the template must have a corresponding key in `replacements`. Any key added to `replacements` that doesn't match a template token is silently ignored — check both sides.

---

## 4. Broken Two-Pass Generation

**Symptom:** Archive snapshots have correct nav links but `index.html` doesn't, or vice versa. Or nav links always show the wrong month.

**Cause:** AI modified `generate_html()` without accounting for the fact that it's called twice — once before the archive is written (no nav links) and once after (nav links resolved). Changes that depend on archive state must go in or after the second call.

**Fix:** Do not merge the two calls into one. The two-pass pattern is intentional — the archive manifest must exist before nav links can be resolved.

---

## 5. Duplicate Element IDs

**Symptom:** Filter dropdown stops working, chart doesn't render, or milestone table is blank after a change.

**Cause:** AI added a new section to the template that reused an existing `id=` attribute (`milestoneRows`, `statusFilter`, `spendChart`, `bars`, `healthCards`, etc.). JavaScript selects these by ID — duplicates cause the first match to win and the second to be ignored.

**Fix:** Before adding any `id=` to the template, search for that string in the existing template and `automate.py`. All IDs in the template must be unique.

---

## 6. CDN Fetch Failures Treated as Errors

**Symptom:** Build crashes or produces broken HTML when internet is unavailable or cdnjs is slow.

**Cause:** AI replaced `fetch_js()` calls with `urllib.request` calls that raise on failure, or removed the empty-string fallback handling.

**Fix:** `fetch_js()` is intentionally fault-tolerant — it returns `""` on any exception. Chart.js and html2pdf.js degrade gracefully when empty. Any new CDN dependency must follow the same pattern.

---

## 7. `derive_status()` Override Order Broken

**Symptom:** Front-loaded or completed milestones incorrectly appear as "At Risk." Or milestones that should be "Watch" appear "On Track."

**Cause:** AI reordered the conditional logic inside `derive_status()` or added an early-return that skips the override dict checks.

**Fix:** The check order is load-bearing: `COMPLETED_MILESTONES` → `FRONT_LOADED_MILESTONES` → negative balance → tight+Ahead → On Track. Do not reorder or merge these conditions.

---

## 8. `manifest.json` Sort Order Corrupted

**Symptom:** Site nav ← Prev / Next → links jump to the wrong month, or the archive index shows months out of order.

**Cause:** AI sorted `archive_manifest()` output alphabetically by filename or label instead of chronologically by parsed date.

**Fix:** `archive_manifest()` sorts by `datetime` object (`sort_key`), not by string. Filenames use `%B-%Y` format (e.g. `April-2026.html`) — alphabetical sort would place April before February.