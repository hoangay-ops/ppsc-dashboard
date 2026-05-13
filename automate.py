#!/usr/bin/env python3
"""
PPSC Dashboard generator — multi-award, drill-down, prior-month forecast.

Run:
    python automate.py
"""
from __future__ import annotations

import json
import re
import shutil
import subprocess
import urllib.request
from collections import defaultdict
from dataclasses import dataclass, asdict, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

try:
    from docx import Document as DocxDocument
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# -------- CONFIG --------
ROOT     = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUT_DIR  = ROOT / "output"

SITE_BASE_URL    = "https://hoangay-ops.github.io/ppsc-dashboard/"
ARCHIVE_BASE_URL = f"{SITE_BASE_URL}archive/"
AUTO_PUSH_GITHUB = True
# Change this to the actual path of your Git repository
GIT_REPO_DIR = Path("/Users/hoangay/Projects/ppsc-dashboard")
# Front-loaded / completed milestone overrides  ── edit as needed
FRONT_LOADED_MILESTONES: Dict[str, str] = {
    "53": "Annual license fees paid upfront in February — no further spend expected.",
}
COMPLETED_MILESTONES: Dict[str, str] = {}

# -------- COLUMN MAP --------
COL_TEXT      = 1
COL_PCT       = 2
COL_BUDGET    = 3
COL_JAN, COL_FEB, COL_MAR = 4, 5, 6
COL_APR, COL_MAY, COL_JUN = 7, 8, 9
COL_JUL, COL_AUG, COL_SEP = 10, 11, 12
COL_OCT, COL_NOV, COL_DEC = 13, 14, 15
COL_TOTAL_ACT  = 16
COL_TOTAL_FCST = 17
COL_EST_SPEND  = 18
COL_EST_BAL    = 19

MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun",
               "Jul","Aug","Sep","Oct","Nov","Dec"]
MONTH_TO_COL = {m: 4+i for i, m in enumerate(MONTH_ORDER)}
MONTH_ALIASES = {
    "jan":"Jan","january":"Jan","feb":"Feb","february":"Feb",
    "mar":"Mar","march":"Mar","apr":"Apr","april":"Apr","may":"May",
    "jun":"Jun","june":"Jun","jul":"Jul","july":"Jul",
    "aug":"Aug","august":"Aug","sep":"Sep","sept":"Sep","september":"Sep",
    "oct":"Oct","october":"Oct","nov":"Nov","november":"Nov",
    "dec":"Dec","december":"Dec",
}
LINE_ITEMS = ["Labor","Travel","ODC"]
AUX_RE = re.compile(
    r"hours.*lcat|lcat.*hours|hours|lcat|story|stories|points.*mile|mile.*points|points\s+by",
    re.IGNORECASE,
)


# -------- DATA MODEL --------
@dataclass
class MilestoneRow:
    milestone_id:      str
    title:             str          # "M59 — Abbreviated Name"
    raw_title:         str          # just the name part
    award_number:      str
    status:            str
    budget:            float
    monthly_spend:     float
    percent_spent:     float        # ytd / budget × 100
    labor_actual:      float
    travel_actual:     float
    odc_actual:        float
    ytd_actual:        float
    total_forecast:    float
    est_total_spend:   float
    est_balance:       float
    expected_spend:    float
    burn_rate_status:  str
    pace_variance:     float
    projected_overrun: float
    exception_note:    str = ""
    people_count:      int   = 0
    total_hours:       float = 0.0
    story_points:      float = 0.0
    workstreams:       list  = field(default_factory=list)


# -------- HELPERS --------
def norm(v: Any) -> str:
    return str(v).replace("\xa0"," ").replace("\u2013","-").replace("\u2014","-").strip()

def sf(v: Any, d: float = 0.0) -> float:
    try:
        return float(v) if pd.notna(v) else d
    except (TypeError, ValueError):
        return d

def fmt_usd(v: float) -> str:
    try:    return f"${v:,.0f}"
    except: return "—"

def fmt_usd_short(v: float) -> str:
    try:
        a, s = abs(v), "-" if v < 0 else ""
        if a >= 1_000_000: return f"{s}${a/1_000_000:.1f}M"
        if a >= 1_000:     return f"{s}${a/1_000:.0f}k"
        return f"{s}${a:.0f}"
    except: return "—"

def bullet_html(dot_cls: str, text: str) -> str:
    return (f'<div class="bullet"><div class="dot {dot_cls or "good"}"></div>'
            f'<div>{text}</div></div>')

def fetch_js(url: str) -> str:
    try:
        with urllib.request.urlopen(url, timeout=10) as r:
            return r.read().decode("utf-8")
    except Exception as e:
        print(f"WARNING fetch {url}: {e}")
        return ""

def month_from_text(text: str) -> Optional[str]:
    s = str(text).lower().replace("_"," ")
    for key, label in MONTH_ALIASES.items():
        if re.search(rf"\b{re.escape(key)}\b", s):
            return label
    return None

def detect_month(path: Path, df: pd.DataFrame) -> str:
    m = month_from_text(path.stem)
    if m: return m
    flat = " ".join(str(x) for x in df.head(10).astype(str).values.flatten())
    return month_from_text(flat) or "Dec"

def is_aux(path: Path) -> bool:
    return bool(AUX_RE.search(path.name))

def choose_sheet(xls: pd.ExcelFile) -> str:
    return xls.sheet_names[0] if len(xls.sheet_names) == 1 else xls.sheet_names[1]

def split_title(text: str) -> Tuple[Optional[str], str]:
    """Handle formats: 'Milestone 44- Directorate', 'Milestone 44 — Name', 'Milestone 44'."""
    m = re.search(r"Milestone\s*(\d+)", text, re.IGNORECASE)
    if not m:
        return None, text
    mid = f"M{m.group(1)}"
    # Split on any dash/em-dash variant with optional surrounding spaces
    remainder = re.split(r"\s*[-—–]+\s*", text, maxsplit=1)
    raw = remainder[1].strip() if len(remainder) > 1 and remainder[1].strip() else ""
    # If the split ate into the milestone number, raw will be wrong — guard it
    if re.match(r"^\d", raw):
        raw = ""
    return mid, raw

def compute_burn(budget: float, ytd: float, est_spend: float,
                 rep_idx: int) -> Tuple[float, str, float, float]:
    elapsed  = rep_idx + 1
    expected = round(budget * elapsed / 12, 2)
    variance = round(ytd - expected, 2)
    allow    = budget / 12 if budget > 0 else 0
    ratio    = (ytd / elapsed / allow) if allow and elapsed else 0
    bst = "Ahead" if ratio > 1.15 else "Behind" if ratio < 0.85 else "On Pace"
    overrun = round(est_spend - budget, 2)
    return expected, bst, variance, overrun

def month_slug(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "-", s).strip("-")

def short_text(text: str, limit: int = 70) -> str:
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if len(text) <= limit: return text
    cut = text[:limit].rsplit(" ", 1)[0]
    return (cut or text[:limit]).rstrip(" .") + "…"

def format_month_year(abbr: str, file_path: Path) -> str:
    full = datetime.strptime(abbr, "%b").strftime("%B")
    m = re.search(r"(20\d{2})", file_path.stem)
    yr = m.group(1) if m else str(datetime.now().year)
    return f"{full} {yr}"


# -------- WORKBOOK PARSING --------
def read_main_workbook(path: Path) -> Tuple[pd.DataFrame, str]:
    """Return (dataframe, reporting_month)."""
    xls    = pd.ExcelFile(path)
    sheet  = choose_sheet(xls)
    df     = pd.read_excel(path, sheet_name=sheet, header=None)
    return df, detect_month(path, df)

def extract_award_number(df: pd.DataFrame, fallback: str) -> str:
    # Scan every cell in the first 20 rows for "award number" nearby
    for _, r in df.iterrows():
        for col_idx in range(len(r)):
            cell = norm(str(r.iloc[col_idx]))
            if "award number" in cell.lower():
                # Check the next several columns on the same row for the value
                for look in range(col_idx + 1, min(col_idx + 8, len(r))):
                    val = norm(str(r.iloc[look]))
                    if val and val.lower() not in ("nan", "", "none"):
                        return val
    # Fallback: extract known award number patterns from filename
    patterns = [
        r"OT[\w]+",          # OT2OD036464
        r"\d{4,5}\.\d+",     # 9877.26
        r"[A-Z]{2}-[\w-]+",  # GS-10F-0033N
    ]
    for pat in patterns:
        m = re.search(pat, fallback)
        if m:
            return m.group(0)
    return fallback

def extract_monthly_row_totals(df: pd.DataFrame) -> Dict[str, float]:
    totals: Dict[str, float] = {m: 0.0 for m in MONTH_ORDER}
    for _, r in df.iterrows():
        text = norm(r[COL_TEXT])
        # Look for the 'Total Milestone' row to capture planned spend per month
        if text.startswith("Total Milestone"):
            for mo in MONTH_ORDER:
                col_idx = MONTH_TO_COL[mo] 
                # This pulls the monthly plan from the prior workbook
                totals[mo] += sf(r[col_idx])
    return totals

def parse_milestones(df: pd.DataFrame, rep_month: str,
                     award_number: str) -> List[MilestoneRow]:
    rep_col = MONTH_TO_COL[rep_month]
    rep_idx = MONTH_ORDER.index(rep_month)
    rows: List[MilestoneRow] = []
    cur_id = cur_title = None
    accum: Dict[str, Dict[str, float]] = {}

    for _, r in df.iterrows():
        text = norm(r[COL_TEXT])

        if re.match(r"Milestone\s+\d+", text, re.IGNORECASE):
            cur_id, cur_title = split_title(text)
            accum = {c: {"budget":0.0,"month":0.0,"ytd":0.0} for c in LINE_ITEMS}

        elif text in LINE_ITEMS and cur_id:
            accum[text] = {
                "budget": sf(r[COL_BUDGET]),
                "month":  sf(r[rep_col]),
                "ytd":    sf(r[COL_TOTAL_ACT]),
            }

        elif "hours" in text.lower() and cur_id:
        # Captures total hours for the milestone from the Total Actual column
            cur_hours = sf(r[COL_TOTAL_ACT]) 
        # Update the latest milestone in our list
            if rows: rows[-1].total_hours = cur_hours

        elif ("staff" in text.lower() or "personnel" in text.lower()) and cur_id:
        # Captures staff count for the milestone
            cur_staff = sf(r[COL_TOTAL_ACT])
            if rows: rows[-1].people_count = int(cur_staff)

        elif text.startswith("Total Milestone") and cur_id:
            budget     = sf(r[COL_BUDGET])
            ytd        = sf(r[COL_TOTAL_ACT])
            est_spend  = sf(r[COL_EST_SPEND])
            est_bal    = sf(r[COL_EST_BAL])
            month_act  = sf(r[rep_col])
            total_fcst = sf(r[COL_TOTAL_FCST])

            # Force consistent percentage calculation
            if budget > 0:
                pct_spent = round((ytd / budget) * 100, 1) #
            else:
                pct_spent = 0.0 #

            overrun = round(est_spend - budget, 2) #

            # Read % spent directly from Excel (col 2) — already calculated

            raw_pct = sf(r[COL_PCT])
            # Excel stores it as a decimal (0.1442) or whole number (14.42)
            if 0 < raw_pct <= 1:
                pct_spent = round(raw_pct * 100, 1)
            elif raw_pct > 1:
                pct_spent = round(raw_pct, 1)
            else:
                pct_spent = round(ytd / budget * 100, 1) if budget else 0.0

            # Projected overrun = est_spend - budget (positive = overrun, negative = under)
            overrun = round(est_spend - budget, 2)

            exp_sp, bst, pace_v, _ = compute_burn(budget, ytd, est_spend, rep_idx)
            mid = cur_id.lstrip("M")

            exception = FRONT_LOADED_MILESTONES.get(mid,"") or COMPLETED_MILESTONES.get(mid,"")
            is_done   = (month_act == 0 and ytd > 0 and est_bal >= 0 and rep_idx >= 3)

            if mid in COMPLETED_MILESTONES or is_done:
                status = "Complete"
            elif mid in FRONT_LOADED_MILESTONES:
                status = "On Track"
            elif bst == "Ahead" and est_bal < 0:
                status = "At Risk"
            elif est_bal < 0:
                status = "Watch"
            elif bst == "Ahead":
                status = "Watch"
            else:
                status = "On Track"

            combined = f"{cur_id} — {cur_title}" if cur_title else cur_id
            rows.append(MilestoneRow(
                milestone_id    = cur_id,
                title           = combined,
                raw_title       = cur_title or cur_id,
                award_number    = award_number,
                status          = status,
                budget          = budget,
                monthly_spend   = month_act,
                percent_spent   = pct_spent,
                labor_actual    = accum.get("Labor",{}).get("month",0.0),
                travel_actual   = accum.get("Travel",{}).get("month",0.0),
                odc_actual      = accum.get("ODC",{}).get("month",0.0),
                ytd_actual      = ytd,
                total_forecast  = total_fcst,
                est_total_spend = est_spend,
                est_balance     = est_bal,
                expected_spend  = exp_sp,
                burn_rate_status= bst,
                pace_variance   = pace_v,
                projected_overrun=overrun,
                exception_note  = exception,
                workstreams = [{"name": c,
                                "val": accum.get(c, {}).get("month", 0.0), # Use 'val' to match HTML
                                "ytd": accum.get(c, {}).get("ytd", 0.0),
                                "budget": accum.get(c, {}).get("budget", 0.0)}
                                for c in LINE_ITEMS]
            ))
    return rows


# -------- PRIOR-MONTH FORECAST --------
def build_forecast_from_prior(all_main_files: List[Path], current_file: Path, current_month: str) -> Dict[str, float]:
    dated = sorted(all_main_files, key=lambda p: p.stat().st_mtime)
    try:
        cur_pos = next(i for i, p in enumerate(dated) if p == current_file)
        # If there's a file before this one, use it. 
        # If not, use the current_file itself to show the current forecast vs actuals.
        prior_path = dated[cur_pos - 1] if cur_pos > 0 else current_file
    except StopIteration:
        return {m: 0.0 for m in MONTH_ORDER}

    print(f"  Forecast source: {prior_path.name}")
    try:
        prior_df, _ = read_main_workbook(prior_path)
        return extract_monthly_row_totals(prior_df)
    except Exception as e:
        return {m: 0.0 for m in MONTH_ORDER}


# -------- AUX (hours / stories) --------
def parse_narrative_doc(path: Path) -> Dict[str, str]:
    if not DOCX_OK: return {}
    try: doc = DocxDocument(path)
    except Exception as e:
        print(f"WARNING docx {path.name}: {e}"); return {}
    narratives: Dict[str, str] = {}
    cur_sec, cur_lines = None, []

    def flush():
        if cur_sec and cur_lines:
            narratives[cur_sec] = " ".join(cur_lines).strip()

    for para in doc.paragraphs:
        t = para.text.strip()
        if not t: continue
        is_mil = bool(re.search(r"^\b(Milestone|M)\s*\d+", t, re.IGNORECASE))
        is_port = (len(t)<100 and any(k in t.lower() for k in
                   ["portfolio summary","costs incurred","financial report","executive summary"]))
        if is_mil or is_port:
            flush(); cur_sec, cur_lines = t, []
        elif cur_sec is not None:
            cur_lines.append(t)
    flush()
    return narratives

def narrative_for(narratives: Dict[str, str], title: str) -> str:
    m = re.search(r"(\d+)", title)
    if not m: return ""
    mid = m.group(1)
    for sec, txt in narratives.items():
        if re.search(rf"\b{mid}\b", sec, re.IGNORECASE):
            return txt
    return ""

def parse_stories_file(path: Path) -> Dict:
    """
    Parse 'Points by Milestones.xlsx' or similar.
    Expects columns: Milestone (or similar), Points (or Story Points).
    Returns totals and per-milestone breakdown.
    """
    try:
        xls = pd.ExcelFile(path)
    except Exception as e:
        print(f"WARNING stories {path.name}: {e}")
        return {}

    def find_col(cols: List[str], candidates: List[str]) -> Optional[str]:
        for c in cols:
            if any(k.lower() in c.lower() for k in candidates):
                return c
        return None

    by_milestone: Dict[str, float] = {}
    total_points = 0.0

    for sheet in xls.sheet_names:
        try:
            # Try to find header row in first 15 rows
            raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=15)
            header_row = 0
            for i in range(len(raw)):
                row_text = " ".join(str(v).lower() for v in raw.iloc[i])
                if any(k in row_text for k in ("milestone","point","story")):
                    header_row = i
                    break
            df = pd.read_excel(path, sheet_name=sheet, header=header_row)
            df.columns = [str(c).strip() for c in df.columns]
        except Exception:
            continue

        mil_col = find_col(df.columns, ["milestone","workstream","task","title","name","id"])
        pts_col = find_col(df.columns, ["point","points","story point","story points","sp"])
        if not mil_col or not pts_col:
            continue

        for _, row in df.iterrows():
            mil = str(row.get(mil_col,"")).strip()
            pts = sf(row.get(pts_col, 0))
            if not mil or mil.lower() in ("nan","none","") or pts == 0:
                continue
            # Normalize milestone key — extract M-number if present
            m = re.search(r"M?(\d+)", mil, re.IGNORECASE)
            key = f"M{m.group(1)}" if m else mil
            by_milestone[key] = by_milestone.get(key, 0.0) + pts
            total_points += pts
        break  # use first valid sheet

    if not by_milestone:
        return {}

    by_mil_list = sorted(
        [{"milestone_id": k, "points": round(v, 1)} for k, v in by_milestone.items()],
        key=lambda x: x["points"], reverse=True
    )
    return {
        "total_points":  round(total_points, 1),
        "total_mils":    len(by_milestone),
        "by_milestone":  by_mil_list,
    }


def build_stories_panel(stories: Dict, all_milestones: List[MilestoneRow]) -> str:
    return """
    <div class="panel" style="margin-top:18px;">
        <h2>Story Points & Velocity</h2>
        <div style="padding:40px; text-align:center; background:rgba(15,23,42,0.02); border:2px dashed rgba(15,23,42,0.1); border-radius:14px;">
            <p style="color:#51627d; font-style:italic; font-size:12px">Story point tracking integration in progress. <br><br> 
            Please ensure 'Points by Milestones.xlsx' is present in the data folder.</p>
        </div>
    </div>
    """

    total   = stories.get("total_points", 0)
    n_mils  = stories.get("total_mils", 0)
    by_mil  = stories.get("by_milestone", [])
    max_pts = by_mil[0]["points"] if by_mil else 1

    # Build a title lookup from parsed milestones
    title_map = {m.milestone_id: m.raw_title for m in all_milestones}

    rows_html = ""
    for item in by_mil[:12]:
        mid   = item["milestone_id"]
        pts   = item["points"]
        label = title_map.get(mid, "")
        bar_w = round(pts / max_pts * 100)
        rows_html += (
            f'<div style="margin-bottom:10px;">'
            f'<div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:3px;">'
            f'<span><strong>{mid}</strong>'
            f'{(" — "+label) if label else ""}</span>'
            f'<span style="font-weight:700;">{pts:,.0f} pts</span></div>'
            f'<div style="height:6px;background:rgba(15,23,42,0.08);border-radius:3px;overflow:hidden;">'
            f'<div style="width:{bar_w}%;height:100%;background:#7c3aed;border-radius:3px;"></div></div>'
            f'</div>'
        )

    return (
        f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;">'
        f'<div style="background:rgba(124,58,237,0.06);border:1px solid rgba(124,58,237,0.18);'
        f'border-radius:14px;padding:12px 16px;">'
        f'<div style="font-size:12px;color:#51627d;text-transform:uppercase;letter-spacing:.06em;">Total Points</div>'
        f'<div style="font-size:22px;font-weight:800;margin-top:4px;">{total:,.0f}</div>'
        f'</div>'
        f'<div style="background:rgba(124,58,237,0.06);border:1px solid rgba(124,58,237,0.18);'
        f'border-radius:14px;padding:12px 16px;">'
        f'<div style="font-size:12px;color:#51627d;text-transform:uppercase;letter-spacing:.06em;">Milestones</div>'
        f'<div style="font-size:22px;font-weight:800;margin-top:4px;">{n_mils}</div>'
        f'</div></div>'
        f'<div style="font-size:13px;font-weight:700;margin-bottom:8px;">Points by milestone</div>'
        f'{rows_html}'
    )
    if not DOCX_OK: return {}
    try: doc = DocxDocument(path)
    except Exception as e:
        print(f"WARNING docx {path.name}: {e}"); return {}
    narratives: Dict[str, str] = {}
    cur_sec, cur_lines = None, []

    def flush():
        if cur_sec and cur_lines:
            narratives[cur_sec] = " ".join(cur_lines).strip()

    for para in doc.paragraphs:
        t = para.text.strip()
        if not t: continue
        is_mil = bool(re.search(r"^\b(Milestone|M)\s*\d+", t, re.IGNORECASE))
        is_port = (len(t)<100 and any(k in t.lower() for k in
                   ["portfolio summary","costs incurred","financial report","executive summary"]))
        if is_mil or is_port:
            flush(); cur_sec, cur_lines = t, []
        elif cur_sec is not None:
            cur_lines.append(t)
    flush()
    return narratives

def narrative_for(narratives: Dict[str, str], title: str) -> str:
    m = re.search(r"(\d+)", title)
    if not m: return ""
    mid = m.group(1)
    for sec, txt in narratives.items():
        if re.search(rf"\b{mid}\b", sec, re.IGNORECASE):
            return txt
    return ""


# -------- ARCHIVE HELPERS --------
def archive_manifest(archive_dir: Path) -> List[Dict]:
    items = []
    for p in archive_dir.glob("*.html"):
        if p.name == "index.html": continue
        try:   dt = datetime.strptime(p.stem, "%B-%Y")
        except ValueError: continue
        items.append({"label": dt.strftime("%B %Y"), "file": p.name, "sort_key": dt})
    items.sort(key=lambda x: x["sort_key"])
    return items

def build_archive_index(archive_dir: Path, current_month_full: str) -> None:
    items = archive_manifest(archive_dir)
    cards = "".join(f"""
      <a class="month-card" href="{it['file']}">
        <div class="month-title">{it['label']}{"  (current)" if it['label']==current_month_full else ""}</div>
        <div class="month-copy">Open archived snapshot</div>
      </a>""" for it in items) or '<div style="color:#51627d;">No archives yet.</div>'

    html = f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>PPSC Archive</title>
<style>
  :root{{--shadow:0 18px 40px rgba(15,23,42,0.10);--radius:22px;}}
  *{{box-sizing:border-box;}}
  body{{margin:0;font-family:Inter,ui-sans-serif,sans-serif;color:#10203a;
    background:linear-gradient(180deg,#fff 0%,#f6f8fc 52%,#eef3fb 100%);min-height:100vh;}}
  .wrap{{max-width:1300px;margin:0 auto;padding:14px;}}
  .hero,.panel{{background:linear-gradient(180deg,rgba(255,255,255,0.96),rgba(248,250,252,0.98));
    border:1px solid rgba(15,23,42,0.10);box-shadow:var(--shadow);border-radius:var(--radius);}}
  .hero{{padding:26px;margin-bottom:18px;}}.panel{{padding:18px;}}
  .header-bar{{display:flex;justify-content:flex-end;margin-bottom:12px;}}
  .header-actions{{display:flex;gap:10px;}}
  .header-actions a{{font-size:12px;padding:8px 12px;border-radius:6px;text-decoration:none;
    border:1px solid #cbd5e1;background:#fff;color:#10203a;}}
  .grid{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-top:16px;}}
  .month-card{{display:block;text-decoration:none;color:#10203a;padding:18px;border-radius:var(--radius);}}
  .month-card:hover{{background:rgba(37,99,235,0.03);}}
  .month-title{{font-size:18px;font-weight:800;margin-bottom:6px;}}
  .month-copy{{color:#51627d;font-size:13px;}}
  @media(max-width:1100px){{.grid{{grid-template-columns:1fr;}}}}
</style></head><body>
<div class="wrap">
  <div class="header-bar"><div class="header-actions">
    <a href="../index.html">Latest</a><a href="./">Archive</a>
  </div></div>
  <section class="hero"><h1>PPSC Dashboard Archive</h1>
    <p style="margin-top:10px;color:#51627d;">Monthly snapshots generated by the automation.</p>
  </section>
  <section class="panel"><div class="grid">{cards}</div></section>
</div></body></html>"""
    (archive_dir / "index.html").write_text(html, encoding="utf-8")

def sync_and_push(out_dir: Path, repo_dir: Path, message: str) -> None:
    repo_arch = repo_dir / "archive"
    repo_arch.mkdir(parents=True, exist_ok=True)
    shutil.copy2(out_dir / "index.html", repo_dir / "index.html")
    shutil.copy2(out_dir / "archive" / "index.html", repo_arch / "index.html")
    mj = out_dir / "archive" / "months.json"
    if mj.exists(): shutil.copy2(mj, repo_arch / "months.json")
    for src in (out_dir / "archive").glob("*.html"):
        if src.name != "index.html":
            shutil.copy2(src, repo_arch / src.name)

    # Always fetch and reset to remote first — avoids merge conflicts.
    # The generated HTML is always the source of truth; we never need local history.
    subprocess.run(["git","-C",str(repo_dir),"fetch","origin"], check=True)
    subprocess.run(["git","-C",str(repo_dir),"reset","--hard","origin/main"], check=True)

    # Now stage and push the freshly generated files
    subprocess.run(["git","-C",str(repo_dir),"add","index.html","archive"], check=True)
    result = subprocess.run(["git","-C",str(repo_dir),"commit","-m",message], check=False)
    if result.returncode not in (0, 1):  # 1 = nothing to commit, that's fine
        raise RuntimeError(f"git commit failed with code {result.returncode}")
    subprocess.run(["git","-C",str(repo_dir),"push"], check=True)


# -------- HTML GENERATION --------
def generate_html(
    all_milestones:       List[MilestoneRow],
    monthly_chart:        List[Dict],
    reporting_month:      str,
    reporting_month_full: str,
    awards_str:           str,
    narratives:           Dict[str, str],
    stories_data:         Dict,
) -> None:

    template_path = next(
        (p for p in [ROOT/"dashboard_template_fixed.html", ROOT/"dashboard_template.html"]
         if p.exists()), None
    )
    if not template_path:
        print("ERROR: dashboard_template.html not found"); return

    html = template_path.read_text(encoding="utf-8")

    OUT_DIR.mkdir(exist_ok=True)
    archive_dir = OUT_DIR / "archive"
    archive_dir.mkdir(exist_ok=True)

    print("Fetching Chart.js…")
    chartjs     = fetch_js("https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js")
    print("Fetching html2pdf…")
    html2pdf_js = fetch_js("https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js")

    at_risk  = [m for m in all_milestones if m.status == "At Risk"]
    on_watch = [m for m in all_milestones if m.status == "Watch"]
    on_track = [m for m in all_milestones if m.status in ("On Track","Complete")]

    budget_tot = sum(m.budget        for m in all_milestones)
    ytd_tot    = sum(m.ytd_actual    for m in all_milestones)
    month_tot  = sum(m.monthly_spend for m in all_milestones)
    bal_tot    = sum(m.est_balance   for m in all_milestones)

    rep_idx          = MONTH_ORDER.index(reporting_month)
    budget_used_pct  = round(ytd_tot / budget_tot * 100, 1) if budget_tot else 0.0
    year_elapsed_pct = round((rep_idx + 1) / 12 * 100, 1)
    budget_pace_gap  = round(budget_used_pct - year_elapsed_pct, 1)
    balance_color    = "#dc2626" if bal_tot < 0 else "#059669"

    prev_month_val = monthly_chart[rep_idx-1]["value"] if rep_idx > 0 else month_tot
    trend_pct   = round((month_tot - prev_month_val) / prev_month_val * 100, 1) if prev_month_val else 0.0
    trend_arrow = "▲" if trend_pct > 0 else "▼" if trend_pct < 0 else "→"
    trend_color = "#dc2626" if trend_pct > 0 else "#059669" if trend_pct < 0 else "#51627d"

    flagged  = at_risk + on_watch
    n_flags  = len(flagged)
    n_crit   = len(at_risk)
    flag_summary = (
        "No milestone flags need escalation." if n_flags == 0 else
        f"{n_flags} flagged, all minor — monitor closely." if n_crit == 0 else
        f"{n_flags} flagged — {n_crit} require{'s' if n_crit==1 else ''} attention."
    )

    tension_msg = (
        f"Budget used is {budget_used_pct:.1f}% versus {year_elapsed_pct:.1f}% "
        f"of the year elapsed. {len(at_risk)} milestone(s) At Risk and {len(on_watch)} on Watch."
    )

    # Watchlist bullets
    watchlist = sorted(
        [m for m in all_milestones if m.status in ("At Risk","Watch") and not m.exception_note],
        key=lambda m: (m.status == "At Risk", abs(m.projected_overrun)),
        reverse=True
    )[:5]

    watchlist_html = "\n".join(
        bullet_html(
            "risk" if m.status=="At Risk" else "warn",
            f"<strong>{m.title}</strong> — {fmt_usd_short(abs(m.projected_overrun))} overrun"
            + (f" — {short_text(narrative_for(narratives, m.title))}" if narrative_for(narratives, m.title) else "")
            + "."
        )
        for m in watchlist
    ) or bullet_html("good", "No milestones currently flagged for attention.")

    # Callouts
    callouts = []
    if bal_tot < 0:
        callouts.append(bullet_html("risk",
            f"<strong>Portfolio balance is negative ({fmt_usd(bal_tot)}).</strong> "
            f"Projected spend is above total budget."))
    else:
        callouts.append(bullet_html("good",
            f"<strong>Portfolio balance is positive at {fmt_usd(bal_tot)}.</strong> "
            f"Budget remains on track."))
    callouts.append(bullet_html("warn" if n_flags else "good",
                                f"<strong>{flag_summary}</strong>"))
    for m in at_risk[:3]:
        callouts.append(bullet_html("risk",
            f"<strong>{m.title} — {fmt_usd_short(abs(m.projected_overrun))} overrun</strong>"
            + (f" — {short_text(narrative_for(narratives, m.title))}" if narrative_for(narratives, m.title) else "")
            + "."))
    callout_html = "\n".join(callouts)

    # Operational highlights
    op_items = []
    credits = sorted([m for m in all_milestones if m.monthly_spend < 0],
                     key=lambda m: m.monthly_spend)
    for m in credits:
        op_items.append(bullet_html("good",
            f"<strong>Credit received — {m.title}:</strong> "
            f"{fmt_usd(abs(m.monthly_spend))} credit posted this month."))
    spikes = [m for m in all_milestones
              if m.monthly_spend > 0 and m.budget > 0
              and m.monthly_spend / m.budget > 0.25]
    for m in sorted(spikes, key=lambda m: m.monthly_spend/m.budget, reverse=True):
        p = round(m.monthly_spend / m.budget * 100, 1)
        op_items.append(bullet_html("warn",
            f"<strong>Spend spike — {m.title}:</strong> "
            f"{fmt_usd(m.monthly_spend)} this month ({p}% of annual budget). "
            f"Verify invoices are correctly period-coded."))
    if not op_items:
        for m in sorted(all_milestones, key=lambda m: m.monthly_spend, reverse=True)[:5]:
            op_items.append(bullet_html("good",
                f"<strong>{m.title}</strong> — {fmt_usd(m.monthly_spend)} this month "
                f"({m.percent_spent:.1f}% of budget spent to date)."))
    operational_html = "\n".join(op_items) or bullet_html("good","No spend data.")

    # ── Chart section (Chart.js) ──
    monthly_json = json.dumps(monthly_chart)
    chart_section = f"""
  <section class="grid" style="margin-bottom:18px;">
    <div class="panel">
      <h2>Monthly Spend Trend — Actuals vs. Prior-Month Forecast</h2>
      <p class="note">Blue bars = actuals. Orange bars = what the previous month's workbook forecast for the same period.</p>
      <div style="display:flex;gap:16px;margin-bottom:8px;font-size:12px;color:#51627d;">
        <span style="display:flex;align-items:center;gap:4px;"><span style="width:10px;height:10px;border-radius:2px;background:#2563eb;display:inline-block;"></span>Actual</span>
        <span style="display:flex;align-items:center;gap:4px;"><span style="width:10px;height:10px;border-radius:2px;background:#d97706;display:inline-block;"></span>Prior forecast</span>
      </div>
      <div style="position:relative;width:100%;height:280px;">
        <canvas id="spendChart" role="img" aria-label="Monthly spend trend"></canvas>
      </div>
      <div class="footer" style="margin-top:8px;">Forecast sourced from the previous month's workbook column totals.</div>
    </div>
    <div class="panel" style="display:flex;flex-direction:column;gap:14px;">
      <div>
        <h2 style="margin-bottom:6px;">Status Mix</h2>
        <p class="note" style="margin-bottom:10px;">On Track = within budget · Watch = balance tightening · At Risk = spending ahead of plan.</p>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;">
          <div class="status-chip status-good"><div class="txt">On Track</div><div class="num">{len(on_track)}</div></div>
          <div class="status-chip status-watch"><div class="txt">Watch</div><div class="num">{len(on_watch)}</div></div>
          <div class="status-chip status-risk"><div class="txt">At Risk</div><div class="num">{len(at_risk)}</div></div>
        </div>
      </div>
      <div style="flex:1;">
        <h2 style="margin-bottom:4px;">Priority Watchlist</h2>
        <p class="note" style="margin-bottom:8px;">Sorted by budget pressure and projected overrun. Use this list to focus the next leadership discussion.</p>
        <div class="bullets">{watchlist_html}</div>
      </div>
    </div>
  </section>
  <script>{chartjs}</script>
  <script>
    (function() {{
      var md = {monthly_json};
      var labels   = md.map(function(d){{ return d.month; }});
      var actuals  = md.map(function(d){{ return d.value > 0 ? d.value : null; }});
      var forecast = md.map(function(d){{ return (d.forecast_val && d.forecast_val > 0) ? d.forecast_val : null; }});
      var canvas = document.getElementById('spendChart');
      if (!canvas) return;
      new Chart(canvas, {{
        type: 'bar',
        data: {{
          labels: labels,
          datasets: [
            {{ label:'Actual Spend', data: actuals,
               backgroundColor:'rgba(37,99,235,0.85)', borderRadius:4, borderSkipped:false, order:2 }},
            {{ label:'Prior Forecast', data: forecast,
               backgroundColor:'rgba(217,119,6,0.75)', borderRadius:4, borderSkipped:false, order:1 }}
          ]
        }},
        options: {{
          responsive:true, maintainAspectRatio:false,
          interaction:{{ mode:'index', intersect:false }},
          plugins:{{ legend:{{ display:false }},
            tooltip:{{ callbacks:{{ label:function(ctx){{
              if (ctx.parsed.y===null) return null;
              return ctx.dataset.label+': $'+Math.abs(ctx.parsed.y).toLocaleString(undefined,{{maximumFractionDigits:0}});
            }} }} }}
          }},
          scales:{{
            x:{{ ticks:{{ autoSkip:false, maxRotation:0, font:{{size:11}} }} }},
            y:{{ ticks:{{ callback:function(v){{
              if (v>=1000000) return '$'+(v/1000000).toFixed(1)+'M';
              if (v>=1000)    return '$'+(v/1000).toFixed(0)+'k';
              return '$'+v;
            }}, font:{{size:11}} }}, min:0 }}
          }}
        }}
      }});
    }})();
  </script>"""

    # Replace the first <section class="grid"> block in the template
    start = html.find('<section class="grid">')
    end   = html.find('</section>', start) + len('</section>') if start != -1 else -1
    if start != -1 and end != -1:
        html = html[:start] + chart_section + html[end:]

    # Milestone span label
    milestone_span = "N/A"
    if all_milestones:
        ids = [m.milestone_id for m in all_milestones]
        milestone_span = f"{ids[0]}–{ids[-1]}"

    slug         = month_slug(reporting_month_full)
    current_file = f"{slug}.html"

    replacements = {
        "__MONTH__":           reporting_month_full,
        "__BALANCE_COLOR__":   balance_color,
        "__TITLE__":           f"PPSC Financial Dashboard — {reporting_month_full}",
        "__SUBTITLE__":        (f"Award(s): {awards_str} · {len(all_milestones)} milestones tracked · "
                                f"{reporting_month_full} reporting period."),
        "__AWARD__":           awards_str,
        "__MILESTONE_SPAN__":  milestone_span,
        "__AT_A_GLANCE__":     tension_msg,
        "__YTD_SPEND_TOTAL__": fmt_usd(ytd_tot),
        "__AWARD_BUDGET__":    fmt_usd(budget_tot),
        "__MONTHLY_SPEND__":   fmt_usd(month_tot),
        "__ON_TRACK__":        str(len(on_track)),
        "__WATCH_ITEMS__":     str(len(at_risk)),
        "__TREND_NOTE__":      "Monthly spend — blue = actuals, orange = prior-month forecast.",
        "__TREND_FOOTER__":    "Forecast sourced from the previous month's workbook column totals.",
        "__STATUS_NOTE__":     "On Track = within budget · Watch = balance tightening · At Risk = spending ahead of plan.",
        "__COUNT_ON_TRACK__":  str(len(on_track)),
        "__COUNT_AT_RISK__":   str(len(at_risk)),
        "__COUNT_AWAITING__":  str(len(on_watch)),
        "__COUNT_NOT_STARTED__":"0",
        "__WATCHLIST_HTML__":  watchlist_html,
        "__TENSION_MSG__":     tension_msg,
        "__OPERATIONAL_NOTE__":"Credits, spend spikes, and notable activity this period.",
        "__OPERATIONAL_HTML__":operational_html,
        "__CALLOUT_NOTE__":    "Action items for At Risk and Watch milestones.",
        "__CALLOUT_HTML__":    callout_html,
        "__CONSOLIDATED_CALLOUTS_HTML__": callout_html,
        "__NARRATIVE_PANEL__":"",
        "__FOOTER__":          (f"Generated {pd.Timestamp.now().strftime('%B %d, %Y %H:%M')} · "
                                f"Source: {reporting_month_full} workbook · Award(s): {awards_str}"),
        "__MONTHLY_DATA_JSON__": monthly_json,
        "__MILESTONES_JSON__":   json.dumps([asdict(m) for m in all_milestones]),
        "__HTML2PDF_JS__":       f"<script>{html2pdf_js}</script>",
        "__CURRENT_FILE__":      current_file,
        "__SITE_BASE_URL__":     SITE_BASE_URL,
        "__ARCHIVE_BASE_URL__":  ARCHIVE_BASE_URL,
        "__MONTHS_JSON_PATH__":  f"{ARCHIVE_BASE_URL}months.json",
        "__BUDGET_USED_PCT__":   f"{budget_used_pct:.1f}",
        "__YEAR_ELAPSED_PCT__":  f"{year_elapsed_pct:.1f}",
        "__BUDGET_PACE_GAP__":   f"{budget_pace_gap:+.1f}",
        "__TREND_ARROW__":       trend_arrow,
        "__TREND_PCT__":         f"{abs(trend_pct):.1f}",
        "__TREND_COLOR__":       trend_color,
        "__FLAGGED_COUNT__":     str(n_flags),
        "__FLAGGED_SUMMARY__":   flag_summary,
        "__MINOR_FLAG_COUNT__":  str(max(n_flags - n_crit, 0)),
        "__PORTFOLIO_BALANCE__": fmt_usd(bal_tot),
        "__TREND_TEXT__":        f"{trend_arrow} {abs(trend_pct):.1f}% vs prior month.",
        "__DECISION_1__":        (f"{n_crit} milestone(s) need immediate review."
                                  if n_crit else "No immediate milestone decisions required."),
        "__DECISION_2__":        f"Budget used is {budget_used_pct:.1f}% versus {year_elapsed_pct:.1f}% of the year elapsed.",
        "__DECISION_3__":        "Confirm whether front-loaded milestones should remain excluded from pace alerts.",
        "__STORIES_PANEL__":     build_stories_panel(stories_data, all_milestones),
        "__HOURS_PANEL__":       "",   # populated separately if hours file present
        "__WHAT_VISUALIZED__":   "Monthly actuals vs. prior-month forecast, spend trend, risk flags.",
    }

    for k, v in replacements.items():
        html = html.replace(k, v)

    remaining = re.findall(r"__[A-Z_]+__", html)
    if remaining:
        print(f"WARNING: Unreplaced placeholders: {set(remaining)}")

    latest_path  = OUT_DIR / "index.html"
    archive_path = archive_dir / f"{slug}.html"
    latest_path.write_text(html, encoding="utf-8")
    archive_path.write_text(html, encoding="utf-8")

    months = archive_manifest(archive_dir)
    (archive_dir / "months.json").write_text(
        json.dumps([{"label": it["label"], "file": it["file"]} for it in months], indent=2),
        encoding="utf-8",
    )
    build_archive_index(archive_dir, reporting_month_full)
    print(f"LATEST:   {latest_path}")
    print(f"ARCHIVED: {archive_path}")


# -------- MAIN PROCESS --------
def process() -> None:
    print(">>> PPSC Dashboard — multi-award process starting")

    all_xlsx = sorted(p for p in DATA_DIR.glob("*.xlsx") if p.is_file())
    if not all_xlsx:
        print("ERROR: No .xlsx files in data/"); return

    main_files = [p for p in all_xlsx if not is_aux(p)]
    if not main_files:
        print("ERROR: No main workbooks found (all files matched auxiliary pattern)"); return

    # Sort by mtime so the newest is the reporting month
    main_files.sort(key=lambda p: p.stat().st_mtime)
    current_file = main_files[-1]

    all_milestones: List[MilestoneRow] = []
    awards_found:   List[str]          = []
    monthly_actuals: Dict[str, float]  = {m: 0.0 for m in MONTH_ORDER}
    reporting_month = "Jan"

    for path in main_files:
        print(f"--- Reading: {path.name}")
        try:
            df, rep_month = read_main_workbook(path)
        except Exception as e:
            print(f"  WARN: {e}"); continue

        if path == current_file:
            reporting_month = rep_month

        award = extract_award_number(df, path.stem)
        if award not in awards_found:
            awards_found.append(award)

        mils = parse_milestones(df, rep_month, award)
        print(f"  Found {len(mils)} milestones  (award: {award}, month: {rep_month})")

        # Only accumulate actuals from the current (newest) file
        if path == current_file:
            all_milestones.extend(mils)
            for mo in MONTH_ORDER:
                monthly_actuals[mo] += extract_monthly_row_totals(df).get(mo, 0.0)
        else:
            # For older award files, still include their milestones
            all_milestones.extend(mils)

    # Prior-month forecast
    prior_forecast = build_forecast_from_prior(main_files, current_file, reporting_month)

    rep_idx = MONTH_ORDER.index(reporting_month)
    budget_tot = sum(m.budget for m in all_milestones)

    monthly_chart = []
    for i, mo in enumerate(MONTH_ORDER):
        actual = monthly_actuals[mo]
        forecast = prior_forecast.get(mo, 0.0)
        monthly_chart.append({
            "month":        mo,
            "value":        round(actual, 2),
            "forecast_val": round(forecast, 2) if forecast > 0 else None, # Ensure this is a number or None
            "is_forecast":  i > rep_idx,
            "value_pct":    round(actual / budget_tot * 100, 1) if budget_tot else 0.0,
    })

    # Narrative docs
    docx_files = sorted(DATA_DIR.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
    narratives: Dict[str, str] = {}
    if docx_files:
        narratives = parse_narrative_doc(docx_files[0])

    # Stories / points file
    stories_files = sorted(
        [p for p in DATA_DIR.glob("*.xlsx") if is_aux(p) and
         any(k in p.name.lower() for k in ("point","story","stories"))],
        key=lambda p: p.stat().st_mtime, reverse=True
    )
    stories_data: Dict = {}
    if stories_files:
        print(f"--- Reading stories: {stories_files[0].name}")
        stories_data = parse_stories_file(stories_files[0])

    reporting_month_full = format_month_year(reporting_month, current_file)
    awards_str = ", ".join(awards_found) if awards_found else "N/A"

    print(f"\nTotal milestones: {len(all_milestones)}  |  Awards: {awards_str}  |  Month: {reporting_month_full}")

    # Write JSON
    OUT_DIR.mkdir(exist_ok=True)
    (OUT_DIR / "dashboard_data.json").write_text(
        json.dumps({
            "summary": {
                "reporting_month": reporting_month_full,
                "awards": awards_str,
                "budget_total":  sum(m.budget        for m in all_milestones),
                "ytd_total":     sum(m.ytd_actual     for m in all_milestones),
                "month_total":   sum(m.monthly_spend  for m in all_milestones),
                "bal_total":     sum(m.est_balance     for m in all_milestones),
            },
            "milestones":  [asdict(m) for m in all_milestones],
            "monthlyData": monthly_chart,
        }, indent=2),
        encoding="utf-8",
    )

    generate_html(all_milestones, monthly_chart, reporting_month,
                  reporting_month_full, awards_str, narratives, stories_data)

    if AUTO_PUSH_GITHUB:
        try:
            sync_and_push(OUT_DIR, GIT_REPO_DIR, f"Dashboard update: {reporting_month_full}")
            print("GITHUB: pushed successfully")
        except Exception as e:
            print(f"GITHUB ERROR: {e}")


if __name__ == "__main__":
    process()