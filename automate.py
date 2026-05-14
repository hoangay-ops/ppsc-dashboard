#!/usr/bin/env python3
"""
PPSC Dashboard generator.
Run from the repo root:  python3 automate.py
"""
from __future__ import annotations

import json
import re
import subprocess
import urllib.request
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

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
ROOT     = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUT_DIR  = ROOT
ARCHIVE_DIR = None # This prevents the script from having a destination for archives

SITE_BASE_URL    = "https://hoangay-ops.github.io/ppsc-dashboard/"
ARCHIVE_BASE_URL = f"{SITE_BASE_URL}archive/"
AUTO_PUSH_GITHUB = True
GIT_REPO_DIR     = ROOT

# Award Year 2: Feb 1 2026 – Jan 31 2027
# Feb=1, Mar=2, Apr=3, May=4, Jun=5, Jul=6, Aug=7, Sep=8, Oct=9, Nov=10, Dec=11, Jan=12
AWARD_YEAR_TOTAL_MONTHS = 12
AWARD_MONTH_POSITION: Dict[str, int] = {
    "Feb":1,"Mar":2,"Apr":3,"May":4,"Jun":5,"Jul":6,
    "Aug":7,"Sep":8,"Oct":9,"Nov":10,"Dec":11,"Jan":12,
}

# Hardcoded forecast defaults (overridden by prior workbook if available)
FORECAST_DEFAULTS: Dict[str, float] = {
    "Jan":0.0, "Feb":0.0, "Mar":2_724_316.0, "Apr":2_660_405.0,
    "May":2_744_523.0, "Jun":2_503_623.0, "Jul":2_362_019.0,
    "Aug":3_737_302.0, "Sep":3_685_565.0, "Oct":608_801.0,
    "Nov":558_196.0,   "Dec":750_064.0,
}

FRONT_LOADED_MILESTONES: Dict[str, str] = {
    "53": "Annual license fees paid upfront in February — no further spend expected.",
}
COMPLETED_MILESTONES: Dict[str, str] = {}

# ─────────────────────────────────────────────
# COLUMN MAP (0-based)
# ─────────────────────────────────────────────
COL_TEXT       = 1
COL_PCT        = 2
COL_BUDGET     = 3
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


# ─────────────────────────────────────────────
# DATA MODEL
# ─────────────────────────────────────────────
@dataclass
class MilestoneRow:
    milestone_id:      str
    title:             str
    raw_title:         str
    award_number:      str
    status:            str
    budget:            float
    monthly_spend:     float
    percent_spent:     float
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
    exception_note:    str   = ""
    people_count:      int   = 0
    total_hours:       float = 0.0
    story_points:      float = 0.0
    workstreams:       list  = field(default_factory=list)


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def norm(v: Any) -> str:
    return str(v).replace("\xa0"," ").replace("\u2013","-").replace("\u2014","-").strip()

def sf(v: Any, d: float = 0.0) -> float:
    try:
        c = str(v).replace("$","").replace(",","").strip()
        return float(c) if c and c.lower() not in ("nan","none","") else d
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
        print(f"  WARNING fetch {url}: {e}")
        return ""

def is_temp_file(path: Path) -> bool:
    return path.name.startswith("~$")

def is_aux(path: Path) -> bool:
    return bool(re.search(
        r"hours|lcat|story|stories|points.*mile|mile.*points|points\s+by",
        path.name, re.IGNORECASE
    ))

def is_valid_xlsx(path: Path) -> bool:
    return (path.suffix.lower() == ".xlsx"
            and not is_temp_file(path)
            and path.stat().st_size > 5000)

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

def choose_sheet(xls: pd.ExcelFile) -> str:
    return xls.sheet_names[0] if len(xls.sheet_names) == 1 else xls.sheet_names[1]

def split_title(text: str) -> Tuple[Optional[str], str]:
    m = re.search(r"Milestone\s*(\d+)", text, re.IGNORECASE)
    if not m: return None, text
    mid   = f"M{m.group(1)}"
    after = text[m.end():].strip()
    raw   = re.sub(r"^[\s\-\u2014\u2013]+", "", after).strip()
    if not raw or re.match(r"^\d", raw): raw = ""
    return mid, raw

def months_elapsed(rep_month: str) -> int:
    return AWARD_MONTH_POSITION.get(rep_month, 1)

def yr_elapsed_pct(rep_month: str) -> float:
    return round(months_elapsed(rep_month) / AWARD_YEAR_TOTAL_MONTHS * 100, 1)

def compute_burn(budget: float, ytd: float, est_spend: float,
                 m_elapsed: int) -> Tuple[float, str, float, float]:
    expected = round(budget * m_elapsed / AWARD_YEAR_TOTAL_MONTHS, 2) if budget > 0 else 0.0
    variance = round(ytd - expected, 2)
    allow    = budget / AWARD_YEAR_TOTAL_MONTHS if budget > 0 else 0
    if allow > 0 and m_elapsed > 0:
        ratio = (ytd / m_elapsed) / allow
        bst = "Ahead" if ratio > 1.15 else "Behind" if ratio < 0.85 else "On Pace"
    else:
        bst = "On Pace"
    return expected, bst, variance, round(est_spend - budget, 2)

def derive_status(mid: str, est_bal: float, budget: float, bst: str,
                  month_act: float, ytd: float, rep_idx: int) -> str:
    if mid in COMPLETED_MILESTONES: return "Complete"
    if month_act == 0 and ytd > 0 and est_bal >= 0 and rep_idx >= 3: return "Complete"
    if mid in FRONT_LOADED_MILESTONES: return "On Track"
    if est_bal < 0: return "At Risk"
    if est_bal < max(budget * 0.02, 500.0) or bst == "Ahead": return "Watch"
    return "On Track"

def month_slug(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "-", s).strip("-")

def short_text(text: str, limit: int = 70) -> str:
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if len(text) <= limit: return text
    cut = text[:limit].rsplit(" ", 1)[0]
    return (cut or text[:limit]).rstrip(" .") + "\u2026"

def format_month_year(abbr: str, file_path: Path) -> str:
    full = datetime.strptime(abbr, "%b").strftime("%B")
    m = re.search(r"(20\d{2})", file_path.stem)
    yr = m.group(1) if m else str(datetime.now().year)
    return f"{full} {yr}"


# ─────────────────────────────────────────────
# WORKBOOK PARSING
# ─────────────────────────────────────────────
def read_main_workbook(path: Path) -> Tuple[pd.DataFrame, str]:
    xls   = pd.ExcelFile(path, engine="openpyxl")
    sheet = choose_sheet(xls)
    df    = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
    return df, detect_month(path, df)

def extract_award_number(df: pd.DataFrame, fallback: str) -> str:
    for _, r in df.head(30).iterrows():
        for ci in range(len(r)):
            if "award number" in norm(str(r.iloc[ci])).lower():
                for look in range(ci+1, min(ci+10, len(r))):
                    val = norm(str(r.iloc[look]))
                    if val and val.lower() not in ("nan","","none"):
                        return val
    for pat in [r"OT[\w]+", r"\d{4,5}\.\d+", r"[A-Z]{2}-[\w-]+"]:
        m = re.search(pat, fallback)
        if m: return m.group(0)
    return fallback

def extract_monthly_row_totals(df: pd.DataFrame) -> Dict[str, float]:
    totals: Dict[str, float] = {m: 0.0 for m in MONTH_ORDER}
    for _, r in df.iterrows():
        if norm(r[COL_TEXT]).startswith("Total Milestone"):
            for mo in MONTH_ORDER:
                totals[mo] += sf(r[MONTH_TO_COL[mo]])
    return totals

def parse_milestones(df: pd.DataFrame, rep_month: str,
                     award_number: str) -> List[MilestoneRow]:
    rep_col   = MONTH_TO_COL[rep_month]
    rep_idx   = MONTH_ORDER.index(rep_month)
    m_el      = months_elapsed(rep_month)
    rows: List[MilestoneRow] = []
    cur_id = cur_title = None
    accum: Dict[str, Dict[str, float]] = {}

    for _, r in df.iterrows():
        text = norm(r[COL_TEXT])
        if re.match(r"Milestone\s+\d+", text, re.IGNORECASE):
            cur_id, cur_title = split_title(text)
            accum = {c: {"budget":0.0,"month":0.0,"ytd":0.0} for c in LINE_ITEMS}
        elif text in LINE_ITEMS and cur_id:
            accum[text] = {"budget":sf(r[COL_BUDGET]),
                           "month": sf(r[rep_col]),
                           "ytd":   sf(r[COL_TOTAL_ACT])}
        elif text.startswith("Total Milestone") and cur_id:
            budget    = sf(r[COL_BUDGET])
            ytd       = sf(r[COL_TOTAL_ACT])
            est_spend = sf(r[COL_EST_SPEND])
            est_bal   = sf(r[COL_EST_BAL])
            month_act = sf(r[rep_col])
            pct       = round(ytd/budget*100, 1) if budget > 0 else 0.0
            exp_sp, bst, pace_v, overrun = compute_burn(budget, ytd, est_spend, m_el)
            mid       = cur_id.lstrip("M")
            exception = FRONT_LOADED_MILESTONES.get(mid,"") or COMPLETED_MILESTONES.get(mid,"")
            status    = derive_status(mid, est_bal, budget, bst, month_act, ytd, rep_idx)
            combined  = f"{cur_id} \u2014 {cur_title}" if cur_title else cur_id
            rows.append(MilestoneRow(
                milestone_id=cur_id, title=combined, raw_title=cur_title or cur_id,
                award_number=award_number, status=status, budget=budget,
                monthly_spend=month_act, percent_spent=pct,
                labor_actual=accum.get("Labor",{}).get("month",0.0),
                travel_actual=accum.get("Travel",{}).get("month",0.0),
                odc_actual=accum.get("ODC",{}).get("month",0.0),
                ytd_actual=ytd, total_forecast=sf(r[COL_TOTAL_FCST]),
                est_total_spend=est_spend, est_balance=est_bal,
                expected_spend=exp_sp, burn_rate_status=bst,
                pace_variance=pace_v, projected_overrun=overrun,
                exception_note=exception,
                workstreams=[{"name":c,
                              "month":accum.get(c,{}).get("month",0.0),
                              "ytd":  accum.get(c,{}).get("ytd",0.0),
                              "budget":accum.get(c,{}).get("budget",0.0)}
                             for c in LINE_ITEMS],
            ))
    return rows


# ─────────────────────────────────────────────
# HOURS PARSER
# ─────────────────────────────────────────────
def parse_hours_file(path: Path) -> Dict:
    try:
        df = pd.read_excel(path, sheet_name=0, header=None, engine="openpyxl")
    except Exception as e:
        print(f"  WARNING hours {path.name}: {e}"); return {}
    rows = df.values.tolist()
    if len(rows) < 3: return {}

    header_idx = 1
    for i in range(min(5, len(rows))):
        if re.search(r"\d{4,5}\.\d+|M\d+|milestone",
                     " ".join(str(v) for v in rows[i]), re.IGNORECASE):
            header_idx = i; break

    header    = [str(x).strip() if x is not None else "" for x in rows[header_idx]]
    mil_codes = header[1:-1]
    gt_col    = len(header) - 1

    def code_to_mid(c: str) -> str:
        m = re.search(r"\.(\d+)$", c)
        return f"M{m.group(1)}" if m else c

    mil_hours:  Dict[str, float] = {c: 0.0 for c in mil_codes}
    lcat_hours: Dict[str, float] = {}
    people_map: Dict[str, int]   = {code_to_mid(c): 0 for c in mil_codes}

    for col_idx, code in enumerate(mil_codes, 1):
        mid   = code_to_mid(code)
        count = sum(1 for row in rows[header_idx+1:]
                    if not str(row[0]).lower().startswith("grand total")
                    and col_idx < len(row) and sf(row[col_idx]) > 0)
        people_map[mid] = count

    for row in rows[header_idx+1:]:
        lcat = str(row[0]).strip() if row[0] is not None else ""
        if not lcat or lcat.lower().startswith("grand total"): continue
        gt = sf(row[gt_col]) if gt_col < len(row) else 0.0
        if gt == 0.0: continue
        lcat_hours[lcat] = lcat_hours.get(lcat, 0.0) + gt
        for col_idx, code in enumerate(mil_codes, 1):
            if col_idx < len(row):
                mil_hours[code] = mil_hours.get(code, 0.0) + sf(row[col_idx])

    total_hours  = sum(lcat_hours.values())
    by_milestone = sorted([{"milestone_id":code_to_mid(c),"hours":round(h,1)}
                            for c,h in mil_hours.items() if h > 0],
                          key=lambda x: x["hours"], reverse=True)
    by_lcat      = sorted([{"lcat":k,"hours":round(v,1)}
                            for k,v in lcat_hours.items() if v > 0],
                          key=lambda x: x["hours"], reverse=True)
    print(f"  Hours: {len(by_lcat)} LCATs, {total_hours:,.0f} hrs, {len(by_milestone)} milestones")
    return {"total_hours":round(total_hours,1),"total_lcats":len(by_lcat),
            "active_mils":len(by_milestone),"by_milestone":by_milestone,
            "by_lcat":by_lcat,"people_map":people_map}

def apply_hours_to_milestones(milestones: List[MilestoneRow], hours: Dict) -> None:
    if not hours: return
    pmap = hours.get("people_map", {})
    hmap = {x["milestone_id"]: x["hours"] for x in hours.get("by_milestone",[])}
    for m in milestones:
        m.people_count = pmap.get(m.milestone_id, 0)
        m.total_hours  = hmap.get(m.milestone_id, 0.0)


# ─────────────────────────────────────────────
# STORIES PARSER
# ─────────────────────────────────────────────
def parse_stories_file(path: Path) -> Dict:
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        print(f"  WARNING stories {path.name}: {e}"); return {}

    def find_col(cols, candidates):
        for c in cols:
            if any(k.lower() in c.lower() for k in candidates): return c
        return None

    by_milestone: Dict[str, float] = {}
    total_points = 0.0
    for sheet in xls.sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sheet, header=None,
                                nrows=15, engine="openpyxl")
            hdr = 0
            for i in range(len(raw)):
                if any(k in " ".join(str(v).lower() for v in raw.iloc[i])
                       for k in ("milestone","point","story")):
                    hdr = i; break
            df = pd.read_excel(path, sheet_name=sheet, header=hdr, engine="openpyxl")
            df.columns = [str(c).strip() for c in df.columns]
        except Exception: continue
        mil_col = find_col(df.columns, ["milestone","workstream","task","title","name","id"])
        pts_col = find_col(df.columns, ["point","points","story point","story points","sp"])
        if not mil_col or not pts_col: continue
        for _, row in df.iterrows():
            mil = str(row.get(mil_col,"")).strip()
            pts = sf(row.get(pts_col, 0))
            if not mil or mil.lower() in ("nan","none","") or pts == 0: continue
            m   = re.search(r"M?(\d+)", mil, re.IGNORECASE)
            key = f"M{m.group(1)}" if m else mil
            by_milestone[key] = by_milestone.get(key, 0.0) + pts
            total_points += pts
        break
    if not by_milestone: return {}
    return {"total_points":round(total_points,1),"total_mils":len(by_milestone),
            "by_milestone":sorted([{"milestone_id":k,"points":round(v,1)}
                                    for k,v in by_milestone.items()],
                                   key=lambda x: x["points"], reverse=True)}

def build_stories_panel(stories: Dict, all_milestones: List[MilestoneRow]) -> str:
    if not stories:
        return (
            '<section class="panel" style="margin-bottom:18px;">'
            '<h2>Story Points &amp; Velocity</h2>'
            '<p class="note">Add <em>Points by Milestones.xlsx</em> to data/ to populate.</p>'
            '</section>'
        )

    total = stories.get("total_points", 0)
    n_mils = stories.get("total_mils", 0)
    by_mil = stories.get("by_milestone", [])
    maxp = by_mil[0]["points"] if by_mil else 1
    tmap = {m.milestone_id: m.raw_title for m in all_milestones}

    def milestone_label(milestone_id: str) -> str:
        return f" — {tmap[milestone_id]}" if milestone_id in tmap else ""

    bars = "".join(
        f'<div style="margin-bottom:10px;">'
        f'<div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px;">'
        f'<span><strong>{it["milestone_id"]}</strong>{milestone_label(it["milestone_id"])}</span>'
        f'<span style="font-weight:700;">{it["points"]:,.0f} pts</span></div>'
        f'<div style="height:6px;background:rgba(15,23,42,0.08);border-radius:3px;overflow:hidden;">'
        f'<div style="width:{round(it["points"] / maxp * 100)}%;height:100%;background:#7c3aed;border-radius:3px;"></div>'
        f'</div></div>'
        for it in by_mil[:15]
    )

    return (
        f'<section class="panel" style="margin-bottom:18px;">'
        f'<h2>Story Points &amp; Velocity</h2>'
        f'<p class="note">Point totals by milestone from <em>Points by Milestones.xlsx</em>.</p>'
        f'<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:14px;margin-bottom:18px;">'
        f'<div style="background:rgba(124,58,237,0.06);border:1px solid rgba(124,58,237,0.18);'
        f'border-radius:14px;padding:14px 16px;">'
        f'<div class="label">Total Points</div><div class="value">{total:,.0f}</div></div>'
        f'<div style="background:rgba(124,58,237,0.06);border:1px solid rgba(124,58,237,0.18);'
        f'border-radius:14px;padding:14px 16px;">'
        f'<div class="label">Milestones with points</div><div class="value">{n_mils}</div></div></div>'
        f'<h3 style="font-size:15px;margin-bottom:10px;">Points by milestone</h3>'
        f'{bars}</section>'
    )

# ─────────────────────────────────────────────
# NARRATIVE
# ─────────────────────────────────────────────
def parse_narrative_doc(path: Path) -> Dict[str, str]:
    if not DOCX_OK: return {}
    try: doc = DocxDocument(path)
    except Exception as e:
        print(f"  WARNING docx {path.name}: {e}"); return {}
    narratives: Dict[str, str] = {}
    cur_sec, cur_lines = None, []
    def flush():
        if cur_sec and cur_lines:
            narratives[cur_sec] = " ".join(cur_lines).strip()
    for para in doc.paragraphs:
        t = para.text.strip()
        if not t: continue
        is_mil  = bool(re.search(r"^\b(Milestone|M)\s*\d+", t, re.IGNORECASE))
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
        if re.search(rf"\b{mid}\b", sec, re.IGNORECASE): return txt
    return ""


# ─────────────────────────────────────────────
# ARCHIVE
# ─────────────────────────────────────────────
def archive_manifest(archive_dir: Path) -> List[Dict]:
    items = []
    for p in archive_dir.glob("*.html"):
        if p.name == "index.html": continue
        try:   dt = datetime.strptime(p.stem, "%B-%Y")
        except ValueError: continue
        items.append({"label":dt.strftime("%B %Y"),"file":p.name,"sort_key":dt})
    items.sort(key=lambda x: x["sort_key"])
    return items

def build_archive_index(archive_dir: Path, current_month_full: str) -> None:
    items = archive_manifest(archive_dir)
    cards = "".join(
        f'<a class="month-card" href="{it["file"]}">'
        f'<div class="month-title">{it["label"]}'
        f'{"  (current)" if it["label"]==current_month_full else ""}</div>'
        f'<div class="month-copy">Open archived snapshot</div></a>'
        for it in items
    ) or '<div style="color:#51627d;">No archives yet.</div>'
    html = (
        '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"/>'
        '<meta name="viewport" content="width=device-width,initial-scale=1.0"/>'
        '<title>PPSC Archive</title><style>'
        ':root{--shadow:0 18px 40px rgba(15,23,42,.10);--radius:22px;}'
        '*{box-sizing:border-box;}'
        'body{margin:0;font-family:Inter,ui-sans-serif,sans-serif;color:#10203a;'
        'background:linear-gradient(180deg,#fff 0%,#f6f8fc 52%,#eef3fb 100%);min-height:100vh;}'
        '.wrap{max-width:1300px;margin:0 auto;padding:14px;}'
        '.hero,.panel{background:linear-gradient(180deg,rgba(255,255,255,.96),rgba(248,250,252,.98));'
        'border:1px solid rgba(15,23,42,.10);box-shadow:var(--shadow);border-radius:var(--radius);}'
        '.hero{padding:26px;margin-bottom:18px;}.panel{padding:18px;}'
        '.header-bar{display:flex;justify-content:flex-end;margin-bottom:12px;}'
        '.header-actions{display:flex;gap:10px;}'
        '.header-actions a{font-size:12px;padding:8px 12px;border-radius:6px;text-decoration:none;'
        'border:1px solid #cbd5e1;background:#fff;color:#10203a;}'
        '.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-top:16px;}'
        '.month-card{display:block;text-decoration:none;color:#10203a;padding:18px;border-radius:var(--radius);}'
        '.month-card:hover{background:rgba(37,99,235,.03);}'
        '.month-title{font-size:18px;font-weight:800;margin-bottom:6px;}'
        '.month-copy{color:#51627d;font-size:13px;}'
        '@media(max-width:1100px){.grid{grid-template-columns:1fr;}}'
        '</style></head><body>'
        '<div class="wrap">'
        '<div class="header-bar"><div class="header-actions">'
        '<a href="../index.html">Latest</a><a href="./">Archive</a>'
        '</div></div>'
        '<section class="hero"><h1>PPSC Dashboard Archive</h1>'
        '<p style="margin-top:10px;color:#51627d;">Monthly snapshots.</p>'
        '</section>'
        f'<section class="panel"><div class="grid">{cards}</div></section>'
        '</div></body></html>'
    )
  # (archive_dir / "index.html").write_text(html, encoding="utf-8")

def git_push(repo_dir: Path, message: str) -> None:
    # 1. Get latest from GitHub
    subprocess.run(["git", "-C", str(repo_dir), "fetch", "origin"], check=True)
    # 2. Merge changes instead of resetting (this keeps your new index.html)
    subprocess.run(["git", "-C", str(repo_dir), "pull", "origin", "main", "--rebase"], check=True)
    # 3. Stage the new file
    subprocess.run(["git", "-C", str(repo_dir), "add", "index.html"], check=True)
    # 4. Commit and Push
    result = subprocess.run(["git", "-C", str(repo_dir), "commit", "-m", message], capture_output=True, text=True)
    subprocess.run(["git", "-C", str(repo_dir), "push"], check=True)

def sync_and_push(out_dir: Path, repo_dir: Path, message: str) -> None:
    print(f"🚀 Syncing dashboard to GitHub: {message}")
    try:
        git_push(repo_dir, message)
        print("✅ Successfully pushed to GitHub Pages.")
    except Exception as e:
        print(f"❌ Git Push failed: {e}")
    
    # Commit and Push
    result = subprocess.run(["git", "-C", str(repo_dir), "commit", "-m", message], capture_output=True, text=True)
    # ... rest of the function

# ─────────────────────────────────────────────
# HTML GENERATION
# ─────────────────────────────────────────────
def build_chart_section(chartjs: str, monthly_json: str, m_elapsed: int,
                         on_track: list, on_watch: list, at_risk: list,
                         watchlist_html: str) -> str:
    """
    Build the chart + status mix section as a plain string.
    Avoids f-string issues with </script> and JS braces.
    """
    sc = '<' + '/script>'   # closing script tag — never inside an f-string

    chart_js_block = (
        '  <script>' + chartjs + sc + '\n'
        '  <script>\n'
        '    (function() {\n'
        '      var md     = ' + monthly_json + ';\n'
        '      var labels = md.map(function(d){ return d.month; });\n'
        '      var actual = md.map(function(d){ return (d.value !== null && d.value > 0) ? d.value : null; });\n'
        '      var fcst   = md.map(function(d){ return (d.forecast_val !== null && d.forecast_val > 0) ? d.forecast_val : null; });\n'
        '      var canvas = document.getElementById(\'spendChart\');\n'
        '      if (!canvas) return;\n'
        '      new Chart(canvas, {\n'
        '        type:\'bar\',\n'
        '        data:{\n'
        '          labels:labels,\n'
        '          datasets:[\n'
        '            {\n'
        '              label:\'Actual Spend\', data:actual,\n'
        '              backgroundColor:\'rgba(37,99,235,0.85)\',\n'
        '              borderRadius:4, borderSkipped:false, order:1,\n'
        '              categoryPercentage:0.8, barPercentage:0.5\n'
        '            },\n'
        '            {\n'
        '              label:\'Forecast\', data:fcst,\n'
        '              backgroundColor:function(ctx){\n'
        '                var idx=ctx.dataIndex;\n'
        '                return (md[idx] && md[idx].is_forecast)\n'
        '                  ? \'rgba(217,119,6,0.40)\'\n'
        '                  : \'rgba(217,119,6,0.80)\';\n'
        '              },\n'
        '              borderRadius:4, borderSkipped:false, order:2,\n'
        '              categoryPercentage:0.8, barPercentage:0.5\n'
        '            }\n'
        '          ]\n'
        '        },\n'
        '        options:{\n'
        '          responsive:true, maintainAspectRatio:false,\n'
        '          interaction:{mode:\'index\',intersect:false},\n'
        '          plugins:{\n'
        '            legend:{display:false},\n'
        '            tooltip:{callbacks:{\n'
        '              label:function(ctx){\n'
        '                if(ctx.parsed.y===null) return null;\n'
        '                return ctx.dataset.label+\': $\'+Math.abs(ctx.parsed.y).toLocaleString(undefined,{maximumFractionDigits:0});\n'
        '              },\n'
        '              afterBody:function(items){\n'
        '                var av=null,fv=null;\n'
        '                items.forEach(function(item){\n'
        '                  if(item.dataset.label===\'Actual Spend\') av=item.parsed.y;\n'
        '                  if(item.dataset.label===\'Forecast\')     fv=item.parsed.y;\n'
        '                });\n'
        '                if(av!==null && fv!==null){\n'
        '                  var d=av-fv, s=d>=0?\'+\':\'\';\n'
        '                  return [\'Variance: \'+s+\'$\'+Math.round(d).toLocaleString()];\n'
        '                }\n'
        '                return [];\n'
        '              }\n'
        '            }}\n'
        '          },\n'
        '          scales:{\n'
        '            x:{ticks:{autoSkip:false,maxRotation:0,font:{size:11}}},\n'
        '            y:{ticks:{callback:function(v){\n'
        '              if(v>=1000000) return \'$\'+(v/1000000).toFixed(1)+\'M\';\n'
        '              if(v>=1000)    return \'$\'+Math.round(v/1000)+\'k\';\n'
        '              return \'$\'+v;\n'
        '            },font:{size:11}},min:0}\n'
        '          }\n'
        '        }\n'
        '      });\n'
        '    })();\n'
        '  ' + sc
    )

    html_part = (
        '\n  <section class="grid" style="margin-bottom:18px;">\n'
        '    <div class="panel">\n'
        '      <h2>Monthly Spend \u2014 Actuals vs. Forecast</h2>\n'
        '      <p class="note">Past months: blue (actual) + orange (forecast) side by side. '
        'Future months: orange forecast only. Hover for variance.</p>\n'
        '      <div style="display:flex;gap:16px;margin-bottom:8px;font-size:12px;color:#51627d;">\n'
        '        <span style="display:flex;align-items:center;gap:4px;">'
        '<span style="width:10px;height:10px;border-radius:2px;background:#2563eb;display:inline-block;"></span>'
        'Actual spend</span>\n'
        '        <span style="display:flex;align-items:center;gap:4px;">'
        '<span style="width:10px;height:10px;border-radius:2px;background:#d97706;display:inline-block;"></span>'
        'Forecast</span>\n'
        '      </div>\n'
        '      <div style="position:relative;width:100%;height:280px;">\n'
        '        <canvas id="spendChart" role="img" aria-label="Monthly spend vs forecast"></canvas>\n'
        '      </div>\n'
        f'      <div class="footer" style="margin-top:8px;">Award Year 2: Feb 1 2026 \u2013 Jan 31 2027 \u00b7 Month {m_elapsed} of {AWARD_YEAR_TOTAL_MONTHS} elapsed.</div>\n'
        '    </div>\n'
        '    <div class="panel" style="display:flex;flex-direction:column;gap:14px;">\n'
        '      <div>\n'
        '        <h2 style="margin-bottom:6px;">Status Mix</h2>\n'
        '        <p class="note" style="margin-bottom:10px;">On Track = within budget \u00b7 Watch = balance tightening \u00b7 At Risk = projected overrun.</p>\n'
        '        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;">\n'
        f'          <div class="status-chip status-good"><div class="txt">On Track</div><div class="num">{len(on_track)}</div></div>\n'
        f'          <div class="status-chip status-watch"><div class="txt">Watch</div><div class="num">{len(on_watch)}</div></div>\n'
        f'          <div class="status-chip status-risk"><div class="txt">At Risk</div><div class="num">{len(at_risk)}</div></div>\n'
        '        </div>\n'
        '      </div>\n'
        '      <div style="flex:1;">\n'
        '        <h2 style="margin-bottom:4px;">Priority Watchlist</h2>\n'
        '        <p class="note" style="margin-bottom:8px;">Sorted by severity and projected overrun. Use this list to focus the next leadership discussion.</p>\n'
        f'        <div class="bullets">{watchlist_html}</div>\n'
        '      </div>\n'
        '    </div>\n'
        '  </section>\n'
        + chart_js_block
    )
    return html_part


def generate_html(
    all_milestones:       List[MilestoneRow],
    monthly_chart:        List[Dict],
    reporting_month:      str,
    reporting_month_full: str,
    awards_str:           str,
    narratives:           Dict[str, str],
    stories_data:         Dict,
    hours_data:           Dict,
) -> None:
    template_path = next(
        (p for p in [ROOT/"dashboard_template_fixed.html", ROOT/"dashboard_template.html"]
         if p.exists()), None
    )
    if not template_path:
        print("ERROR: dashboard_template.html not found"); return

    html = template_path.read_text(encoding="utf-8")
if ARCHIVE_DIR:
    ARCHIVE_DIR.mkdir(exist_ok=True)
    print("  Fetching Chart.js\u2026")
    chartjs     = fetch_js("https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js")
    print("  Fetching html2pdf\u2026")
    html2pdf_js = fetch_js("https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js")

    at_risk  = [m for m in all_milestones if m.status == "At Risk"]
    on_watch = [m for m in all_milestones if m.status == "Watch"]
    on_track = [m for m in all_milestones if m.status in ("On Track","Complete")]

    budget_tot = sum(m.budget        for m in all_milestones)
    ytd_tot    = sum(m.ytd_actual    for m in all_milestones)
    month_tot  = sum(m.monthly_spend for m in all_milestones)
    bal_tot    = sum(m.est_balance   for m in all_milestones)

    rep_idx         = MONTH_ORDER.index(reporting_month)
    m_el            = months_elapsed(reporting_month)
    budget_used_pct = round(ytd_tot/budget_tot*100, 1) if budget_tot else 0.0
    yr_pct          = yr_elapsed_pct(reporting_month)
    pace_gap        = round(budget_used_pct - yr_pct, 1)
    balance_color   = "#dc2626" if bal_tot < 0 else "#059669"

    prev_val    = monthly_chart[rep_idx-1]["value"] or month_tot if rep_idx > 0 else month_tot
    trend_pct   = round((month_tot - prev_val)/prev_val*100, 1) if prev_val else 0.0
    trend_arrow = "\u25b2" if trend_pct>0 else "\u25bc" if trend_pct<0 else "\u2192"
    trend_color = "#dc2626" if trend_pct>0 else "#059669" if trend_pct<0 else "#51627d"

    flagged = at_risk + on_watch
    n_flags = len(flagged)
    n_crit  = len(at_risk)
    flag_summary = (
        "No milestone flags need escalation." if n_flags==0 else
        f"{n_flags} flagged, all minor \u2014 monitor closely." if n_crit==0 else
        f"{n_flags} flagged \u2014 {n_crit} require{'s' if n_crit==1 else ''} attention."
    )
    tension_msg = (
        f"Budget used is {budget_used_pct:.1f}% versus {yr_pct:.1f}% of the award year elapsed "
        f"(Feb 2026 \u2013 Jan 2027). {len(at_risk)} milestone(s) At Risk and {len(on_watch)} on Watch."
    )

    watchlist = sorted(
        [m for m in all_milestones if m.status in ("At Risk","Watch") and not m.exception_note],
        key=lambda m: (m.status=="At Risk", abs(m.projected_overrun)), reverse=True
    )[:5]
    watchlist_html = "\n".join(
        bullet_html("risk" if m.status=="At Risk" else "warn",
            f"<strong>{m.title}</strong> \u2014 {fmt_usd_short(abs(m.projected_overrun))} projected overrun"
            + (f" \u2014 {short_text(narrative_for(narratives,m.title))}"
               if narrative_for(narratives,m.title) else "") + ".")
        for m in watchlist
    ) or bullet_html("good","No milestones currently flagged for attention.")

    callouts = [
        bullet_html("risk" if bal_tot<0 else "good",
            f"<strong>Portfolio balance is {'negative' if bal_tot<0 else 'positive'} "
            f"at {fmt_usd(bal_tot)}.</strong> "
            f"{'Projected spend exceeds budget.' if bal_tot<0 else 'Budget on track.'}"),
        bullet_html("warn" if n_flags else "good", f"<strong>{flag_summary}</strong>"),
    ]
    for m in at_risk[:3]:
        callouts.append(bullet_html("risk",
            f"<strong>{m.title} \u2014 {fmt_usd_short(abs(m.projected_overrun))} overrun</strong>"
            + (f" \u2014 {short_text(narrative_for(narratives,m.title))}"
               if narrative_for(narratives,m.title) else "") + "."))
    callout_html = "\n".join(callouts)

    op_items = []
    for m in sorted([m for m in all_milestones if m.monthly_spend<0],
                    key=lambda m: m.monthly_spend):
        op_items.append(bullet_html("good",
            f"<strong>Credit \u2014 {m.title}:</strong> {fmt_usd(abs(m.monthly_spend))} posted."))
    for m in sorted([m for m in all_milestones
                     if m.monthly_spend>0 and m.budget>0 and m.monthly_spend/m.budget>0.25],
                    key=lambda m: m.monthly_spend/m.budget, reverse=True):
        p = round(m.monthly_spend/m.budget*100, 1)
        op_items.append(bullet_html("warn",
            f"<strong>Spend spike \u2014 {m.title}:</strong> "
            f"{fmt_usd(m.monthly_spend)} this month ({p}% of annual budget)."))
    if not op_items:
        for m in sorted(all_milestones, key=lambda m: m.monthly_spend, reverse=True)[:5]:
            op_items.append(bullet_html("good",
                f"<strong>{m.title}</strong> \u2014 {fmt_usd(m.monthly_spend)} this month "
                f"({m.percent_spent:.1f}% of budget spent to date)."))
    operational_html = "\n".join(op_items) or bullet_html("good","No spend data.")

    # Hours panel
    if hours_data:
        h_total = hours_data.get("total_hours",0)
        h_lcats = hours_data.get("total_lcats",0)
        h_mils  = hours_data.get("active_mils",0)
        h_bm    = hours_data.get("by_milestone",[])
        h_bl    = hours_data.get("by_lcat",[])
        maxm    = h_bm[0]["hours"] if h_bm else 1
        maxl    = h_bl[0]["hours"] if h_bl else 1
        tmap    = {m.milestone_id: m.raw_title for m in all_milestones}

        def hbar(lbl, val, mx, color):
            pct = round(val/mx*100) if mx else 0
            return (f'<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:12px;">'
                    f'<div style="width:130px;flex:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#51627d;" title="{lbl}">{lbl}</div>'
                    f'<div style="flex:1;height:8px;background:rgba(15,23,42,.08);border-radius:4px;overflow:hidden;">'
                    f'<div style="width:{pct}%;height:100%;background:{color};border-radius:4px;"></div></div>'
                    f'<div style="width:52px;text-align:right;color:#51627d;">{val:,.1f}</div></div>')

        mil_bars  = "".join(hbar(
            f'{r["milestone_id"]}' + (f' \u2014 {tmap[r["milestone_id"]]}' if r["milestone_id"] in tmap else ""),
            r["hours"], maxm, "#2563eb") for r in h_bm[:8])
        lcat_bars = "".join(hbar(r["lcat"], r["hours"], maxl, "#059669") for r in h_bl[:8])

        hours_panel_html = (
            f'<section class="panel" style="margin-bottom:18px;">'
            f'<h2>NORC People &amp; Hours \u2014 {reporting_month_full}</h2>'
            f'<p class="note">Direct labor charged by LCAT across all active milestones.</p>'
            f'<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:16px;">'
            f'<div class="card" style="min-height:auto;padding:14px;"><div class="label">Staff charged</div>'
            f'<div class="value">{h_lcats}</div><div class="delta">unique LCATs billing</div></div>'
            f'<div class="card" style="min-height:auto;padding:14px;"><div class="label">Total hours</div>'
            f'<div class="value">{h_total:,.0f}</div><div class="delta">across all milestones</div></div>'
            f'<div class="card" style="min-height:auto;padding:14px;"><div class="label">Active milestones</div>'
            f'<div class="value">{h_mils}</div><div class="delta">with hours charged</div></div></div>'
            f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;">'
            f'<div><div style="font-size:13px;font-weight:700;margin-bottom:8px;">Hours by milestone</div>{mil_bars}</div>'
            f'<div><div style="font-size:13px;font-weight:700;margin-bottom:8px;">Hours by LCAT</div>{lcat_bars}</div>'
            f'</div></section>'
        )
    else:
        hours_panel_html = ('<section class="panel" style="margin-bottom:18px;">'
                            '<h2>NORC People &amp; Hours</h2>'
                            '<p class="note">Add <em>9877 Hours by LCAT*.xlsx</em> to data/ to populate.</p>'
                            '</section>')

    monthly_json = json.dumps(monthly_chart)
    chart_section = build_chart_section(
        chartjs, monthly_json, m_el, on_track, on_watch, at_risk, watchlist_html
    )

    start = html.find('<section class="grid">')
    end   = html.find('</section>', start) + len('</section>') if start != -1 else -1
    if start != -1 and end != -1:
        html = html[:start] + chart_section + html[end:]

    milestone_span = "N/A"
    if all_milestones:
        ids = [m.milestone_id for m in all_milestones]
        milestone_span = f"{ids[0]}\u2013{ids[-1]}"

    slug         = month_slug(reporting_month_full)
    current_file = f"{slug}.html"
    sc_tag       = '<' + '/script>'
    pdf_js_tag   = '<script>' + html2pdf_js + sc_tag

    replacements = {
        "__MONTH__":            reporting_month_full,
        "__BALANCE_COLOR__":    balance_color,
        "__TITLE__":            f"PPSC Financial Dashboard \u2014 {reporting_month_full}",
        "__SUBTITLE__":         f"Award(s): {awards_str} \u00b7 {len(all_milestones)} milestones \u00b7 {reporting_month_full}.",
        "__AWARD__":            awards_str,
        "__MILESTONE_SPAN__":   milestone_span,
        "__AT_A_GLANCE__":      tension_msg,
        "__YTD_SPEND_TOTAL__":  fmt_usd(ytd_tot),
        "__AWARD_BUDGET__":     fmt_usd(budget_tot),
        "__MONTHLY_SPEND__":    fmt_usd(month_tot),
        "__ON_TRACK__":         str(len(on_track)),
        "__WATCH_ITEMS__":      str(len(at_risk)),
        "__TREND_NOTE__":       "Monthly spend \u2014 blue = actuals, orange = forecast.",
        "__TREND_FOOTER__":     f"Award Year 2: Feb 2026 \u2013 Jan 2027. Month {m_el} of {AWARD_YEAR_TOTAL_MONTHS}.",
        "__STATUS_NOTE__":      "On Track = within budget \u00b7 Watch = tightening \u00b7 At Risk = overrun.",
        "__COUNT_ON_TRACK__":   str(len(on_track)),
        "__COUNT_AT_RISK__":    str(len(at_risk)),
        "__COUNT_AWAITING__":   str(len(on_watch)),
        "__COUNT_NOT_STARTED__":"0",
        "__WATCHLIST_HTML__":   watchlist_html,
        "__TENSION_MSG__":      tension_msg,
        "__OPERATIONAL_NOTE__": "Credits, spend spikes, and notable activity this period.",
        "__OPERATIONAL_HTML__": operational_html,
        "__CALLOUT_NOTE__":     "Action items for At Risk and Watch milestones.",
        "__CALLOUT_HTML__":     callout_html,
        "__CONSOLIDATED_CALLOUTS_HTML__": callout_html,
        "__NARRATIVE_PANEL__":  "",
        "__FOOTER__":           (f"Generated {pd.Timestamp.now().strftime('%B %d, %Y %H:%M')} "
                                 f"\u00b7 {reporting_month_full} \u00b7 Award(s): {awards_str}"),
        "__MONTHLY_DATA_JSON__":  monthly_json,
        "__MILESTONES_JSON__":    json.dumps([asdict(m) for m in all_milestones]),
        "__HTML2PDF_JS__":        pdf_js_tag,
        "__CURRENT_FILE__":       current_file,
        "__SITE_BASE_URL__":      SITE_BASE_URL,
        "__ARCHIVE_BASE_URL__":   ARCHIVE_BASE_URL,
        "__MONTHS_JSON_PATH__":   f"{ARCHIVE_BASE_URL}months.json",
        "__BUDGET_USED_PCT__":    f"{budget_used_pct:.1f}",
        "__YEAR_ELAPSED_PCT__":   f"{yr_pct:.1f}",
        "__BUDGET_PACE_GAP__":    f"{pace_gap:+.1f}",
        "__TREND_ARROW__":        trend_arrow,
        "__TREND_PCT__":          f"{abs(trend_pct):.1f}",
        "__TREND_COLOR__":        trend_color,
        "__FLAGGED_COUNT__":      str(n_flags),
        "__FLAGGED_SUMMARY__":    flag_summary,
        "__MINOR_FLAG_COUNT__":   str(max(n_flags-n_crit,0)),
        "__PORTFOLIO_BALANCE__":  fmt_usd(bal_tot),
        "__TREND_TEXT__":         f"{trend_arrow} {abs(trend_pct):.1f}% vs prior month.",
        "__DECISION_1__":         (f"{n_crit} milestone(s) need immediate review."
                                   if n_crit else "No immediate decisions required."),
        "__DECISION_2__":         f"Budget used {budget_used_pct:.1f}% vs {yr_pct:.1f}% of award year (Feb 2026 \u2013 Jan 2027).",
        "__DECISION_3__":         "Confirm whether front-loaded milestones should remain excluded from pace alerts.",
        "__STORIES_PANEL__":      build_stories_panel(stories_data, all_milestones),
        "__HOURS_PANEL__":        hours_panel_html,
        "__WHAT_VISUALIZED__":    "Monthly actuals vs. forecast, spend trend, risk flags.",
    }

    for k, v in replacements.items():
        html = html.replace(k, v)

    remaining = re.findall(r"__[A-Z_]+__", html)
    if remaining:
        print(f"  WARNING unreplaced placeholders: {set(remaining)}")

    latest_path  = OUT_DIR / "index.html"
  # archive_path = ARCHIVE_DIR / f"{slug}.html"
    latest_path.write_text(html,  encoding="utf-8")
    archive_path.write_text(html, encoding="utf-8")

    months = archive_manifest(ARCHIVE_DIR)
    (ARCHIVE_DIR / "months.json").write_text(
        json.dumps([{"label":it["label"],"file":it["file"]} for it in months], indent=2),
        encoding="utf-8"
    )
    build_archive_index(ARCHIVE_DIR, reporting_month_full)
    print(f"  LATEST:   {latest_path}")
    print(f"  ARCHIVED: {archive_path}")


# ─────────────────────────────────────────────
# MAIN PROCESS
# ─────────────────────────────────────────────
def process() -> None:
    print(">>> PPSC Dashboard \u2014 starting")

    all_xlsx = sorted(
        [p for p in DATA_DIR.glob("*.xlsx") if is_valid_xlsx(p)],
        key=lambda p: p.stat().st_mtime
    )
    if not all_xlsx:
        print("ERROR: No valid .xlsx files in data/"); return

    main_files = [p for p in all_xlsx if not is_aux(p)]
    aux_files  = [p for p in all_xlsx if is_aux(p)]
    if not main_files:
        print("ERROR: No main workbooks found."); return

    print(f"Main: {[p.name for p in main_files]}")
    print(f"Aux:  {[p.name for p in aux_files]}")

    current_file     = main_files[-1]
    all_milestones:  List[MilestoneRow]   = []
    awards_found:    List[str]            = []
    monthly_actuals: Dict[str, float]     = {m: 0.0 for m in MONTH_ORDER}
    reporting_month  = "Jan"

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
        print(f"  {len(mils)} milestones (award: {award}, month: {rep_month})")
        all_milestones.extend(mils)
        if path == current_file:
            for mo in MONTH_ORDER:
                monthly_actuals[mo] += extract_monthly_row_totals(df).get(mo, 0.0)

    # Forecast: prior workbook overrides defaults where non-zero
    prior_forecast: Dict[str, float] = dict(FORECAST_DEFAULTS)
    if len(main_files) >= 2:
        print(f"--- Prior workbook forecast: {main_files[-2].name}")
        try:
            prior_df, _ = read_main_workbook(main_files[-2])
            for mo, val in extract_monthly_row_totals(prior_df).items():
                if val > 0:
                    prior_forecast[mo] = val
        except Exception as e:
            print(f"  WARNING: {e} \u2014 using hardcoded defaults.")
    else:
        print("  Single workbook \u2014 using hardcoded forecast defaults.")

    # Hours
    hours_data: Dict = {}
    hours_files = sorted([p for p in aux_files
                          if "hours" in p.name.lower() or "lcat" in p.name.lower()],
                         key=lambda p: p.stat().st_mtime)
    if hours_files:
        print(f"--- Hours: {hours_files[-1].name}")
        hours_data = parse_hours_file(hours_files[-1])
        apply_hours_to_milestones(all_milestones, hours_data)

    # Stories
    stories_data: Dict = {}
    stories_files = sorted([p for p in aux_files
                             if any(k in p.name.lower() for k in ("point","story","stories"))],
                           key=lambda p: p.stat().st_mtime)
    if stories_files:
        print(f"--- Stories: {stories_files[-1].name}")
        stories_data = parse_stories_file(stories_files[-1])

    # Narrative
    narratives: Dict[str, str] = {}
    docx_files = [p for p in sorted(DATA_DIR.glob("*.docx"),
                                    key=lambda p: p.stat().st_mtime, reverse=True)
                  if not p.name.startswith("~$")]
    if docx_files:
        print(f"--- Narrative: {docx_files[0].name}")
        narratives = parse_narrative_doc(docx_files[0])

    rep_idx    = MONTH_ORDER.index(reporting_month)
    budget_tot = sum(m.budget for m in all_milestones)

    monthly_chart = []
    for i, mo in enumerate(MONTH_ORDER):
        actual   = monthly_actuals[mo]
        forecast = prior_forecast.get(mo, 0.0)
        monthly_chart.append({
            "month":        mo,
            "value":        round(actual,2) if actual > 0 else None,
            "forecast_val": round(forecast,2) if forecast > 0 else None,
            "is_forecast":  i > rep_idx,
            "value_pct":    round(actual/budget_tot*100,1) if budget_tot and actual>0 else 0.0,
        })

    reporting_month_full = format_month_year(reporting_month, current_file)
    awards_str = ", ".join(awards_found) if awards_found else "N/A"

    at_r = sum(1 for m in all_milestones if m.status=="At Risk")
    wch  = sum(1 for m in all_milestones if m.status=="Watch")
    ont  = sum(1 for m in all_milestones if m.status in ("On Track","Complete"))
    print(f"\nMilestones: {len(all_milestones)} | Awards: {awards_str} | Month: {reporting_month_full}")
    print(f"Status: {ont} On Track \u00b7 {wch} Watch \u00b7 {at_r} At Risk")

    (ROOT / "output").mkdir(exist_ok=True)
    (ROOT / "output" / "dashboard_data.json").write_text(
        json.dumps({
            "summary": {"reporting_month":reporting_month_full,"awards":awards_str,
                        "budget_total":sum(m.budget for m in all_milestones),
                        "ytd_total":sum(m.ytd_actual for m in all_milestones),
                        "month_total":sum(m.monthly_spend for m in all_milestones),
                        "bal_total":sum(m.est_balance for m in all_milestones)},
            "milestones":  [asdict(m) for m in all_milestones],
            "monthlyData": monthly_chart,
        }, indent=2), encoding="utf-8"
    )

    generate_html(all_milestones, monthly_chart, reporting_month,
                  reporting_month_full, awards_str, narratives,
                  stories_data, hours_data)

    if AUTO_PUSH_GITHUB:
        try:
            # Re-generate files after git reset wipes them
            generate_html(all_milestones, monthly_chart, reporting_month,
                          reporting_month_full, awards_str, narratives,
                          stories_data, hours_data)
            git_push(GIT_REPO_DIR, f"Dashboard update: {reporting_month_full}")
            print("GITHUB: pushed successfully")
        except Exception as e:
            print(f"GITHUB ERROR: {e}")


if __name__ == "__main__":
    process()