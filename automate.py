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
ROOT        = Path(__file__).resolve().parent
DATA_DIR    = ROOT / "data"
OUT_DIR     = ROOT
ARCHIVE_DIR = ROOT / "archive"

SITE_BASE_URL    = "https://hoangay-ops.github.io/ppsc-dashboard/"
ARCHIVE_BASE_URL = f"{SITE_BASE_URL}archive/"
AUTO_PUSH_GITHUB = True
GIT_REPO_DIR     = ROOT

# Award Year 2: Feb 1 2026 – Jan 31 2027
AWARD_YEAR_TOTAL_MONTHS = 12
AWARD_MONTH_POSITION: Dict[str, int] = {
    "Feb":1,"Mar":2,"Apr":3,"May":4,"Jun":5,"Jul":6,
    "Aug":7,"Sep":8,"Oct":9,"Nov":10,"Dec":11,"Jan":12,
}

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

# Credit/adjustment keywords for narrative scanning
CREDIT_PATTERNS = [
    re.compile(r"\bcredit\b",        re.IGNORECASE),
    re.compile(r"\badjustment\b",    re.IGNORECASE),
    re.compile(r"\bwrite[- ]?off\b", re.IGNORECASE),
    re.compile(r"\brefund\b",        re.IGNORECASE),
    re.compile(r"\brecovery\b",      re.IGNORECASE),
    re.compile(r"\brevers",          re.IGNORECASE),
]

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

def _safe_json(obj: Any) -> str:
    """JSON-encode and escape '</' so browsers don't close <script> tags early."""
    return json.dumps(obj).replace("</", "<\\/")

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
        bst   = "Ahead" if ratio > 1.15 else "Behind" if ratio < 0.85 else "On Pace"
    else:
        bst = "On Pace"
    return expected, bst, variance, round(est_spend - budget, 2)

def derive_status(mid: str, est_bal: float, budget: float, bst: str,
                  month_act: float, ytd: float, rep_idx: int) -> str:
    if mid in COMPLETED_MILESTONES: return "Complete"
    if month_act == 0 and ytd > 0 and est_bal >= 0 and rep_idx >= 3: return "Complete"
    if mid in FRONT_LOADED_MILESTONES: return "On Track"
    if est_bal < 0: return "At Risk"
    if est_bal < max(budget * 0.02, 500.0) and bst == "Ahead": return "Watch"
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
    m    = re.search(r"(20\d{2})", file_path.stem)
    yr   = m.group(1) if m else str(datetime.now().year)
    return f"{full} {yr}"

def is_credit_text(text: str) -> bool:
    return any(p.search(text) for p in CREDIT_PATTERNS)

def extract_milestone_ids_from_text(text: str) -> List[str]:
    found = re.findall(r"[Mm]ilestone[s]?\s*(\d+)", text) or re.findall(r"\bM(\d+)\b", text)
    return sorted({f"M{n}" for n in found}, key=lambda x: int(x[1:]))


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
            accum[text] = {
                "budget": sf(r[COL_BUDGET]),
                "month":  sf(r[rep_col]),
                "ytd":    sf(r[COL_TOTAL_ACT]),
            }
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
                workstreams=[{
                    "name":   c,
                    "month":  accum.get(c,{}).get("month",0.0),
                    "ytd":    accum.get(c,{}).get("ytd",0.0),
                    "budget": accum.get(c,{}).get("budget",0.0),
                } for c in LINE_ITEMS],
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
    personnel:  Dict[str, Dict]  = {}
    people_map: Dict[str, int]   = {code_to_mid(c): 0 for c in mil_codes}

    for col_idx, code in enumerate(mil_codes, 1):
        mid   = code_to_mid(code)
        count = sum(1 for row in rows[header_idx+1:]
                    if not str(row[0]).lower().startswith("grand total")
                    and col_idx < len(row) and sf(row[col_idx]) > 0)
        people_map[mid] = count

    for row in rows[header_idx+1:]:
        label = str(row[0]).strip() if row[0] is not None else ""
        if not label or label.lower().startswith("grand total"): continue
        gt = sf(row[gt_col]) if gt_col < len(row) else 0.0
        if gt == 0.0: continue
        lcat_hours[label] = lcat_hours.get(label, 0.0) + gt
        is_name = " " in label and label != label.upper()
        if is_name:
            per_mil = []
            for col_idx, code in enumerate(mil_codes, 1):
                h = sf(row[col_idx]) if col_idx < len(row) else 0.0
                if h > 0:
                    per_mil.append({"id": code_to_mid(code), "hours": round(h, 1)})
            if label not in personnel:
                personnel[label] = {"total": 0.0, "milestones": per_mil}
            personnel[label]["total"] = round(personnel[label]["total"] + gt, 1)
        for col_idx, code in enumerate(mil_codes, 1):
            if col_idx < len(row):
                mil_hours[code] = mil_hours.get(code, 0.0) + sf(row[col_idx])

    total_hours  = sum(lcat_hours.values())
    by_milestone = sorted(
        [{"milestone_id": code_to_mid(c), "hours": round(h, 1),
          "people": people_map.get(code_to_mid(c), 0)}
         for c, h in mil_hours.items() if h > 0],
        key=lambda x: x["hours"], reverse=True,
    )
    by_lcat = sorted(
        [{"lcat": k, "hours": round(v, 1)} for k, v in lcat_hours.items() if v > 0],
        key=lambda x: x["hours"], reverse=True,
    )
    by_personnel = sorted(
        [{"name": k, "hours": round(v["total"], 1), "milestones": v["milestones"]}
         for k, v in personnel.items()],
        key=lambda x: x["hours"], reverse=True,
    )
    print(f"  Hours: {len(by_lcat)} LCATs, {total_hours:,.0f} hrs, "
          f"{len(by_milestone)} milestones, {len(by_personnel)} personnel rows")
    return {
        "total_hours":   round(total_hours, 1),
        "total_lcats":   len(by_lcat),
        "active_mils":   len(by_milestone),
        "by_milestone":  by_milestone,
        "by_lcat":       by_lcat,
        "by_personnel":  by_personnel,
        "people_map":    people_map,
    }

def apply_hours_to_milestones(milestones: List[MilestoneRow], hours: Dict) -> None:
    if not hours: return
    pmap = hours.get("people_map", {})
    hmap = {x["milestone_id"]: x["hours"] for x in hours.get("by_milestone", [])}
    for m in milestones:
        m.people_count = pmap.get(m.milestone_id, 0)
        m.total_hours  = hmap.get(m.milestone_id, 0.0)

def apply_labor_rates(hours_data: Dict, all_milestones: List[MilestoneRow]) -> None:
    if not hours_data: return
    labor_map: Dict[str, float] = {}
    for m in all_milestones:
        ws_labor = next((w for w in m.workstreams if w["name"] == "Labor"), None)
        if ws_labor:
            labor_map[m.milestone_id] = ws_labor.get("ytd", 0.0)
    total_hrs   = hours_data.get("total_hours", 0.0)
    total_labor = sum(labor_map.values())
    blended     = round(total_labor / total_hrs, 2) if total_hrs > 0 else 0.0
    for entry in hours_data.get("by_milestone", []):
        mid  = entry["milestone_id"]
        hrs  = entry["hours"]
        lab  = labor_map.get(mid, 0.0)
        rate = round(lab / hrs, 2) if hrs > 0 and lab > 0 else blended
        entry["implied_rate"] = rate
        entry["implied_cost"] = round(hrs * rate, 2)
    for entry in hours_data.get("by_lcat", []):
        entry["implied_cost"] = round(entry["hours"] * blended, 2)
    for entry in hours_data.get("by_personnel", []):
        entry["implied_cost"] = round(entry["hours"] * blended, 2)
    hours_data["blended_rate"] = blended
    hours_data["total_labor"]  = round(total_labor, 2)


# ─────────────────────────────────────────────
# PANEL BUILDERS
# ─────────────────────────────────────────────
def build_stories_panel(stories: Dict, all_milestones: List[MilestoneRow]) -> str:
    if not stories:
        return (
            '<section class="panel" style="margin-bottom:18px;">'
            '<h2>Story Points &amp; Velocity</h2>'
            '<p class="note">Add <em>Points by Milestones.xlsx</em> to data/ to populate.</p>'
            '</section>'
        )
    total  = stories.get("total_points", 0)
    n_mils = stories.get("total_mils", 0)
    by_mil = stories.get("by_milestone", [])
    maxp   = by_mil[0]["points"] if by_mil else 1
    tmap   = {m.milestone_id: m.raw_title for m in all_milestones}

    bars = "".join(
        f'<div style="margin-bottom:10px;">'
        f'<div style="display:flex;justify-content:space-between;font-size:13px;margin-bottom:3px;">'
        f'<span><strong>{it["milestone_id"]}</strong>'
        f'{(" — " + tmap[it["milestone_id"]]) if it["milestone_id"] in tmap else ""}</span>'
        f'<span style="font-weight:700;">{it["points"]:,.0f} pts</span></div>'
        f'<div style="height:6px;background:rgba(15,23,42,0.08);border-radius:3px;overflow:hidden;">'
        f'<div style="width:{round(it["points"]/maxp*100)}%;height:100%;'
        f'background:#7c3aed;border-radius:3px;"></div>'
        f'</div></div>'
        for it in by_mil[:15]
    )
    kpi = ("background:rgba(124,58,237,0.06);border:1px solid rgba(124,58,237,0.18);"
           "border-radius:14px;padding:14px 16px;")
    return (
        f'<section class="panel" style="margin-bottom:18px;">'
        f'<h2>Story Points &amp; Velocity</h2>'
        f'<p class="note">Point totals by milestone.</p>'
        f'<div style="display:grid;grid-template-columns:repeat(2,1fr);gap:14px;margin-bottom:18px;">'
        f'<div style="{kpi}"><div class="label">Total Points</div>'
        f'<div class="value">{total:,.0f}</div></div>'
        f'<div style="{kpi}"><div class="label">Milestones with points</div>'
        f'<div class="value">{n_mils}</div></div></div>'
        f'<h3 style="font-size:15px;margin-bottom:10px;">Points by milestone</h3>'
        f'{bars}</section>'
    )


def build_hours_panel(hours_data: Dict) -> str:
    """Sortable, collapsible Hours & LCAT panel with implied costs."""
    if not hours_data:
        return (
            '<section class="panel" style="margin-bottom:18px;">'
            '<h2>Hours &amp; LCAT Breakdown</h2>'
            '<p class="note">Add an <em>Hours by LCAT.xlsx</em> file to data/ to populate.</p>'
            '</section>'
        )

    total     = hours_data.get("total_hours", 0)
    n_lcats   = hours_data.get("total_lcats", 0)
    blended   = hours_data.get("blended_rate", 0.0)
    tot_labor = hours_data.get("total_labor", 0.0)
    by_lcat   = hours_data.get("by_lcat", [])
    by_pers   = hours_data.get("by_personnel", [])

    kpi_s = ("background:rgba(37,99,235,0.06);border:1px solid rgba(37,99,235,0.18);"
             "border-radius:14px;padding:14px 16px;")
    kpis = (
        f'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px;">'
        f'<div style="{kpi_s}"><div class="label">Total Hours</div>'
        f'<div class="value">{total:,.0f}</div></div>'
        f'<div style="{kpi_s}"><div class="label">Active LCATs</div>'
        f'<div class="value">{n_lcats}</div></div>'
        f'<div style="{kpi_s}"><div class="label">Blended Rate</div>'
        f'<div class="value">{fmt_usd(blended)}/hr</div></div>'
        f'<div style="{kpi_s}"><div class="label">Implied Labor Cost</div>'
        f'<div class="value">{fmt_usd(tot_labor) if tot_labor else "—"}</div></div>'
        f'</div>'
    )

    # Embed data as JSON attributes on hidden elements — avoids </script> injection
    pers_json = _safe_json(by_pers[:50])
    lcat_json = _safe_json(by_lcat)
    total_str = str(total)

    # Build the panel — JS reads data from hidden spans, not inline script vars
    sc = "<" + "/script>"
    panel = (
        f'<section class="panel" style="margin-bottom:18px;" id="hp_panel">'
        f'<h2>Hours &amp; LCAT Breakdown</h2>'
        f'<p class="note">Labor hours and implied costs. Blended rate = total labor YTD &divide; total hours.</p>'
        f'{kpis}'
        f'<span id="hp_pers_data" style="display:none;">{pers_json}</span>'
        f'<span id="hp_lcat_data" style="display:none;">{lcat_json}</span>'
        f'<span id="hp_total_hrs" style="display:none;">{total_str}</span>'

        # Personnel header + table
        f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">'
        f'<h3 style="font-size:15px;margin:0;">Top Personnel by Hours</h3>'
        f'<button onclick="hpTogglePers()" id="hp_persBtn" style="font-size:12px;border:1px solid #cbd5e1;'
        f'background:#fff;border-radius:6px;padding:4px 10px;cursor:pointer;">Show all</button></div>'
        f'<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:18px;">'
        f'<thead><tr style="border-bottom:2px solid #e2e8f0;text-align:left;">'
        f'<th onclick="hpSortPers(\'name\')" style="cursor:pointer;padding:6px 4px;">Name ⇅</th>'
        f'<th onclick="hpSortPers(\'hours\')" style="cursor:pointer;padding:6px 4px;text-align:right;">Hours ⇅</th>'
        f'<th onclick="hpSortPers(\'implied_cost\')" style="cursor:pointer;padding:6px 4px;text-align:right;">Est. Cost ⇅</th>'
        f'<th style="padding:6px 4px;">Milestones</th>'
        f'</tr></thead><tbody id="hp_persTbody"></tbody></table>'

        # LCAT header + table
        f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">'
        f'<h3 style="font-size:15px;margin:0;">LCAT Breakdown</h3>'
        f'<button onclick="hpToggleLcat()" id="hp_lcatBtn" style="font-size:12px;border:1px solid #cbd5e1;'
        f'background:#fff;border-radius:6px;padding:4px 10px;cursor:pointer;">Collapse</button></div>'
        f'<table style="width:100%;border-collapse:collapse;font-size:13px;" id="hp_lcatTable">'
        f'<thead><tr style="border-bottom:2px solid #e2e8f0;text-align:left;">'
        f'<th onclick="hpSortLcat(\'lcat\')" style="cursor:pointer;padding:6px 4px;">LCAT ⇅</th>'
        f'<th onclick="hpSortLcat(\'hours\')" style="cursor:pointer;padding:6px 4px;text-align:right;">Hours ⇅</th>'
        f'<th onclick="hpSortLcat(\'implied_cost\')" style="cursor:pointer;padding:6px 4px;text-align:right;">Est. Cost ⇅</th>'
        f'<th style="padding:6px 4px;text-align:right;">% of Total</th>'
        f'</tr></thead><tbody id="hp_lcatTbody"></tbody></table>'
        f'</section>'

        # Script reads from hidden spans — no </script> risk inside f-string
        f'<script>'
        f'(function(){{'
        f'var pd=JSON.parse(document.getElementById("hp_pers_data").textContent);'
        f'var ld=JSON.parse(document.getElementById("hp_lcat_data").textContent);'
        f'var th=parseFloat(document.getElementById("hp_total_hrs").textContent)||1;'
        f'var pk="hours",pa=false,lk="hours",la=false,pAll=false,lCol=false;'
        f'function rp(){{'
        f'var r=pd.slice().sort(function(a,b){{var v=pa?1:-1;return (a[pk]||0)<(b[pk]||0)?v:-v;}});'
        f'r=pAll?r:r.slice(0,5);'
        f'document.getElementById("hp_persTbody").innerHTML=r.map(function(x){{'
        f'var m=(x.milestones||[]).map(function(y){{return y.id;}}).join(", ");'
        f'var c=x.implied_cost?"$"+Math.round(x.implied_cost).toLocaleString():"\u2014";'
        f'return "<tr style=\'border-bottom:1px solid #f1f5f9;\'>"'
        f'+"<td style=\'padding:5px 4px;\'>"+x.name+"</td>"'
        f'+"<td style=\'padding:5px 4px;text-align:right;font-weight:700;\'>"+x.hours.toLocaleString()+"</td>"'
        f'+"<td style=\'padding:5px 4px;text-align:right;\'>"+c+"</td>"'
        f'+"<td style=\'padding:5px 4px;color:#51627d;font-size:12px;\'>"+m+"</td>"'
        f'+"</tr>";}}).join("")'
        f'||"<tr><td colspan=4 style=\'color:#51627d;padding:8px;\'>No personnel rows detected.</td></tr>";'
        f'document.getElementById("hp_persBtn").textContent=pAll?"Show top 5":"Show all ("+pd.length+")";'
        f'}}'
        f'function rl(){{'
        f'var r=ld.slice().sort(function(a,b){{var v=la?1:-1;return (a[lk]||0)<(b[lk]||0)?v:-v;}});'
        f'document.getElementById("hp_lcatTbody").innerHTML=r.map(function(x){{'
        f'var pct=th>0?(x.hours/th*100).toFixed(1)+"%":"\u2014";'
        f'var c=x.implied_cost?"$"+Math.round(x.implied_cost).toLocaleString():"\u2014";'
        f'var bw=th>0?Math.round(x.hours/th*100):0;'
        f'return "<tr style=\'border-bottom:1px solid #f1f5f9;\'>"'
        f'+"<td style=\'padding:5px 4px;\'>"+x.lcat+"</td>"'
        f'+"<td style=\'padding:5px 4px;text-align:right;font-weight:700;\'>"+x.hours.toLocaleString()'
        f'+"<div style=\'height:4px;background:rgba(15,23,42,0.07);border-radius:2px;margin-top:3px;\'>"'
        f'+"<div style=\'width:"+bw+"%;height:100%;background:#2563eb;border-radius:2px;\'></div></div></td>"'
        f'+"<td style=\'padding:5px 4px;text-align:right;\'>"+c+"</td>"'
        f'+"<td style=\'padding:5px 4px;text-align:right;color:#51627d;\'>"+pct+"</td>"'
        f'+"</tr>";}}).join("");'
        f'document.getElementById("hp_lcatTable").style.display=lCol?"none":"";'
        f'document.getElementById("hp_lcatBtn").textContent=lCol?"Expand":"Collapse";'
        f'}}'
        f'window.hpSortPers=function(k){{if(pk===k)pa=!pa;else{{pk=k;pa=false;}}rp();}};'
        f'window.hpTogglePers=function(){{pAll=!pAll;rp();}};'
        f'window.hpSortLcat=function(k){{if(lk===k)la=!la;else{{lk=k;la=false;}}rl();}};'
        f'window.hpToggleLcat=function(){{lCol=!lCol;rl();}};'
        f'rp();rl();'
        f'}})();'
        f'</script>'  # this is a real Python string ending, not injected HTML
    )
    return panel


# ─────────────────────────────────────────────
# NARRATIVE PARSER
# ─────────────────────────────────────────────
def parse_narrative_doc(path: Path) -> Dict[str, Any]:
    if not DOCX_OK: return {}
    try: doc = DocxDocument(path)
    except Exception as e:
        print(f"  WARNING docx {path.name}: {e}"); return {}

    by_milestone: Dict[str, str] = {}
    credits:      List[Dict]     = []
    portfolio_summary            = ""
    cur_sec, cur_lines           = None, []

    def flush():
        nonlocal portfolio_summary
        if not cur_sec: return
        body = " ".join(cur_lines).strip()
        m = re.search(r"(\d+)", cur_sec)
        if m:
            by_milestone[f"M{m.group(1)}"] = body
        elif any(k in cur_sec.lower() for k in
                 ["portfolio summary","costs incurred","financial report","executive summary"]):
            portfolio_summary = body
        for line in cur_lines:
            if is_credit_text(line):
                credits.append({"section": cur_sec, "text": line.strip()})

    for para in doc.paragraphs:
        t = para.text.strip()
        if not t: continue
        is_mil  = bool(re.search(r"^\b(Milestone|M)\s*\d+", t, re.IGNORECASE))
        is_port = (len(t) < 100 and any(k in t.lower() for k in
                   ["portfolio summary","costs incurred","financial report","executive summary"]))
        if is_mil or is_port:
            flush(); cur_sec, cur_lines = t, []
        elif cur_sec is not None:
            cur_lines.append(t)
    flush()

    print(f"  Narrative: {len(by_milestone)} milestone sections, {len(credits)} credit/adjustment items")
    return {"by_milestone": by_milestone, "credits": credits, "portfolio_summary": portfolio_summary}

def narrative_for(narratives: Dict[str, Any], title: str) -> str:
    m = re.search(r"(\d+)", title)
    if not m: return ""
    return narratives.get("by_milestone", {}).get(f"M{m.group(1)}", "")


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
            raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=15, engine="openpyxl")
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
    return {
        "total_points": round(total_points, 1),
        "total_mils":   len(by_milestone),
        "by_milestone": sorted(
            [{"milestone_id": k, "points": round(v, 1)} for k, v in by_milestone.items()],
            key=lambda x: x["points"], reverse=True,
        ),
    }


# ─────────────────────────────────────────────
# ARCHIVE
# ─────────────────────────────────────────────
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
    archive_dir.mkdir(parents=True, exist_ok=True)
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
    (archive_dir / "index.html").write_text(html, encoding="utf-8")
    manifest = [{"label": it["label"], "file": it["file"]} for it in items]
    (archive_dir / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")


# ─────────────────────────────────────────────
# GIT
# ─────────────────────────────────────────────
def git_push(repo_dir: Path, message: str) -> None:
    subprocess.run(["git", "-C", str(repo_dir), "stash"],            check=False)
    subprocess.run(["git", "-C", str(repo_dir), "fetch", "origin"],  check=True)
    subprocess.run(["git", "-C", str(repo_dir), "pull", "origin", "main", "--rebase"], check=True)
    subprocess.run(["git", "-C", str(repo_dir), "stash", "pop"],     check=False)
    for cmd in [
        ["git", "-C", str(repo_dir), "add", "index.html"],
        ["git", "-C", str(repo_dir), "add", "-f", "archive"],
        ["git", "-C", str(repo_dir), "add", "output/dashboard_data.json"],
    ]:
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0:
            print(f"  WARNING git add: {r.stderr.strip()} ({cmd[-1]})")
    subprocess.run(["git", "-C", str(repo_dir), "commit", "-m", message], check=False)
    subprocess.run(["git", "-C", str(repo_dir), "push"], check=True)

def sync_and_push(repo_dir: Path, message: str) -> None:
    print(f"🚀 Pushing to GitHub: {message}")
    try:
        git_push(repo_dir, message)
        print("✅ Pushed successfully.")
    except Exception as e:
        print(f"❌ Git push failed: {e}")


# ─────────────────────────────────────────────
# HTML GENERATION
# ─────────────────────────────────────────────
def generate_html(
    all_milestones:       List[MilestoneRow],
    monthly_chart:        List[Dict],
    reporting_month:      str,
    reporting_month_full: str,
    awards_str:           str,
    narratives:           Dict[str, Any],
    stories_data:         Dict,
    hours_data:           Dict,
    prev_link:            str = "",
    prev_label:           str = "",
    next_link:            str = "",
    next_label:           str = "",
) -> None:
    template_path = next(
        (p for p in [ROOT/"dashboard_template_fixed.html", ROOT/"dashboard_template.html"]
         if p.exists()), None
    )
    if not template_path:
        print("❌ ERROR: dashboard_template*.html not found"); return

    html = template_path.read_text(encoding="utf-8")

    print("  Fetching Chart.js…")
    chartjs     = fetch_js("https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js")
    print("  Fetching html2pdf…")
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
    budget_used_pct = round(ytd_tot / budget_tot * 100, 1) if budget_tot else 0.0
    yr_pct          = yr_elapsed_pct(reporting_month)
    pace_gap        = round(budget_used_pct - yr_pct, 1)
    balance_color   = "#dc2626" if bal_tot < 0 else "#059669"

    prev_val    = monthly_chart[rep_idx-1]["value"] or month_tot if rep_idx > 0 else month_tot
    trend_pct   = round((month_tot - prev_val) / prev_val * 100, 1) if prev_val else 0.0
    trend_arrow = "▲" if trend_pct > 0 else "▼" if trend_pct < 0 else "→"
    trend_color = "#dc2626" if trend_pct > 0 else "#059669" if trend_pct < 0 else "#51627d"

    n_flags      = len(at_risk) + len(on_watch)
    n_crit       = len(at_risk)
    flag_summary = (
        "No milestone flags need escalation." if n_flags == 0 else
        f"{n_flags} flagged, all minor — monitor closely." if n_crit == 0 else
        f"{n_flags} flagged — {n_crit} require{'s' if n_crit==1 else ''} attention."
    )
    tension_msg = (
        f"Budget used is {budget_used_pct:.1f}% versus {yr_pct:.1f}% of the award year elapsed. "
        f"{n_crit} milestone(s) At Risk and {len(on_watch)} on Watch."
    )

    watchlist = sorted(
        [m for m in all_milestones if m.status in ("At Risk","Watch") and not m.exception_note],
        key=lambda m: (m.status == "At Risk", abs(m.projected_overrun)), reverse=True,
    )[:5]
    watchlist_html = "\n".join(
        bullet_html(
            "risk" if m.status == "At Risk" else "warn",
            f"<strong>{m.title}</strong> — {fmt_usd_short(abs(m.projected_overrun))} projected overrun"
            + (f" — {short_text(narrative_for(narratives, m.title))}"
               if narrative_for(narratives, m.title) else "") + "."
        ) for m in watchlist
    ) or bullet_html("good", "No milestones currently flagged for attention.")

    callout_html = "\n".join([
        bullet_html("risk" if bal_tot < 0 else "good",
            f"<strong>Portfolio balance is {'negative' if bal_tot<0 else 'positive'} at {fmt_usd(bal_tot)}.</strong>"),
        bullet_html("warn" if n_flags else "good", f"<strong>{flag_summary}</strong>"),
    ])

    # Operational highlights — top spenders + credits from narrative
    credits = narratives.get("credits", []) if isinstance(narratives, dict) else []
    op_items: List[str] = []
    for m in sorted(all_milestones, key=lambda x: x.monthly_spend, reverse=True)[:5]:
        narr = narrative_for(narratives, m.title)
        note = f" — {short_text(narr)}" if narr else ""
        op_items.append(bullet_html("good",
            f"<strong>{m.title}</strong> — {fmt_usd(m.monthly_spend)} this month{note}."))
    for credit in credits[:5]:
        mids = extract_milestone_ids_from_text(credit["text"])
        tag  = f" [{', '.join(mids)}]" if mids else f" [{credit['section']}]"
        op_items.append(bullet_html("warn",
            f"<strong>Credit/Adjustment{tag}:</strong> {short_text(credit['text'], 120)}"))
    operational_html = "\n".join(op_items)

    # Milestone popover JS
    mil_title_map  = {m.milestone_id: m.title for m in all_milestones}
    mil_title_json = _safe_json(mil_title_map)
    mil_popover_js = (
        "<script>"
        "(function(){"
        "var tm=" + mil_title_json + ";"
        "document.querySelectorAll('.mil-count-link').forEach(function(el){"
        "el.addEventListener('click',function(e){"
        "e.stopPropagation();"
        "var ex=document.getElementById('mil-popover');"
        "if(ex){ex.remove();return;}"
        "var mids=(el.dataset.mids||'').split(',').map(function(x){return x.trim();}).filter(Boolean);"
        "var lines=mids.map(function(id){"
        "return \"<div style='padding:4px 0;border-bottom:1px solid #f1f5f9;'>\""
        "+\"<strong>\"+id+\"</strong>\"+(tm[id]?\" \u2014 \"+tm[id]:\"\")+\"</div>\";"
        "}).join('');"
        "var pop=document.createElement('div');"
        "pop.id='mil-popover';"
        "pop.style.cssText='position:fixed;z-index:9999;background:#fff;border:1px solid #e2e8f0;"
        "border-radius:12px;box-shadow:0 8px 32px rgba(15,23,42,0.15);"
        "padding:14px 16px;max-width:420px;font-size:13px;max-height:60vh;overflow-y:auto;';"
        "pop.innerHTML=\"<div style='font-weight:700;margin-bottom:8px;'>Milestones (\"+mids.length+\")</div>\"+lines;"
        "var r=el.getBoundingClientRect();"
        "pop.style.top=(r.bottom+6)+'px';"
        "pop.style.left=Math.min(r.left,window.innerWidth-440)+'px';"
        "document.body.appendChild(pop);"
        "document.addEventListener('click',function rm(){pop.remove();document.removeEventListener('click',rm);});"
        "});"
        "});"
        "})();"
        "<" + "/script>"
    )

    sc_tag     = "<" + "/script>"
    pdf_js_tag = "<script>" + html2pdf_js + sc_tag
    slug        = month_slug(reporting_month_full)

    milestone_span = (
        f"{all_milestones[0].milestone_id}–{all_milestones[-1].milestone_id}"
        if all_milestones else "N/A"
    )
    monthly_json = _safe_json(monthly_chart)

    replacements = {
        "__MONTH__":            reporting_month_full,
        "__BALANCE_COLOR__":    balance_color,
        "__TITLE__":            f"PPSC Financial Dashboard — {reporting_month_full}",
        "__SUBTITLE__":         f"Award(s): {awards_str} · {reporting_month_full}",
        "__AWARD__":            awards_str,
        "__MILESTONE_SPAN__":   milestone_span,
        "__YTD_SPEND_TOTAL__":  fmt_usd(ytd_tot),
        "__AWARD_BUDGET__":     fmt_usd(budget_tot),
        "__MONTHLY_SPEND__":    fmt_usd(month_tot),
        "__WATCHLIST_HTML__":   watchlist_html,
        "__OPERATIONAL_HTML__": operational_html,
        "__CONSOLIDATED_CALLOUTS_HTML__": callout_html,
        "__FOOTER__":           f"Generated {datetime.now().strftime('%B %d, %Y %H:%M')}",
        "__STORIES_PANEL__":    build_stories_panel(stories_data, all_milestones),
        "__HOURS_PANEL__":      build_hours_panel(hours_data),
        "__TREND_COLOR__":      trend_color,
        "__SITE_BASE_URL__":    SITE_BASE_URL,
        "__ARCHIVE_BASE_URL__": ARCHIVE_BASE_URL,
        "__MONTHS_JSON_PATH__": f"{ARCHIVE_BASE_URL}manifest.json",
        "__DECISION_1__":       (f"{n_crit} milestone(s) need immediate review."
                                 if n_crit else "No immediate decisions required."),
        "__DECISION_2__":       f"Budget used {budget_used_pct:.1f}% vs {yr_pct:.1f}% of award year.",
        "__DECISION_3__":       "Review LCAT allocation for high-burn milestones.",
        "__BUDGET_USED_PCT__":  f"{budget_used_pct:.1f}",
        "__YEAR_ELAPSED_PCT__": f"{yr_pct:.1f}",
        "__BUDGET_PACE_GAP__":  f"{pace_gap:+.1f}",
        "__TREND_ARROW__":      trend_arrow,
        "__TREND_PCT__":        f"{abs(trend_pct):.1f}",
        "__TREND_TEXT__":       f"{trend_arrow} {abs(trend_pct):.1f}% vs prior month.",
        "__FLAGGED_COUNT__":    str(n_flags),
        "__FLAGGED_SUMMARY__":  flag_summary,
        "__PORTFOLIO_BALANCE__":fmt_usd(bal_tot),
        "__AT_A_GLANCE__":      tension_msg,
        "__AT_RISK_MIDS__":     ",".join(m.milestone_id for m in at_risk),
        "__WATCH_MIDS__":       ",".join(m.milestone_id for m in on_watch),
        "__ON_TRACK_MIDS__":    ",".join(m.milestone_id for m in on_track),
        "__COUNT_AT_RISK__":    str(len(at_risk)),
        "__COUNT_WATCH__":      str(len(on_watch)),
        "__COUNT_ON_TRACK__":   str(len(on_track)),
        "__HTML2PDF_JS__":      pdf_js_tag,
        "__CURRENT_FILE__":     f"{slug}.html",
        "__MIL_POPOVER_JS__":   mil_popover_js,
        "__MILESTONES_JSON__":  _safe_json([asdict(m) for m in all_milestones]),
        "__MONTHLY_DATA_JSON__":monthly_json,
    }

    for k, v in replacements.items():
        html = html.replace(k, str(v))

    out = ROOT / "index.html"
    out.write_text(html, encoding="utf-8")
    print(f"✅ index.html written ({out.stat().st_size:,} bytes)")


# ─────────────────────────────────────────────
# MAIN PROCESS
# ─────────────────────────────────────────────
def process() -> None:
    print(">>> PPSC Dashboard — starting")

    all_xlsx = sorted(
        [p for p in DATA_DIR.glob("*.xlsx") if is_valid_xlsx(p)],
        key=lambda p: p.stat().st_mtime,
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
    all_milestones:  List[MilestoneRow] = []
    awards_found:    List[str]          = []
    monthly_actuals: Dict[str, float]   = {m: 0.0 for m in MONTH_ORDER}
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

    # Forecast
    prior_forecast: Dict[str, float] = dict(FORECAST_DEFAULTS)
    if len(main_files) >= 2:
        print(f"--- Prior workbook forecast: {main_files[-2].name}")
        try:
            prior_df, _ = read_main_workbook(main_files[-2])
            for mo, val in extract_monthly_row_totals(prior_df).items():
                if val > 0:
                    prior_forecast[mo] = val
        except Exception as e:
            print(f"  WARNING: {e} — using hardcoded defaults.")
    else:
        print("  Single workbook — using hardcoded forecast defaults.")

    # Hours
    hours_data: Dict = {}
    hours_files = sorted(
        [p for p in aux_files if "hours" in p.name.lower() or "lcat" in p.name.lower()],
        key=lambda p: p.stat().st_mtime,
    )
    if hours_files:
        print(f"--- Hours: {hours_files[-1].name}")
        hours_data = parse_hours_file(hours_files[-1])
        apply_hours_to_milestones(all_milestones, hours_data)

    # Stories
    stories_data: Dict = {}
    stories_files = sorted(
        [p for p in aux_files if any(k in p.name.lower() for k in ("point","story","stories"))],
        key=lambda p: p.stat().st_mtime,
    )
    if stories_files:
        print(f"--- Stories: {stories_files[-1].name}")
        stories_data = parse_stories_file(stories_files[-1])

    # Narrative
    narratives: Dict[str, Any] = {}
    docx_files = [p for p in sorted(DATA_DIR.glob("*.docx"),
                                    key=lambda p: p.stat().st_mtime, reverse=True)
                  if not p.name.startswith("~$")]
    if docx_files:
        print(f"--- Narrative: {docx_files[0].name}")
        narratives = parse_narrative_doc(docx_files[0])

    # Implied labor rates
    apply_labor_rates(hours_data, all_milestones)

    # Chart data
    rep_idx    = MONTH_ORDER.index(reporting_month)
    budget_tot = sum(m.budget for m in all_milestones)
    monthly_chart = []
    for i, mo in enumerate(MONTH_ORDER):
        actual   = monthly_actuals[mo]
        forecast = prior_forecast.get(mo, 0.0)
        monthly_chart.append({
            "month":            mo,
            "value":            round(actual, 2)   if actual   > 0 else None,
            "forecast_val":     round(forecast, 2) if forecast > 0 else None,
            "is_forecast":      i > rep_idx,
            "is_current_month": i == rep_idx,
            "value_pct":        round(actual / budget_tot * 100, 1)
                                if budget_tot and actual > 0 else 0.0,
        })

    reporting_month_full = format_month_year(reporting_month, current_file)
    awards_str           = ", ".join(awards_found) if awards_found else "N/A"

    at_r = sum(1 for m in all_milestones if m.status == "At Risk")
    wch  = sum(1 for m in all_milestones if m.status == "Watch")
    ont  = sum(1 for m in all_milestones if m.status in ("On Track","Complete"))
    print(f"\nMilestones: {len(all_milestones)} | Awards: {awards_str} | Month: {reporting_month_full}")
    print(f"Status: {ont} On Track · {wch} Watch · {at_r} At Risk")

    # JSON output
    (ROOT / "output").mkdir(exist_ok=True)
    (ROOT / "output" / "dashboard_data.json").write_text(
        json.dumps({
            "summary": {
                "reporting_month": reporting_month_full, "awards": awards_str,
                "budget_total":    sum(m.budget        for m in all_milestones),
                "ytd_total":       sum(m.ytd_actual    for m in all_milestones),
                "month_total":     sum(m.monthly_spend for m in all_milestones),
                "bal_total":       sum(m.est_balance   for m in all_milestones),
            },
            "milestones":  [asdict(m) for m in all_milestones],
            "monthlyData": monthly_chart,
        }, indent=2), encoding="utf-8",
    )

    # ── Step 1: first HTML pass (no nav — archive not written yet) ────────
    generate_html(
        all_milestones, monthly_chart, reporting_month,
        reporting_month_full, awards_str, narratives,
        stories_data, hours_data,
    )

    # ── Step 2: archive snapshot + manifest ───────────────────────────────
    slug          = month_slug(reporting_month_full)
    snapshot_path = None
    if ARCHIVE_DIR is not None:
        ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
        snapshot_path = ARCHIVE_DIR / f"{slug}.html"
        snapshot_path.write_text(
            (ROOT / "index.html").read_text(encoding="utf-8"), encoding="utf-8"
        )
        build_archive_index(ARCHIVE_DIR, reporting_month_full)
        print(f"  Archive snapshot → {snapshot_path.name}")

    # ── Step 3: resolve prev/next from updated manifest ───────────────────
    prev_link = prev_label = next_link = next_label = ""
    if ARCHIVE_DIR is not None and ARCHIVE_DIR.exists():
        arc_items  = archive_manifest(ARCHIVE_DIR)
        arc_labels = [it["label"] for it in arc_items]
        try:
            idx = arc_labels.index(reporting_month_full)
        except ValueError:
            idx = -1
        if idx > 0:
            prev_link  = f"{ARCHIVE_BASE_URL}{arc_items[idx-1]['file']}"
            prev_label = arc_items[idx-1]["label"]
        if 0 <= idx < len(arc_items) - 1:
            next_link  = f"{ARCHIVE_BASE_URL}{arc_items[idx+1]['file']}"
            next_label = arc_items[idx+1]["label"]

    # ── Step 4: re-generate with nav baked in ────────────────────────────
    generate_html(
        all_milestones, monthly_chart, reporting_month,
        reporting_month_full, awards_str, narratives,
        stories_data, hours_data,
        prev_link=prev_link, prev_label=prev_label,
        next_link=next_link, next_label=next_label,
    )

    # ── Step 5: overwrite archive with nav version ────────────────────────
    if snapshot_path is not None:
        snapshot_path.write_text(
            (ROOT / "index.html").read_text(encoding="utf-8"), encoding="utf-8"
        )

    # ── Step 6: push ──────────────────────────────────────────────────────
    if AUTO_PUSH_GITHUB:
        sync_and_push(GIT_REPO_DIR, f"Dashboard update: {reporting_month_full}")


if __name__ == "__main__":
    process()