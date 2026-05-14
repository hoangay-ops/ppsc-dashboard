from pathlib import Path
import pandas as pd
import re

DATA_DIR = Path("data")

def norm(v):
    return str(v).replace("\xa0"," ").replace("\u2013","-").replace("\u2014","-").strip()

def choose_sheet(xls):
    return xls.sheet_names[0] if len(xls.sheet_names) == 1 else xls.sheet_names[1]

all_xlsx = sorted(
    [p for p in DATA_DIR.glob("*.xlsx")
     if not p.name.startswith("~$") and p.stat().st_size > 5000],
    key=lambda p: p.stat().st_mtime
)
main_files = [p for p in all_xlsx
              if not re.search(r"hours|lcat|story|stories|points", p.name, re.IGNORECASE)]
if not main_files:
    print("ERROR: no main workbook found in data/"); raise SystemExit

path = main_files[-1]
print(f"Workbook: {path.name}\n")

xls   = pd.ExcelFile(path, engine="openpyxl")
sheet = choose_sheet(xls)
df    = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")

print(f"Sheet: '{sheet}'   Shape: {df.shape}\n")

print("=== Row 0 — column index : value ===")
for ci, val in enumerate(df.iloc[0]):
    print(f"  col {ci:>2}: {repr(val)}")

print()

in_block = False
rows_printed = 0
print("=== First Milestone block ===")
for ridx, row in df.iterrows():
    text = norm(row[1]) if len(row) > 1 else ""
    if re.match(r"Milestone\s+\d+", text, re.IGNORECASE):
        in_block = True
    if in_block:
        print(f"\n  row {ridx}:")
        for ci, val in enumerate(row):
            print(f"    col {ci:>2}: {repr(val)}")
        rows_printed += 1
        if rows_printed >= 8:
            break
