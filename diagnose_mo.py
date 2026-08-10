#!/usr/bin/env python3
"""
Diagnostic script — investigates why the MO file (202601_202606_MO.xlsx)
is dropping ~17% of rows and producing zero rows for June 2026.

Run this locally (same folder as update_dashboard.py works from):
    python3 diagnose_mo.py

It does NOT modify or push anything — read-only inspection.
"""
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("Run: pip3 install pandas openpyxl --break-system-packages")
    sys.exit(1)

ONEDRIVE_FOLDER = Path.home() / "Library/CloudStorage/OneDrive-PT.TelekomunikasiIndonesia/Performancedb"
FILE = ONEDRIVE_FOLDER / "202601_202606_MO.xlsx"

if not FILE.exists():
    print(f"❌  File not found: {FILE}")
    sys.exit(1)

print(f"📂  Loading {FILE.name} ({FILE.stat().st_size / 1024 / 1024:.1f} MB)...")
df = pd.read_excel(FILE, engine='openpyxl')
df.columns = [c.strip().lower() for c in df.columns]
print(f"✓  Loaded {len(df):,} rows, {len(df.columns)} columns\n")

print("Columns:", list(df.columns))
print()

col = 'c_statusdate'
if col not in df.columns:
    print(f"❌  Column '{col}' not found! Available columns with 'status' or 'date' in the name:")
    for c in df.columns:
        if 'status' in c or 'date' in c:
            print("   -", c)
    sys.exit(1)

raw = df[col]
print(f"── Raw '{col}' column ──")
print("dtype:", raw.dtype)
print("null/NaN (raw, before parsing):", int(raw.isna().sum()), f"({raw.isna().sum()/len(df)*100:.1f}%)")
print("sample of 10 raw values:")
print(raw.dropna().head(10).to_string())
print()

parsed = pd.to_datetime(raw, errors='coerce')
n_bad = int(parsed.isna().sum())
n_raw_notnull_but_bad = int((raw.notna() & parsed.isna()).sum())
print(f"── After pd.to_datetime(errors='coerce') ──")
print(f"Total unparseable (NaT): {n_bad:,} ({n_bad/len(df)*100:.1f}%)")
print(f"  - of which, raw was NOT null (i.e. had a value but couldn't parse): {n_raw_notnull_but_bad:,}")
print(f"  - of which, raw WAS null/empty to begin with: {n_bad - n_raw_notnull_but_bad:,}")
print()

if n_raw_notnull_but_bad > 0:
    print("Sample of raw values that failed to parse as dates (up to 20):")
    bad_mask = raw.notna() & parsed.isna()
    print(df.loc[bad_mask, col].head(20).to_string())
    print()
    print("Value type breakdown of the failed-to-parse values:")
    print(df.loc[bad_mask, col].apply(lambda v: type(v).__name__).value_counts())
    print()

print(f"── Valid parsed dates ──")
valid = parsed.dropna()
if len(valid):
    print("Min date:", valid.min())
    print("Max date:", valid.max())
    print()
    print("Row count by year-month:")
    print(valid.dt.strftime('%Y-%m').value_counts().sort_index())
else:
    print("No valid dates at all!")

print()
print(f"── June 2026 specifically ──")
june_mask = (parsed >= '2026-06-01') & (parsed < '2026-07-01')
print(f"Rows with c_statusdate in June 2026: {int(june_mask.sum()):,}")
if june_mask.sum() == 0:
    print("→ Confirms: genuinely ZERO rows have a June 2026 status date in this file.")
    # Check if there's any hint of June data via other date-like columns
    other_date_cols = [c for c in df.columns if 'date' in c and c != col]
    print(f"\nOther date-like columns to check for June 2026 signal: {other_date_cols}")
    for c in other_date_cols:
        try:
            p = pd.to_datetime(df[c], errors='coerce')
            jm = ((p >= '2026-06-01') & (p < '2026-07-01')).sum()
            if jm > 0:
                print(f"   ⚠  '{c}' HAS {jm:,} rows falling in June 2026 — maybe this is the column that should anchor the period instead of c_statusdate for rows still in progress?")
        except Exception:
            pass

print()
print("Done. Paste this entire output back so it can be analyzed.")
