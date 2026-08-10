#!/usr/bin/env python3
"""
Dashboard Update Script v3
--------------------------
Entities: TTI, PSRE, FFG, TTRFFG, FFGGAUL

File naming (place in OneDrive folder):
  202506_TTI.csv          → TTI
  202506_PSRE.csv         → PS/RE
  202506_FFGGROUP.xlsx    → FFG + TTRFFG + FFGGAUL  (multi-sheet: FFG / TTRFFG / GAUL)
  OR separate files:
  202506_FFG.xlsx         → FFG only
  202506_TTRFFG.xlsx      → TTRFFG only
  202506_FFGGAUL.xlsx     → FFGGAUL only

FFGGROUP multi-sheet format (recommended):
  Sheet "FFG"    — columns: org_1, org_2, org_3, sto, order_id, flag_exclude_all, f_ffg
  Sheet "TTRFFG" — columns: org_1, org_2, org_3, sto, order_id, flag_exclude_all, manja, compliance
  Sheet "GAUL"   — columns: tiket (FFG/GAUL), org_1, org_2, org_3, sto, order_id

FFG denominator logic:
  FFG for period YYYYMM uses TTI (comply + not_comply) from period YYYYMM-2.
  Anomaly rows are EXCLUDED from the denominator per business rule.

KPI 2025 vs KPI 2026 (dual formula computation):
  TTI and FFG each have TWO independent formulas — the original/legacy one
  (KPI 2025 menu) and a new one (KPI 2026 menu). For any period >=
  KPI_2026_CUTOVER_PERIOD ("202601"), the SAME source file (e.g.
  202601_TTI.csv) is run through BOTH formulas, producing two separate
  results that are pushed to two separate files: data/kpis_2025.json
  (legacy formula) and data/kpis_2026.json (2026 formula). Periods before
  the cutover only ever get the legacy formula. MO/PDA only ever have a
  2026 formula (kpis_2026.json only); PSRE/TTRFFG/FFGGAUL only ever have
  the legacy formula (kpis_2025.json only). See entity_formula_versions().

Usage:
    python3 update_dashboard.py
"""

import os, re, sys, json, base64
from datetime import datetime
from pathlib import Path
from glob import glob

# ─── Dependency check ─────────────────────────────────────────────────────────
def check_deps():
    missing = []
    try: import pandas
    except ImportError: missing.append("pandas")
    try: import openpyxl
    except ImportError: missing.append("openpyxl")
    try: import requests
    except ImportError: missing.append("requests")
    try: import numpy
    except ImportError: missing.append("numpy")
    if missing:
        print(f"❌  Missing packages. Run:\n    pip3 install {' '.join(missing)} --break-system-packages\n")
        sys.exit(1)

check_deps()
import pandas as pd
import numpy as np
import requests

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════
ONEDRIVE_FOLDER = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-PT.TelekomunikasiIndonesia/Performancedb"
)

from dashboard_env import (get_github_config, get_target_period, github_request,
                           GitHubUnreachable, explain_network_error, preflight)

GITHUB_TOKEN, GITHUB_REPO, GITHUB_BRANCH = get_github_config()

# "" = process ALL periods in folder; or e.g. "202606".
# Set via the TARGET_PERIOD env var — deploy.py sets it per period, so there is
# no longer any need to sed-edit this line.
TARGET_PERIOD = get_target_period()

# ── CSV upload limit ──────────────────────────────────────────────────────────
# Only push raw/processed CSVs for the N most recent periods.
# Older periods already have their CSVs on GitHub — no need to re-upload them
# every run. Set to 0 to skip CSV uploads entirely (UI-only run), or a large
# number (e.g. 99) to always upload everything.
CSV_UPLOAD_MAX_PERIODS = 1   # ← change to 2 if you want current + previous month

# Deduplication column for TTI / PSRE
DEDUP_COLUMN = "order_id"

# For FFG/TTRFFG/FFGGAUL use trouble_no as the unique key (one ticket = one row)
DEDUP_COLUMN_FFG = "trouble_no"

# SLA targets (%) — 2025 and earlier
TARGETS = {
    "TTI":      93.31,
    "FFG":      98.29,   # from piv ffg (confirm with team)
    "TTRFFG":   80.81,   # from piv ttr
    "PSRE":     75.17,
    "FFGGAUL":  91.76,   # from piv gaul
}
DEFAULT_TARGET = 95.0

# SLA targets (%) — used whenever a result is computed with the '2026'
# formula version (see entity_formula_versions() / resolve_target()).
# "TTI" here = Fulfill AO (same entity key as legacy TTI, new 2026 formula).
TARGETS_2026 = {
    "TTI": 95.50,   # Fulfill AO
    "MO":  95.00,   # Fulfill MO
    "PDA": 100.00,  # Fulfill PDA
    "FFG": 98.50,   # FFG (2026 WILSUS rule)
}


def resolve_target(entity, formula_version="legacy"):
    """
    SLA target (%) for an entity, keyed by WHICH FORMULA produced the result
    (not by calendar period — see entity_formula_versions()). The '2026'
    version uses TARGETS_2026 when the entity has a 2026-specific target
    (Fulfill AO/MO/PDA, FFG); the 'legacy' version always uses the original
    TARGETS/DEFAULT_TARGET lookup, unchanged, regardless of which calendar
    period it's being computed for.
    """
    if formula_version == "2026" and entity in TARGETS_2026:
        return TARGETS_2026[entity]
    return TARGETS.get(entity, DEFAULT_TARGET)

# Entities to silently skip during file discovery — files that happen to
# match the YYYYMM_ENTITY naming convention but aren't part of any KPI on
# this dashboard (e.g. a one-off export someone dropped in the OneDrive
# folder). Add an entity name here to have its file(s) ignored entirely.
IGNORED_ENTITIES = {"HOMEID"}

# Entities whose source files are known/expected to have no reliable unique
# row identifier (no 'order_id' column) — dedup is intentionally skipped for
# these, so we stay quiet about it instead of printing a "column not found"
# warning that reads like something's wrong.
NO_DEDUP_ENTITIES = {"MO", "PDA"}

# FFG denominator: TTI from N months prior (anomaly rows excluded)
FFG_TTI_LAG_MONTHS = 2

# Hardcoded TTI denominators for months where TTI raw file is unavailable.
# Value = comply + not_comply (anomaly already excluded).
# Used as fallback when the TTI file for ref_period is missing.
TTI_DENOM_HARDCODE = {
    "202511": {"comply": 143_591, "not_comply": 0, "anomaly": 0, "total": 143_591, "sla_pct": 0, "by_area": []},  # Nov 2025 → FFG Jan 2026
    "202512": {"comply": 150_622, "not_comply": 0, "anomaly": 0, "total": 150_622, "sla_pct": 0, "by_area": []},  # Dec 2025 → FFG Feb 2026
}

# Entity names that are treated as FFGGROUP (multi-sheet combined file)
FFGGROUP_ALIASES = {"FFGGROUP", "FFGTTRFFGFFGGAUL", "FFGTTR", "FFGALL"}

# ── Ranged (multi-month-in-one-file) entities ─────────────────────────────────
# Some entities (e.g. MO) ship one workbook covering several months at once,
# named START_END_ENTITY.xlsx (e.g. 202601_202606_MO.xlsx) instead of the
# usual one-file-per-month convention. The period for each row is derived
# from a per-row status-date column, not the filename.
# Add an entry here for any other entity that adopts this convention.
RANGED_STATUS_DATE_COL = {
    "MO":  "c_statusdate",
    "PDA": "orca_provcomp_date",
}
# ══════════════════════════════════════════════════════════════════════════════


# ─── File loading ─────────────────────────────────────────────────────────────

def detect_delimiter(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        sample = f.read(4096)
    for d in ['|', '\t', ';', ',']:
        if sample.count(d) > 10:
            return d
    return ','


def load_file(filepath, sheet=0):
    """Load CSV or Excel. sheet=0 means first sheet; sheet='Name' for named sheet."""
    ext = Path(filepath).suffix.lower()
    if ext == '.csv':
        return pd.read_csv(filepath, sep=detect_delimiter(filepath),
                           low_memory=False, encoding='utf-8', encoding_errors='replace')
    return pd.read_excel(filepath, engine='openpyxl', sheet_name=sheet)


def split_ranged_file(filepath, entity, start_period, end_period):
    """
    Load a multi-month workbook (e.g. 202601_202606_MO.xlsx) and split it into
    {period: dataframe} using that entity's status-date column (see
    RANGED_STATUS_DATE_COL). Only periods within [start_period, end_period]
    (inclusive) are kept; rows with an unparseable/missing status date are
    dropped (they can't be assigned to any period) and counted for visibility.
    """
    status_col = RANGED_STATUS_DATE_COL.get(entity)
    if not status_col:
        print(f"   ⚠  No status-date column configured for ranged entity '{entity}' — skipping")
        return {}

    df = load_file(filepath)
    df.columns = [c.strip().lower() for c in df.columns]
    if status_col not in df.columns:
        print(f"   ⚠  Ranged file for {entity}: column '{status_col}' not found — skipping")
        return {}

    df['_period_bucket'] = pd.to_datetime(df[status_col], errors='coerce').dt.strftime('%Y%m')
    n_bad = int(df['_period_bucket'].isna().sum())
    if n_bad:
        print(f"   ⚠  {n_bad:,} row(s) with unparseable/missing '{status_col}' — excluded from all periods")

    out = {}
    for period, sub in df.dropna(subset=['_period_bucket']).groupby('_period_bucket'):
        if start_period <= period <= end_period:
            out[period] = sub.drop(columns=['_period_bucket']).copy()
    return out


def _expand_ffggroup(filepath):
    """
    Multi-sheet FFGGROUP xlsx → dict: {entity: (filepath, sheet_name)}
    Looks for sheets named FFG, TTRFFG, GAUL (case-insensitive).
    Falls back to Sheet1/first sheet as TTRFFG if no named sheets found.
    """
    import openpyxl
    wb = openpyxl.load_workbook(filepath, read_only=True)
    sheets = wb.sheetnames
    wb.close()

    # Check for named sheets (FFG / TTRFFG / GAUL)
    # Accept 'TTR' as alias for TTRFFG, 'TTR FFG' also accepted
    sheet_map = [
        (['FFG'],                    'FFG'),
        (['TTRFFG', 'TTR', 'TTR FFG'], 'TTRFFG'),
        (['GAUL', 'FFG GAUL'],       'FFGGAUL'),
    ]
    named = {}
    for sheet_keys, entity in sheet_map:
        matched = next((s for s in sheets if s.strip().upper() in sheet_keys), None)
        if matched:
            named[entity] = (filepath, matched)

    if named:
        # Multi-sheet format — each KPI from its own sheet.
        # If FFG or FFGGAUL sheets are absent, fall back to the first sheet
        # so those KPIs are still computed (the single-sheet formulas work on any sheet).
        first_sheet = sheets[0]
        if 'FFG' not in named:
            print(f"   ℹ  FFGGROUP: no FFG sheet — computing FFG from '{first_sheet}'")
            named['FFG'] = (filepath, first_sheet)
        if 'FFGGAUL' not in named:
            print(f"   ℹ  FFGGROUP: no GAUL sheet — computing FFGGAUL from '{first_sheet}'")
            named['FFGGAUL'] = (filepath, first_sheet)
        return named

    # Single-sheet format — all 3 KPIs derived from the same sheet
    # FFG:     distinct order_id = NC count; denominator from TTI 2 months prior
    # TTRFFG:  flag=0 AND first_cust_assign not null → date_close - first_cust_assign ≤ 3h
    # FFGGAUL: FFG = distinct order_id; Total = all rows; SLA = FFG/Total
    first_sheet = sheets[0]
    print(f"   ℹ  FFGGROUP: single-sheet format ('{first_sheet}') — applying FFG / TTRFFG / FFGGAUL formulas")
    return {
        'FFG':     (filepath, first_sheet),
        'TTRFFG':  (filepath, first_sheet),
        'FFGGAUL': (filepath, first_sheet),
    }


# ─── Hierarchy normalisation ───────────────────────────────────────────────────

def _normalize_hierarchy(df):
    """
    FFG/TTRFFG/FFGGAUL use org_1/org_2/org_3 for hierarchy.
    Rename them to tsel_area/tsel_regional/tsel_branch so aggregate_entity works.
    """
    rename = {}
    if 'org_1' in df.columns and 'tsel_area'     not in df.columns: rename['org_1'] = 'tsel_area'
    if 'org_2' in df.columns and 'tsel_regional' not in df.columns: rename['org_2'] = 'tsel_regional'
    if 'org_3' in df.columns and 'tsel_branch'   not in df.columns: rename['org_3'] = 'tsel_branch'
    return df.rename(columns=rename) if rename else df


# ─── STO → Area/Regional/Branch lookup (Fulfill MO only) ─────────────────────
# Fulfill MO's source workbook does NOT ship org_1/org_2/org_3 (unlike TTI/
# FFG/PDA) — instead it has a handful of overlapping/ambiguous hierarchy
# columns (REGIONAL_LAMA/WITEL_LAMA/DATEL_LAMA, REGIONAL_BARU/WITEL_BARU/
# DATEL_BARU, TERRITORY_TIF/REGIONAL_TIF/DISTRICT_TIF, ...) with no single
# obviously-correct one. Per user confirmation, MO DOES have a reliable STO
# column though — so Area/Regional/Branch for MO is looked up by STO code
# using the exact same STO→hierarchy relationship TTI's own rows already
# establish (TTI carries both an STO and a trusted tsel_area/regional/branch
# per row). This file was built once from ~867k TTI rows across 202601-
# 202606 (majority-vote per STO to resolve a small number of naming-variant
# duplicates, e.g. "BALI NUSRA" vs "BALINUSRA") — see the accompanying
# sto_hierarchy_map.json shipped alongside this script. Regenerate that file
# (same TTI-derived majority-vote approach) if new STO codes start appearing
# unmapped in Fulfill MO's Data Checklist / area breakdown.
STO_HIERARCHY_MAP_FILE = "sto_hierarchy_map.json"
_STO_HIERARCHY_MAP_CACHE = None

def _load_sto_hierarchy_map():
    global _STO_HIERARCHY_MAP_CACHE
    if _STO_HIERARCHY_MAP_CACHE is not None:
        return _STO_HIERARCHY_MAP_CACHE
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), STO_HIERARCHY_MAP_FILE)
    try:
        with open(path, 'r', encoding='utf-8') as f:
            _STO_HIERARCHY_MAP_CACHE = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError) as ex:
        print(f"   ⚠  Could not load {STO_HIERARCHY_MAP_FILE} ({ex}) — Fulfill MO area breakdown will be empty")
        _STO_HIERARCHY_MAP_CACHE = {}
    return _STO_HIERARCHY_MAP_CACHE


def _apply_sto_hierarchy(df):
    """
    Derive tsel_area/tsel_regional/tsel_branch for Fulfill MO from its 'sto'
    column via the STO→hierarchy lookup (see module comment above). Rows
    whose STO code isn't in the map get NaN for these three columns — they
    still count toward MO's overall total/comply/not_comply (aggregate_entity
    computes those from the whole dataframe, not per-area), they just won't
    appear under any area in the by_area/breakdown_by_area tables, exactly
    like any other missing-hierarchy row would.
    """
    if 'sto' not in df.columns or 'tsel_area' in df.columns:
        return df
    sto_map = _load_sto_hierarchy_map()
    if not sto_map:
        return df
    sto_key = df['sto'].astype(str).str.strip().str.upper()
    df = df.copy()
    df['tsel_area']     = sto_key.map(lambda s: sto_map.get(s, {}).get('tsel_area'))
    df['tsel_regional'] = sto_key.map(lambda s: sto_map.get(s, {}).get('tsel_regional'))
    df['tsel_branch']   = sto_key.map(lambda s: sto_map.get(s, {}).get('tsel_branch'))
    n_unmapped = int(df['tsel_area'].isna().sum())
    if n_unmapped:
        print(f"   ⚠  {n_unmapped:,} Fulfill MO row(s) with an STO code not in {STO_HIERARCHY_MAP_FILE} — excluded from area breakdown (still counted in totals)")
    return df


# ─── SLA formulas ─────────────────────────────────────────────────────────────

def _fulfill_ao_formula(df):
    """
    Fulfill AO SLA — KPI 2026 menu, effective for periods >= 202601.
    New business rule (replaces the legacy 3x24h TTI rule for 2026+ data only;
    periods before 202601 keep using _tti_formula unchanged):

      - Comply threshold is 1x24 hours (24h), not 72h.
      - Duration = tgl_pc (complete) minus a reference date, tried in order:
          1. tgl_manja_awal (Manja)
          2. tgl_pi           (only if Manja is NULL)
          3. tgl_reg          (only if both Manja and PI are NULL)
        Unlike the legacy formula, a NEGATIVE duration does NOT fall through
        to the next reference — it is clamped to 0 and flagged Comply
        immediately (completed at/before the reference time = trivially OK).
      - tgl_pc empty → anomaly (excluded from SLA, same as legacy).
    """
    for col in ['tgl_pc', 'tgl_manja_awal', 'tgl_pi', 'tgl_reg']:
        df[col] = pd.to_datetime(df.get(col, pd.NaT), errors='coerce') \
                  if col in df.columns else pd.NaT

    def hrs(a, b): return (a - b).dt.total_seconds() / 3600
    dm = hrs(df['tgl_pc'], df['tgl_manja_awal'])
    dp = hrs(df['tgl_pc'], df['tgl_pi'])
    dr = hrs(df['tgl_pc'], df['tgl_reg'])

    has_m = df['tgl_manja_awal'].notna()
    has_p = df['tgl_pi'].notna()
    has_r = df['tgl_reg'].notna()
    empty = df['tgl_pc'].isna()

    use_m = has_m
    use_p = ~has_m & has_p
    use_r = ~has_m & ~has_p & has_r

    ref     = pd.Series(np.select([use_m, use_p, use_r], ['Manja', 'PI', 'RE'], default='Anomaly'), index=df.index)
    raw_dur = pd.Series(np.select([use_m, use_p, use_r], [dm, dp, dr],           default=np.nan),    index=df.index)
    # Negative duration (completed at/before the reference) → clamp to 0 (instant Comply)
    dur = raw_dur.clip(lower=0)

    has = ref.isin(['Manja', 'PI', 'RE'])
    df['_comply'] = pd.Series(
        np.select([empty, has & (dur <= 24), has & (dur > 24)],
                  ['anomaly pc empty', 'Comply', 'Not Comply'], default='anomaly all minus'),
        index=df.index)

    df['_category'] = pd.Series(np.select(
        [use_m, use_p, use_r], ['manja', 'pi', 're'], default='anomaly'
    ), index=df.index)

    # ── Detail columns for per-row export ────────────────────────────────────
    df['_ref_start']      = ref
    df['_duration_h']     = dur.round(2)       # clamped duration used for SLA
    df['_raw_duration_h'] = raw_dur.round(2)   # unclamped, for auditing negatives
    _reason_map = {
        'manja': 'tgl_manja_awal tersedia (durasi negatif di-clamp ke 0 = Comply)',
        'pi':    'tgl_manja_awal kosong → pakai tgl_pi (durasi negatif di-clamp ke 0 = Comply)',
        're':    'tgl_manja_awal & tgl_pi kosong → pakai tgl_reg (durasi negatif di-clamp ke 0 = Comply)',
        'anomaly': 'tgl_pc kosong atau semua referensi (Manja/PI/RE) kosong',
    }
    df['_ref_reason'] = df['_category'].map(_reason_map).fillna('tidak diketahui')
    return df


def _fulfill_mo_formula(df):
    """
    Fulfill MO SLA — KPI 2026 menu.
    Source file convention differs from other entities: MO ships as a single
    multi-month workbook (e.g. 202601_202606_MO.xlsx) that gets split into
    one dataframe per month via split_ranged_file() (using c_statusdate)
    BEFORE this formula ever runs — by the time apply_sla_formula() calls
    this, df is already scoped to a single period.

    Same 1x24h + clamp-negative-to-Comply pattern as Fulfill AO, different
    column names:
      - Comply threshold: 1x24 hours (24h).
      - Duration = c_statusdate (completion) minus a reference date, tried
        in order:
          1. tanggal_manja      (Manja)
          2. date_pi_startwork  (only if Manja is NULL)
          3. datecreated        (only if both Manja and PI are NULL)
        A NEGATIVE duration clamps to 0 and is flagged Comply immediately —
        it does NOT fall through to the next reference.
      - c_statusdate empty → anomaly (excluded from SLA).
    """
    for col in ['c_statusdate', 'tanggal_manja', 'date_pi_startwork', 'datecreated']:
        df[col] = pd.to_datetime(df.get(col, pd.NaT), errors='coerce') \
                  if col in df.columns else pd.NaT

    def hrs(a, b): return (a - b).dt.total_seconds() / 3600
    dm = hrs(df['c_statusdate'], df['tanggal_manja'])
    dp = hrs(df['c_statusdate'], df['date_pi_startwork'])
    dc = hrs(df['c_statusdate'], df['datecreated'])

    has_m = df['tanggal_manja'].notna()
    has_p = df['date_pi_startwork'].notna()
    has_c = df['datecreated'].notna()
    empty = df['c_statusdate'].isna()

    use_m = has_m
    use_p = ~has_m & has_p
    use_c = ~has_m & ~has_p & has_c

    ref     = pd.Series(np.select([use_m, use_p, use_c], ['Manja', 'PI', 'Created'], default='Anomaly'), index=df.index)
    raw_dur = pd.Series(np.select([use_m, use_p, use_c], [dm, dp, dc],                default=np.nan),    index=df.index)
    # Negative duration (completed at/before the reference) → clamp to 0 (instant Comply)
    dur = raw_dur.clip(lower=0)

    has = ref.isin(['Manja', 'PI', 'Created'])
    df['_comply'] = pd.Series(
        np.select([empty, has & (dur <= 24), has & (dur > 24)],
                  ['anomaly statusdate empty', 'Comply', 'Not Comply'], default='anomaly all minus'),
        index=df.index)

    df['_category'] = pd.Series(np.select(
        [use_m, use_p, use_c], ['manja', 'pi', 'created'], default='anomaly'
    ), index=df.index)

    # ── Detail columns for per-row export ────────────────────────────────────
    df['_ref_start']      = ref
    df['_duration_h']     = dur.round(2)       # clamped duration used for SLA
    df['_raw_duration_h'] = raw_dur.round(2)   # unclamped, for auditing negatives
    _reason_map = {
        'manja':   'tanggal_manja tersedia (durasi negatif di-clamp ke 0 = Comply)',
        'pi':      'tanggal_manja kosong → pakai date_pi_startwork (durasi negatif di-clamp ke 0 = Comply)',
        'created': 'tanggal_manja & date_pi_startwork kosong → pakai datecreated (durasi negatif di-clamp ke 0 = Comply)',
        'anomaly': 'c_statusdate kosong atau semua referensi (Manja/PI/Created) kosong',
    }
    df['_ref_reason'] = df['_category'].map(_reason_map).fillna('tidak diketahui')
    return df


def _fulfill_pda_formula(df):
    """
    Fulfill PDA SLA — KPI 2026 menu.
    Source file convention: like MO, PDA ships as a single multi-month
    workbook (e.g. 202601_202606_PDA.xlsx) split into one dataframe per
    month via split_ranged_file() (using orca_provcomp_date) BEFORE this
    formula ever runs.

    Comply threshold depends on orca_order_type, per row:
      - 'migrate' (same STO / satu STO)      → 1x24h (24h)
      - 'create'  (different STO / beda STO) → 2x24h (48h)
      - any other/unrecognized order type    → flagged anomaly (does not
        default silently to either threshold — see 'anomaly unknown order
        type' below)

    Duration = orca_provcomp_date (completion) minus a reference date, tried
    in order:
      1. bima_customerassign_date
      2. orca_pi_date               (only if bima_customerassign_date NULL)
      3. orca_external_created_dt   (only if both above are NULL)
    A NEGATIVE duration clamps to 0 and is flagged Comply immediately — it
    does NOT fall through to the next reference.
    orca_provcomp_date empty → anomaly (excluded from SLA).
    """
    for col in ['orca_provcomp_date', 'bima_customerassign_date', 'orca_pi_date', 'orca_external_created_dt']:
        df[col] = pd.to_datetime(df.get(col, pd.NaT), errors='coerce') \
                  if col in df.columns else pd.NaT

    def hrs(a, b): return (a - b).dt.total_seconds() / 3600
    dm = hrs(df['orca_provcomp_date'], df['bima_customerassign_date'])
    dp = hrs(df['orca_provcomp_date'], df['orca_pi_date'])
    dc = hrs(df['orca_provcomp_date'], df['orca_external_created_dt'])

    has_m = df['bima_customerassign_date'].notna()
    has_p = df['orca_pi_date'].notna()
    has_c = df['orca_external_created_dt'].notna()
    empty = df['orca_provcomp_date'].isna()

    use_m = has_m
    use_p = ~has_m & has_p
    use_c = ~has_m & ~has_p & has_c

    ref     = pd.Series(np.select([use_m, use_p, use_c], ['CustomerAssign', 'PI', 'Created'], default='Anomaly'), index=df.index)
    raw_dur = pd.Series(np.select([use_m, use_p, use_c], [dm, dp, dc],                        default=np.nan),    index=df.index)
    # Negative duration (completed at/before the reference) → clamp to 0 (instant Comply)
    dur = raw_dur.clip(lower=0)

    # Per-row comply threshold: migrate (same STO) = 24h, create (different STO) = 48h
    order_type = df.get('orca_order_type', pd.Series('', index=df.index)).astype(str).str.strip().str.lower()
    is_migrate  = order_type == 'migrate'
    is_create   = order_type == 'create'
    known_type  = is_migrate | is_create
    threshold   = pd.Series(np.select([is_migrate, is_create], [24, 48], default=np.nan), index=df.index)

    has = ref.isin(['CustomerAssign', 'PI', 'Created'])
    df['_comply'] = pd.Series(
        np.select(
            [empty, has & ~known_type, has & known_type & (dur <= threshold), has & known_type & (dur > threshold)],
            ['anomaly provcomp empty', 'anomaly unknown order type', 'Comply', 'Not Comply'],
            default='anomaly all minus'),
        index=df.index)

    df['_category'] = pd.Series(np.select(
        [use_m, use_p, use_c], ['customerassign', 'pi', 'created'], default='anomaly'
    ), index=df.index)

    # ── Detail columns for per-row export ────────────────────────────────────
    df['_order_type']      = order_type
    df['_sla_threshold_h'] = threshold
    df['_ref_start']       = ref
    df['_duration_h']      = dur.round(2)       # clamped duration used for SLA
    df['_raw_duration_h']  = raw_dur.round(2)   # unclamped, for auditing negatives
    _reason_map = {
        'customerassign': 'bima_customerassign_date tersedia (durasi negatif di-clamp ke 0 = Comply)',
        'pi':             'bima_customerassign_date kosong → pakai orca_pi_date (durasi negatif di-clamp ke 0 = Comply)',
        'created':        'bima_customerassign_date & orca_pi_date kosong → pakai orca_external_created_dt (durasi negatif di-clamp ke 0 = Comply)',
        'anomaly':        'orca_provcomp_date kosong, tipe order tidak dikenal (bukan migrate/create), atau semua referensi kosong',
    }
    df['_ref_reason'] = df['_category'].map(_reason_map).fillna('tidak diketahui')
    return df


def _tti_formula(df):
    """TTI SLA: duration from Manja/PI/RE ref to tgl_pc ≤ 72h."""
    for col in ['tgl_pc', 'tgl_manja_awal', 'tgl_pi', 'tgl_reg']:
        df[col] = pd.to_datetime(df.get(col, pd.NaT), errors='coerce') \
                  if col in df.columns else pd.NaT

    def hrs(a, b): return (a - b).dt.total_seconds() / 3600
    dm = hrs(df['tgl_pc'], df['tgl_manja_awal'])
    dp = hrs(df['tgl_pc'], df['tgl_pi'])
    dr = hrs(df['tgl_pc'], df['tgl_reg'])

    ok_m = df['tgl_manja_awal'].notna() & (dm >= 0)
    ok_p = df['tgl_pi'].notna()          & (dp >= 0)
    ok_r = df['tgl_reg'].notna()         & (dr >= 0)
    empty = df['tgl_pc'].isna()

    dur = pd.Series(np.select([ok_m,ok_p,ok_r],[dm.round(4),dp.round(4),dr.round(4)],default=np.nan),index=df.index)
    ref = pd.Series(np.select([ok_m,ok_p,ok_r,empty],['Manja','PI','RE','Anomaly'],default='Anomaly'),index=df.index)
    has = ref.isin(['Manja','PI','RE'])
    df['_comply'] = pd.Series(
        np.select([empty, has&(dur<=72), has&(dur>72)],
                  ['anomaly pc empty','Comply','Not Comply'], default='anomaly all minus'),
        index=df.index)

    manja_minus = df['tgl_manja_awal'].notna() & (dm < 0)
    manja_null  = df['tgl_manja_awal'].isna()
    pi_minus    = df['tgl_pi'].notna() & (dp < 0)
    pi_null     = df['tgl_pi'].isna()
    df['_category'] = pd.Series(np.select(
        [ok_m,
         ~ok_m & ok_p & manja_minus,
         ~ok_m & ok_p & manja_null,
         ~ok_m & ~ok_p & ok_r & pi_minus,
         ~ok_m & ~ok_p & ok_r & pi_null],
        ['manja','pi_manja_minus','pi_manja_null','re_pi_minus','re_pi_null'],
        default='anomaly'
    ), index=df.index)

    # ── Detail columns for per-row export ────────────────────────────────────
    df['_ref_start']  = ref   # Manja / PI / RE / Anomaly
    df['_duration_h'] = dur.round(2)   # duration in hours (NaN for anomaly rows)
    _reason_map = {
        'manja':          'tgl_manja_awal valid & positif',
        'pi_manja_minus': 'tgl_manja_awal negatif → pakai tgl_pi',
        'pi_manja_null':  'tgl_manja_awal kosong → pakai tgl_pi',
        're_pi_minus':    'tgl_pi negatif → pakai tgl_reg',
        're_pi_null':     'tgl_pi kosong → pakai tgl_reg',
        'anomaly':        'tgl_pc kosong atau semua referensi tidak valid',
    }
    df['_ref_reason'] = df['_category'].map(_reason_map).fillna('tidak diketahui')
    return df


def _fmt_eu(val):
    """Format a float as European number string: period thousands, comma decimal.
    e.g. 1234.56 → '1.234,56' ; 51.06 → '51,06'. Used for Mac Excel."""
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ''
        return f"{float(val):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return ''


def build_tti_raw_csv(df, period):
    """
    Export the raw/original TTI columns (as loaded from DB, after dedup).
    Uses semicolon delimiter so Mac Excel reads European-formatted numbers correctly.
    """
    # Exclude internal computed columns (_comply, _category, _ref_start, etc.)
    raw_cols = [c for c in df.columns if not c.startswith('_')]
    out = df[raw_cols].copy()
    out.insert(0, 'Periode', period)
    return out.to_csv(index=False, sep=';')


def build_tti_detail_csv(df, period):
    """
    Build a per-row TTI PROCESSED detail CSV for dashboard download.
    Columns: Periode, Area, Regional, Branch, STO, Order_ID,
             Reference_Start, Ref_Reason, Duration_Hours, SLA_Status.
    Uses semicolon delimiter + European number format (. thousands, , decimal)
    so Mac Excel opens it correctly without import wizard.
    """
    col_map = {
        'tsel_area':           'tsel_area',
        'tsel_regional':       'tsel_regional',
        'tsel_branch':         'tsel_branch',
        'tsel_cluster':        'tsel_cluster',
        'tsel_region_network': 'tsel_region_network',
        'tsel_nop':            'tsel_nop',
        'sto':                 'sto',
        'order_id':            'order_id',
        '_ref_start':          'reference_for_counting',
        '_ref_reason':         'reason_of_reference',
        '_duration_h':         'Duration_Hours',
        '_comply':             'SLA_Status',
    }
    sel = {k: v for k, v in col_map.items() if k in df.columns}
    out = df[list(sel.keys())].rename(columns=sel).copy()
    out.insert(0, 'Periode', period)
    # Clean SLA_Status labels
    if 'SLA_Status' in out.columns:
        out['SLA_Status'] = out['SLA_Status'].replace({
            'anomaly pc empty':  'Anomaly',
            'anomaly all minus': 'Anomaly',
            'anomaly':           'Anomaly',
        })
    # European number format for Duration_Hours
    if 'Duration_Hours' in out.columns:
        out['Duration_Hours'] = out['Duration_Hours'].apply(_fmt_eu)
    # Semicolon delimiter — standard for European Mac Excel
    return out.to_csv(index=False, sep=';')


def build_mo_raw_csv(df, period):
    """Export the raw/original Fulfill MO columns. Semicolon CSV for Mac Excel."""
    raw_cols = [c for c in df.columns if not c.startswith('_')]
    out = df[raw_cols].copy()
    out.insert(0, 'Periode', period)
    return out.to_csv(index=False, sep=';')


def build_mo_detail_csv(df, period):
    """
    Per-row Fulfill MO PROCESSED detail CSV — same shape as build_tti_detail_csv,
    keyed off MO's own reference/duration/comply columns from _fulfill_mo_formula().
    """
    col_map = {
        'tsel_area':           'tsel_area',
        'tsel_regional':       'tsel_regional',
        'tsel_branch':         'tsel_branch',
        'sto':                 'sto',
        '_ref_start':          'reference_for_counting',
        '_ref_reason':         'reason_of_reference',
        '_duration_h':         'Duration_Hours',
        '_comply':             'SLA_Status',
    }
    sel = {k: v for k, v in col_map.items() if k in df.columns}
    out = df[list(sel.keys())].rename(columns=sel).copy()
    out.insert(0, 'Periode', period)
    if 'SLA_Status' in out.columns:
        out['SLA_Status'] = out['SLA_Status'].replace({
            'anomaly statusdate empty': 'Anomaly',
            'anomaly all minus':        'Anomaly',
            'anomaly':                  'Anomaly',
        })
    if 'Duration_Hours' in out.columns:
        out['Duration_Hours'] = out['Duration_Hours'].apply(_fmt_eu)
    return out.to_csv(index=False, sep=';')


def build_pda_raw_csv(df, period):
    """Export the raw/original Fulfill PDA columns. Semicolon CSV for Mac Excel."""
    raw_cols = [c for c in df.columns if not c.startswith('_')]
    out = df[raw_cols].copy()
    out.insert(0, 'Periode', period)
    return out.to_csv(index=False, sep=';')


def build_pda_detail_csv(df, period):
    """
    Per-row Fulfill PDA PROCESSED detail CSV — includes the per-row order type
    and SLA threshold (24h migrate / 48h create) alongside the usual
    reference/duration/comply columns from _fulfill_pda_formula().
    """
    col_map = {
        'tsel_area':           'tsel_area',
        'tsel_regional':       'tsel_regional',
        'tsel_branch':         'tsel_branch',
        'sto':                 'sto',
        '_order_type':         'Order_Type',
        '_sla_threshold_h':    'SLA_Threshold_Hours',
        '_ref_start':          'reference_for_counting',
        '_ref_reason':         'reason_of_reference',
        '_duration_h':         'Duration_Hours',
        '_comply':             'SLA_Status',
    }
    sel = {k: v for k, v in col_map.items() if k in df.columns}
    out = df[list(sel.keys())].rename(columns=sel).copy()
    out.insert(0, 'Periode', period)
    if 'SLA_Status' in out.columns:
        out['SLA_Status'] = out['SLA_Status'].replace({
            'anomaly provcomp empty':      'Anomaly',
            'anomaly unknown order type':  'Anomaly',
            'anomaly all minus':           'Anomaly',
            'anomaly':                     'Anomaly',
        })
    for col in ('Duration_Hours', 'SLA_Threshold_Hours'):
        if col in out.columns:
            out[col] = out[col].apply(_fmt_eu)
    return out.to_csv(index=False, sep=';')


def build_ffg_raw_csv(df, period):
    """
    Export the raw/original FFG columns (as loaded from DB, after dedup).
    Uses semicolon delimiter so Mac Excel reads correctly.
    """
    raw_cols = [c for c in df.columns if not c.startswith('_')]
    out = df[raw_cols].copy()
    out.insert(0, 'Periode', period)
    return out.to_csv(index=False, sep=';')


def build_ffg_detail_csv(df, period):
    """
    Build a per-row FFG PROCESSED detail CSV for dashboard download.
    Columns: Periode, org_1, org_2, org_3, sto, order_id, trouble_no, flag_exclude_all, Wilsus.
    Uses semicolon delimiter for Mac Excel compatibility.

    - org_1/2/3 come from tsel_area/tsel_regional/tsel_branch (after _normalize_hierarchy)
    - Wilsus = 1 if f_ttr column == 'WILSUS' (case-insensitive), else 0
    """
    wanted_src = [
        ('tsel_area',     'org_1'),
        ('tsel_regional', 'org_2'),
        ('tsel_branch',   'org_3'),
        ('sto',           'sto'),
        ('order_id',      'order_id'),
        ('trouble_no',    'trouble_no'),
        ('flag_exclude_all', 'flag_exclude_all'),
    ]
    src_cols  = [s for s, _ in wanted_src if s in df.columns]
    rename_map = {s: d for s, d in wanted_src if s in df.columns}

    result = df[src_cols].rename(columns=rename_map).copy()
    result.insert(0, 'Periode', period)

    # is_duplicate_order: 1 if same order_id appears more than once in the dataset
    if 'order_id' in df.columns:
        order_counts = df.groupby('order_id').size()
        result['is_duplicate_order'] = df['order_id'].map(order_counts).gt(1).astype(int).values
    else:
        result['is_duplicate_order'] = 0

    # Wilsus: 1 if f_ttr == 'WILSUS', else 0
    if 'f_ttr' in df.columns:
        result['Wilsus'] = (
            df['f_ttr'].astype(str).str.strip().str.upper() == 'WILSUS'
        ).astype(int).values
    else:
        result['Wilsus'] = 0

    return result.to_csv(index=False, sep=';')


def build_ttrffg_raw_csv(df, period):
    """
    Export the raw/original TTRFFG columns (as loaded from DB, after dedup).
    Uses semicolon delimiter so Mac Excel reads correctly.
    """
    raw_cols = [c for c in df.columns if not c.startswith('_')]
    out = df[raw_cols].copy()
    out.insert(0, 'Periode', period)
    return out.to_csv(index=False, sep=';')


def build_ttrffg_detail_csv(df, period):
    """
    Build a per-row TTRFFG PROCESSED detail CSV for dashboard download.
    Columns: Periode, Area, Witel, Branch, STO, Order_ID,
             Flag_Exclude, Duration_Hours, SLA_Status.
    Duration_Hours = date_close - first_cust_assign (hours).
    Uses semicolon delimiter + European number format for Mac Excel.
    """
    # Recompute duration (not stored on df by _ttrffg_formula)
    d_close  = pd.to_datetime(df['date_close'],        errors='coerce') if 'date_close'        in df.columns else pd.NaT
    d_assign = pd.to_datetime(df['first_cust_assign'], errors='coerce') if 'first_cust_assign' in df.columns else pd.NaT
    dur_h    = (d_close - d_assign).dt.total_seconds() / 3600

    col_map = {
        'tsel_area':        'Area',
        'tsel_regional':    'Witel',
        'tsel_branch':      'Branch',
        'sto':              'STO',
        'order_id':         'Order_ID',
        'flag_exclude_all': 'Flag_Exclude',
        '_comply':          'SLA_Status',
    }
    sel = {k: v for k, v in col_map.items() if k in df.columns}
    out = df[list(sel.keys())].rename(columns=sel).copy()
    out.insert(0, 'Periode', period)
    # Insert Duration_Hours after other columns
    out['Duration_Hours'] = dur_h.values
    out['Duration_Hours'] = out['Duration_Hours'].apply(_fmt_eu)
    # Clean SLA_Status
    if 'SLA_Status' in out.columns:
        out['SLA_Status'] = out['SLA_Status'].replace({'anomaly': 'Anomaly'})
    return out.to_csv(index=False, sep=';')


def build_psre_raw_csv(df, period):
    """
    Export the raw/original PS/RE columns (as loaded from DB, after dedup & AO filter).
    Uses semicolon delimiter so Mac Excel reads correctly.
    """
    raw_cols = [c for c in df.columns if not c.startswith('_')]
    out = df[raw_cols].copy()
    out.insert(0, 'Periode', period)
    return out.to_csv(index=False, sep=';')


def build_psre_detail_csv(df, period):
    """
    Build a per-row PS/RE PROCESSED detail CSV for dashboard download.
    Columns: Periode, order_id, tsel_area, tsel_regional, tsel_branch, sto, service_id, Status.
    Uses semicolon delimiter for Mac Excel compatibility.
    Status values: Comply | Not Comply | Anomaly (Cancel <7hr) | Anomaly (Complete Exclude)
    """
    wanted = [
        ('order_id',      'order_id'),
        ('tsel_area',     'tsel_area'),
        ('tsel_regional', 'tsel_regional'),
        ('tsel_branch',   'tsel_branch'),
        ('sto',           'sto'),
        ('service_id',    'service_id'),
        ('_comply',       'Status'),
    ]
    src_cols   = [s for s, _ in wanted if s in df.columns]
    rename_map = {s: d for s, d in wanted if s in df.columns}

    out = df[src_cols].rename(columns=rename_map).copy()
    out.insert(0, 'Periode', period)

    # Readable status labels
    if 'Status' in out.columns:
        out['Status'] = out['Status'].replace({
            'anomaly cancel 7hari':     'Anomaly (Cancel <7hr)',
            'anomaly complete exclude':  'Anomaly (Complete Exclude)',
        })
    return out.to_csv(index=False, sep=';')


def _psre_formula(df):
    """PS/RE SLA: Complete AO / (Total AO − Cancel<7hari − Complete Exclude)."""
    if 'transaksi' in df.columns:
        df.attrs['re_total'] = len(df)   # all rows before AO filter (RE + AO)
        df = df[df['transaksi'].str.strip().str.upper() == 'AO'].copy()
        print(f"   (filtered to AO: {len(df):,} rows)", end="  ", flush=True)

    for col in ['tgl_pc', 'tgl_ps', 'tgl_ps_complete', 'tgl_cancel']:
        df[col] = pd.to_datetime(df.get(col, pd.NaT), errors='coerce') \
                  if col in df.columns else pd.NaT

    f7 = pd.to_numeric(
        df['f_cancel_7hari'] if 'f_cancel_7hari' in df.columns
        else pd.Series(0, index=df.index),
        errors='coerce'
    ).fillna(0)

    cancel_filled  = df['tgl_cancel'].notna()
    cancel_7hari   = cancel_filled & (f7 == 1)
    cancel_regular = cancel_filled & ~cancel_7hari
    any_filled     = df['tgl_pc'].notna() | df['tgl_ps'].notna() | df['tgl_ps_complete'].notna()
    ps_filled      = df['tgl_ps'].notna() | df['tgl_ps_complete'].notna()
    complete       = ~cancel_filled & any_filled & ps_filled
    comp_exclude   = ~cancel_filled & any_filled & ~ps_filled

    df['_comply'] = pd.Series(np.select(
        [cancel_7hari, cancel_regular, complete, comp_exclude],
        ['anomaly cancel 7hari', 'Not Comply', 'Comply', 'anomaly complete exclude'],
        default='Not Comply'
    ), index=df.index)
    df['_category'] = pd.Series(np.select(
        [cancel_7hari, cancel_regular, complete, comp_exclude],
        ['cancel_7hari', 'cancel', 'complete', 'complete_exclude'],
        default='on_progress'
    ), index=df.index)
    return df


def _ttrffg_formula(df):
    """
    TTR FFG SLA.
    Filter: flag_exclude_all = 0  AND  first_cust_assign is not null.
    Duration = date_close - first_cust_assign (hours).
    ≤ 3 h → Comply  |  > 3 h → Not Comply  |  all others → anomaly (excluded).
    SLA = Comply / (Comply + Not Comply).

    NOTE (2026-08): a change dropping the flag_exclude_all==0 requirement
    was tested and verified to match a manual reference reconciliation for
    period 202607 exactly (644 Comply / 86 Not Comply / 88.22% SLA vs this
    version's 315/26/92.38%) — but per explicit user instruction the
    flag_exclude_all gate was kept/restored anyway, so this file's SLA %
    intentionally does NOT match that reconciliation. See conversation
    history if revisiting this.
    """
    flag = pd.to_numeric(
        df['flag_exclude_all'] if 'flag_exclude_all' in df.columns
        else pd.Series(1, index=df.index),
        errors='coerce'
    ).fillna(1).astype(int)

    d_close  = pd.to_datetime(df['date_close'],       errors='coerce') if 'date_close'       in df.columns else pd.NaT
    d_assign = pd.to_datetime(df['first_cust_assign'],errors='coerce') if 'first_cust_assign' in df.columns else pd.NaT

    dur_h = (d_close - d_assign).dt.total_seconds() / 3600

    valid = (flag == 0) & d_assign.notna() & d_close.notna()

    # WILSUS rows are excluded from SLA (anomaly), regardless of duration
    is_wilsus = df['f_ttr'].str.strip().str.upper() == 'WILSUS' \
                if 'f_ttr' in df.columns else pd.Series(False, index=df.index)

    df['_comply'] = 'anomaly'
    df.loc[valid & ~is_wilsus & (dur_h <= 3), '_comply'] = 'Comply'
    df.loc[valid & ~is_wilsus & (dur_h >  3), '_comply'] = 'Not Comply'
    # valid & is_wilsus → stays 'anomaly' (WILSUS excluded from denominator)
    return df


def _ffggaul_formula(df):
    """
    FFG GAUL SLA.
    Filter:      flag_exclude_all = 0.
    FFG          = unique order_ids among flag=0 rows.
    Total ticket = count of flag=0 rows.
    GAUL         = Total ticket - FFG  (duplicate order_id = repeat ticket).
    SLA          = FFG / Total ticket.

    Implementation: flag=1 rows → anomaly (excluded).
    Among flag=0: first occurrence of each order_id → Comply (FFG),
    subsequent occurrences → Not Comply (GAUL).
    """
    if 'order_id' not in df.columns:
        print("   ⚠  FFG GAUL: 'order_id' column not found")
        df['_comply'] = 'anomaly'
        return df

    flag = pd.to_numeric(
        df['flag_exclude_all'] if 'flag_exclude_all' in df.columns
        else pd.Series(0, index=df.index),
        errors='coerce'
    ).fillna(1).astype(int)

    df['_comply'] = 'anomaly'   # flag=1 rows excluded
    valid_idx = df.index[flag == 0]
    # Among valid rows: first occurrence of order_id → Comply, rest → Not Comply
    is_dup = df.loc[valid_idx].duplicated(subset=['order_id'], keep='first')
    df.loc[valid_idx, '_comply'] = is_dup.map({False: 'Comply', True: 'Not Comply'})
    return df


# ── KPI file split ─────────────────────────────────────────────────────────
# 2025 and 2026 use different formulas even for KPIs sharing the same entity
# name (legacy TTI vs Fulfill AO, FFG-with-exclusions vs FFG-no-exclusions,
# etc.), so their results are pushed to two SEPARATE JSON files rather than
# one shared kpis.json. This is what actually keeps the two menus' numbers
# from ever mixing — NOT a date cutover: for TTI and FFG, one source file
# for a period >= KPI_2026_CUTOVER_PERIOD is run through BOTH formulas (see
# entity_formula_versions()) — the legacy 3x24h TTI rule / FFG-with-
# exclusions rule for KPI 2025, and the new Fulfill AO 1x24h rule / FFG-no-
# exclusions rule for KPI 2026 — and each result goes to its own file. So
# the same calendar period can legitimately appear in BOTH kpis_2025.json
# (under the legacy formula) and kpis_2026.json (under the 2026 formula) —
# that's by design, not a bug.
KPI_2026_CUTOVER_PERIOD = "202601"
KPI_2025_FILE = "data/kpis_2025.json"
KPI_2026_FILE = "data/kpis_2026.json"


def entity_formula_versions(entity, period):
    """
    Which formula version(s) this entity/period combination should be
    computed with, and therefore which output file(s) it lands on:
      'legacy' → data/kpis_2025.json (KPI 2025 menu) — the original formula,
                 unchanged, for every period regardless of calendar year.
      '2026'   → data/kpis_2026.json (KPI 2026 menu) — the new formula.

    TTI and FFG are the two entities that have BOTH a legacy and a 2026
    formula. For periods >= KPI_2026_CUTOVER_PERIOD, the SAME source file is
    run through both formulas independently — e.g. 202601_TTI.csv produces
    one result for KPI 2025 (legacy 3x24h rule) and a separate result for
    KPI 2026 (Fulfill AO 1x24h rule). This is intentional, per explicit
    instruction: KPI 2025 and KPI 2026 must never share a formula even when
    reading the identical file.

    MO/PDA only ever have a 2026 formula (no legacy version was ever
    defined for them) — they stay 2026-only regardless of period.
    PSRE/TTRFFG/FFGGAUL have no 2026 formula/tab defined at all — they stay
    legacy-only regardless of period.
    """
    is_2026_eligible = bool(period) and period >= KPI_2026_CUTOVER_PERIOD
    if entity in ("TTI", "FFG"):
        return ["legacy", "2026"] if is_2026_eligible else ["legacy"]
    if entity in ("MO", "PDA"):
        return ["2026"]
    return ["legacy"]


def prepare_df(df, entity):
    """
    Load-time normalization shared by every formula version: lower-case
    columns, hierarchy rename, dedup. This is VERSION-INDEPENDENT — it's the
    same cleaned dataframe whether the legacy or 2026 formula is about to be
    applied to it, so it only needs to run once per (entity, period) even
    when that entity computes two formula versions from the same file.
    """
    df = df.copy()
    df.columns = [c.strip().lower() for c in df.columns]
    df = _normalize_hierarchy(df)
    if entity == "MO":
        df = _apply_sto_hierarchy(df)

    # NO_DEDUP_ENTITIES (MO, PDA): known to have no reliable unique row
    # identifier — dedup is intentionally skipped, no warning needed.
    if entity not in NO_DEDUP_ENTITIES:
        dedup_col = DEDUP_COLUMN_FFG if entity in ("FFG", "TTRFFG", "FFGGAUL") else DEDUP_COLUMN
        if dedup_col:
            col = dedup_col.lower()
            if col in df.columns:
                before = len(df)
                df = df.drop_duplicates(subset=[col])
                removed = before - len(df)
                if removed:
                    print(f"   🔁 Removed {removed:,} duplicate rows (column: {col})")
            else:
                print(f"   ⚠  Dedup column '{col}' not found — skipping dedup")
    return df


def apply_sla_formula(df, entity="TTI", formula_version="legacy"):
    """
    Apply the SLA formula for ONE (entity, formula_version) combination to
    an ALREADY-prepared dataframe (see prepare_df — normalize/dedup already
    done, so this can safely be called twice on two copies of the same
    prepared df when an entity has both a legacy and a 2026 formula).
    """
    df = df.copy()
    if entity == "PSRE":    return _psre_formula(df)
    if entity == "TTRFFG":  return _ttrffg_formula(df)
    if entity == "FFGGAUL": return _ffggaul_formula(df)
    if entity == "FFG":     return df   # FFG uses separate aggregate_ffg()
    if entity == "MO":      return _fulfill_mo_formula(df)   # KPI 2026 menu: "Fulfill MO", 1x24h rule
    if entity == "PDA":     return _fulfill_pda_formula(df)  # KPI 2026 menu: "Fulfill PDA", 1x24h/2x24h rule
    if entity == "TTI":
        if formula_version == "2026":
            return _fulfill_ao_formula(df)   # KPI 2026 menu: "Fulfill AO", 1x24h rule
        return _tti_formula(df)              # KPI 2025 menu: legacy 3x24h rule
    return df


# ─── Statistics helpers ────────────────────────────────────────────────────────

def stats(grp):
    c  = int((grp['_comply'] == 'Comply').sum())
    nc = int((grp['_comply'] == 'Not Comply').sum())
    an = int((~grp['_comply'].isin(['Comply', 'Not Comply'])).sum())
    pct = round(c / (c + nc) * 100, 2) if (c + nc) > 0 else 0.0
    return c, nc, an, pct


# Explicit, fixed category schemes per (entity, formula) — no guessing from
# whichever categories happen to be present in a given area's rows. Using a
# fixed scheme per formula guarantees legacy TTI (pre-2026) always reports
# its original 5-key breakdown even for a small area where only 'manja' rows
# happen to occur, instead of silently shrinking to fewer keys.
BREAKDOWN_SCHEME_LEGACY_TTI = ['manja', 'pi_manja_minus', 'pi_manja_null', 're_pi_minus', 're_pi_null']
BREAKDOWN_SCHEME_FULFILL_AO = ['manja', 'pi', 're']       # Fulfill AO (TTI, periods >= cutover)
BREAKDOWN_SCHEME_FULFILL_MO = ['manja', 'pi', 'created']  # Fulfill MO
BREAKDOWN_SCHEME_FULFILL_PDA = ['customerassign', 'pi', 'created']  # Fulfill PDA


def breakdown_scheme_for(entity, formula_version="legacy"):
    """
    Which fixed category scheme applies to this entity/formula_version
    combination. Returns None for entities with no reference breakdown at
    all. Keyed by formula_version (not period) for the same reason as
    resolve_target() — TTI can produce a legacy-formula AND a 2026-formula
    result for the identical period, and each needs its own scheme.
    """
    if entity == "TTI":
        return BREAKDOWN_SCHEME_FULFILL_AO if formula_version == "2026" else BREAKDOWN_SCHEME_LEGACY_TTI
    if entity == "MO":
        return BREAKDOWN_SCHEME_FULFILL_MO
    if entity == "PDA":
        return BREAKDOWN_SCHEME_FULFILL_PDA
    return None


def breakdown_stats(grp, cats):
    """
    Reference breakdown (which date column produced the duration for each
    row), using the FIXED category list `cats` given by breakdown_scheme_for
    — never inferred from what happens to be present in this particular
    group, so the same entity/period always reports the same set of keys
    regardless of how small or skewed an individual area's data is.
    """
    if '_category' not in grp.columns or not cats:
        return None
    result = {}
    for cat in cats:
        mask = grp['_category'] == cat
        result[cat] = {
            'comply':     int((mask & (grp['_comply'] == 'Comply')).sum()),
            'not_comply': int((mask & (grp['_comply'] == 'Not Comply')).sum()),
        }
    result['anomaly'] = int((grp['_category'] == 'anomaly').sum())
    return result


# ─── Standard aggregation (TTI / PSRE / TTRFFG / FFGGAUL) ────────────────────

def aggregate_entity(df, entity, formula_version="legacy"):
    c, nc, an, pct = stats(df)
    total = c + nc + an

    by_area = []
    breakdown_by_area = []

    area_col = 'tsel_area'     if 'tsel_area'     in df.columns else None
    reg_col  = 'tsel_regional' if 'tsel_regional' in df.columns else None
    br_col   = 'tsel_branch'   if 'tsel_branch'   in df.columns else None
    sto_col  = 'sto'           if 'sto'           in df.columns else None

    for area, ag in (df.groupby(area_col) if area_col else []):
        ac, anc, aan, apct = stats(ag)
        by_reg = []
        for reg, rg in (ag.groupby(reg_col) if reg_col else []):
            rc, rnc, ran, rpct = stats(rg)
            by_br = []
            for br, bg in (rg.groupby(br_col) if br_col else []):
                bc, bnc, ban, bpct = stats(bg)
                by_sto = []
                for sto, sg in (bg.groupby(sto_col) if sto_col else []):
                    sc, snc, san, spct = stats(sg)
                    by_sto.append({"sto":str(sto),"comply":sc,"not_comply":snc,"anomaly":san,"sla_pct":spct})
                by_sto.sort(key=lambda x: x['sto'])
                row = {"branch":str(br),"comply":bc,"not_comply":bnc,"anomaly":ban,"sla_pct":bpct}
                if by_sto: row["by_sto"] = by_sto
                by_br.append(row)
            by_br.sort(key=lambda x: x['branch'])
            by_reg.append({"regional":str(reg),"comply":rc,"not_comply":rnc,"anomaly":ran,"sla_pct":rpct,"by_branch":by_br})
        by_reg.sort(key=lambda x: x['regional'])
        by_area.append({"area":str(area),"comply":ac,"not_comply":anc,"anomaly":aan,"sla_pct":apct,"by_regional":by_reg})

        scheme = breakdown_scheme_for(entity, formula_version)
        if scheme:
            bd = breakdown_stats(ag, scheme)
            if bd:
                breakdown_by_area.append({"area": str(area), **bd})

    by_area.sort(key=lambda x: x['area'])
    breakdown_by_area.sort(key=lambda x: x['area'])

    result = {
        "total": total, "comply": c, "not_comply": nc, "anomaly": an,
        "sla_pct": pct, "target": resolve_target(entity, formula_version),
        "by_area": by_area
    }
    if breakdown_by_area:
        result["breakdown_by_area"] = breakdown_by_area

    # PS/RE breakdown card
    if entity == "PSRE" and '_category' in df.columns:
        re_total   = df.attrs.get('re_total', len(df))
        re_ao      = len(df)
        cancel_7h  = int((df['_category'] == 'cancel_7hari').sum())
        ps_ao      = c   # Comply = complete
        not_ps_ao  = nc  # Not Comply = cancel_regular + on_progress
        exclude_ao = int((df['_category'] == 'complete_exclude').sum())
        re_ao_nett = re_ao - exclude_ao
        result['psre_breakdown'] = {
            're_total':   re_total,
            're_ao':      re_ao,
            're_ao_nett': re_ao_nett,
            'cancel_7hari': cancel_7h,
            'ps_ao':      ps_ao,
            'not_ps_ao':  not_ps_ao,
            'exclude_ao': exclude_ao,
        }

    # WILSUS count for TTRFFG: rows where f_ttr == 'WILSUS' AND first_cust_assign is not null
    if entity == "TTRFFG":
        is_wilsus = df['f_ttr'].str.strip().str.upper() == 'WILSUS' \
                    if 'f_ttr' in df.columns else pd.Series(False, index=df.index)
        has_assign = pd.to_datetime(
            df['first_cust_assign'] if 'first_cust_assign' in df.columns else pd.NaT,
            errors='coerce'
        ).notna()
        wilsus_mask = is_wilsus & has_assign
        flag_val = pd.to_numeric(
            df['flag_exclude_all'] if 'flag_exclude_all' in df.columns
            else pd.Series(1, index=df.index), errors='coerce'
        ).fillna(1).astype(int)
        wilsus_flag0_mask = wilsus_mask & (flag_val == 0)
        result['wilsus'] = int(wilsus_mask.sum())
        result['wilsus_flag0'] = int(wilsus_flag0_mask.sum())
        # Per-area wilsus breakdown (flag0 + all)
        wilsus_by_area = []
        if area_col:
            for area, ag in df.groupby(area_col):
                idx = ag.index
                wilsus_by_area.append({
                    'area':       str(area),
                    'wilsus_all': int(wilsus_mask[idx].sum()),
                    'wilsus_flag0': int(wilsus_flag0_mask[idx].sum()),
                })
            wilsus_by_area.sort(key=lambda x: -x['wilsus_all'])
        result['wilsus_by_area'] = wilsus_by_area

    return result


# ─── FFG special aggregation ──────────────────────────────────────────────────

def aggregate_ffg(df, period, merged_months, formula_version="legacy"):
    """
    FFG SLA = (TTI_ref_comply + TTI_ref_not_comply - FFG_NC) /
               (TTI_ref_comply + TTI_ref_not_comply)

    Denominator = TTI comply + not_comply from 2 months prior.
    Anomaly rows are EXCLUDED from denominator per business rule.

    NC rows = rows in FFG file with flag_exclude_all=0 (and f_ffg='FFG-NOTC' if present).

    `merged_months` must be the merged dict MATCHING formula_version (i.e.
    merged_2025 when formula_version='legacy', merged_2026 when '2026') —
    the caller is responsible for passing the right one, so the TTI
    denominator this FFG result uses is always the TTI result computed with
    the SAME formula (legacy TTI denom for legacy FFG, Fulfill-AO TTI denom
    for 2026 FFG). If the reference period predates the 2026 cutover, only a
    legacy TTI result (or TTI_DENOM_HARDCODE) will ever exist for it, and
    both formula versions correctly fall back to that same number.
    """
    # ── Reference period (N-2 months) ──
    dt = datetime.strptime(period, '%Y%m')
    m, y = dt.month - FFG_TTI_LAG_MONTHS, dt.year
    if m <= 0:
        m += 12
        y -= 1
    ref_period = f"{y}{m:02d}"

    tti_ref = merged_months.get(ref_period, {}).get("TTI")
    if not tti_ref:
        tti_ref = TTI_DENOM_HARDCODE.get(ref_period)
        if tti_ref:
            print(f"   ℹ  FFG {period}: using hardcoded TTI denominator for {ref_period} "
                  f"({tti_ref['comply'] + tti_ref['not_comply']:,} orders)")
        else:
            raise ValueError(
                f"TTI data for reference period {ref_period} not found.\n"
                f"   FFG {period} requires TTI {ref_period}. "
                f"Ensure {ref_period}_TTI.csv is in the folder, or add it to TTI_DENOM_HARDCODE."
            )

    # ── Helpers ──
    def tti_denom(obj):
        """Comply + not_comply only — anomaly excluded from FFG denominator."""
        return obj.get("comply", 0) + obj.get("not_comply", 0)

    def nc_count(mask, col=None, val=None):
        if col and col in df.columns:
            return int((mask & (df[col] == val)).sum())
        return int(mask.sum())

    # ── Filter flag_exclude_all=0 (exclude flagged rows from NC count) ──
    # '2026' formula version: NO exclusions or exceptions of any kind —
    # every row counts, full stop. flag_exclude_all is not consulted at all
    # (this supersedes the earlier WILSUS-only override: WILSUS rows are
    # simply included along with everything else now, no special-casing
    # needed). The 'legacy' formula version is completely unaffected —
    # flag_exclude_all alone still decides eligibility there, exactly as it
    # always has, regardless of calendar period.
    if formula_version == "2026":
        eligible = pd.Series(True, index=df.index)
    else:
        flag = pd.to_numeric(
            df['flag_exclude_all'] if 'flag_exclude_all' in df.columns
            else pd.Series(0, index=df.index),
            errors='coerce'
        ).fillna(1).astype(int)
        eligible = (flag == 0)

    df = df[eligible].copy()

    id_col = 'order_id' if 'order_id' in df.columns else None

    area_col = 'tsel_area'     if 'tsel_area'     in df.columns else None
    reg_col  = 'tsel_regional' if 'tsel_regional' in df.columns else None
    br_col   = 'tsel_branch'   if 'tsel_branch'   in df.columns else None
    sto_col  = 'sto'           if 'sto'           in df.columns else None

    def nc_distinct(sub):
        """Distinct order_id count in a sub-dataframe (= FFG not-comply count)."""
        if id_col and id_col in sub.columns:
            return sub[id_col].nunique()
        return len(sub)

    # ── Overall ──
    nc_ov  = nc_distinct(df)
    den_ov = tti_denom(tti_ref)
    co_ov  = den_ov - nc_ov
    sla_ov = round(co_ov / den_ov * 100, 2) if den_ov else 0.0

    # ── Per area → regional → branch → STO ──
    by_area = []
    for ao in tti_ref.get("by_area", []):
        area  = ao["area"]
        den_a = tti_denom(ao)
        adf   = df[df[area_col] == area] if area_col else df
        nc_a  = nc_distinct(adf)
        co_a  = den_a - nc_a
        sla_a = round(co_a / den_a * 100, 2) if den_a else 0.0

        by_reg = []
        for ro in ao.get("by_regional", []):
            reg   = ro["regional"]
            den_r = tti_denom(ro)
            rdf   = adf[adf[reg_col] == reg] if reg_col else adf
            nc_r  = nc_distinct(rdf)
            co_r  = den_r - nc_r
            sla_r = round(co_r / den_r * 100, 2) if den_r else 0.0

            by_br = []
            for bo in ro.get("by_branch", []):
                br    = bo["branch"]
                den_b = tti_denom(bo)
                bdf   = rdf[rdf[br_col] == br] if br_col else rdf
                nc_b  = nc_distinct(bdf)
                co_b  = den_b - nc_b
                sla_b = round(co_b / den_b * 100, 2) if den_b else 0.0

                by_sto = []
                for so in bo.get("by_sto", []):
                    sto   = so["sto"]
                    den_s = tti_denom(so)
                    sdf   = bdf[bdf[sto_col] == sto] if sto_col else bdf
                    nc_s  = nc_distinct(sdf)
                    co_s  = den_s - nc_s
                    sla_s = round(co_s / den_s * 100, 2) if den_s else 0.0
                    by_sto.append({"sto": sto, "comply": co_s, "not_comply": nc_s,
                                   "anomaly": 0, "sla_pct": sla_s})

                row = {"branch": br, "comply": co_b, "not_comply": nc_b,
                       "anomaly": 0, "sla_pct": sla_b}
                if by_sto: row["by_sto"] = by_sto
                by_br.append(row)

            by_reg.append({"regional": reg, "comply": co_r, "not_comply": nc_r,
                           "anomaly": 0, "sla_pct": sla_r, "by_branch": by_br})

        by_area.append({"area": area, "comply": co_a, "not_comply": nc_a,
                        "anomaly": 0, "sla_pct": sla_a, "by_regional": by_reg})

    return {
        "total":      den_ov,
        "comply":     co_ov,
        "not_comply": nc_ov,
        "anomaly":    0,
        "sla_pct":    sla_ov,
        "target":     resolve_target("FFG", formula_version),
        "by_area":    by_area,
        "_ffg_ref_period": ref_period,   # informational
    }


# ─── GitHub helpers ───────────────────────────────────────────────────────────

def read_existing_kpis(filepath="data/kpis.json"):
    """
    Fetch and parse an existing KPI JSON file from GitHub. Used for BOTH
    data/kpis_2025.json and data/kpis_2026.json — pass the filepath for
    whichever one you need (see KPI_2025_FILE / KPI_2026_FILE below).
    """
    url     = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filepath}"
    headers = {"Authorization": f"token {GITHUB_TOKEN}",
               "Accept": "application/vnd.github.v3+json"}
    r = github_request("GET", url, headers=headers, params={"ref": GITHUB_BRANCH})

    if r.status_code == 404:
        return {"months": {}}
    if r.status_code != 200:
        print(f"   ⚠  GitHub returned HTTP {r.status_code} for {filepath} — will start fresh")
        return {"months": {}}

    info = r.json()
    try:
        content = info.get("content", "").replace("\n", "")
        if content:
            return json.loads(base64.b64decode(content).decode('utf-8'))

        blob_sha = info.get("sha")
        if blob_sha:
            print(f"   ℹ  {filepath} > 1 MB — fetching via Git Blobs API …", end="  ", flush=True)
            blob_url = f"https://api.github.com/repos/{GITHUB_REPO}/git/blobs/{blob_sha}"
            blob_r   = github_request("GET", blob_url, headers=headers)
            if blob_r.status_code == 200:
                bc = blob_r.json().get("content", "").replace("\n", "")
                if bc:
                    print("✓")
                    return json.loads(base64.b64decode(bc).decode('utf-8'))
            print(f"❌ HTTP {blob_r.status_code}")
    except Exception as e:
        print(f"   ⚠  Could not parse existing {filepath}: {e}")

    return {"months": {}}


def push_to_github(content, filepath, message):
    url     = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filepath}"
    headers = {"Authorization": f"token {GITHUB_TOKEN}",
               "Accept": "application/vnd.github.v3+json"}
    sha = None
    r = github_request("GET", url, headers=headers, params={"ref": GITHUB_BRANCH})
    if r.status_code == 200:
        sha = r.json().get("sha")
    payload = {
        "message": message,
        "content": base64.b64encode(content.encode()).decode(),
        "branch":  GITHUB_BRANCH
    }
    if sha:
        payload["sha"] = sha
    r = github_request("PUT", url, headers=headers, json=payload)
    if r.status_code not in (200, 201):
        print(f"\n   ❌  HTTP {r.status_code} for {filepath}: {r.text[:160]}")
        return False
    return True


def period_label(p):
    try: return datetime.strptime(p, '%Y%m').strftime('%B %Y')
    except: return p


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    print("\n🚀  Dashboard Update Script v3")
    print(f"    Folder : {ONEDRIVE_FOLDER}\n")

    if not os.path.isdir(ONEDRIVE_FOLDER):
        print(f"❌  Folder not found: {ONEDRIVE_FOLDER}"); sys.exit(1)
    if GITHUB_TOKEN == "YOUR_GITHUB_TOKEN":
        print("❌  Set GITHUB_TOKEN in the script."); sys.exit(1)

    # ── Connectivity check, before any expensive parsing ─────────────────────
    # Reading these workbooks takes minutes (202601_202606_MO.xlsx alone is
    # 128 MB). Failing at the push step after all that work is miserable, so
    # confirm GitHub is reachable first — it costs a couple of seconds.
    print("🌐  Checking connection to GitHub …", end="  ", flush=True)
    ok, detail = preflight()
    if not ok:
        print("✗")
        explain_network_error(RuntimeError(detail))
        print("    Stopped before processing any data — nothing was wasted.\n")
        sys.exit(2)
    print(f"✓  ({detail})\n")

    # ── Discover all files ────────────────────────────────────────────────────
    all_files = glob(os.path.join(ONEDRIVE_FOLDER, "*.csv")) + \
                glob(os.path.join(ONEDRIVE_FOLDER, "*.xlsx"))
    pattern        = re.compile(r'(\d{6})_([a-zA-Z0-9]+)\.(csv|xlsx)$', re.IGNORECASE)
    # Ranged (multi-month-in-one-file) filenames: START_END_ENTITY.ext, e.g.
    # 202601_202606_MO.xlsx. Checked first so its two period tokens aren't
    # mistaken for the single-period pattern above.
    ranged_pattern = re.compile(r'(\d{6})_(\d{6})_([a-zA-Z0-9]+)\.(csv|xlsx)$', re.IGNORECASE)

    by_period = {}
    ranged_files = []   # (start_period, end_period, entity, filepath)
    ignored_found = []  # (filename, entity) — for the summary print below
    for f in all_files:
        base = os.path.basename(f)
        rm = ranged_pattern.search(base)
        if rm:
            start_p, end_p, ent, _ext = rm.groups()
            ent = ent.upper()
            if ent in IGNORED_ENTITIES:
                ignored_found.append((base, ent))
                continue
            ranged_files.append((start_p, end_p, ent, f))
            continue
        m = pattern.search(base)
        if m:
            ent = m.group(2).upper()
            if ent in IGNORED_ENTITIES:
                ignored_found.append((base, ent))
                continue
            by_period.setdefault(m.group(1), {})[ent] = f

    if ignored_found:
        for base, ent in ignored_found:
            print(f"⏭   Ignoring {base} (entity '{ent}' is in IGNORED_ENTITIES)")
        print()

    # Split ranged files by their per-row status date and merge the resulting
    # per-period dataframes into by_period. filepath=None marks "already
    # loaded" — the load-site below uses the dataframe directly instead of
    # calling load_file() again.
    for start_p, end_p, ent, filepath in ranged_files:
        print(f"📦  Ranged file: {os.path.basename(filepath)} → entity {ent}, periods {start_p}–{end_p}")
        split = split_ranged_file(filepath, ent, start_p, end_p)
        for period, sub_df in sorted(split.items()):
            by_period.setdefault(period, {})[ent] = (None, sub_df)
            print(f"   ↳  {period}: {len(sub_df):,} rows")
        print()

    if not by_period:
        print(f"❌  No files found in: {ONEDRIVE_FOLDER}")
        print("   Files must be named like: 202506_TTI.csv or 202506_FFGGROUP.xlsx")
        sys.exit(1)

    # ── Decide which periods to process ──────────────────────────────────────
    if TARGET_PERIOD:
        # Accepts one period ("202607") or a comma-separated list
        # ("202601,202602"). The list form lets deploy.py handle several
        # changed periods in a single run, so a ranged file like
        # 202601_202606_MO.xlsx is split once instead of once per period.
        wanted  = [p.strip() for p in TARGET_PERIOD.split(",") if p.strip()]
        missing = [p for p in wanted if p not in by_period]
        if missing:
            print(f"❌  Period(s) not found: {', '.join(missing)}. "
                  f"Available: {', '.join(sorted(by_period))}")
            sys.exit(1)
        periods_to_process = {p: by_period[p] for p in wanted}
    else:
        periods_to_process = by_period   # ALL periods

    # Expand FFGGROUP aliases into individual entities per period
    for period in list(periods_to_process.keys()):
        ef = periods_to_process[period]
        for alias in list(ef.keys()):
            if alias in FFGGROUP_ALIASES:
                filepath = ef.pop(alias)
                sub = _expand_ffggroup(filepath)
                # sub = {entity: (filepath, sheet)}
                ef.update(sub)
            else:
                # Wrap plain filepath as (filepath, 0) for consistency
                ef[alias] = (ef[alias], 0) if not isinstance(ef[alias], tuple) else ef[alias]
        periods_to_process[period] = ef

    print(f"📅  Periods  : {', '.join(sorted(periods_to_process))}")
    entities_all = sorted({e for ef in periods_to_process.values() for e in ef})
    print(f"📁  Entities : {', '.join(entities_all)}\n")

    # ── Read existing GitHub data ONCE — separately per file ──────────────────
    existing_2025 = read_existing_kpis(KPI_2025_FILE)
    existing_2026 = read_existing_kpis(KPI_2026_FILE)
    merged_2025 = existing_2025.get("months", {})
    merged_2026 = existing_2026.get("months", {})
    prev_count_2025 = len(merged_2025)
    prev_count_2026 = len(merged_2026)

    # ── Process each period ───────────────────────────────────────────────────
    any_processed_2025 = False
    any_processed_2026 = False
    all_sorted_periods = sorted(periods_to_process.keys())
    last_period        = all_sorted_periods[-1]
    # Periods eligible for CSV upload (most recent N only)
    csv_upload_periods = set(
        all_sorted_periods[-CSV_UPLOAD_MAX_PERIODS:]
        if CSV_UPLOAD_MAX_PERIODS > 0 else []
    )

    VERSION_TAG = {"legacy": "2025", "2026": "2026"}   # for progress messages only

    for period in sorted(periods_to_process.keys()):
        label        = period_label(period)
        entity_files = periods_to_process[period]
        print(f"── {label} ({period})  [{', '.join(sorted(entity_files))}]")

        # Results are split by OUTPUT FILE, not by calendar period — a
        # period >= KPI_2026_CUTOVER_PERIOD can produce entries in BOTH
        # dicts for the same entity (TTI/FFG), one per formula version.
        month_data_2025 = {}
        month_data_2026 = {}

        # Process all entities EXCEPT FFG (FFG needs TTI from 2 months prior)
        for entity in sorted(entity_files):
            if entity == "FFG":
                continue
            filepath, sheet = entity_files[entity]
            print(f"   ⏳  {entity} ...", end="  ", flush=True)
            try:
                if filepath is None:
                    # Ranged file (e.g. MO): sheet holds the already-split,
                    # already-lowercased-columns sub-dataframe for this period.
                    raw_df = sheet
                else:
                    raw_df = load_file(filepath, sheet=sheet)
                prepared = prepare_df(raw_df, entity)

                versions   = entity_formula_versions(entity, period)
                legacy_fdf = None   # kept for CSV export below (always the
                                     # legacy-formula df, matching the KPI
                                     # 2025 page's "download detail" button)
                v2026_fdf  = None   # 2026-formula df, when computed — feeds
                                     # the separate tti2026_* CSV export so
                                     # the KPI 2026 page never downloads
                                     # legacy-formula rows under its own
                                     # "Fulfill AO" download buttons.
                parts = []
                for version in versions:
                    fdf = apply_sla_formula(prepared.copy(), entity, formula_version=version)
                    res = aggregate_entity(fdf, entity, formula_version=version)
                    if version == "legacy":
                        month_data_2025[entity] = res
                        legacy_fdf = fdf
                    else:
                        month_data_2026[entity] = res
                        v2026_fdf = fdf
                    parts.append(f"[{VERSION_TAG[version]}] {res['sla_pct']}%")
                print(f"✓   SLA {'  '.join(parts)}   ({', '.join(v for v in versions)})")

                csv_df = legacy_fdf if legacy_fdf is not None else fdf

                # Push per-row detail files for TTI (raw + processed) —
                # always the legacy-formula df. Only for the most recent
                # CSV_UPLOAD_MAX_PERIODS periods.
                if entity == "TTI" and period in csv_upload_periods:
                    for label_str, fn, builder in [
                        ("raw",       f"data/tti_raw_{period}.csv",    build_tti_raw_csv),
                        ("processed", f"data/tti_detail_{period}.csv", build_tti_detail_csv),
                    ]:
                        try:
                            csv_str = builder(csv_df, period)
                            print(f"   📁  Pushing {fn} ({label_str}, {len(csv_df):,} rows) ...", end="  ", flush=True)
                            if push_to_github(csv_str, fn, f"TTI {label_str}: {period}"):
                                print("✓")
                            else:
                                print("⚠  push failed (non-critical)")
                        except Exception as ex:
                            print(f"⚠  TTI {label_str} export failed: {ex}")
                elif entity == "TTI":
                    print(f"   ⏭   Skipping CSV upload for {period} (not in latest {CSV_UPLOAD_MAX_PERIODS} period(s))")
                # Separate 2026-formula (Fulfill AO) detail export — only when
                # a 2026-version result was actually computed for this period.
                # Distinct filenames (tti2026_*) so the KPI 2026 page's
                # download buttons never serve legacy-formula rows.
                if entity == "TTI" and v2026_fdf is not None and period in csv_upload_periods:
                    for label_str, fn, builder in [
                        ("raw",       f"data/tti2026_raw_{period}.csv",    build_tti_raw_csv),
                        ("processed", f"data/tti2026_detail_{period}.csv", build_tti_detail_csv),
                    ]:
                        try:
                            csv_str = builder(v2026_fdf, period)
                            print(f"   📁  Pushing {fn} ({label_str}, {len(v2026_fdf):,} rows) ...", end="  ", flush=True)
                            if push_to_github(csv_str, fn, f"Fulfill AO {label_str}: {period}"):
                                print("✓")
                            else:
                                print("⚠  push failed (non-critical)")
                        except Exception as ex:
                            print(f"⚠  Fulfill AO {label_str} export failed: {ex}")
                if entity == "MO" and period in csv_upload_periods:
                    for label_str, fn, builder in [
                        ("raw",       f"data/mo_raw_{period}.csv",    build_mo_raw_csv),
                        ("processed", f"data/mo_detail_{period}.csv", build_mo_detail_csv),
                    ]:
                        try:
                            csv_str = builder(csv_df, period)
                            print(f"   📁  Pushing {fn} ({label_str}, {len(csv_df):,} rows) ...", end="  ", flush=True)
                            if push_to_github(csv_str, fn, f"Fulfill MO {label_str}: {period}"):
                                print("✓")
                            else:
                                print("⚠  push failed (non-critical)")
                        except Exception as ex:
                            print(f"⚠  Fulfill MO {label_str} export failed: {ex}")
                elif entity == "MO":
                    print(f"   ⏭   Skipping MO CSV upload for {period} (not in latest {CSV_UPLOAD_MAX_PERIODS} period(s))")
                if entity == "PDA" and period in csv_upload_periods:
                    for label_str, fn, builder in [
                        ("raw",       f"data/pda_raw_{period}.csv",    build_pda_raw_csv),
                        ("processed", f"data/pda_detail_{period}.csv", build_pda_detail_csv),
                    ]:
                        try:
                            csv_str = builder(csv_df, period)
                            print(f"   📁  Pushing {fn} ({label_str}, {len(csv_df):,} rows) ...", end="  ", flush=True)
                            if push_to_github(csv_str, fn, f"Fulfill PDA {label_str}: {period}"):
                                print("✓")
                            else:
                                print("⚠  push failed (non-critical)")
                        except Exception as ex:
                            print(f"⚠  Fulfill PDA {label_str} export failed: {ex}")
                elif entity == "PDA":
                    print(f"   ⏭   Skipping PDA CSV upload for {period} (not in latest {CSV_UPLOAD_MAX_PERIODS} period(s))")
                if entity == "TTRFFG" and period in csv_upload_periods:
                    for label_str, fn, builder in [
                        ("raw",       f"data/ttrffg_raw_{period}.csv",    build_ttrffg_raw_csv),
                        ("processed", f"data/ttrffg_detail_{period}.csv", build_ttrffg_detail_csv),
                    ]:
                        try:
                            csv_str = builder(csv_df, period)
                            print(f"   📁  Pushing {fn} ({label_str}, {len(csv_df):,} rows) ...", end="  ", flush=True)
                            if push_to_github(csv_str, fn, f"TTRFFG {label_str}: {period}"):
                                print("✓")
                            else:
                                print("⚠  push failed (non-critical)")
                        except Exception as ex:
                            print(f"⚠  TTRFFG {label_str} export failed: {ex}")
                elif entity == "TTRFFG":
                    print(f"   ⏭   Skipping TTRFFG CSV upload for {period} (not in latest {CSV_UPLOAD_MAX_PERIODS} period(s))")
                if entity == "PSRE" and period in csv_upload_periods:
                    # Raw = push the original source file as-is (YYYYMM_PSRE.csv)
                    try:
                        src_path = entity_files["PSRE"][0]
                        raw_fn   = f"data/{period}_PSRE.csv"
                        with open(src_path, 'r', encoding='utf-8-sig', errors='replace') as _f:
                            raw_content = _f.read()
                        print(f"   📁  Pushing {raw_fn} (raw original) ...", end="  ", flush=True)
                        if push_to_github(raw_content, raw_fn, f"PSRE raw: {period}"):
                            print("✓")
                        else:
                            print("⚠  push failed (non-critical)")
                    except Exception as ex:
                        print(f"⚠  PSRE raw upload failed: {ex}")
                    # Processed = structured detail columns
                    try:
                        csv_str = build_psre_detail_csv(csv_df, period)
                        det_fn  = f"data/psre_detail_{period}.csv"
                        print(f"   📁  Pushing {det_fn} (processed, {len(csv_df):,} rows) ...", end="  ", flush=True)
                        if push_to_github(csv_str, det_fn, f"PSRE processed: {period}"):
                            print("✓")
                        else:
                            print("⚠  push failed (non-critical)")
                    except Exception as ex:
                        print(f"⚠  PSRE processed export failed: {ex}")
                elif entity == "PSRE":
                    print(f"   ⏭   Skipping PSRE CSV upload for {period} (not in latest {CSV_UPLOAD_MAX_PERIODS} period(s))")
            except Exception as e:
                print(f"❌  {e}")

        # Merge this period's non-FFG results into BOTH merged dicts (each
        # entity's result only lands in the dict(s) matching the formula
        # version(s) it was actually computed with) — done BEFORE FFG so
        # FFG can reference this period's TTI in the same run.
        if month_data_2025:
            merged_2025[period] = {**merged_2025.get(period, {}), **month_data_2025}
            any_processed_2025 = True
        if month_data_2026:
            merged_2026[period] = {**merged_2026.get(period, {}), **month_data_2026}
            any_processed_2026 = True

        # Process FFG (requires TTI from ref_period already in the matching
        # merged dict — a 2026 January/February FFG's ref_period falls in
        # 2025, so it naturally misses merged_2026 and falls through to
        # TTI_DENOM_HARDCODE, exactly as before the file split).
        if "FFG" in entity_files:
            filepath, sheet = entity_files["FFG"]
            print(f"   ⏳  FFG ...", end="  ", flush=True)
            try:
                raw_df   = load_file(filepath, sheet=sheet)
                prepared = prepare_df(raw_df, "FFG")

                versions = entity_formula_versions("FFG", period)
                parts = []
                for version in versions:
                    target_merged = merged_2025 if version == "legacy" else merged_2026
                    try:
                        res = aggregate_ffg(prepared.copy(), period, target_merged, formula_version=version)
                        target_merged.setdefault(period, {})["FFG"] = res
                        if version == "legacy":
                            month_data_2025["FFG"] = res
                            any_processed_2025 = True
                        else:
                            month_data_2026["FFG"] = res
                            any_processed_2026 = True
                        ref = res.get("_ffg_ref_period", "?")
                        parts.append(f"[{VERSION_TAG[version]}] {res['sla_pct']}% (denom: TTI {ref})")
                    except Exception as ve:
                        print(f"\n   ❌  FFG [{VERSION_TAG[version]}] {period}: {ve}", end="  ")
                if parts:
                    print(f"✓   SLA {'  '.join(parts)}")
                else:
                    print("❌  both formula versions failed")

                # Push per-row detail files for FFG (raw + processed) — once
                # per period regardless of how many versions were computed
                # (the eligibility filter is applied inside aggregate_ffg,
                # not here, so `prepared` is the same dedup'd-but-unfiltered
                # df for either version). Only for the most recent
                # CSV_UPLOAD_MAX_PERIODS periods.
                if period in csv_upload_periods:
                    for label_str, fn, builder in [
                        ("raw",       f"data/ffg_raw_{period}.csv",    build_ffg_raw_csv),
                        ("processed", f"data/ffg_detail_{period}.csv", build_ffg_detail_csv),
                    ]:
                        try:
                            csv_str = builder(prepared, period)
                            print(f"   📁  Pushing {fn} ({label_str}, {len(prepared):,} rows) ...", end="  ", flush=True)
                            if push_to_github(csv_str, fn, f"FFG {label_str}: {period}"):
                                print("✓")
                            else:
                                print("⚠  push failed (non-critical)")
                        except Exception as ex:
                            print(f"⚠  FFG {label_str} export failed: {ex}")
                else:
                    print(f"   ⏭   Skipping FFG CSV upload for {period} (not in latest {CSV_UPLOAD_MAX_PERIODS} period(s))")
            except Exception as e:
                print(f"❌  {e}")

        print()

    if not any_processed_2025 and not any_processed_2026:
        print("❌  No data processed."); sys.exit(1)

    last_label = period_label(last_period)

    def _push_kpi_file(merged, prev_count, filepath, year_label):
        if not merged:
            print(f"   ⏭   {filepath}: no {year_label} data to push, skipping")
            return
        periods = sorted(merged.keys())
        file_last = periods[-1]
        output = {
            "last_updated":       file_last,
            "last_updated_label": period_label(file_last),
            "generated_at":       datetime.now().isoformat(timespec='seconds'),
            "months":             merged
        }
        json_str = json.dumps(output, ensure_ascii=False, indent=2)
        print(f"   Months in {filepath} : {len(merged)}  (was {prev_count} on GitHub)")
        print(f"📤  Pushing {filepath} ...", end="  ", flush=True)
        if push_to_github(json_str, filepath, f"{year_label} KPI update: {last_label}"):
            print("✓")
        else:
            print("❌  Push failed. Check token.")
            local = os.path.join(os.path.dirname(__file__), *filepath.split('/'))
            os.makedirs(os.path.dirname(local), exist_ok=True)
            with open(local, 'w') as f: f.write(json_str)
            print(f"   Saved locally: {local}")

    # ── Two separate pushes — only for whichever year(s) this run touched ────
    if any_processed_2025:
        _push_kpi_file(merged_2025, prev_count_2025, KPI_2025_FILE, "2025")
    if any_processed_2026:
        _push_kpi_file(merged_2026, prev_count_2026, KPI_2026_FILE, "2026")

    html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'index.html')
    if os.path.isfile(html_path):
        print(f"📤  Pushing index.html ...", end="  ", flush=True)
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        if push_to_github(html_content, "index.html", f"Dashboard update: {last_label}"):
            print("✓")
        else:
            print("❌  index.html push failed (non-critical)")

    print(f"\n✅  Done! → https://{GITHUB_REPO.split('/')[0]}.github.io/{GITHUB_REPO.split('/')[1]}/\n")


def _run():
    """Wrap main() so a blocked network prints guidance, not a traceback."""
    try:
        main()
    except GitHubUnreachable as e:
        explain_network_error(e)
        sys.exit(2)
    except KeyboardInterrupt:
        print("\n\n⏹   Cancelled. Nothing further was uploaded.\n")
        sys.exit(130)


if __name__ == "__main__":
    _run()
