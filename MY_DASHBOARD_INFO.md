# Dashboard Quick Reference

## GitHub
- **Repo:** kusumaosanda/Performancedb
- **Dashboard URL:** https://kusumaosanda.github.io/Performancedb/
- **Token location:** `.env` in the project folder (never committed, never pushed):
  ```
  GITHUB_TOKEN=your_token_here
  ```
  Create it once with `cp .env.example .env`, then paste your token in.
  The token used to be hardcoded in `update_dashboard.py` / `push_ui.py` /
  `remove_period.py` — it isn't any more, because the project's own source is
  now pushed to the same repo.
- **Get/renew token:** github.com → Settings → Developer settings → Personal access tokens → Tokens (classic)
  - Required scope: `repo`

---

## Folder Paths

| What | Path |
|------|------|
| Dashboard project files | `/Users/970129/Claude/Projects/Dashboard Project/` |
| Update script | `/Users/970129/Claude/Projects/Dashboard Project/update_dashboard.py` |
| KPI data file (KPI 2025 menu / legacy formula) | `/Users/970129/Claude/Projects/Dashboard Project/data/kpis_2025.json` |
| KPI data file (KPI 2026 menu / new formula) | `/Users/970129/Claude/Projects/Dashboard Project/data/kpis_2026.json` |
| MO area/regional/branch lookup (required — must stay next to update_dashboard.py) | `/Users/970129/Claude/Projects/Dashboard Project/sto_hierarchy_map.json` |
| OneDrive data folder | `~/Library/CloudStorage/OneDrive-PT.TelekomunikasiIndonesia/Performancedb/` |

> Note: `data/kpis.json` (singular) is an old, superseded file. Since the KPI 2025/2026 formula split, results are written to **both** `kpis_2025.json` and `kpis_2026.json` in the same run.

---

## How to Update Dashboard (Monthly)

### Step 1 — Place files in OneDrive
Put renamed data files here:
```
~/Library/CloudStorage/OneDrive-PT.TelekomunikasiIndonesia/Performancedb/
```
File naming format:
```
YYYYMM_tti.csv         (or .xlsx)   — legacy (KPI 2025) AND new (KPI 2026) formula, both computed
YYYYMM_ffggroup.csv    (or .xlsx)   — multi-sheet file covering FFG + TTRFFG + FFGGAUL, legacy formula only
YYYYMM_psre.csv                     — legacy formula only
YYYYMM_mo.csv          (or .xlsx)   — KPI 2026 only (Fulfill MO); needs sto_hierarchy_map.json for area breakdown
YYYYMM_pda.csv         (or .xlsx)   — KPI 2026 only (Fulfill PDA)
```
Lowercase or uppercase — both work.

---

### Step 2 — Open Terminal
Press **Cmd + Space**, type **Terminal**, press Enter.

---

### Step 3 — Run one command

```bash
cd "/Users/970129/Claude/Projects/Dashboard Project" && ./deploy.sh
```

That's it. `deploy.sh` compares the OneDrive folder against `.deploy_state.json`
and processes **only the periods whose files you actually changed**, then pushes.
If nothing changed it says so and stops without touching GitHub.

Other modes:

| Command | What it does |
|---------|--------------|
| `./deploy.sh` | Process changed periods, then push |
| `./deploy.sh --check` | Show what *would* run — changes nothing |
| `./deploy.sh 202607` | Force one period, ignoring change detection |
| `./deploy.sh 202601 202602` | Force several periods (one pipeline run) |
| `./deploy.sh --all` | Reprocess every period in the folder |
| `./deploy.sh --ui` | Push `index.html` only (UI edits, no data run) |
| `./deploy.sh --source` | Push this project's own code/docs to the repo |
| `./deploy.sh --mark-deployed` | Accept the folder as already published (baseline) |
| `./deploy.sh --reset` | Forget the state — next run reprocesses everything |

> The old `sed -i ''` juggling of `TARGET_PERIOD` is gone. `update_dashboard.py`
> now reads `TARGET_PERIOD` from the environment and accepts a comma-separated
> list, so several months are handled in a single run — which matters because
> `202601_202606_MO.xlsx` is 128 MB and gets re-split on every run.

---

### Step 4 — Wait for completion
The script prints progress. When it says **✅ Done**, dashboard is live at the URL above.

---

### Troubleshooting

| Problem | Fix |
|---------|-----|
| Script hangs at "Pushing to GitHub…" | Wait 2–3 min. If still stuck, close Terminal, check GitHub repo, open new Terminal and re-run |
| `No files found` error | Check file names match `YYYYMM_entity.csv` format |
| Missing package message | Run the `pip3 install …` command shown, then re-run script |
| Token expired / 401 error | Get a new token: github.com → Settings → Developer settings → Personal access tokens → Tokens (classic). Paste it into `.env` as `GITHUB_TOKEN=…` — nothing else needs editing |
| `❌ No GitHub token found` | `.env` is missing. Run `cp .env.example .env`, paste your token, then `chmod 600 .env` |
| `./deploy.sh` says "Nothing changed" but you did update a file | OneDrive may not have finished syncing, or the file kept its size and timestamp. Force it: `./deploy.sh 202607` |
| Deploy processed 7 periods when you only changed one | The baseline was never set. Run `./deploy.sh --mark-deployed` once, then future runs are incremental |
| Breakdown table shows "no data" | Normal on first run — re-run script after the latest index.html update to generate breakdown data |
| Ran the script successfully but the live dashboard still looks old | Almost always browser/GitHub Pages caching, not a failed push — hard-refresh (Cmd+Shift+R) or wait a few minutes before assuming something broke |
| MO area breakdown missing some entries | Some STO codes in the raw file aren't in `sto_hierarchy_map.json` — those rows still count in the MO totals but drop out of the area breakdown. See CLAUDE.md for how that map was built |
| TTR FFG SLA % doesn't match a manual reconciliation table | This is a known, intentional discrepancy — see the NOTE in `_ttrffg_formula()`'s docstring and the "Known gotchas" section of CLAUDE.md before changing the formula |

---

## File Naming Rules
- Pattern: `YYYYMM_entity` — e.g. `202607_tti.xlsx`
- Lowercase or uppercase entity name both work
- CSV and XLSX both supported

## KPI Targets

**KPI 2025 menu (legacy formula, `TARGETS` in `update_dashboard.py`):**

| KPI | Target |
|-----|--------|
| TTI | 93.31% |
| FFG | 98.29% *(marked "confirm with team" in code — double-check before trusting)* |
| TTR FFG | 80.81% |
| PS-RE | 75.17% |
| FFG GAUL | 91.76% |
| (any other entity) | 95.00% default |

**KPI 2026 menu (new formula, `TARGETS_2026` in `update_dashboard.py`):**

| KPI | Target |
|-----|--------|
| Fulfill AO (source file `TTI`) | 95.50% |
| Fulfill MO | 95.00% |
| Fulfill PDA | 100.00% |
| FFG (2026 WILSUS rule) | 98.50% |

These live in code, not in this doc — if a target changes, update `TARGETS` / `TARGETS_2026` in `update_dashboard.py` and this table together.
