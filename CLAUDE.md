# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

A Telkom Indonesia / Telkomsel performance-management dashboard: a single static `index.html` (Chart.js + xlsx.js, no build step) deployed to GitHub Pages at `https://kusumaosanda.github.io/Performancedb/`, backed by a Python pipeline (`update_dashboard.py`) that ingests monthly raw KPI exports and republishes them as JSON the dashboard fetches client-side.

It also contains one unrelated standalone tool: `Notulen-Rapat-Generator.html`, a single-file Minutes-of-Meeting generator (Web Speech API recording, transcript editing with a self-learning typo-correction dictionary, auto-generated action plan, Word/PDF export styled to match Telkom/Telkomsel MoM templates). It has no dependency on the dashboard pipeline — treat it as a separate app that happens to live in the same folder.

## Commands

**The one command** — `deploy.sh` is the normal entry point. It fingerprints every file in the OneDrive folder (size + mtime), compares against `.deploy_state.json`, and runs the pipeline for **only the periods whose files changed**, then pushes:
```bash
./deploy.sh
```
| Flag | Effect |
|------|--------|
| `--check` | Dry run — print what would be processed, touch nothing |
| `202607` (positional) | Force specific period(s), bypassing change detection |
| `--all` | Reprocess every period in the folder |
| `--ui` | `index.html` only (wraps `push_ui.py`) |
| `--source` | Push the project's own code/docs to the repo |
| `--mark-deployed` | Baseline the current folder as already published |
| `--reset` | Delete the state file |

Change detection expands ranged filenames: `202601_202606_MO.xlsx` fingerprints against all six periods it covers, so editing it marks all six dirty. Multiple dirty periods are passed to a **single** pipeline run (`TARGET_PERIOD="202601,202602"`) rather than one run each — re-splitting that 128 MB workbook per period would be brutal.

**Direct pipeline invocation** (what `deploy.sh` calls; `TARGET_PERIOD` is read from the environment and accepts a comma-separated list — the old `sed -i ''` dance is gone):
```bash
TARGET_PERIOD="202607" python3 update_dashboard.py
```
Missing-dependency errors print the exact `pip3 install ... --break-system-packages` command to run.

`push.sh` still exists and still works (it wraps `push_ui.py`); `./deploy.sh --ui` is the same thing with the token loaded from `.env`.

**Where raw input files go** (not part of the repo): `~/Library/CloudStorage/OneDrive-PT.TelekomunikasiIndonesia/Performancedb/`, named `YYYYMM_entity.csv` or `.xlsx` (e.g. `202607_TTI.csv`, `202607_FFGGROUP.xlsx`).

There is no test suite, linter, or build process in this repo — `index.html` is opened directly / viewed live, and Python scripts are run ad hoc.

## Architecture

**Dual KPI formula split (the thing most likely to trip you up):** TTI and FFG each have two independent, deliberately-never-shared formulas:
- **legacy** ("KPI 2025" menu) — the original formula, applies to every period.
- **2026** ("KPI 2026" menu) — a new formula, applies from `KPI_2026_CUTOVER_PERIOD = "202601"` onward.

For any period ≥ the cutover, the *same* source file (e.g. `202601_TTI.csv`) is run through **both** formulas independently, producing two separate results written to two separate files: `data/kpis_2025.json` and `data/kpis_2026.json`. MO/PDA only ever compute the 2026 formula (2026-only, no legacy version exists); PSRE/TTRFFG/FFGGAUL only ever compute the legacy formula (no 2026 tab exists for them). This routing logic lives in `entity_formula_versions(entity, period)` in `update_dashboard.py` — check that function first whenever KPI numbers look like they landed on the wrong menu.

**Data flow:**
1. Monthly raw exports dropped in the OneDrive folder above, one file per entity (or a multi-sheet `FFGGROUP` file covering FFG + TTRFFG + FFGGAUL in one workbook).
2. `update_dashboard.py` loads them, dedupes (`order_id` for TTI/PSRE, `trouble_no` for FFG/TTRFFG/FFGGAUL — one ticket = one row), applies the per-entity formula(s) (`_tti_formula`, `_fulfill_ao_formula`, `_fulfill_mo_formula`, `_ffggaul_formula`, etc.), and for Fulfill MO additionally joins against `sto_hierarchy_map.json` (`_apply_sto_hierarchy`) to build an area/regional/branch breakdown — rows with an unmapped STO code still count toward totals but are dropped from the area breakdown.
3. FFG's denominator for period `YYYYMM` uses TTI comply+not_comply counts from period `YYYYMM-2` (a 2-month lag), with anomaly rows excluded.
4. Results are written to `data/kpis_2025.json` / `data/kpis_2026.json`, plus `data/home.json` and `data/projects.json` (separate, simpler datasets). Raw/processed detail CSVs are also pushed, capped to the `CSV_UPLOAD_MAX_PERIODS` most recent periods (default 1) to avoid re-uploading history every run.
5. Everything is pushed directly to the `kusumaosanda/Performancedb` GitHub repo via the API (`push_to_github()` in `update_dashboard.py`, or the smaller `push_file()` in `push_ui.py`) — not through a local git workflow.
6. `index.html` fetches all four JSON files client-side on load (see the `DATA` object around line 1634) and renders KPI views, a project tracker, and a calendar with no server involved.

**Other scripts** (not part of the regular pipeline, run standalone when needed):
- `diagnose_mo.py` — read-only diagnostic for figuring out why a specific MO Excel file is dropping rows; hardcodes a specific filename, meant to be edited per investigation.
- `remove_period.py YYYYMM` — deletes one period's entry from a KPI JSON on GitHub. It targets the older, unsplit `data/kpis.json`, which predates the 2025/2026 formula split — check whether it needs updating to target `kpis_2025.json`/`kpis_2026.json` before relying on it.

## Known gotchas (learned the hard way — read before touching formulas)

**TTR FFG SLA formula is intentionally "wrong" relative to at least one reference table — do not "fix" it without asking first.** `_ttrffg_formula()` in `update_dashboard.py` filters on `flag_exclude_all == 0`. In August 2026 this was found to undercount Comply/Not Comply against a user-supplied manual reconciliation for period 202607 (the fix — dropping the `flag_exclude_all` condition — was verified bit-for-bit against real data: 644 Comply / 86 Not Comply / 88.22% SLA vs. the current 315/26/92.38%). The user was told about the discrepancy and explicitly chose to keep the original `flag_exclude_all`-gated version anyway ("revert to SLA 92,38%"), so the function's docstring carries a NOTE documenting this. If TTR FFG numbers look wrong against a reference table again, check that docstring first and confirm with the user before changing the filter — don't assume it's a bug to silently fix.

**The office network blocks `api.github.com` — deploys only work off it.** On the user's Telkom corporate network the TLS handshake to GitHub times out (`[Errno 60] Operation timed out`), so every push fails before authentication is even attempted. This is not a token problem. `./deploy.sh --netcheck` diagnoses it in ~15 seconds, and `update_dashboard.py` now runs the same probe *before* parsing any workbooks so a blocked network fails fast instead of after minutes of work. If the network needs a proxy, set `HTTPS_PROXY=` in `.env` — every script picks it up automatically.

**Check the live file before diagnosing "it didn't update" — and check it *after* the user's push, not before.** Compare byte length and SHA-256 of the local `index.html` against both `raw.githubusercontent.com` (the repo) and `kusumaosanda.github.io` (the Pages CDN, which is a separate cache layer). A stale check run moments before the user's push looks identical to a failed push; re-check before concluding anything.

**The dashboard's "🔄 Clear Cache" button does not reload `index.html`.** `clearDashboardCache()` only re-fetches `kpis_2025.json` / `kpis_2026.json`. For any UI change the user needs a hard refresh (Cmd+Shift+R) — clicking that button and seeing no change is not evidence the push failed.

**"The dashboard isn't updating" is almost always GitHub Pages / browser cache, not a code bug.** This has happened twice: the user runs `update_dashboard.py`, it completes and pushes successfully, but the live site looks unchanged for a few minutes. Verify by fetching the live JSON/HTML with a cache-busting query param and `{cache:'no-store'}` (or just check the file on `raw.githubusercontent.com`) before assuming the push failed or the code is broken.

**Never widen `sto_hierarchy_map.json` coverage without the user's explicit sign-off on the mapping source.** MO's raw export has no `org_1/2/3` columns — only an `sto` column — so its area/regional/branch breakdown is derived by joining against `sto_hierarchy_map.json` (1,221 STO → hierarchy entries, majority-voted from all six `*_tti.csv` files, per explicit user instruction to mirror TTI's own mapping). Rows whose STO isn't in the map still count toward MO's totals but silently drop out of the area breakdown — if a user reports "MO area breakdown missing entries," check for unmapped STOs before assuming another bug.

## Credentials

`GITHUB_TOKEN` lives in `.env` next to the scripts (mode 600, listed in `.gitignore`, excluded from `deploy.sh --source`). All three scripts get it via `dashboard_env.get_github_config()`; a real environment variable overrides the file. Do not print, log, or echo this value back in full.

It was previously hardcoded in `update_dashboard.py`, `push_ui.py`, `remove_period.py` and `MY_DASHBOARD_INFO.md`. That mattered once the project's own source started being pushed to the same (Pages-serving) `Performancedb` repo. **The old token was exposed on disk for a long time and should be treated as compromised until rotated.** `push_source()` also refuses to upload any file still containing `ghp_` as a backstop.

**Do not put the project under git without dealing with `/Users/970129/.git` first** — an empty repo is initialised at the *home directory*, so the project folder is nested inside it and a `git add .` would stage `.bash_history`, `.mysql_history`, `.claude.json` and the rest of the home dir. The deploy path is pure GitHub REST API and never invokes git, which sidesteps this entirely.

## Existing docs in this repo

- `MY_DASHBOARD_INFO.md` — quick-reference for the monthly update workflow (token location, file naming, step-by-step run instructions, troubleshooting table, KPI targets). Kept in sync with the `kpis_2025.json` / `kpis_2026.json` split and the entity/file-naming list (2026-08).
- `HOW-TO-DEPLOY.md` — describes an older Google-Sheets-based architecture (`GOOGLE_SHEET_ID` in `index.html`) that no longer matches the current code. Treat `MY_DASHBOARD_INFO.md` and this file's Architecture section as the source of truth instead.
