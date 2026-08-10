#!/usr/bin/env python3
"""
deploy.py — one command to publish the dashboard.

Looks at the OneDrive Performancedb folder, works out which periods have new
or changed files since the last successful deploy, runs the pipeline for just
those periods, and pushes the result to GitHub Pages.

Normal use — after dropping new files into OneDrive:

    ./deploy.sh

Other modes:

    ./deploy.sh --check           show what would run, change nothing
    ./deploy.sh 202607            force one period (ignores change detection)
    ./deploy.sh 202601 202602     force several periods
    ./deploy.sh --all             reprocess every period in the folder
    ./deploy.sh --ui              push index.html only (UI edits, no data run)
    ./deploy.sh --source          push the project's own code/docs to the repo
    ./deploy.sh --reset           forget the change-detection state
    ./deploy.sh --mark-deployed   accept the folder as already published
                                  (run this once, now, so the first real
                                  ./deploy.sh doesn't reprocess all 7 periods)
    ./deploy.sh --netcheck        test the connection to GitHub and stop

State lives in .deploy_state.json next to this file. It records the size and
modification time of every source file at the last successful deploy; a period
is "changed" when one of its files is added, removed, resized, or re-saved.
"""
import base64
import json
import os
import re
import subprocess
import sys
from datetime import datetime

from dashboard_env import (get_github_config, ENV_PATH, github_request,
                           GitHubUnreachable, explain_network_error, preflight)

HERE = os.path.dirname(os.path.abspath(__file__))
STATE_PATH = os.path.join(HERE, ".deploy_state.json")
ONEDRIVE_FOLDER = os.path.expanduser(
    "~/Library/CloudStorage/OneDrive-PT.TelekomunikasiIndonesia/Performancedb"
)

SINGLE_RE = re.compile(r"^(\d{6})_([a-zA-Z0-9]+)\.(csv|xlsx)$", re.IGNORECASE)
RANGED_RE = re.compile(r"^(\d{6})_(\d{6})_([a-zA-Z0-9]+)\.(csv|xlsx)$", re.IGNORECASE)

# Backstop against publishing a credential. Matches the *shape* of a real
# GitHub token, not the bare "ghp_" prefix — an earlier version searched for
# the prefix alone and so blocked this very file (which contains the pattern in
# its own guard) and CLAUDE.md (which documents it). A classic PAT is a 4-char
# prefix plus 36 alphanumerics; fine-grained ones start with github_pat_.
SECRET_RE = re.compile(
    rb"gh[pousr]_[A-Za-z0-9]{36,}"      # ghp_/gho_/ghu_/ghs_/ghr_ classic PATs
    rb"|github_pat_[A-Za-z0-9_]{30,}"   # fine-grained PATs
)

# Project files pushed by --source. Deliberately excludes .env, .deploy_state.json,
# __pycache__ and data/ (the pipeline pushes data/ itself).
SOURCE_FILES = [
    "index.html",
    "Notulen-Rapat-Generator.html",
    "update_dashboard.py",
    "push_ui.py",
    "remove_period.py",
    "diagnose_mo.py",
    "deploy.py",
    "deploy.sh",
    "dashboard_env.py",
    "sto_hierarchy_map.json",
    "CLAUDE.md",
    "MY_DASHBOARD_INFO.md",
    "HOW-TO-DEPLOY.md",
    ".gitignore",
    ".env.example",
]


# ── helpers ───────────────────────────────────────────────────────────────────

def months_between(start, end):
    """['202601', ..., '202606'] for a ranged filename's two period tokens."""
    out, y, m = [], int(start[:4]), int(start[4:])
    ey, em = int(end[:4]), int(end[4:])
    while (y, m) <= (ey, em) and len(out) < 240:
        out.append(f"{y}{m:02d}")
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


def scan_folder():
    """Map period -> {filename: "size:mtime"} for everything in OneDrive."""
    if not os.path.isdir(ONEDRIVE_FOLDER):
        print(f"❌  Data folder not found: {ONEDRIVE_FOLDER}")
        sys.exit(1)

    periods = {}
    for name in sorted(os.listdir(ONEDRIVE_FOLDER)):
        if not name.lower().endswith((".csv", ".xlsx")) or name.startswith((".", "~$")):
            continue
        path = os.path.join(ONEDRIVE_FOLDER, name)
        try:
            st = os.stat(path)
        except OSError:
            continue
        stamp = f"{st.st_size}:{int(st.st_mtime)}"

        rm = RANGED_RE.match(name)
        if rm:
            targets = months_between(rm.group(1), rm.group(2))
        else:
            sm = SINGLE_RE.match(name)
            if not sm:
                continue
            targets = [sm.group(1)]

        for p in targets:
            periods.setdefault(p, {})[name] = stamp
    return periods


def load_state():
    if not os.path.isfile(STATE_PATH):
        return {}
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f).get("periods", {})
    except (ValueError, OSError):
        print("⚠   .deploy_state.json unreadable — treating every period as new.")
        return {}


def save_state(periods, deployed):
    """Record fingerprints for the periods that just deployed successfully."""
    state = load_state()
    for p in deployed:
        state[p] = periods.get(p, {})
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {"updated_at": datetime.now().isoformat(timespec="seconds"),
             "periods": state},
            f, indent=2,
        )


def changed_periods(periods, state):
    """Periods whose file set or file fingerprints differ from last deploy."""
    return sorted(p for p, files in periods.items() if state.get(p) != files)


def describe_change(period, files, state):
    old = state.get(period)
    if old is None:
        return "new period"
    added = sorted(set(files) - set(old))
    removed = sorted(set(old) - set(files))
    edited = sorted(f for f in files if f in old and files[f] != old[f])
    bits = []
    if added:
        bits.append("added " + ", ".join(added))
    if edited:
        bits.append("updated " + ", ".join(edited))
    if removed:
        bits.append("removed " + ", ".join(removed))
    return "; ".join(bits) or "unchanged"


# ── actions ───────────────────────────────────────────────────────────────────

def run_pipeline(period_list):
    """Run update_dashboard.py once for the given periods."""
    env = dict(os.environ)
    env["TARGET_PERIOD"] = ",".join(period_list)
    print(f"\n▶   Running pipeline for: {', '.join(period_list)}\n" + "─" * 70)
    r = subprocess.run([sys.executable, os.path.join(HERE, "update_dashboard.py")],
                       cwd=HERE, env=env)
    return r.returncode == 0


def run_ui_push():
    print("\n▶   Pushing index.html only\n" + "─" * 70)
    r = subprocess.run([sys.executable, os.path.join(HERE, "push_ui.py")], cwd=HERE)
    return r.returncode == 0


def push_source():
    """Upload the project's own code and docs to the repo via the Contents API."""
    token, repo, branch = get_github_config()
    headers = {"Authorization": f"token {token}",
               "Accept": "application/vnd.github.v3+json"}

    print("\n▶   Pushing project source to GitHub\n" + "─" * 70)
    ok = failed = skipped = 0
    for rel in SOURCE_FILES:
        path = os.path.join(HERE, rel)
        if not os.path.isfile(path):
            print(f"   ⏭   {rel} (not present)")
            skipped += 1
            continue
        with open(path, "rb") as f:
            blob = f.read()
        leak = SECRET_RE.search(blob)
        if leak:
            print(f"   ⛔  {rel} — contains what looks like a live token "
                  f"({leak.group(0)[:8].decode()}…), NOT pushed")
            failed += 1
            continue

        url = f"https://api.github.com/repos/{repo}/contents/{rel}"
        r = github_request("GET", url, headers=headers, params={"ref": branch})
        sha = r.json().get("sha") if r.status_code == 200 else None
        payload = {"message": f"Update {rel}",
                   "content": base64.b64encode(blob).decode(),
                   "branch": branch}
        if sha:
            payload["sha"] = sha
        pr = github_request("PUT", url, headers=headers, json=payload)
        if pr.status_code in (200, 201):
            print(f"   ✓   {rel}")
            ok += 1
        else:
            print(f"   ❌  {rel} — HTTP {pr.status_code}")
            failed += 1

    print(f"\n   {ok} pushed, {skipped} skipped, {failed} failed")
    return failed == 0


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    args = sys.argv[1:]
    flags = {a for a in args if a.startswith("-")}
    explicit = [a for a in args if re.fullmatch(r"\d{6}", a)]
    unknown = [a for a in args if a.startswith("-") and a not in {
        "--check", "--dry-run", "--all", "--ui", "--source", "--reset",
        "--mark-deployed", "--netcheck", "-h", "--help"}]

    if unknown or "-h" in flags or "--help" in flags:
        if unknown:
            print(f"❌  Unknown option(s): {', '.join(unknown)}\n")
        print(__doc__)
        sys.exit(1 if unknown else 0)

    print("\n" + "═" * 70)
    print("  DASHBOARD DEPLOY")
    print(f"  Data   : {ONEDRIVE_FOLDER}")
    print(f"  Secrets: {ENV_PATH}")
    print("═" * 70)

    if "--reset" in flags:
        if os.path.isfile(STATE_PATH):
            os.remove(STATE_PATH)
            print("\n🧹  Change-detection state cleared — next run reprocesses everything.\n")
        else:
            print("\n🧹  No state file to clear.\n")
        return

    if "--netcheck" in flags:
        print("\n🌐  Testing connection to api.github.com …", end="  ", flush=True)
        ok, detail = preflight()
        if ok:
            print(f"✓  reachable ({detail})\n")
            return
        print("✗")
        explain_network_error(RuntimeError(detail))
        sys.exit(2)

    if "--mark-deployed" in flags:
        periods = scan_folder()
        save_state(periods, sorted(periods))
        print(f"\n📌  Baselined {len(periods)} period(s): {', '.join(sorted(periods))}")
        print("    These are now treated as already published. From here on,")
        print("    ./deploy.sh only processes files you actually change.\n")
        return

    if "--ui" in flags:
        sys.exit(0 if run_ui_push() else 1)

    if "--source" in flags and not explicit and "--all" not in flags:
        sys.exit(0 if push_source() else 1)

    periods = scan_folder()
    state = load_state()

    if "--all" in flags:
        todo = sorted(periods)
        reason = "--all: every period in the folder"
    elif explicit:
        missing = [p for p in explicit if p not in periods]
        if missing:
            print(f"\n❌  No files for period(s): {', '.join(missing)}")
            print(f"    Available: {', '.join(sorted(periods))}\n")
            sys.exit(1)
        todo = sorted(explicit)
        reason = "requested on the command line"
    else:
        todo = changed_periods(periods, state)
        reason = "changed since last deploy"

    print(f"\n📂  {len(periods)} period(s) in folder: {', '.join(sorted(periods))}")

    if not todo:
        print("\n✅  Nothing changed since the last deploy — nothing to do.")
        print("    Drop new files in OneDrive and re-run, or force one:")
        print("      ./deploy.sh 202607        ./deploy.sh --all\n")
        return

    print(f"\n🔎  {len(todo)} period(s) to process ({reason}):")
    for p in todo:
        print(f"      {p}  —  {describe_change(p, periods.get(p, {}), state)}")

    if "--check" in flags or "--dry-run" in flags:
        print("\n(--check) Nothing was run or pushed.\n")
        return

    if not run_pipeline(todo):
        print("\n❌  Pipeline failed — state not updated, nothing marked as deployed.")
        print("    Fix the error above and re-run ./deploy.sh\n")
        sys.exit(1)

    save_state(periods, todo)
    print("\n" + "─" * 70)
    print(f"💾  Marked as deployed: {', '.join(todo)}")

    if "--source" in flags and not push_source():
        sys.exit(1)

    _, repo, _ = get_github_config()
    owner, name = repo.split("/")
    print(f"\n✅  Live → https://{owner}.github.io/{name}/")
    print("    If the page still looks old, that's GitHub Pages/browser cache —")
    print("    wait a minute, then hard-refresh with Cmd+Shift+R.\n")


if __name__ == "__main__":
    try:
        main()
    except GitHubUnreachable as e:
        explain_network_error(e)
        sys.exit(2)
    except KeyboardInterrupt:
        print("\n\n⏹   Cancelled.\n")
        sys.exit(130)
