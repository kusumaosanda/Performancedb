#!/usr/bin/env python3
"""
Remove a period from kpis.json on GitHub.

Usage:
    python3 remove_period.py 202506
"""

import sys, json, base64
import requests
from dashboard_env import get_github_config

GITHUB_TOKEN, GITHUB_REPO, GITHUB_BRANCH = get_github_config()

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 remove_period.py YYYYMM")
        print("Example: python3 remove_period.py 202506")
        sys.exit(1)

    period = sys.argv[1].strip()

    url     = f"https://api.github.com/repos/{GITHUB_REPO}/contents/data/kpis.json"
    headers = {"Authorization": f"token {GITHUB_TOKEN}",
               "Accept": "application/vnd.github.v3+json"}

    print(f"📥  Fetching kpis.json from GitHub…")
    r = requests.get(url, headers=headers, params={"ref": GITHUB_BRANCH})
    if r.status_code != 200:
        print(f"❌  Could not fetch file: HTTP {r.status_code}")
        sys.exit(1)

    info = r.json()
    data = json.loads(base64.b64decode(info["content"]).decode("utf-8"))
    sha  = info["sha"]

    months = data.get("months", {})
    if period not in months:
        print(f"⚠   Period '{period}' not found. Available: {', '.join(sorted(months.keys()))}")
        sys.exit(0)

    del months[period]
    # Update last_updated to the latest remaining month
    remaining = sorted(months.keys())
    if remaining:
        latest = remaining[-1]
        try:
            from datetime import datetime
            dt = datetime.strptime(latest, "%Y%m")
            data["last_updated"] = latest
            data["last_updated_label"] = dt.strftime("%B %Y")
        except Exception:
            pass

    print(f"🗑   Removing period '{period}'…")
    content  = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode()).decode()
    payload  = {"message": f"Remove period {period}", "content": content,
                "sha": sha, "branch": GITHUB_BRANCH}
    put_r = requests.put(url, headers=headers, json=payload)

    if put_r.status_code in (200, 201):
        print(f"✅  Done — '{period}' removed. Remaining months: {', '.join(remaining)}")
    else:
        print(f"❌  Push failed: HTTP {put_r.status_code}")
        print(put_r.text[:300])

if __name__ == "__main__":
    main()
