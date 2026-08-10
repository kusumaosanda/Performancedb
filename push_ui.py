#!/usr/bin/env python3
"""
push_ui.py  —  Push only index.html to GitHub Pages.
Run this after Claude finishes editing to go live instantly.

Usage:
    python3 push_ui.py
"""
import os, sys, base64
from dashboard_env import (get_github_config, github_request,
                           GitHubUnreachable, explain_network_error)

GITHUB_TOKEN, GITHUB_REPO, GITHUB_BRANCH = get_github_config()
HEADERS = {
    "Authorization": f"token {GITHUB_TOKEN}",
    "Accept": "application/vnd.github.v3+json",
}

def push_file(content_str, filepath, message):
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{filepath}"
    r = github_request("GET", url, headers=HEADERS, params={"ref": GITHUB_BRANCH})
    if r.status_code == 401:
        print("\n❌  GitHub rejected the token (HTTP 401).")
        print("    It has expired or been revoked. Generate a new one and put it")
        print("    in .env as GITHUB_TOKEN=…  (scope: repo)")
        sys.exit(1)
    sha = r.json().get("sha") if r.status_code == 200 else None
    payload = {
        "message": message,
        "content": base64.b64encode(content_str.encode("utf-8")).decode(),
        "branch":  GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha
    r = github_request("PUT", url, headers=HEADERS, json=payload)
    if r.status_code not in (200, 201):
        print(f"\n❌  GitHub returned HTTP {r.status_code}")
        print(f"    {r.text[:200]}")
        return False
    return True

here = os.path.dirname(os.path.abspath(__file__))
html_path = os.path.join(here, "index.html")

if not os.path.isfile(html_path):
    print("❌  index.html not found"); sys.exit(1)

with open(html_path, "r", encoding="utf-8") as f:
    html = f.read()

print("📤  Pushing index.html ...", end="  ", flush=True)
try:
    ok = push_file(html, "index.html", "UI update: index.html")
except GitHubUnreachable as e:
    explain_network_error(e)
    sys.exit(2)

if ok:
    print("✓")
    print(f"\n✅  Live → https://{GITHUB_REPO.split('/')[0]}.github.io/{GITHUB_REPO.split('/')[1]}/\n")
else:
    print("\n❌  Push failed — see the error above.")
    sys.exit(1)
