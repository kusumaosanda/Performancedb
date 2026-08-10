#!/usr/bin/env python3
"""
dashboard_env.py — one place that knows where the GitHub credentials live.

The PAT used to be hardcoded in update_dashboard.py / push_ui.py /
remove_period.py. Since the source of this project is now pushed to the same
public Performancedb repo, the token had to move out of the tracked files.

It now lives in `.env` next to this file (mode 600, listed in .gitignore).
Real environment variables win over `.env`, so you can also do:

    GITHUB_TOKEN=ghp_xxx python3 update_dashboard.py
"""
import os
import sys
import time

_HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(_HERE, ".env")

DEFAULT_REPO = "kusumaosanda/Performancedb"
DEFAULT_BRANCH = "main"


def _load_env_file(path=ENV_PATH):
    """Read simple KEY=VALUE lines from .env without overriding real env vars."""
    if not os.path.isfile(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


def get_github_config(required=True):
    """Return (token, repo, branch). Exits with a readable message if unset."""
    _load_env_file()
    token = os.environ.get("GITHUB_TOKEN", "").strip()
    repo = os.environ.get("GITHUB_REPO", DEFAULT_REPO).strip()
    branch = os.environ.get("GITHUB_BRANCH", DEFAULT_BRANCH).strip()

    if required and (not token or token == "YOUR_GITHUB_TOKEN"):
        print("❌  No GitHub token found.")
        print(f"    Expected GITHUB_TOKEN in: {ENV_PATH}")
        print("    Create it with:  cp .env.example .env   then paste your token.")
        print("    New token: github.com → Settings → Developer settings →")
        print("               Personal access tokens (classic) → scope: repo")
        sys.exit(1)

    return token, repo, branch


def get_target_period():
    """TARGET_PERIOD from the environment. '' means 'process every period'.

    deploy.py sets this per run so nothing has to sed-edit update_dashboard.py.
    """
    _load_env_file()
    return os.environ.get("TARGET_PERIOD", "").strip()


# ══════════════════════════════════════════════════════════════════════════════
#  NETWORK
#
#  The scripts originally called requests.get/put with no timeout at all, so a
#  blocked or proxied network made them hang indefinitely and then dump a raw
#  traceback. Everything now goes through github_request(): bounded timeouts,
#  retries with backoff, and a readable explanation when GitHub is unreachable.
# ══════════════════════════════════════════════════════════════════════════════

CONNECT_TIMEOUT = 15     # seconds to establish TCP+TLS — fails fast behind a firewall
READ_TIMEOUT    = 180    # seconds to wait for a response body (big CSV pushes)
MAX_ATTEMPTS    = 3
RETRY_BACKOFF   = 4      # seconds, multiplied by attempt number


class GitHubUnreachable(RuntimeError):
    """Raised when GitHub could not be reached after every retry."""


def _proxy_note():
    _load_env_file()
    proxies = {k: os.environ[k] for k in
               ("HTTPS_PROXY", "https_proxy", "HTTP_PROXY", "http_proxy")
               if os.environ.get(k)}
    if proxies:
        name, val = next(iter(proxies.items()))
        return f"    A proxy IS configured ({name}={val}) — check it is correct and running."
    return ("    No proxy is configured. If this machine needs one to reach the\n"
            "    internet, add it to .env:  HTTPS_PROXY=http://user:pass@host:port")


def explain_network_error(err):
    """Print actionable guidance instead of a 60-line traceback."""
    print("\n" + "─" * 70)
    print("❌  Could not reach api.github.com.")
    print(f"    ({type(err).__name__}: {str(err)[:140]})")
    print("─" * 70)
    print("\n    This is a connectivity problem, not a problem with your data.")
    print("    Nothing was uploaded, and nothing local was changed.\n")
    print("    Check, in this order:")
    print("      1. Test the connection directly:")
    print("           curl -sS -m 15 -o /dev/null -w 'HTTP %{http_code}\\n' \\")
    print("             https://api.github.com/rate_limit")
    print("      2. If that also times out, GitHub is blocked on this network.")
    print("         Try a different network (personal hotspot) to confirm.")
    print("      3. Corporate network? You likely need a proxy.")
    print(_proxy_note())
    print("      4. On VPN? Try disconnecting — split-tunnel setups often break")
    print("         outbound TLS to GitHub.\n")


def github_request(method, url, *, headers=None, retries=MAX_ATTEMPTS, **kwargs):
    """requests wrapper: bounded timeout, retry on transient network failure.

    Raises GitHubUnreachable when every attempt fails. HTTP error *statuses*
    (401, 404, 409…) are returned as normal responses — only transport-level
    failures are retried, since a 401 will never fix itself.
    """
    import requests

    kwargs.setdefault("timeout", (CONNECT_TIMEOUT, READ_TIMEOUT))
    last = None
    for attempt in range(1, retries + 1):
        try:
            return requests.request(method, url, headers=headers, **kwargs)
        except (requests.exceptions.Timeout,
                requests.exceptions.ConnectionError,
                requests.exceptions.SSLError) as e:
            last = e
            if attempt < retries:
                wait = RETRY_BACKOFF * attempt
                print(f"\n   ⚠  Network error (attempt {attempt}/{retries}) — "
                      f"retrying in {wait}s…", flush=True)
                time.sleep(wait)
    raise GitHubUnreachable(last)


def preflight():
    """Quick reachability probe before doing any real work.

    Returns (ok: bool, detail: str). Cheap — one unauthenticated GET with a
    short timeout, so a blocked network fails in seconds instead of hanging.
    """
    import requests
    try:
        r = requests.get("https://api.github.com/rate_limit",
                         timeout=(CONNECT_TIMEOUT, 20))
        return True, f"HTTP {r.status_code}"
    except Exception as e:
        return False, f"{type(e).__name__}: {str(e)[:120]}"
