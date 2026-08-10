#!/bin/bash
# deploy.sh — publish the dashboard in one command.
#
#   ./deploy.sh              process whatever changed in OneDrive, then push
#   ./deploy.sh --check      show what would run, change nothing
#   ./deploy.sh 202607       force one period
#   ./deploy.sh --all        reprocess every period
#   ./deploy.sh --ui         push index.html only
#   ./deploy.sh --source     push this project's own code/docs to the repo
#   ./deploy.sh --reset      forget change-detection state
#
# Runs from anywhere — it cd's to its own folder first.
set -euo pipefail
cd "$(dirname "$0")"
exec python3 deploy.py "$@"
