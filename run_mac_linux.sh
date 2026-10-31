#!/usr/bin/env bash
set -euo pipefail
cd -- "$(dirname -- "$0")"
if [ ! -x .venv/bin/python ]; then
  echo "Run the first-time setup in README.md before using this launcher."
  exit 1
fi
exec .venv/bin/python aqua_scraper.py "$@"
