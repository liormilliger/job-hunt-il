#!/usr/bin/env bash
# job-hunt-il one-shot installer (macOS/Linux; Google Docs/Sheets — no
# Microsoft Office needed, no LibreOffice needed either)
set -e
echo "=== job-hunt-il install ==="

# Recommended: run this from inside a virtualenv, e.g.
#   python3 -m venv .venv && source .venv/bin/activate
pip3 install anthropic python-docx playwright \
    google-api-python-client google-auth-httplib2 google-auth-oauthlib
python3 -m playwright install chromium
(cd scripts/cv_render && npm install)
python3 scripts/setup.py

echo
echo "Next:"
echo "  1. Set your ANTHROPIC_API_KEY environment variable if you haven't"
echo "     (docs/API_KEY.md)."
echo "  2. Connect Google Sheets/Docs, one-time, ~5 minutes (docs/GOOGLE_SETUP.md)."
echo "  3. Open Claude Code in this folder and say:"
echo "       \"run the job-hunt-il onboarding\""
echo "Claude fills in the rest - you never edit a config file by hand."
