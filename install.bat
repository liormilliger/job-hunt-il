@echo off
REM job-hunt-il one-shot installer (Windows; Google Docs/Sheets - no
REM Microsoft Office needed)
echo === job-hunt-il install ===
pip install anthropic python-docx playwright google-api-python-client google-auth-httplib2 google-auth-oauthlib
python -m playwright install chromium
pushd scripts\cv_render
call npm install
popd
python scripts\setup.py
echo.
echo Next:
echo   1. Set your ANTHROPIC_API_KEY environment variable if you haven't (docs\API_KEY.md).
echo   2. Connect Google Sheets/Docs, one-time, ~5 minutes (docs\GOOGLE_SETUP.md).
echo   3. Open Claude Code in this folder and say:
echo        "run the job-hunt-il onboarding"
echo Claude fills in the rest - you never edit a config file by hand.
pause
