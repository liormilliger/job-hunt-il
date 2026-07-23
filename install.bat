@echo off
REM job-hunt-il one-shot installer (Windows; MS Word required for PDF export)
echo === job-hunt-il install ===
pip install openpyxl anthropic python-docx playwright docx2pdf
python -m playwright install chromium
pushd scripts\cv_render
call npm install
popd
python scripts\setup.py
echo.
echo Next: set your ANTHROPIC_API_KEY environment variable if you haven't,
echo then open Claude Code in this folder and say:
echo    "run the job-hunt-il onboarding"
echo Claude fills in the rest - you never edit a config file by hand.
pause
