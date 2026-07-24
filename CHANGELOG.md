# Changelog

All notable changes to job-hunt-il are logged here.

## 2026-07-24 — Move to Google Sheets + Google Docs (Mac/Linux/Windows, no Microsoft Office)

Replaced the Excel + Microsoft Word dependency with Google Sheets + Google
Drive, so the pipeline runs the same on Windows, macOS, and Linux without
Office installed.

### New modules (`scripts/`)

- **`gsheets_compat.py`** — an openpyxl-compatible facade backed by the
  Sheets API. Every script that used to call `ws.cell(...)`, `.font = `,
  `.append()`, and so on against a local .xlsx now does the identical thing
  against a live Google Sheet, with almost no call-site changes.
- **`google_auth.py`** — shared OAuth flow against the user's own Google
  account (not a service account), scoped to `spreadsheets` +
  `drive.file` only.
- **`drive_pdf.py`** — replaces `docx2pdf`/Word: uploads the generated
  `.docx` to Drive, exports it as PDF, deletes the temporary Doc.

### Edited

- All seven tracker-touching scripts (both crawlers, `sync_tracker.py`,
  `brief.py`, `add_tracker_row.py`, `clean_failed_run.py`,
  `generate_applications.py`) now import from `gsheets_compat` instead of
  `openpyxl`. The "close Excel first" locking logic is gone — Sheets
  handles concurrent access natively.
- `setup.py` now creates a Google Sheet and saves its spreadsheet ID into
  `config.json`, instead of writing a local `.xlsx` file.
- `config.example.json` gained `google_credentials` / `google_token` fields.
- Added `install.sh` for Mac/Linux; stripped Word/openpyxl/docx2pdf out of
  `install.bat` too, so Windows now runs the same Google-based path.

### Docs

- New **`docs/GOOGLE_SETUP.md`** (bilingual, mirrors `docs/API_KEY.md`'s
  style) walking through the one-time OAuth client setup (~5 minutes).
- `README.md`, `SKILL.md`, `docs/SCHEDULING.md`, `docs/HEBREW_CV_TEMPLATE.md`,
  `docs/ONBOARDING.md`, `docs/API_KEY.md` updated to drop Windows/Word/Excel-
  only language and reflect the cross-platform Google Docs/Sheets flow.

### Known caveats

- **Untested end-to-end.** All edited Python files pass a syntax check, but
  this was built without live Google credentials to exercise it against.
  Treat `./install.sh` (or `install.bat`) → `python scripts/setup.py` as the
  first real smoke test.
- **Hebrew/RTL PDF fidelity is unverified.** Google Drive's docx→PDF
  conversion isn't guaranteed pixel-identical to Microsoft Word's. If you use
  a Hebrew CV template, check the very first generated PDF's RTL layout
  before trusting it fully.
