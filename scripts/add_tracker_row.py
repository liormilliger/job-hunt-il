#!/usr/bin/env python3
"""
Permanent job tracker updater. Replaces the old write-a-temp-script-per-application workflow.

How it works:
1. Claude (or the user) writes the new application's details into new_row.json
   in the workdir (path comes from config.json).
2. Run:  python add_tracker_row.py
3. The script appends one formatted row to the tracker and renames
   new_row.json to last_row_added.json so it can't be double-imported.

new_row.json format (all strings unless noted):
{
  "date": "10/06/2026",          // optional, defaults to today
  "company": "Acme Ltd",
  "role": "Supply Chain Manager",
  "description": "One-liner about the role",
  "score": 85,                    // number
  "cv_path": "/Users/you/my-job-hunt/Acme_SC_Manager/CV.pdf",
  "cl_path": "/Users/you/my-job-hunt/Acme_SC_Manager/Cover_Letter.pdf",
  "status": "Draft",              // Draft | Applied | Phone Screen | Interview | Offer | Withdrawn
  "notes": ""
}

Rules preserved from the old workflow:
- Row 2 is the sample row, never overwritten. First empty row from row 3 down.
- Status values must match what the morning dashboard understands (see list above).
- The tracker is a Google Sheet now, so there's no "close Excel first" lock to
  worry about — Sheets handles concurrent access itself.
"""

import json
import os
import sys
from datetime import datetime

from gsheets_compat import load_workbook, Font, Alignment, Border, Side

import jh_config

_CFG = jh_config.load()
HERE = os.path.dirname(os.path.abspath(__file__))
TRACKER = _CFG["tracker"]
ROW_FILE = os.path.join(_CFG["workdir"], "new_row.json")
DONE_FILE = os.path.join(_CFG["workdir"], "last_row_added.json")

VALID_STATUSES = {"Draft", "Applied", "Phone Screen", "Interview", "Offer", "Withdrawn", "Rejected"}


def to_file_url(local_path):
    # Works for both Windows-style ("C:\...") and POSIX-style ("/Users/...")
    # paths — backslash-to-slash is a no-op on the latter.
    return "file:///" + local_path.replace("\\", "/")


def main():
    if not os.path.exists(ROW_FILE):
        print(f"ERROR: {ROW_FILE} not found. Write the application data there first.")
        sys.exit(1)

    with open(ROW_FILE, encoding="utf-8") as f:
        data = json.load(f)
    # new_row.json may hold one application (object) or several (list)
    rows = data if isinstance(data, list) else [data]

    print("Loading tracker...")
    wb = load_workbook(TRACKER)
    ws = wb.active

    added = []
    for d in rows:
        add_one(ws, d, added)

    print("Saving tracker...")
    wb.save(TRACKER)

    if os.path.exists(DONE_FILE):
        os.remove(DONE_FILE)
    os.rename(ROW_FILE, DONE_FILE)

    for line in added:
        print("OK: " + line)
    print("OK: input renamed to last_row_added.json")


def add_one(ws, d, added):
    required = ["company", "role"]
    missing = [k for k in required if not d.get(k)]
    if missing:
        print(f"ERROR: a row is missing: {', '.join(missing)}")
        sys.exit(1)
    # Applications made outside the filing workflow (applied directly on a site)
    # legitimately have no score or CV/CL files - those fields stay blank.

    status = d.get("status", "Draft")
    if status not in VALID_STATUSES:
        print(f"ERROR: status '{status}' not in {sorted(VALID_STATUSES)} "
              "(the morning dashboard only understands these).")
        sys.exit(1)

    empty_row = None
    for row_num in range(3, 1000):
        if ws[f"A{row_num}"].value is None:
            empty_row = row_num
            break
    if not empty_row:
        print("ERROR: no empty row found.")
        sys.exit(1)

    entry = {
        "A": d.get("date") or datetime.now().strftime("%d/%m/%Y"),
        "B": d["company"],
        "C": d["role"],
        "D": d.get("description", ""),
        "E": d.get("score", ""),
        "F": f'=HYPERLINK("{to_file_url(d["cv_path"])}", "CV")' if d.get("cv_path") else "",
        "G": f'=HYPERLINK("{to_file_url(d["cl_path"])}", "Cover Letter")' if d.get("cl_path") else "",
        "H": status,
        "I": d.get("notes", ""),
    }

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, value in entry.items():
        cell = ws[f"{col}{empty_row}"]
        cell.value = value
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        if col == "E":
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(bold=True)
        cell.border = border

    ws.row_dimensions[empty_row].height = 22
    added.append(f"'{d['company']} - {d['role']}' added at row {empty_row}")


if __name__ == "__main__":
    main()
