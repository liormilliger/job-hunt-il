#!/usr/bin/env python3
"""
Tracker sync: promote crawler candidates to real applications.

How to use it:
1. In the Candidates tab, change a row's Status from "New" to one of:
   Applied / Phone Screen / Interview / Offer  (case doesn't matter)
2. Run this script (it also runs automatically every morning before the brief):
   - the row is COPIED to the Applications sheet with today's date and your status
   - the Candidates row status becomes "Moved" so it's never copied twice
   - nothing is ever deleted

Dedup: if Applications already has a row with the same company + role,
the candidate is just marked "Moved" and not copied again.

The tracker is a Google Sheet, so there's no "close Excel first" lock to
worry about.
"""

import os
import sys
from datetime import datetime

from gsheets_compat import load_workbook, Font, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
import jh_config
TRACKER = jh_config.load()["tracker"]

PROMOTE_STATUSES = {"applied": "Applied", "phone screen": "Phone Screen",
                    "interview": "Interview", "offer": "Offer"}

# Candidates tab columns (0-based): 0 Date Found, 1 Company, 2 Job Role,
# 3 Location, 4 International, 5 Match Score, 6 Strengths, 7 Key Gaps,
# 8 Source, 9 URL, 10 JD snippet, 11 Status


def first_empty_row(ws, start=3):
    for r in range(start, 2000):
        if ws[f"A{r}"].value is None:
            return r
    return None


def main():
    try:
        wb = load_workbook(TRACKER)
    except FileNotFoundError as e:
        print(f"ERROR: Google auth not set up yet ({e})")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: could not load the tracker Google Sheet: {e}")
        sys.exit(1)

    if "Candidates" not in wb.sheetnames:
        print("No Candidates sheet found. Nothing to sync.")
        sys.exit(0)
    ws_cand = wb["Candidates"]
    ws_app = wb["Applications"] if "Applications" in wb.sheetnames else wb.active

    # Existing applications for dedup (company+role, lowercase)
    existing = set()
    for row in ws_app.iter_rows(min_row=2, values_only=True):
        if row and (row[1] or row[2]):
            existing.add((str(row[1] or "").strip().lower(), str(row[2] or "").strip().lower()))

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    promoted, skipped = [], []
    for r in range(2, ws_cand.max_row + 1):
        status_cell = ws_cand.cell(row=r, column=12)
        status = str(status_cell.value or "").strip().lower()
        if status not in PROMOTE_STATUSES:
            continue
        nice_status = PROMOTE_STATUSES[status]

        company = str(ws_cand.cell(row=r, column=2).value or "").strip()
        role = str(ws_cand.cell(row=r, column=3).value or "").strip()
        jd = " ".join(str(ws_cand.cell(row=r, column=11).value or "").split())
        # some crawler snippets captured site boilerplate instead of the JD - don't copy junk
        if "תנאי שימוש" in jd or "מדיניות פרטיות" in jd or jd.startswith("×"):
            jd = ""
        source = str(ws_cand.cell(row=r, column=9).value or "").strip()
        url = str(ws_cand.cell(row=r, column=10).value or "").strip()
        score = ws_cand.cell(row=r, column=6).value or 0

        key = (company.lower(), role.lower())
        if key in existing:
            skipped.append(f"{company} - {role} (already in Applications)")
        else:
            target = first_empty_row(ws_app)
            if not target:
                print("ERROR: no empty row in Applications sheet.")
                sys.exit(1)
            values = {
                "A": datetime.now().strftime("%d/%m/%Y"),
                "B": company,
                "C": role,
                "D": jd[:120],
                "E": int(score) if isinstance(score, (int, float)) else 0,
                "F": "",  # CV link - add when a tailored CV exists
                "G": "",  # Cover letter link
                "H": nice_status,
                "I": f"Promoted from Candidates ({source}) {url}".strip(),
            }
            for col, value in values.items():
                cell = ws_app[f"{col}{target}"]
                cell.value = value
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                if col == "E":
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(bold=True)
                cell.border = border
            ws_app.row_dimensions[target].height = 22
            existing.add(key)
            promoted.append(f"{company} - {role} -> {nice_status}")

        status_cell.value = "Moved"

    if not promoted and not skipped:
        print("Nothing to sync - no Candidates rows with a promotable status.")
        sys.exit(0)

    wb.save(TRACKER)
    print(f"Synced. {len(promoted)} promoted, {len(skipped)} already existed.")
    for p in promoted:
        print("  + " + p)
    for s in skipped:
        print("  = " + s)


if __name__ == "__main__":
    main()
