# -*- coding: utf-8 -*-
"""Pipeline status brief. Read-only; safe to run any time.

Usage:  python brief.py
Run sync_tracker.py first so Candidates promotions are reflected.
"""
import datetime
import sys
import gsheets_compat as openpyxl
import jh_config

sys.stdout.reconfigure(encoding="utf-8")

STALE_DAYS = 14


def parse_date(s):
    try:
        return datetime.datetime.strptime(str(s), "%d/%m/%Y").date()
    except (ValueError, TypeError):
        return None


def main():
    cfg = jh_config.load()
    wb = openpyxl.load_workbook(cfg["tracker"], read_only=True)
    today = datetime.date.today()

    # Applications
    apps = wb["Applications"]
    hdr = [c.value for c in next(apps.iter_rows(min_row=1, max_row=1))]

    def ci(sheet_hdr, name):
        for i, h in enumerate(sheet_hdr):
            if h and name.lower() in str(h).lower():
                return i
        return None

    c_stat = ci(hdr, "Status")
    counts, stale, interviews = {}, [], []
    for row in apps.iter_rows(min_row=2, values_only=True):
        status = str(row[c_stat] or "").strip()
        if not status or status == "Draft":
            continue
        counts[status] = counts.get(status, 0) + 1
        d = parse_date(row[0])
        if status in ("Phone Screen", "Interview"):
            interviews.append((row[1], row[2], status))
        elif status == "Applied" and d and (today - d).days >= STALE_DAYS:
            stale.append((row[1], row[2], (today - d).days))

    # Candidates worth a look: New, high score
    cands = wb["Candidates"]
    chdr = [c.value for c in next(cands.iter_rows(min_row=1, max_row=1))]
    cc_stat, cc_score = ci(chdr, "Status"), ci(chdr, "Match Score")
    floor = cfg.get("generate_score_floor", 70)
    fresh = []
    for row in cands.iter_rows(min_row=2, values_only=True):
        if str(row[cc_stat] or "") == "New":
            try:
                score = int(row[cc_score])
            except (TypeError, ValueError):
                continue
            if score >= floor:
                fresh.append((score, row[1], row[2]))
    fresh.sort(reverse=True)

    print(f"=== Job hunt brief — {today.strftime('%d/%m/%Y')} ===")
    print("\nPipeline:", ", ".join(f"{k}: {v}" for k, v in sorted(counts.items()))
          or "nothing filed yet")
    if interviews:
        print("\nActive conversations:")
        for co, role, st in interviews:
            print(f"  [{st}] {co} — {role}")
    if stale:
        print(f"\nNo response {STALE_DAYS}+ days (consider a follow-up):")
        for co, role, days in stale:
            print(f"  {co} — {role} ({days}d)")
    if fresh:
        print(f"\nNew candidates scoring {floor}+ (top 10):")
        for score, co, role in fresh[:10]:
            print(f"  [{score}] {co} — {role}")
    print()


if __name__ == "__main__":
    main()
