# Automating the daily brief

The brief is on-demand by design (`python scripts/sync_tracker.py` then
`python scripts/brief.py`). If you want it to run every morning by itself:

## Windows (Task Scheduler)

1. Create `run_brief.bat` in the skill folder:

```bat
@echo off
cd /d "%~dp0"
python scripts\sync_tracker.py >> brief_log.txt 2>&1
python scripts\brief.py >> brief_log.txt 2>&1
```

2. Task Scheduler → Create Basic Task → Daily, pick a time (e.g. 07:00) →
   Start a program → point at `run_brief.bat`.
3. In the task's settings, enable "Run only when user is logged on" — the
   scripts read your user environment (API key env var).

The output accumulates in `brief_log.txt`. If you want it in your inbox
instead, the simplest reliable route is a mail step at the end of the .bat
using whatever mail CLI you already trust — the skill deliberately doesn't
ship email credentials handling.

## macOS / Linux (cron)

```cron
0 7 * * 1-5  cd /path/to/job-hunt-il && python scripts/sync_tracker.py && python scripts/brief.py >> brief_log.txt 2>&1
```

## Notes

- The sync + brief are read-mostly and safe to run unattended; the one rule
  is that the tracker must not be open in Excel at run time (sync writes to
  it). Schedule it for before you start your day.
- Crawling is NOT suited to unattended scheduling: LinkedIn needs a real,
  logged-in browser window. Run crawls manually.
