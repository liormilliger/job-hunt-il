# job-hunt-il

An end-to-end job-hunt pipeline for the **Israeli market**, packaged as a
[Claude Code](https://claude.com/claude-code) skill. Built from a real,
running job hunt — every piece here earned its place by surviving actual use.

## What it does

- **Crawls Israeli job sources** — Jobify360 (aggregates the Israeli boards)
  and LinkedIn (Israel geo) — with a real browser, so it sees what you'd see.
- **Scores every job 0–100 against *your* profile** using Claude, with honest
  strengths/gaps per job. A score floor keeps noise out.
- **Tracks everything in one Google Sheet** — a Candidates sheet for crawler
  finds and an Applications sheet for jobs you actually pursue, with
  automatic promotion when you change a status. No Excel required; open it
  from a browser or the Google Sheets app.
- **Generates tailored CV + cover letter packages** (Word-format .docx + PDF)
  per job: English via docx-js, **Hebrew via template-fill of your own Word
  CV** — the only Hebrew/RTL approach that renders reliably. PDF export goes
  through Google Drive's conversion, so no Word or LibreOffice install is
  needed on your machine.
- **Preps you for interviews** — live company & interviewer research, fit
  hooks, gap-framing scripts, questions to ask, and current Israeli salary
  benchmarks.
- **Daily pipeline brief** — filed / awaiting / interviews / fresh
  high-scoring candidates / stale applications needing a follow-up.

## What makes it different

The **facts/framing split**. Your `profile.md` holds only what's true (real
numbers, real dates). Your `positioning.md` holds how to present it (which
story leads for which role type, and your answer to your career's hard
question). Claude writes every document from those two files and is forbidden
from inventing metrics. Tailored, but never fabricated.

## Requirements

- **Windows, macOS, or Linux** — no Microsoft Office required. Documents are
  generated as real .docx files and converted to PDF through Google Drive,
  and the tracker is a Google Sheet, so nothing here depends on a specific OS
  or an Office install.
- Python 3.11+ (`anthropic`, `python-docx`, `playwright`,
  `google-api-python-client`, `google-auth-oauthlib`)
- Node.js (for the English CV renderer)
- An Anthropic API key (in an environment variable — never stored in files)
- A Google account, with a one-time OAuth setup (~5 minutes) so the pipeline
  can create/edit a Sheet and convert documents to PDF in your own Drive —
  see [docs/GOOGLE_SETUP.md](docs/GOOGLE_SETUP.md)

## Getting started

Three things are yours to do. Everything else, Claude does.

### 1. Install (once)

```bash
git clone <this repo> job-hunt-il
cd job-hunt-il
./install.sh        # macOS/Linux
install.bat          # Windows
```

The installer installs the Python packages, Chromium, and the CV renderer,
and scaffolds the config. Then:
- Set your Anthropic API key as an environment variable named
  `ANTHROPIC_API_KEY` — 5-minute walkthrough (English + Hebrew, with real
  cost numbers): [docs/API_KEY.md](docs/API_KEY.md).
- Connect Google Sheets/Docs — 5-minute one-time setup:
  [docs/GOOGLE_SETUP.md](docs/GOOGLE_SETUP.md). The first script that touches
  the tracker or generates a PDF will open a browser for you to sign in;
  after that it's silent.

### 2. Talk to Claude (once, ~20 minutes)

Open Claude Code in this folder (or copy the folder into `~/.claude/skills/`)
and say:

> **"run the job-hunt-il onboarding"**

Claude interviews you — start by handing it your current CV, in any
language — and then does ALL the file work itself: fills `config.json`,
builds `profile.md` (the facts of your career) and `positioning.md` (how to
tell them), creates your Google Sheet tracker, and reads everything back for
your sign-off. You never edit a config file by hand.

The full interview script is in [docs/ONBOARDING.md](docs/ONBOARDING.md) if
you want to see what's coming.

### 3. Then just talk

| You say | What happens |
|---|---|
| "run the crawlers" | Claude runs both crawlers; a Chrome window opens — **logging in to LinkedIn/Jobify is your part** — then every job is scored 0–100 against your profile and lands in the tracker |
| "generate CVs for the top 5" (or mark **Gen CV = Y** yourself in the tracker) | Claude marks the rows and runs the generator: each job gets a folder with a tailored CV + cover letter, .docx and PDF. Hebrew jobs are rendered by filling YOUR OWN Hebrew Word CV ([docs/HEBREW_CV_TEMPLATE.md](docs/HEBREW_CV_TEMPLATE.md)) |
| "add this job: \<url\>" | Claude fetches the posting, scores it honestly, and files the full package: folder, CV, cover letter, tracker row |
| "I have an interview at X" | Live research on the company and interviewer, your strongest hooks, gap-framing scripts, questions to ask, current Israeli salary benchmarks |
| "how's my pipeline?" | Claude syncs the tracker (Applied/Interview promotions) and gives you the brief: filed, awaiting, stale, fresh finds worth a look |
| "the run failed, clean it" | Claude removes exactly the failed batch (guarded — rows you touched are never deleted); the next crawl re-finds and re-scores them |

The two things that stay manual by design: **logging in** to the job boards
(it's your account, in a real browser window), and **deciding which jobs to
pursue**. Everything mechanical belongs to Claude.

To automate the daily brief without asking, see
[docs/SCHEDULING.md](docs/SCHEDULING.md).

## Honesty notes

- Job boards change and fight automation; the crawlers are built
  defensively (they fail loudly, never silently) but expect occasional
  selector maintenance.
- Scoring costs real API money (~a few agorot per job with Haiku). If your
  API credit runs out mid-run, failed rows are flagged for cleanup — nothing
  is lost silently.

## License

MIT
