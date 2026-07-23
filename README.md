# job-hunt-il

An end-to-end job-hunt pipeline for the **Israeli market**, packaged as a
[Claude Code](https://claude.com/claude-code) skill. Built from a real,
running job hunt — every piece here earned its place by surviving actual use.

## What it does

- **Crawls Israeli job sources** — Jobify360 (aggregates the Israeli boards)
  and LinkedIn (Israel geo) — with a real browser, so it sees what you'd see.
- **Scores every job 0–100 against *your* profile** using Claude, with honest
  strengths/gaps per job. A score floor keeps noise out.
- **Tracks everything in one Excel file** — a Candidates sheet for crawler
  finds and an Applications sheet for jobs you actually pursue, with
  automatic promotion when you change a status.
- **Generates tailored CV + cover letter packages** (Word + PDF) per job:
  English via docx-js, **Hebrew via template-fill of your own Word CV** — the
  only Hebrew/RTL approach that renders reliably in real Microsoft Word.
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

- Windows with Microsoft Word (for PDF export of the generated documents)
- Python 3.11+ (`openpyxl`, `anthropic`, `python-docx`, `playwright`, `docx2pdf`)
- Node.js (for the English CV renderer)
- An Anthropic API key (in an environment variable — never stored in files)

## Getting started

Three things are yours to do. Everything else, Claude does.

### 1. Install (once)

```bash
git clone <this repo> job-hunt-il
cd job-hunt-il
install.bat
```

`install.bat` installs the Python packages, Chromium, and the CV renderer,
and scaffolds the config. Then set your Anthropic API key as an environment
variable named `ANTHROPIC_API_KEY` (get one at console.anthropic.com —
scoring a job costs fractions of an agora with Haiku; a full tailored CV
package is a few agorot with Sonnet).

### 2. Talk to Claude (once, ~20 minutes)

Open Claude Code in this folder (or copy the folder into `~/.claude/skills/`)
and say:

> **"run the job-hunt-il onboarding"**

Claude interviews you — start by handing it your current CV, in any
language — and then does ALL the file work itself: fills `config.json`,
builds `profile.md` (the facts of your career) and `positioning.md` (how to
tell them), creates your Excel tracker, and reads everything back for your
sign-off. You never edit a config file by hand.

The full interview script is in [docs/ONBOARDING.md](docs/ONBOARDING.md) if
you want to see what's coming.

### 3. Then just talk

| You say | What happens |
|---|---|
| "run the crawlers" | Claude runs both crawlers; a Chrome window opens — **logging in to LinkedIn/Jobify is your part** — then every job is scored 0–100 against your profile and lands in the tracker |
| "generate CVs for the top 5" (or mark **Gen CV = Y** yourself in Excel) | Claude marks the rows and runs the generator: each job gets a folder with a tailored CV + cover letter, Word and PDF. Hebrew jobs are rendered by filling YOUR OWN Hebrew Word CV ([docs/HEBREW_CV_TEMPLATE.md](docs/HEBREW_CV_TEMPLATE.md)) |
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
