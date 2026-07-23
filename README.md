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

### Step 1 — install (10 minutes, once)

```bash
git clone <this repo> job-hunt-il
cd job-hunt-il
pip install openpyxl anthropic python-docx playwright docx2pdf
python -m playwright install chromium
cd scripts/cv_render && npm install && cd ../..
```

Set your Anthropic API key as an environment variable named
`ANTHROPIC_API_KEY` (get one at console.anthropic.com — scoring a job costs
fractions of an agora with Haiku; a full tailored CV package is a few
agorot with Sonnet).

### Step 2 — run setup twice

```bash
python scripts/setup.py    # checks dependencies, creates config.json
```

Open `config.json` and fill in your name, a `workdir` folder (where your
tracker and applications will live), and your target job titles. Then:

```bash
python scripts/setup.py    # creates the tracker + scaffolds profile.md, positioning.md
```

### Step 3 — the onboarding interview (the real setup)

Copy this folder into `~/.claude/skills/` and open a Claude Code session.
Say: **"run the job-hunt-il onboarding interview."** Claude will interview
you — start by handing it your current CV, in any language — and build
`profile.md` (the facts of your career) and `positioning.md` (how to tell
them). Budget 20–30 minutes and take it seriously: every CV the pipeline
ever writes comes from these two files, and Claude is forbidden from
inventing anything that isn't in them.

The full interview script is in [docs/ONBOARDING.md](docs/ONBOARDING.md), so
you can see exactly what's asked and why before you start.

### Step 4 — first crawl

```bash
cd scripts
python crawl_jobify.py      # Israeli boards via Jobify360 (needs your Jobify login)
python crawl_linkedin.py    # LinkedIn, Israel geo (needs your LinkedIn login)
```

A real Chrome window opens; log in when asked. Every job found is scored
0–100 against your profile; anything at or above the floor lands in your
tracker's Candidates sheet with honest strengths and gaps.

### Step 5 — generate applications

Open the tracker, put **Y** in the "Gen CV" column for the candidates you
want to pursue, close Excel, then:

```bash
python generate_applications.py
```

Each marked job gets a folder with a tailored CV + cover letter (Word and
PDF). Hebrew-language jobs are rendered by filling YOUR OWN Hebrew Word CV
(see [docs/HEBREW_CV_TEMPLATE.md](docs/HEBREW_CV_TEMPLATE.md)) — the only
approach we found that survives Word's RTL handling.

### Step 6 — run the pipeline day to day

```bash
python sync_tracker.py   # promotes candidates you marked Applied/Interview
python brief.py          # your pipeline at a glance
```

Change a candidate's Status to Applied / Phone Screen / Interview / Offer in
Excel and the sync moves it to the Applications sheet. `brief.py` shows what's
filed, what's stale, and which new finds are worth a look. To automate the
brief every morning, see [docs/SCHEDULING.md](docs/SCHEDULING.md). If a crawl
run fails mid-scoring (API credit ran out), nothing is lost:
`python clean_failed_run.py DD/MM/YYYY` removes exactly the failed batch and
the next crawl re-finds and re-scores those jobs.

## Daily use

| You say | What happens |
|---|---|
| "run the crawlers" | Crawl + score into the tracker's Candidates sheet |
| "generate CVs for today's finds" | Tailored packages for rows you marked Gen CV = Y |
| "add this job: \<url\>" | Score, folder, CV + cover, tracker row |
| "I have an interview at X" | Research + prep brief with salary data |
| "how's my pipeline?" | Sync + status brief |
| "the run failed, clean it" | Guarded removal of the failed batch only |

## Honesty notes

- Job boards change and fight automation; the crawlers are built
  defensively (they fail loudly, never silently) but expect occasional
  selector maintenance.
- Scoring costs real API money (~a few agorot per job with Haiku). If your
  API credit runs out mid-run, failed rows are flagged for cleanup — nothing
  is lost silently.

## License

MIT
