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

## Install

```bash
git clone <this repo> job-hunt-il
cd job-hunt-il
pip install openpyxl anthropic python-docx playwright docx2pdf
python -m playwright install chromium
cd scripts/cv_render && npm install && cd ../..
python scripts/setup.py        # creates config.json — edit it
python scripts/setup.py        # scaffolds tracker + profile + positioning
```

To use it as a Claude Code skill, copy (or clone) this folder into
`~/.claude/skills/`. Then ask Claude to interview you to fill `profile.md`
and `positioning.md` — that conversation is the real setup.

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
