---
name: job-hunt-il
description: >
  End-to-end job-hunt pipeline for the Israeli market: crawl Israeli job
  boards (Jobify360 + LinkedIn Israel), score every job against the user's
  real profile with Claude, track everything in one Excel tracker, generate
  tailored CVs and cover letters (English and Hebrew, including reliable
  Hebrew Word rendering), prep for interviews, and produce a daily pipeline
  brief. Use when the user wants to crawl for jobs, score job fits, generate
  a tailored CV/cover letter, add a job to their tracker, prep for an
  interview, clean a failed run, or see their pipeline status.
---

# job-hunt-il — Israeli job-hunt pipeline

One skill, several modes. All state lives in the user's workdir (see
`config.json`); all code lives here. Nothing personal is stored in the skill
folder itself.

**Before any mode:** if `config.json` doesn't exist next to this file, run
setup first. Never hardcode user facts — everything personal comes from
`config.json`, `profile.md`, and `positioning.md`.

## Modes

### setup — first run
```
python scripts/setup.py
```
Interactive: checks dependencies (Python packages, Node for the English CV
renderer, Microsoft Word for PDF export, Playwright Chromium), copies
`config.example.json` → `config.json`, creates the Excel tracker, and scaffolds
`profile.md` + `positioning.md` from `templates/`.

After the script runs, YOU (Claude) run the onboarding interview. Work through
these areas in order; write the answers into `config.json`, `profile.md`, and
`positioning.md` as you go. If the user uploads or points to an existing CV,
READ IT FIRST and pre-fill everything you can — then confirm and fill gaps
instead of asking from zero.

1. **Personal basics** → config + profile header: full name (English + Hebrew
   spelling), city/area, email + phone as they should appear on a CV,
   LinkedIn URL. Never ask for or store passwords or payment details.
2. **Their CV** → the richest source. Ask them to drop in their current CV
   (any language/format). Extract roles, dates, team sizes, systems, real
   metrics into `profile.md`. Read numbers back to the user for confirmation —
   a wrong metric in profile.md poisons every generated document.
3. **Career facts not on the CV** → per role: what did you BUILD (systems
   implemented, processes created, zero-to-ones)? Any numbers you're sure of?
   Anything on the CV that's outdated or wrong?
4. **Target positions** → config `search_titles`: which titles, which
   seniority band, acceptable pivots. Also ask for the Hebrew AND English
   tab-title keywords Jobify360 would use to categorize their role → config
   `role_keywords`. This is separate from search_titles and easy to forget;
   outside the operations/supply-chain default wordlist, skipping it means
   Jobify silently never opens the user's own relevant tabs.
5. **Market scope** → within Israel: commute base and limits →
   `exclude_locations`; remote/hybrid preferences; industries to avoid, if any.
5b. **Score thresholds** → `min_score` (floor for landing in the tracker at
   all, default 65) and `generate_score_floor` (floor for getting a full CV
   package, default 70). Confirm or adjust — never leave silently defaulted.
6. **Compensation floor** → profile constraints (used for fit judgment only —
   never appears in documents; say that out loud so they answer honestly).
7. **The hard question** → positioning.md: every career has one (gaps, short
   stints, a pivot). Draft the honest, confident narrative WITH them.
8. **Hebrew CV template** → if they have a Hebrew CV in Word, wire it as
   `hebrew_cv_template` (explain the rules in docs/HEBREW_CV_TEMPLATE.md).
   If not, note that Hebrew jobs will get English packages. Also confirm
   `language_default` (default "auto" = job-posting language decides the CV
   language) — some users want every application in one fixed language
   regardless of the posting.
9. **API key** → confirm the env var named in `api_key_env` is set; if not,
   walk them through setting it. Never write the key into any file.

The facts/framing split is the core of this system: `profile.md` = what is
true, `positioning.md` = how to present it. End the onboarding by reading both
files back to the user for sign-off.

### crawl — find and score jobs
```
python scripts/crawl_jobify.py      # Jobify360 (aggregates Israeli boards)
python scripts/crawl_linkedin.py    # LinkedIn, Israel geo
```
Run one at a time (both write the same tracker — never in parallel). The user
must be at the computer: LinkedIn opens a real Chrome window and may need a
login. Every found job is scored 0-100 against `profile.md` by the scoring
model; jobs under `min_score` or in `exclude_locations` are skipped; the rest
land in the tracker's **Candidates** sheet as status "New".

Notes for you:
- The tracker must be CLOSED in Excel before any crawl or write. Check for a
  `~$` lock file next to the tracker; if present, ask the user to close Excel.
- If scoring fails (API error / credit limit), rows are written with score =
  `min_score` and Strengths = "SCORING FAILED — review/rescore manually". They
  are visible, not silently dropped. Clean them with `clean` and re-run after
  the API problem is fixed.

### generate — tailored CV + cover letter packages
```
python scripts/generate_applications.py
```
Processes Candidates rows with score ≥ `generate_score_floor`, `Gen CV` = Y,
and an empty `CV Folder`. For each: tailors a CV and cover letter to the job
description using the tailoring model (facts from `profile.md`, framing from
`positioning.md`), renders Word + PDF, writes a
`<DD-MM-YYYY>/<Company>_<Role>/` folder, and writes the folder path back to
the tracker.

- **English CVs** render via `scripts/cv_render/render_docx.js` (docx-js).
- **Hebrew CVs** render by FILLING the user's own Hebrew Word CV (the
  `hebrew_cv_template` in config) — swapping run text only, never building a
  docx from scratch. From-scratch Hebrew docx garbles in real Word; template
  fill is the only approach that survives. Rules the user's template must
  follow: `docs/HEBREW_CV_TEMPLATE.md`. If no template is configured, Hebrew
  jobs get an English CV and a note.
- Language is auto-detected per job (Hebrew JD → Hebrew CV) unless config says
  otherwise.

### add — file a job the user found themselves
The user pastes a URL or a job description. You:
1. Fetch/read the JD, score it against `profile.md` honestly (be strict).
2. If score < `generate_score_floor`, present the analysis and ask before
   filing.
3. Otherwise: create the `<Company>_<Role>/` folder, generate the package
   (same as generate mode), and append a tracker row via
   `python scripts/add_tracker_row.py` (write `new_row.json` first — format at
   the top of that script). Status starts as "Draft".

### prep — interview preparation
When the user has an interview, research live (web) and produce a prep brief
saved into the application's folder:
- The company: what it does, size, markets, recent news, the specific
  operational reality behind the role.
- The interviewer: role, background, what language they think in — and flag
  name ambiguities before the user walks in with the wrong name.
- Fit: the user's 3-4 strongest hooks for THIS role, from `profile.md`.
- Gaps: honest list, each with a framing script (never denial — concede,
  reframe, pivot to adjacent strength).
- Questions to ask: ones that prove the user understands the business.
- Salary: pull current Israeli market data (Manpower salary tables and
  similar) for the role level; give target / anchor / walk-away. Warn the
  user if their own floor is below market. Board-listing "AI estimated"
  salary ranges are not employer data — say so.

### brief — pipeline status
```
python scripts/sync_tracker.py    # promote Candidates → Applications first
python scripts/brief.py           # then print the status brief
```
The brief shows: applications filed / awaiting response / interviews ahead,
new high-scoring candidates worth a look, and stale applications (no movement
in 14+ days). Present it conversationally, not as a raw dump. To automate the
daily run, see `docs/SCHEDULING.md`.

### clean — remove a failed crawl batch
```
python scripts/clean_failed_run.py DD/MM/YYYY
```
Deletes ONLY rows that match BOTH the given Date Found AND the full failure
signature (score = min_score, "SCORING FAILED" in Strengths, status "New", no
CV Folder). Rows the user has touched (status changed, CV generated) are never
deleted. Always report matched/kept/deleted counts. Cleaned jobs are
re-found and re-scored on the next crawl (dedup no longer sees them).

## Conventions (enforce everywhere)

- **The tracker is the single source of truth.** Never maintain a parallel
  list in markdown. Candidates = crawler finds; Applications = things the
  user actually pursued. `sync_tracker.py` promotes rows when the user sets
  Status to Applied / Phone Screen / Interview / Offer.
- **Statuses** (exact strings): New, Skip, Moved (Candidates); Draft, Applied,
  Phone Screen, Interview, Offer, Rejected, Withdrawn (Applications).
- **One folder per application**: `<Company>_<Role>/` with CV + cover letter
  (Word and PDF). Crawler-generated packages live under a `<DD-MM-YYYY>/` date
  folder.
- **Facts discipline**: every CV line traces to `profile.md`. No invented
  metrics, ever. If a tailored draft contains a number not in the profile,
  remove it.
- **Writing style**: plain language, no AI-tells ("leverage", "seamless",
  "passionate"), sentence variety, em-dashes sparingly. Hebrew: direct
  register, no melitzot.
- **Never write to the tracker while it's open in Excel.**
- **API key** comes from the env var named in config. Never write it to any
  file.
