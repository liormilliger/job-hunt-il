"""
Generate application packages — job-hunt-il
=========================================================
On-demand script (run whenever you're ready to review the day's candidates,
not automatically after each crawl).

Scans the tracker's Candidates sheet for rows the user has marked
"Y" in the "Gen CV" column that haven't been processed yet (tracked via the
"CV Folder" column), and for each one generates a complete application package:
  job-hunt-26/<DD-MM-YYYY>/<Company>_<Role>/
    CV_<Name>.docx + .pdf
    Cover_Letter.docx + .pdf   (cover email opening + interview prep notes)

Scope: GOING FORWARD ONLY. The first time this script runs, every existing
Candidates row gets backfilled with "(pre-feature)" in the new CV Folder
column so it's permanently excluded (going-forward-only scope).

Language: auto-detected per candidate (Hebrew chars in title/JD -> Hebrew
CV; otherwise English), matching the existing JD-language rule.

Source of truth for facts: profile.md. The model is instructed to only
reorder/select/rephrase from these facts — never invent new numbers or
claims. Invented facts in a CV are career-damaging; the prompt treats this
as the highest-priority rule.

Run:  python generate_cvs_from_candidates.py
"""

import sys
import os
import json
import re
import subprocess
import time
import datetime
import anthropic
import gsheets_compat as openpyxl
from gsheets_compat import Font, PatternFill, Alignment
from drive_pdf import convert as docx2pdf_convert
from hebrew_render import render_cv_from_template, render_cover_from_template

sys.stdout.reconfigure(encoding="utf-8")

# ── CONFIG ───────────────────────────────────────────────────────────────────
import jh_config
_CFG = jh_config.load()
HERE = os.path.dirname(os.path.abspath(__file__))          # skill scripts/
JOB_HUNT_ROOT = _CFG["output_dir"]                          # user workdir output
TRACKER_PATH = _CFG["tracker"]
PROFILE_PATH = _CFG["profile_md"]
CANDIDATE_NAME = _CFG.get("candidate_name", "Candidate")
_CV_BASENAME = "CV_" + re.sub(r"[^A-Za-z0-9֐-׿]+", "_", CANDIDATE_NAME).strip("_")
HEBREW_TEMPLATE = _CFG.get("hebrew_cv_template", "")
RENDER_SCRIPT = os.path.join(HERE, "cv_render", "render_docx.js")
CV_FOLDER_HEADER = "CV Folder"
GEN_CV_HEADER = "Gen CV"
PRE_FEATURE_SENTINEL = "(pre-feature)"

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
if not ANTHROPIC_API_KEY:
    try:
        import winreg
        _reg = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment")
        ANTHROPIC_API_KEY, _ = winreg.QueryValueEx(_reg, "ANTHROPIC_API_KEY")
        winreg.CloseKey(_reg)
    except Exception:
        pass

WRITING_RULES = """
- American spelling.
- Plain English. No jargon, buzzwords, or AI tells (no "Certainly!", "leverage", "synergy", "seamless", etc.).
- Get to the point fast. No hollow openers, no summary padding.
- Vary sentence length. Don't bullet everything.
- NO EM DASHES in any prose sentence (the summary, bullets, cover email, framing, or angle text).
  Use a period, comma, or colon instead. This is a hard rule, not a stylistic suggestion. The user has
  flagged em-dash overuse as an AI tell multiple times. The "Company — Role" header format and
  section heading separators (shown in the JSON shape below) are a fixed template convention, not
  prose. Keep those as shown. Just keep em dashes out of anything you actually write as a sentence.
"""

# Positioning guidance — mirrors job-hunt-26/positioning.md. Kept inline because
# this script runs standalone (can't read CLAUDE.md). If positioning.md changes
# materially, update this block too.
# Framing comes from the user's positioning.md; fall back to a neutral rule.
POSITIONING = jh_config.positioning_text(_CFG) or """
CHOOSE THE FRAMING FROM THE ROLE: read the job description first, then decide
which side of the candidate's experience should lead the summary and bullets.
Present the rest as supporting range, not the headline.

BUILT-AND-FINISHED: for every role, lead the bullets with what the candidate
BUILT and DELIVERED (a zero-to-one, a system rollout, a turnaround), not what
they maintained.

Do not invent reasons for any career move. Do not use gendered language about
the candidate unless the profile states pronouns explicitly.
"""

_HEBREW_RE = re.compile(r"[֐-׿]")


def detect_language(text: str) -> str:
    """Hebrew if the text contains Hebrew characters, else English."""
    return "hebrew" if _HEBREW_RE.search(text or "") else "english"


# ── PROFILE FACTS (source of truth — never invent beyond this) ──────────────
def load_profile_facts() -> str:
    with open(PROFILE_PATH, "r", encoding="utf-8") as f:
        return f.read()


# ── TRACKER HELPERS ──────────────────────────────────────────────────────────
def ensure_cv_folder_column(ws):
    """Idempotent: add the CV Folder column once. On first creation, backfill
    every existing row with the pre-feature sentinel so old candidates are
    never picked up (going-forward-only scope)."""
    header_row = ws[1]
    for cell in header_row:
        if cell.value == CV_FOLDER_HEADER:
            return cell.column

    col_idx = ws.max_column + 1
    header_cell = ws.cell(row=1, column=col_idx, value=CV_FOLDER_HEADER)
    header_cell.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    header_cell.fill = PatternFill("solid", fgColor="1F4E79")
    header_cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 24

    backfilled = 0
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=2).value:  # has data (Company column)
            ws.cell(row=row, column=col_idx, value=PRE_FEATURE_SENTINEL)
            backfilled += 1
    print(f"Created '{CV_FOLDER_HEADER}' column. Backfilled {backfilled} pre-existing rows "
          f"as '{PRE_FEATURE_SENTINEL}' (going-forward-only scope).")
    return col_idx


def ensure_gen_cv_column(ws):
    """Idempotent: add the Gen CV column once. No backfill — all rows start
    blank. User marks Y on any row they want a CV generated for."""
    for cell in ws[1]:
        if cell.value == GEN_CV_HEADER:
            return cell.column

    col_idx = ws.max_column + 1
    header_cell = ws.cell(row=1, column=col_idx, value=GEN_CV_HEADER)
    header_cell.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    header_cell.fill = PatternFill("solid", fgColor="375623")  # dark green
    header_cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 10
    print(f"Created '{GEN_CV_HEADER}' column. Mark rows with 'Y' to queue them for CV generation.")
    return col_idx


def find_eligible_rows(ws, cv_col, gen_cv_col):
    """Gen CV = 'Y' (case-insensitive) and CV Folder column blank."""
    eligible = []
    for row in range(2, ws.max_row + 1):
        gen_cv = str(ws.cell(row=row, column=gen_cv_col).value or "").strip().upper()
        cv_folder = ws.cell(row=row, column=cv_col).value
        if gen_cv == "Y" and not cv_folder:
            eligible.append(row)
    return eligible


def sanitize_folder_name(text: str) -> str:
    text = re.sub(r"[^\w\s-]", "", text, flags=re.UNICODE).strip()
    text = re.sub(r"\s+", "_", text)
    return text[:60] or "Unknown"


# ── CLAUDE: TAILOR CONTENT ──────────────────────────────────────────────────
def tailor_application(profile_facts: str, candidate: dict, language: str) -> dict:
    """Returns a dict with both the CV spec and cover-letter spec (matching
    render_docx.js's expected JSON shape), tailored to this specific role."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    lang_instructions = (
        "Write ALL output in Hebrew (the job posting is in Hebrew). "
        "Use natural, professional Hebrew phrasing, not a literal translation. "
        "DATE FORMAT — hard rule: use numeric MM/YYYY dates only (e.g. 11/2024 - 06/2026). "
        "NEVER use Hebrew month names (נוב׳, ספט׳, אפר׳, etc.) anywhere in the JSON. "
        "EXPERIENCE TITLE FORMAT: the 'title' field for each experience entry must be "
        "the Company name in English, then ' — ' (em dash with spaces), then the Hebrew "
        "job title. Example: 'Acme Logistics — מנהל תפעול'. The ' — ' here is the fixed "
        "template separator between company and title; it is not prose and is allowed."
        if language == "hebrew" else
        "Write ALL output in English."
    )

    prompt = f"""You are tailoring a job application package for {CANDIDATE_NAME}. You will produce a tailored CV and a cover letter +
interview prep package as a single JSON object.

CRITICAL RULE: Use ONLY the facts given below in CANDIDATE PROFILE FACTS. You may
reorder, select, emphasize, and rephrase, but you may NEVER invent a new number,
metric, company detail, or claim that isn't already in these facts. Invented facts in a CV are
career-damaging for the candidate — treat this as the highest-priority rule. If you are unsure whether something is a fact
or an inference, leave it out. Where the profile gives a precomputed number (e.g. total years of
experience), use that number exactly as given. Do not calculate your own
total by adding up date ranges; that produces wrong numbers.

WRITING STYLE — apply to every sentence you write:
{WRITING_RULES}

POSITIONING — decide what to lead with based on THIS role, then frame accordingly:
{POSITIONING}

SUMMARY ENDING — hard rule, checked every time: the professional summary must NOT end
on a sentence about availability or relocation; close on something substantive. Bad example (do not do this): "...
Available immediately and open to relocation to East Africa." Good example: put that
fact in the middle, and close on something substantive instead, e.g. "... Available
immediately for relocation to East Africa. Comfortable building operations from zero
or turning around an underperforming site." Before finalizing, check your own summary
field: if the last sentence is about availability or relocation, rewrite it so it isn't.

{lang_instructions}

CANDIDATE PROFILE FACTS (markdown, source of truth):
{profile_facts}

JOB POSTING:
Title: {candidate['title']}
Company: {candidate['company']}
Source: {candidate['source']}
Description (truncated): {candidate['jd_snippet']}
Already-identified strengths: {candidate['strengths']}
Already-identified gaps: {candidate['gaps']}
Fit score: {candidate['score']}/100

Produce a JSON object with this EXACT shape (no markdown fences, no commentary,
just the JSON):

{{
  "cv": {{
    "type": "cv",
    "rtl": {"true" if language == "hebrew" else "false"},
    "name": "<{CANDIDATE_NAME}, in the CV's language (transliterate to Hebrew if the profile gives a Hebrew name)>",
    "contact": "<contact line, same format as the profile, translated if Hebrew>",
    "summaryHeading": "<'Professional Summary' or 'סיכום מקצועי'>",
    "summary": "<3-4 sentence summary tailored to this role, using only profile facts. The LAST sentence must be substantive (impact, scope, capability). Availability/relocation, if mentioned at all, goes in the SECOND or THIRD sentence, never the last one.>",
    "experienceHeading": "<'Work Experience' or 'ניסיון תעסוקתי'>",
    "experience": [
      {{"title": "<Company — Role>", "dates": "<date range>", "bullets": ["<bullet tailored to this role, from profile facts only>", "..."]}}
    ],
    "educationHeading": "<'Education' or 'השכלה'>",
    "education": [
      {{"degree": "...", "school": "...", "honors": "...", "dates": "..."}}
    ],
    "skillsHeading": "<'Skills' or 'כישורים'>",
    "skills": [
      {{"label": "<category label, with trailing colon>", "items": "<items separated by ' \\u00b7 '>"}}
    ]
  }},
  "cover": {{
    "type": "cover",
    "rtl": {"true" if language == "hebrew" else "false"},
    "meta": {{
      "<'Role' or 'תפקיד'>": "{candidate['title']}",
      "<'Company' or 'חברה'>": "{candidate['company']}",
      "<'Date' or 'תאריך'>": "{candidate['date']}",
      "<'Fit Score' or 'ציון התאמה'>": "{candidate['score']} / 100",
      "<'Source' or 'מקור'>": "{candidate['source']}"
    }},
    "headingCover": "<'Cover Email Opening' or 'פתיח מכתב מועמדות'>",
    "coverEmail": "<2-4 sentence opening paragraph for the application email, direct and specific, no hollow openers>",
    "headingPrep": "<'Interview Prep \\u2014 Gaps and How to Frame Them' or 'הכנה לראיון \\u2014 חוסרים ואיך למסגר אותם'>",
    "gaps": [
      {{
        "gap": "<short label for the gap, derived from 'Already-identified gaps' above>",
        "framingLabel": "<'Framing' or 'מסגור'>",
        "framing": "<concrete framing strategy using real profile facts>",
        "angleLabel": "<'Angle' or 'זווית'>",
        "angle": "<a sentence the candidate could literally say in the interview>"
      }}
    ]
  }}
}}

CV LENGTH — aim for one page, like the established CVs. It's fine if it bleeds a little onto a
second page; don't cut content or force artificially terse bullets just to avoid that. As a
starting point:
- The two most recent roles can have 2-3 bullets each, whichever are most
  relevant to this specific job posting.
- Older roles normally get one concise bullet each, combining the most relevant
  facts into a single sentence, but use two if the role is genuinely relevant to
  this job and one would lose something that matters.
- The summary is 3-4 sentences, not more.
Include all roles from the profile's career timeline."""

    def call(messages, max_tokens=8000):
        msg = client.messages.create(
            model=_CFG.get("tailoring_model", "claude-sonnet-5"),
            max_tokens=max_tokens,
            messages=messages,
        )
        raw = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text").strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw).strip()
        return raw

    # Hebrew responses use noticeably more tokens per character than English
    # for equivalent content, which truncated the JSON mid-string at the old
    # max_tokens=4000 (silent failure: "Unterminated string" on json.loads).
    # 8000 covers it; this retry is a second line of defense in case an
    # unusually long response still gets cut off.
    messages = [{"role": "user", "content": prompt}]
    raw = call(messages)
    try:
        content = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"  WARNING: JSON parse failed ({e}). Likely truncated output. Retrying with more room...")
        messages.append({"role": "assistant", "content": raw})
        messages.append({"role": "user", "content": (
            "That output was not valid, complete JSON (it looks like it got cut off). "
            "Produce the full JSON object again, identical content, but make sure it is "
            "complete and well-formed. Output ONLY the JSON object, nothing else."
        )})
        raw = call(messages, max_tokens=12000)
        content = json.loads(raw)

    violations = find_em_dash_violations(content)
    if violations:
        print(f"  WARNING: em dash found in prose ({len(violations)} field(s)). Retrying once...")
        messages.append({"role": "assistant", "content": raw})
        messages.append({"role": "user", "content": (
            "These prose fields still contain an em dash, which is not allowed there: "
            + "; ".join(f"{path}: \"{text}\"" for path, text in violations)
            + ". Rewrite the full JSON object with those sentences fixed (period, comma, or colon "
            "instead of the em dash). Keep everything else the same. Output ONLY the JSON object."
        )})
        raw = call(messages)
        content = json.loads(raw)
        remaining = find_em_dash_violations(content)
        if remaining:
            print(f"  WARNING: {len(remaining)} em dash(es) still present after retry. Proceeding anyway.")

    return content


def find_em_dash_violations(content: dict) -> list:
    """Check only prose fields (not header/heading template strings) for a
    literal em dash. Returns [(field_path, offending_text), ...]."""
    violations = []
    cv = content.get("cv", {})
    if "—" in (cv.get("summary") or ""):
        violations.append(("cv.summary", cv["summary"]))
    for i, role in enumerate(cv.get("experience", [])):
        for j, b in enumerate(role.get("bullets", [])):
            if "—" in b:
                violations.append((f"cv.experience[{i}].bullets[{j}]", b))
    cover = content.get("cover", {})
    if "—" in (cover.get("coverEmail") or ""):
        violations.append(("cover.coverEmail", cover["coverEmail"]))
    for i, gap in enumerate(cover.get("gaps", [])):
        if "—" in (gap.get("framing") or ""):
            violations.append((f"cover.gaps[{i}].framing", gap["framing"]))
        if "—" in (gap.get("angle") or ""):
            violations.append((f"cover.gaps[{i}].angle", gap["angle"]))
    return violations


# ── RENDER: docx via Node, then PDF via Google Drive ──────────────────────────
def render_docx(spec: dict, out_path: str):
    tmp_json = out_path + ".tmp.json"
    with open(tmp_json, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False)
    try:
        # encoding="utf-8" matters here: node's stdout/stderr can contain
        # Hebrew (e.g. in printed file paths), and without an explicit
        # encoding Python falls back to the Windows default (cp1252), which
        # crashes the subprocess output reader thread on non-Latin bytes.
        result = subprocess.run(
            ["node", RENDER_SCRIPT, tmp_json, out_path],
            capture_output=True, text=True, timeout=60, encoding="utf-8",
        )
        if result.returncode != 0:
            raise RuntimeError(f"render_docx.js failed: {result.stderr}")
    finally:
        if os.path.exists(tmp_json):
            os.remove(tmp_json)


def docx_to_pdf(docx_path: str, pdf_path: str, retries: int = 2):
    """Converts via Google Drive (upload -> auto-convert -> export PDF -> delete
    temp Doc); see drive_pdf.py. This is a network round-trip per document, so
    transient failures (rate limit, dropped connection) are worth a retry
    before giving up."""
    last_error = None
    for attempt in range(retries + 1):
        try:
            docx2pdf_convert(docx_path, pdf_path, cfg=_CFG)
            return
        except Exception as e:
            last_error = e
            if attempt < retries:
                print(f"    Drive PDF conversion failed ({e}), retrying in 3s...")
                time.sleep(3)
    raise last_error


# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    if not ANTHROPIC_API_KEY:
        print("ERROR: ANTHROPIC_API_KEY not set. Open a new terminal and try again.")
        return

    if not os.path.exists(RENDER_SCRIPT):
        print(f"ERROR: render script not found at {RENDER_SCRIPT}")
        return

    today_dash = datetime.date.today().strftime("%d-%m-%Y")
    today_slash = datetime.date.today().strftime("%d/%m/%Y")

    print(f"\n{'='*60}")
    print(f"Generate CVs from Candidates — {today_slash}")
    print(f"{'='*60}\n")

    wb = openpyxl.load_workbook(TRACKER_PATH)
    if "Candidates" not in wb.sheetnames:
        print("No Candidates sheet found. Nothing to do.")
        return
    ws = wb["Candidates"]

    cv_col = ensure_cv_folder_column(ws)
    gen_cv_col = ensure_gen_cv_column(ws)
    wb.save(TRACKER_PATH)  # persist new columns immediately, even if we crash later

    eligible_rows = find_eligible_rows(ws, cv_col, gen_cv_col)
    print(f"{len(eligible_rows)} candidate(s) queued (Gen CV = Y, not yet processed).\n")

    if not eligible_rows:
        print("Nothing to generate. Done.")
        return

    profile_facts = load_profile_facts()
    generated = 0

    for row in eligible_rows:
        company = str(ws.cell(row=row, column=2).value or "").strip()
        title = str(ws.cell(row=row, column=3).value or "").strip()
        location = str(ws.cell(row=row, column=4).value or "").strip()
        score = ws.cell(row=row, column=6).value
        strengths = str(ws.cell(row=row, column=7).value or "").strip()
        gaps = str(ws.cell(row=row, column=8).value or "").strip()
        source = str(ws.cell(row=row, column=9).value or "").strip()
        url = str(ws.cell(row=row, column=10).value or "").strip()
        jd_snippet = str(ws.cell(row=row, column=11).value or "").strip()

        print(f"Processing: {title[:60]} @ {company[:30]} (score {score})")

        language = detect_language(title + " " + jd_snippet)
        print(f"  Language: {language}")

        candidate = {
            "title": title, "company": company, "location": location,
            "score": score, "strengths": strengths, "gaps": gaps,
            "source": f"{source} — {url}" if url else source,
            "jd_snippet": jd_snippet, "date": today_slash,
        }

        try:
            content = tailor_application(profile_facts, candidate, language)
        except Exception as e:
            print(f"  ERROR tailoring content: {e}")
            print("  Skipping this candidate (will retry next run).")
            continue

        folder_name = f"{sanitize_folder_name(company)}_{sanitize_folder_name(title)}"
        rel_folder = os.path.join(today_dash, folder_name)
        abs_folder = os.path.join(JOB_HUNT_ROOT, rel_folder)
        os.makedirs(abs_folder, exist_ok=True)

        try:
            # Hebrew fills the user's own Word CV template (correct RTL in Word);
            # English keeps render_docx.js. From-scratch docx bidi garbles Hebrew
            # in Word, so Hebrew reuses the template's proven paragraph structures.
            cv_docx = os.path.join(abs_folder, _CV_BASENAME + ".docx")
            cv_pdf = os.path.join(abs_folder, _CV_BASENAME + ".pdf")
            if language == "hebrew" and HEBREW_TEMPLATE:
                render_cv_from_template(content["cv"], cv_docx, template_path=HEBREW_TEMPLATE)
            elif language == "hebrew":
                print("  NOTE: no hebrew_cv_template configured — rendering English CV "
                      "for a Hebrew job (see docs/HEBREW_CV_TEMPLATE.md).")
                render_docx(content["cv"], cv_docx)
            else:
                render_docx(content["cv"], cv_docx)
            docx_to_pdf(cv_docx, cv_pdf)

            cl_docx = os.path.join(abs_folder, "Cover_Letter.docx")
            cl_pdf = os.path.join(abs_folder, "Cover_Letter.pdf")
            if language == "hebrew" and HEBREW_TEMPLATE:
                render_cover_from_template(content["cover"], cl_docx, template_path=HEBREW_TEMPLATE)
            else:
                render_docx(content["cover"], cl_docx)
            docx_to_pdf(cl_docx, cl_pdf)
        except Exception as e:
            print(f"  ERROR rendering documents: {e}")
            print("  Skipping this candidate (will retry next run).")
            continue

        ws.cell(row=row, column=cv_col, value=rel_folder.replace("\\", "/"))
        generated += 1
        print(f"  Done -> {rel_folder}")

    wb.save(TRACKER_PATH)
    print(f"\n{'='*60}")
    print(f"Done. Generated {generated} application package(s).")
    print(f"File: {TRACKER_PATH}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
