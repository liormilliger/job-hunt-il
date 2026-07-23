# Your Hebrew CV template — the rules

Hebrew CV generation works by **filling your own Word CV**, not by building a
document from scratch. This is a hard-won lesson: every from-scratch approach
to Hebrew/RTL in .docx (python-docx, any recipe) eventually produces a file
that looks fine in previews but garbles in real Microsoft Word — reversed
numbers, split words, wrong-side punctuation. The only structure that renders
reliably in Word is a file that was authored in Word. So the pipeline takes
YOUR CV and swaps only the text inside it.

## What the pipeline replaces
- The professional summary paragraph (tailored per job).
- The bullet lines under each role (tailored per job).

## What it keeps verbatim
- Your name, contact line, links.
- Role headers (company + dates), education, skills, section headings.
- All fonts, spacing, and paragraph settings — that's the point.

## Rules your .docx must follow

1. **Author it in Microsoft Word**, in Hebrew, as a normal RTL document.
   If it looks right in your Word, it qualifies.
2. **Structure**: contact info near the top (a line containing your email),
   one summary paragraph, then per-role sections: a **bold header line**
   containing the company name and dates, followed by bullet lines
   (Word list paragraphs) describing the role.
3. **Company names in headers must be recognizable** — the pipeline matches
   tailored content to roles by finding the company name in the header text.
   Use the same company names in your profile.md.
4. **Keep it to designs you'd send anyway** — no text boxes, tables, or
   multi-column layouts for the experience section (plain paragraphs +
   bullets). Headers/footers are fine.

## Setting it up
Put the file's path in `config.json` → `hebrew_cv_template`.

## Editing it later
Edit the file in Word whenever you like — the pipeline reads it live. One
warning: don't run a generation while the template file is open in Word with
AutoSave on; Word's re-save can strip the paragraph settings that make the
bidi work. Close it first.

## If you don't have a Hebrew CV
Leave `hebrew_cv_template` empty. Hebrew-language jobs will get an English
package and a note in the output. (Most Israeli employers accept English CVs,
but for Hebrew-first companies a Hebrew CV noticeably helps.)
