# The onboarding interview — what it asks and how it's framed

This is the script Claude follows on first run (after `install.bat` has
installed dependencies and scaffolded the config). Claude does every file
edit itself — config.json, profile.md, positioning.md — the user only answers. It's a conversation, not a form: one area at a time,
follow-ups only where answers are thin, and everything gets read back for
approval before it's saved. Budget 20–30 minutes.

The tone rule for Claude: **you are building the file that every future CV is
written from. Never accept a vague number, and never write down a number the
user didn't actually say.**

---

## Opening (sets the frame)

> "Before anything can run, I need to build two files with you: `profile.md` —
> the facts of your career, and `positioning.md` — how we tell them. Every CV
> and cover letter this pipeline ever writes comes from these two files, and I
> can't invent anything that isn't in them. So the better this conversation,
> the better every application. Easiest way to start: got a current CV in any
> language? Drop it here."

## 1. Their CV (if they have one — most do)

Claude reads it and pre-fills, then confirms instead of interrogating:

> "I've pulled out what I can. Let me read the numbers back, because these
> have to be exactly right: team of 40 at Delta Foods, 18% fewer picking
> errors, three shifts of 12 at Rimon. Confirm each, correct anything, and
> tell me if any of these is an estimate rather than a number you'd defend in
> an interview."

## 2. Personal basics

> "How should your name appear on documents — English and Hebrew spellings?
> Which email and phone go on the CV? LinkedIn URL?"

(Claude never asks for passwords, ID numbers, or payment details — and says
so if the user offers them.)

## 3. What the CV doesn't say

Per role, one question:

> "At [company] — what did you BUILD? A system you implemented, a process
> that didn't exist before you, a team you stood up. That's what leads every
> bullet we write."

## 4. Target positions

> "Which titles are you hunting? Give me the real list, including stretch
> titles and acceptable pivots. These become the crawler searches, so a title
> you don't list is a job you'll never see.
>
> One more thing this feeds: Jobify360 sorts its own listings into role tabs,
> and I need the Hebrew AND English words that would appear on YOUR tabs —
> not just your exact title. If you're outside logistics/operations/supply
> chain, the default keyword list won't catch your tabs at all, so tell me
> now or Jobify will quietly skip half your relevant postings."

## 5. Market scope (Israel)

> "Where do you live, and what's your honest commute limit? Any cities or
> regions to exclude outright? Remote/hybrid preferences? Industries you
> won't touch?"

## 5b. Score thresholds

> "Two numbers decide what you actually see. `min_score` (default 65) is the
> floor — anything scoring below it never lands in your tracker at all.
> `generate_score_floor` (default 70) is higher: a job has to clear this
> before I'll write it a full CV and cover letter, even if it's in your
> tracker. Keep the defaults if you're not sure — they came from someone who
> tuned them after seeing what 'too much noise' looks like. Want them
> stricter or looser?"

## 6. Compensation floor

> "What's the monthly gross below which you'd decline? This never appears in
> any document — it's only used to judge fit — so give me the real number,
> not the negotiating number."

## 7. The hard question

> "Every career has one thing an interviewer will poke at — a gap, short
> stints, a sector change, a step down. What's yours? We're going to write
> the honest, confident version of that answer now, once, so it's baked into
> every cover letter and you never improvise it in an interview."

## 8. Hebrew CV template

> "Do you have a Hebrew CV as a Word file? If yes, I'll wire it in — the
> pipeline fills YOUR file, keeping your design, because that's the only way
> Hebrew reliably survives Word. If not, Hebrew-language jobs get an English
> package, which most Israeli employers accept.
>
> By default, the language of the JOB POSTING decides the language of your
> CV — Hebrew posting gets a Hebrew CV, English gets English. Do you want
> that, or should every application go out in one language regardless of how
> the posting was written?"

## 9. API key

> "Scoring and CV tailoring run on your own Anthropic API key. Is
> ANTHROPIC_API_KEY set as an environment variable? If not, I'll walk you
> through it — it never gets written into a file."

## Close (mandatory)

Claude reads back BOTH the finished files AND the config settings this
conversation set — not just profile/positioning. Half the answers from this
interview (titles, keywords, exclusions, thresholds, template path, language
default) live in config.json, and none of them get a second look if the close
only covers the two markdown files.

> "This is profile.md — the facts I'll never deviate from. This is
> positioning.md — how we frame them. Anything wrong in these files will be
> wrong in every application, so read them like a proofreader, not like a
> form. And here's what went into config.json: target titles, Jobify
> keywords, excluded locations, score thresholds, Hebrew template, language
> default — one line each, so you can catch anything I misheard. Sign off,
> or fix now."
