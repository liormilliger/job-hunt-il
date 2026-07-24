# Getting an Anthropic API key (5 minutes) · השגת מפתח API של Anthropic

The pipeline scores jobs and writes CVs using Claude through *your own* API
key. It's pay-per-use and cheap at this scale: scoring a job costs fractions
of an agora, a full tailored CV + cover letter package costs a few agorot.
A typical week of active job hunting costs less than a cup of coffee.

## English

1. **Create an account** at [console.anthropic.com](https://console.anthropic.com)
   (this is the developer console — separate from a claude.ai chat account).
2. **Add credit**: Settings → Billing → add a payment method and buy credits.
   The minimum purchase (usually $5) will last you a long time here.
3. **Set a spend limit** while you're there (Settings → Limits). Pick
   something like $10/month. If a crawl run ever goes wild, this is the wall
   it hits — the pipeline is built to fail loudly and recover when that
   happens, but the wall is still worth having.
4. **Create the key**: API Keys → Create Key. Name it `job-hunt-il`. Copy it
   immediately — it's shown only once.
5. **Store it as an environment variable** (never in a file):
   - macOS/Linux: add a line to your shell profile (`~/.zshrc` for the
     default macOS shell, `~/.bashrc` on most Linux setups):
     `export ANTHROPIC_API_KEY="sk-ant-..."`, then `source ~/.zshrc` (or open
     a new terminal).
   - Windows: Start → search "environment variables" → "Edit environment
     variables for your account" → New → Name: `ANTHROPIC_API_KEY`,
     Value: your key → OK.
   - Open a **new** terminal afterward (already-open windows don't see it).
6. Done. The pipeline reads the key from the environment; it is never
   written to any file, and `config.json` never contains it.

## עברית

1. **פותחים חשבון** ב-[console.anthropic.com](https://console.anthropic.com)
   (זו קונסולת המפתחים — נפרדת מחשבון הצ'אט של claude.ai).
2. **טוענים קרדיט**: Settings → Billing → מוסיפים אמצעי תשלום וקונים קרדיט.
   הרכישה המינימלית (בדרך כלל 5$) תספיק לכם להרבה זמן.
3. **מגדירים תקרת הוצאה** באותו מקום (Settings → Limits). משהו כמו 10$
   לחודש. אם ריצה תשתולל, זה הקיר שהיא תפגוש — המערכת בנויה להיכשל
   בקול רם ולהתאושש, אבל שיהיה קיר.
4. **יוצרים מפתח**: API Keys → Create Key. תנו לו שם `job-hunt-il` והעתיקו
   מיד — הוא מוצג פעם אחת בלבד.
5. **שומרים אותו כמשתנה סביבה** (אף פעם לא בקובץ):
   - macOS/Linux: מוסיפים שורה לקובץ הפרופיל של השל (`~/.zshrc` ב-macOS,
     `~/.bashrc` ברוב הלינוקסים): `export ANTHROPIC_API_KEY="sk-ant-..."`,
     ואז `source ~/.zshrc` (או פותחים טרמינל חדש).
   - Windows: התחל → מחפשים "environment variables" → "עריכת משתני סביבה
     עבור החשבון שלך" → New → שם: `ANTHROPIC_API_KEY`, ערך: המפתח → OK.
   - פותחים טרמינל **חדש** אחר כך (חלונות פתוחים לא רואים את השינוי).
6. זהו. המערכת קוראת את המפתח מהסביבה בלבד — הוא לא נכתב לשום קובץ,
   ו-config.json לעולם לא מכיל אותו.

## Cost reality check · בדיקת מציאות על עלויות

| Action | Model | Rough cost |
|---|---|---|
| Score one job | Haiku | < 0.05 agorot |
| Full crawl run (50 jobs scored) | Haiku | ~ 1–2 agorot |
| One tailored CV + cover package | Sonnet | ~ 10–20 agorot |

If your credit runs out mid-crawl, nothing breaks silently: failed rows are
flagged "SCORING FAILED" in the tracker, and one command
(`clean_failed_run.py`) removes them so the next crawl re-scores them.
