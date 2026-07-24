# Connecting Google Sheets + Docs (5 minutes) · חיבור ל-Google Sheets ו-Docs

The tracker is a Google Sheet, and generated CVs/cover letters are converted
to PDF through Google Drive — both live in *your own* Google account (not a
shared or third-party one), the same way the Anthropic key is yours alone.
This is a one-time setup.

## English

1. **Create (or reuse) a Google Cloud project**: go to
   [console.cloud.google.com](https://console.cloud.google.com), and either
   pick an existing project or create a new one (top-left project picker →
   "New Project"). Name doesn't matter — e.g. "job-hunt-il".
2. **Enable two APIs** for that project: search the top bar for "Google
   Sheets API" → Enable, then search for "Google Drive API" → Enable.
3. **Configure the OAuth consent screen** (APIs & Services → OAuth consent
   screen): choose **External**, fill in an app name (e.g. "job-hunt-il") and
   your email for the required fields, and save through the remaining steps
   with defaults. Since this is a personal tool, you don't need to submit it
   for Google verification — add yourself under "Test users" so Google lets
   you sign in without a warning.
4. **Create an OAuth client ID** (APIs & Services → Credentials → Create
   Credentials → OAuth client ID): Application type = **Desktop app**. Name
   it anything. Click Create, then **Download JSON**.
5. **Save the downloaded file** as `credentials.json` in your workdir (the
   same folder as your tracker and profile.md — the path in config.json's
   `workdir`). This matches config.json's default `google_credentials` value;
   if you name or place it differently, update that field.
6. **First run**: the first script that touches the tracker (`setup.py`,
   any crawler, `generate_applications.py`, `sync_tracker.py`, `brief.py`)
   will open a browser window asking you to sign in and approve access.
   Approve it — you'll see a warning that the app isn't verified by Google
   (normal for a personal-use OAuth client); click through via "Advanced" →
   "Go to job-hunt-il (unsafe)". After that, a `token.json` is cached in your
   workdir and every future run is silent, until the token is revoked or you
   delete it.
7. Done. Everything the pipeline creates (the tracker Sheet, temporary Docs
   used for PDF export) lives in your own Drive under your own account —
   nothing is shared with or owned by a separate service account.

### What access does this actually grant?

Two scopes only: edit access to Sheets, and `drive.file` (files this app
creates or you explicitly open with it) — not blanket access to your whole
Drive. You can review or revoke access anytime at
[myaccount.google.com/permissions](https://myaccount.google.com/permissions).

### One thing worth checking

PDF export for tailored CVs/cover letters now goes through Google Drive's
docx→PDF conversion instead of Microsoft Word. It's generally faithful, but
if you use a Hebrew CV template (`docs/HEBREW_CV_TEMPLATE.md`), open the PDF
of the first one this pipeline generates and confirm the RTL layout still
looks right — Drive's conversion isn't guaranteed pixel-identical to Word's.

## עברית

1. **יוצרים (או משתמשים בקיים) פרויקט ב-Google Cloud**: נכנסים ל-
   [console.cloud.google.com](https://console.cloud.google.com), ובוחרים
   פרויקט קיים או יוצרים חדש. השם לא משנה — למשל "job-hunt-il".
2. **מפעילים שני APIs** לפרויקט: מחפשים "Google Sheets API" → Enable, ואז
   "Google Drive API" → Enable.
3. **מגדירים את מסך הסכמת OAuth** (APIs & Services → OAuth consent screen):
   בוחרים **External**, ממלאים שם אפליקציה והאימייל שלכם, וממשיכים עם ברירות
   המחדל. מוסיפים את עצמכם תחת "Test users" כדי לא לקבל אזהרת אימות.
4. **יוצרים OAuth client ID** (APIs & Services → Credentials → Create
   Credentials → OAuth client ID): סוג אפליקציה = **Desktop app**. יוצרים,
   ואז **מורידים את קובץ ה-JSON**.
5. **שומרים את הקובץ** בשם `credentials.json` בתיקיית העבודה שלכם (אותה
   תיקייה של הטראקר ו-profile.md — הנתיב ב-`workdir` בתוך config.json).
6. **הרצה ראשונה**: הסקריפט הראשון שנוגע בטראקר יפתח דפדפן לאישור גישה.
   מאשרים (אזהרת "אפליקציה לא מאומתת" תופיע — זה תקין לשימוש אישי) דרך
   "Advanced" → "Go to job-hunt-il (unsafe)". אחרי זה נוצר `token.json`
   ושמור, וכל הרצה הבאה שקטה.
7. זהו. כל מה שהמערכת יוצרת (הטראקר, מסמכים זמניים להמרת PDF) נמצא בחשבון
   Google שלכם בלבד.
