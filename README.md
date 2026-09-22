# SimplyDoors — New Project Intake

The form customers fill in to start a project. Static page on GitHub Pages,
with a Google Apps Script backend that writes to a spreadsheet and emails
Adem.

- **Live form:** https://simplydoorsaa.github.io/Simplydoors-customer-intake/
- **Submissions land in:** the "SimplyDoors Project Intake" spreadsheet,
  `Submissions` tab, plus an email to adem@simplydoors.com with any photos
  attached.
- **Backend code:** `apps-script/Code.gs` (this file is the master copy — the
  Apps Script editor holds a deployed copy of it).

## How a submission travels

1. The customer taps Submit. The page saves the submission **on their own
   device first**, then sends it. They see "Got it" either way.
2. Anything that could not be sent — no signal, backend slow — is retried
   automatically the next time the page is open or the device is online.
   Each submission carries an id so a retry can't create a duplicate row.
3. The backend writes the row to the sheet, **then** emails. If the email
   fails, the row is still saved and its status column reads `failed`.

## Deploying a change to the backend

Editing `Code.gs` in the Apps Script editor changes nothing on its own, and
"New deployment" creates a **new address** that the form doesn't know about.

Always: **Deploy → Manage deployments → pencil icon on the live deployment →
Version: New version → Deploy.** Leave "Who has access" on **Anyone**.

"Only myself" is what broke the form for customers between 4 and 13 September
2026: it works in your own signed-in browser and sends everyone else to a
Google sign-in page.

## Checking it works

- Automatic: `.github/workflows/health.yml` runs every 30 minutes from
  GitHub's servers and fails (emailing you) if the page is down or the
  backend asks for a sign-in. See the repo's Actions tab.
- By hand, the way a customer sees it: open the live form on your phone with
  wifi off, submit a test, then delete the row from the sheet.

## Known limits

- Up to 3 files, 20 MB total, because the backend emails them and Gmail caps a
  message at about 25 MB.
- `APP_TOKEN` in the page is not a secret — it is visible to anyone who views
  the source. It stops accidental posts, not a determined one. Spam
  protection (a Turnstile or reCAPTCHA check) is still to be added.
- Photos live only in the notification email. They are not filed in Drive yet.
