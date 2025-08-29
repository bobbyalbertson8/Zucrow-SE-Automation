
# Purchase Order Notifier — GitHub Assets Edition

Send a polished HTML email when a row’s status becomes **Ordered**. Logos are hosted from your GitHub repo’s **assets/** folder—no Drive fiddling. All runtime values (headers, ordered value, currency, etc.) are set via **Script Properties** (no hard-coded emails). **Cost** is read from your sheet and shown in the email.

---

## How logos work (GitHub `assets/`)
1. In your GitHub repository, create an **assets/** folder at the repo root.
2. Upload two images (PNG, SVG, etc.). Example:
   - `assets/spectral_logo.png`
   - `assets/purdue_prop_logo.png`
3. Open `src/Config.gs` and set:
   ```js
   var GITHUB_CONFIG = {
     username: "your-github-username",
     repository: "your-repo-name",
     branch: "main",
     logo:  { filename: "spectral_logo.png",  altText: "Spectral Energies",  maxWidth: "200px", maxHeight: "100px" },
     logo2: { filename: "purdue_prop_logo.png", altText: "Purdue Propulsion", maxWidth: "200px", maxHeight: "100px" }
   };
   ```
4. Pick a strategy via Script Properties or in `DEFAULTS`:
   - `LOGO_STRATEGY` = `primary` | `secondary` | `both` | `conditional`

The script builds raw GitHub URLs like:
```
https://raw.githubusercontent.com/<user>/<repo>/<branch>/assets/<filename>
```

---

## Quick Setup (Apps Script UI)
1. Open your Google Sheet → **Extensions → Apps Script**.
2. Create files and paste contents from `src/*` and `src/email.html`. Also paste `appsscript.json` (Project settings → show manifest).
3. Set **Script Properties** (Project settings → Script properties):
   - `SHEET_NAME` (blank = first sheet)
   - `STATUS_HEADER` (default `Status`)
   - `STATUS_ORDERED_VALUE` (default `Ordered`)
   - `EMAIL_HEADER` (default `Requester Email`)
   - `NAME_HEADER` (optional, default `Requester Name`)
   - `COST_HEADER` (default `Cost`)
   - `NOTIFIED_HEADER` (default `Notified`)
   - `LOGO_STRATEGY` (`primary|secondary|both|conditional`, default `both`)
   - `CURRENCY` (default `USD`)
4. **Installable trigger**: Triggers (clock icon) → Add Trigger → Function `onEditHandler` → From spreadsheet → On edit → Save & authorize.
5. Edit a row’s `Status` to match `STATUS_ORDERED_VALUE` and verify the email.

---

## Setup with `clasp`
```bash
unzip purchase-notifier-github.zip
cd purchase-notifier-github
clasp login
clasp create --title "Purchase Notifier (GitHub Assets)" --type sheets
# Put returned scriptId into .clasp.json OR run: clasp pull
clasp push
clasp open
```
Then set **Script Properties** and add the **installable trigger** as above.

---

## Notes
- The email’s **Cost** field is formatted to currency when numeric.
- Dedupe: writes ISO timestamp to `Notified` column after sending.
- Conditional logos: when `LOGO_STRATEGY="conditional"`, secondary logo is used for keywords like `purdue`, `@purdue.edu`, `university`, `student` (editable in `DEFAULTS`).
- Uses `MailApp.sendEmail` (works with installable triggers).

---

## Troubleshooting
- No email? Ensure the edited column is your `STATUS_HEADER`, and the new value equals `STATUS_ORDERED_VALUE` exactly.
- Permissions? Use an **installable** onEdit trigger (simple triggers can’t send email).
- Logos not showing? Confirm repo is **public** (or GH raw URL is accessible) and filenames match exactly.

MIT License
