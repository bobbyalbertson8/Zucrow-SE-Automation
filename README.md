# Project Setup Guide

This guide shows a brand‑new user how to (1) grab the files from GitHub (even if they’re stored as a `.zip`), and (2) install and run the Google Apps Script inside a Google Sheet. It’s written for beginners and includes both a quick path and an advanced, CLI‑based path.

---

## Contents

* [What’s in this repo](#whats-in-this-repo)
* [Prerequisites](#prerequisites)
* [Getting the code from GitHub](#getting-the-code-from-github)

  * [Option A — Download as ZIP (easiest)](#option-a--download-as-zip-easiest)
  * [Option B — Clone with Git](#option-b--clone-with-git)
  * [Option C — If the repo contains a `.zip` of the source](#option-c--if-the-repo-contains-a-zip-of-the-source)
* [Installing the Apps Script into Google Sheets](#installing-the-apps-script-into-google-sheets)

  * [Method 1 — Copy/Paste (no tools)](#method-1--copypaste-no-tools)
  * [Method 2 — Using the Apps Script CLI (`clasp`)](#method-2--using-the-apps-script-cli-clasp)
* [Configuration](#configuration)

  * [Script properties](#script-properties)
  * [Email sender & Gmail API notes](#email-sender--gmail-api-notes)
  * [Permissions & scopes](#permissions--scopes)
* [Triggers (make it run automatically)](#triggers-make-it-run-automatically)
* [Testing](#testing)
* [Troubleshooting](#troubleshooting)
* [Folder structure](#folder-structure)
* [FAQ](#faq)

---

## What’s in this repo

* **Google Apps Script** source files (for Google Sheets / Forms automation). Typical use: send a notification email on Form submit or when a Sheet row changes, and include key details from the row.
* Optional **GitHub Action** snippet showing how to auto‑extract zips added to the repo (useful if you store script bundles as `.zip`).

> You do **not** need GitHub to *run* the script; GitHub is just where the code lives for version control.

---

## Prerequisites

* A Google account with access to **Google Drive**, **Google Sheets**, and **Google Apps Script**.
* If you plan to send emails programmatically, ensure your account can authorize the permissions (you’ll be prompted the first time you run the script).
* Optional (for advanced installs):

  * **Node.js** LTS
  * **npm** (comes with Node)
  * **clasp** (Apps Script CLI)

---

## Getting the code from GitHub

### Option A — Download as ZIP (easiest)

1. Open the repository page in your browser: `https://github.com/<your-username>/<your-repo>`
2. Click the green **Code** button → **Download ZIP**.
3. On your computer, extract the archive:

   * **Windows**: Right‑click the file → **Extract All…**
   * **macOS**/**Linux**: Double‑click the `.zip` or run `unzip file.zip` in Terminal.
4. You now have the project folder with all source files.

### Option B — Clone with Git

```bash
# Install Git first: https://git-scm.com/downloads
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
```

### Option C — If the repo contains a `.zip` of the source

> Sometimes the repo holds a `.zip` file that contains the actual source.

* **From the web UI:** click the `.zip` → **Download** (top‑right), then extract locally.
* **If you cloned:**

  ```bash
  cd <your-repo>
  unzip source-bundle.zip -d extracted_source
  ```

  Replace `source-bundle.zip` with the real filename.

---

## Installing the Apps Script into Google Sheets

You have two ways to get the code into Apps Script: quick **copy/paste** or the **CLI**.

### Method 1 — Copy/Paste (no tools)

1. Create or open the **Google Sheet** that will host the automation.
2. In the Sheet, go to **Extensions → Apps Script**. A new tab opens the Apps Script editor.
3. In the left file tree, click the **+** and create files that match those in this repo (e.g., `Code.gs`, `Email.gs`, `Config.gs`).
4. Open each file in your local project and **copy** its contents into the corresponding Apps Script file.
5. If this repo includes an `appsscript.json`, open **Project Settings** (gear icon) → enable **Show "appsscript.json" manifest file**. Create `appsscript.json` and paste in the provided manifest.
6. Click **Save**.

### Method 2 — Using the Apps Script CLI (`clasp`)

This is great if you’ll update the script often.

1. Install `clasp`:

   ```bash
   npm install -g @google/clasp
   ```
2. Log in:

   ```bash
   clasp login
   ```

   A browser window will open; select the Google account that owns the Sheet.
3. In your local project folder (where `appsscript.json` and `.gs` files live):

   ```bash
   clasp create --type sheets --title "My Sheet Automation"
   ```

   Or, if you already created a container‑bound project from the Sheet’s **Extensions → Apps Script**, open that editor, click **Project Settings** → copy the **Script ID**, then in your terminal:

   ```bash
   clasp clone <YOUR_SCRIPT_ID>
   ```
4. Push your local files to Apps Script:

   ```bash
   clasp push
   ```
5. Confirm the files appear in the Script Editor.

> **Tip:** With `clasp`, future changes are as simple as editing locally then running `clasp push`.

---

## Configuration

### Script properties

Many projects use **Script Properties** to avoid hard‑coding emails and settings.

1. In the Apps Script editor: **Project Settings → Script properties → Add property**.
2. Add keys/values you see referenced in `Config.gs` (examples):

   * `NOTIFY_TO` = `someone@example.com` (comma‑separate for multiple)
   * `CC_REQUESTOR` = `true` (if you want to CC the person from the form/row)
   * `ENV` = `prod` or `dev`
3. Click **Save**.

Inside code, read them with:

```javascript
const PROPS = PropertiesService.getScriptProperties();
const NOTIFY_TO = PROPS.getProperty('NOTIFY_TO');
```

### Email sender & Gmail API notes

* If your code uses `MailApp.sendEmail`, the **from** address is your account and requires basic Gmail permission. For many cases, `MailApp` is enough.
* If your code uses **`GmailApp`** or the **Gmail API** (advanced service), you’ll need to authorize additional scopes, and in some domains you must enable the API:

  1. In the script editor, click **Services** (puzzle‑piece icon) → **+** → add **Gmail API** if the code uses `Gmail.Users.*` (advanced service).
  2. Ensure the `appsscript.json` includes the needed **oauthScopes** (see below).
  3. On first run you’ll be prompted to authorize.

> If you saw an error like: *“You do not have permission to call gmail.users.settings.sendAs.list…”* you’re using the advanced Gmail API and must enable the service + have the right scopes and account permissions. If you don’t need sender aliases, prefer `MailApp` or `GmailApp` to simplify.

### Permissions & scopes

If this repo includes a manifest, it may look like this (example):

```json
{
  "timeZone": "America/Indiana/Indianapolis",
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": [
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly"
  ]
}
```

Adjust to match what your functions actually use. Fewer scopes = easier approval.

---

## Triggers (make it run automatically)

Open **Apps Script editor → Triggers (clock icon)** and add an **installable trigger**:

* **Form submit** (if linked to a Google Form):

  * Event source: **From form**
  * Event type: **On form submit**
  * Choose function: e.g., `onFormSubmit(e)`

* **Sheet edit** (react to changes in a specific tab/columns):

  * Event source: **From spreadsheet**
  * Event type: **On edit**
  * Choose function: e.g., `onEdit(e)`

* **Time‑driven** (scheduled runs): set **Type of time based trigger** and interval, then pick your main function.

> **Note:** Simple triggers (function names exactly `onEdit` / `onOpen` / `onFormSubmit`) have limited permissions. If you need Gmail/Drive access, use **installable triggers** created from the Triggers UI.

---

## Testing

1. In the Apps Script editor, select a function like `sendForRow_` and click **Run**.
2. The first time, you’ll see an **Authorization** prompt → grant permissions.
3. Watch **Executions** (left sidebar) for logs and errors.
4. If you’re handling Form submits, send a **test response** through the Form and verify the email.

---

## Troubleshooting

**❗ “Specified permissions are not sufficient to call MailApp.sendEmail.”**
You’re calling a service that needs authorization. Use an **installable trigger** (not simple trigger), and re‑authorize from the editor when prompted. Ensure the account actually has Gmail enabled.

**❗ “Gmail is not defined”**
Your code references `Gmail.*` (advanced service) but you didn’t enable it. In the editor, click **Services** → **+** → add **Gmail API**. Or change code to use `GmailApp`/`MailApp` if advanced endpoints aren’t required.

**❗ “You do not have permission to call gmail.users.settings.sendAs.list …”**
This is an **advanced Gmail API** endpoint. You need the service enabled, correct scopes in `appsscript.json`, and sometimes admin‑granted permissions on Workspace domains. If you only need to send mail, prefer `MailApp.sendEmail` (simpler) or `GmailApp.sendEmail`.

**❗ “Syntax error: Illegal character”**
This usually happens when you pasted smart quotes or copied from a formatted doc. Re‑paste as plain text. In VS Code, run **Convert to UTF‑8**, or in the Script Editor, delete and re‑type the quotes.

**❗ “TypeError: redeclaration of const CONFIG.”**
Make sure `const CONFIG = { ... }` appears **only once** in the entire project. If multiple files declare it, rename one (e.g., `CONFIG_EMAIL`) or export as properties. Avoid duplicating the same constants across files.

**❗ Emails not sending from the expected address**
`MailApp` sends from your account. Sending **as** an alias requires the alias in Gmail Settings and, for the advanced API, access to `sendAs`. Otherwise stick with your default account.

---

## Folder structure

A typical layout in this repo:

```
/ (repo root)
├─ src/
│  ├─ Code.gs            # Entry points: onFormSubmit(e), onEdit(e), runOnce()
│  ├─ Email.gs           # Email helpers build subject/body, inline HTML templates
│  ├─ Config.gs          # Single CONFIG object or property readers
│  └─ Utils.gs           # Parsing, validation, logging
├─ appsscript.json       # Manifest with oauthScopes, timeZone, etc.
├─ .clasp.json           # (optional) clasp config
├─ .github/workflows/
│  └─ extract-zip.yml    # (optional) auto‑extract zips pushed to repo
└─ README.md             # This file
```

---

## FAQ

**Q: Can GitHub “auto‑extract” a zip I upload?**
Not by default. You (or a GitHub Action) must extract it. See the included example workflow in `.github/workflows/extract-zip.yml`:

```yaml
name: Extract Zip
on:
  push:
    paths:
      - '*.zip'
jobs:
  unzip:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Extract zips
        run: |
          for f in *.zip; do
            unzip -o "$f" -d "${f%.zip}"
          done
      - name: Commit extracted files
        run: |
          git config user.name "github-actions"
          git config user.email "github-actions@github.com"
          git add .
          git commit -m "Auto-extract zip" || echo "No changes"
          git push
```

**Q: Do I have to use the advanced Gmail API?**
No. For most cases, `MailApp.sendEmail` or `GmailApp.sendEmail` is sufficient and simpler to authorize.

**Q: I’m on a Google Workspace domain with admin restrictions.**
Ask your admin to allow the required Gmail/Drive scopes, or avoid advanced endpoints. Installable triggers also run **as you**, so you must have permission to act.

---

### You’re done!

At this point, your script should be in the Google Sheet, configured, authorized, and either triggered on form submit or on edit. If something’s off, check **Executions** logs and the **Troubleshooting** section above.
