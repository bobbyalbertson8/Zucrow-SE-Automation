# Order Notifier (Split by Functionality)

This is a split, maintainable Apps Script project that sends a **branded "Order Placed" email (with Cost)** to the
**requester email in the sheet row** when the status/ordered column is changed to an "ordered" value (e.g., "yes").

## Structure
```
.zucrow-se-order-notifier-split/
  .clasp.json            ← set your Script ID
  src/
    appsscript.json      ← manifest (minimal scopes)
    CONFIG.gs            ← reads Script Properties (headers, logos, recipients)
    UTIL.gs              ← helpers (headers map, currency, logos)
    EMAIL.gs             ← builds HTML/Plain for "Order Placed" (includes Cost + logos)
    STATUS.gs            ← onEdit trigger to detect ordered state and send email
```

## Deploy
1. Copy your **Script ID** from Apps Script editor → Project Settings.
2. Update `.clasp.json` with that ID.
3. `clasp login` → `clasp push` from the repo root (this folder).
4. In Apps Script: set **Script properties** (see below).
5. Add an installable trigger: **onEdit** → From spreadsheet → On edit.

## Script Properties (configure to match your sheet)
- `SHEET_ID` (blank if this is **bound** to your responses sheet)
- `SHEET_NAME` (optional specific tab)

- `STATUS_HEADER` (default: `Ordered`)
- `ORDERED_VALUES` (default: `yes,ordered,order placed,true`)
- `EMAIL_HEADER` (default: `Email`)           ← requester email column
- `ITEM_HEADER` (default: `Item`)
- `VENDOR_HEADER` (default: `Vendor`)
- `LINK_HEADER` (default: `Link`)
- `JUSTIFICATION_HEADER` (default: `Justification`)
- `COST_HEADER` (default: `Cost`)
- `ORDER_NUMBER_HEADER` (default: `Order Number`)

- `LOGO_STRATEGY` (`primary|secondary|both|none`)
- `LOGO_URL`, `LOGO_ALT`, `LOGO_MAX_WIDTH`, `LOGO_MAX_HEIGHT`
- `LOGO2_URL`, `LOGO2_ALT`

- `NOTIFY_TO` (optional BCC list, comma-separated)
- `CC_REQUESTER` (`true|false`)

## Behavior
- When a cell in the `STATUS_HEADER` column is edited to any value in `ORDERED_VALUES`, the script:
  - Reads the row’s **Email** (per `EMAIL_HEADER`),
  - Builds a clean HTML email including **Cost** (formatted) and logos,
  - Sends to the requester. Optionally BCCs your admin list (`NOTIFY_TO`) and CCs the requester.

> Note: For bulk paste edits, Apps Script does not provide old values. The notifier will send for any rows whose new status equals one of the ordered values.
