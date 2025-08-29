# Purchase Order Notifier (Apps Script)

This repo contains your Apps Script project split into modules for GitHub.

## Structure
```
purchase-order-notifier/
├─ .clasp.json
├─ src/
│  ├─ appsscript.json
│  ├─ index.gs
│  ├─ config/config.gs
│  ├─ github/repo.gs
│  ├─ email/buildEmailContent.gs
│  ├─ email/sendAdvancedGmail.gs
│  ├─ menu/index.gs
│  ├─ sheet/mapping.gs
│  ├─ sheet/duplicates.gs
│  ├─ sheet/rate_limit.gs
│  └─ util/utils.gs
└─ assets/
```

## Quick start
1. Install clasp: `npm i -g @google/clasp`
2. `clasp login`
3. Set Script ID in `.clasp.json` (Apps Script → Project Settings).
4. `clasp push`
5. In Apps Script, enable **Gmail API** (Services → + → Gmail API).
6. Add an **installable** trigger for `onEdit`.
