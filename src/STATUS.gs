/**
 * Installable trigger: From spreadsheet → On edit
 * When status/ordered column changes to a configured ordered value (e.g., "yes"),
 * notify the requester email listed in the row with a branded message that includes Cost.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const cfg = getCfg();

    // Establish sheet
    const ss = cfg.sheetId ? SpreadsheetApp.openById(cfg.sheetId) : SpreadsheetApp.getActiveSpreadsheet();
    const sh = cfg.sheetName ? ss.getSheetByName(cfg.sheetName) : e.range.getSheet();

    // Headers
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

    const statusIdx = findHeaderIndex_(headers, cfg.statusHeader);
    const emailIdx = findHeaderIndex_(headers, cfg.emailHeader);
    if (statusIdx === -1 || emailIdx === -1) return;

    const editedRow = e.range.getRow();
    const editedCol = e.range.getColumn();
    if (editedRow <= 1) return; // skip header row
    if (editedCol !== statusIdx + 1) return; // only react to edits in status column

    // New value → check if it's one of our "ordered" values
    const newVal = String(e.range.getValue() || '').trim().toLowerCase();
    const ordered = cfg.orderedValues.map(v => v.toLowerCase());
    if (!ordered.includes(newVal)) return;

    // Grab full row values
    const row = sh.getRange(editedRow, 1, 1, lastCol).getValues()[0];
    const requesterEmail = String(getCell_(row, emailIdx)).trim();
    if (!requesterEmail) return;

    // Build email
    const msg = buildOrderPlacedEmail_(headers, row, cfg);

    // Compose recipients
    const to = requesterEmail;
    const options = { name: 'Purchase Notifier', htmlBody: msg.html };
    if (cfg.ccRequester && requesterEmail) options.cc = requesterEmail; // optional self-cc
    if (cfg.notifyTo.length) options.bcc = cfg.notifyTo.join(',');

    MailApp.sendEmail(to, msg.subject, msg.plain, options);
  } catch (err) {
    console.error(err);
  }
}
