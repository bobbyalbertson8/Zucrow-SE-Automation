
/**
 * Main: installable onEdit trigger handler
 */

function onEdit(e) { onEditHandler(e); }

function onEditHandler(e) {
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);

  // Ensure required headers exist
  const must = [cfg.STATUS_HEADER, cfg.EMAIL_HEADER];
  must.forEach(h => { if (!headerMap[h]) throw new Error('Missing header: ' + h); });

  const statusCol = headerMap[cfg.STATUS_HEADER];
  const emailCol = headerMap[cfg.EMAIL_HEADER];
  const notifiedCol = headerMap[cfg.NOTIFIED_HEADER] || ensureHeader_(sheet, cfg.NOTIFIED_HEADER);

  if (!e || !e.range || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 1) return;         // header
  if (col !== statusCol) return; // only when status is edited

  const rowObj = getRowObject_(sheet, row, headerMap);
  const newStatus = e.value;
  if (!newStatus || String(newStatus).trim() !== String(cfg.STATUS_ORDERED_VALUE).trim()) return;

  // Dedupe
  const notifiedCell = sheet.getRange(row, notifiedCol).getValue();
  if (notifiedCell && String(notifiedCell).trim()) return;

  const recipient = rowObj[cfg.EMAIL_HEADER];
  if (!recipient || String(recipient).indexOf('@') < 0) throw new Error('Invalid or missing email in row ' + row);

  const name = headerMap[cfg.NAME_HEADER] ? rowObj[cfg.NAME_HEADER] : '';
  const cc = (cfg.CC_HEADER && headerMap[cfg.CC_HEADER]) ? rowObj[cfg.CC_HEADER] : '';

  const costRaw = headerMap[cfg.COST_HEADER] ? rowObj[cfg.COST_HEADER] : '';
  const costFmt = formatCurrency_(costRaw, cfg.CURRENCY || 'USD');

  const item = headerMap['Item'] ? rowObj['Item'] : '';
  const qty  = headerMap['Quantity'] ? rowObj['Quantity'] : '';
  const po   = headerMap['PO Number'] ? rowObj['PO Number'] : '';

  const data = {
    name: name || '',
    recipient: recipient,
    cc: cc || '',
    status: newStatus,
    cost: costFmt,
    costRaw: costRaw,
    item: item,
    quantity: qty,
    poNumber: po,
    rowNumber: row,
    timestamp: nowIso_()
  };

  const subject = 'Purchase Order Placed' + (po ? (' – ' + po) : '');
  const htmlBody = buildEmailHtml_(data, cfg);
  sendNotificationEmail_(recipient, subject, htmlBody, cc);

  setCell_(sheet, row, notifiedCol, nowIso_());
}

function ensureHeader_(sheet, headerName) {
  const lastCol = sheet.getLastColumn() || 1;
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let c = 0; c < header.length; c++) {
    if (String(header[c]).trim() === headerName) return c + 1;
  }
  sheet.getRange(1, lastCol + 1).setValue(headerName);
  return lastCol + 1;
}
