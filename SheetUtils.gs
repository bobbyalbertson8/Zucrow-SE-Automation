
/**
 * Sheet utilities
 */

function getActiveSheet_(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (cfg.SHEET_NAME && cfg.SHEET_NAME.trim()) {
    const sh = ss.getSheetByName(cfg.SHEET_NAME.trim());
    if (sh) return sh;
    throw new Error('SHEET_NAME not found: ' + cfg.SHEET_NAME);
  }
  return ss.getSheets()[0];
}

function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn() || 1;
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  header.forEach((name, idx) => {
    if (name && String(name).trim()) {
      map[String(name).trim()] = idx + 1;
    }
  });
  return map;
}

function getRowObject_(sheet, row, headerMap) {
  const lastCol = sheet.getLastColumn() || 1;
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const obj = {};
  Object.keys(headerMap).forEach(h => obj[h] = values[headerMap[h] - 1]);
  return obj;
}

function setCell_(sheet, row, col, value) {
  sheet.getRange(row, col).setValue(value);
}

function nowIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function formatCurrency_(num, currency) {
  try {
    if (typeof num === 'number') return Utilities.formatString('%s %s', currency, num.toFixed(2));
    const parsed = parseFloat(String(num).replace(/[^0-9.\-]/g, ''));
    if (!isNaN(parsed)) return Utilities.formatString('%s %s', currency, parsed.toFixed(2));
  } catch (e) {}
  return String(num);
}
