function headersToIndexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h).trim().toLowerCase()] = i);
  return map;
}
function findHeaderIndex_(headers, name) {
  const target = String(name || '').trim().toLowerCase();
  if (!target) return -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === target) return i;
  }
  return -1;
}
function getCell_(row, idx) {
  if (!Array.isArray(row) || idx < 0 || idx >= row.length) return '';
  const v = row[idx];
  return (v === null || v === undefined) ? '' : v;
}
function formatCurrency_(val) {
  try {
    const n = (typeof val === 'number') ? val : Number(String(val).replace(/[^0-9.+-]/g, ''));
    if (!isFinite(n)) return String(val);
    return Utilities.formatString('$%,.2f', n);
  } catch (e) { return String(val); }
}
function buildLogoHtml_(logo) {
  if (!logo) return '';
  const one = logo.url ? `<img src="${logo.url}" alt="${logo.alt || 'Logo'}" style="max-width:${logo.maxWidth||'180px'};max-height:${logo.maxHeight||'80px'};height:auto;width:auto;display:inline-block;vertical-align:middle;">` : '';
  const two = logo.url2 ? `<img src="${logo.url2}" alt="${logo.alt2 || ''}" style="max-width:${logo.maxWidth||'180px'};max-height:${logo.maxHeight||'80px'};height:auto;width:auto;display:inline-block;vertical-align:middle;margin-left:12px;">` : '';
  let inner = '';
  switch (String(logo.strategy||'primary').toLowerCase()) {
    case 'both': inner = (one + two) || ''; break;
    case 'secondary': inner = two || one; break;
    case 'primary': default: inner = one || two; break;
    case 'none': inner = ''; break;
  }
  return inner ? `<div style="text-align:center;margin:0 0 16px">${inner}</div>` : '';
}
