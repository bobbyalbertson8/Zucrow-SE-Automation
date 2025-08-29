/**
 * Build the "Order Placed" mail (includes Cost, logos, etc.).
 * Reads the requester email from the row using the configured EMAIL_HEADER.
 */
function buildOrderPlacedEmail_(headers, row, cfg) {
  const m = headersToIndexMap_(headers);
  const item = getCell_(row, m[cfg.itemHeader.toLowerCase()]) || getCell_(row, m['item name']) || getCell_(row, m['product']);
  const vendor = getCell_(row, m[cfg.vendorHeader.toLowerCase()]) || getCell_(row, m['supplier']);
  const link = getCell_(row, m[cfg.linkHeader.toLowerCase()]) || getCell_(row, m['url']);
  const justification = getCell_(row, m[cfg.justificationHeader.toLowerCase()]) || getCell_(row, m['reason']);
  const costRaw = getCell_(row, m[cfg.costHeader.toLowerCase()]) || getCell_(row, m['price']) || getCell_(row, m['amount']);
  const cost = formatCurrency_(costRaw);
  const orderNum = getCell_(row, m[cfg.orderNumberHeader.toLowerCase()]);

  const subject = ['Order Placed', item].filter(Boolean).join(' - ');

  let html = '<div style="font-family:-apple-system,BlinkMacSystemFont,\\'Segoe UI\\',Roboto,sans-serif;max-width:650px;margin:0 auto;padding:16px;background:#ffffff;border-radius:8px">';
  html += buildLogoHtml_(cfg.logo);
  html += '<h2 style="margin:0 0 12px">Order Placed</h2>';
  html += '<table cellpadding="6" cellspacing="0" style="border-collapse:collapse">';
  if (item) html += `<tr><td><b>Item</b></td><td>${item}</td></tr>`;
  if (vendor) html += `<tr><td><b>Vendor</b></td><td>${vendor}</td></tr>`;
  if (costRaw !== '') html += `<tr><td><b>Cost</b></td><td>${cost}</td></tr>`;
  if (orderNum) html += `<tr><td><b>Order #</b></td><td>${orderNum}</td></tr>`;
  if (link) html += `<tr><td><b>Link</b></td><td><a href="${link}">${link}</a></td></tr>`;
  if (justification) html += `<tr><td><b>Justification</b></td><td>${justification}</td></tr>`;
  html += '</table>';
  html += '<p style="margin-top:16px">You will receive a follow-up when shipping details are available.</p>';
  html += '</div>';

  const plain = [
    'Order Placed',
    item ? `Item: ${item}` : '',
    vendor ? `Vendor: ${vendor}` : '',
    (costRaw !== '') ? `Cost: ${cost}` : '',
    orderNum ? `Order #: ${orderNum}` : '',
    link ? `Link: ${link}` : '',
    justification ? `Justification: ${justification}` : ''
  ].filter(Boolean).join('\\n');

  return { subject, html, plain };
}
