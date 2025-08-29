/**
 * Main: installable onEdit trigger handler
 * Updated to work with your specific sheet headers
 */

function onEdit(e) { 
  try {
    onEditHandler(e); 
  } catch (error) {
    console.error('Error in onEdit:', error.toString());
    // Optional: Send error notification to admin
    // MailApp.sendEmail('admin@example.com', 'Script Error', error.toString());
  }
}

function onEditHandler(e) {
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);

  console.log('Headers found:', Object.keys(headerMap));
  console.log('Looking for status header:', cfg.STATUS_HEADER);
  console.log('Looking for email header:', cfg.EMAIL_HEADER);

  // Ensure required headers exist
  const must = [cfg.STATUS_HEADER, cfg.EMAIL_HEADER];
  must.forEach(h => { 
    if (!headerMap[h]) {
      console.error('Available headers:', Object.keys(headerMap));
      throw new Error('Missing header: ' + h + '. Available headers: ' + Object.keys(headerMap).join(', ')); 
    }
  });

  const statusCol = headerMap[cfg.STATUS_HEADER];
  const emailCol = headerMap[cfg.EMAIL_HEADER];
  const notifiedCol = headerMap[cfg.NOTIFIED_HEADER] || ensureHeader_(sheet, cfg.NOTIFIED_HEADER);

  // Validate edit event
  if (!e || !e.range) {
    console.log('No valid edit event');
    return;
  }
  
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) {
    console.log('Edit spans multiple cells, ignoring');
    return;
  }

  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  console.log('Edit detected - Row:', row, 'Col:', col, 'Status Col:', statusCol);
  
  if (row === 1) {
    console.log('Header row edited, ignoring');
    return;
  }
  
  if (col !== statusCol) {
    console.log('Non-status column edited, ignoring');
    return;
  }

  const rowObj = getRowObject_(sheet, row, headerMap);
  const newStatus = e.value;
  
  console.log('New status value:', newStatus);
  console.log('Expected trigger value:', cfg.STATUS_ORDERED_VALUE);
  
  if (!newStatus || String(newStatus).trim() !== String(cfg.STATUS_ORDERED_VALUE).trim()) {
    console.log('Status does not match trigger value, ignoring');
    return;
  }

  // Check if already notified (dedupe)
  const notifiedCell = sheet.getRange(row, notifiedCol).getValue();
  if (notifiedCell && String(notifiedCell).trim()) {
    console.log('Already notified, skipping');
    return;
  }

  const recipient = rowObj[cfg.EMAIL_HEADER];
  if (!recipient || String(recipient).indexOf('@') < 0) {
    throw new Error('Invalid or missing email in row ' + row + ': ' + recipient);
  }

  console.log('Preparing to send email to:', recipient);

  // Extract data from row
  const name = headerMap[cfg.NAME_HEADER] ? rowObj[cfg.NAME_HEADER] : '';
  const cc = (cfg.CC_HEADER && headerMap[cfg.CC_HEADER]) ? rowObj[cfg.CC_HEADER] : '';
  
  const costRaw = headerMap[cfg.COST_HEADER] ? rowObj[cfg.COST_HEADER] : '';
  const costFmt = formatCurrency_(costRaw, cfg.CURRENCY || 'USD');

  // Map additional fields that might exist in your sheet
  const item = headerMap['Purchase order description (program and pur'] ? 
               rowObj['Purchase order description (program and pur'] : 
               (headerMap['Item'] ? rowObj['Item'] : '');
  const qty = headerMap['Quantity'] ? rowObj['Quantity'] : '';
  const po = headerMap['Purchase Order Number'] ? rowObj['Purchase Order Number'] : '';
  const quotePdf = headerMap['Quote PDF'] ? rowObj['Quote PDF'] : '';

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
    quotePdf: quotePdf,
    rowNumber: row,
    timestamp: nowIso_()
  };

  console.log('Email data prepared:', data);

  const subject = 'Purchase Order Placed' + (po ? (' – ' + po) : '');
  const htmlBody = buildEmailHtml_(data, cfg);
  
  sendNotificationEmail_(recipient, subject, htmlBody, cc);
  console.log('Email sent successfully');

  // Mark as notified
  setCell_(sheet, row, notifiedCol, nowIso_());
  console.log('Notification timestamp recorded');
}

function ensureHeader_(sheet, headerName) {
  const lastCol = sheet.getLastColumn() || 1;
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Check if header already exists
  for (let c = 0; c < header.length; c++) {
    if (String(header[c]).trim() === headerName) {
      console.log('Header', headerName, 'found at column', c + 1);
      return c + 1;
    }
  }
  
  // Add new header
  console.log('Adding new header:', headerName, 'at column', lastCol + 1);
  sheet.getRange(1, lastCol + 1).setValue(headerName);
  return lastCol + 1;
}

/**
 * Test function to verify configuration and setup
 */
function testConfiguration() {
  console.log('=== Testing Configuration ===');
  
  try {
    const cfg = getConfig_();
    console.log('Configuration loaded:', cfg);
    
    const sheet = getActiveSheet_(cfg);
    console.log('Sheet found:', sheet.getName());
    
    const headerMap = getHeaderMap_(sheet);
    console.log('Headers mapped:', headerMap);
    
    // Test required headers
    console.log('Testing required headers...');
    console.log('STATUS_HEADER "' + cfg.STATUS_HEADER + '" found:', !!headerMap[cfg.STATUS_HEADER]);
    console.log('EMAIL_HEADER "' + cfg.EMAIL_HEADER + '" found:', !!headerMap[cfg.EMAIL_HEADER]);
    console.log('NAME_HEADER "' + cfg.NAME_HEADER + '" found:', !!headerMap[cfg.NAME_HEADER]);
    console.log('COST_HEADER "' + cfg.COST_HEADER + '" found:', !!headerMap[cfg.COST_HEADER]);
    
    // Test data from a sample row (row 2 if exists)
    if (sheet.getLastRow() >= 2) {
      console.log('Testing sample row 2...');
      const rowObj = getRowObject_(sheet, 2, headerMap);
      console.log('Sample row data:', rowObj);
      
      const email = rowObj[cfg.EMAIL_HEADER];
      console.log('Email validation:', email, email && email.indexOf('@') > 0 ? 'VALID' : 'INVALID');
    }
    
    console.log('=== Configuration Test Complete ===');
    return true;
    
  } catch (error) {
    console.error('Configuration test failed:', error.toString());
    return false;
  }
}

/**
 * Manual trigger function for testing email sending
 */
function testSendEmail() {
  console.log('=== Testing Email Send ===');
  
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  if (sheet.getLastRow() < 2) {
    console.error('No data rows found for testing');
    return;
  }
  
  // Use row 2 for testing
  const rowObj = getRowObject_(sheet, 2, headerMap);
  const recipient = rowObj[cfg.EMAIL_HEADER];
  
  if (!recipient || recipient.indexOf('@') < 0) {
    console.error('Invalid email in test row:', recipient);
    return;
  }
  
  const testData = {
    name: rowObj[cfg.NAME_HEADER] || 'Test User',
    recipient: recipient,
    cc: '',
    status: 'Test',
    cost: 'USD 100.00',
    costRaw: 100,
    item: 'Test Item',
    quantity: '1',
    poNumber: 'TEST-001',
    quotePdf: '',
    rowNumber: 2,
    timestamp: nowIso_()
  };
  
  const subject = 'Test Email - Purchase Order Placed';
  const htmlBody = buildEmailHtml_(testData, cfg);
  
  try {
    sendNotificationEmail_(recipient, subject, htmlBody, '');
    console.log('Test email sent successfully to:', recipient);
  } catch (error) {
    console.error('Failed to send test email:', error.toString());
  }
}
