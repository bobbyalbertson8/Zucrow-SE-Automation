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

/**
 * COMPLETE SYSTEM TEST - Simulates the entire onEdit workflow
 * This is what you want to test the whole system!
 */
function testCompleteSystem() {
  console.log('=== COMPLETE SYSTEM TEST ===');
  console.log('This will simulate editing the Order Placed? column and trigger the full workflow');
  
  try {
    const cfg = getConfig_();
    const sheet = getActiveSheet_(cfg);
    const headerMap = getHeaderMap_(sheet);
    
    if (sheet.getLastRow() < 2) {
      console.error('No data rows found. Please add some test data first.');
      return false;
    }
    
    // Find a test row (look for one that hasn't been notified yet)
    let testRow = -1;
    const notifiedCol = headerMap[cfg.NOTIFIED_HEADER] || ensureHeader_(sheet, cfg.NOTIFIED_HEADER);
    
    for (let i = 2; i <= sheet.getLastRow(); i++) {
      const notifiedValue = sheet.getRange(i, notifiedCol).getValue();
      if (!notifiedValue || !String(notifiedValue).trim()) {
        testRow = i;
        break;
      }
    }
    
    if (testRow === -1) {
      console.log('All rows already notified. Creating a new test row...');
      testRow = sheet.getLastRow() + 1;
      
      // Add test data
      const testRowData = [
        nowIso_(), // Timestamp
        'Test User', // Name
        'test@example.com', // Email - CHANGE THIS TO YOUR EMAIL FOR TESTING
        'Test Purchase Order Description',
        'TEST-' + Date.now(),
        'https://example.com/quote.pdf',
        '150.00',
        '', // Order Placed? - leave blank initially
        '' // Notified - leave blank
      ];
      
      sheet.getRange(testRow, 1, 1, testRowData.length).setValues([testRowData]);
      console.log('Test row created at row', testRow);
    }
    
    console.log('Using test row:', testRow);
    
    // Get current data
    const rowObj = getRowObject_(sheet, testRow, headerMap);
    console.log('Test row data:', rowObj);
    
    const email = rowObj[cfg.EMAIL_HEADER];
    if (!email || email.indexOf('@') < 0) {
      console.error('Invalid email in test row. Please update the email in row', testRow, 'to a valid email address.');
      return false;
    }
    
    console.log('Test email recipient:', email);
    console.log('IMPORTANT: Make sure this email is one you can check!');
    
    // Simulate the onEdit event
    const statusCol = headerMap[cfg.STATUS_HEADER];
    console.log('Simulating edit on row', testRow, 'column', statusCol);
    
    // Create mock edit event
    const mockRange = {
      getRow: () => testRow,
      getColumn: () => statusCol,
      getNumRows: () => 1,
      getNumColumns: () => 1
    };
    
    const mockEvent = {
      range: mockRange,
      value: cfg.STATUS_ORDERED_VALUE
    };
    
    console.log('Mock event created with value:', cfg.STATUS_ORDERED_VALUE);
    
    // Set the status value in the sheet to trigger the email
    sheet.getRange(testRow, statusCol).setValue(cfg.STATUS_ORDERED_VALUE);
    console.log('Status cell updated to trigger value');
    
    // Call the handler directly with mock event
    console.log('Calling onEditHandler...');
    onEditHandler(mockEvent);
    
    // Check if notification was recorded
    const finalNotifiedValue = sheet.getRange(testRow, notifiedCol).getValue();
    
    console.log('=== TEST RESULTS ===');
    console.log('Row tested:', testRow);
    console.log('Email recipient:', email);
    console.log('Final notified value:', finalNotifiedValue);
    
    if (finalNotifiedValue && String(finalNotifiedValue).trim() && !String(finalNotifiedValue).includes('ERROR')) {
      console.log('✅ SUCCESS: Email sent and notification recorded');
      console.log('Check the email inbox for:', email);
      return true;
    } else if (String(finalNotifiedValue).includes('ERROR')) {
      console.log('❌ ERROR: Email failed -', finalNotifiedValue);
      return false;
    } else {
      console.log('⚠️  WARNING: No notification timestamp recorded');
      return false;
    }
    
  } catch (error) {
    console.error('❌ SYSTEM TEST FAILED:', error.toString());
    return false;
  }
}

/**
 * Advanced system test with menu integration and UI
 */
function runCompleteSystemTest() {
  console.log('Starting advanced system test...');
  
  try {
    // Run configuration test first
    const configResult = testConfiguration();
    if (!configResult) {
      throw new Error('Configuration test failed. Check the logs and fix issues before proceeding.');
    }
    
    // Ask user for confirmation
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Complete System Test', 
      'This will:\n\n' +
      '1. Test your configuration\n' +
      '2. Create or find a test row\n' +
      '3. Simulate the onEdit trigger\n' +
      '4. Send a REAL email\n' +
      '5. Update tracking columns\n\n' +
      'Make sure the test email address is one you can access!\n\n' +
      'Continue with complete system test?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      console.log('System test cancelled by user');
      return;
    }
    
    // Run the complete test
    const testResult = testCompleteSystem();
    
    // Show results to user
    const resultMessage = testResult ? 
      '✅ SYSTEM TEST PASSED!\n\n' +
      'The complete workflow executed successfully:\n' +
      '• Configuration validated\n' +
      '• Test data prepared\n' +
      '• Email sent successfully\n' +
      '• Notification recorded\n\n' +
      'Check the recipient\'s email to verify the message was received.\n' +
      'Check the Execution Transcript for detailed logs.' :
      
      '❌ SYSTEM TEST FAILED\n\n' +
      'The test encountered errors. Please:\n' +
      '1. Check the Execution Transcript for error details\n' +
      '2. Verify your configuration settings\n' +
      '3. Ensure email addresses are valid\n' +
      '4. Check that all required columns exist\n\n' +
      'Run testConfiguration() first to diagnose issues.';
    
    ui.alert('System Test Results', resultMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Advanced system test failed:', error.toString());
    SpreadsheetApp.getUi().alert(
      'Test Failed', 
      'System test failed with error:\n\n' + error.toString() + '\n\nCheck the Execution Transcript for details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Create test menu for easy access to testing functions
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧪 Purchase Order Tests')
    .addItem('🔍 Test Configuration', 'testConfiguration')
    .addItem('📧 Test Email Send', 'testSendEmail')
    .addSeparator()
    .addItem('🚀 Complete System Test', 'runCompleteSystemTest')
    .addItem('🔄 Manual Complete Test', 'testCompleteSystem')
    .addSeparator()
    .addItem('⚙️ Setup Script Properties', 'setupScriptPropertiesMenu')
    .addItem('📊 View Current Config', 'viewCurrentConfigMenu')
    .addToUi();
}

/**
 * Menu wrapper for setup script properties
 */
function setupScriptPropertiesMenu() {
  setupScriptProperties();
  SpreadsheetApp.getUi().alert(
    'Setup Complete', 
    'Script properties have been configured for your sheet structure.\n\nYou can now run tests or use the system normally.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Menu wrapper for viewing current configuration
 */
function viewCurrentConfigMenu() {
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  const message = 
    'CURRENT CONFIGURATION:\n\n' +
    'Sheet Name: ' + (cfg.SHEET_NAME || 'First sheet (auto)') + '\n' +
    'Status Header: "' + cfg.STATUS_HEADER + '"\n' +
    'Trigger Value: "' + cfg.STATUS_ORDERED_VALUE + '"\n' +
    'Email Header: "' + cfg.EMAIL_HEADER + '"\n' +
    'Name Header: "' + cfg.NAME_HEADER + '"\n' +
    'Cost Header: "' + cfg.COST_HEADER + '"\n' +
    'Notified Header: "' + cfg.NOTIFIED_HEADER + '"\n\n' +
    'SHEET ANALYSIS:\n' +
    'Active Sheet: "' + sheet.getName() + '"\n' +
    'Headers Found: ' + Object.keys(headerMap).length + '\n' +
    'Data Rows: ' + Math.max(0, sheet.getLastRow() - 1) + '\n\n' +
    'COLUMN MAPPINGS:\n' +
    'Status Column: ' + (headerMap[cfg.STATUS_HEADER] || 'NOT FOUND') + '\n' +
    'Email Column: ' + (headerMap[cfg.EMAIL_HEADER] || 'NOT FOUND') + '\n' +
    'Name Column: ' + (headerMap[cfg.NAME_HEADER] || 'NOT FOUND') + '\n' +
    'Cost Column: ' + (headerMap[cfg.COST_HEADER] || 'NOT FOUND');
  
  SpreadsheetApp.getUi().alert('Current Configuration', message, SpreadsheetApp.getUi().ButtonSet.OK);
}
