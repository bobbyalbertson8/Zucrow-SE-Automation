/**
 * Ultra-Robust Purchase Order System
 * Handles complete lifecycle: Submitted → Ordered → In Transit → Received
 * Includes comprehensive logo testing and configuration
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

/**
 * Form submission trigger handler
 */
function onFormSubmit(e) {
  try {
    onFormSubmitHandler(e);
  } catch (error) {
    console.error('Error in onFormSubmit:', error.toString());
    MailApp.sendEmail('admin@example.com', 'Form Submit Error', error.toString());
  }
}

function onFormSubmitHandler(e) {
  console.log('Form submission detected');
  
  const cfg = getConfig_();
  
  const purchasingEmail = cfg.PURCHASING_EMAIL;
  const purchasingCcEmail = cfg.PURCHASING_CC_EMAIL;
  
  if (!purchasingEmail) {
    console.error('No purchasing email configured. Please set PURCHASING_EMAIL in script properties.');
    return;
  }
  
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  const submittedRow = sheet.getLastRow();
  const rowObj = getRowObject_(sheet, submittedRow, headerMap);
  
  console.log('Processing form submission for row:', submittedRow);
  
  const data = buildEmailData_(rowObj, headerMap, cfg, submittedRow, 'Submitted');
  
  console.log('Purchasing team notification data prepared:', data);
  
  const subject = 'New Purchase Request Submitted' + (data.name ? (' from ' + data.name) : '');
  const htmlBody = buildPurchasingTeamEmailHtml_(data, cfg);
  
  try {
    sendNotificationEmail_(purchasingEmail, subject, htmlBody, purchasingCcEmail);
    console.log('Purchasing team notification sent successfully');
  } catch (emailError) {
    console.error('Failed to send purchasing team notification:', emailError.toString());
  }
}

/**
 * Enhanced onEdit handler with multiple status triggers
 */
function onEditHandler(e) {
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);

  console.log('Headers found:', Object.keys(headerMap));

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
  if (!e || !e.range || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) {
    console.log('Invalid edit event');
    return;
  }

  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  if (row === 1 || col !== statusCol) {
    console.log('Not a status column edit on data row');
    return;
  }

  const newStatus = e.value;
  const statusKey = String(newStatus || '').trim().toLowerCase();
  
  console.log('Status change detected:', newStatus, 'Key:', statusKey);

  // Define status triggers and their configurations
  const statusTriggers = {
    'yes': { 
      configKey: 'STATUS_ORDERED_VALUE', 
      emailType: 'ordered',
      subject: 'Purchase Order Placed',
      requiresNotificationCheck: true
    },
    'in transit': { 
      configKey: 'STATUS_TRANSIT_VALUE', 
      emailType: 'transit',
      subject: 'Order In Transit',
      requiresNotificationCheck: false
    },
    'received': { 
      configKey: 'STATUS_RECEIVED_VALUE', 
      emailType: 'received',
      subject: 'Order Delivered',
      requiresNotificationCheck: false
    }
  };

  const trigger = statusTriggers[statusKey];
  
  if (!trigger) {
    console.log('Status does not match any trigger values:', statusKey);
    return;
  }

  // Check if we should skip notification (for "Yes"/ordered status only)
  if (trigger.requiresNotificationCheck) {
    const notifiedCell = sheet.getRange(row, notifiedCol).getValue();
    if (notifiedCell && String(notifiedCell).trim()) {
      console.log('Already notified for ordered status, skipping');
      return;
    }
  }

  const rowObj = getRowObject_(sheet, row, headerMap);
  const recipient = rowObj[cfg.EMAIL_HEADER];
  
  if (!recipient || String(recipient).indexOf('@') < 0) {
    throw new Error('Invalid or missing email in row ' + row + ': ' + recipient);
  }

  console.log('Preparing to send', trigger.emailType, 'email to:', recipient);

  const data = buildEmailData_(rowObj, headerMap, cfg, row, newStatus);
  data.emailType = trigger.emailType;

  const subject = trigger.subject + (data.poNumber ? (' – ' + data.poNumber) : '');
  const htmlBody = buildStatusEmailHtml_(data, cfg);
  
  sendNotificationEmail_(recipient, subject, htmlBody, data.cc || '');
  console.log(trigger.emailType.toUpperCase(), 'email sent successfully');

  // Mark as notified (for ordered status) or log the action
  if (trigger.requiresNotificationCheck) {
    setCell_(sheet, row, notifiedCol, nowIso_());
    console.log('Notification timestamp recorded');
  } else {
    console.log(trigger.emailType.toUpperCase(), 'notification sent for row', row);
  }
}

/**
 * Build comprehensive email data object
 */
function buildEmailData_(rowObj, headerMap, cfg, row, status) {
  const name = headerMap[cfg.NAME_HEADER] ? rowObj[cfg.NAME_HEADER] : '';
  const email = headerMap[cfg.EMAIL_HEADER] ? rowObj[cfg.EMAIL_HEADER] : '';
  const cc = (cfg.CC_HEADER && headerMap[cfg.CC_HEADER]) ? rowObj[cfg.CC_HEADER] : '';
  
  const costRaw = headerMap[cfg.COST_HEADER] ? rowObj[cfg.COST_HEADER] : '';
  const costFmt = formatCurrency_(costRaw, cfg.CURRENCY || 'USD');

  // Enhanced field mapping with multiple possible column names
  const description = getFieldValue_(rowObj, headerMap, [
    'Purchase order description (program and pur',
    'Purchase order description (program and purpose)',
    'Description', 
    'Item Description',
    'Item',
    'Product'
  ]);

  const quantity = getFieldValue_(rowObj, headerMap, ['Quantity', 'Qty', 'Amount']);
  const poNumber = getFieldValue_(rowObj, headerMap, ['Purchase Order Number', 'PO Number', 'PO #']);
  const quotePdf = getFieldValue_(rowObj, headerMap, ['Quote PDF', 'Quote', 'PDF']);
  const messageKey = getFieldValue_(rowObj, headerMap, ['MessageKey', 'Message Key', 'Tracking']);
  
  // Get timestamp - try multiple possible column names
  let timestamp = getFieldValue_(rowObj, headerMap, ['Timestamp', 'Date', 'Submitted', 'Created']);
  if (!timestamp) timestamp = nowIso_();

  return {
    name: name || '',
    recipient: email || '',
    requesterEmail: email || '',
    cc: cc || '',
    status: status || '',
    cost: costFmt,
    costRaw: costRaw,
    description: description || '',
    item: description || '', // Maintain compatibility
    quantity: quantity || '',
    poNumber: poNumber || '',
    quotePdf: quotePdf || '',
    messageKey: messageKey || '',
    rowNumber: row,
    timestamp: timestamp,
    isSubmissionNotification: status === 'Submitted'
  };
}

/**
 * Get field value by trying multiple possible column names
 */
function getFieldValue_(rowObj, headerMap, possibleNames) {
  for (const name of possibleNames) {
    if (headerMap[name] && rowObj[name]) {
      return rowObj[name];
    }
  }
  return '';
}

/**
 * Enhanced header ensuring with better column detection
 */
function ensureHeader_(sheet, headerName) {
  const lastCol = sheet.getLastColumn() || 1;
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  for (let c = 0; c < header.length; c++) {
    if (String(header[c]).trim() === headerName) {
      console.log('Header', headerName, 'found at column', c + 1);
      return c + 1;
    }
  }
  
  console.log('Adding new header:', headerName, 'at column', lastCol + 1);
  sheet.getRange(1, lastCol + 1).setValue(headerName);
  return lastCol + 1;
}

/**
 * Enhanced configuration test with logo validation
 */
function testConfiguration() {
  console.log('=== Testing Enhanced Configuration ===');
  
  try {
    const cfg = getConfig_();
    console.log('Configuration loaded:', cfg);
    
    const sheet = getActiveSheet_(cfg);
    console.log('Sheet found:', sheet.getName());
    
    const headerMap = getHeaderMap_(sheet);
    console.log('Headers mapped:', headerMap);
    
    // Test all status trigger values
    console.log('Testing status trigger values...');
    console.log('ORDERED trigger "' + cfg.STATUS_ORDERED_VALUE + '" configured');
    console.log('TRANSIT trigger "' + cfg.STATUS_TRANSIT_VALUE + '" configured');
    console.log('RECEIVED trigger "' + cfg.STATUS_RECEIVED_VALUE + '" configured');
    
    // Test required headers
    console.log('Testing required headers...');
    console.log('STATUS_HEADER "' + cfg.STATUS_HEADER + '" found:', !!headerMap[cfg.STATUS_HEADER]);
    console.log('EMAIL_HEADER "' + cfg.EMAIL_HEADER + '" found:', !!headerMap[cfg.EMAIL_HEADER]);
    console.log('NAME_HEADER "' + cfg.NAME_HEADER + '" found:', !!headerMap[cfg.NAME_HEADER]);
    console.log('COST_HEADER "' + cfg.COST_HEADER + '" found:', !!headerMap[cfg.COST_HEADER]);
    
    // Test purchasing team configuration
    console.log('Testing purchasing team configuration...');
    console.log('PURCHASING_EMAIL:', cfg.PURCHASING_EMAIL || 'NOT SET');
    console.log('PURCHASING_CC_EMAIL:', cfg.PURCHASING_CC_EMAIL || 'NOT SET');
    
    // Test logo configuration
    console.log('Testing logo configuration...');
    console.log('LOGO_URL:', cfg.LOGO_URL);
    console.log('LOGO2_URL:', cfg.LOGO2_URL);
    const hasPlaceholders = cfg.LOGO_URL.includes('your-github-username') || cfg.LOGO_URL.includes('your-repo-name');
    console.log('Logo config status:', hasPlaceholders ? '⚠️ NEEDS UPDATE' : '✅ CONFIGURED');
    
    // Test enhanced field detection
    if (sheet.getLastRow() >= 2) {
      console.log('Testing enhanced field detection on row 2...');
      const rowObj = getRowObject_(sheet, 2, headerMap);
      const testData = buildEmailData_(rowObj, headerMap, cfg, 2, 'Test');
      
      console.log('Enhanced data extraction results:');
      console.log('- Description:', testData.description ? 'FOUND' : 'NOT FOUND');
      console.log('- Quantity:', testData.quantity ? 'FOUND' : 'NOT FOUND');
      console.log('- PO Number:', testData.poNumber ? 'FOUND' : 'NOT FOUND');
      console.log('- Message Key:', testData.messageKey ? 'FOUND' : 'NOT FOUND');
      console.log('- Cost:', testData.cost);
      
      const email = testData.recipient;
      console.log('Email validation:', email, email && email.indexOf('@') > 0 ? 'VALID' : 'INVALID');
    }
    
    console.log('=== Enhanced Configuration Test Complete ===');
    return true;
    
  } catch (error) {
    console.error('Configuration test failed:', error.toString());
    return false;
  }
}

/**
 * Test logo selection logic
 */
function testLogoSelection() {
  console.log('=== Testing Logo Selection Logic ===');
  
  const cfg = getConfig_();
  
  // Test different email scenarios
  const testEmails = [
    { email: 'user@spectralenergies.com', expected: 'Spectral only' },
    { email: 'student@purdue.edu', expected: 'Both logos' },
    { email: 'researcher@purdue.edu', expected: 'Both logos' },
    { email: 'john@gmail.com', expected: 'Spectral only (default)' },
    { email: 'admin@spectralenergies.com', expected: 'Spectral only' }
  ];
  
  console.log('\n📋 Logo Selection Test Results:');
  console.log('─'.repeat(60));
  
  testEmails.forEach(test => {
    const testData = {
      recipient: test.email,
      requesterEmail: test.email,
      name: 'Test User'
    };
    
    // Call the logo selection function from Email.gs
    const selectedLogos = selectLogos_(testData, cfg);
    const result = selectedLogos.length === 2 ? 'Both logos' : 'Spectral only';
    const status = result.includes(test.expected.split(' ')[0]) ? '✅' : '❌';
    
    console.log(`${status} ${test.email.padEnd(25)} → ${result} (expected: ${test.expected})`);
  });
  
  console.log('─'.repeat(60));
  console.log('\n🔗 Logo URLs configured:');
  console.log('Spectral Energies:', cfg.LOGO_URL);
  console.log('Purdue Propulsion:', cfg.LOGO2_URL);
  console.log('');
  console.log('💡 GitHub Configuration:');
  console.log('   Username:', GITHUB_CONFIG.username);
  console.log('   Repository:', GITHUB_CONFIG.repository);
  console.log('   Branch:', GITHUB_CONFIG.branch);
  console.log('');
  if (cfg.LOGO_URL.includes('your-github-username') || cfg.LOGO2_URL.includes('your-repo-name')) {
    console.log('⚠️  WARNING: GitHub config still has placeholder values!');
    console.log('   Update GITHUB_CONFIG in Config.gs with your actual username and repo name.');
  } else {
    console.log('✅ GitHub configuration looks good!');
  }
}

/**
 * Test that logo URLs are accessible
 */
function testLogoUrls() {
  console.log('=== Testing Logo URL Accessibility ===');
  
  const cfg = getConfig_();
  
  console.log('🔗 Testing logo URLs:');
  console.log('Spectral logo:', cfg.LOGO_URL);
  console.log('Purdue logo:', cfg.LOGO2_URL);
  console.log('');
  
  // Test if URLs contain placeholders
  const hasPlaceholders = cfg.LOGO_URL.includes('your-github-username') || 
                         cfg.LOGO_URL.includes('your-repo-name') ||
                         cfg.LOGO2_URL.includes('your-github-username') || 
                         cfg.LOGO2_URL.includes('your-repo-name');
  
  if (hasPlaceholders) {
    console.log('❌ CONFIGURATION ERROR:');
    console.log('   Your GitHub config still has placeholder values.');
    console.log('   Update these in Config.gs:');
    console.log('   - username: "' + GITHUB_CONFIG.username + '"');
    console.log('   - repository: "' + GITHUB_CONFIG.repository + '"');
    console.log('');
    return false;
  }
  
  console.log('✅ Configuration appears correct.');
  console.log('');
  console.log('🧪 To test if images actually load:');
  console.log('   1. Copy these URLs into your browser');
  console.log('   2. They should display your logo images');
  console.log('   3. If they show "404 Not Found", check:');
  console.log('      - Repository is public');
  console.log('      - Files exist in /assets/ folder');
  console.log('      - Filenames match exactly (case-sensitive)');
  
  return true;
}

/**
 * Test individual email sending
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
  
  const rowObj = getRowObject_(sheet, 2, headerMap);
  const recipient = rowObj[cfg.EMAIL_HEADER];
  
  if (!recipient || recipient.indexOf('@') < 0) {
    console.error('Invalid email in test row:', recipient);
    return;
  }
  
  const testData = buildEmailData_(rowObj, headerMap, cfg, 2, 'Test');
  testData.emailType = 'ordered';
  
  const subject = 'Test Email - Purchase Order System';
  const htmlBody = buildStatusEmailHtml_(testData, cfg);
  
  try {
    sendNotificationEmail_(recipient, subject, htmlBody, '');
    console.log('✅ Test email sent successfully to:', recipient);
    console.log('Email content preview (first 200 chars):', htmlBody.substring(0, 200) + '...');
  } catch (error) {
    console.error('❌ Failed to send test email:', error.toString());
  }
}

/**
 * Test purchasing team email notification
 */
function testPurchasingTeamEmail() {
  console.log('=== Testing Purchasing Team Email ===');
  
  const cfg = getConfig_();
  
  if (!cfg.PURCHASING_EMAIL) {
    console.error('❌ PURCHASING_EMAIL not configured.');
    console.log('Please run setupPurchasingTeamEmails() first or use the menu:');
    console.log('👥 Setup Purchasing Team Emails');
    return false;
  }
  
  console.log('📧 Purchasing team email configured:', cfg.PURCHASING_EMAIL);
  if (cfg.PURCHASING_CC_EMAIL) {
    console.log('📧 CC email configured:', cfg.PURCHASING_CC_EMAIL);
  }
  
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  if (sheet.getLastRow() < 2) {
    console.error('❌ No data rows found for testing. Please add test data first.');
    return false;
  }
  
  const testRow = sheet.getLastRow();
  const rowObj = getRowObject_(sheet, testRow, headerMap);
  
  console.log('📋 Using test data from row:', testRow);
  console.log('📋 Requester name:', rowObj[cfg.NAME_HEADER] || 'Not provided');
  console.log('📋 Requester email:', rowObj[cfg.EMAIL_HEADER] || 'Not provided');
  
  const data = buildEmailData_(rowObj, headerMap, cfg, testRow, 'Submitted');
  data.isSubmissionNotification = true;
  
  console.log('💼 Purchasing team notification data prepared:');
  console.log('   - Requester:', data.name || 'Unknown');
  console.log('   - Description:', data.description ? 'Provided' : 'Missing');
  console.log('   - Cost:', data.cost || 'Not specified');
  console.log('   - PO Number:', data.poNumber || 'Not assigned');
  
  const subject = 'TEST: New Purchase Request Submitted' + (data.name ? (' from ' + data.name) : '');
  
  try {
    console.log('📤 Building purchasing team email...');
    const htmlBody = buildPurchasingTeamEmailHtml_(data, cfg);
    
    console.log('📤 Sending email to purchasing team...');
    console.log('   - Primary recipient:', cfg.PURCHASING_EMAIL);
    if (cfg.PURCHASING_CC_EMAIL) {
      console.log('   - CC recipient:', cfg.PURCHASING_CC_EMAIL);
    }
    
    sendNotificationEmail_(cfg.PURCHASING_EMAIL, subject, htmlBody, cfg.PURCHASING_CC_EMAIL);
    
    console.log('✅ SUCCESS: Purchasing team email sent successfully!');
    console.log('');
    console.log('🎯 What to check:');
    console.log('   1. Check ' + cfg.PURCHASING_EMAIL + ' for the email');
    if (cfg.PURCHASING_CC_EMAIL) {
      console.log('   2. Check ' + cfg.PURCHASING_CC_EMAIL + ' for the CC copy');
    }
    console.log('   3. Verify the email contains all purchase request details');
    console.log('   4. Check that the "Action Required" styling appears correctly');
    console.log('   5. Test the "Reply to Requester" button if requester email exists');
    
    return true;
    
  } catch (error) {
    console.error('❌ FAILED: Purchasing team email test failed:', error.toString());
    console.log('');
    console.log('🔧 Troubleshooting steps:');
    console.log('   1. Check that PURCHASING_EMAIL is a valid email address');
    console.log('   2. Ensure the purchasing-team-email.html template exists');
    console.log('   3. Verify email permissions are granted to the script');
    console.log('   4. Check the execution transcript for detailed error info');
    
    return false;
  }
}

/**
 * Test form submission notification
 */
function testFormSubmitNotification() {
  console.log('=== Testing Form Submit Notification ===');
  console.log('This tests the onFormSubmit handler that triggers when forms are submitted');
  
  const cfg = getConfig_();
  
  if (!cfg.PURCHASING_EMAIL) {
    console.error('❌ PURCHASING_EMAIL not configured.');
    console.log('Please run setupPurchasingTeamEmails() first or use the menu.');
    return false;
  }
  
  console.log('✅ Purchasing team email configured:', cfg.PURCHASING_EMAIL);
  
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  if (sheet.getLastRow() < 2) {
    console.error('❌ No data rows found for testing');
    console.log('Please add some test data to your sheet first.');
    return false;
  }
  
  const submissionRow = sheet.getLastRow();
  console.log('📝 Simulating form submission for row:', submissionRow);
  
  const mockEvent = {
    range: sheet.getRange(submissionRow, 1),
    values: [sheet.getRange(submissionRow, 1, 1, sheet.getLastColumn()).getValues()[0]]
  };
  
  const rowObj = getRowObject_(sheet, submissionRow, headerMap);
  console.log('📋 Submission data preview:');
  console.log('   - Requester:', rowObj[cfg.NAME_HEADER] || 'Not provided');
  console.log('   - Email:', rowObj[cfg.EMAIL_HEADER] || 'Not provided');
  console.log('   - Description:', (rowObj['Purchase order description (program and pur'] || '').substring(0, 50) + '...');
  console.log('   - Cost:', rowObj[cfg.COST_HEADER] || 'Not specified');
  
  try {
    console.log('🚀 Calling onFormSubmitHandler...');
    onFormSubmitHandler(mockEvent);
    
    console.log('✅ SUCCESS: Form submission test completed successfully');
    console.log('');
    console.log('📧 Check these email addresses:');
    console.log('   - Primary:', cfg.PURCHASING_EMAIL);
    if (cfg.PURCHASING_CC_EMAIL) {
      console.log('   - CC:', cfg.PURCHASING_CC_EMAIL);
    }
    console.log('');
    console.log('🎯 The purchasing team should have received a "New Purchase Request" email');
    console.log('with all the submission details and "Action Required" styling.');
    
    return true;
    
  } catch (error) {
    console.error('❌ FAILED: Form submission test failed:', error.toString());
    console.log('');
    console.log('🔧 Check:');
    console.log('   1. Purchasing team emails are configured correctly');
    console.log('   2. purchasing-team-email.html template exists');
    console.log('   3. Required headers exist in your sheet');
    
    return false;
  }
}

/**
 * Test specific email status types
 */
function testSpecificEmailType() {
  console.log('=== Testing Specific Email Types ===');
  
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  if (sheet.getLastRow() < 2) {
    console.error('No data rows found for testing');
    return;
  }
  
  const rowObj = getRowObject_(sheet, 2, headerMap);
  const recipient = rowObj[cfg.EMAIL_HEADER];
  
  if (!recipient || recipient.indexOf('@') < 0) {
    console.error('Invalid email in test row. Update row 2 with a valid email address.');
    return;
  }
  
  const emailTypes = [
    { type: 'ordered', status: 'Yes', subject: 'Test: Order Confirmation' },
    { type: 'transit', status: 'In Transit', subject: 'Test: In Transit Notification' },
    { type: 'received', status: 'Received', subject: 'Test: Delivery Confirmation' }
  ];
  
  emailTypes.forEach((emailTest, index) => {
    console.log(`\n--- Testing ${emailTest.type} email ---`);
    
    try {
      const testData = buildEmailData_(rowObj, headerMap, cfg, 2, emailTest.status);
      testData.emailType = emailTest.type;
      
      const htmlBody = buildStatusEmailHtml_(testData, cfg);
      
      if (index > 0) {
        Utilities.sleep(2000);
      }
      
      sendNotificationEmail_(recipient, emailTest.subject, htmlBody, '');
      console.log(`✅ ${emailTest.type} email sent successfully`);
      
    } catch (error) {
      console.error(`❌ ${emailTest.type} email failed:`, error.toString());
    }
  });
  
  console.log('\n=== All email type tests completed ===');
  console.log('Check your email inbox:', recipient);
}

/**
 * Test both requester and purchasing team emails together
 */
function testBothEmailTypes() {
  console.log('=== Testing Both Requester & Purchasing Team Emails ===');
  
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  if (sheet.getLastRow() < 2) {
    console.error('❌ No data rows found for testing');
    return false;
  }
  
  console.log('\n🏢 PART 1: Testing Purchasing Team Email');
  console.log('─'.repeat(50));
  const purchasingResult = testPurchasingTeamEmail();
  
  Utilities.sleep(2000);
  
  console.log('\n👤 PART 2: Testing Requester Email');
  console.log('─'.repeat(50));
  const requesterResult = testSendEmail();
  
  console.log('\n📊 COMBINED TEST RESULTS');
  console.log('─'.repeat(50));
  console.log('Purchasing team email:', purchasingResult ? '✅ PASSED' : '❌ FAILED');
  console.log('Requester email:', requesterResult ? '✅ PASSED' : '❌ FAILED');
  
  if (purchasingResult && requesterResult) {
    console.log('🎉 SUCCESS: Both email types working correctly!');
    console.log('');
    console.log('📧 Check these inboxes:');
    console.log('   • Purchasing team:', cfg.PURCHASING_EMAIL);
    if (cfg.PURCHASING_CC_EMAIL) {
      console.log('   • Purchasing CC:', cfg.PURCHASING_CC_EMAIL);
    }
    
    const testRow = sheet.getLastRow();
    const rowObj = getRowObject_(sheet, testRow, headerMap);
    const requesterEmail = rowObj[cfg.EMAIL_HEADER];
    if (requesterEmail) {
      console.log('   • Requester:', requesterEmail);
    }
    
    return true;
  } else {
    console.log('⚠️  Some email tests failed - check individual results above');
    return false;
  }
}

/**
 * Create comprehensive test row with all fields
 */
function createComprehensiveTestRow() {
  console.log('=== Creating Comprehensive Test Row ===');
  
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  const testRow = sheet.getLastRow() + 1;
  const timestamp = nowIso_();
  
  const testData = [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  headers.forEach((header, index) => {
    const headerStr = String(header).trim();
    
    if (headerStr === 'Name') {
      testData[index] = 'Test User Enhanced';
    } else if (headerStr === 'Email') {
      testData[index] = 'test@example.com';
    } else if (headerStr.includes('Purchase order description')) {
      testData[index] = 'Enhanced Test Equipment for Laboratory Research Project - High Precision Digital Multimeter with Data Logging Capability';
    } else if (headerStr === 'Purchase Order Number') {
      testData[index] = 'TEST-ENH-' + Date.now();
    } else if (headerStr === 'Quote PDF') {
      testData[index] = 'https://example.com/enhanced-quote.pdf';
    } else if (headerStr === 'Cost ($)') {
      testData[index] = '2450.00';
    } else if (headerStr === 'Order Placed?') {
      testData[index] = '';
    } else if (headerStr === 'Notified') {
      testData[index] = '';
    } else if (headerStr === 'MessageKey') {
      testData[index] = 'ENH-MSG-' + Date.now();
    } else if (headerStr === 'Timestamp') {
      testData[index] = timestamp;
    } else {
      testData[index] = '';
    }
  });
  
  sheet.getRange(testRow, 1, 1, testData.length).setValues([testData]);
  
  console.log('✅ Comprehensive test row created at row:', testRow);
  console.log('Test data preview:', testData.slice(0, 8));
  
  return testRow;
}

/**
 * Test all status change notifications
 */
function testAllStatusNotifications() {
  console.log('=== Testing All Status Notifications ===');
  
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  if (sheet.getLastRow() < 2) {
    console.error('No data rows found for testing. Please add test data first.');
    return false;
  }
  
  const testStatuses = [
    { value: 'Yes', type: 'Ordered' },
    { value: 'In Transit', type: 'Transit' },
    { value: 'Received', type: 'Received' }
  ];
  
  const testRow = 2;
  const statusCol = headerMap[cfg.STATUS_HEADER];
  
  for (const statusTest of testStatuses) {
    console.log(`\n--- Testing ${statusTest.type} Status ---`);
    
    try {
      const mockRange = {
        getRow: () => testRow,
        getColumn: () => statusCol,
        getNumRows: () => 1,
        getNumColumns: () => 1
      };
      
      const mockEvent = {
        range: mockRange,
        value: statusTest.value
      };
      
      console.log(`Setting status to: ${statusTest.value}`);
      sheet.getRange(testRow, statusCol).setValue(statusTest.value);
      
      console.log('Calling onEditHandler...');
      onEditHandler(mockEvent);
      
      console.log(`✅ ${statusTest.type} notification test completed`);
      
      Utilities.sleep(1000);
      
    } catch (error) {
      console.error(`❌ ${statusTest.type} notification test failed:`, error.toString());
      return false;
    }
  }
  
  console.log('=== All Status Notification Tests Complete ===');
  return true;
}

/**
 * Complete system test with all status notifications
 */
function testCompleteSystem() {
  console.log('=== COMPLETE SYSTEM TEST ===');
  console.log('This will simulate the entire purchase order lifecycle');
  
  try {
    const cfg = getConfig_();
    const sheet = getActiveSheet_(cfg);
    const headerMap = getHeaderMap_(sheet);
    
    if (sheet.getLastRow() < 2) {
      console.error('No data rows found. Please add some test data first.');
      return false;
    }
    
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
      testRow = createComprehensiveTestRow();
    }
    
    console.log('Using test row:', testRow);
    
    const rowObj = getRowObject_(sheet, testRow, headerMap);
    const email = rowObj[cfg.EMAIL_HEADER];
    
    if (!email || email.indexOf('@') < 0) {
      console.error('Invalid email in test row. Please update the email in row', testRow, 'to a valid email address.');
      return false;
    }
    
    console.log('Test email recipient:', email);
    console.log('IMPORTANT: Make sure this email is one you can check!');
    
    const statusCol = headerMap[cfg.STATUS_HEADER];
    const statuses = [
      { value: cfg.STATUS_ORDERED_VALUE, name: 'Ordered' },
      { value: cfg.STATUS_TRANSIT_VALUE, name: 'In Transit' },
      { value: cfg.STATUS_RECEIVED_VALUE, name: 'Received' }
    ];
    
    let allPassed = true;
    
    for (const statusTest of statuses) {
      console.log(`\n--- Testing ${statusTest.name} Status ---`);
      
      try {
        const mockRange = {
          getRow: () => testRow,
          getColumn: () => statusCol,
          getNumRows: () => 1,
          getNumColumns: () => 1
        };
        
        const mockEvent = {
          range: mockRange,
          value: statusTest.value
        };
        
        sheet.getRange(testRow, statusCol).setValue(statusTest.value);
        console.log(`Status set to: ${statusTest.value}`);
        
        onEditHandler(mockEvent);
        console.log(`✅ ${statusTest.name} notification sent successfully`);
        
        Utilities.sleep(1000);
        
      } catch (error) {
        console.error(`❌ ${statusTest.name} test failed:`, error.toString());
        allPassed = false;
      }
    }
    
    const finalNotifiedValue = sheet.getRange(testRow, notifiedCol).getValue();
    
    console.log('\n=== COMPLETE TEST RESULTS ===');
    console.log('Test row:', testRow);
    console.log('Email recipient:', email);
    console.log('Notified timestamp:', finalNotifiedValue || 'Not recorded');
    
    if (allPassed) {
      console.log('✅ SUCCESS: All lifecycle emails sent successfully');
      console.log('Check the email inbox for:', email);
      console.log('You should receive 3 emails: Ordered → In Transit → Received');
      return true;
    } else {
      console.log('⚠️  WARNING: Some tests failed - check error messages above');
      return false;
    }
    
  } catch (error) {
    console.error('❌ COMPLETE SYSTEM TEST FAILED:', error.toString());
    return false;
  }
}

/**
 * Comprehensive email system test
 */
function testCompleteEmailSystem() {
  console.log('=== COMPREHENSIVE EMAIL SYSTEM TEST ===');
  console.log('This tests the entire email notification system');
  console.log('');
  
  let allPassed = true;
  const results = {};
  
  console.log('🔧 TEST 1: Configuration Validation');
  console.log('─'.repeat(40));
  results.config = testConfiguration();
  if (!results.config) allPassed = false;
  
  Utilities.sleep(1000);
  
  console.log('\n🏢 TEST 2: Purchasing Team Notification');
  console.log('─'.repeat(40));
  results.purchasing = testPurchasingTeamEmail();
  if (!results.purchasing) allPassed = false;
  
  Utilities.sleep(2000);
  
  console.log('\n👤 TEST 3: All Requester Email Types');
  console.log('─'.repeat(40));
  results.requester = testSpecificEmailType();
  if (!results.requester) allPassed = false;
  
  Utilities.sleep(1000);
  
  console.log('\n📝 TEST 4: Form Submission Handler');
  console.log('─'.repeat(40));
  results.formSubmit = testFormSubmitNotification();
  if (!results.formSubmit) allPassed = false;
  
  console.log('\n📊 COMPREHENSIVE TEST RESULTS');
  console.log('═'.repeat(50));
  console.log('Configuration validation:', results.config ? '✅ PASSED' : '❌ FAILED');
  console.log('Purchasing team emails:', results.purchasing ? '✅ PASSED' : '❌ FAILED');
  console.log('Requester emails (all types):', results.requester ? '✅ PASSED' : '❌ FAILED');
  console.log('Form submission handler:', results.formSubmit ? '✅ PASSED' : '❌ FAILED');
  console.log('');
  
  if (allPassed) {
    console.log('🎉 SUCCESS: Complete email system is working perfectly!');
    console.log('');
    console.log('📧 EMAILS SENT TO:');
    const cfg = getConfig_();
    console.log('   • Purchasing team:', cfg.PURCHASING_EMAIL);
    if (cfg.PURCHASING_CC_EMAIL) {
      console.log('   • Purchasing CC:', cfg.PURCHASING_CC_EMAIL);
    }
    
    const sheet = getActiveSheet_(cfg);
    const headerMap = getHeaderMap_(sheet);
    const testRow = sheet.getLastRow();
    const rowObj = getRowObject_(sheet, testRow, headerMap);
    const requesterEmail = rowObj[cfg.EMAIL_HEADER];
    if (requesterEmail) {
      console.log('   • Requester (3 emails):', requesterEmail);
    }
    
    console.log('');
    console.log('🎯 TOTAL EMAILS SENT: ~5-6 test emails');
    console.log('Your email notification system is fully operational!');
    
  } else {
    console.log('⚠️  Some tests failed. Review the individual test results above.');
    console.log('Fix any issues and run the failed tests individually.');
  }
  
  return allPassed;
}

/**
 * Run enhanced system test with UI feedback
 */
function runEnhancedSystemTest() {
  console.log('Starting enhanced system test...');
  
  try {
    const configResult = testConfiguration();
    if (!configResult) {
      throw new Error('Configuration test failed. Check the logs and fix issues before proceeding.');
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Enhanced System Test', 
      'This will test the complete enhanced system:\n\n' +
      '1. Create comprehensive test data\n' +
      '2. Test all status notifications (Ordered, In Transit, Received)\n' +
      '3. Send REAL emails for each status\n' +
      '4. Include enhanced descriptions and all fields\n\n' +
      'Make sure the test email is one you can access!\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      console.log('Enhanced system test cancelled');
      return;
    }
    
    const testRow = createComprehensiveTestRow();
    const allStatusResult = testAllStatusNotifications();
    
    const resultMessage = allStatusResult ? 
      '✅ ENHANCED SYSTEM TEST PASSED!\n\n' +
      'Successfully tested:\n' +
      '• Configuration validation\n' +
      '• Comprehensive test data creation\n' +
      '• Ordered status notification\n' +
      '• In Transit status notification\n' +
      '• Received status notification\n' +
      '• Enhanced email content with descriptions\n\n' +
      'Check all emails sent to verify content and formatting.\n' +
      'Test row created at: Row ' + testRow :
      
      '❌ ENHANCED SYSTEM TEST FAILED\n\n' +
      'Check the Execution Transcript for detailed error information.';
    
    ui.alert('Enhanced Test Results', resultMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Enhanced system test failed:', error.toString());
    SpreadsheetApp.getUi().alert(
      'Test Failed', 
      'Enhanced system test failed:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Enhanced menu with comprehensive testing options including logo tests
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧪 Enhanced Purchase Order System')
    .addSubMenu(ui.createMenu('🔍 Basic Tests')
      .addItem('📋 Test Configuration', 'testConfiguration')
      .addItem('📧 Test Requester Email', 'testSendEmail')
      .addItem('🏢 Test Purchasing Team Email', 'testPurchasingTeamEmail')
      .addItem('📝 Test Form Submit Handler', 'testFormSubmitNotification')
    )
    .addSubMenu(ui.createMenu('📨 Email Template Tests')
      .addItem('👤 Test All Requester Emails', 'testSpecificEmailType')
      .addItem('🏢👤 Test Both Email Types', 'testBothEmailTypes')
      .addItem('📧 Test All Status Notifications', 'testAllStatusNotifications')
      .addItem('🧪 Test Email Templates', 'testEmailTemplates')
      .addItem('🖼️ Test Logo Selection', 'testLogoSelection')
      .addItem('🔗 Test Logo URLs', 'testLogoUrls')
      .addItem('🔍 Debug Email Logos', 'debugEmailLogos')
      .addItem('📋 Test Template Loading', 'testEmailTemplateLoading')
      .addItem('📁 Check Template Files', 'checkTemplateFiles')
      .addItem('📊 Test Enhanced Data', 'testEnhancedDataProcessing')
    )
    .addSubMenu(ui.createMenu('🚀 Complete Tests')
      .addItem('🔄 Complete System Test', 'testCompleteSystem')
      .addItem('📧 Complete Email System Test', 'testCompleteEmailSystem')
      .addItem('🎛️ Enhanced System Test', 'runEnhancedSystemTest')
      .addItem('📊 Create Test Row', 'createComprehensiveTestRow')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Setup & Configuration')
      .addItem('🔧 Setup Enhanced Properties', 'setupEnhancedPropertiesMenu')
      .addItem('👥 Setup Purchasing Team Emails', 'setupPurchasingTeamEmailsMenu')
      .addItem('🪄 Quick Setup Wizard', 'runQuickSetupWizard')
      .addItem('✅ Validate Configuration', 'validateAndSuggestConfiguration')
    )
    .addItem('📋 View Current Config', 'viewEnhancedConfigMenu')
    .addToUi();
}

// Enhanced menu functions
function setupEnhancedPropertiesMenu() {
  setupEnhancedProperties();
  SpreadsheetApp.getUi().alert(
    'Enhanced Setup Complete', 
    'All enhanced script properties configured including multiple status triggers.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setupPurchasingTeamEmailsMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const emailResponse = ui.prompt(
    'Setup Purchasing Team Email',
    'Enter the primary purchasing team email address:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (emailResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const purchasingEmail = emailResponse.getResponseText().trim();
  if (!purchasingEmail || purchasingEmail.indexOf('@') < 0) {
    ui.alert('Invalid Email', 'Please enter a valid email address.', ui.ButtonSet.OK);
    return;
  }
  
  const ccResponse = ui.prompt(
    'Setup Purchasing Team CC Email',
    'Enter the CC email address (optional):',
    ui.ButtonSet.OK_CANCEL
  );
  
  let ccEmail = '';
  if (ccResponse.getSelectedButton() === ui.Button.OK) {
    ccEmail = ccResponse.getResponseText().trim();
  }
  
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'PURCHASING_EMAIL': purchasingEmail,
    'PURCHASING_CC_EMAIL': ccEmail
  });
  
  ui.alert('Setup Complete', 
    `Purchasing team emails configured:\nPrimary: ${purchasingEmail}\nCC: ${ccEmail || 'Not set'}`, 
    ui.ButtonSet.OK);
}

function viewEnhancedConfigMenu() {
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  const message = 
    'ENHANCED SYSTEM CONFIGURATION:\n\n' +
    '📋 BASIC SETTINGS:\n' +
    `Sheet: "${sheet.getName()}"\n` +
    `Status Header: "${cfg.STATUS_HEADER}"\n` +
    `Email Header: "${cfg.EMAIL_HEADER}"\n` +
    `Name Header: "${cfg.NAME_HEADER}"\n` +
    `Cost Header: "${cfg.COST_HEADER}"\n\n` +
    
    '🔄 STATUS TRIGGERS:\n' +
    `Ordered: "${cfg.STATUS_ORDERED_VALUE}"\n` +
    `In Transit: "${cfg.STATUS_TRANSIT_VALUE}"\n` +
    `Received: "${cfg.STATUS_RECEIVED_VALUE}"\n\n` +
    
    '👥 PURCHASING TEAM:\n' +
    `Primary: ${cfg.PURCHASING_EMAIL || 'NOT SET'}\n` +
    `CC: ${cfg.PURCHASING_CC_EMAIL || 'NOT SET'}\n\n` +
    
    '🖼️ LOGO CONFIGURATION:\n' +
    `Spectral Logo: ${cfg.LOGO_URL.includes('your-github') ? 'NEEDS UPDATE' : 'CONFIGURED'}\n` +
    `Purdue Logo: ${cfg.LOGO2_URL.includes('your-github') ? 'NEEDS UPDATE' : 'CONFIGURED'}\n\n` +
    
    '📊 SHEET STATUS:\n' +
    `Headers Found: ${Object.keys(headerMap).length}\n` +
    `Data Rows: ${Math.max(0, sheet.getLastRow() - 1)}`;
  
  SpreadsheetApp.getUi().alert('Enhanced Configuration', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Debug logo data flow in email generation
 */
function debugEmailLogos() {
  console.log('=== Debugging Logo Data Flow ===');
  
  const cfg = getConfig_();
  const sheet = getActiveSheet_(cfg);
  const headerMap = getHeaderMap_(sheet);
  
  if (sheet.getLastRow() < 2) {
    console.error('No test data found. Please add data to row 2.');
    return;
  }
  
  const rowObj = getRowObject_(sheet, 2, headerMap);
  const testData = buildEmailData_(rowObj, headerMap, cfg, 2, 'Test');
  testData.emailType = 'ordered';
  
  console.log('📧 Test email recipient:', testData.recipient);
  
  // Test logo selection
  console.log('\n🖼️ Testing logo selection...');
  const selectedLogos = selectLogos_(testData, cfg);
  console.log('Number of logos selected:', selectedLogos.length);
  selectedLogos.forEach((logo, index) => {
    console.log(`Logo ${index + 1}:`, logo.url);
  });
  
  // Test email template generation
  console.log('\n📧 Testing email template generation...');
  try {
    const htmlBody = buildStatusEmailHtml_(testData, cfg);
    console.log('Email HTML generated successfully');
    console.log('HTML length:', htmlBody.length, 'characters');
    
    // Check if logos are in the HTML
    const hasSpectralLogo = htmlBody.includes('spectral_logo.png');
    const hasPurdueLogo = htmlBody.includes('purdue_prop_logo.png');
    const hasLogoSection = htmlBody.includes('<div style="display:inline-block;">');
    
    console.log('\n🔍 Logo presence in HTML:');
    console.log('Contains spectral logo URL:', hasSpectralLogo);
    console.log('Contains purdue logo URL:', hasPurdueLogo);
    console.log('Has logo section HTML:', hasLogoSection);
    
    // Show first part of HTML for inspection
    console.log('\n📋 HTML Preview (first 500 chars):');
    console.log(htmlBody.substring(0, 500));
    
    if (!hasLogoSection) {
      console.log('⚠️ WARNING: No logo section found in HTML template');
      console.log('This suggests the template might not be loading the logos properly');
    }
    
  } catch (error) {
    console.error('❌ Email template generation failed:', error.toString());
  }
}

/**
 * Test email template loading specifically
 */
function testEmailTemplateLoading() {
  console.log('=== Testing Email Template Loading ===');
  
  try {
    // Test if we can load the email template files
    const templates = ['email-ordered', 'email-transit', 'email-received', 'purchasing-team-email'];
    
    templates.forEach(templateName => {
      try {
        console.log(`\nTesting template: ${templateName}`);
        const template = HtmlService.createTemplateFromFile(templateName);
        console.log(`✅ ${templateName} loaded successfully`);
        
        // Test with mock data
        template.data = { name: 'Test', recipient: 'test@purdue.edu' };
        template.cfg = getConfig_();
        template.logos = [{ url: 'test-url', alt: 'Test Logo', maxW: '200px', maxH: '80px' }];
        
        const html = template.evaluate().getContent();
        const hasLogoCode = html.includes('logos.forEach') || html.includes('logo.url');
        console.log(`   Logo code present: ${hasLogoCode}`);
        
      } catch (templateError) {
        console.error(`❌ ${templateName} failed to load:`, templateError.toString());
      }
    });
    
  } catch (error) {
    console.error('Template loading test failed:', error.toString());
  }
}

/**
 * Check which HTML template files exist
 */
function checkTemplateFiles() {
  console.log('=== Checking HTML Template Files ===');
  
  const requiredTemplates = [
    'email-ordered',
    'email-transit', 
    'email-received',
    'purchasing-team-email'
  ];
  
  let allFound = true;
  
  requiredTemplates.forEach(templateName => {
    try {
      const template = HtmlService.createTemplateFromFile(templateName);
      console.log(`✅ ${templateName}.html - EXISTS`);
    } catch (error) {
      console.log(`❌ ${templateName}.html - MISSING`);
      console.log(`   Error: ${error.toString()}`);
      allFound = false;
    }
  });
  
  if (!allFound) {
    console.log('');
    console.log('⚠️  MISSING TEMPLATE FILES DETECTED');
    console.log('The system will use fallback templates (basic HTML without enhanced styling)');
    console.log('');
    console.log('📋 TO FIX:');
    console.log('1. Make sure you have created these HTML files in your Apps Script project:');
    requiredTemplates.forEach(name => {
      console.log(`   - ${name}.html`);
    });
    console.log('2. Each file should contain the enhanced email template with logo support');
  } else {
    console.log('');
    console.log('✅ All required template files found!');
  }
  
  return allFound;
}

// End of Code.gs file
