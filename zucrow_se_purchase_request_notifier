/**
 * Enhanced Purchase Order Email Notification System
 * Improved robustness for Google Sheets input handling
 * Compatible with Google Apps Script - ES5 syntax only
 * 
 * Setup Instructions:
 * 1. Add Gmail API service (Services > Gmail API)
 * 2. Set up trigger: onEdit, From spreadsheet, On edit
 * 3. Configure CONFIG section below
 * 4. Run initializeSystem() once after setup
 */

// Enhanced Configuration with validation
var CONFIG = {
  // Auto-detect sheet - will use the first sheet with form responses or main data sheet
  // Leave empty to auto-detect, or specify a sheet name
  SHEET_NAME: '', // Set to '' for auto-detection
  
  // Column headers (case-insensitive) with alternatives
  COLUMN_MAPPINGS: {
    email: ['email', 'email address', 'e-mail'],
    name: ['name', 'full name', 'requester name', 'user name'],
    po: ['purchase order number', 'po number', 'po #', 'order number'],
    description: ['purchase order description', 'description', 'order description', 'item description'],
    quote: ['quote pdf', 'quote', 'quote file', 'pdf quote'],
    ordered: ['order placed?', 'order placed', 'ordered', 'status', 'order status']
  },
  
  // Helper columns
  H_NOTIFIED: 'Notified',
  H_MESSAGEKEY: 'MessageKey',
  
  // Enhanced "yes" values with more variations
  YES_VALUES: ['yes', 'y', 'true', '1', 'placed', 'ordered', 'complete', 'done', 
               'completed', 'finished', 'sent', 'order placed', 'order sent', 
               'confirmed', 'approved', 'processed'],
  
  // Input validation settings
  MAX_EMAIL_LENGTH: 254,
  MAX_PO_LENGTH: 50,
  MAX_DESCRIPTION_LENGTH: 1000,
  MAX_NAME_LENGTH: 100,
  
  // Email settings
  EMAIL_SUBJECT: 'Your order has been placed',
  EMAIL_GREETING: 'Hello',
  EMAIL_SIGNATURE: '-- Purchasing Team',
  ATTACH_QUOTE_PDF: false,
  
  // Sender configuration
  EMAIL_SENDER: 'auto', // 'auto', 'config', or 'user@example.com'
  FALLBACK_SENDERS: [],
  
  // Safety settings
  MAX_EMAILS_PER_HOUR: 50,
  MAX_EMAILS_PER_DAY: 200,
  DUPLICATE_CHECK_DAYS: 30,
  BACKUP_SHEET_NAME: 'Email_Backup_Log',
  RATE_LIMIT_SHEET_NAME: 'Rate_Limit_Log'
};

/**
 * Creates menu when spreadsheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Order Notifier')
    .addItem('Test email for selected row', 'testEmailForRow')
    .addItem('Check setup & validate config', 'checkSetup')
    .addItem('Setup data validation rules', 'setupDataValidation')
    .addSeparator()
    .addItem('Configure sender email', 'configureSenderEmail')
    .addItem('View current configuration', 'viewCurrentConfig')
    .addSeparator()
    .addItem('View email backup log', 'viewEmailBackupLog')
    .addItem('View rate limit status', 'viewRateLimitStatus')
    .addItem('Reset rate limits (emergency)', 'resetRateLimits')
    .addSeparator()
    .addItem('Initialize system', 'initializeSystem')
    .addItem('Audit data integrity', 'auditDataIntegrity')
    .addToUi();
}

/**
 * Auto-detect the main data sheet
 */
function getMainSheet() {
  try {
    var ss = SpreadsheetApp.getActive();
    
    if (CONFIG.SHEET_NAME && CONFIG.SHEET_NAME.trim() !== '') {
      var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
      if (sheet) {
        logToSheet('Using configured sheet: ' + CONFIG.SHEET_NAME);
        return sheet;
      } else {
        logToSheet('WARNING: Configured sheet "' + CONFIG.SHEET_NAME + '" not found, attempting auto-detection');
      }
    }
    
    var sheets = ss.getSheets();
    var candidates = [];
    
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName().toLowerCase();
      
      if (sheetName.indexOf('log') !== -1 || 
          sheetName.indexOf('config') !== -1 ||
          sheetName.indexOf('rate_limit') !== -1 ||
          sheetName.indexOf('backup') !== -1) {
        continue;
      }
      
      if (sheet.getLastRow() > 1 && sheet.getLastColumn() > 3) {
        var headers = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues()[0];
        var headerText = headers.join(' ').toLowerCase();
        
        var score = 0;
        if (headerText.indexOf('email') !== -1) score += 10;
        if (headerText.indexOf('order') !== -1) score += 8;
        if (headerText.indexOf('purchase') !== -1) score += 8;
        if (headerText.indexOf('name') !== -1) score += 5;
        if (headerText.indexOf('description') !== -1) score += 5;
        if (headerText.indexOf('timestamp') !== -1) score += 7;
        if (sheetName.indexOf('response') !== -1) score += 15;
        if (sheetName.indexOf('form') !== -1) score += 10;
        
        if (score > 10) {
          candidates.push({
            sheet: sheet,
            score: score,
            name: sheet.getName()
          });
        }
      }
    }
    
    if (candidates.length > 0) {
      candidates.sort(function(a, b) { return b.score - a.score; });
      var selected = candidates[0];
      logToSheet('Auto-detected main sheet: "' + selected.name + '" (score: ' + selected.score + ')');
      return selected.sheet;
    }
    
    for (var j = 0; j < sheets.length; j++) {
      if (sheets[j].getLastRow() > 1) {
        logToSheet('Fallback: Using first sheet with data: "' + sheets[j].getName() + '"');
        return sheets[j];
      }
    }
    
    if (sheets.length > 0) {
      logToSheet('Last resort: Using first sheet: "' + sheets[0].getName() + '"');
      return sheets[0];
    }
    
    return null;
    
  } catch (error) {
    logToSheet('getMainSheet ERROR: ' + error.toString());
    return null;
  }
}

/**
 * Get the current user's email for sending
 */
function getSenderEmail() {
  try {
    var senderEmail = '';
    
    switch (CONFIG.EMAIL_SENDER.toLowerCase()) {
      case 'auto':
        senderEmail = Session.getActiveUser().getEmail();
        break;
      case 'config':
        senderEmail = getSenderFromConfig();
        break;
      default:
        if (isValidEmail(CONFIG.EMAIL_SENDER)) {
          senderEmail = CONFIG.EMAIL_SENDER;
        } else {
          logToSheet('WARNING: Invalid EMAIL_SENDER config, falling back to current user');
          senderEmail = Session.getActiveUser().getEmail();
        }
        break;
    }
    
    if (!senderEmail || !isValidEmail(senderEmail)) {
      throw new Error('Could not determine valid sender email');
    }
    
    logToSheet('Using sender email: ' + senderEmail);
    return senderEmail;
    
  } catch (error) {
    logToSheet('getSenderEmail ERROR: ' + error.toString());
    
    for (var i = 0; i < CONFIG.FALLBACK_SENDERS.length; i++) {
      var fallback = CONFIG.FALLBACK_SENDERS[i];
      if (isValidEmail(fallback)) {
        logToSheet('Using fallback sender: ' + fallback);
        return fallback;
      }
    }
    
    try {
      var currentUser = Session.getActiveUser().getEmail();
      if (currentUser && isValidEmail(currentUser)) {
        logToSheet('Final fallback to current user: ' + currentUser);
        return currentUser;
      }
    } catch (sessionError) {
      logToSheet('Could not get current user email: ' + sessionError.toString());
    }
    
    throw new Error('No valid sender email available');
  }
}

/**
 * Get sender email from Config sheet
 */
function getSenderFromConfig() {
  try {
    var configSheet = SpreadsheetApp.getActive().getSheetByName('Config');
    if (!configSheet) {
      logToSheet('Config sheet not found for sender email');
      return '';
    }
    
    var data = configSheet.getRange(1, 1, 20, 2).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0] || '').toUpperCase().trim();
      var value = String(data[i][1] || '').trim();
      
      if ((key === 'SENDER_EMAIL' || key === 'FROM_EMAIL' || key === 'EMAIL_FROM') && isValidEmail(value)) {
        return value;
      }
    }
    
    logToSheet('No valid sender email found in Config sheet');
    return '';
    
  } catch (error) {
    logToSheet('getSenderFromConfig ERROR: ' + error.toString());
    return '';
  }
}

/**
 * Enhanced onEdit trigger
 */
function onEdit(e) {
  try {
    if (!e || !e.source || !e.range) {
      return;
    }
    
    var sheet = e.source.getActiveSheet();
    if (!sheet) {
      return;
    }
    
    var mainSheet = getMainSheet();
    if (!mainSheet || sheet.getName() !== mainSheet.getName()) {
      return;
    }
    
    var row = e.range.getRow();
    if (row === 1) {
      return;
    }
    
    var mapping = getColumnMapping(sheet);
    if (!mapping || !mapping.email || !mapping.ordered) {
      logToSheet('onEdit: Required columns not found');
      return;
    }
    
    var editedColumn = e.range.getColumn();
    if (editedColumn !== mapping.ordered) {
      return;
    }
    
    var newValue = e.value;
    if (!newValue) {
      return;
    }
    
    var stringValue = String(newValue).trim();
    if (!stringValue || !isYesValue(stringValue)) {
      return;
    }
    
    logToSheet('onEdit: Processing row ' + row + ' - order marked as placed');
    processOrderRow(sheet, row, mapping);
    
  } catch (error) {
    logToSheet('onEdit ERROR: ' + error.toString());
  }
}

/**
 * Enhanced column mapping
 */
function getColumnMapping(sheet) {
  try {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      return null;
    }
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var mapping = {};
    
    for (var i = 0; i < headers.length; i++) {
      var header = normalizeHeader(headers[i]);
      var colNum = i + 1;
      
      for (var fieldType in CONFIG.COLUMN_MAPPINGS) {
        var possibleNames = CONFIG.COLUMN_MAPPINGS[fieldType];
        for (var j = 0; j < possibleNames.length; j++) {
          if (header === possibleNames[j].toLowerCase() || 
              header.indexOf(possibleNames[j].toLowerCase()) === 0) {
            if (!mapping[fieldType]) {
              mapping[fieldType] = colNum;
              break;
            }
          }
        }
      }
      
      if (header === CONFIG.H_NOTIFIED.toLowerCase()) {
        mapping.notified = colNum;
      } else if (header === CONFIG.H_MESSAGEKEY.toLowerCase()) {
        mapping.messageKey = colNum;
      }
    }
    
    return mapping;
    
  } catch (error) {
    logToSheet('getColumnMapping ERROR: ' + error.toString());
    return null;
  }
}

/**
 * Normalize header text
 */
function normalizeHeader(header) {
  if (!header) return '';
  return String(header)
    .toLowerCase()
    .trim()
    .replace(/[^\w\s]/g, '')
    .replace(/\s+/g, ' ');
}

/**
 * Get row data
 */
function getRowData(sheet, row, mapping) {
  try {
    var lastCol = sheet.getLastColumn();
    var values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    
    var data = {
      email: '',
      name: '',
      po: '',
      description: '',
      quote: '',
      notified: '',
      messageKey: ''
    };
    
    if (mapping.email) {
      data.email = validateAndCleanEmail(values[mapping.email - 1]);
    }
    
    if (mapping.name) {
      data.name = validateAndCleanText(values[mapping.name - 1], CONFIG.MAX_NAME_LENGTH, 'Name');
    }
    
    if (mapping.po) {
      data.po = validateAndCleanText(values[mapping.po - 1], CONFIG.MAX_PO_LENGTH, 'PO Number');
    }
    
    if (mapping.description) {
      data.description = validateAndCleanText(values[mapping.description - 1], CONFIG.MAX_DESCRIPTION_LENGTH, 'Description');
    }
    
    if (mapping.quote) {
      data.quote = cleanString(values[mapping.quote - 1]);
    }
    
    if (mapping.notified) {
      data.notified = cleanString(values[mapping.notified - 1]);
    }
    
    if (mapping.messageKey) {
      data.messageKey = cleanString(values[mapping.messageKey - 1]);
    }
    
    return data;
    
  } catch (error) {
    logToSheet('getRowData ERROR for row ' + row + ': ' + error.toString());
    return null;
  }
}

/**
 * Validate email
 */
function validateAndCleanEmail(value) {
  try {
    if (!value) return '';
    
    var email = String(value).trim().toLowerCase();
    email = email.replace(/^["'\[\(]+|["'\]\)]+$/g, '');
    
    if (email.length > CONFIG.MAX_EMAIL_LENGTH) {
      throw new Error('Email too long: ' + email.length + ' chars');
    }
    
    var emailPattern = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;
    
    if (!emailPattern.test(email)) {
      throw new Error('Invalid email format: ' + email);
    }
    
    return email;
    
  } catch (error) {
    logToSheet('Email validation error: ' + error.toString());
    return '';
  }
}

/**
 * Validate text
 */
function validateAndCleanText(value, maxLength, fieldName) {
  try {
    if (!value) return '';
    
    var text = String(value).trim();
    text = text.replace(/\s+/g, ' ');
    text = text.replace(/[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]/g, '');
    
    if (text.length > maxLength) {
      logToSheet('WARNING: ' + fieldName + ' truncated from ' + text.length + ' to ' + maxLength + ' chars');
      text = text.substring(0, maxLength - 3) + '...';
    }
    
    return text;
    
  } catch (error) {
    logToSheet('Text validation error for ' + fieldName + ': ' + error.toString());
    return String(value || '').trim();
  }
}

/**
 * Check yes values
 */
function isYesValue(value) {
  try {
    if (!value) return false;
    
    var cleanValue = String(value).toLowerCase().trim();
    cleanValue = cleanValue.replace(/[!?\.,:;]/g, '');
    
    if (CONFIG.YES_VALUES.indexOf(cleanValue) !== -1) {
      return true;
    }
    
    var partialMatches = ['order placed', 'order sent', 'has been ordered', 'order confirmed'];
    for (var i = 0; i < partialMatches.length; i++) {
      if (cleanValue.indexOf(partialMatches[i]) !== -1) {
        return true;
      }
    }
    
    if (/^[✓✔☑x\*]$/.test(cleanValue) || cleanValue === 'check' || cleanValue === 'checked') {
      return true;
    }
    
    return false;
    
  } catch (error) {
    logToSheet('isYesValue error: ' + error.toString());
    return false;
  }
}

/**
 * Process order row
 */
function processOrderRow(sheet, row, mapping) {
  var lock = LockService.getDocumentLock();
  
  try {
    if (!lock.tryLock(10000)) {
      throw new Error('Could not acquire lock after 10 seconds');
    }
    
    if (!sheet || !row || row < 2 || !mapping || !mapping.email || !mapping.ordered) {
      throw new Error('Invalid parameters');
    }
    
    ensureHelperColumns(sheet, mapping);
    mapping = getColumnMapping(sheet);
    
    var rowData = getRowData(sheet, row, mapping);
    if (!rowData) {
      throw new Error('Failed to extract row data');
    }
    
    var validationErrors = validateRowData(rowData);
    if (validationErrors.length > 0) {
      throw new Error('Validation failed: ' + validationErrors.join('; '));
    }
    
    if (rowData.notified && rowData.notified !== '') {
      throw new Error('Already notified on: ' + rowData.notified);
    }
    
    if (isDuplicateEmail(rowData.email, rowData.po, rowData.description)) {
      throw new Error('Duplicate email prevented - same order already sent recently');
    }
    
    checkRateLimit();
    
    var emailData = buildEmailContent(rowData);
    if (!emailData || !emailData.subject || !emailData.htmlBody) {
      throw new Error('Failed to build email content');
    }
    
    var attachments = [];
    if (CONFIG.ATTACH_QUOTE_PDF && rowData.quote) {
      try {
        var attachment = getQuoteAttachment(rowData.quote);
        if (attachment) {
          attachments.push(attachment);
        }
      } catch (attachError) {
        logToSheet('Attachment warning: ' + attachError.toString());
      }
    }
    
    var replyTo = getReplyToEmail();
    sendOrderEmail(rowData.email, emailData, attachments, replyTo);
    
    logEmailAttempt(rowData.email, rowData.po, rowData.description, 'SUCCESS', null);
    
    try {
      var now = new Date();
      sheet.getRange(row, mapping.notified).setValue(now);
      
      var messageKey = generateMessageKey(rowData.email, rowData.po, rowData.description);
      if (mapping.messageKey) {
        sheet.getRange(row, mapping.messageKey).setValue(messageKey);
      }
    } catch (updateError) {
      logToSheet('Warning: Failed to update tracking columns: ' + updateError.toString());
    }
    
    logToSheet('SUCCESS: Email sent to ' + rowData.email + ' for row ' + row);
    
  } catch (error) {
    logToSheet('ERROR row ' + row + ': ' + error.toString());
    
    try {
      var errorData = getRowData(sheet, row, mapping);
      if (errorData) {
        logEmailAttempt(errorData.email, errorData.po, errorData.description, 'FAILED', error.toString());
      }
    } catch (logError) {
      // Ignore logging errors
    }
    
    try {
      if (sheet && mapping && mapping.notified) {
        var errorMsg = error.toString();
        if (errorMsg.length > 100) {
          errorMsg = errorMsg.substring(0, 97) + '...';
        }
        sheet.getRange(row, mapping.notified).setValue('ERROR: ' + errorMsg);
      }
    } catch (markError) {
      logToSheet('Could not mark error in sheet: ' + markError.toString());
    }
    
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (releaseError) {
        logToSheet('Warning: Failed to release lock: ' + releaseError.toString());
      }
    }
  }
}

/**
 * Validate row data
 */
function validateRowData(data) {
  var errors = [];
  
  if (!data) {
    errors.push('No data provided');
    return errors;
  }
  
  if (!data.email) {
    errors.push('Email is required');
  } else if (!isValidEmail(data.email)) {
    errors.push('Invalid email format: ' + data.email);
  }
  
  if (!data.po) {
    errors.push('PO number is missing');
  }
  
  if (!data.description) {
    errors.push('Description is missing');
  }
  
  return errors;
}

/**
 * Configure sender email settings
 */
function configureSenderEmail() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    var currentSender = '';
    try {
      currentSender = getSenderEmail();
    } catch (e) {
      currentSender = 'Error: ' + e.message;
    }
    
    var message = 'EMAIL SENDER CONFIGURATION\n\n';
    message += 'Current sender: ' + currentSender + '\n\n';
    message += 'Choose how emails should be sent:\n\n';
    message += '1. AUTO - Use current user email (recommended)\n';
    message += '2. CONFIG - Use email from Config sheet\n';
    message += '3. SPECIFIC - Use a specific email address\n\n';
    message += 'Which option would you like to use?';
    
    var response = ui.alert('Configure Sender Email', message, ui.ButtonSet.YES_NO_CANCEL);
    
    var selectedOption = '';
    if (response === ui.Button.YES) {
      selectedOption = 'auto';
    } else if (response === ui.Button.NO) {
      var optionResponse = ui.prompt('Sender Email Options', 
        'Enter your choice:\n\n' +
        '1 = AUTO (current user)\n' +
        '2 = CONFIG (from Config sheet)\n' +
        '3 = SPECIFIC (enter email address)\n\n' +
        'Enter 1, 2, or 3:', 
        ui.ButtonSet.OK_CANCEL);
      
      if (optionResponse.getSelectedButton() === ui.Button.OK) {
        var choice = optionResponse.getResponseText().trim();
        
        if (choice === '1') {
          selectedOption = 'auto';
        } else if (choice === '2') {
          selectedOption = 'config';
          var configSheet = SpreadsheetApp.getActive().getSheetByName('Config');
          if (!configSheet) {
            var createSheet = ui.alert('Create Config Sheet', 
              'Config sheet does not exist. Create it now?', 
              ui.ButtonSet.YES_NO);
            if (createSheet === ui.Button.YES) {
              configSheet = SpreadsheetApp.getActive().insertSheet('Config');
              configSheet.getRange(1, 1, 2, 2).setValues([
                ['SENDER_EMAIL', 'your-email@example.com'],
                ['REPLY_TO', 'your-reply-email@example.com']
              ]);
              configSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
              ui.alert('Config Sheet Created', 
                'Config sheet created with example values. Please edit the SENDER_EMAIL value.', 
                ui.ButtonSet.OK);
            }
          } else {
            ui.alert('Using Config Sheet', 
              'Make sure the Config sheet has a row with:\nSENDER_EMAIL | your-email@domain.com', 
              ui.ButtonSet.OK);
          }
        } else if (choice === '3') {
          var emailResponse = ui.prompt('Specific Email Address', 
            'Enter the email address to send from:', 
            ui.ButtonSet.OK_CANCEL);
          if (emailResponse.getSelectedButton() === ui.Button.OK) {
            var emailAddress = emailResponse.getResponseText().trim();
            if (isValidEmail(emailAddress)) {
              selectedOption = emailAddress;
            } else {
              ui.alert('Invalid Email', 'Please enter a valid email address.', ui.ButtonSet.OK);
              return;
            }
          }
        } else {
          ui.alert('Invalid Choice', 'Please enter 1, 2, or 3.', ui.ButtonSet.OK);
          return;
        }
      }
    } else {
      return;
    }
    
    if (selectedOption) {
      logToSheet('Sender email configuration changed to: ' + selectedOption);
      CONFIG.EMAIL_SENDER = selectedOption;
      
      var confirmMessage = 'Sender email configuration updated!\n\n';
      confirmMessage += 'New setting: ' + selectedOption + '\n\n';
      
      try {
        var newSender = getSenderEmail();
        confirmMessage += 'Resolved sender email: ' + newSender + '\n\n';
        confirmMessage += 'Note: This setting will reset when the script reloads. For permanent changes, modify the CONFIG.EMAIL_SENDER value in the code.';
      } catch (e) {
        confirmMessage += 'Error testing new configuration: ' + e.message;
      }
      
      ui.alert('Configuration Updated', confirmMessage, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Configuration Error', 'Failed to configure sender email: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * View current configuration
 */
function viewCurrentConfig() {
  try {
    var sheet = getMainSheet();
    var message = 'CURRENT SYSTEM CONFIGURATION\n\n';
    
    if (sheet) {
      message += 'Main Sheet: "' + sheet.getName() + '"\n';
      var mapping = getColumnMapping(sheet);
      if (mapping) {
        message += 'Email Column: ' + (mapping.email ? 'Column ' + mapping.email : 'Not found') + '\n';
        message += 'Order Status Column: ' + (mapping.ordered ? 'Column ' + mapping.ordered : 'Not found') + '\n';
        message += 'PO Number Column: ' + (mapping.po ? 'Column ' + mapping.po : 'Not found') + '\n';
      }
    } else {
      message += 'Main Sheet: Not found\n';
    }
    
    message += '\n';
    
    message += 'Sender Config: ' + CONFIG.EMAIL_SENDER + '\n';
    try {
      var currentSender = getSenderEmail();
      message += 'Active Sender: ' + currentSender + '\n';
    } catch (e) {
      message += 'Active Sender: ERROR - ' + e.message + '\n';
    }
    
    message += '\n';
    
    message += 'SAFETY SETTINGS:\n';
    message += 'Rate Limits: ' + CONFIG.MAX_EMAILS_PER_HOUR + '/hour, ' + CONFIG.MAX_EMAILS_PER_DAY + '/day\n';
    message += 'Duplicate Check: ' + CONFIG.DUPLICATE_CHECK_DAYS + ' days\n';
    message += 'Attach PDFs: ' + (CONFIG.ATTACH_QUOTE_PDF ? 'Yes' : 'No') + '\n';
    
    message += '\n';
    
    try {
      var currentUser = Session.getActiveUser().getEmail();
      message += 'Current User: ' + currentUser + '\n';
    } catch (e) {
      message += 'Current User: Could not determine\n';
    }
    
    // Skip trigger check to avoid permission error
    message += 'onEdit Trigger: Check manually in Apps Script > Triggers\n';
    
    SpreadsheetApp.getUi().alert('Current Configuration', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Configuration Error', 'Failed to get current configuration: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Include all other necessary functions with proper syntax...
// (I'll continue with the rest in a separate response due to length limits)

/**
 * Build email content
 */
function buildEmailContent(data) {
  var greeting = data.name ? 'Hello ' + data.name + ',' : CONFIG.EMAIL_GREETING + ',';
  
  var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">';
  htmlBody += '<p>' + greeting + '</p>';
  htmlBody += '<p>Great news! Your order has been placed and is being processed.</p>';
  htmlBody += '<table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; margin: 20px 0; width: 100%;">';
  htmlBody += '<tr style="background-color: #f0f0f0;"><th colspan="2" style="text-align: left; padding: 12px;">Order Details</th></tr>';
  
  if (data.po) {
    htmlBody += '<tr><td style="font-weight: bold; background-color: #f9f9f9;">PO Number:</td><td>' + escapeHtml(data.po) + '</td></tr>';
  }
  
  if (data.description) {
    htmlBody += '<tr><td style="font-weight: bold; background-color: #f9f9f9;">Description:</td><td>' + escapeHtml(data.description) + '</td></tr>';
  }
  
  htmlBody += '<tr><td style="font-weight: bold; background-color: #f9f9f9;">Status:</td><td style="color: green; font-weight: bold;">Order Placed</td></tr>';
  htmlBody += '<tr><td style="font-weight: bold; background-color: #f9f9f9;">Notification Date:</td><td>' + new Date().toLocaleString() + '</td></tr>';
  htmlBody += '</table>';
  htmlBody += '<p>If you have any questions about your order, please reply to this email.</p>';
  htmlBody += '<p style="margin-top: 30px;">' + CONFIG.EMAIL_SIGNATURE + '</p>';
  htmlBody += '</div>';
  
  var textBody = greeting + '\n\n';
  textBody += 'Great news! Your order has been placed and is being processed.\n\n';
  textBody += 'ORDER DETAILS:\n';
  textBody += '==============\n';
  
  if (data.po) {
    textBody += 'PO Number: ' + data.po + '\n';
  }
  
  if (data.description) {
    textBody += 'Description: ' + data.description + '\n';
  }
  
  textBody += 'Status: Order Placed\n';
  textBody += 'Notification Date: ' + new Date().toLocaleString() + '\n\n';
  textBody += 'If you have any questions about your order, please reply to this email.\n\n';
  textBody += CONFIG.EMAIL_SIGNATURE;
  
  return {
    subject: CONFIG.EMAIL_SUBJECT,
    htmlBody: htmlBody,
    textBody: textBody
  };
}

/**
 * Send email using Gmail API
 */
function sendOrderEmail(toEmail, emailData, attachments, replyTo) {
  try {
    var boundary = 'boundary' + Date.now();
    var nl = '\r\n';
    
    var headers = [];
    headers.push('To: ' + toEmail);
    headers.push('Subject: ' + emailData.subject);
    headers.push('MIME-Version: 1.0');
    
    if (replyTo && isValidEmail(replyTo)) {
      headers.push('Reply-To: ' + replyTo);
    }
    
    var body = '';
    
    if (attachments.length > 0) {
      headers.push('Content-Type: multipart/mixed; boundary="' + boundary + '"');
      
      body += '--' + boundary + nl;
      body += 'Content-Type: multipart/alternative; boundary="' + boundary + 'alt"' + nl + nl;
      
      body += '--' + boundary + 'alt' + nl;
      body += 'Content-Type: text/plain; charset=UTF-8' + nl + nl;
      body += emailData.textBody + nl + nl;
      
      body += '--' + boundary + 'alt' + nl;
      body += 'Content-Type: text/html; charset=UTF-8' + nl + nl;
      body += emailData.htmlBody + nl + nl;
      
      body += '--' + boundary + 'alt--' + nl;
      
      for (var i = 0; i < attachments.length; i++) {
        var attachment = attachments[i];
        var filename = attachment.getName() || 'attachment.pdf';
        
        body += '--' + boundary + nl;
        body += 'Content-Type: ' + attachment.getContentType() + nl;
        body += 'Content-Transfer-Encoding: base64' + nl;
        body += 'Content-Disposition: attachment; filename="' + filename + '"' + nl + nl;
        body += Utilities.base64Encode(attachment.getBytes()) + nl + nl;
      }
      
      body += '--' + boundary + '--';
      
    } else {
      headers.push('Content-Type: multipart/alternative; boundary="' + boundary + '"');
      
      body += '--' + boundary + nl;
      body += 'Content-Type: text/plain; charset=UTF-8' + nl + nl;
      body += emailData.textBody + nl + nl;
      
      body += '--' + boundary + nl;
      body += 'Content-Type: text/html; charset=UTF-8' + nl + nl;
      body += emailData.htmlBody + nl + nl;
      
      body += '--' + boundary + '--';
    }
    
    var rawMessage = headers.join(nl) + nl + nl + body;
    var encodedMessage = Utilities.base64EncodeWebSafe(rawMessage);
    
    Gmail.Users.Messages.send({
      raw: encodedMessage
    }, 'me');
    
  } catch (error) {
    throw new Error('Failed to send email: ' + error.toString());
  }
}

/**
 * Check for duplicate emails
 */
function isDuplicateEmail(email, po, description) {
  try {
    if (!email) return false;
    
    var messageKey = generateMessageKey(email, po, description);
    var sheet = getMainSheet();
    if (!sheet) return false;
    
    var mapping = getColumnMapping(sheet);
    if (!mapping || !mapping.messageKey || !mapping.notified) return false;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    
    var cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - CONFIG.DUPLICATE_CHECK_DAYS);
    
    var batchSize = 100;
    for (var startRow = 2; startRow <= lastRow; startRow += batchSize) {
      var endRow = Math.min(startRow + batchSize - 1, lastRow);
      var range = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn());
      var values = range.getValues();
      
      for (var i = 0; i < values.length; i++) {
        var rowMessageKey = cleanString(values[i][mapping.messageKey - 1]);
        var rowNotified = values[i][mapping.notified - 1];
        
        if (rowMessageKey === messageKey && rowNotified) {
          var notifiedDate = null;
          
          if (rowNotified instanceof Date) {
            notifiedDate = rowNotified;
          } else if (typeof rowNotified === 'string' && rowNotified.trim()) {
            notifiedDate = new Date(rowNotified.trim());
          }
          
          if (notifiedDate && notifiedDate > cutoffDate) {
            logToSheet('DUPLICATE DETECTED: MessageKey ' + messageKey + ' sent on ' + notifiedDate.toLocaleString());
            return true;
          }
        }
      }
    }
    
    return false;
    
  } catch (error) {
    logToSheet('isDuplicateEmail ERROR: ' + error.toString());
    return false;
  }
}

/**
 * Rate limiting
 */
function checkRateLimit() {
  try {
    var ss = SpreadsheetApp.getActive();
    var rateLimitSheet = ss.getSheetByName(CONFIG.RATE_LIMIT_SHEET_NAME);
    
    if (!rateLimitSheet) {
      rateLimitSheet = ss.insertSheet(CONFIG.RATE_LIMIT_SHEET_NAME);
      rateLimitSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Type', 'Count']]);
      rateLimitSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      logToSheet('Created rate limit tracking sheet');
    }
    
    var now = new Date();
    var oneHourAgo = new Date(now.getTime() - (60 * 60 * 1000));
    var oneDayAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000));
    
    var lastRow = rateLimitSheet.getLastRow();
    var hourlyCount = 0;
    var dailyCount = 0;
    
    if (lastRow > 1) {
      var checkRows = Math.min(200, lastRow - 1);
      var data = rateLimitSheet.getRange(lastRow - checkRows + 1, 1, checkRows, 3).getValues();
      
      for (var i = 0; i < data.length; i++) {
        var timestamp = data[i][0];
        if (timestamp instanceof Date) {
          if (timestamp > oneHourAgo) {
            hourlyCount++;
          }
          if (timestamp > oneDayAgo) {
            dailyCount++;
          }
        }
      }
    }
    
    if (hourlyCount >= CONFIG.MAX_EMAILS_PER_HOUR) {
      throw new Error('Hourly email limit exceeded: ' + hourlyCount + '/' + CONFIG.MAX_EMAILS_PER_HOUR);
    }
    
    if (dailyCount >= CONFIG.MAX_EMAILS_PER_DAY) {
      throw new Error('Daily email limit exceeded: ' + dailyCount + '/' + CONFIG.MAX_EMAILS_PER_DAY);
    }
    
    rateLimitSheet.appendRow([now, 'EMAIL_SENT', 1]);
    
    var maxEntries = 500;
    if (rateLimitSheet.getLastRow() > maxEntries) {
      var excessRows = rateLimitSheet.getLastRow() - maxEntries;
      rateLimitSheet.deleteRows(2, excessRows);
    }
    
    logToSheet('Rate limit check passed - Hourly: ' + (hourlyCount + 1) + '/' + CONFIG.MAX_EMAILS_PER_HOUR + 
               ', Daily: ' + (dailyCount + 1) + '/' + CONFIG.MAX_EMAILS_PER_DAY);
    
    return true;
    
  } catch (error) {
    logToSheet('Rate limit check failed: ' + error.toString());
    throw error;
  }
}

/**
 * Log email attempts
 */
function logEmailAttempt(email, po, description, status, errorMessage) {
  try {
    var ss = SpreadsheetApp.getActive();
    var backupSheet = ss.getSheetByName(CONFIG.BACKUP_SHEET_NAME);
    
    if (!backupSheet) {
      backupSheet = ss.insertSheet(CONFIG.BACKUP_SHEET_NAME);
      var headers = ['Timestamp', 'Email', 'PO Number', 'Description', 'Status', 'Error Message', 'MessageKey'];
      backupSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      backupSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      
      backupSheet.setColumnWidth(1, 150);
      backupSheet.setColumnWidth(2, 200);
      backupSheet.setColumnWidth(3, 120);
      backupSheet.setColumnWidth(4, 300);
      backupSheet.setColumnWidth(5, 100);
      backupSheet.setColumnWidth(6, 300);
      backupSheet.setColumnWidth(7, 150);
      
      logToSheet('Created email backup tracking sheet');
    }
    
    var messageKey = generateMessageKey(email, po, description);
    
    backupSheet.appendRow([
      new Date(),
      email,
      po || '',
      description || '',
      status,
      errorMessage || '',
      messageKey
    ]);
    
    var lastRow = backupSheet.getLastRow();
    if (lastRow > 1001) {
      backupSheet.deleteRows(2, lastRow - 1001);
    }
    
  } catch (error) {
    logToSheet('Failed to log email attempt: ' + error.toString());
  }
}

/**
 * Utility functions
 */
function getQuoteAttachment(quoteText) {
  try {
    if (!quoteText) return null;
    
    var fileId = extractFileId(quoteText);
    if (!fileId) {
      logToSheet('No valid Drive file ID found in: ' + quoteText);
      return null;
    }
    
    try {
      var file = DriveApp.getFileById(fileId);
    } catch (driveError) {
      if (driveError.toString().indexOf('permissions') !== -1) {
        logToSheet('PERMISSION ERROR: Drive access not authorized.');
        return null;
      } else {
        throw driveError;
      }
    }
    
    var sizeMB = file.getSize() / (1024 * 1024);
    if (sizeMB > 25) {
      logToSheet('WARNING: File too large (' + sizeMB.toFixed(2) + 'MB): ' + file.getName());
      return null;
    }
    
    logToSheet('Attaching PDF: ' + file.getName() + ' (' + sizeMB.toFixed(2) + 'MB)');
    return file.getAs(MimeType.PDF);
    
  } catch (error) {
    logToSheet('Attachment error: ' + error.toString());
    return null;
  }
}

function extractFileId(text) {
  if (!text) return null;
  
  var match = text.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (match) return match[1];
  
  match = text.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (match) return match[1];
  
  if (/^[a-zA-Z0-9_-]{10,}$/.test(text.trim())) {
    return text.trim();
  }
  
  return null;
}

function getReplyToEmail() {
  try {
    var configSheet = SpreadsheetApp.getActive().getSheetByName('Config');
    if (!configSheet) return '';
    
    var data = configSheet.getRange(1, 1, 10, 2).getValues();
    
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0] || '').toUpperCase().trim();
      var value = String(data[i][1] || '').trim();
      
      if (key === 'REPLY_TO' && isValidEmail(value)) {
        return value;
      }
    }
    
    return '';
    
  } catch (error) {
    logToSheet('Config sheet error: ' + error.toString());
    return '';
  }
}

function generateMessageKey(email, po, description) {
  var text = email + '|' + po + '|' + description;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  return Utilities.base64Encode(digest).substring(0, 16);
}

function ensureHelperColumns(sheet, mapping) {
  var needsUpdate = false;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  if (!mapping.notified) {
    headers.push(CONFIG.H_NOTIFIED);
    needsUpdate = true;
  }
  
  if (!mapping.messageKey) {
    headers.push(CONFIG.H_MESSAGEKEY);
    needsUpdate = true;
  }
  
  if (needsUpdate) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function cleanString(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }
    
    var str = String(value).trim();
    
    if (str.startsWith('#')) {
      logToSheet('Warning: Cell contains formula error: ' + str);
      return '';
    }
    
    return str;
    
  } catch (error) {
    logToSheet('cleanString error: ' + error.toString());
    return '';
  }
}

function isValidEmail(email) {
  try {
    if (!email) return false;
    
    var trimmed = String(email).trim();
    
    if (trimmed.length < 5 || trimmed.length > CONFIG.MAX_EMAIL_LENGTH) {
      return false;
    }
    
    var pattern = /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;
    
    return pattern.test(trimmed);
    
  } catch (error) {
    logToSheet('isValidEmail error: ' + error.toString());
    return false;
  }
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function logToSheet(message) {
  try {
    console.log(message);
    
    var ss = SpreadsheetApp.getActive();
    var logSheet = ss.getSheetByName('Automation_Log');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('Automation_Log');
      logSheet.getRange(1, 1, 1, 2).setValues([['Timestamp', 'Message']]);
      logSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    }
    
    logSheet.appendRow([new Date(), message]);
    
    var lastRow = logSheet.getLastRow();
    if (lastRow > 501) {
      logSheet.deleteRows(2, lastRow - 501);
    }
    
  } catch (error) {
    console.error('Logging failed: ' + error.toString());
  }
}

/**
 * Setup and validation functions
 */
function setupDataValidation() {
  try {
    var sheet = getMainSheet();
    if (!sheet) {
      throw new Error('Could not find main data sheet');
    }
    
    var mapping = getColumnMapping(sheet);
    if (!mapping) {
      throw new Error('Could not determine column mapping');
    }
    
    var lastRow = Math.max(sheet.getLastRow(), 100);
    
    if (mapping.email) {
      var emailRule = SpreadsheetApp.newDataValidation()
        .requireTextIsEmail()
        .setAllowInvalid(false)
        .setHelpText('Please enter a valid email address')
        .build();
      sheet.getRange(2, mapping.email, lastRow-1, 1).setDataValidation(emailRule);
      logToSheet('Applied email validation to column ' + mapping.email);
    }
    
    if (mapping.ordered) {
      var statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['', 'Yes', 'No', 'Order Placed', 'Pending', 'Confirmed', 'Complete'])
        .setAllowInvalid(true)
        .setHelpText('Select order status')
        .build();
      sheet.getRange(2, mapping.ordered, lastRow-1, 1).setDataValidation(statusRule);
      logToSheet('Applied status validation to column ' + mapping.ordered);
    }
    
    logToSheet('Data validation setup completed for sheet: ' + sheet.getName());
    return true;
    
  } catch (error) {
    logToSheet('setupDataValidation ERROR: ' + error.toString());
    return false;
  }
}

function validateConfiguration() {
  try {
    var issues = [];
    
    if (typeof Gmail === 'undefined') {
      issues.push('Gmail API service not available');
    }
    
    var sheet = getMainSheet();
    if (!sheet) {
      issues.push('Could not find main data sheet');
    } else {
      logToSheet('Using main sheet: ' + sheet.getName());
      var mapping = getColumnMapping(sheet);
      if (!mapping) {
        issues.push('Could not determine column mapping');
      } else {
        if (!mapping.email) issues.push('Email column not found');
        if (!mapping.ordered) issues.push('Order status column not found');
        if (!mapping.po) issues.push('PO Number column not found');
      }
    }
    
    try {
      var senderEmail = getSenderEmail();
      if (!senderEmail) {
        issues.push('Could not determine sender email');
      } else {
        logToSheet('Sender email configured: ' + senderEmail);
      }
    } catch (senderError) {
      issues.push('Sender email error: ' + senderError.toString());
    }
    
    if (!CONFIG.EMAIL_SUBJECT || CONFIG.EMAIL_SUBJECT.trim() === '') {
      issues.push('EMAIL_SUBJECT is empty');
    }
    
    var triggers = ScriptApp.getProjectTriggers();
    var hasOnEditTrigger = false;
    
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onEdit' && 
          triggers[i].getEventType() === ScriptApp.EventType.ON_EDIT) {
        hasOnEditTrigger = true;
        break;
      }
    }
    
    if (!hasOnEditTrigger) {
      issues.push('No onEdit trigger found');
    }
    
    if (issues.length > 0) {
      logToSheet('CONFIGURATION ISSUES FOUND:');
      for (var j = 0; j < issues.length; j++) {
        logToSheet('- ' + issues[j]);
      }
      return false;
    } else {
      logToSheet('Configuration validation passed');
      return true;
    }
    
  } catch (error) {
    logToSheet('validateConfiguration ERROR: ' + error.toString());
    return false;
  }
}

function checkSetup() {
  try {
    var issues = [];
    var warnings = [];
    
    var configValid = validateConfiguration();
    if (!configValid) {
      issues.push('Configuration validation failed - check logs');
    }
    
    if (typeof Gmail === 'undefined') {
      issues.push('Gmail API service not added');
    }
    
    var sheet = getMainSheet();
    if (!sheet) {
      issues.push('Could not find main data sheet');
    } else {
      var mapping = getColumnMapping(sheet);
      if (!mapping || !mapping.email) {
        issues.push('Email column not found');
      }
      if (!mapping || !mapping.ordered) {
        issues.push('Order Placed column not found');
      }
      
      if (!mapping || !mapping.notified) {
        warnings.push('Notified column will be created automatically');
      }
      if (!mapping || !mapping.messageKey) {
        warnings.push('MessageKey column will be created automatically');
      }
    }
    
    try {
      var senderEmail = getSenderEmail();
      logToSheet('Current sender email: ' + senderEmail);
    } catch (senderError) {
      issues.push('Sender email configuration issue: ' + senderError.message);
    }
    
    var message = '';
    
    if (issues.length > 0) {
      message += 'ISSUES FOUND:\n\n';
      for (var i = 0; i < issues.length; i++) {
        message += '❌ ' + issues[i] + '\n';
      }
      message += '\nPlease fix these issues before using the system.\n';
    } else {
      message += '✅ Setup validation passed!\n\n';
    }
    
    if (warnings.length > 0) {
      message += '\nINFO:\n';
      for (var j = 0; j < warnings.length; j++) {
        message += '⚠️ ' + warnings[j] + '\n';
      }
    }
    
    if (sheet) {
      message += '\nCurrent Configuration:\n';
      message += '• Main sheet: "' + sheet.getName() + '"\n';
      try {
        var currentSender = getSenderEmail();
        message += '• Sender email: ' + currentSender + '\n';
      } catch (e) {
        message += '• Sender email: ERROR - ' + e.message + '\n';
      }
    }
    
    message += '\nSafety features enabled:\n';
    message += '• Duplicate prevention (' + CONFIG.DUPLICATE_CHECK_DAYS + ' days)\n';
    message += '• Rate limiting (' + CONFIG.MAX_EMAILS_PER_HOUR + '/hour, ' + CONFIG.MAX_EMAILS_PER_DAY + '/day)\n';
    message += '• Auto sheet detection\n';
    message += '• Dynamic sender configuration\n';
    
    SpreadsheetApp.getUi().alert('Setup Check Results', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Setup check failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function testEmailForRow() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var mainSheet = getMainSheet();
    
    if (!mainSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Could not find main data sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (sheet.getName() !== mainSheet.getName()) {
      SpreadsheetApp.getUi().alert('Error', 'Please select the main data sheet "' + mainSheet.getName() + '" first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var activeRange = sheet.getActiveRange();
    if (!activeRange) {
      SpreadsheetApp.getUi().alert('Error', 'Please select a cell first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var row = activeRange.getRow();
    if (row === 1) {
      SpreadsheetApp.getUi().alert('Error', 'Please select a data row, not the header.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var mapping = getColumnMapping(sheet);
    if (!mapping || !mapping.email || !mapping.ordered) {
      SpreadsheetApp.getUi().alert('Error', 'Required columns not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var rowData = getRowData(sheet, row, mapping);
    if (!rowData) {
      SpreadsheetApp.getUi().alert('Error', 'Could not read row data.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var ui = SpreadsheetApp.getUi();
    var message = 'TEST EMAIL DETAILS:\n\n';
    message += 'Sheet: ' + sheet.getName() + '\n';
    message += 'Row: ' + row + '\n';
    message += 'Email: ' + (rowData.email || '(empty)') + '\n';
    message += 'Name: ' + (rowData.name || '(empty)') + '\n';
    message += 'PO: ' + (rowData.po || '(empty)') + '\n';
    
    try {
      var senderEmail = getSenderEmail();
      message += 'Sender: ' + senderEmail + '\n';
    } catch (senderError) {
      message += 'Sender: ERROR - ' + senderError.message + '\n';
    }
    
    message += '\nThis will send a REAL email. Continue?';
    
    var response = ui.alert('Confirm Test Email', message, ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    logToSheet('MANUAL TEST: Processing row ' + row);
    
    try {
      processOrderRow(sheet, row, mapping);
      logToSheet('TEST EMAIL: Sent to ' + rowData.email);
    } catch (error) {
      if (error.toString().indexOf('Duplicate email prevented') !== -1) {
        var overrideResponse = ui.alert('Duplicate Detected', 
          'This email was already sent recently. Send anyway for testing?', 
          ui.ButtonSet.YES_NO);
        
        if (overrideResponse === ui.Button.YES) {
          if (mapping.notified) {
            sheet.getRange(row, mapping.notified).setValue('');
          }
          processOrderRow(sheet, row, mapping);
        }
      } else {
        throw error;
      }
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Test Failed', 'Test failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function initializeSystem() {
  try {
    logToSheet('=== SYSTEM INITIALIZATION STARTED ===');
    
    var configValid = validateConfiguration();
    if (!configValid) {
      throw new Error('Configuration validation failed');
    }
    
    var validationSetup = setupDataValidation();
    if (validationSetup) {
      logToSheet('Data validation rules applied successfully');
    }
    
    var sheet = getMainSheet();
    if (sheet) {
      var mapping = getColumnMapping(sheet);
      if (mapping) {
        ensureHelperColumns(sheet, mapping);
        logToSheet('Helper columns ensured');
      }
    }
    
    logToSheet('=== SYSTEM INITIALIZATION COMPLETED ===');
    
    var message = 'System initialization completed successfully!\n\n';
    message += 'Configuration:\n';
    if (sheet) {
      message += '• Main sheet: "' + sheet.getName() + '"\n';
    }
    try {
      var senderEmail = getSenderEmail();
      message += '• Sender email: ' + senderEmail + '\n';
    } catch (e) {
      message += '• Sender email: ERROR - ' + e.message + '\n';
    }
    
    message += '\nSafety features enabled:\n';
    message += '• Duplicate email prevention\n';
    message += '• Rate limiting\n';
    message += '• Input validation\n';
    message += '• Auto sheet detection\n';
    message += '• Dynamic sender configuration\n\n';
    message += 'The system is now ready to use.';
    
    SpreadsheetApp.getUi().alert('System Initialized', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    logToSheet('INITIALIZATION ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert('Initialization Failed', 
      'System initialization failed: ' + error.toString(), 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function viewEmailBackupLog() {
  try {
    var backupSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.BACKUP_SHEET_NAME);
    
    if (!backupSheet) {
      SpreadsheetApp.getUi().alert('Info', 'No email backup log found yet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    backupSheet.activate();
    
    var lastRow = backupSheet.getLastRow();
    var message = 'Email backup log contains ' + Math.max(0, lastRow - 1) + ' entries.';
    
    SpreadsheetApp.getUi().alert('Email Backup Log', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to view backup log: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function viewRateLimitStatus() {
  try {
    var rateLimitSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.RATE_LIMIT_SHEET_NAME);
    
    if (!rateLimitSheet) {
      SpreadsheetApp.getUi().alert('Rate Limit Status', 'No emails sent yet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var now = new Date();
    var oneHourAgo = new Date(now.getTime() - (60 * 60 * 1000));
    var oneDayAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000));
    
    var lastRow = rateLimitSheet.getLastRow();
    var hourlyCount = 0;
    var dailyCount = 0;
    
    if (lastRow > 1) {
      var checkRows = Math.min(200, lastRow - 1);
      var data = rateLimitSheet.getRange(lastRow - checkRows + 1, 1, checkRows, 3).getValues();
      
      for (var i = 0; i < data.length; i++) {
        var timestamp = data[i][0];
        if (timestamp instanceof Date) {
          if (timestamp > oneHourAgo) {
            hourlyCount++;
          }
          if (timestamp > oneDayAgo) {
            dailyCount++;
          }
        }
      }
    }
    
    var message = 'CURRENT RATE LIMIT STATUS:\n\n';
    message += 'Last Hour: ' + hourlyCount + '/' + CONFIG.MAX_EMAILS_PER_HOUR + ' emails\n';
    message += 'Last 24 Hours: ' + dailyCount + '/' + CONFIG.MAX_EMAILS_PER_DAY + ' emails\n\n';
    
    if (hourlyCount >= CONFIG.MAX_EMAILS_PER_HOUR) {
      message += 'HOURLY LIMIT REACHED\n';
    } else if (dailyCount >= CONFIG.MAX_EMAILS_PER_DAY) {
      message += 'DAILY LIMIT REACHED\n';
    } else {
      message += 'Within limits - emails can be sent\n';
    }
    
    SpreadsheetApp.getUi().alert('Rate Limit Status', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to check rate limit status: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function resetRateLimits() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Reset Rate Limits', 
      'Are you sure you want to reset the rate limits?', 
      ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    var rateLimitSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.RATE_LIMIT_SHEET_NAME);
    
    if (rateLimitSheet) {
      var lastRow = rateLimitSheet.getLastRow();
      if (lastRow > 1) {
        rateLimitSheet.deleteRows(2, lastRow - 1);
      }
      logToSheet('ADMIN ACTION: Rate limits reset by user');
    }
    
    ui.alert('Rate Limits Reset', 'Rate limits have been reset.', ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to reset rate limits: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function auditDataIntegrity() {
  try {
    var sheet = getMainSheet();
    if (!sheet) {
      throw new Error('Could not find main data sheet');
    }
    
    var mapping = getColumnMapping(sheet);
    if (!mapping) {
      throw new Error('Could not determine column mapping');
    }
    
    var lastRow = sheet.getLastRow();
    var issues = [];
    var fixes = 0;
    
    logToSheet('=== DATA INTEGRITY AUDIT STARTED ===');
    logToSheet('Auditing sheet: ' + sheet.getName());
    
    for (var row = 2; row <= lastRow; row++) {
      var rowData = getRowData(sheet, row, mapping);
      
      if (rowData.email && !isValidEmail(rowData.email)) {
        issues.push('Row ' + row + ': Invalid email format - ' + rowData.email);
      }
      
      if (mapping.ordered) {
        var orderStatus = sheet.getRange(row, mapping.ordered).getValue();
        if (isYesValue(orderStatus) && !rowData.po) {
          issues.push('Row ' + row + ': Order marked as placed but no PO number');
        }
      }
      
      if (rowData.messageKey && !rowData.notified) {
        issues.push('Row ' + row + ': Has MessageKey but no notification timestamp');
      }
      
      if (rowData.email && rowData.email !== rowData.email.toLowerCase().trim()) {
        sheet.getRange(row, mapping.email).setValue(rowData.email.toLowerCase().trim());
        fixes++;
      }
    }
    
    logToSheet('Audit complete - Found ' + issues.length + ' issues, made ' + fixes + ' automatic fixes');
    
    if (issues.length > 0) {
      logToSheet('ISSUES FOUND:');
      for (var i = 0; i < Math.min(issues.length, 20); i++) {
        logToSheet('- ' + issues[i]);
      }
      if (issues.length > 20) {
        logToSheet('... and ' + (issues.length - 20) + ' more issues');
      }
    }
    
    logToSheet('=== DATA INTEGRITY AUDIT COMPLETED ===');
    
    var message = 'Data integrity audit completed for "' + sheet.getName() + '".\n\n';
    message += 'Found ' + issues.length + ' issues\n';
    message += 'Made ' + fixes + ' automatic fixes\n\n';
    message += 'Check the Automation_Log for detailed results.';
    
    SpreadsheetApp.getUi().alert('Audit Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    logToSheet('Data audit failed: ' + error.toString());
    SpreadsheetApp.getUi().alert('Audit Failed', 'Data integrity audit failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
