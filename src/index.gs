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
    
    if (rowData.notified && rowData.notified !== '' && String(rowData.notified).slice(0,6).toUpperCase() !== 'ERROR:') {
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
    
    logToSheet('SUCCESS: Email sent to ' + rowData.email + ' for row ' + row + ' (GitHub-powered)');
    
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
