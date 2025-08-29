function logToSheet(message) {
  try {
    console.log(message);
    
    var ss = SpreadsheetApp.getActive();
    var logSheet = ss.getSheetByName('Automation_Log');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('Automation_Log');
      logSheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Message', 'GitHub Version']]);
      logSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      logSheet.setColumnWidth(1, 150);
      logSheet.setColumnWidth(2, 500);
      logSheet.setColumnWidth(3, 120);
      logSheet.getRange(1, 1, 1, 3).setBackground('#f0f0f0');
    }
    
    logSheet.appendRow([new Date(), message, '2.0-GitHub-DualLogo']);
    
    // Keep only the last 500 entries to prevent sheet from getting too large
    var lastRow = logSheet.getLastRow();
    if (lastRow > 501) {
      var excessRows = lastRow - 501;
      logSheet.deleteRows(2, excessRows);
    }
    
    // Auto-resize columns if needed
    logSheet.autoResizeColumns(1, 3);
    
  } catch (error) {
    console.error('Logging failed: ' + error.toString());
    // Fallback to console only if sheet logging fails
    try {
      console.log('FALLBACK LOG: ' + message);
    } catch (consoleError) {
      // If even console fails, fail silently to prevent infinite loops
    }
  }
}

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
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function cleanString(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }
    
    var str = String(value).trim();
    
    if (str.charAt(0) === '#') {
      logToSheet('Warning: Cell contains formula error: ' + str);
      return '';
    }
    
    return str;
    
  } catch (error) {
    logToSheet('cleanString error: ' + error.toString());
    return '';
  }
}