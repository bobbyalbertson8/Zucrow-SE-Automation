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

function logEmailAttempt(email, po, description, status, errorMessage) {
  try {
    var ss = SpreadsheetApp.getActive();
    var backupSheet = ss.getSheetByName(CONFIG.BACKUP_SHEET_NAME);
    
    if (!backupSheet) {
      backupSheet = ss.insertSheet(CONFIG.BACKUP_SHEET_NAME);
      var headers = ['Timestamp', 'Email', 'PO Number', 'Description', 'Status', 'Error Message', 'MessageKey', 'GitHub Version'];
      backupSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      backupSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      
      backupSheet.setColumnWidth(1, 150);
      backupSheet.setColumnWidth(2, 200);
      backupSheet.setColumnWidth(3, 120);
      backupSheet.setColumnWidth(4, 300);
      backupSheet.setColumnWidth(5, 100);
      backupSheet.setColumnWidth(6, 300);
      backupSheet.setColumnWidth(7, 150);
      backupSheet.setColumnWidth(8, 120);
      
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
      messageKey,
      '2.0-GitHub-DualLogo'
    ]);
    
    var lastRow = backupSheet.getLastRow();
    if (lastRow > 1001) {
      backupSheet.deleteRows(2, lastRow - 1001);
    }
    
  } catch (error) {
    logToSheet('Failed to log email attempt: ' + error.toString());
  }
}