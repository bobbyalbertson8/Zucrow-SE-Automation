function generateMessageKey(email, po, description) {
  var text = email + '|' + po + '|' + description;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  return Utilities.base64Encode(digest).substring(0, 16);
}

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
        
        if (rowMessageKey === messageKey && rowNotified && String(rowNotified).slice(0,6).toUpperCase() !== 'ERROR:') {
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
