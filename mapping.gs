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
      
      // Skip system sheets
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
    
    // Fallback options
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

function normalizeHeader(header) {
  if (!header) return '';
  return String(header)
    .toLowerCase()
    .trim()
    .replace(/[^\w\s]/g, '')
    .replace(/\s+/g, ' ');
}

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
    logToSheet('Added helper columns to sheet');
  }
}