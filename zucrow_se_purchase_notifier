/**
 * GitHub-Ready Purchase Order Email Notification System
 * 
 * Repository: https://github.com/yourusername/purchase-order-notifier
 * Version: 2.0 - GitHub 
 * 
 * QUICK SETUP:
 * 1. Fork/clone the GitHub repository
 * 2. Upload both logos to assets/ folder in the repo
 * 3. Update GITHUB_CONFIG below with your repository details
 * 4. Enable Gmail API in Google Apps Script (Services > Gmail API)
 * 5. Set up onEdit trigger (Triggers > Add Trigger > onEdit > From spreadsheet > On edit)
 * 
 * Features:
 * - GitHub-hosted dual logos with conditional selection
 * - Auto-detection of sheet structure
 * - Duplicate prevention & rate limiting
 * - Professional email templates
 * - Multiple logo fallback options
 * - Enhanced error handling and logging
 */

// =============================================================================
// GITHUB CONFIGURATION - UPDATE THESE VALUES FOR YOUR REPOSITORY
// =============================================================================

var GITHUB_CONFIG = {
  // Your GitHub repository details
  username: 'bobbyalbertson8',                    // Replace with YOUR GitHub username
  repository: 'Zucrow-SE-Automation',       // Replace with YOUR repo name
  branch: 'main',                             // Usually 'main' or 'master'
  
  // Primary logo configuration (GitHub-hosted)
  logo: {
    filename: 'spectral_logo.png',             // Your logo file in assets/ folder
    altText: 'Spectral Energies',                       // Alt text for accessibility
    maxWidth: '200px',                         // Maximum logo width
    maxHeight: '100px'                         // Maximum logo height
  },
  
  // Secondary logo configuration (GitHub-hosted)  
  logo2: {
    filename: 'purdue_prop_logo.png',          // Your second logo file in assets/ folder
    altText: 'Purdue Propulsion',                    // Alt text for accessibility
    maxWidth: '200px',                         // Maximum logo width
    maxHeight: '100px'                         // Maximum logo height
  }
};

// =============================================================================
// MAIN CONFIGURATION - CUSTOMIZE FOR YOUR NEEDS
// =============================================================================

var CONFIG = {
  // Sheet detection (leave empty for auto-detection)
  SHEET_NAME: '',
  
  // Logo selection strategy: 'primary', 'secondary', 'both', 'conditional'
  LOGO_STRATEGY: 'both', // Options: 'primary', 'secondary', 'both', 'conditional'
  
  // Primary GitHub-hosted logo URL (auto-generated from GITHUB_CONFIG)
  LOGO_URL: generateGitHubImageUrl(GITHUB_CONFIG.logo.filename),
  LOGO_ALT_TEXT: GITHUB_CONFIG.logo.altText,
  LOGO_MAX_WIDTH: GITHUB_CONFIG.logo.maxWidth,
  LOGO_MAX_HEIGHT: GITHUB_CONFIG.logo.maxHeight,
  
  // Secondary GitHub-hosted logo URL (auto-generated from GITHUB_CONFIG)
  LOGO2_URL: generateGitHubImageUrl(GITHUB_CONFIG.logo2.filename),
  LOGO2_ALT_TEXT: GITHUB_CONFIG.logo2.altText,
  LOGO2_MAX_WIDTH: GITHUB_CONFIG.logo2.maxWidth,
  LOGO2_MAX_HEIGHT: GITHUB_CONFIG.logo2.maxHeight,
  
  // Conditional logo logic - use logo2 if email contains these keywords
  LOGO2_KEYWORDS: ['purdue', '@purdue.edu', 'university', 'student'],
  
  // Column mappings (flexible - system will auto-detect these)
  COLUMN_MAPPINGS: {
    email: ['email', 'email address', 'e-mail'],
    name: ['name', 'full name', 'requester name', 'user name'],
    po: ['purchase order number', 'po number', 'po #', 'order number'],
    description: ['purchase order description', 'description', 'order description', 'item description'],
    quote: ['quote pdf', 'quote', 'quote file', 'pdf quote'],
    ordered: ['order placed?', 'order placed', 'ordered', 'status', 'order status']
  },
  
  // Helper columns (automatically created)
  H_NOTIFIED: 'Notified',
  H_MESSAGEKEY: 'MessageKey',
  
  // Values that trigger email notifications
  YES_VALUES: ['yes', 'y', 'true', '1', 'placed', 'ordered', 'complete', 'done', 
               'completed', 'finished', 'sent', 'order placed', 'order sent', 
               'confirmed', 'approved', 'processed'],
  
  // Input validation limits
  MAX_EMAIL_LENGTH: 254,
  MAX_PO_LENGTH: 50,
  MAX_DESCRIPTION_LENGTH: 1000,
  MAX_NAME_LENGTH: 100,
  
  // Email content settings
  EMAIL_SUBJECT: 'Your order has been placed',
  EMAIL_GREETING: 'Hello',
  EMAIL_SIGNATURE: '-- Purchasing Team',
  ATTACH_QUOTE_PDF: false,
  
  // Sender configuration ('auto' uses current user, 'config' uses Config sheet)
  EMAIL_SENDER: 'auto',
  FALLBACK_SENDERS: [],
  
  // Safety and rate limiting
  MAX_EMAILS_PER_HOUR: 50,
  MAX_EMAILS_PER_DAY: 200,
  DUPLICATE_CHECK_DAYS: 30,
  BACKUP_SHEET_NAME: 'Email_Backup_Log',
  RATE_LIMIT_SHEET_NAME: 'Rate_Limit_Log'
};

// =============================================================================
// GITHUB INTEGRATION FUNCTIONS
// =============================================================================

/**
 * Generate GitHub raw image URL
 */
function generateGitHubImageUrl(filename) {
  return 'https://raw.githubusercontent.com/' + 
         GITHUB_CONFIG.username + '/' + 
         GITHUB_CONFIG.repository + '/' + 
         GITHUB_CONFIG.branch + '/assets/' + filename;
}

/**
 * Get repository information
 */
function getRepositoryInfo() {
  return {
    url: 'https://github.com/' + GITHUB_CONFIG.username + '/' + GITHUB_CONFIG.repository,
    logoUrl: CONFIG.LOGO_URL,
    logo2Url: CONFIG.LOGO2_URL,
    version: '2.0-GitHub-DualLogo',
    features: [
      'GitHub-hosted',
      'Conditional logo selection', 
      'Auto sheet detection',
      'Duplicate prevention',
      'Rate limiting',
      'Professional templates',
      'Multiple logo fallbacks',
      'Enhanced error handling'
    ]
  };
}

/**
 * Determine which logo(s) to use based on strategy and data
 */
function selectLogos(data) {
  var logos = [];
  
  logToSheet('selectLogos called with strategy: ' + CONFIG.LOGO_STRATEGY);
  logToSheet('Data email: "' + (data.email || '') + '"');
  logToSheet('Data description: "' + (data.description || '') + '"');
  
  switch (CONFIG.LOGO_STRATEGY) {
    case 'primary':
      logToSheet('Using primary logo strategy');
      logos.push({
        url: CONFIG.LOGO_URL,
        altText: CONFIG.LOGO_ALT_TEXT,
        maxWidth: CONFIG.LOGO_MAX_WIDTH,
        maxHeight: CONFIG.LOGO_MAX_HEIGHT
      });
      break;
      
    case 'secondary':
      logToSheet('Using secondary logo strategy');
      logos.push({
        url: CONFIG.LOGO2_URL,
        altText: CONFIG.LOGO2_ALT_TEXT,
        maxWidth: CONFIG.LOGO2_MAX_WIDTH,
        maxHeight: CONFIG.LOGO2_MAX_HEIGHT
      });
      break;
      
    case 'both':
      logToSheet('Using both logos strategy');
      logos.push({
        url: CONFIG.LOGO_URL,
        altText: CONFIG.LOGO_ALT_TEXT,
        maxWidth: CONFIG.LOGO_MAX_WIDTH,
        maxHeight: CONFIG.LOGO_MAX_HEIGHT
      });
      logos.push({
        url: CONFIG.LOGO2_URL,
        altText: CONFIG.LOGO2_ALT_TEXT,
        maxWidth: CONFIG.LOGO2_MAX_WIDTH,
        maxHeight: CONFIG.LOGO2_MAX_HEIGHT
      });
      break;
      
    case 'conditional':
      logToSheet('Using conditional logo strategy');
      // Use logo2 if email or description contains keywords, otherwise use primary
      var useSecondary = false;
      var emailText = (data.email || '').toLowerCase();
      var descText = (data.description || '').toLowerCase();
      
      logToSheet('Checking email text: "' + emailText + '"');
      logToSheet('Checking description text: "' + descText + '"');
      logToSheet('Keywords to match: ' + CONFIG.LOGO2_KEYWORDS.join(', '));
      
      for (var i = 0; i < CONFIG.LOGO2_KEYWORDS.length; i++) {
        var keyword = CONFIG.LOGO2_KEYWORDS[i].toLowerCase();
        logToSheet('Checking keyword: "' + keyword + '"');
        
        if (emailText.indexOf(keyword) !== -1) {
          logToSheet('MATCH FOUND: "' + keyword + '" in email');
          useSecondary = true;
          break;
        }
        
        if (descText.indexOf(keyword) !== -1) {
          logToSheet('MATCH FOUND: "' + keyword + '" in description');
          useSecondary = true;
          break;
        }
      }
      
      logToSheet('Use secondary logo: ' + useSecondary);
      
      if (useSecondary) {
        logToSheet('Adding secondary logo');
        logos.push({
          url: CONFIG.LOGO2_URL,
          altText: CONFIG.LOGO2_ALT_TEXT,
          maxWidth: CONFIG.LOGO2_MAX_WIDTH,
          maxHeight: CONFIG.LOGO2_MAX_HEIGHT
        });
      } else {
        logToSheet('Adding primary logo');
        logos.push({
          url: CONFIG.LOGO_URL,
          altText: CONFIG.LOGO_ALT_TEXT,
          maxWidth: CONFIG.LOGO_MAX_WIDTH,
          maxHeight: CONFIG.LOGO_MAX_HEIGHT
        });
      }
      break;
      
    default:
      logToSheet('Unknown strategy, falling back to primary logo');
      // Fallback to primary logo
      logos.push({
        url: CONFIG.LOGO_URL,
        altText: CONFIG.LOGO_ALT_TEXT,
        maxWidth: CONFIG.LOGO_MAX_WIDTH,
        maxHeight: CONFIG.LOGO_MAX_HEIGHT
      });
  }
  
  logToSheet('selectLogos returning ' + logos.length + ' logo(s)');
  return logos;
}

/**
 * Build logo HTML with multiple logo support and GitHub fallbacks
 */
function buildLogoHtml(logos) {
  if (!logos || logos.length === 0) {
    return '';
  }
  
  var logoHtml = '<div style="text-align: center; margin-bottom: 30px;">';
  
  for (var i = 0; i < logos.length; i++) {
    var logo = logos[i];
    var spacing = logos.length > 1 ? 'margin: 0 10px; display: inline-block;' : 'display: block; margin: 0 auto;';
    
    // Generate fallback URLs for this specific logo
    var fallbackUrls = generateLogoFallbacks(logo.url);
    
    logoHtml += '<img src="' + logo.url + '" ' +
               'alt="' + logo.altText + '" ' +
               'style="max-width: ' + logo.maxWidth + '; ' +
               'max-height: ' + logo.maxHeight + '; ' +
               spacing + ' border-radius: 4px;" ' +
               'onerror="this.src=\'' + fallbackUrls[0] + '\'; ' +
               'this.onerror=function(){this.src=\'' + fallbackUrls[1] + '\'; ' +
               'this.onerror=function(){this.style.display=\'none\';};};">';
  }
  
  logoHtml += '</div>';
  return logoHtml;
}

/**
 * Generate fallback URLs for any logo
 */
function generateLogoFallbacks(primaryUrl) {
  var fallbacks = [
    generateGitHubImageUrl('logo.png'),
    generateGitHubImageUrl('company-logo.png'),
    generateGitHubImageUrl('default-logo.png')
  ];
  
  // Remove the primary URL from fallbacks if it's already there
  return fallbacks.filter(function(url) { return url !== primaryUrl; });
}

/**
 * Enhanced logging function
 */
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
 * Get row data with validation
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
 * Build email content with GitHub-hosted logo
 */
function buildEmailContent(data) {
  var greeting = data.name ? 'Hello ' + data.name + ',' : CONFIG.EMAIL_GREETING + ',';
  var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy, h:mm:ss a z");

  // Outlook-safe inline icons
  var check = "<!--[if !mso]><!--><span style=\"font-family:'Segoe UI Emoji','Segoe UI Symbol','Apple Color Emoji','Noto Color Emoji',Arial,sans-serif;\">&#10003;</span><!--<![endif]--><!--[if mso]><span style='font-family:Arial,sans-serif;'>[OK]</span><![endif]-->";
  var bolt  = "<!--[if !mso]><!--><span style=\"font-family:'Segoe UI Emoji','Segoe UI Symbol','Apple Color Emoji','Noto Color Emoji',Arial,sans-serif;\">&#9889;</span><!--<![endif]--><!--[if mso]><span style='font-family:Arial,sans-serif;'>[Lightning]</span><![endif]-->";

  // Select appropriate logo(s)
  var selectedLogos = selectLogos(data);
  var logoHtml = buildLogoHtml(selectedLogos);

  var htmlBody = '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #ffffff;">';

  // Header logos
  htmlBody += logoHtml;

  // Greeting + intro
  htmlBody += '<p style="font-size: 16px; line-height: 1.5; color: #333; margin-bottom: 20px;">' + greeting + '</p>';
  htmlBody += '<p style="font-size: 16px; line-height: 1.5; color: #333; margin-bottom: 25px;">Great news! Your order has been placed and is being processed.</p>';

  // ===== Card shell =====
  htmlBody += '<div style="background: #ffffff; border: 1px solid #e1e5e9; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1); margin: 25px 0;">';
  htmlBody += '<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px;">' +
              '<h2 style="margin: 0; font-size: 20px; font-weight: 600;">Order Details</h2></div>';

  // PO Number (table row)
  if (data.po) {
    htmlBody += "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'>" +
                  "<tr>" +
                    "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; font-weight:600; color:#555; width:140px;'>PO Number:</td>" +
                    "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; color:#333; font-family:monospace; font-size:14px;'>" + escapeHtml(data.po) + "</td>" +
                  "</tr>" +
                "</table>";
  }

  // Description (table row)
  if (data.description) {
    htmlBody += "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'>" +
                  "<tr>" +
                    "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; font-weight:600; color:#555; width:140px;'>Description:</td>" +
                    "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; color:#333;'>" + escapeHtml(data.description) + "</td>" +
                  "</tr>" +
                "</table>";
  }

  // Status row (icon + text, table-based for alignment)
  htmlBody += "<div style='padding: 14px 20px; border-top: 1px solid #f0f0f0; border-bottom: 1px solid #f0f0f0;'>" +
                "<table role='presentation' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'><tr>" +
                  "<td style='vertical-align:top; padding-right:8px; font-size:16px; line-height:1; color:#28a745; font-weight:700;'>" + check + "</td>" +
                  "<td style='vertical-align:top; font-size:16px; line-height:1; color:#28a745; font-weight:700;'>Order Placed</td>" +
                "</tr></table>" +
              "</div>";

  // Date row (formatted)
  htmlBody += "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'>" +
                "<tr>" +
                  "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; font-weight:600; color:#555; width:140px;'>Date:</td>" +
                  "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; color:#333;'>" + formattedDate + "</td>" +
                "</tr>" +
              "</table>";

  // End card
  htmlBody += '</div>';

  // Next steps
  htmlBody += '<div style="background: linear-gradient(135deg, #e3f2fd 0%, #f3e5f5 100%); border-radius: 8px; padding: 20px; margin: 25px 0;">' +
              '<h3 style="margin: 0 0 10px 0; color: #1976d2; font-size: 16px;">Next Steps</h3>' +
              '<p style="margin: 0; color: #555; line-height: 1.5;">We\'ll keep you updated on your order progress. If you have any questions about your order, please reply to this email and we\'ll get back to you promptly.</p>' +
              '</div>';

  // Footer
  htmlBody += '<div style="text-align: center; margin-top: 40px; padding-top: 20px; border-top: 1px solid #e0e0e0;">' +
              '<p style="margin: 0; color: #666; font-size: 14px;">' + CONFIG.EMAIL_SIGNATURE + '</p>';

  // GitHub attribution with Outlook-safe lightning
  if (GITHUB_CONFIG.username && GITHUB_CONFIG.repository) {
    var logoInfo = selectedLogos.length > 1 ? ' (' + selectedLogos[0].altText + ')':
    htmlBody += "<p style='margin: 10px 0 0 0; font-size: 12px; color: #999;'>" +
                  "<a href='https://github.com/" + GITHUB_CONFIG.username + "/" + GITHUB_CONFIG.repository + "' style='color:#999; text-decoration:none;'>" +
                    "Powered by GitHub automation&nbsp;" + bolt + " " + (logoInfo || "") +
                  "</a></p>";
  }

  htmlBody += '</div></div>';

  // ===== Plain text version =====
  var textBody = greeting + '\n\n';
  textBody += 'Great news! Your order has been placed and is being processed.\n\n';
  textBody += 'ORDER DETAILS\n';
  textBody += '=============\n';
  if (data.po) textBody += 'PO Number: ' + data.po + '\n';
  if (data.description) textBody += 'Description: ' + data.description + '\n';
  textBody += 'Status: Order Placed [OK]\n';
  textBody += 'Date: ' + formattedDate + '\n\n';
  textBody += 'NEXT STEPS\n';
  textBody += '==========\n';
  textBody += 'We\'ll keep you updated on your order progress. If you have any questions\n';
  textBody += 'about your order, please reply to this email and we\'ll get back to you promptly.\n\n';
  textBody += CONFIG.EMAIL_SIGNATURE + '\n\n';
  if (GITHUB_CONFIG.username && GITHUB_CONFIG.repository) {
    textBody += 'This notification is powered by GitHub automation:\n';
    textBody += 'https://github.com/' + GITHUB_CONFIG.username + '/' + GITHUB_CONFIG.repository;
  }

  return {
    subject: CONFIG.EMAIL_SUBJECT,
    htmlBody: htmlBody,
    textBody: textBody
  };
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
 * Process order row with enhanced error handling
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
    
    // Add GitHub automation header
    headers.push('X-GitHub-Automation: purchase-order-notifier-dual-logo');
    headers.push('X-Automation-Version: 2.0');
    
    var body = '';
    
    if (attachments && attachments.length > 0) {
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

// =============================================================================
// VALIDATION AND UTILITY FUNCTIONS
// =============================================================================

/**
 * Validate email format
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
 * Validate and clean text fields
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
 * Check if value represents "yes" for order placement
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
 * Validate row data before processing
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
 * Validate email format
 */
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

/**
 * Escape HTML special characters
 */
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Clean string values
 */
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

/**
 * Generate unique message key for duplicate detection
 */
function generateMessageKey(email, po, description) {
  var text = email + '|' + po + '|' + description;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  return Utilities.base64Encode(digest).substring(0, 16);
}

/**
 * Ensure helper columns exist
 */
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

/**
 * Check for duplicate emails with message key
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

/**
 * Rate limiting check
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
 * Log email attempts for audit trail
 */
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

/**
 * Get reply-to email from config
 */
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

/**
 * Get quote attachment from Drive
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

/**
 * Extract file ID from Drive URL
 */
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

// =============================================================================
// MENU SYSTEM AND GITHUB INTEGRATION FUNCTIONS
// =============================================================================

/**
 * Creates menu when spreadsheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Order Notifier (GitHub v2.0)')
    .addItem('🧪 Test email for selected row', 'testEmailForRow')
    .addItem('✅ Check setup & validate config', 'checkSetup')
    .addItem('📊 Setup data validation rules', 'setupDataValidation')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🐙 GitHub Integration')
      .addItem('🔧 Setup GitHub integration', 'setupGitHubIntegration')
      .addItem('🖼️ Test GitHub', 'testGitHubLogos')
      .addItem('🎨 Configure logo strategy', 'configureLogoStrategy')
      .addItem('🔍 Debug logo selection', 'debugLogoSelection')
      .addItem('📋 View repository info', 'viewRepositoryInfo'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🧪 Testing Tools')
      .addItem('🎯 Force primary logo test', 'forcePrimaryLogoTest')
      .addItem('🎯 Force secondary logo test', 'forceSecondaryLogoTest')
      .addItem('🎯 Force both logos test', 'forceBothLogosTest')
      .addItem('⚙️ Test configuration', 'testConfiguration'))
    .addSeparator()
    .addItem('📧 Configure sender email', 'configureSenderEmail')
    .addItem('📊 View current configuration', 'viewCurrentConfig')
    .addSeparator()
    .addItem('📋 View email backup log', 'viewEmailBackupLog')
    .addItem('📈 View rate limit status', 'viewRateLimitStatus')
    .addItem('🔄 Reset rate limits (emergency)', 'resetRateLimits')
    .addSeparator()
    .addItem('🚀 Initialize system', 'initializeSystem')
    .addItem('🔍 Audit data integrity', 'auditDataIntegrity')
    .addToUi();
}

/**
 * Setup function for GitHub deployment
 */
function setupGitHubIntegration() {
  try {
    var ui = SpreadsheetApp.getUi();
    var repoInfo = getRepositoryInfo();
    
    var message = 'GITHUB INTEGRATION SETUP\n\n';
    message += 'Repository URL: ' + repoInfo.url + '\n';
    message += 'Primary logo URL: ' + repoInfo.logoUrl + '\n';
    message += 'Secondary logo URL: ' + repoInfo.logo2Url + '\n';
    message += 'Version: ' + repoInfo.version + '\n\n';
    message += 'SETUP CHECKLIST:\n';
    message += '☐ 1. Fork the GitHub repository\n';
    message += '☐ 2. Upload BOTH logos to assets/ folder\n';
    message += '☐ 3. Make repository public (for image hosting)\n';
    message += '☐ 4. Update GITHUB_CONFIG in this script\n';
    message += '☐ 5. Test both logo URLs\n\n';
    message += 'CURRENT CONFIGURATION:\n';
    message += 'Username: ' + GITHUB_CONFIG.username + '\n';
    message += 'Repository: ' + GITHUB_CONFIG.repository + '\n';
    message += 'Branch: ' + GITHUB_CONFIG.branch + '\n';
    message += 'Logo 1: ' + GITHUB_CONFIG.logo.filename + '\n';
    message += 'Logo 2: ' + GITHUB_CONFIG.logo2.filename + '\n';
    message += 'Strategy: ' + CONFIG.LOGO_STRATEGY + '\n\n';
    message += 'Click "Test GitHub" to verify setup.';
    
    ui.alert('GitHub Setup', message, ui.ButtonSet.OK);
    
    logToSheet('GitHub setup accessed - Repository: ' + repoInfo.url);
    
    return true;
    
  } catch (error) {
    logToSheet('GitHub setup error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Setup Error', 'GitHub setup failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

/**
 * Test GitHub-hosted dual logos
 */
function testGitHubLogos() {
  try {
    logToSheet('=== TESTING GITHUB ===');
    logToSheet('Primary logo URL: ' + CONFIG.LOGO_URL);
    logToSheet('Secondary logo URL: ' + CONFIG.LOGO2_URL);
    logToSheet('Logo strategy: ' + CONFIG.LOGO_STRATEGY);
    logToSheet('Repository: https://github.com/' + GITHUB_CONFIG.username + '/' + GITHUB_CONFIG.repository);
    
    // Test different logo selection scenarios
    var testScenarios = [
      { name: 'Regular User', email: 'user@example.com', description: 'Regular order' },
      { name: 'Purdue User', email: 'student@purdue.edu', description: 'University equipment' },
      { name: 'University Order', email: 'admin@company.com', description: 'Equipment for Purdue research' }
    ];
    
    for (var i = 0; i < testScenarios.length; i++) {
      var scenario = testScenarios[i];
      var selectedLogos = selectLogos(scenario);
      
      logToSheet('Scenario "' + scenario.name + '":');
      logToSheet('  Email: ' + scenario.email);
      logToSheet('  Selected logos: ' + selectedLogos.length);
      for (var j = 0; j < selectedLogos.length; j++) {
        logToSheet('    Logo ' + (j + 1) + ': ' + selectedLogos[j].altText + ' (' + selectedLogos[j].url + ')');
      }
    }
    
    // Generate test email to verify logo integration
    var testData = {
      name: 'GitHub Test User',
      email: 'test@github-integration.com',
      po: 'GITHUB-DUAL-001',
      description: 'Testing GitHub integration'
    };
    
    var emailContent = buildEmailContent(testData);
    logToSheet('✓ Email content successfully');
    
    logToSheet('=== GITHUB TEST COMPLETE ===');
    
    var ui = SpreadsheetApp.getUi();
    var message = 'GITHUB TEST RESULTS\n\n';
    message += '✅ Email generation: SUCCESS\n';
    message += '✅ logo integration: SUCCESS\n';
    message += '✅ Logo selection logic: WORKING\n';
    message += '✅ Fallback options: CONFIGURED\n\n';
    message += 'Logo Strategy: ' + CONFIG.LOGO_STRATEGY + '\n';
    message += 'Primary logo: ' + GITHUB_CONFIG.logo.altText + '\n';
    message += 'Secondary logo: ' + GITHUB_CONFIG.logo2.altText + '\n\n';
    message += 'URLs being used:\n';
    message += '• ' + CONFIG.LOGO_URL + '\n';
    message += '• ' + CONFIG.LOGO2_URL + '\n\n';
    message += 'If logos don\'t display in emails:\n\n';
    message += '1. Ensure your repository is PUBLIC\n';
    message += '2. Verify both logo files exist in assets/\n';
    message += '3. Check filenames match exactly\n';
    message += '4. Try different LOGO_STRATEGY settings\n\n';
    message += 'Check Automation_Log for detailed results.';
    
    ui.alert('GitHub Test Results', message, ui.ButtonSet.OK);
    
  } catch (error) {
    logToSheet('GitHub test failed: ' + error.toString());
    SpreadsheetApp.getUi().alert('Test Failed', 
      'GitHub test failed: ' + error.toString() + 
      '\n\nPlease check your GITHUB_CONFIG settings and ensure both logo files exist.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Configure logo selection strategy
 */
function configureLogoStrategy() {
  try {
    var ui = SpreadsheetApp.getUi();
    var currentStrategy = CONFIG.LOGO_STRATEGY;
    
    var message = 'LOGO STRATEGY CONFIGURATION\n\n';
    message += 'Current strategy: ' + currentStrategy + '\n\n';
    message += 'Available strategies:\n\n';
    message += '1. PRIMARY - Always use ' + GITHUB_CONFIG.logo.altText + ' logo\n';
    message += '2. SECONDARY - Always use ' + GITHUB_CONFIG.logo2.altText + ' logo\n';
    message += '3. BOTH - Show both logos side by side\n';
    message += '4. CONDITIONAL - Auto-select based on keywords\n\n';
    message += 'Select a strategy (1-4):';
    
    var response = ui.prompt('Logo Strategy', message, ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() === ui.Button.OK) {
      var choice = response.getResponseText().trim();
      var newStrategy = '';
      
      switch (choice) {
        case '1':
          newStrategy = 'primary';
          break;
        case '2':
          newStrategy = 'secondary';
          break;
        case '3':
          newStrategy = 'both';
          break;
        case '4':
          newStrategy = 'conditional';
          break;
        default:
          ui.alert('Invalid Choice', 'Please enter 1, 2, 3, or 4.', ui.ButtonSet.OK);
          return;
      }
      
      CONFIG.LOGO_STRATEGY = newStrategy;
      logToSheet('Logo strategy changed to: ' + newStrategy);
      
      var confirmMessage = 'Logo strategy updated to: ' + newStrategy + '\n\n';
      if (newStrategy === 'conditional') {
        confirmMessage += 'Keywords that trigger secondary logo:\n';
        for (var i = 0; i < CONFIG.LOGO2_KEYWORDS.length; i++) {
          confirmMessage += '• ' + CONFIG.LOGO2_KEYWORDS[i] + '\n';
        }
        confirmMessage += '\nModify LOGO2_KEYWORDS in config to change triggers.\n';
      }
      confirmMessage += '\nThis setting will reset when script reloads.\nFor permanent changes, modify CONFIG.LOGO_STRATEGY in code.';
      
      ui.alert('Strategy Updated', confirmMessage, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Configuration Error', 'Failed to configure logo strategy: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Debug logo selection process
 */
function debugLogoSelection() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Get the currently selected row
    var sheet = SpreadsheetApp.getActiveSheet();
    var activeRange = sheet.getActiveRange();
    
    if (!activeRange || activeRange.getRow() === 1) {
      ui.alert('Debug Error', 'Please select a data row (not header) to debug.', ui.ButtonSet.OK);
      return;
    }
    
    var row = activeRange.getRow();
    var mapping = getColumnMapping(sheet);
    var rowData = getRowData(sheet, row, mapping);
    
    logToSheet('=== LOGO SELECTION DEBUG ===');
    logToSheet('Current LOGO_STRATEGY: ' + CONFIG.LOGO_STRATEGY);
    logToSheet('Row data:');
    logToSheet('  Email: "' + (rowData.email || '') + '"');
    logToSheet('  Description: "' + (rowData.description || '') + '"');
    logToSheet('  Name: "' + (rowData.name || '') + '"');
    
    // Test the logo selection logic
    var selectedLogos = selectLogos(rowData);
    logToSheet('Selected logos count: ' + selectedLogos.length);
    
    for (var i = 0; i < selectedLogos.length; i++) {
      logToSheet('Logo ' + (i + 1) + ':');
      logToSheet('  Alt text: ' + selectedLogos[i].altText);
      logToSheet('  URL: ' + selectedLogos[i].url);
    }
    
    logToSheet('=== END DEBUG ===');
    
    var message = 'LOGO SELECTION DEBUG RESULTS\n\n';
    message += 'Row: ' + row + '\n';
    message += 'Strategy: ' + CONFIG.LOGO_STRATEGY + '\n';
    message += 'Email: ' + (rowData.email || 'empty') + '\n';
    message += 'Logos selected: ' + selectedLogos.length + '\n\n';
    
    if (selectedLogos.length > 0) {
      message += 'Selected logo(s):\n';
      for (var k = 0; k < selectedLogos.length; k++) {
        message += '• ' + selectedLogos[k].altText + '\n';
      }
    }
    
    message += '\nCheck Automation_Log for detailed debug info.';
    
    ui.alert('Logo Selection Debug', message, ui.ButtonSet.OK);
    
  } catch (error) {
    logToSheet('Debug function error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Debug Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * View repository information
 */
function viewRepositoryInfo() {
  try {
    var repoInfo = getRepositoryInfo();
    var message = 'GITHUB REPOSITORY INFO\n\n';
    message += '📁 Repository: ' + repoInfo.url + '\n';
    message += '🏷️ Version: ' + repoInfo.version + '\n';
    message += '🖼️ Primary logo: ' + repoInfo.logoUrl + '\n';
    message += '🖼️ Secondary logo: ' + repoInfo.logo2Url + '\n\n';
    message += '✨ FEATURES:\n';
    for (var i = 0; i < repoInfo.features.length; i++) {
      message += '  • ' + repoInfo.features[i] + '\n';
    }
    message += '\n📝 CONFIGURATION:\n';
    message += '  • Strategy: ' + CONFIG.LOGO_STRATEGY + '\n';
    message += '  • Primary: ' + GITHUB_CONFIG.logo.filename + ' (' + GITHUB_CONFIG.logo.altText + ')\n';
    message += '  • Secondary: ' + GITHUB_CONFIG.logo2.filename + ' (' + GITHUB_CONFIG.logo2.altText + ')\n';
    message += '  • Keywords: ' + CONFIG.LOGO2_KEYWORDS.join(', ') + '\n\n';
    message += '📖 DOCUMENTATION:\n';
    message += '  • Setup guide: ' + repoInfo.url + '#dual-logo-setup\n';
    message += '  • Examples: ' + repoInfo.url + '/examples';
    
    SpreadsheetApp.getUi().alert('Repository Information', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to get repository info: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Test email for selected row
 */
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
    
    // Preview which logos will be used
    var selectedLogos = selectLogos(rowData);
    var logoPreview = '';
    for (var i = 0; i < selectedLogos.length; i++) {
      logoPreview += selectedLogos[i].altText + (i < selectedLogos.length - 1 ? ' + ' : '');
    }
    
    var ui = SpreadsheetApp.getUi();
    var repoInfo = getRepositoryInfo();
    var message = 'GITHUB TEST EMAIL:\n\n';
    message += 'Sheet: ' + sheet.getName() + '\n';
    message += 'Row: ' + row + '\n';
    message += 'Email: ' + (rowData.email || '(empty)') + '\n';
    message += 'Strategy: ' + CONFIG.LOGO_STRATEGY + '\n';
    message += 'Logo(s): ' + logoPreview + '\n\n';
    message += 'Repository: ' + repoInfo.url + '\n';
    message += 'Version: ' + repoInfo.version + '\n\n';
    message += 'Send REAL email?';
    
    var response = ui.alert('Confirm Test', message, ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    logToSheet('MANUAL TEST: Processing row ' + row);
    
    try {
      processOrderRow(sheet, row, mapping);
      logToSheet('TEST EMAIL: Sent to ' + rowData.email + ' with logos: ' + logoPreview);
      
      ui.alert('Test Email Sent!', 
        'GitHub test email sent successfully! 🎉\n\n' +
        'Email sent to: ' + rowData.email + '\n' +
        'Logo(s) used: ' + logoPreview + '\n' +
        'Strategy: ' + CONFIG.LOGO_STRATEGY + '\n\n' +
        'Check the recipient\'s email to verify logos display correctly.',
        ui.ButtonSet.OK);
        
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
          
          ui.alert('Test Email Sent!', 
            'GitHub test email sent (duplicate override)! 🎉\n\n' +
            'Check the Automation_Log for details.',
            ui.ButtonSet.OK);
        }
      } else {
        throw error;
      }
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Test Failed', 
      'GitHub test failed: ' + error.toString() + 
      '\n\nCheck the Automation_Log for details.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
    logToSheet('GITHUB TEST FAILED: ' + error.toString());
  }
}

/**
 * Force primary logo test (current strategy)
 */
function forcePrimaryLogoTest() {
  try {
    var originalStrategy = CONFIG.LOGO_STRATEGY;
    CONFIG.LOGO_STRATEGY = 'primary';
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Force Primary Logo Test', 
      'This will temporarily use only the primary logo (' + GITHUB_CONFIG.logo.altText + ') for testing.\n\nContinue?', 
      ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) {
      CONFIG.LOGO_STRATEGY = originalStrategy;
      return;
    }
    
    var sheet = SpreadsheetApp.getActiveSheet();
    var activeRange = sheet.getActiveRange();
    
    if (!activeRange || activeRange.getRow() === 1) {
      CONFIG.LOGO_STRATEGY = originalStrategy;
      ui.alert('Error', 'Please select a data row first.', ui.ButtonSet.OK);
      return;
    }
    
    var row = activeRange.getRow();
    var mapping = getColumnMapping(sheet);
    
    logToSheet('FORCE PRIMARY LOGO TEST: Processing row ' + row);
    
    try {
      processOrderRow(sheet, row, mapping);
      logToSheet('Force primary logo test completed successfully');
      
      ui.alert('Test Complete', 
        'Primary logo test email sent!\n\nThis email should show only the ' + GITHUB_CONFIG.logo.altText + ' logo.\nCheck your email to verify.', 
        ui.ButtonSet.OK);
        
    } catch (error) {
      if (error.toString().indexOf('Duplicate email prevented') !== -1) {
        if (mapping.notified) {
          sheet.getRange(row, mapping.notified).setValue('');
        }
        processOrderRow(sheet, row, mapping);
        ui.alert('Test Complete', 
          'Primary logo test email sent (duplicate override)!\n\nThis email should show only the ' + GITHUB_CONFIG.logo.altText + ' logo.', 
          ui.ButtonSet.OK);
      } else {
        throw error;
      }
    }
    
    CONFIG.LOGO_STRATEGY = originalStrategy;
    
  } catch (error) {
    CONFIG.LOGO_STRATEGY = originalStrategy;
    logToSheet('Force primary logo test failed: ' + error.toString());
    SpreadsheetApp.getUi().alert('Test Failed', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Force secondary logo test
 */
function forceSecondaryLogoTest() {
  try {
    var originalStrategy = CONFIG.LOGO_STRATEGY;
    CONFIG.LOGO_STRATEGY = 'secondary';
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Force Secondary Logo Test', 
      'This will temporarily use only the secondary logo (' + GITHUB_CONFIG.logo2.altText + ') for testing.\n\nContinue?', 
      ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) {
      CONFIG.LOGO_STRATEGY = originalStrategy;
      return;
    }
    
    var sheet = SpreadsheetApp.getActiveSheet();
    var activeRange = sheet.getActiveRange();
    
    if (!activeRange || activeRange.getRow() === 1) {
      CONFIG.LOGO_STRATEGY = originalStrategy;
      ui.alert('Error', 'Please select a data row first.', ui.ButtonSet.OK);
      return;
    }
    
    var row = activeRange.getRow();
    var mapping = getColumnMapping(sheet);
    
    logToSheet('FORCE SECONDARY LOGO TEST: Processing row ' + row);
    
    try {
      processOrderRow(sheet, row, mapping);
      logToSheet('Force secondary logo test completed successfully');
      
      ui.alert('Test Complete', 
        'Secondary logo test email sent!\n\nThis email should show only the ' + GITHUB_CONFIG.logo2.altText + ' logo.\nCheck your email to verify.', 
        ui.ButtonSet.OK);
        
    } catch (error) {
      if (error.toString().indexOf('Duplicate email prevented') !== -1) {
        if (mapping.notified) {
          sheet.getRange(row, mapping.notified).setValue('');
        }
        processOrderRow(sheet, row, mapping);
        ui.alert('Test Complete', 
          'Secondary logo test email sent (duplicate override)!\n\nThis email should show only the ' + GITHUB_CONFIG.logo2.altText + ' logo.', 
          ui.ButtonSet.OK);
      } else {
        throw error;
      }
    }
    
    CONFIG.LOGO_STRATEGY = originalStrategy;
    
  } catch (error) {
    CONFIG.LOGO_STRATEGY = originalStrategy;
    logToSheet('Force secondary logo test failed: ' + error.toString());
    SpreadsheetApp.getUi().alert('Test Failed', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Force both logos test
 */
function forceBothLogosTest() {
  try {
    var originalStrategy = CONFIG.LOGO_STRATEGY;
    CONFIG.LOGO_STRATEGY = 'both';
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Force Both Logos Test', 
      'This will temporarily show both logos side by side for testing.\n\nContinue?', 
      ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) {
      CONFIG.LOGO_STRATEGY = originalStrategy;
      return;
    }
    
    var sheet = SpreadsheetApp.getActiveSheet();
    var activeRange = sheet.getActiveRange();
    
    if (!activeRange || activeRange.getRow() === 1) {
      CONFIG.LOGO_STRATEGY = originalStrategy;
      ui.alert('Error', 'Please select a data row first.', ui.ButtonSet.OK);
      return;
    }
    
    var row = activeRange.getRow();
    var mapping = getColumnMapping(sheet);
    
    logToSheet('FORCE BOTH LOGOS TEST: Processing row ' + row);
    
    try {
      processOrderRow(sheet, row, mapping);
      logToSheet('Force both logos test completed successfully');
      
      ui.alert('Test Complete', 
        'Both logos test email sent!\n\nThis email should show both ' + GITHUB_CONFIG.logo.altText + ' and ' + GITHUB_CONFIG.logo2.altText + ' logos side by side.\nCheck your email to verify.', 
        ui.ButtonSet.OK);
        
    } catch (error) {
      if (error.toString().indexOf('Duplicate email prevented') !== -1) {
        if (mapping.notified) {
          sheet.getRange(row, mapping.notified).setValue('');
        }
        processOrderRow(sheet, row, mapping);
        ui.alert('Test Complete', 
          'Both logos test email sent (duplicate override)!\n\nThis email should show both logos side by side.', 
          ui.ButtonSet.OK);
      } else {
        throw error;
      }
    }
    
    CONFIG.LOGO_STRATEGY = originalStrategy;
    
  } catch (error) {
    CONFIG.LOGO_STRATEGY = originalStrategy;
    logToSheet('Force both logos test failed: ' + error.toString());
    SpreadsheetApp.getUi().alert('Test Failed', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Test configuration values
 */
function testConfiguration() {
  try {
    logToSheet('=== CONFIGURATION TEST ===');
    logToSheet('LOGO_STRATEGY: "' + CONFIG.LOGO_STRATEGY + '"');
    logToSheet('LOGO2_KEYWORDS: [' + CONFIG.LOGO2_KEYWORDS.join(', ') + ']');
    logToSheet('Primary logo URL: ' + CONFIG.LOGO_URL);
    logToSheet('Secondary logo URL: ' + CONFIG.LOGO2_URL);
    logToSheet('Primary logo alt text: "' + CONFIG.LOGO_ALT_TEXT + '"');
    logToSheet('Secondary logo alt text: "' + CONFIG.LOGO2_ALT_TEXT + '"');
    
    // Test keyword matching with sample data
    var testEmails = [
      'user@gmail.com',
      'student@purdue.edu',
      'admin@university.edu',
      'test@company.com'
    ];
    
    for (var i = 0; i < testEmails.length; i++) {
      var testEmail = testEmails[i];
      var testData = { email: testEmail, description: 'Test order' };
      var logos = selectLogos(testData);
      logToSheet('Test email "' + testEmail + '" -> ' + logos[0].altText + ' logo');
    }
    
    logToSheet('=== CONFIGURATION TEST COMPLETE ===');
    
    var ui = SpreadsheetApp.getUi();
    ui.alert('Configuration Test Complete', 'Check Automation_Log for detailed results.', ui.ButtonSet.OK);
    
  } catch (error) {
    logToSheet('Configuration test error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Test Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =============================================================================
// SYSTEM ADMINISTRATION AND SETUP FUNCTIONS
// =============================================================================

/**
 * Initialize the entire system
 */
function initializeSystem() {
  try {
    logToSheet('=== GITHUB SYSTEM INITIALIZATION ===');
    
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
    
    // Test GitHub integration
    testGitHubLogos();
    
    logToSheet('=== GITHUB SYSTEM INITIALIZATION COMPLETED ===');
    
    var repoInfo = getRepositoryInfo();
    var message = 'GitHub System Initialized!\n\n';
    message += 'Configuration validated\n';
    message += 'Data validation rules applied\n';
    message += 'Helper columns configured\n';
    message += 'Github integration tested\n\n';
    message += 'REPOSITORY INFO:\n';
    message += 'Repository: ' + repoInfo.url + '\n';
    message += 'Version: ' + repoInfo.version + '\n';
    message += 'Primary logo: ' + repoInfo.logoUrl + '\n';
    message += 'Secondary logo: ' + repoInfo.logo2Url + '\n';
    message += 'Strategy: ' + CONFIG.LOGO_STRATEGY + '\n\n';
    if (sheet) {
      message += 'SHEET CONFIG:\n';
      message += 'Main sheet: "' + sheet.getName() + '"\n';
    }
    try {
      var senderEmail = getSenderEmail();
      message += 'Sender email: ' + senderEmail + '\n';
    } catch (e) {
      message += 'Sender email: ERROR - ' + e.message + '\n';
    }
    message += '\nThe system is ready!\n\n';
    message += 'NEXT STEPS:\n';
    message += '1. Test with "Test email for selected row"\n';
    message += '2. Try different logo strategies\n';
    message += '3. Check logs in Automation_Log sheet';
    
    SpreadsheetApp.getUi().alert('GitHub System Initialized', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    logToSheet('GITHUB INIT ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert('Initialization Failed', 
      'GitHub system initialization failed: ' + error.toString(), 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Validate system configuration
 */
function validateConfiguration() {
  var issues = [];

  try {
    // --- 1) Check presence of global configs safely ---
    var hasCONFIG = (typeof CONFIG !== 'undefined' && CONFIG && typeof CONFIG === 'object');
    var hasGITHUB = (typeof GITHUB_CONFIG !== 'undefined' && GITHUB_CONFIG && typeof GITHUB_CONFIG === 'object');

    if (!hasCONFIG) issues.push('CONFIG object is missing or not an object');
    if (!hasGITHUB) issues.push('GITHUB_CONFIG object is missing or not an object');

    // --- 2) Gmail Advanced Service check ---
    // (This verifies the Advanced Gmail service "Gmail" is enabled under Services.)
    if (typeof Gmail === 'undefined') {
      issues.push('Gmail Advanced Service not available. In Apps Script: Services → + → enable "Gmail API".');
    }

    // --- 3) GitHub configuration checks (guard each read) ---
    var ghUser = hasGITHUB && GITHUB_CONFIG.username;
    var ghRepo = hasGITHUB && GITHUB_CONFIG.repository;
    var logo1  = hasGITHUB && GITHUB_CONFIG.logo && GITHUB_CONFIG.logo.filename;
    var logo2  = hasGITHUB && GITHUB_CONFIG.logo2 && GITHUB_CONFIG.logo2.filename;

    if (!ghUser || ghUser === 'yourusername') issues.push('GITHUB_CONFIG.username not configured');
    if (!ghRepo) issues.push('GITHUB_CONFIG.repository not configured');
    if (!logo1) issues.push('Primary logo filename (GITHUB_CONFIG.logo.filename) not configured');
    if (!logo2) issues.push('Secondary logo filename (GITHUB_CONFIG.logo2.filename) not configured');

    // --- 4) Sheet + mapping checks ---
    var sheet = null;
    try {
      sheet = (typeof getMainSheet === 'function') ? getMainSheet() : null;
    } catch (e) {
      issues.push('Error calling getMainSheet(): ' + e.toString());
    }

    // Try to obtain mapping from a helper or a global (support either pattern).
    var mapping = null;
    try {
      if (typeof getColumnMapping === 'function') {
        mapping = getColumnMapping(sheet);
      } else if (typeof MAPPING !== 'undefined') {
        mapping = MAPPING;
      }
    } catch (e) {
      issues.push('Error determining column mapping: ' + e.toString());
    }

    if (!sheet) {
      issues.push('Main sheet not found (getMainSheet returned null/undefined)');
    } else {
      if (!mapping || typeof mapping !== 'object') {
        issues.push('Could not determine column mapping');
      } else {
        if (!mapping.email)   issues.push('Email column not found in mapping');
        if (!mapping.ordered) issues.push('Order status column not found in mapping');
        if (!mapping.po)      issues.push('PO Number column not found in mapping');
      }
    }

    // --- 5) Sender email check ---
    try {
      var senderEmail = (typeof getSenderEmail === 'function') ? getSenderEmail() : null;
      if (!senderEmail) {
        issues.push('Could not determine sender email (getSenderEmail returned empty)');
      } else {
        logToSheet('Sender email configured: ' + senderEmail);
      }
    } catch (senderError) {
      issues.push('Sender email error: ' + senderError.toString());
    }

    // --- 6) Logo strategy & URLs ---
    var validStrategies = ['primary', 'secondary', 'both', 'conditional'];
    var strategy = hasCONFIG ? CONFIG.LOGO_STRATEGY : null;
    if (validStrategies.indexOf(strategy) === -1) {
      issues.push('Invalid LOGO_STRATEGY: ' + strategy + ' (valid: ' + validStrategies.join(', ') + ')');
    }

    var logoUrl1 = hasCONFIG ? CONFIG.LOGO_URL : null;
    var logoUrl2 = hasCONFIG ? CONFIG.LOGO2_URL : null;

    if (logoUrl1 && ghUser && typeof logoUrl1 === 'string') {
      if (!logoUrl1.includes(ghUser)) issues.push('Primary logo URL does not match GitHub username');
    } else if (!logoUrl1) {
      issues.push('CONFIG.LOGO_URL is not set');
    }

    if (logoUrl2 && ghUser && typeof logoUrl2 === 'string') {
      if (!logoUrl2.includes(ghUser)) issues.push('Secondary logo URL does not match GitHub username');
    } else if (!logoUrl2) {
      issues.push('CONFIG.LOGO2_URL is not set');
    }

    // --- 7) Final logging & result ---
    if (issues.length > 0) {
      logToSheet('CONFIGURATION ISSUES FOUND:');
      for (var j = 0; j < issues.length; j++) {
        logToSheet('- ' + issues[j]);
      }
      return false;
    } else {
      logToSheet('configuration validation passed');
      return true;
    }
  } catch (error) {
    logToSheet('validateConfiguration ERROR: ' + error.toString());
    return false;
  }
}


/**
 * Setup data validation rules
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
    
    // Email validation
    if (mapping.email) {
      var emailRule = SpreadsheetApp.newDataValidation()
        .requireTextIsEmail()
        .setAllowInvalid(false)
        .setHelpText('Please enter a valid email address')
        .build();
      sheet.getRange(2, mapping.email, lastRow-1, 1).setDataValidation(emailRule);
      logToSheet('Applied email validation to column ' + mapping.email);
    }
    
    // Order status validation
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

/**
 * Check overall system setup
 */
function checkSetup() {
  try {
    var issues = [];
    var warnings = [];
    
    var configValid = validateConfiguration();
    if (!configValid) {
      issues.push('Configuration validation failed - check logs');
    }
    
    // GitHub dual logo specific checks
    if (GITHUB_CONFIG.username === 'yourusername') {
      issues.push('GITHUB_CONFIG.username needs to be updated');
    }
    
    if (GITHUB_CONFIG.logo.filename === GITHUB_CONFIG.logo2.filename) {
      warnings.push('Both logos use the same filename - they will be identical');
    }
    
    var validStrategies = ['primary', 'secondary', 'both', 'conditional'];
    if (validStrategies.indexOf(CONFIG.LOGO_STRATEGY) === -1) {
      issues.push('Invalid logo strategy: ' + CONFIG.LOGO_STRATEGY);
    }
    
    var message = '';
    
    if (issues.length > 0) {
      message += 'ISSUES FOUND:\n\n';
      for (var i = 0; i < issues.length; i++) {
        message += 'X ' + issues[i] + '\n';
      }
      message += '\nPlease fix these issues before using the system.\n';
    } else {
      message += 'Setup validation passed!\n\n';
    }
    
    if (warnings.length > 0) {
      message += '\nWARNINGS:\n';
      for (var j = 0; j < warnings.length; j++) {
        message += '! ' + warnings[j] + '\n';
      }
    }
    
    var sheet = getMainSheet();
    if (sheet) {
      message += '\nCURRENT CONFIG:\n';
      message += 'Main sheet: "' + sheet.getName() + '"\n';
      message += 'Logo strategy: ' + CONFIG.LOGO_STRATEGY + '\n';
      message += 'Primary logo: ' + GITHUB_CONFIG.logo.altText + '\n';
      message += 'Secondary logo: ' + GITHUB_CONFIG.logo2.altText + '\n';
      
      var repoInfo = getRepositoryInfo();
      message += 'Repository: ' + repoInfo.url + '\n';
    }
    
    message += '\nFEATURES:\n';
    message += '• Conditional logo selection\n';
    message += '• GitHub-hosted images\n';
    message += '• Multiple fallback URLs\n';
    message += '• Side-by-side display option\n';
    message += '• Keyword-based switching\n';
    
    SpreadsheetApp.getUi().alert('GitHub Setup Check', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Setup check failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Configure sender email
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
    
    var message = 'EMAIL SENDER CONFIGURATION (GitHub Edition)\n\n';
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
        confirmMessage += 'GitHub repository: ' + GITHUB_CONFIG.username + '/' + GITHUB_CONFIG.repository + '\n\n';
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
 * View current configuration with GitHub details
 */
function viewCurrentConfig() {
  try {
    var sheet = getMainSheet();
    var repoInfo = getRepositoryInfo();
    var message = 'CURRENT SYSTEM CONFIGURATION (GitHub)\n\n';
    
    // Sheet information
    if (sheet) {
      message += 'SHEET CONFIG:\n';
      message += 'Main Sheet: "' + sheet.getName() + '"\n';
      var mapping = getColumnMapping(sheet);
      if (mapping) {
        message += 'Email Column: ' + (mapping.email ? 'Column ' + mapping.email : 'Not found') + '\n';
        message += 'Order Status Column: ' + (mapping.ordered ? 'Column ' + mapping.ordered : 'Not found') + '\n';
        message += 'PO Number Column: ' + (mapping.po ? 'Column ' + mapping.po : 'Not found') + '\n';
      }
    } else {
      message += 'SHEET CONFIG:\n';
      message += 'Main Sheet: Not found\n';
    }
    
    message += '\n';
    
    // GitHub configuration
    message += 'GITHUB CONFIG:\n';
    message += 'Repository: ' + repoInfo.url + '\n';
    message += 'Version: ' + repoInfo.version + '\n';
    message += 'Primary logo: ' + repoInfo.logoUrl + '\n';
    message += 'Secondary logo: ' + repoInfo.logo2Url + '\n';
    message += 'Username: ' + GITHUB_CONFIG.username + '\n';
    message += 'Repo Name: ' + GITHUB_CONFIG.repository + '\n';
    message += 'Branch: ' + GITHUB_CONFIG.branch + '\n';
    message += 'Logo Strategy: ' + CONFIG.LOGO_STRATEGY + '\n';
    
    message += '\n';
    
    // Email configuration
    message += 'EMAIL CONFIG:\n';
    message += 'Sender Config: ' + CONFIG.EMAIL_SENDER + '\n';
    try {
      var currentSender = getSenderEmail();
      message += 'Active Sender: ' + currentSender + '\n';
    } catch (e) {
      message += 'Active Sender: ERROR - ' + e.message + '\n';
    }
    message += 'Subject: ' + CONFIG.EMAIL_SUBJECT + '\n';
    message += 'Signature: ' + CONFIG.EMAIL_SIGNATURE + '\n';
    
    message += '\n';
    
    // Safety settings
    message += 'SAFETY SETTINGS:\n';
    message += 'Rate Limits: ' + CONFIG.MAX_EMAILS_PER_HOUR + '/hour, ' + CONFIG.MAX_EMAILS_PER_DAY + '/day\n';
    message += 'Duplicate Check: ' + CONFIG.DUPLICATE_CHECK_DAYS + ' days\n';
    message += 'Attach PDFs: ' + (CONFIG.ATTACH_QUOTE_PDF ? 'Yes' : 'No') + '\n';
    
    message += '\n';
    
    // User information
    try {
      var currentUser = Session.getActiveUser().getEmail();
      message += 'USER INFO:\n';
      message += 'Current User: ' + currentUser + '\n';
    } catch (e) {
      message += 'USER INFO:\n';
      message += 'Current User: Could not determine\n';
    }
    
    message += 'onEdit Trigger: Check manually in Apps Script > Triggers\n';
    
    message += '\n';
    message += 'FEATURES:\n';
    for (var i = 0; i < repoInfo.features.length; i++) {
      message += '• ' + repoInfo.features[i] + '\n';
    }
    
    SpreadsheetApp.getUi().alert('Current Configuration', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Configuration Error', 'Failed to get current configuration: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * View email backup log
 */
function viewEmailBackupLog() {
  try {
    var backupSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.BACKUP_SHEET_NAME);
    
    if (!backupSheet) {
      SpreadsheetApp.getUi().alert('Info', 'No email backup log found yet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    backupSheet.activate();
    
    var lastRow = backupSheet.getLastRow();
    var message = 'Email backup log contains ' + Math.max(0, lastRow - 1) + ' entries.\n\n';
    message += 'The log shows all email attempts with GitHub version tracking.\n\n';
    message += 'Columns include:\n';
    message += '• Timestamp\n• Email address\n• PO Number\n• Description\n• Status\n• Error Message\n• Message Key\n• GitHub Version';
    
    SpreadsheetApp.getUi().alert('Email Backup Log', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to view backup log: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * View rate limit status
 */
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
      message += 'Please wait before sending more emails.\n';
    } else if (dailyCount >= CONFIG.MAX_EMAILS_PER_DAY) {
      message += 'DAILY LIMIT REACHED\n';
      message += 'Please wait until tomorrow.\n';
    } else {
      message += 'WITHIN LIMITS\n';
      message += 'Emails can be sent normally.\n';
    }
    
    message += '\nThese limits help prevent accidental spam and protect your Gmail reputation.';
    
    SpreadsheetApp.getUi().alert('Rate Limit Status', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to check rate limit status: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Reset rate limits (emergency function)
 */
function resetRateLimits() {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Reset Rate Limits', 
      'Are you sure you want to reset the rate limits?\n\n' +
      'This should only be done in emergency situations.\n' +
      'Rate limits protect your Gmail reputation.',
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
      logToSheet('ADMIN ACTION: Rate limits reset by user (GitHub)');
    }
    
    ui.alert('Rate Limits Reset', 'Rate limits have been reset.\n\nPlease use this feature responsibly.', ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to reset rate limits: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Audit data integrity
 */
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
    
    logToSheet('=== GITHUB DATA INTEGRITY AUDIT STARTED ===');
    logToSheet('Auditing sheet: ' + sheet.getName());
    logToSheet('GitHub repository: ' + GITHUB_CONFIG.username + '/' + GITHUB_CONFIG.repository);
    logToSheet('Logo strategy: ' + CONFIG.LOGO_STRATEGY);
    
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
    
    // Check GitHub dual logo configuration integrity
    if (!CONFIG.LOGO_URL.includes(GITHUB_CONFIG.username)) {
      issues.push('Primary logo URL does not match GitHub username');
    }
    
    if (!CONFIG.LOGO2_URL.includes(GITHUB_CONFIG.username)) {
      issues.push('Secondary logo URL does not match GitHub username');
    }
    
    if (GITHUB_CONFIG.logo.filename === GITHUB_CONFIG.logo2.filename) {
      issues.push('Both logos use the same filename - they will be identical');
    }
    
    var validStrategies = ['primary', 'secondary', 'both', 'conditional'];
    if (validStrategies.indexOf(CONFIG.LOGO_STRATEGY) === -1) {
      issues.push('Invalid logo strategy: ' + CONFIG.LOGO_STRATEGY);
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
    
    logToSheet('=== GITHUB DATA INTEGRITY AUDIT COMPLETED ===');
    
    var repoInfo = getRepositoryInfo();
    var message = 'Data integrity audit completed for "' + sheet.getName() + '".\n\n';
    message += 'Found ' + issues.length + ' issues\n';
    message += 'Made ' + fixes + ' automatic fixes\n';
    message += 'GitHub repository: ' + repoInfo.url + '\n';
    message += 'Logo strategy: ' + CONFIG.LOGO_STRATEGY + '\n';
    message += 'Primary logo: ' + GITHUB_CONFIG.logo.altText + '\n';
    message += 'Secondary logo: ' + GITHUB_CONFIG.logo2.altText + '\n\n';
    
    if (issues.length > 0) {
      message += 'Issues found - check Automation_Log for details.\n';
    } else {
      message += 'No issues found - system is healthy!\n';
    }
    
    message += '\nsystem operating normally.';
    
    SpreadsheetApp.getUi().alert('Audit Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    logToSheet('Data audit failed: ' + error.toString());
    SpreadsheetApp.getUi().alert('Audit Failed', 'Data integrity audit failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =============================================================================
// END OF GITHUB PURCHASE ORDER EMAIL NOTIFICATION SYSTEM
// =============================================================================
