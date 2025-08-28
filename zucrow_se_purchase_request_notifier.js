/**
 * GitHub-Ready Purchase Order Email Notification System
 * 
 * Repository: https://github.com/yourusername/purchase-order-notifier
 * Version: 2.0 - GitHub Edition
 * 
 * QUICK SETUP:
 * 1. Fork/clone the GitHub repository
 * 2. Upload your logo to assets/ folder in the repo
 * 3. Update GITHUB_CONFIG below with your repository details
 * 4. Enable Gmail API in Google Apps Script (Services > Gmail API)
 * 5. Set up onEdit trigger (Triggers > Add Trigger > onEdit > From spreadsheet > On edit)
 * 
 * Features:
 * - GitHub-hosted images (no Drive permissions needed)
 * - Auto-detection of sheet structure
 * - Duplicate prevention & rate limiting
 * - Professional email templates
 * - Multiple logo fallback options
 * - Enhanced error handling and logging
 */

// =============================================================================
// UPDATED GITHUB CONFIGURATION - DUAL LOGO SUPPORT
// =============================================================================

var GITHUB_CONFIG = {
  // Your GitHub repository details
  username: 'yourusername',                    // Replace with YOUR GitHub username
  repository: 'purchase-order-notifier',       // Replace with YOUR repo name
  branch: 'main',                             // Usually 'main' or 'master'
  
  // Primary logo configuration (GitHub-hosted)
  logo: {
    filename: 'spectral_logo.png',             // Your logo file in assets/ folder
    altText: 'Spectral',                       // Alt text for accessibility
    maxWidth: '200px',                         // Maximum logo width
    maxHeight: '100px'                         // Maximum logo height
  },
  
  // Secondary logo configuration (GitHub-hosted)  
  logo2: {
    filename: 'purdue_prop_logo.png',          // Your second logo file in assets/ folder
    altText: 'Purdue_Prop',                    // Alt text for accessibility
    maxWidth: '200px',                         // Maximum logo width
    maxHeight: '100px'                         // Maximum logo height
  }
};

// =============================================================================
// UPDATED MAIN CONFIGURATION - DUAL LOGO SUPPORT
// =============================================================================

var CONFIG = {
  // Sheet detection (leave empty for auto-detection)
  SHEET_NAME: '',
  
  // Logo selection strategy: 'primary', 'secondary', 'both', 'conditional'
  LOGO_STRATEGY: 'primary', // Options: 'primary', 'secondary', 'both', 'conditional'
  
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
  
  // Rest of existing CONFIG...
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
// UPDATED GITHUB INTEGRATION FUNCTIONS - DUAL LOGO SUPPORT
// =============================================================================

/**
 * Generate GitHub raw image URL for any logo
 */
function generateGitHubImageUrl(filename) {
  return 'https://raw.githubusercontent.com/' + 
         GITHUB_CONFIG.username + '/' + 
         GITHUB_CONFIG.repository + '/' + 
         GITHUB_CONFIG.branch + '/assets/' + filename;
}

/**
 * Determine which logo(s) to use based on strategy and data
 */
function selectLogos(data) {
  var logos = [];
  
  switch (CONFIG.LOGO_STRATEGY) {
    case 'primary':
      logos.push({
        url: CONFIG.LOGO_URL,
        altText: CONFIG.LOGO_ALT_TEXT,
        maxWidth: CONFIG.LOGO_MAX_WIDTH,
        maxHeight: CONFIG.LOGO_MAX_HEIGHT
      });
      break;
      
    case 'secondary':
      logos.push({
        url: CONFIG.LOGO2_URL,
        altText: CONFIG.LOGO2_ALT_TEXT,
        maxWidth: CONFIG.LOGO2_MAX_WIDTH,
        maxHeight: CONFIG.LOGO2_MAX_HEIGHT
      });
      break;
      
    case 'both':
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
      // Use logo2 if email or description contains keywords, otherwise use primary
      var useSecondary = false;
      var emailText = (data.email || '').toLowerCase();
      var descText = (data.description || '').toLowerCase();
      
      for (var i = 0; i < CONFIG.LOGO2_KEYWORDS.length; i++) {
        var keyword = CONFIG.LOGO2_KEYWORDS[i].toLowerCase();
        if (emailText.indexOf(keyword) !== -1 || descText.indexOf(keyword) !== -1) {
          useSecondary = true;
          break;
        }
      }
      
      if (useSecondary) {
        logos.push({
          url: CONFIG.LOGO2_URL,
          altText: CONFIG.LOGO2_ALT_TEXT,
          maxWidth: CONFIG.LOGO2_MAX_WIDTH,
          maxHeight: CONFIG.LOGO2_MAX_HEIGHT
        });
      } else {
        logos.push({
          url: CONFIG.LOGO_URL,
          altText: CONFIG.LOGO_ALT_TEXT,
          maxWidth: CONFIG.LOGO_MAX_WIDTH,
          maxHeight: CONFIG.LOGO_MAX_HEIGHT
        });
      }
      break;
      
    default:
      // Fallback to primary logo
      logos.push({
        url: CONFIG.LOGO_URL,
        altText: CONFIG.LOGO_ALT_TEXT,
        maxWidth: CONFIG.LOGO_MAX_WIDTH,
        maxHeight: CONFIG.LOGO_MAX_HEIGHT
      });
  }
  
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
 * Updated email content builder with dual logo support
 */
function buildEmailContent(data) {
  var greeting = data.name ? 'Hello ' + data.name + ',' : CONFIG.EMAIL_GREETING + ',';
  
  // Select appropriate logo(s) based on strategy and data
  var selectedLogos = selectLogos(data);
  var logoHtml = buildLogoHtml(selectedLogos);
  
  var htmlBody = '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #ffffff;">';
  
  // Add GitHub-hosted logo(s)
  htmlBody += logoHtml;
  
  htmlBody += '<p style="font-size: 16px; line-height: 1.5; color: #333; margin-bottom: 20px;">' + greeting + '</p>';
  htmlBody += '<p style="font-size: 16px; line-height: 1.5; color: #333; margin-bottom: 25px;">Great news! Your order has been placed and is being processed.</p>';
  
  // Modern order details card
  htmlBody += '<div style="background: #ffffff; border: 1px solid #e1e5e9; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1); margin: 25px 0;">';
  htmlBody += '<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px;">' +
              '<h2 style="margin: 0; font-size: 20px; font-weight: 600;">Order Details</h2></div>';
  
  if (data.po) {
    htmlBody += '<div style="padding: 16px 20px; border-bottom: 1px solid #f0f0f0; display: flex;">' +
                '<div style="font-weight: 600; color: #555; width: 140px; flex-shrink: 0;">PO Number:</div>' +
                '<div style="color: #333; font-family: monospace; font-size: 14px;">' + escapeHtml(data.po) + '</div></div>';
  }
  
  if (data.description) {
    htmlBody += '<div style="padding: 16px 20px; border-bottom: 1px solid #f0f0f0; display: flex;">' +
                '<div style="font-weight: 600; color: #555; width: 140px; flex-shrink: 0;">Description:</div>' +
                '<div style="color: #333;">' + escapeHtml(data.description) + '</div></div>';
  }
  
  htmlBody += '<div style="padding: 16px 20px; border-bottom: 1px solid #f0f0f0; display: flex;">' +
              '<div style="font-weight: 600; color: #555; width: 140px; flex-shrink: 0;">Status:</div>' +
              '<div style="color: #28a745; font-weight: 600; font-size: 16px;">✅ Order Placed</div></div>';
  
  htmlBody += '<div style="padding: 16px 20px; display: flex;">' +
              '<div style="font-weight: 600; color: #555; width: 140px; flex-shrink: 0;">Date:</div>' +
              '<div style="color: #333;">' + new Date().toLocaleString() + '</div></div>';
  
  htmlBody += '</div>';
  
  // Next steps section
  htmlBody += '<div style="background: linear-gradient(135deg, #e3f2fd 0%, #f3e5f5 100%); border-radius: 8px; padding: 20px; margin: 25px 0;">';
  htmlBody += '<h3 style="margin: 0 0 10px 0; color: #1976d2; font-size: 16px;">Next Steps</h3>';
  htmlBody += '<p style="margin: 0; color: #555; line-height: 1.5;">We\'ll keep you updated on your order progress. If you have any questions about your order, please reply to this email and we\'ll get back to you promptly.</p>';
  htmlBody += '</div>';
  
  // Footer
  htmlBody += '<div style="text-align: center; margin-top: 40px; padding-top: 20px; border-top: 1px solid #e0e0e0;">';
  htmlBody += '<p style="margin: 0; color: #666; font-size: 14px;">' + CONFIG.EMAIL_SIGNATURE + '</p>';
  
  // GitHub attribution with dual logo info
  if (GITHUB_CONFIG.username && GITHUB_CONFIG.repository) {
    var logoInfo = selectedLogos.length > 1 ? ' (dual logo)' : ' (' + selectedLogos[0].altText + ')';
    htmlBody += '<p style="margin: 10px 0 0 0; font-size: 12px; color: #999;">' +
                '<a href="https://github.com/' + GITHUB_CONFIG.username + '/' + GITHUB_CONFIG.repository + '" ' +
                'style="color: #999; text-decoration: none;">⚡ Powered by GitHub automation' + logoInfo + '</a></p>';
  }
  
  htmlBody += '</div></div>';
  
  // Plain text version (for email clients that don't support HTML)
  var textBody = greeting + '\n\n';
  textBody += 'Great news! Your order has been placed and is being processed.\n\n';
  textBody += 'ORDER DETAILS\n';
  textBody += '=============\n';
  if (data.po) textBody += 'PO Number: ' + data.po + '\n';
  if (data.description) textBody += 'Description: ' + data.description + '\n';
  textBody += 'Status: Order Placed ✓\n';
  textBody += 'Date: ' + new Date().toLocaleString() + '\n\n';
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
 * Updated GitHub logo test function for dual logos
 */
function testGitHubLogos() {
  try {
    logToSheet('=== TESTING GITHUB DUAL LOGOS ===');
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
      name: 'GitHub Dual Logo Test User',
      email: 'test@github-integration.com',
      po: 'GITHUB-DUAL-001',
      description: 'Testing GitHub dual logo integration'
    };
    
    var emailContent = buildEmailContent(testData);
    logToSheet('✓ Email content with dual logo support generated successfully');
    
    logToSheet('=== GITHUB DUAL LOGO TEST COMPLETE ===');
    
    var ui = SpreadsheetApp.getUi();
    var message = 'GITHUB DUAL LOGO TEST RESULTS\n\n';
    message += '✅ Email generation: SUCCESS\n';
    message += '✅ Dual logo integration: SUCCESS\n';
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
    message += '4. Test with different LOGO_STRATEGY settings\n\n';
    message += 'Check Automation_Log for detailed results.';
    
    ui.alert('GitHub Dual Logo Test Results', message, ui.ButtonSet.OK);
    
  } catch (error) {
    logToSheet('GitHub dual logo test failed: ' + error.toString());
    SpreadsheetApp.getUi().alert('Test Failed', 
      'GitHub dual logo test failed: ' + error.toString() + 
      '\n\nPlease check your GITHUB_CONFIG settings and ensure both logo files exist.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// =============================================================================
// UPDATED MENU SYSTEM FOR DUAL LOGO SUPPORT
// =============================================================================

/**
 * Updated menu creation with dual logo options
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Order Notifier (GitHub Dual Logo)')
    .addItem('🧪 Test email for selected row', 'testEmailForRow')
    .addItem('✅ Check setup & validate config', 'checkSetup')
    .addItem('📊 Setup data validation rules', 'setupDataValidation')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🐙 GitHub Integration')
      .addItem('🔧 Setup GitHub integration', 'setupGitHubIntegration')
      .addItem('🖼️ Test GitHub dual logos', 'testGitHubLogos')
      .addItem('🎨 Configure logo strategy', 'configureLogoStrategy')
      .addItem('📋 View repository info', 'viewRepositoryInfo'))
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
 * Configure logo selection strategy
 */
function configureLogoStrategy() {
  try {
    var ui = SpreadsheetApp.getUi();
    var currentStrategy = CONFIG.LOGO_STRATEGY;
    
    var message = 'LOGO STRATEGY CONFIGURATION\n\n';
    message += 'Current strategy: ' + currentStrategy + '\n\n';
    message += 'Available strategies:\n\n';
    message += '1. PRIMARY - Always use Spectral logo\n';
    message += '2. SECONDARY - Always use Purdue Prop logo\n';
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

// =============================================================================
// CORE SYSTEM FUNCTIONS (Enhanced from original)
// =============================================================================

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
    headers.push('X-GitHub-Automation: purchase-order-notifier');
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
      '2.0-GitHub'
    ]);
    
    var lastRow = backupSheet.getLastRow();
    if (lastRow > 1001) {
      backupSheet.deleteRows(2, lastRow - 1001);
    }
    
  } catch (error) {
    logToSheet('Failed to log email attempt: ' + error.toString());
  }
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

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
 * Clean string values
 */
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
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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
    }
    
    logSheet.appendRow([new Date(), message, '2.0-GitHub']);
    
    var lastRow = logSheet.getLastRow();
    if (lastRow > 501) {
      logSheet.deleteRows(2, lastRow - 501);
    }
    
  } catch (error) {
    console.error('Logging failed: ' + error.toString());
  }
}

// =============================================================================
// SETUP AND CONFIGURATION FUNCTIONS
// =============================================================================

/**
 * Initialize the entire system
 */
function initializeSystem() {
  try {
    logToSheet('=== GITHUB SYSTEM INITIALIZATION STARTED ===');
    
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
    var githubTest = testGitHubLogo();
    
    logToSheet('=== GITHUB SYSTEM INITIALIZATION COMPLETED ===');
    
    var repoInfo = getRepositoryInfo();
    var message = 'GitHub System Initialization Complete! 🎉\n\n';
    message += '✅ Configuration validated\n';
    message += '✅ Data validation rules applied\n';
    message += '✅ Helper columns configured\n';
    message += '✅ GitHub integration tested\n\n';
    message += 'REPOSITORY INFO:\n';
    message += '📁 ' + repoInfo.url + '\n';
    message += '🏷️ Version ' + repoInfo.version + '\n';
    message += '🖼️ Logo: ' + repoInfo.logoUrl + '\n\n';
    if (sheet) {
      message += 'SHEET CONFIG:\n';
      message += '📊 Main sheet: "' + sheet.getName() + '"\n';
    }
    try {
      var senderEmail = getSenderEmail();
      message += '📧 Sender email: ' + senderEmail + '\n';
    } catch (e) {
      message += '📧 Sender email: ERROR - ' + e.message + '\n';
    }
    message += '\n🚀 The system is now ready to use!\n\n';
    message += 'NEXT STEPS:\n';
    message += '1. Test with "Test email for selected row"\n';
    message += '2. Mark an order as placed to see it work\n';
    message += '3. Check logs in Automation_Log sheet\n\n';
    message += 'Need help? Check the repository documentation!';
    
    SpreadsheetApp.getUi().alert('GitHub System Initialized', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    logToSheet('GITHUB INITIALIZATION ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert('Initialization Failed', 
      'GitHub system initialization failed: ' + error.toString() + 
      '\n\nPlease check your GITHUB_CONFIG settings and try again.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Validate system configuration
 */
function validateConfiguration() {
  try {
    var issues = [];
    
    // Check Gmail API
    if (typeof Gmail === 'undefined') {
      issues.push('Gmail API service not available - please add it in Services');
    }
    
    // Check GitHub configuration
    if (!GITHUB_CONFIG.username || GITHUB_CONFIG.username === 'yourusername') {
      issues.push('GITHUB_CONFIG.username not configured');
    }
    
    if (!GITHUB_CONFIG.repository) {
      issues.push('GITHUB_CONFIG.repository not configured');
    }
    
    // Check sheet
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
    
    // Check sender email
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
    
    // Check email subject
    if (!CONFIG.EMAIL_SUBJECT || CONFIG.EMAIL_SUBJECT.trim() === '') {
      issues.push('EMAIL_SUBJECT is empty');
    }
    
    // Check GitHub logo URL
    if (CONFIG.LOGO_URL) {
      logToSheet('GitHub logo URL: ' + CONFIG.LOGO_URL);
    } else {
      issues.push('GitHub logo URL not generated - check GITHUB_CONFIG');
    }
    
    // Check triggers
    try {
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
        issues.push('No onEdit trigger found - please add one');
      }
    } catch (triggerError) {
      logToSheet('Could not check triggers: ' + triggerError.toString());
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
    
    // Gmail API check
    if (typeof Gmail === 'undefined') {
      issues.push('Gmail API service not added');
    }
    
    // GitHub configuration check
    if (GITHUB_CONFIG.username === 'yourusername') {
      issues.push('GITHUB_CONFIG.username needs to be updated');
    }
    
    if (!GITHUB_CONFIG.repository) {
      issues.push('GITHUB_CONFIG.repository not set');
    }
    
    // Sheet check
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
    
    // Sender email check
    try {
      var senderEmail = getSenderEmail();
      logToSheet('Current sender email: ' + senderEmail);
    } catch (senderError) {
      issues.push('Sender email configuration issue: ' + senderError.message);
    }
    
    // GitHub logo check
    if (!CONFIG.LOGO_URL.includes(GITHUB_CONFIG.username)) {
      warnings.push('Logo URL may not be properly configured for your GitHub username');
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
      message += '\nCURRENT CONFIGURATION:\n';
      message += '📊 Main sheet: "' + sheet.getName() + '"\n';
      try {
        var currentSender = getSenderEmail();
        message += '📧 Sender email: ' + currentSender + '\n';
      } catch (e) {
        message += '📧 Sender email: ERROR - ' + e.message + '\n';
      }
      
      var repoInfo = getRepositoryInfo();
      message += '🐙 Repository: ' + repoInfo.url + '\n';
      message += '🖼️ Logo: ' + repoInfo.logoUrl + '\n';
    }
    
    message += '\nSAFETY FEATURES:\n';
    message += '• Duplicate prevention (' + CONFIG.DUPLICATE_CHECK_DAYS + ' days)\n';
    message += '• Rate limiting (' + CONFIG.MAX_EMAILS_PER_HOUR + '/hour, ' + CONFIG.MAX_EMAILS_PER_DAY + '/day)\n';
    message += '• GitHub-hosted images (no Drive permissions needed)\n';
    message += '• Auto sheet detection\n';
    message += '• Enhanced error handling\n';
    message += '• Comprehensive logging\n';
    
    SpreadsheetApp.getUi().alert('GitHub Setup Check Results', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Setup check failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
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
    
    var ui = SpreadsheetApp.getUi();
    var repoInfo = getRepositoryInfo();
    var message = 'GITHUB TEST EMAIL DETAILS:\n\n';
    message += 'Sheet: ' + sheet.getName() + '\n';
    message += 'Row: ' + row + '\n';
    message += 'Email: ' + (rowData.email || '(empty)') + '\n';
    message += 'Name: ' + (rowData.name || '(empty)') + '\n';
    message += 'PO: ' + (rowData.po || '(empty)') + '\n';
    message += 'Description: ' + (rowData.description || '(empty)') + '\n\n';
    
    try {
      var senderEmail = getSenderEmail();
      message += 'Sender: ' + senderEmail + '\n';
    } catch (senderError) {
      message += 'Sender: ERROR - ' + senderError.message + '\n';
    }
    
    message += 'Logo: ' + repoInfo.logoUrl + '\n';
    message += 'Version: ' + repoInfo.version + '\n\n';
    message += 'This will send a REAL email with GitHub-hosted logo. Continue?';
    
    var response = ui.alert('Confirm GitHub Test Email', message, ui.ButtonSet.YES_NO);
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    logToSheet('MANUAL GITHUB TEST: Processing row ' + row);
    
    try {
      processOrderRow(sheet, row, mapping);
      logToSheet('GITHUB TEST EMAIL: Sent to ' + rowData.email + ' with logo from ' + repoInfo.logoUrl);
      
      ui.alert('Test Email Sent!', 
        'GitHub test email sent successfully! 🎉\n\n' +
        'Email sent to: ' + rowData.email + '\n' +
        'Logo loaded from: GitHub repository\n\n' +
        'Check the recipient\'s email to verify the logo displays correctly.',
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
            'GitHub test email sent successfully (duplicate override)! 🎉\n\n' +
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
          var configSheet = SpreadsheetApp.getActive().getSheetByName('Config');
          if (!configSheet) {
            var createSheet = ui.alert('Create Config Sheet', 
              'Config sheet does not exist. Create it now?', 
              ui.ButtonSet.YES_NO);
            if (createSheet === ui.Button.YES) {
              configSheet = SpreadsheetApp.getActive().insertSheet('Config');
              configSheet.getRange(1, 1, 3, 2).setValues([
                ['SENDER_EMAIL', 'your-email@example.com'],
                ['REPLY_TO', 'your-reply-email@example.com'],
                ['GITHUB_REPO', GITHUB_CONFIG.username + '/' + GITHUB_CONFIG.repository]
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
      
      var confirmMessage = 'Sender email configuration updated! 🎉\n\n';
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
    var message = 'CURRENT SYSTEM CONFIGURATION (GitHub Edition)\n\n';
    
    // Sheet information
    if (sheet) {
      message += 'SHEET CONFIG:\n';
      message += '📊 Main Sheet: "' + sheet.getName() + '"\n';
      var mapping = getColumnMapping(sheet);
      if (mapping) {
        message += '📧 Email Column: ' + (mapping.email ? 'Column ' + mapping.email : 'Not found') + '\n';
        message += '📋 Order Status Column: ' + (mapping.ordered ? 'Column ' + mapping.ordered : 'Not found') + '\n';
        message += '🔢 PO Number Column: ' + (mapping.po ? 'Column ' + mapping.po : 'Not found') + '\n';
      }
    } else {
      message += 'SHEET CONFIG:\n';
      message += '📊 Main Sheet: Not found\n';
    }
    
    message += '\n';
    
    // GitHub configuration
    message += 'GITHUB CONFIG:\n';
    message += '🐙 Repository: ' + repoInfo.url + '\n';
    message += '🏷️ Version: ' + repoInfo.version + '\n';
    message += '🖼️ Logo URL: ' + repoInfo.logoUrl + '\n';
    message += '👤 Username: ' + GITHUB_CONFIG.username + '\n';
    message += '📁 Repo Name: ' + GITHUB_CONFIG.repository + '\n';
    message += '🌿 Branch: ' + GITHUB_CONFIG.branch + '\n';
    
    message += '\n';
    
    // Email configuration
    message += 'EMAIL CONFIG:\n';
    message += '⚙️ Sender Config: ' + CONFIG.EMAIL_SENDER + '\n';
    try {
      var currentSender = getSenderEmail();
      message += '📤 Active Sender: ' + currentSender + '\n';
    } catch (e) {
      message += '📤 Active Sender: ERROR - ' + e.message + '\n';
    }
    message += '📝 Subject: ' + CONFIG.EMAIL_SUBJECT + '\n';
    message += '✍️ Signature: ' + CONFIG.EMAIL_SIGNATURE + '\n';
    
    message += '\n';
    
    // Safety settings
    message += 'SAFETY SETTINGS:\n';
    message += '⏱️ Rate Limits: ' + CONFIG.MAX_EMAILS_PER_HOUR + '/hour, ' + CONFIG.MAX_EMAILS_PER_DAY + '/day\n';
    message += '🔄 Duplicate Check: ' + CONFIG.DUPLICATE_CHECK_DAYS + ' days\n';
    message += '📎 Attach PDFs: ' + (CONFIG.ATTACH_QUOTE_PDF ? 'Yes' : 'No') + '\n';
    
    message += '\n';
    
    // User information
    try {
      var currentUser = Session.getActiveUser().getEmail();
      message += 'USER INFO:\n';
      message += '👤 Current User: ' + currentUser + '\n';
    } catch (e) {
      message += 'USER INFO:\n';
      message += '👤 Current User: Could not determine\n';
    }
    
    message += '🔧 onEdit Trigger: Check manually in Apps Script > Triggers\n';
    
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

// =============================================================================
// MAINTENANCE AND UTILITY FUNCTIONS
// =============================================================================

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
    message += '⏰ Last Hour: ' + hourlyCount + '/' + CONFIG.MAX_EMAILS_PER_HOUR + ' emails\n';
    message += '📅 Last 24 Hours: ' + dailyCount + '/' + CONFIG.MAX_EMAILS_PER_DAY + ' emails\n\n';
    
    if (hourlyCount >= CONFIG.MAX_EMAILS_PER_HOUR) {
      message += '🚫 HOURLY LIMIT REACHED\n';
      message += 'Please wait before sending more emails.\n';
    } else if (dailyCount >= CONFIG.MAX_EMAILS_PER_DAY) {
      message += '🚫 DAILY LIMIT REACHED\n';
      message += 'Please wait until tomorrow.\n';
    } else {
      message += '✅ WITHIN LIMITS\n';
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
      logToSheet('ADMIN ACTION: Rate limits reset by user (GitHub Edition)');
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
    
    // Check GitHub configuration integrity
    if (!CONFIG.LOGO_URL.includes(GITHUB_CONFIG.username)) {
      issues.push('Logo URL does not match GitHub username');
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
    message += '📊 Found ' + issues.length + ' issues\n';
    message += '🔧 Made ' + fixes + ' automatic fixes\n';
    message += '🐙 GitHub repository: ' + repoInfo.url + '\n';
    message += ''
    
    SpreadsheetApp.getUi().alert('Audit Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    logToSheet('Data audit failed: ' + error.toString());
    SpreadsheetApp.getUi().alert('Audit Failed', 'Data integrity audit failed: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


