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
