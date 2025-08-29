/**
 * Ultra-Robust Configuration System
 * Supports complete purchase order lifecycle with multiple status triggers
 */

// --- GitHub assets configuration (edit these to your repo) ---
var GITHUB_CONFIG = {
  username: 'bobbyalbertson8',  // UPDATE THIS
  repository: 'zucrow-se-purchase-notifier',      // UPDATE THIS
  branch: 'main',
  logo: { 
    filename: 'spectral_logo.png', 
    altText: 'Spectral Energies', 
    maxWidth: '200px', 
    maxHeight: '100px' 
  },
  logo2: { 
    filename: 'purdue_prop_logo.png', 
    altText: 'Purdue Propulsion', 
    maxWidth: '200px', 
    maxHeight: '100px' 
  }
};

function generateGitHubImageUrl_(filename) {
  return 'https://raw.githubusercontent.com/' +
         GITHUB_CONFIG.username + '/' +
         GITHUB_CONFIG.repository + '/' +
         GITHUB_CONFIG.branch + '/assets/' + filename;
}

// --- Enhanced defaults with complete lifecycle support ---
const DEFAULTS = {
  // Basic sheet configuration
  SHEET_NAME: '',                           // blank = first sheet
  EMAIL_HEADER: 'Email',                    // Column containing requester email
  NAME_HEADER: 'Name',                      // Column containing requester name
  CC_HEADER: '',                            // CC column (if exists)
  COST_HEADER: 'Cost ($)',                  // Column containing cost
  NOTIFIED_HEADER: 'Notified',              // Column for tracking notifications
  
  // Status configuration
  STATUS_HEADER: 'Order Placed?',           // The column that triggers status changes
  
  // Multiple status trigger values
  STATUS_ORDERED_VALUE: 'Yes',              // Triggers "order placed" email
  STATUS_TRANSIT_VALUE: 'In Transit',       // Triggers "in transit" email  
  STATUS_RECEIVED_VALUE: 'Received',        // Triggers "received" email
  
  // Purchasing team notification settings
  PURCHASING_EMAIL: '',                     // Primary purchasing team email
  PURCHASING_CC_EMAIL: '',                  // CC purchasing team email (optional)
  
  // Visual and branding
  LOGO_STRATEGY: 'conditional',             // Logo selection strategy
  LOGO2_KEYWORDS: ['purdue', '@purdue.edu', 'university', 'student'],
  LOGO_PRIMARY_ALT: 'Spectral Energies',
  LOGO_SECONDARY_ALT: 'Purdue Propulsion',
  CURRENCY: 'USD'
};

function getConfig_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const cfg = Object.assign({}, DEFAULTS, props);

  // Resolve logo URLs from GitHub assets
  cfg.LOGO_URL = generateGitHubImageUrl_(GITHUB_CONFIG.logo.filename);
  cfg.LOGO2_URL = generateGitHubImageUrl_(GITHUB_CONFIG.logo2.filename);
  cfg.LOGO_ALT_TEXT = GITHUB_CONFIG.logo.altText || cfg.LOGO_PRIMARY_ALT;
  cfg.LOGO2_ALT_TEXT = GITHUB_CONFIG.logo2.altText || cfg.LOGO_SECONDARY_ALT;
  cfg.LOGO_MAX_WIDTH = GITHUB_CONFIG.logo.maxWidth || '200px';
  cfg.LOGO_MAX_HEIGHT = GITHUB_CONFIG.logo.maxHeight || '100px';
  cfg.LOGO2_MAX_WIDTH = GITHUB_CONFIG.logo2.maxWidth || '200px';
  cfg.LOGO2_MAX_HEIGHT = GITHUB_CONFIG.logo2.maxHeight || '100px';

  return cfg;
}

/**
 * Enhanced setup function for all script properties
 */
function setupEnhancedProperties() {
  const props = PropertiesService.getScriptProperties();
  
  const settings = {
    // Basic headers - adjust these to match your exact column names
    'STATUS_HEADER': 'Order Placed?',
    'EMAIL_HEADER': 'Email',
    'NAME_HEADER': 'Name', 
    'COST_HEADER': 'Cost ($)',
    'NOTIFIED_HEADER': 'Notified',
    
    // Status trigger values - adjust these to match your dropdown options
    'STATUS_ORDERED_VALUE': 'Yes',           // What appears in dropdown for "ordered"
    'STATUS_TRANSIT_VALUE': 'In Transit',    // What appears in dropdown for "in transit"
    'STATUS_RECEIVED_VALUE': 'Received',     // What appears in dropdown for "received"
    
    // Other settings
    'CURRENCY': 'USD',
    'LOGO_STRATEGY': 'conditional'
  };
  
  props.setProperties(settings);
  console.log('Enhanced script properties configured:', settings);
  console.log('✅ System now supports complete purchase order lifecycle');
  console.log('Status triggers configured:');
  console.log('- Ordered:', settings.STATUS_ORDERED_VALUE);
  console.log('- In Transit:', settings.STATUS_TRANSIT_VALUE); 
  console.log('- Received:', settings.STATUS_RECEIVED_VALUE);
}

/**
 * Setup purchasing team email addresses
 */
function setupPurchasingTeamEmails() {
  const props = PropertiesService.getScriptProperties();
  
  // UPDATE THESE EMAIL ADDRESSES FOR YOUR PURCHASING TEAM
  const purchasingSettings = {
    'PURCHASING_EMAIL': 'rjalbert@purdue.edu',        // CHANGE THIS
    'PURCHASING_CC_EMAIL': 'rjalbert@purdue.edu'        // CHANGE THIS (optional)
  };
  
  props.setProperties(purchasingSettings);
  console.log('Purchasing team emails configured:', purchasingSettings);
  console.log('Form submissions will now notify the purchasing team');
}

/**
 * Complete setup function - configures everything at once
 */
function setupCompleteEnhancedConfiguration() {
  console.log('Setting up complete enhanced configuration...');
  
  setupEnhancedProperties();
  setupPurchasingTeamEmails();
  
  console.log('✅ Complete enhanced configuration setup finished!');
  console.log('');
  console.log('SYSTEM CAPABILITIES:');
  console.log('📝 Form submissions → Purchasing team notifications');
  console.log('✅ Order placed → Requester confirmation');
  console.log('🚚 In transit → Requester tracking update');
  console.log('📦 Received → Requester delivery confirmation');
  console.log('');
  console.log('NEXT STEPS:');
  console.log('1. Run testConfiguration() to verify setup');
  console.log('2. Install onFormSubmit trigger for form notifications');
  console.log('3. Test with runEnhancedSystemTest()');
}

/**
 * Customize status trigger values for your specific dropdown options
 */
function customizeStatusTriggers() {
  console.log('=== Customize Status Triggers ===');
  console.log('Current default values:');
  console.log('- Ordered: "Yes"');
  console.log('- In Transit: "In Transit"');
  console.log('- Received: "Received"');
  console.log('');
  console.log('If your dropdown uses different values, update the setupEnhancedProperties() function');
  console.log('or run this function with your custom values:');
  
  // Example of how to set custom values
  const props = PropertiesService.getScriptProperties();
  
  // UNCOMMENT AND MODIFY THESE IF YOUR DROPDOWN VALUES ARE DIFFERENT:
  /*
  const customStatusValues = {
    'STATUS_ORDERED_VALUE': 'Approved',           // If your dropdown says "Approved" instead of "Yes"
    'STATUS_TRANSIT_VALUE': 'Shipped',            // If your dropdown says "Shipped" instead of "In Transit"
    'STATUS_RECEIVED_VALUE': 'Delivered',         // If your dropdown says "Delivered" instead of "Received"
  };
  
  props.setProperties(customStatusValues);
  console.log('Custom status triggers configured:', customStatusValues);
  */
  
  console.log('Current configured values:');
  const currentProps = props.getProperties();
  console.log('- Ordered trigger:', currentProps.STATUS_ORDERED_VALUE || 'NOT SET');
  console.log('- Transit trigger:', currentProps.STATUS_TRANSIT_VALUE || 'NOT SET');
  console.log('- Received trigger:', currentProps.STATUS_RECEIVED_VALUE || 'NOT SET');
}

/**
 * Advanced configuration for complex sheet structures
 */
function setupAdvancedFieldMapping() {
  console.log('=== Advanced Field Mapping Setup ===');
  
  const props = PropertiesService.getScriptProperties();
  
  // If your column names are different, update these mappings
  const advancedSettings = {
    // Adjust these if your column names don't match the defaults
    'STATUS_HEADER': 'Order Placed?',                    // The dropdown column
    'EMAIL_HEADER': 'Email',                             // Requester email column
    'NAME_HEADER': 'Name',                               // Requester name column
    'COST_HEADER': 'Cost ($)',                           // Cost column
    
    // Advanced field mappings (the system will auto-detect these)
    'DESCRIPTION_HEADER': 'Purchase order description (program and pur',  // Description column
    'QUANTITY_HEADER': 'Quantity',                       // Quantity column  
    'PO_NUMBER_HEADER': 'Purchase Order Number',         // PO number column
    'QUOTE_PDF_HEADER': 'Quote PDF',                     // Quote/PDF column
    'MESSAGE_KEY_HEADER': 'MessageKey',                  // Tracking/message key column
    'TIMESTAMP_HEADER': 'Timestamp',                     // Timestamp column
  };
  
  props.setProperties(advancedSettings);
  console.log('Advanced field mappings configured:', advancedSettings);
  console.log('✅ System will now use enhanced field detection');
}

/**
 * Validate current configuration and suggest improvements
 */
function validateAndSuggestConfiguration() {
  console.log('=== Configuration Validation & Suggestions ===');
  
  try {
    const cfg = getConfig_();
    const sheet = getActiveSheet_({SHEET_NAME: cfg.SHEET_NAME});
    const headerMap = getHeaderMap_(sheet);
    
    console.log('📋 CURRENT SHEET ANALYSIS:');
    console.log('Sheet name:', sheet.getName());
    console.log('Total columns:', sheet.getLastColumn());
    console.log('Data rows:', Math.max(0, sheet.getLastRow() - 1));
    console.log('');
    
    console.log('🔍 HEADER DETECTION:');
    Object.keys(headerMap).forEach(header => {
      console.log(`- "${header}" → Column ${headerMap[header]}`);
    });
    console.log('');
    
    // Validate required headers
    console.log('✅ REQUIRED HEADER VALIDATION:');
    const requiredHeaders = ['STATUS_HEADER', 'EMAIL_HEADER', 'NAME_HEADER', 'COST_HEADER'];
    let allValid = true;
    
    requiredHeaders.forEach(configKey => {
      const headerName = cfg[configKey];
      const found = headerMap[headerName];
      const status = found ? '✅ FOUND' : '❌ MISSING';
      console.log(`${configKey} (${headerName}): ${status}`);
      if (!found) allValid = false;
    });
    console.log('');
    
    // Validate status triggers
    console.log('🔄 STATUS TRIGGER VALIDATION:');
    console.log(`Ordered trigger: "${cfg.STATUS_ORDERED_VALUE}"`);
    console.log(`Transit trigger: "${cfg.STATUS_TRANSIT_VALUE}"`);
    console.log(`Received trigger: "${cfg.STATUS_RECEIVED_VALUE}"`);
    console.log('');
    
    // Validate email configuration
    console.log('📧 EMAIL CONFIGURATION:');
    console.log(`Purchasing email: ${cfg.PURCHASING_EMAIL || '❌ NOT SET'}`);
    console.log(`Purchasing CC: ${cfg.PURCHASING_CC_EMAIL || 'Not set (optional)'}`);
    console.log('');
    
    // Provide suggestions
    console.log('💡 SUGGESTIONS:');
    if (!allValid) {
      console.log('⚠️  Some required headers are missing. Check your column names.');
    }
    if (!cfg.PURCHASING_EMAIL) {
      console.log('⚠️  Set up purchasing team emails with setupPurchasingTeamEmails()');
    }
    if (sheet.getLastRow() < 3) {
      console.log('⚠️  Add some test data rows to fully test the system');
    }
    
    console.log('');
    console.log('🚀 RECOMMENDED NEXT STEPS:');
    console.log('1. Fix any missing headers or configuration');
    console.log('2. Run setupCompleteEnhancedConfiguration() if needed');
    console.log('3. Test with runEnhancedSystemTest()');
    console.log('4. Install form submit trigger if using Google Forms');
    
    return allValid;
    
  } catch (error) {
    console.error('❌ Configuration validation failed:', error.toString());
    return false;
  }
}

/**
 * Quick setup wizard for new installations
 */
function runQuickSetupWizard() {
  console.log('🧙‍♂️ QUICK SETUP WIZARD 🧙‍♂️');
  console.log('');
  console.log('This wizard will configure your enhanced purchase order system...');
  console.log('');
  
  // Step 1: Basic configuration
  console.log('Step 1/4: Setting up basic configuration...');
  setupEnhancedProperties();
  console.log('✅ Basic configuration complete');
  console.log('');
  
  // Step 2: Validate sheet structure
  console.log('Step 2/4: Validating sheet structure...');
  const isValid = validateAndSuggestConfiguration();
  if (!isValid) {
    console.log('⚠️  Sheet validation found issues. Please review and fix them.');
  } else {
    console.log('✅ Sheet structure is valid');
  }
  console.log('');
  
  // Step 3: Email setup reminder
  console.log('Step 3/4: Email configuration...');
  const cfg = getConfig_();
  if (!cfg.PURCHASING_EMAIL) {
    console.log('⚠️  Purchasing team email not configured');
    console.log('   Run setupPurchasingTeamEmails() to set this up');
  } else {
    console.log('✅ Purchasing team emails already configured');
  }
  console.log('');
  
  // Step 4: Testing reminder
  console.log('Step 4/4: Testing setup...');
  console.log('📋 To test your system:');
  console.log('   • Run testConfiguration() for basic validation');
  console.log('   • Run runEnhancedSystemTest() for complete testing');
  console.log('   • Install onFormSubmit trigger if using Google Forms');
  console.log('');
  
  console.log('🎉 QUICK SETUP WIZARD COMPLETE! 🎉');
  console.log('Your enhanced purchase order system is ready for testing.');
}
