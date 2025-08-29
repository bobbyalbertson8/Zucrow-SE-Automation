/**
 * GitHub assets + Script Properties configuration
 * Updated with defaults that match your sheet structure
 */

// --- GitHub assets configuration (edit these to your repo) ---
var GITHUB_CONFIG = {
  username: 'your-github-username',  // UPDATE THIS
  repository: 'your-repo-name',      // UPDATE THIS
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

// --- Updated defaults to match your sheet structure ---
const DEFAULTS = {
  SHEET_NAME: '',                           // blank = first sheet
  STATUS_HEADER: 'Order Placed?',           // Your actual column name
  STATUS_ORDERED_VALUE: 'Yes',              // The dropdown value that triggers email
  EMAIL_HEADER: 'Email',                    // Your actual column name  
  NAME_HEADER: 'Name',                      // Your actual column name
  CC_HEADER: '',                            // You don't have a CC column
  COST_HEADER: 'Cost ($)',                  // Your actual column name
  NOTIFIED_HEADER: 'Notified',              // Your actual column name
  LOGO_STRATEGY: 'conditional',             // Show Purdue logo for @purdue.edu emails
  LOGO2_KEYWORDS: ['purdue', '@purdue.edu', 'university', 'student'],
  LOGO_PRIMARY_ALT: 'Spectral Energies',
  LOGO_SECONDARY_ALT: 'Purdue Propulsion',
  CURRENCY: 'USD'
};

function getConfig_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const cfg = Object.assign({}, DEFAULTS, props);

  // Resolve logo URLs from GitHub assets (if you have them)
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
 * Helper function to set up Script Properties for your specific sheet
 * Run this once to configure the script for your sheet structure
 */
function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  
  const settings = {
    'STATUS_HEADER': 'Order Placed?',
    'STATUS_ORDERED_VALUE': 'Yes',        // Change this if your dropdown uses different values
    'EMAIL_HEADER': 'Email',
    'NAME_HEADER': 'Name', 
    'COST_HEADER': 'Cost ($)',
    'NOTIFIED_HEADER': 'Notified',
    'CURRENCY': 'USD',
    'LOGO_STRATEGY': 'conditional'        // Will show Purdue logo for @purdue.edu emails
  };
  
  props.setProperties(settings);
  console.log('Script properties configured:', settings);
  console.log('You can now test with testConfiguration()');
}
