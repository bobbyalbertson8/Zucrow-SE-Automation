
/**
 * GitHub assets + Script Properties configuration
 */

// --- GitHub assets configuration (edit these to your repo) ---
var GITHUB_CONFIG = {
  username: 'your-github-username',
  repository: 'your-repo-name',
  branch:     'main',
  logo:  { filename: 'spectral_logo.png',  altText: 'Spectral Energies',  maxWidth: '200px', maxHeight: '100px' },
  logo2: { filename: 'purdue_prop_logo.png', altText: 'Purdue Propulsion', maxWidth: '200px', maxHeight: '100px' }
};

function generateGitHubImageUrl_(filename) {
  return 'https://raw.githubusercontent.com/' +
         GITHUB_CONFIG.username + '/' +
         GITHUB_CONFIG.repository + '/' +
         GITHUB_CONFIG.branch + '/assets/' + filename;
}

// --- Defaults (Script Properties override all of these) ---
const DEFAULTS = {
  SHEET_NAME: '',                 // blank = first sheet
  STATUS_HEADER: 'Status',
  STATUS_ORDERED_VALUE: 'Ordered',
  EMAIL_HEADER: 'Requester Email',
  NAME_HEADER: 'Requester Name',
  CC_HEADER: '',                  // optional column that contains CCs
  COST_HEADER: 'Cost',
  NOTIFIED_HEADER: 'Notified',
  LOGO_STRATEGY: 'both',          // 'primary' | 'secondary' | 'both' | 'conditional'
  LOGO2_KEYWORDS: ['purdue', '@purdue.edu', 'university', 'student'],
  LOGO_PRIMARY_ALT: 'Primary Logo',
  LOGO_SECONDARY_ALT: 'Secondary Logo',
  CURRENCY: 'USD'
};

function getConfig_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const cfg = Object.assign({}, DEFAULTS, props);

  // Resolve logo URLs from GitHub assets
  cfg.LOGO_URL  = generateGitHubImageUrl_(GITHUB_CONFIG.logo.filename);
  cfg.LOGO2_URL = generateGitHubImageUrl_(GITHUB_CONFIG.logo2.filename);
  cfg.LOGO_ALT_TEXT  = GITHUB_CONFIG.logo.altText || cfg.LOGO_PRIMARY_ALT;
  cfg.LOGO2_ALT_TEXT = GITHUB_CONFIG.logo2.altText || cfg.LOGO_SECONDARY_ALT;
  cfg.LOGO_MAX_WIDTH  = GITHUB_CONFIG.logo.maxWidth  || '200px';
  cfg.LOGO_MAX_HEIGHT = GITHUB_CONFIG.logo.maxHeight || '100px';
  cfg.LOGO2_MAX_WIDTH  = GITHUB_CONFIG.logo2.maxWidth  || '200px';
  cfg.LOGO2_MAX_HEIGHT = GITHUB_CONFIG.logo2.maxHeight || '100px';

  return cfg;
}
