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