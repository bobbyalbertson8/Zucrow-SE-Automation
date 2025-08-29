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

function generateGitHubImageUrl(filename) {
  return 'https://raw.githubusercontent.com/' + 
         GITHUB_CONFIG.username + '/' + 
         GITHUB_CONFIG.repository + '/' + 
         GITHUB_CONFIG.branch + '/assets/' + filename;
}

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

function generateLogoFallbacks(primaryUrl) {
  var fallbacks = [
    generateGitHubImageUrl('logo.png'),
    generateGitHubImageUrl('company-logo.png'),
    generateGitHubImageUrl('default-logo.png')
  ];
  
  // Remove the primary URL from fallbacks if it's already there
  return fallbacks.filter(function(url) { return url !== primaryUrl; });
}
