
/**
 * Email and logos
 */

function selectLogos_(data, cfg) {
  const logos = [];
  switch (cfg.LOGO_STRATEGY) {
    case 'primary':
      logos.push({ url: cfg.LOGO_URL,  alt: cfg.LOGO_ALT_TEXT,  maxW: cfg.LOGO_MAX_WIDTH,  maxH: cfg.LOGO_MAX_HEIGHT });
      break;
    case 'secondary':
      logos.push({ url: cfg.LOGO2_URL, alt: cfg.LOGO2_ALT_TEXT, maxW: cfg.LOGO2_MAX_WIDTH, maxH: cfg.LOGO2_MAX_HEIGHT });
      break;
    case 'both':
      logos.push({ url: cfg.LOGO_URL,  alt: cfg.LOGO_ALT_TEXT,  maxW: cfg.LOGO_MAX_WIDTH,  maxH: cfg.LOGO_MAX_HEIGHT });
      logos.push({ url: cfg.LOGO2_URL, alt: cfg.LOGO2_ALT_TEXT, maxW: cfg.LOGO2_MAX_WIDTH, maxH: cfg.LOGO2_MAX_HEIGHT });
      break;
    case 'conditional':
      const emailText = (data.recipient || '').toLowerCase();
      const descText  = (data.item || data.description || '').toLowerCase();
      const hit = (cfg.LOGO2_KEYWORDS || []).some(k => emailText.includes(k.toLowerCase()) || descText.includes(k.toLowerCase()));
      if (hit) {
        logos.push({ url: cfg.LOGO2_URL, alt: cfg.LOGO2_ALT_TEXT, maxW: cfg.LOGO2_MAX_WIDTH, maxH: cfg.LOGO2_MAX_HEIGHT });
      } else {
        logos.push({ url: cfg.LOGO_URL,  alt: cfg.LOGO_ALT_TEXT,  maxW: cfg.LOGO_MAX_WIDTH,  maxH: cfg.LOGO_MAX_HEIGHT });
      }
      break;
    default:
      logos.push({ url: cfg.LOGO_URL,  alt: cfg.LOGO_ALT_TEXT,  maxW: cfg.LOGO_MAX_WIDTH,  maxH: cfg.LOGO_MAX_HEIGHT });
  }
  return logos;
}

function buildEmailHtml_(data, cfg) {
  const template = HtmlService.createTemplateFromFile('email');
  template.data = data;
  template.cfg = cfg;
  template.logos = selectLogos_(data, cfg);
  return template.evaluate().getContent();
}

function sendNotificationEmail_(recipient, subject, htmlBody, cc) {
  const options = { htmlBody: htmlBody };
  if (cc && String(cc).trim()) options.cc = cc;
  MailApp.sendEmail(recipient, subject, 'HTML required', options);
}
