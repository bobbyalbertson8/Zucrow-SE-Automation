/**
 * Ultra-Robust Email System
 * Handles all lifecycle status emails with enhanced content
 */

function selectLogos_(data, cfg) {
  const logos = [];
  
  // Get email addresses to check
  const emailText = (data.recipient || data.requesterEmail || '').toLowerCase();
  
  console.log('Logo selection - checking email:', emailText);
  
  // Your specific logic: 
  // - spectralenergies.com email = Spectral logo only
  // - purdue email = both logos
  // - anything else = Spectral logo only (default)
  
  if (emailText.includes('@purdue.edu') || emailText.includes('purdue')) {
    // Purdue email - show BOTH logos (Spectral first, then Purdue)
    console.log('Purdue email detected - using both logos');
    logos.push({ 
      url: cfg.LOGO_URL,  
      alt: cfg.LOGO_ALT_TEXT,  
      maxW: cfg.LOGO_MAX_WIDTH,  
      maxH: cfg.LOGO_MAX_HEIGHT 
    });
    logos.push({ 
      url: cfg.LOGO2_URL, 
      alt: cfg.LOGO2_ALT_TEXT, 
      maxW: cfg.LOGO2_MAX_WIDTH, 
      maxH: cfg.LOGO2_MAX_HEIGHT 
    });
  } else if (emailText.includes('@spectralenergies.com') || emailText.includes('spectralenergies')) {
    // Spectral Energies email - show Spectral logo only
    console.log('Spectral Energies email detected - using Spectral logo only');
    logos.push({ 
      url: cfg.LOGO_URL,  
      alt: cfg.LOGO_ALT_TEXT,  
      maxW: cfg.LOGO_MAX_WIDTH,  
      maxH: cfg.LOGO_MAX_HEIGHT 
    });
  } else {
    // Default case - show Spectral logo only
    console.log('Default case - using Spectral logo only');
    logos.push({ 
      url: cfg.LOGO_URL,  
      alt: cfg.LOGO_ALT_TEXT,  
      maxW: cfg.LOGO_MAX_WIDTH,  
      maxH: cfg.LOGO_MAX_HEIGHT 
    });
  }
  
  console.log('Selected', logos.length, 'logo(s) for email');
  logos.forEach((logo, index) => {
    console.log(`Logo ${index + 1}: ${logo.url}`);
  });
  return logos;
}

/**
 * Enhanced email builder for all status types
 */
function buildStatusEmailHtml_(data, cfg) {
  // Select appropriate template based on email type
  const templateName = getTemplateForEmailType_(data.emailType);
  
  console.log('Building email with template:', templateName);
  console.log('Email type:', data.emailType);
  
  try {
    const template = HtmlService.createTemplateFromFile(templateName);
    template.data = data;
    template.cfg = cfg;
    template.logos = selectLogos_(data, cfg);
    
    // Add enhanced data processing
    template.enhancedData = enhanceEmailData_(data);
    
    console.log('Template data prepared with', template.logos.length, 'logos');
    
    const htmlContent = template.evaluate().getContent();
    console.log('Email HTML generated successfully, length:', htmlContent.length);
    
    return htmlContent;
  } catch (error) {
    console.error('Error building email HTML with template:', templateName, error.toString());
    console.log('Falling back to basic template');
    // Fallback to basic template
    return buildFallbackEmailHtml_(data, cfg);
  }
}

/**
 * Get template name based on email type
 */
function getTemplateForEmailType_(emailType) {
  const templateMap = {
    'ordered': 'email-ordered',
    'transit': 'email-transit', 
    'received': 'email-received',
    'submission': 'purchasing-team-email'
  };
  
  const templateName = templateMap[emailType] || 'email-ordered';
  console.log('Selected template:', templateName, 'for email type:', emailType);
  return templateName;
}

/**
 * Enhance email data with additional processing
 */
function enhanceEmailData_(data) {
  return {
    // Enhanced description processing
    shortDescription: truncateText_(data.description || data.item || '', 100),
    fullDescription: data.description || data.item || '',
    hasLongDescription: (data.description || data.item || '').length > 100,
    
    // Status-specific messaging
    statusMessage: getStatusMessage_(data.emailType, data.status),
    statusIcon: getStatusIcon_(data.emailType),
    statusColor: getStatusColor_(data.emailType),
    
    // Tracking information
    hasTracking: !!(data.messageKey || data.poNumber),
    trackingInfo: data.messageKey || data.poNumber || '',
    
    // Timeline information
    isUrgent: isUrgentOrder_(data.cost, data.costRaw),
    estimatedDelivery: getEstimatedDelivery_(data.emailType),
    
    // Enhanced formatting
    formattedCost: formatEnhancedCurrency_(data.costRaw, data.cost, cfg.CURRENCY),
    formattedDate: formatFriendlyDate_(data.timestamp)
  };
}

/**
 * Get status-specific messaging
 */
function getStatusMessage_(emailType, status) {
  const messages = {
    'ordered': 'Your purchase request has been approved and ordered. We\'ll keep you updated on the progress.',
    'transit': 'Great news! Your order is now in transit and on its way to you.',
    'received': 'Your order has been successfully delivered and received. Thank you for your purchase request!'
  };
  
  return messages[emailType] || `Status updated to: ${status}`;
}

/**
 * Get status-specific icons
 */
function getStatusIcon_(emailType) {
  const icons = {
    'ordered': '✅',
    'transit': '🚚',
    'received': '📦'
  };
  
  return icons[emailType] || '📋';
}

/**
 * Get status-specific colors
 */
function getStatusColor_(emailType) {
  const colors = {
    'ordered': '#10b981',    // Green
    'transit': '#f59e0b',    // Amber  
    'received': '#3b82f6'    // Blue
  };
  
  return colors[emailType] || '#6b7280';
}

/**
 * Check if order is urgent based on cost
 */
function isUrgentOrder_(costRaw, costFormatted) {
  try {
    const cost = parseFloat(String(costRaw || costFormatted || '0').replace(/[^0-9.-]/g, ''));
    return cost >= 1000; // Orders over $1000 are considered urgent
  } catch (e) {
    return false;
  }
}

/**
 * Get estimated delivery timeframe
 */
function getEstimatedDelivery_(emailType) {
  const estimates = {
    'ordered': '5-10 business days',
    'transit': '2-5 business days',
    'received': 'Delivered'
  };
  
  return estimates[emailType] || 'TBD';
}

/**
 * Enhanced currency formatting
 */
function formatEnhancedCurrency_(costRaw, costFormatted, currency) {
  try {
    if (costFormatted && costFormatted.includes(currency)) {
      return costFormatted; // Already formatted
    }
    
    const cost = parseFloat(String(costRaw || costFormatted || '0').replace(/[^0-9.-]/g, ''));
    if (!isNaN(cost)) {
      return `${currency} ${cost.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    }
  } catch (e) {
    console.log('Currency formatting error:', e.toString());
  }
  
  return String(costFormatted || costRaw || '0');
}

/**
 * Format timestamp in friendly way
 */
function formatFriendlyDate_(timestamp) {
  try {
    const date = new Date(timestamp);
    return date.toLocaleDateString('en-US', { 
      weekday: 'long',
      year: 'numeric', 
      month: 'long', 
      day: 'numeric',
      hour: 'numeric',
      minute: '2-digit'
    });
  } catch (e) {
    return String(timestamp || '');
  }
}

/**
 * Truncate text with ellipsis
 */
function truncateText_(text, maxLength) {
  if (!text || text.length <= maxLength) {
    return text;
  }
  return text.substring(0, maxLength - 3) + '...';
}

/**
 * Build purchasing team email (for form submissions)
 */
function buildPurchasingTeamEmailHtml_(data, cfg) {
  try {
    const template = HtmlService.createTemplateFromFile('purchasing-team-email');
    template.data = data;
    template.cfg = cfg;
    template.logos = selectLogos_(data, cfg);
    template.enhancedData = enhanceEmailData_(data);
    return template.evaluate().getContent();
  } catch (error) {
    console.error('Error building purchasing team email:', error.toString());
    return buildFallbackEmailHtml_(data, cfg);
  }
}

/**
 * Fallback email builder with logos for error cases
 */
function buildFallbackEmailHtml_(data, cfg) {
  const statusIcon = getStatusIcon_(data.emailType);
  const statusMessage = getStatusMessage_(data.emailType, data.status);
  const logos = selectLogos_(data, cfg);
  
  console.log('Building fallback email with', logos.length, 'logos');
  
  // Build logo HTML
  let logoHtml = '';
  if (logos && logos.length > 0) {
    logoHtml = `
      <div style="background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 50%, #1e40af 100%); padding:30px 40px; text-align:center; margin-bottom: 20px;">
        <div style="display:inline-block;">`;
    
    logos.forEach((logo, index) => {
      logoHtml += `
          <img src="${logo.url}" 
               alt="${logo.alt}" 
               style="max-height:80px; height:auto; max-width:${logo.maxW}; ${index > 0 ? 'margin-left:30px;' : ''} vertical-align:middle; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.15); background:white; padding:10px;" 
               onerror="this.style.display='none'"/>`;
    });
    
    logoHtml += `
        </div>
        ${logos.length > 1 ? 
          '<div style="color:white; font-size:14px; margin-top:15px; font-weight:500;">Spectral Energies & Purdue Propulsion Partnership</div>' :
          '<div style="color:white; font-size:14px; margin-top:15px; font-weight:500;">Spectral Energies Purchase Order System</div>'
        }
      </div>`;
  }
  
  return `
    <html>
    <body style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 0;">
      ${logoHtml}
      <div style="padding: 20px;">
        <div style="background: #f8fafc; padding: 30px; border-radius: 10px;">
          <h1 style="color: #1f2937; margin-bottom: 20px;">${statusIcon} Purchase Order Update</h1>
          
          ${data.name ? `<p><strong>Hello ${data.name},</strong></p>` : '<p><strong>Hello,</strong></p>'}
          
          <p>${statusMessage}</p>
          
          <div style="background: white; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #3b82f6;">
            <h3 style="margin-top: 0; color: #1f2937;">Order Details</h3>
            ${data.poNumber ? `<p><strong>PO Number:</strong> ${data.poNumber}</p>` : ''}
            ${data.description ? `<p><strong>Description:</strong> ${data.description}</p>` : ''}
            ${data.quantity ? `<p><strong>Quantity:</strong> ${data.quantity}</p>` : ''}
            ${data.cost ? `<p><strong>Cost:</strong> ${data.cost}</p>` : ''}
            <p><strong>Status:</strong> ${data.status}</p>
            <p><strong>Updated:</strong> ${data.timestamp}</p>
          </div>
          
          <p style="color: #6b7280; font-size: 14px;">
            If you have questions about your order, please reply to this email.
          </p>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * Legacy function - now routes to enhanced system
 */
function buildEmailHtml_(data, cfg) {
  // Determine email type for legacy compatibility
  data.emailType = data.emailType || 'ordered';
  return buildStatusEmailHtml_(data, cfg);
}

/**
 * Send notification email (unchanged)
 */
function sendNotificationEmail_(recipient, subject, htmlBody, cc) {
  const options = { htmlBody: htmlBody };
  if (cc && String(cc).trim()) options.cc = cc;
  MailApp.sendEmail(recipient, subject, 'HTML required', options);
}

/**
 * Test email template generation
 */
function testEmailTemplates() {
  console.log('=== Testing Email Templates ===');
  
  const cfg = getConfig_();
  
  // Create test data
  const testData = {
    name: 'Test User',
    recipient: 'test@example.com',
    requesterEmail: 'test@example.com',
    status: 'Test Status',
    cost: 'USD 1,234.56',
    costRaw: 1234.56,
    description: 'High-Performance Digital Oscilloscope with Advanced Triggering Capabilities and Data Logging Features for Laboratory Testing and Research Applications',
    item: 'Digital Oscilloscope',
    quantity: '2',
    poNumber: 'TEST-12345',
    quotePdf: 'https://example.com/quote.pdf',
    messageKey: 'MSG-TEST-789',
    rowNumber: 42,
    timestamp: new Date().toISOString()
  };
  
  // Test each email type
  const emailTypes = ['ordered', 'transit', 'received'];
  
  emailTypes.forEach(emailType => {
    console.log(`\nTesting ${emailType} email template...`);
    
    try {
      testData.emailType = emailType;
      const htmlContent = buildStatusEmailHtml_(testData, cfg);
      
      const wordCount = htmlContent.split(' ').length;
      const hasDescription = htmlContent.includes(testData.description);
      const hasEnhancedData = htmlContent.includes('enhancedData');
      
      console.log(`✅ ${emailType} template generated successfully`);
      console.log(`   - Content length: ${htmlContent.length} characters`);
      console.log(`   - Word count: ${wordCount} words`);
      console.log(`   - Includes description: ${hasDescription}`);
      console.log(`   - Uses enhanced data: ${hasEnhancedData}`);
      
    } catch (error) {
      console.error(`❌ ${emailType} template failed:`, error.toString());
    }
  });
  
  console.log('\n=== Email Template Testing Complete ===');
}

/**
 * Test enhanced data processing
 */
function testEnhancedDataProcessing() {
  console.log('=== Testing Enhanced Data Processing ===');
  
  const testData = {
    emailType: 'ordered',
    status: 'Yes',
    description: 'This is a very long description of a sophisticated piece of equipment that includes many technical specifications and details about its capabilities and intended use cases in our research facility.',
    costRaw: 2500.50,
    cost: 'USD 2,500.50',
    messageKey: 'MSG-ABC-123',
    poNumber: 'PO-2023-456',
    timestamp: new Date().toISOString()
  };
  
  try {
    const enhancedData = enhanceEmailData_(testData);
    
    console.log('✅ Enhanced data processing successful:');
    console.log('   - Short description:', enhancedData.shortDescription);
    console.log('   - Has long description:', enhancedData.hasLongDescription);
    console.log('   - Status message:', enhancedData.statusMessage);
    console.log('   - Status icon:', enhancedData.statusIcon);
    console.log('   - Status color:', enhancedData.statusColor);
    console.log('   - Is urgent:', enhancedData.isUrgent);
    console.log('   - Estimated delivery:', enhancedData.estimatedDelivery);
    console.log('   - Formatted cost:', enhancedData.formattedCost);
    console.log('   - Formatted date:', enhancedData.formattedDate);
    
  } catch (error) {
    console.error('❌ Enhanced data processing failed:', error.toString());
  }
  
  console.log('\n=== Enhanced Data Processing Test Complete ===');
}
