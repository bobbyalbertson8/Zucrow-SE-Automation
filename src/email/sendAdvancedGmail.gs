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
