function buildEmailContent(data) {
  var greeting = data.name ? 'Hello ' + data.name + ',' : CONFIG.EMAIL_GREETING + ',';
  var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy, h:mm:ss a z");

  // Outlook-safe inline icons
  var check = "<!--[if !mso]><!--><span style=\"font-family:'Segoe UI Emoji','Segoe UI Symbol','Apple Color Emoji','Noto Color Emoji',Arial,sans-serif;\">&#10003;</span><!--<![endif]--><!--[if mso]><span style='font-family:Arial,sans-serif;'>[OK]</span><![endif]-->";
  var bolt  = "<!--[if !mso]><!--><span style=\"font-family:'Segoe UI Emoji','Segoe UI Symbol','Apple Color Emoji','Noto Color Emoji',Arial,sans-serif;\">&#9889;</span><!--<![endif]--><!--[if mso]><span style='font-family:Arial,sans-serif;'>[Lightning]</span><![endif]-->";

  // Select appropriate logo(s)
  var selectedLogos = selectLogos(data);
  var logoHtml = buildLogoHtml(selectedLogos);

  var htmlBody = '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #ffffff;">';

  // Header logos
  htmlBody += logoHtml;

  // Greeting + intro
  htmlBody += '<p style="font-size: 16px; line-height: 1.5; color: #333; margin-bottom: 20px;">' + greeting + '</p>';
  htmlBody += '<p style="font-size: 16px; line-height: 1.5; color: #333; margin-bottom: 25px;">Great news! Your order has been placed and is being processed.</p>';

  // ===== Card shell =====
  htmlBody += '<div style="background: #ffffff; border: 1px solid #e1e5e9; border-radius: 12px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1); margin: 25px 0;">';
  htmlBody += '<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px;">' +
              '<h2 style="margin: 0; font-size: 20px; font-weight: 600;">Order Details</h2></div>';

  // PO Number (table row)
  if (data.po) {
    htmlBody += "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'>" +
                  "<tr>" +
                    "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; font-weight:600; color:#555; width:140px;'>PO Number:</td>" +
                    "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; color:#333; font-family:monospace; font-size:14px;'>" + escapeHtml(data.po) + "</td>" +
                  "</tr>" +
                "</table>";
  }

  // Description (table row)
  if (data.description) {
    htmlBody += "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'>" +
                  "<tr>" +
                    "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; font-weight:600; color:#555; width:140px;'>Description:</td>" +
                    "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; color:#333;'>" + escapeHtml(data.description) + "</td>" +
                  "</tr>" +
                "</table>";
  }

  // Status row (icon + text, table-based for alignment)
  htmlBody += "<div style='padding: 14px 20px; border-top: 1px solid #f0f0f0; border-bottom: 1px solid #f0f0f0;'>" +
                "<table role='presentation' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'><tr>" +
                  "<td style='vertical-align:top; padding-right:8px; font-size:16px; line-height:1; color:#28a745; font-weight:700;'>" + check + "</td>" +
                  "<td style='vertical-align:top; font-size:16px; line-height:1; color:#28a745; font-weight:700;'>Order Placed</td>" +
                "</tr></table>" +
              "</div>";

  // Date row (formatted)
  htmlBody += "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'>" +
                "<tr>" +
                  "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; font-weight:600; color:#555; width:140px;'>Date:</td>" +
                  "<td style='padding:16px 20px; border-top:1px solid #f0f0f0; color:#333;'>" + formattedDate + "</td>" +
                "</tr>" +
              "</table>";

  // End card
  htmlBody += '</div>';

  // Next steps
  htmlBody += '<div style="background: linear-gradient(135deg, #e3f2fd 0%, #f3e5f5 100%); border-radius: 8px; padding: 20px; margin: 25px 0;">' +
              '<h3 style="margin: 0 0 10px 0; color: #1976d2; font-size: 16px;">Next Steps</h3>' +
              '<p style="margin: 0; color: #555; line-height: 1.5;">We\'ll keep you updated on your order progress. If you have any questions about your order, please reply to this email and we\'ll get back to you promptly.</p>' +
              '</div>';

  // Footer
  htmlBody += '<div style="text-align: center; margin-top: 40px; padding-top: 20px; border-top: 1px solid #e0e0e0;">' +
              '<p style="margin: 0; color: #666; font-size: 14px;">' + CONFIG.EMAIL_SIGNATURE + '</p>';

  // GitHub attribution with Outlook-safe lightning
  if (GITHUB_CONFIG.username && GITHUB_CONFIG.repository) {
    var logoInfo = selectedLogos.length > 1 ? ' (' + selectedLogos[0].altText + ')':
    htmlBody += "<p style='margin: 10px 0 0 0; font-size: 12px; color: #999;'>" +
                  "<a href='https://github.com/" + GITHUB_CONFIG.username + "/" + GITHUB_CONFIG.repository + "' style='color:#999; text-decoration:none;'>" +
                    "Powered by GitHub automation&nbsp;" + bolt + " " + (logoInfo || "") +
                  "</a></p>";
  }

  htmlBody += '</div></div>';

  // ===== Plain text version =====
  var textBody = greeting + '\n\n';
  textBody += 'Great news! Your order has been placed and is being processed.\n\n';
  textBody += 'ORDER DETAILS\n';
  textBody += '=============\n';
  if (data.po) textBody += 'PO Number: ' + data.po + '\n';
  if (data.description) textBody += 'Description: ' + data.description + '\n';
  textBody += 'Status: Order Placed [OK]\n';
  textBody += 'Date: ' + formattedDate + '\n\n';
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