# Purchase Order Email Notification System

An automated Google Sheets system that sends email notifications when purchase orders are marked as placed. Perfect for tracking and notifying users about order status changes.

## 🚀 Features

- **Automatic Email Notifications**: Sends professional emails when orders are marked as "placed"
- **Smart Sheet Detection**: Automatically finds your main data sheet
- **Duplicate Prevention**: Prevents sending duplicate emails for the same order
- **Rate Limiting**: Built-in safety limits to prevent email spam
- **Error Handling**: Comprehensive logging and error tracking
- **Flexible Column Mapping**: Works with various spreadsheet layouts
- **PDF Attachments**: Optional quote PDF attachment support
- **Admin Tools**: Testing, monitoring, and configuration tools

## 📋 Prerequisites

Before setting up this system, you'll need:

1. **Google Account** with access to Google Sheets and Gmail
2. **Google Sheets** spreadsheet with purchase order data
3. **Basic permissions** to edit Google Apps Script projects

## 🛠️ Initial Setup

### Step 1: Prepare Your Spreadsheet

Your spreadsheet should have columns for:

**Required Columns:**
- **Email Address** (e.g., "Email", "Email Address", "E-mail")
- **Purchase Order Number** (e.g., "PO Number", "Purchase Order Number", "PO #")
- **Order Description** (e.g., "Description", "Order Description", "Item Description")
- **Order Status** (e.g., "Order Placed?", "Order Placed", "Status")

**Optional Columns:**
- **Requester Name** (e.g., "Name", "Full Name", "Requester Name")
- **Quote PDF** (e.g., "Quote PDF", "Quote File") - for Google Drive file links

**Example Spreadsheet Layout:**
| Email | Name | PO Number | Description | Quote PDF | Order Placed? |
|-------|------|-----------|-------------|-----------|---------------|
| user@example.com | John Doe | PO-2024-001 | Office supplies | https://drive.google.com/... | No |

### Step 2: Open Google Apps Script

1. In your Google Sheet, go to **Extensions** → **Apps Script**
2. Delete any existing code in the script editor
3. Copy and paste the entire Purchase Order Notification System code
4. Save the project (Ctrl+S or Cmd+S)
5. Give your project a name like "Purchase Order Notifier"

### Step 3: Enable Gmail API

1. In the Apps Script editor, click **Services** (+ icon) in the left sidebar
2. Find **Gmail API** and click **Add**
3. Leave the default settings and click **Save**

### Step 4: Configure the System

1. In the script, find the `CONFIG` section at the top
2. Modify settings as needed:

```javascript
var CONFIG = {
  SHEET_NAME: '', // Leave empty for auto-detection
  EMAIL_SUBJECT: 'Your order has been placed',
  EMAIL_GREETING: 'Hello',
  EMAIL_SIGNATURE: '-- Purchasing Team',
  ATTACH_QUOTE_PDF: false, // Set to true if you want to attach PDFs
  EMAIL_SENDER: 'auto', // Uses current user's email
  // ... other settings
};
```

### Step 5: Set Up the Trigger

1. In Apps Script, click **Triggers** (clock icon) in the left sidebar
2. Click **+ Add Trigger**
3. Configure as follows:
   - **Function to run:** `onEdit`
   - **Event source:** From spreadsheet
   - **Event type:** On edit
   - **Failure notification settings:** Choose your preference
4. Click **Save**
5. Grant necessary permissions when prompted

### Step 6: Initialize the System

1. In Apps Script, select the function **`initializeSystem`** from the dropdown
2. Click **Run** (▶️ button)
3. Grant permissions when prompted:
   - View and manage spreadsheets
   - Send email on your behalf
   - Connect to external services
4. Check that initialization completed successfully

## ✅ Testing the System

### Method 1: Use the Built-in Menu

1. Go back to your Google Sheet
2. You should see a new **"Order Notifier"** menu
3. Click **Order Notifier** → **"Check setup & validate config"**
4. Resolve any issues shown

### Method 2: Test a Single Row

1. Select a row with test data in your spreadsheet
2. Go to **Order Notifier** → **"Test email for selected row"**
3. Confirm when prompted - this sends a real email!

### Method 3: Live Test

1. Change a row's "Order Placed?" status to "Yes" or "Order Placed"
2. The system should automatically send an email
3. Check the **"Automation_Log"** sheet for detailed logs

## 📧 How It Works

1. **Trigger Activation**: When you edit the "Order Placed?" column
2. **Value Check**: System checks if the new value means "yes" (accepts: "yes", "y", "true", "1", "placed", "ordered", "complete", etc.)
3. **Validation**: Verifies email address, PO number, and description are present
4. **Duplicate Check**: Ensures this exact order hasn't been emailed recently
5. **Rate Limiting**: Confirms within hourly/daily email limits
6. **Email Generation**: Creates professional HTML and text email
7. **Send & Track**: Sends email and records timestamp in spreadsheet

## 🎛️ Admin Tools

Access these through the **"Order Notifier"** menu:

- **Check setup & validate config**: Verify system configuration
- **Configure sender email**: Change who emails are sent from
- **View current configuration**: See all current settings
- **View email backup log**: See all email attempts (success/failure)
- **View rate limit status**: Check current email sending limits
- **Test email for selected row**: Send test email for debugging
- **Audit data integrity**: Check for data issues in your sheet

## 🔧 Customization

### Email Template

Modify the `buildEmailContent()` function to customize:
- Email subject line
- HTML formatting
- Email signature
- Additional information included

### Column Recognition

The system automatically recognizes columns with these names (case-insensitive):

- **Email**: "email", "email address", "e-mail"
- **Name**: "name", "full name", "requester name", "user name"
- **PO Number**: "purchase order number", "po number", "po #", "order number"
- **Description**: "purchase order description", "description", "order description"
- **Status**: "order placed?", "order placed", "ordered", "status", "order status"

To add custom column names, modify the `COLUMN_MAPPINGS` in the CONFIG section.

### "Yes" Values

The system recognizes these values as "order placed":
- "yes", "y", "true", "1"
- "placed", "ordered", "complete", "done"
- "confirmed", "approved", "processed"
- And many more...

Add custom values to `YES_VALUES` in the CONFIG section.

## 🛡️ Safety Features

### Rate Limiting
- **50 emails per hour** (configurable)
- **200 emails per day** (configurable)
- Prevents accidental mass email sending

### Duplicate Prevention
- **30-day duplicate check** (configurable)
- Prevents sending the same order notification multiple times
- Uses email + PO number + description as unique identifier

### Error Handling
- Comprehensive error logging
- Failed attempts are recorded with reasons
- System continues operating even if individual emails fail

## 📊 Monitoring & Logs

The system creates several tracking sheets:

- **Automation_Log**: Detailed system activity and errors
- **Email_Backup_Log**: Record of all email attempts
- **Rate_Limit_Log**: Email sending frequency tracking

## ❗ Troubleshooting

### Common Issues

**"Gmail API service not available"**
- Solution: Add Gmail API service in Apps Script → Services

**"No onEdit trigger found"**
- Solution: Create trigger in Apps Script → Triggers → Add Trigger

**"Could not find main data sheet"**
- Solution: Ensure your sheet has proper column headers or set `SHEET_NAME` in CONFIG

**"Email column not found"**
- Solution: Make sure you have a column with "email" in the name

**"Required columns not found"**
- Solution: Verify your sheet has email, PO number, description, and status columns

### Getting Help

1. Check the **Automation_Log** sheet for detailed error messages
2. Use **"Check setup & validate config"** from the menu
3. Verify your column headers match expected names
4. Test with a single row first using **"Test email for selected row"**

## 🔒 Security & Privacy

- Uses your Google account credentials
- Emails are sent from your Gmail account
- No external services required
- All data stays within your Google account
- Rate limiting prevents abuse

## 📝 Version History

- **v1.0**: Initial release with basic email notifications
- **v2.0**: Added duplicate prevention and rate limiting
- **v3.0**: Enhanced error handling and admin tools
- **v4.0**: Smart sheet detection and improved reliability

## 📄 License

This script is provided as-is for internal business use. Modify as needed for your organization.

---

**Need help?** Check the troubleshooting section above or review the detailed logs in your **Automation_Log** sheet.
