# GitHub Dual Logo Purchase Order Email Notification System

An automated Google Sheets system that sends professional email notifications when purchase orders are placed, featuring dual logo support with GitHub-hosted images and intelligent logo selection.

## Features

- **Dual Logo Support**: Display two different logos based on configurable strategies
- **GitHub-Hosted Images**: No Google Drive permissions needed - logos served from your public GitHub repository
- **Intelligent Logo Selection**: Conditional logic to choose logos based on email content, recipient, or order details
- **Professional Email Templates**: Modern, responsive HTML emails with fallback text versions
- **Auto-Detection**: Automatically detects spreadsheet structure and column mappings
- **Duplicate Prevention**: Prevents sending multiple emails for the same order
- **Rate Limiting**: Built-in protection against accidental spam
- **Comprehensive Logging**: Detailed audit trail of all email activities
- **Error Recovery**: Graceful handling of failures with detailed error reporting

## Logo Display Strategies

1. **Primary**: Always display the primary logo (Spectral)
2. **Secondary**: Always display the secondary logo (Purdue Prop)
3. **Both**: Display both logos side by side
4. **Conditional**: Automatically select logo based on keywords in email address or description

## Quick Setup

### 1. GitHub Repository Setup

1. Create a public GitHub repository (must be public for image hosting)
2. Create an `assets/` folder in your repository
3. Upload your logo files:
   - `spectral_logo.png` (primary logo)
   - `purdue_prop_logo.png` (secondary logo)

### 2. Google Sheets Setup

1. Create or open your purchase order spreadsheet
2. Ensure you have these columns (names are flexible):
   - Email Address
   - Name (optional)
   - Purchase Order Number
   - Description
   - Order Placed? (triggers email when set to "Yes")

### 3. Google Apps Script Configuration

1. In Google Sheets, go to Extensions > Apps Script
2. Replace all code with the provided script
3. Enable Gmail API:
   - Click Services (+ icon)
   - Add "Gmail API"
4. Update the configuration:

```javascript
var GITHUB_CONFIG = {
  username: 'your-github-username',        // Your actual GitHub username
  repository: 'your-repo-name',            // Your repository name
  branch: 'main',                          // Usually 'main'
  
  logo: {
    filename: 'spectral_logo.png',         // Primary logo filename
    altText: 'Spectral',
    maxWidth: '200px',
    maxHeight: '100px'
  },
  
  logo2: {
    filename: 'purdue_prop_logo.png',      // Secondary logo filename  
    altText: 'Purdue_Prop',
    maxWidth: '200px',
    maxHeight: '100px'
  }
};
```

### 4. Trigger Setup

1. In Apps Script, click "Triggers" (alarm clock icon)
2. Add trigger:
   - Function: `onEdit`
   - Event source: "From spreadsheet"
   - Event type: "On edit"

### 5. System Initialization

1. Save the script and return to your spreadsheet
2. Refresh the page to see the new menu
3. Use menu: "Initialize system"
4. Run: "Check setup & validate config"

## Configuration Options

### Logo Strategy Configuration

Access via menu: GitHub Integration > Configure logo strategy

- **Primary**: Always use Spectral logo
- **Secondary**: Always use Purdue Prop logo  
- **Both**: Show both logos side by side
- **Conditional**: Auto-select based on keywords

### Conditional Logic Keywords

The system checks for these keywords to determine when to use the secondary logo:

```javascript
LOGO2_KEYWORDS: ['purdue', '@purdue.edu', 'university', 'student']
```

If an email address or order description contains any of these keywords, the Purdue Prop logo will be displayed instead of the Spectral logo.

### Email Settings

```javascript
var CONFIG = {
  LOGO_STRATEGY: 'conditional',           // Logo selection strategy
  EMAIL_SUBJECT: 'Your order has been placed',
  EMAIL_SIGNATURE: '-- Purchasing Team',
  MAX_EMAILS_PER_HOUR: 50,               // Safety limits
  MAX_EMAILS_PER_DAY: 200,
  DUPLICATE_CHECK_DAYS: 30,              // Prevent duplicate emails
  // ... other settings
};
```

## Usage

### Automatic Operation
1. Add purchase order data to your spreadsheet
2. When ready, change "Order Placed?" column to "Yes"
3. Email automatically sends with appropriate logo(s)
4. Status is tracked in "Notified" column

### Manual Testing
1. Select a data row in your spreadsheet
2. Use menu: "Test email for selected row"
3. System shows which logo(s) will be used
4. Confirm to send test email

## Menu Functions

### GitHub Integration
- **Setup GitHub integration**: Configuration wizard
- **Test GitHub dual logos**: Verify both logos load correctly
- **Configure logo strategy**: Change logo display behavior
- **Debug logo selection**: Troubleshoot logo selection issues
- **View repository info**: Display current GitHub settings

### System Management
- **Check setup & validate config**: Verify system configuration
- **Initialize system**: Complete setup process
- **View email backup log**: Audit trail of all emails sent
- **View rate limit status**: Check current usage limits

## Troubleshooting

### Logos Not Displaying

1. **Repository not public**: GitHub raw URLs only work with public repositories
2. **File paths incorrect**: Verify files are in `assets/` folder with exact filenames
3. **Wrong GitHub config**: Check username and repository name in script
4. **Logo strategy issue**: Use "Debug logo selection" to see selection logic

### Common Issues

**"Configuration validation failed"**
- Update GITHUB_CONFIG with your actual repository details
- Ensure Gmail API is enabled in Apps Script
- Verify your spreadsheet has required columns

**"Required columns not found"**
- System auto-detects columns by name
- Ensure columns exist: Email, Order Status, PO Number, Description
- Column names are flexible (e.g., "Email Address", "E-mail" both work)

**"Duplicate email prevented"**
- System prevents sending same order multiple times
- Override available for testing
- Adjust DUPLICATE_CHECK_DAYS in config

### Debug Tools

Use the built-in debugging functions:

1. **Debug logo selection**: Shows which logo will be selected for a specific row
2. **Test GitHub dual logos**: Verifies logo URLs and selection logic
3. **View current configuration**: Displays all current settings
4. **Audit data integrity**: Checks for data issues

## Security & Privacy

- **Repository must be public**: Required for GitHub image hosting
- **No sensitive data**: Never commit passwords, API keys, or private information to GitHub
- **Google Apps Script**: Main script stays in your private Google Apps Script project
- **Rate limiting**: Built-in protection prevents accidental spam
- **Audit logging**: Complete record of all email activities

## File Structure

```
your-github-repo/
├── README.md                    # This file
├── assets/
│   ├── spectral_logo.png       # Primary logo
│   ├── purdue_prop_logo.png    # Secondary logo
│   └── (optional fallbacks)
└── (other repository files)
```

## Technical Details

### Email Format
- **HTML emails**: Modern, responsive design with fallback images
- **Text fallback**: Plain text version for compatibility
- **Professional styling**: Clean, corporate appearance
- **Mobile responsive**: Optimized for all devices

### GitHub Integration
- Uses GitHub raw content URLs: `https://raw.githubusercontent.com/username/repo/branch/assets/filename.png`
- Fallback chain for missing images
- Version tracking in email footers
- Repository attribution (optional)

### Google Sheets Integration
- Auto-detects column structure
- Flexible column naming
- Helper columns added automatically
- Data validation rules applied

## Support

### Getting Help
1. Check the Automation_Log sheet for detailed error messages
2. Use built-in debug functions to troubleshoot
3. Verify GitHub repository is public and contains logo files
4. Ensure Gmail API is properly enabled

### Common Questions

**Q: Can I use private GitHub repositories?**
A: No, GitHub raw URLs require public repositories for image hosting in emails.

**Q: How do I change which logo displays?**
A: Use the "Configure logo strategy" menu option or modify LOGO2_KEYWORDS for conditional logic.

**Q: Can I add more logos?**
A: The current system supports two logos. Additional logos would require code modifications.

**Q: What image formats are supported?**
A: PNG, JPG, and GIF are supported. PNG recommended for logos with transparency.

## Version History

- **v2.0 - GitHub Dual Logo Edition**: Added dual logo support, GitHub integration, conditional selection
- **v1.x**: Single logo Google Drive-based system (deprecated)

## License

This project is provided as-is for educational and business use. Modify as needed for your organization.

---

**Repository**: `https://github.com/yourusername/purchase-order-notifier`
**Version**: 2.0 - GitHub Dual Logo Edition
