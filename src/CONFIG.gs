/**
 * Configuration via Script Properties (no hardcoding).
 * Set in: Apps Script → Project Settings → Script properties
 */
const PROP = Object.freeze({
  SHEET_ID: 'SHEET_ID',                 // Optional: Spreadsheet ID (blank if script is bound)
  SHEET_NAME: 'SHEET_NAME',             // Optional: Tab name with responses

  // Column headers (case-insensitive exact match)
  STATUS_HEADER: 'STATUS_HEADER',       // e.g., "Ordered" or "Status"
  ORDERED_VALUES: 'ORDERED_VALUES',     // CSV: values considered "ordered" (e.g., "yes,ordered,order placed,true")
  EMAIL_HEADER: 'EMAIL_HEADER',         // Header to email requester (e.g., "Email", "Requester Email", "Email Address")
  ITEM_HEADER: 'ITEM_HEADER',           // e.g., "Item" or "Item Name"
  VENDOR_HEADER: 'VENDOR_HEADER',       // e.g., "Vendor" or "Supplier"
  LINK_HEADER: 'LINK_HEADER',           // e.g., "Link" or "URL"
  JUSTIFICATION_HEADER: 'JUSTIFICATION_HEADER', // e.g., "Justification" or "Reason"
  COST_HEADER: 'COST_HEADER',           // e.g., "Cost" or "Price" or "Amount"
  ORDER_NUMBER_HEADER: 'ORDER_NUMBER_HEADER',   // optional: "Order Number" / "PO Number"

  // Branding
  LOGO_STRATEGY: 'LOGO_STRATEGY',       // 'primary' | 'secondary' | 'both' | 'none'
  LOGO_URL: 'LOGO_URL',
  LOGO_ALT: 'LOGO_ALT',
  LOGO_MAX_WIDTH: 'LOGO_MAX_WIDTH',     // e.g., "180px"
  LOGO_MAX_HEIGHT: 'LOGO_MAX_HEIGHT',   // e.g., "80px"
  LOGO2_URL: 'LOGO2_URL',
  LOGO2_ALT: 'LOGO2_ALT',

  // Notifications (fallback recipients if you also want lab admins notified)
  NOTIFY_TO: 'NOTIFY_TO',               // comma-separated additional recipients
  CC_REQUESTER: 'CC_REQUESTER'          // "true" | "false"
});

function getProp_(key, def='') {
  const p = PropertiesService.getScriptProperties();
  const v = p.getProperty(key);
  return (v === null || v === undefined || v === '') ? def : v;
}

function getCSVProp_(key, defCSV='') {
  return getProp_(key, defCSV).split(',').map(s => s.trim()).filter(Boolean);
}

function getCfg() {
  return {
    sheetId: getProp_(PROP.SHEET_ID, ''),
    sheetName: getProp_(PROP.SHEET_NAME, ''),

    statusHeader: getProp_(PROP.STATUS_HEADER, 'Ordered'),
    orderedValues: getCSVProp_(PROP.ORDERED_VALUES, 'yes,ordered,order placed,true'),
    emailHeader: getProp_(PROP.EMAIL_HEADER, 'Email'),
    itemHeader: getProp_(PROP.ITEM_HEADER, 'Item'),
    vendorHeader: getProp_(PROP.VENDOR_HEADER, 'Vendor'),
    linkHeader: getProp_(PROP.LINK_HEADER, 'Link'),
    justificationHeader: getProp_(PROP.JUSTIFICATION_HEADER, 'Justification'),
    costHeader: getProp_(PROP.COST_HEADER, 'Cost'),
    orderNumberHeader: getProp_(PROP.ORDER_NUMBER_HEADER, 'Order Number'),

    logo: {
      strategy: getProp_(PROP.LOGO_STRATEGY, 'primary'),
      url: getProp_(PROP.LOGO_URL, ''),
      alt: getProp_(PROP.LOGO_ALT, 'Logo'),
      maxWidth: getProp_(PROP.LOGO_MAX_WIDTH, '180px'),
      maxHeight: getProp_(PROP.LOGO_MAX_HEIGHT, '80px'),
      url2: getProp_(PROP.LOGO2_URL, ''),
      alt2: getProp_(PROP.LOGO2_ALT, '')
    },

    notifyTo: getCSVProp_(PROP.NOTIFY_TO, ''),
    ccRequester: String(getProp_(PROP.CC_REQUESTER, 'false')).toLowerCase() === 'true'
  };
}

/** Seed demo properties once (edit values to your needs, then run). */
function seedPropertiesDemo() {
  PropertiesService.getScriptProperties().setProperties({
    [PROP.SHEET_ID]: '',
    [PROP.SHEET_NAME]: '',

    [PROP.STATUS_HEADER]: 'Ordered',
    [PROP.ORDERED_VALUES]: 'yes,ordered,order placed,true',
    [PROP.EMAIL_HEADER]: 'Email',
    [PROP.ITEM_HEADER]: 'Item',
    [PROP.VENDOR_HEADER]: 'Vendor',
    [PROP.LINK_HEADER]: 'Link',
    [PROP.JUSTIFICATION_HEADER]: 'Justification',
    [PROP.COST_HEADER]: 'Cost',
    [PROP.ORDER_NUMBER_HEADER]: 'Order Number',

    [PROP.LOGO_STRATEGY]: 'primary',
    [PROP.LOGO_URL]: 'https://raw.githubusercontent.com/your-org/your-repo/main/assets/logo.png',
    [PROP.LOGO_ALT]: 'Your Lab',
    [PROP.LOGO_MAX_WIDTH]: '180px',
    [PROP.LOGO_MAX_HEIGHT]: '80px',
    [PROP.LOGO2_URL]: '',
    [PROP.LOGO2_ALT]: '',

    [PROP.NOTIFY_TO]: '',
    [PROP.CC_REQUESTER]: 'false'
  }, true);
  Logger.log('Demo properties set. Update values under Project Settings → Script properties.');
}
