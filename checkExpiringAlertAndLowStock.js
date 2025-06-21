
// --- Expiring Product Alert Configuration ---
const EXPIRING_SOON_DAYS = 30; // Products expiring within this many days are considered "nearly expiring"
const ALERT_EMAIL_THRESHOLD_1 = 35; // Send email when expiry is within this many days
const ALERT_EMAIL_THRESHOLD_2 = 1;  // Send email when expiry is within this many days (24 hours)


const NOTIFICATION_EMAIL = SCRIPT_PROPERTIES.getProperty('NOTIFICATION_EMAIL');
const HIGHLIGHT_COLOR_EXPIRED = '#FFCDD2'; // Light red
const HIGHLIGHT_COLOR_EXPIRING_SOON = '#FFF59D'; // Light yellow
const HIGHLIGHT_COLOR_NORMAL = '#FFFFFF'; // Normal (white background)


// --- Email Sending Functions ---
function sendExpiringProductEmail(subject, products, thresholdDays) {
  Logger.log('sendExpiringProductEmail called with subject: %s, products: %s, thresholdDays: %s', subject, JSON.stringify(products), thresholdDays);
  let emailBody = `<html><body><p>Dear Team,</p><p>The following product(s) are `;

  if (!Array.isArray(products)) {
    Logger.log(`sendExpiringProductEmail: Invalid 'products': ${JSON.stringify(products)} (Type: ${typeof products})`);
    return;
  }
  if (!subject || typeof subject !== 'string') {
    Logger.log(`sendExpiringProductEmail: Invalid 'subject': ${subject}`);
    return;
  }
  if (typeof thresholdDays !== 'number') {
    Logger.log(`sendExpiringProductEmail: Invalid 'thresholdDays': ${thresholdDays}`);
    return;
  }

  if (thresholdDays === 1) {
    emailBody += `<b>expiring in 1 day or have already expired</b>:</p>`;
  } else if (thresholdDays === 35) {
    emailBody += `<b>expiring in ${thresholdDays} days</b>:</p>`;
  } else {
    emailBody += `<b>expiring soon (within ${thresholdDays} days)</b>:</p>`;
  }

  emailBody += `<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">` +
    `<tr><th>Product Name</th><th>Batch Number</th><th>Expiry Date</th><th>Days Remaining</th><th>Clinic Location</th></tr>`;

  products.forEach(p => {
    emailBody += `<tr>` +
      `<td>${p.productName}</td>` +
      `<td>${p.batchNumber}</td>` +
      `<td>${p.expiryDate}</td>` +
      `<td>${p.daysRemaining}</td>` +
      `<td>${p.clinic}</td>` +
      `</tr>`;
  });

  emailBody += `</table>` +
    `<p>Please take appropriate action.</p>` +
    `<p>Best regards,<br>Your Inventory Management System</p></body></html>`;

  try {
    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: subject,
      htmlBody: emailBody
    });
    Logger.log(`Expiry alert email sent successfully to ${NOTIFICATION_EMAIL}. Subject: "${subject}"`);
  } catch (e) {
    Logger.log(`Error sending expiry alert email: ${e.message}`);
  }
}

// --- Check and Alert for Expiring Products ---
function checkAndAlertExpiringProducts() {
  const ss = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const sheet = ss.getSheetByName(GOOGLE_SHEET_NAME);
  if (!sheet) {
    Logger.log(`ERROR: Google Sheet tab named '${GOOGLE_SHEET_NAME}' not found in spreadsheet ID '${GOOGLE_SHEET_ID}'.`);
    return;
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];

  const requiredHeaders = [
    'Entry_ID',
    'Expiry_Date',
    'Product_Name',
    'Batch_Number',
    'Status',
    'Clinic_Location'
  ];

  const headerIndices = {};
  const missingHeaders = [];

  requiredHeaders.forEach(header => {
    const index = headers.indexOf(header);
    if (index === -1) {
      missingHeaders.push(header);
    }
    headerIndices[header] = index;
  });

  if (missingHeaders.length > 0) {
    Logger.log(`ERROR: Missing required headers for expiry check: ${missingHeaders.join(', ')}. Please ensure exact spelling and capitalization.`);
    return;
  }

  const entryIdCol = headerIndices['Entry_ID'];
  const expiryDateCol = headerIndices['Expiry_Date'];
  const productNameCol = headerIndices['Product_Name'];
  const batchNumberCol = headerIndices['Batch_Number'];
  const statusCol = headerIndices['Status'];
  const clinicLocationCol = headerIndices['Clinic_Location'];

  const today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize today to start of day

  const expiringSoonProducts = []; // For sheet highlighting, not direct email
  const oneDayAlertProducts = [];
  const thirtyFiveDayAlertProducts = [];
  const expiredProducts = [];

  // Retrieve previous email alert status from User Properties
  const userProperties = PropertiesService.getUserProperties();
  const sentAlertsJson = userProperties.getProperty('sentExpiringProductAlerts');
  const sentAlerts = sentAlertsJson ? JSON.parse(sentAlertsJson) : {};
  Logger.log('Loaded sentExpiringProductAlerts: ' + JSON.stringify(sentAlerts));

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const entryId = row[entryIdCol];
    const expiryDateValue = row[expiryDateCol];
    const productName = row[productNameCol];
    const batchNumber = row[batchNumberCol];
    const currentStatus = row[statusCol];
    const clinicLocation = row[clinicLocationCol];

    if (!expiryDateValue) continue; // Skip if no expiry date

    const expiryDate = new Date(expiryDateValue);
    expiryDate.setHours(0, 0, 0, 0); // Normalize expiry date to start of day

    const diffTime = expiryDate.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); // Days remaining including today

    let newStatus = currentStatus;
    let highlightColor = null;

    // --- Status and Highlighting Logic ---
  // const expiringSoonProducts = []; // For sheet highlighting, not direct email
  // const oneDayAlertProducts = [];
  // const thirtyFiveDayAlertProducts = [];
  // const expiredProducts = [];

//   const EXPIRING_SOON_DAYS = 30; // Products expiring within this many days are considered "nearly expiring"
// const ALERT_EMAIL_THRESHOLD_1 = 35; // Send email when expiry is within this many days
// const ALERT_EMAIL_THRESHOLD_2 = 1;  // Send email when expiry is within this many days (24 hours)

    if (diffDays <= 0) {
      newStatus = 'Expired';
      highlightColor = HIGHLIGHT_COLOR_EXPIRED;
      expiredProducts.push({
        rowNum: i + 1,
        entryId,
        productName,
        batchNumber,
        expiryDate: Utilities.formatDate(expiryDate, ss.getSpreadsheetTimeZone(), "MMM dd, yyyy"),
        daysRemaining: 'EXPIRED',
        clinic: clinicLocation
      });

    } else if (diffDays > EXPIRING_SOON_DAYS && diffDays <= ALERT_EMAIL_THRESHOLD_1) {
      newStatus = `${diffDays} expiring`;
      highlightColor = HIGHLIGHT_COLOR_EXPIRING_SOON;
      thirtyFiveDayAlertProducts.push({
        rowNum: i + 1,
        entryId,
        productName,
        batchNumber,
        expiryDate: Utilities.formatDate(expiryDate, ss.getSpreadsheetTimeZone(), "MMM dd, yyyy"),
        daysRemaining: diffDays,
        clinic: clinicLocation
      });

    } else if (diffDays > ALERT_EMAIL_THRESHOLD_2 && diffDays <= EXPIRING_SOON_DAYS) {
      newStatus = `${diffDays} expiring`;
      highlightColor = HIGHLIGHT_COLOR_EXPIRING_SOON;
      expiringSoonProducts.push({
        rowNum: i + 1,
        entryId,
        productName,
        batchNumber,
        expiryDate: Utilities.formatDate(expiryDate, ss.getSpreadsheetTimeZone(), "MMM dd, yyyy"),
        daysRemaining: diffDays,
        clinic: clinicLocation
      });

    } else if (diffDays === ALERT_EMAIL_THRESHOLD_2) {
      newStatus = `${diffDays} expiring`;
      highlightColor = HIGHLIGHT_COLOR_EXPIRING_SOON;
      oneDayAlertProducts.push({
        rowNum: i + 1,
        entryId,
        productName,
        batchNumber,
        expiryDate: Utilities.formatDate(expiryDate, ss.getSpreadsheetTimeZone(), "MMM dd, yyyy"),
        daysRemaining: diffDays,
        clinic: clinicLocation
      });

    } else {
      newStatus = 'Active';
      highlightColor = HIGHLIGHT_COLOR_NORMAL;
    }


    // Apply status update if changed
    if (newStatus !== currentStatus) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      Logger.log(`Row ${i + 1}: Status updated to '${newStatus}' for ${productName} (ID: ${entryId}).`);
    }

    // Apply highlighting
    const currentRowRange = sheet.getRange(i + 1, 1, 1, headers.length);
    if (currentRowRange.getBackground() !== highlightColor) { // Only update if color needs changing
      currentRowRange.setBackground(highlightColor);
      Logger.log(`Row ${i + 1}: Background color updated to '${highlightColor}' for ${productName} (ID: ${entryId}).`);
    }

    // // --- Email Alert Logic ---
    // const productAlertStatus = sentAlerts[entryId] || {};


    // // 35-day alert
    // if (diffDays === ALERT_EMAIL_THRESHOLD_1 && !productAlertStatus.thirtyFiveDaySent) {
    //   thirtyFiveDayAlertProducts.push({
    //     productName: productName,
    //     batchNumber: batchNumber,
    //     expiryDate: Utilities.formatDate(expiryDate, ss.getSpreadsheetTimeZone(), "MMM dd, yyyy"),
    //     daysRemaining: diffDays,
    //     clinic: clinicLocation
    //   });
    //   productAlertStatus.thirtyFiveDaySent = true;
    //   Logger.log(`Preparing 35-day email alert for ${productName} (ID: ${entryId}).`);
    // }

    // // 1-day alert (or expired)
    // // Check if diffDays is 1 (today is the day before expiry) OR diffDays is 0 or less (product expired today or in the past)
    // if (diffDays <= ALERT_EMAIL_THRESHOLD_2 && !productAlertStatus.oneDaySent) {
    //   oneDayAlertProducts.push({
    //     productName: productName,
    //     batchNumber: batchNumber,
    //     expiryDate: Utilities.formatDate(expiryDate, ss.getSpreadsheetTimeZone(), "MMM dd, yyyy"),
    //     daysRemaining: diffDays,
    //     clinic: clinicLocation
    //   });
    //   productAlertStatus.oneDaySent = true;
    //   Logger.log(`Preparing 1-day/expired email alert for ${productName} (ID: ${entryId}).`);
    // }

    // // Store updated alert status for this product
    // sentAlerts[entryId] = productAlertStatus;
  }

  // Save updated email alert status
  userProperties.setProperty('sentExpiringProductAlerts', JSON.stringify(sentAlerts));
  Logger.log('Updated expiry email alert status in user properties.');

  // --- Compile and Send Expiring Product Emails ---
  Logger.log('Contents of thirtyFiveDayAlertProducts before sending email: ' + JSON.stringify(thirtyFiveDayAlertProducts));
  if (thirtyFiveDayAlertProducts.length > 0) {
    sendExpiringProductEmail(
      `Zoho Inventory Alert: ${thirtyFiveDayAlertProducts.length} Product(s) Expiring in ${ALERT_EMAIL_THRESHOLD_1} Days!`,
      thirtyFiveDayAlertProducts,
      ALERT_EMAIL_THRESHOLD_1
    );
  }

  Logger.log('Contents of oneDayAlertProducts before sending email: ' + JSON.stringify(oneDayAlertProducts));
  if (oneDayAlertProducts.length > 0) {
    sendExpiringProductEmail(
      `Zoho Inventory URGENT Alert: ${oneDayAlertProducts.length} Product(s) Expiring Soon/Expired!`,
      oneDayAlertProducts,
      ALERT_EMAIL_THRESHOLD_2
    );
  }

  Logger.log('Contents of expiringSoonProducts before sending email: ' + JSON.stringify(expiringSoonProducts));
  if (expiringSoonProducts.length > 0) {
    sendExpiringProductEmail(
      `Zoho Inventory URGENT Alert: ${expiringSoonProducts.length} Product(s) Expiring Soon`,
      expiringSoonProducts,
      ALERT_EMAIL_THRESHOLD_2
    );
  }

  Logger.log('Contents of expiredProducts before sending email: ' + JSON.stringify(expiredProducts));
  if (expiredProducts.length > 0) {
    sendExpiringProductEmail(
      `Zoho Inventory URGENT Alert: ${expiredProducts.length} Product(s) Expired`,
      expiredProducts,
      ALERT_EMAIL_THRESHOLD_2
    );
  }

  Logger.log('Finished checking for expiring products, updating sheet, and sending alerts.');
  try {
    SpreadsheetApp.getUi().alert('Inventory Expiry Check Complete', 'Expired and expiring products have been checked, statuses/highlights updated, and emails sent if necessary.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('UI alert skipped for checkAndAlertExpiringProducts as no active Spreadsheet UI context was found.');
  }
}

function sendLowStockEmail(subject, products) {
  let emailBody = `<html><body><p>Dear Team,</p><p>The following product(s) are running <b>low on stock</b>:</p>` +
    `<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">` +
    `<tr><th>Product Name</th><th>Batch Number</th><th>Current Quantity</th><th>Clinic Location</th></tr>`;

  products.forEach(p => {
    emailBody += `<tr>` +
      `<td>${p.productName}</td>` +
      `<td>${p.batchNumber}</td>` +
      `<td>${p.currentQuantity}</td>` +
      `<td>${p.clinic}</td>` +
      `</tr>`;
  });

  emailBody += `</table>` +
    `<p>Please consider restocking these items.</p>` +
    `<p>Best regards,<br>Your Inventory Management System</p></body></html>`;

  try {
    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: subject,
      htmlBody: emailBody
    });
    Logger.log(`Low stock alert email sent successfully to ${NOTIFICATION_EMAIL}. Subject: "${subject}"`);
  } catch (e) {
    Logger.log(`Error sending low stock alert email: ${e.message}`);
  }
}


// --- Check and Alert for Low Stock Products ---
function checkAndAlertLowStockProducts() {
  const LOW_STOCK_THRESHOLD = 5;
  const HIGHLIGHT_COLOR_LOW_STOCK = '#FFECB3'; // Light orange

  const ss = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const sheet = ss.getSheetByName(GOOGLE_SHEET_NAME);
  if (!sheet) {
    Logger.log(`ERROR: Google Sheet tab named '${GOOGLE_SHEET_NAME}' not found in spreadsheet ID '${GOOGLE_SHEET_ID}'.`);
    return;
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];

  const requiredHeaders = [
    'Entry_ID',
    'Batch_Number',
    'Product_Name',
    'Current_Stock_Count',
    'Clinic_Location',
    'Status'
  ];

  const headerIndices = {};
  const missingHeaders = [];

  requiredHeaders.forEach(header => {
    const index = headers.indexOf(header);
    if (index === -1) {
      missingHeaders.push(header);
    }
    headerIndices[header] = index;
  });

  if (missingHeaders.length > 0) {
    Logger.log(`ERROR: Missing required headers for low stock check: ${missingHeaders.join(', ')}. Please ensure exact spelling and capitalization.`);
    return;
  }

  const entryIdCol = headerIndices['Entry_ID'];
  const batchNumberCol = headerIndices['Batch_Number'];
  const productNameCol = headerIndices['Product_Name'];
  const quantityCol = headerIndices['Current_Stock_Count'];
  const clinicCol = headerIndices['Clinic_Location'];
  const statusCol = headerIndices['Status'];

  const lowStockProducts = [];

  const userProperties = PropertiesService.getUserProperties();
  const sentLowStockAlertsJson = userProperties.getProperty('sentLowStockAlerts');
  const sentLowStockAlerts = sentLowStockAlertsJson ? JSON.parse(sentLowStockAlertsJson) : {};
  Logger.log('Loaded sentLowStockAlerts: ' + JSON.stringify(sentLowStockAlerts));

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const entryId = row[entryIdCol];
    const batchNumber = row[batchNumberCol];
    const productName = row[productNameCol];
    const quantity = parseFloat(row[quantityCol]);
    const clinic = row[clinicCol];
    const currentStatus = row[statusCol];

    if (isNaN(quantity)) {
      Logger.log(`Row ${i + 1}: Skipping due to invalid quantity.`);
      continue;
    }

    const rowNum = i + 1;
    let highlightColor = null;

    if (quantity <= LOW_STOCK_THRESHOLD) {
      highlightColor = HIGHLIGHT_COLOR_LOW_STOCK;

      if (!sentLowStockAlerts[entryId] || sentLowStockAlerts[entryId].lastNotifiedQuantity !== quantity) {
        lowStockProducts.push({
          productName,
          batchNumber,
          currentQuantity: quantity,
          clinic
        });

        // Update alert status
        sentLowStockAlerts[entryId] = {
          notified: true,
          lastNotifiedQuantity: quantity,
          timestamp: new Date().toISOString()
        };

        // Update sheet status and highlight
        sheet.getRange(rowNum, statusCol + 1).setValue('Low Stock');
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground(highlightColor);

        Logger.log(`Row ${rowNum}: Marked as Low Stock for ${productName} (Batch: ${batchNumber}, ID: ${entryId}).`);
      } else {
        Logger.log(`Row ${rowNum}: Alert already sent for ${productName} (ID: ${entryId}) at same quantity.`);
      }

    } else {
      // If stock is fine now, reset alert and status if needed
      if (sentLowStockAlerts[entryId]) {
        delete sentLowStockAlerts[entryId];
        Logger.log(`Row ${rowNum}: Reset alert for ${productName} (ID: ${entryId}) due to quantity restored.`);
      }

      if (currentStatus === 'Low Stock') {
        sheet.getRange(rowNum, statusCol + 1).setValue('Active');
        Logger.log(`Row ${rowNum}: Status reset to Active for ${productName} (ID: ${entryId}).`);
      }

      // Clear low-stock highlight only if not expiring/expired
      const expiryDateCol = headers.indexOf('Expiry_Date');
      const expiryValue = row[expiryDateCol];
      let isExpiring = false;

      if (expiryDateCol !== -1 && expiryValue) {
        const expiryDate = new Date(expiryValue);
        expiryDate.setHours(0, 0, 0, 0);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const diffDays = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));
        isExpiring = diffDays <= EXPIRING_SOON_DAYS;
      }

      if (!isExpiring) {
        sheet.getRange(rowNum, 1, 1, headers.length).setBackground(HIGHLIGHT_COLOR_NORMAL);
        Logger.log(`Row ${rowNum}: Cleared low-stock highlight.`);
      }
    }
  }

  // Save updated alert state
  userProperties.setProperty('sentLowStockAlerts', JSON.stringify(sentLowStockAlerts));
  Logger.log('Updated low stock alert status in user properties.');

  if (lowStockProducts.length > 0) {
    sendLowStockEmail(
      `Zoho Inventory Alert: ${lowStockProducts.length} Product(s) Low on Stock!`,
      lowStockProducts
    );
  }

  Logger.log('Finished checking for low stock products.');
  try {
    SpreadsheetApp.getUi().alert('Low Stock Check Complete', 'Statuses, highlights, and alerts sent.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('UI alert skipped (likely in headless mode).');
  }
}


function runInventoryChecks() {
  checkAndAlertExpiringProducts(); // Always runs first
  checkAndAlertLowStockProducts(); // Runs after expiry
}
