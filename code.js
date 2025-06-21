// --- Configuration Variables (MUST BE SET IN SCRIPT PROPERTIES) ---
// Go to Project Settings (the gear icon) > Script properties > Add new property
// Key: ZOHO_CLIENT_ID, Value: [Your Zoho Client ID]
// Key: ZOHO_CLIENT_SECRET, Value: [Your Zoho Client Secret]
// Key: ZOHO_ORGANIZATION_ID, Value: [Your Zoho Organization ID]
// Key: ZOHO_APP_LINK_NAME, Value: [Your Zoho Creator Application's Link Name]
// Key: ZOHO_REPORT_LINK_NAME, Value: [Your Zoho Creator Report's Link Name]
// Key: GOOGLE_SHEET_ID, Value: [Your Google Sheet ID from its URL]

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

const ZOHO_CLIENT_ID = SCRIPT_PROPERTIES.getProperty('ZOHO_CLIENT_ID');
const ZOHO_CLIENT_SECRET = SCRIPT_PROPERTIES.getProperty('ZOHO_CLIENT_SECRET');
const ZOHO_ORGANIZATION_ID = SCRIPT_PROPERTIES.getProperty('ZOHO_ORGANIZATION_ID');
const ZOHO_APP_LINK_NAME = SCRIPT_PROPERTIES.getProperty('ZOHO_APP_LINK_NAME');
const ZOHO_REPORT_LINK_NAME = SCRIPT_PROPERTIES.getProperty('ZOHO_REPORT_LINK_NAME');
const GOOGLE_SHEET_ID = SCRIPT_PROPERTIES.getProperty('GOOGLE_SHEET_ID');


// Zoho API Scope (Permissions needed for your script)
// This scope dictates what data your script can access in Zoho Creator.
// If you change this, you MUST re-authorize your script with Zoho.
const ZOHO_API_SCOPE = 'ZohoCreator.report.READ,ZohoCreator.report.CREATE,ZohoCreator.report.UPDATE'; // Allows reading data from Creator reports.

// Google Sheet Details (less sensitive, can be kept in code or also in Script Properties)
const GOOGLE_SHEET_NAME = 'Sheet1'; // Ensure this matches your Google Sheet tab name exactly (case-sensitive)

// --- OAuth2 Service Configuration ---
// This function sets up the OAuth2 service for Zoho Creator.
// It defines how the script communicates with Zoho's authorization server.
function getZohoOAuthService() {
  return OAuth2.createService('ZohoCreator')
        // Set the base URLs for Zoho's OAuth authorization and token endpoints.
        .setAuthorizationBaseUrl('https://accounts.zoho.com/oauth/v2/auth')
        .setTokenUrl('https://accounts.zoho.com/oauth/v2/token')
        
        // Set the Client ID and Secret obtained from the Zoho Developer Console.
        .setClientId(ZOHO_CLIENT_ID)
        .setClientSecret(ZOHO_CLIENT_SECRET)
        
        // Set the API scope (permissions) that your script requires.
        .setScope(ZOHO_API_SCOPE)
        
        // Define the callback function that Zoho will redirect to after user authorization.
        // The name 'authCallback' must match the function defined below.
        .setCallbackFunction('authCallback')
        
        // Specify where to store the obtained access and refresh tokens.
        // UserProperties is secure and specific to the user running the script.
        .setPropertyStore(PropertiesService.getUserProperties());
}

// --- OAuth2 Callback Function ---
// This function is automatically called by the OAuth2 library after a user
// successfully authorizes the script via the browser. It receives the authorization
// code and exchanges it for access and refresh tokens, storing them.
function authCallback(request) {
  const service = getZohoOAuthService();
  const authorized = service.handleCallback(request); // Handles the code exchange and token storage

  if (authorized) {
    Logger.log('Zoho Creator authorization complete. Access and refresh tokens have been obtained and stored securely.');
    return HtmlService.createHtmlOutput('Success! Zoho Creator authorization complete. You can close this tab.');
  } else {
    Logger.log('Zoho Creator authorization failed. Please check the Apps Script logs for errors.');
    return HtmlService.createHtmlOutput('Denied. Zoho Creator authorization failed. Please check logs and try again.');
  }
}

// --- ONE-TIME Authorization Initiator Function (RUN THIS FIRST!) ---
// You MUST run this function ONCE to authorize your script with Zoho.
// It will check for existing tokens and, if none, provide a URL for manual authorization.
// This is the function you run from the Apps Script editor (select authorizeZohoOAuth from the dropdown).
function authorizeZohoOAuth() {
  const service = getZohoOAuthService();

  // Check if an access token already exists and is valid or can be refreshed.
  if (service.hasAccess()) {
    Logger.log('Access token already exists and is valid or refreshed. Zoho OAuth is ready.');
    Logger.log('No need to re-authorize unless the API scope changes, tokens are explicitly revoked, or issues arise.');
    return;
  }

  // If no valid tokens are found, get the authorization URL.
  // This URL will be printed to the Apps Script Logger.
  const authorizationUrl = service.getAuthorizationUrl();
  Logger.log('No valid Zoho Creator access tokens found. Initial authorization is required.');
  Logger.log('Please open the following URL in your web browser to grant your script access to Zoho Creator:');
  Logger.log('**************************************************************************************************');
  Logger.log(authorizationUrl);
  Logger.log('**************************************************************************************************');
  Logger.log('IMPORTANT: After granting permission, you will be redirected to a Google page displaying "Success!".');
  Logger.log('The "Invalid Redirect Uri" error indicates a mismatch in Zoho Developer Console settings (See point 2 below).');


  // Provide a UI alert if running from the Sheets UI (e.g., from a custom menu).
  // This helps guide the user to the logs for the URL.
  try {
    SpreadsheetApp.getUi().alert(
      'Zoho Creator Authorization Required',
      'Please open the Google Apps Script Logs (Ctrl+Enter or Cmd+Enter) to find the authorization URL. Copy and paste it into your browser to grant Zoho Creator access. Once complete, you can close that browser tab.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    // This catches errors if the script is run from the Apps Script editor without an open sheet UI.
    Logger.log('UI alert skipped as no active Spreadsheet UI context was found (running directly from editor).');
  }

  // --- CRITICAL MANUAL STEPS AFTER RUNNING THIS FUNCTION ---
  Logger.log('');
  Logger.log('--- CRITICAL CONFIGURATION STEPS ---');
  Logger.log('1. Ensure your Zoho Developer Console Client is a "Server-based Application."');
  Logger.log('2. The "Authorized Redirect URIs" in your Zoho Developer Console MUST EXACTLY MATCH the URL found in the log above.');
  Logger.log('    Example decoded Redirect URI: https://script.google.com/macros/d/YOUR_SCRIPT_ID_HERE/usercallback');
  Logger.log('    Even a single missing character, extra space, or incorrect casing will cause "Invalid Redirect Uri".');
  Logger.log('3. Your ZOHO_CLIENT_ID and ZOHO_CLIENT_SECRET in Script Properties must be correct.');
  Logger.log('------------------------------------');
}


// --- Main Data Synchronization Function ---
// This function fetches data from Zoho Creator and appends new records to a Google Sheet.
function syncZohoToGoogleSheets() {
  const service = getZohoOAuthService();

  if (!service.hasAccess()) {
    Logger.log('ERROR: No valid Zoho Creator access token available. Please run authorizeZohoOAuth() first.');
    return;
  }

  const accessToken = service.getAccessToken();
  if (!accessToken) {
    Logger.log('FATAL ERROR: Failed to obtain access token even after authorization check. Review OAuth setup.');
    return;
  }

  const ss = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  const sheet = ss.getSheetByName(GOOGLE_SHEET_NAME);
  if (!sheet) {
    Logger.log(`ERROR: Google Sheet tab named '${GOOGLE_SHEET_NAME}' not found in spreadsheet ID '${GOOGLE_SHEET_ID}'.`);
    return;
  }

  const existingData = sheet.getDataRange().getValues();
  const headers = existingData.length > 0 ? existingData[0] : [];

  const entryIdColumnIndex = headers.indexOf('Entry_ID');
  if (entryIdColumnIndex === -1) {
    Logger.log("ERROR: 'Entry_ID' column header not found in Google Sheet. Ensure your first row has 'Entry_ID' as a header.");
    return;
  }

  const existingEntryIds = new Set();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][entryIdColumnIndex]) {
      existingEntryIds.add(existingData[i][entryIdColumnIndex]);
    }
  }
  Logger.log(`Google Sheet currently contains ${existingEntryIds.size} unique records (based on 'Entry_ID').`);

  let page = 1;
  let moreRecordsToFetch = true;
  const newRowsToAppend = [];

  while (moreRecordsToFetch) {
    const zohoApiUrl = `https://creator.zoho.com/api/v2/${ZOHO_ORGANIZATION_ID}/${ZOHO_APP_LINK_NAME}/report/${ZOHO_REPORT_LINK_NAME}`; // 🔄 MODIFIED to remove looping bug

    Logger.log(`Attempting to fetch Zoho Creator data from URL: ${zohoApiUrl} (Page: ${page})`);

    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Zoho-oauthtoken ${accessToken}`
      },
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(zohoApiUrl, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if (responseCode === 200) {
        const data = JSON.parse(responseBody);
        const records = data.data;

        if (records && records.length > 0) {
          Logger.log(`Successfully fetched ${records.length} records from Zoho Creator (Page: ${page}).`);

          let newRecordsOnThisPage = 0; // ✅ ADDED

          for (const record of records) {
            const zohoRecordId = record.ID;

            if (!existingEntryIds.has(zohoRecordId)) {
              const newRow = new Array(headers.length).fill('');

              newRow[headers.indexOf('Entry_ID')] = zohoRecordId;
              newRow[headers.indexOf('Batch_Number')] = record.Batch_Number;
              newRow[headers.indexOf('Product_Name')] = record.Product_Name;
              newRow[headers.indexOf('Quantity_Received')] = record.Quantity_Received;
              newRow[headers.indexOf('Manufacturing_Date')] = record.Manufacturing_Date;
              newRow[headers.indexOf('Expiry_Date')] = record.Expiry_Date;
              newRow[headers.indexOf('Clinic_Location')] = record.Clinic_Location;
              newRow[headers.indexOf('Current_Stock_Count')] = record.Quantity_Received;
              newRow[headers.indexOf('Status')] = 'Active';

              newRowsToAppend.push(newRow);
              existingEntryIds.add(zohoRecordId);
              newRecordsOnThisPage++; // ✅ ADDED
            }
          }

          if (newRecordsOnThisPage === 0) { // ✅ ADDED
            Logger.log("All records on this page are already imported. Stopping further fetch."); // ✅ ADDED
            moreRecordsToFetch = false; // ✅ ADDED
          } else {
            page++; // 🔄 MODIFIED: increment page only if new records were found
          }

        } else {
          Logger.log("No more records returned from Zoho. Stopping fetch.");
          moreRecordsToFetch = false;
        }

      } else {
        Logger.log(`ERROR: Zoho Creator API call failed.`);
        Logger.log(`HTTP Status: ${responseCode}.`);
        Logger.log(`Response Body: ${responseBody}.`);
        moreRecordsToFetch = false;
      }

    } catch (e) {
      Logger.log(`FATAL EXCEPTION: An unexpected error occurred during Zoho Creator API fetch: ${e.message}.`);
      Logger.log(`Stack Trace: ${e.stack}`);
      moreRecordsToFetch = false;
    }
  }

  if (newRowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRowsToAppend.length, headers.length).setValues(newRowsToAppend);
    Logger.log(`SUCCESS: Appended ${newRowsToAppend.length} new records to Google Sheet.`);
  } else {
    Logger.log("INFO: No new records to append to Google Sheet during this synchronization run.");
  }
}

