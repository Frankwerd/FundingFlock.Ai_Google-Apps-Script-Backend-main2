/**
 * @file Handles authenticated GET and POST requests for the FundingFlock.AI Web App.
 * @version 3.0 (Adapted for FundingFlock communication)
 */

/**
 * Main entry point for all GET requests to the Web App.
 * This function's primary role is to serve the landing page after a user completes the OAuth flow.
 * @param {GoogleAppsScript.Events.DoGet} e The event parameter from the GET request.
 * @returns {GoogleAppsScript.Content.TextOutput | GoogleAppsScript.HTML.HtmlOutput} A JSON or HTML response.
 */
function doGet(e) {
  const FUNC_NAME = "WebApp_doGet";
  try {
    const action = e.parameter.action;
    Logger.log(`[${FUNC_NAME}] Received GET request. Action: "${action}". User: ${Session.getEffectiveUser().getEmail()}`);

    // If an action parameter exists, it's an API call we don't support via GET in this version.
    if (action) {
      return createJsonResponse({ success: false, error: `Unknown GET action: ${action}` });
    } else {
      // If no action is specified, it's the OAuth redirect. Show the user-friendly landing page.
      return createAuthLandingPage();
    }
  } catch (error) {
    Logger.log(`[${FUNC_NAME}] CRITICAL Unhandled Error in doGet: ${error.toString()}\nStack: ${error.stack}`);
    return createJsonResponse({ success: false, error: `An unexpected server error occurred: ${error.message}` });
  }
}

/**
 * Handles POST requests from the extension, which is the primary method of communication.
 * @param {GoogleAppsScript.Events.DoPost} e The event parameter.
 * @returns {GoogleAppsScript.Content.TextOutput} A JSON response.
 */
function doPost(e) {
  const FUNC_NAME = "WebApp_doPost";
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    Logger.log(`[${FUNC_NAME}] Received POST from ${userEmail}.`);

    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;

    switch (action) {
      case 'createTrackerSheet':
        Logger.log(`[${FUNC_NAME}] Action: createTrackerSheet`);
        return handleCreateSheet(postData);

      case 'setApiKey':
        Logger.log(`[${FUNC_NAME}] Action: setApiKey`);
        const apiKey = postData.apiKey;
        if (!apiKey || typeof apiKey !== 'string' || apiKey.length < 35) {
          throw new Error('Invalid or missing API key.');
        }
        // GEMINI_API_KEY_PROPERTY is a global constant from Config.js
        PropertiesService.getUserProperties().setProperty(GEMINI_API_KEY_PROPERTY, apiKey);
        Logger.log(`[${FUNC_NAME} SUCCESS] Saved Gemini API key.`);
        return createJsonResponse({ status: 'success', message: 'API key was successfully saved.' });

      default:
        throw new Error(`Unknown POST action: "${action}".`);
    }
  } catch (error) {
    Logger.log(`[${FUNC_NAME} CRITICAL ERROR] ${error.toString()}`);
    return createJsonResponse({ status: 'error', message: `Server error: ${error.message}` });
  }
}

/**
 * Creates a copy of the template sheet for the user. Does NOT run the full setup.
 * @param {object} postData The parsed data from the POST request.
 * @returns {GoogleAppsScript.Content.TextOutput} A JSON response.
 */
function handleCreateSheet(postData) {
  const FUNC_NAME = "handleCreateSheet";
  const userEmail = Session.getEffectiveUser().getEmail();
  const userProps = PropertiesService.getUserProperties();
  const userSheetIdPropKey = 'userFundingFlockSheetId'; // Use a unique key for FF

  const existingSheetId = userProps.getProperty(userSheetIdPropKey);

  // 1. Check if a valid Sheet ID is already stored for this user.
  if (existingSheetId) {
    try {
      const existingSheet = SpreadsheetApp.openById(existingSheetId);
      Logger.log(`[${FUNC_NAME}] Found existing, valid sheet for ${userEmail}.`);
      return createJsonResponse({
        status: 'success',
        message: 'Sheet already exists.',
        sheetId: existingSheetId,
        sheetUrl: existingSheet.getUrl()
      });
    } catch (e) {
      // The stored ID was invalid (e.g., user deleted the sheet).
      Logger.log(`[${FUNC_NAME}] Stored sheet ID was invalid. Clearing property and creating new sheet.`);
      userProps.deleteProperty(userSheetIdPropKey);
    }
  }

  // 2. If no valid ID, create a new sheet by copying the template.
  Logger.log(`[${FUNC_NAME}] Creating new sheet from template ID: ${TEMPLATE_SHEET_ID}`);
  const templateFile = DriveApp.getFileById(TEMPLATE_SHEET_ID);
  const newFileName = `${TARGET_SPREADSHEET_FILENAME} - ${userEmail}`;
  const newSheetFile = templateFile.makeCopy(newFileName);
  const newSheetId = newSheetFile.getId();
  
  // 3. Store the new ID for the user.
  userProps.setProperty(userSheetIdPropKey, newSheetId);
  Logger.log(`[${FUNC_NAME}] Successfully created new sheet. ID: ${newSheetId}`);
  
  // 4. Return success message, prompting user for the next step.
  return createJsonResponse({
    status: 'success',
    message: 'Sheet created. Please open it and run "Finalize Project Setup" from the menu.',
    sheetId: newSheetId,
    sheetUrl: newSheetFile.getUrl()
  });
}

/**
 * Creates a generic HTML response for the OAuth landing page.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output for the landing page.
 */
function createAuthLandingPage() {
    const htmlOutput = `
      <!DOCTYPE html><html><head><title>FundingFlock.AI Authorization</title>
      <style>body{font-family:sans-serif;margin:20px;background-color:#FFF0E5;color:#333;text-align:center;}.container{background-color:#fff;padding:30px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);display:inline-block;}h1{color:#96616B;}</style></head><body><div class="container">
      <h1>FundingFlock.AI</h1><p>Authorization successful!</p><p>You can now close this tab and return to the extension.</p>
      </div></body></html>`;
    return HtmlService.createHtmlOutput(htmlOutput)
      .setTitle("FundingFlock.AI Authorization")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Creates a standardized JSON response object.
 * @param {object} payload The JSON payload to send.
 * @returns {GoogleAppsScript.Content.TextOutput} The JSON response object.
 */
function createJsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
