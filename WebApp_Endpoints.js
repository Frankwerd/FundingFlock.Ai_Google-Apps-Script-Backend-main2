/**
 * @file Handles authenticated GET and POST requests for the CareerSuite.AI Web App.
 * @version 2.1
 */

/**
 * Main entry point for all GET requests to the Web App.
 * It routes requests based on an 'action' parameter after validating authentication.
 * @param {GoogleAppsScript.Events.DoGet} e The event parameter from the GET request.
 * @returns {GoogleAppsScript.Content.TextOutput | GoogleAppsScript.HTML.HtmlOutput} A JSON or HTML response.
 */
function doGet(e) {
  const FUNC_NAME = "WebApp_doGet";
  try {
    const action = e.parameter.action;
    Logger.log(`[${FUNC_NAME}] Received GET request. Action: "${action}". User: ${Session.getEffectiveUser().getEmail()}`);

    // Route to the correct function based on the 'action' parameter.
    // If an 'action' is provided, we treat it as an API call from our extension.
    if (action) {
      switch (action) {
        case 'getOrCreateSheet':
          return doGet_getOrCreateSheet(e); // Handles sheet creation/retrieval.

        case 'getWeeklyApplicationData':
          return doGet_WeeklyApplicationData(e); // Handles chart data.

        case 'getApiKeyForScript':
          return doGet_getApiKeyForScript(e); // Handles key sync for copied sheets.

        default:
          // If the action is unknown, return a clear JSON error.
          return createJsonResponse({ success: false, error: `Unknown action: ${action}` });
      }
    } else {
      // If no action is specified, it's likely a user visiting the URL directly
      // or the initial OAuth redirect. Show a user-friendly landing page.
      return createAuthLandingPage();
    }

  } catch (error) {
    Logger.log(`[${FUNC_NAME}] CRITICAL Unhandled Error in doGet: ${error.toString()}\nStack: ${error.stack}`);
    return createJsonResponse({ success: false, error: `An unexpected server error occurred: ${error.message}` });
  }
}

/**
 * Handles POST requests to the web app, primarily for saving data like the API key.
 * @param {GoogleAppsScript.Events.DoPost} e The event parameter from the POST request.
 * @returns {GoogleAppsScript.Content.TextOutput} A JSON response.
 */
function doPost(e) {
  const FUNC_NAME = "WebApp_doPost";
  try {
    // --- User & Action Identification ---
    const userEmail = Session.getEffectiveUser().getEmail();
    Logger.log(`[${FUNC_NAME}] Received POST request from user: ${userEmail}.`);

    const action = e && e.parameter ? e.parameter.action : null;

    // === ACTION: setApiKey ===
    if (action === 'setApiKey') {
      Logger.log(`[${FUNC_NAME}] Routing to 'setApiKey' action.`);
      
      const postData = JSON.parse(e.postData.contents);
      const apiKey = postData.apiKey;

      if (!apiKey || typeof apiKey !== 'string' || apiKey.length < 35) {
        Logger.log(`[${FUNC_NAME} ERROR] 'setApiKey' failed: Invalid API key provided.`);
        return createJsonResponse({
          status: 'error',
          message: 'Invalid or missing API key provided in the request.'
        });
      }

      // GEMINI_API_KEY_PROPERTY is the constant from Config.gs
      PropertiesService.getUserProperties().setProperty(GEMINI_API_KEY_PROPERTY, apiKey);

      Logger.log(`[${FUNC_NAME} SUCCESS] Saved Gemini API key to UserProperties for user ${userEmail}.`);
      return createJsonResponse({
        status: 'success',
        message: 'API key was successfully saved to the backend.'
      });
    }

    // Fallback for an unknown POST action.
    Logger.log(`[${FUNC_NAME} WARN] An unknown POST action was requested: "${action}".`);
    return createJsonResponse({
      status: 'error',
      message: 'Unknown or unsupported POST action requested.'
    });

  } catch (error) {
    // Global Error Handling for any unexpected errors.
    Logger.log(`[${FUNC_NAME} CRITICAL ERROR] Error in doPost: ${error.toString()}\nStack: ${error.stack}`);
    return createJsonResponse({
      status: 'error',
      message: `Failed to complete POST action due to a server-side error: ${error.message}`
    });
  }
}


/**
 * Handles the logic for getting an existing sheet or creating a new one for the user.
 * This is the primary endpoint for the extension's "Manage Job Tracker" button.
 * @param {GoogleAppsScript.Events.DoGet} e The event parameter from the GET request.
 * @returns {GoogleAppsScript.Content.TextOutput} A JSON response.
 */
function doGet_getOrCreateSheet(e) {
  const FUNC_NAME = "doGet_getOrCreateSheet";
  const userEmail = Session.getEffectiveUser().getEmail();
  const userProps = PropertiesService.getUserProperties();
  const existingSheetId = userProps.getProperty('userMjmSheetId');
  
  // 1. Check if a valid Sheet ID is already stored for the user.
  if (existingSheetId) {
    try {
      const existingSheet = SpreadsheetApp.openById(existingSheetId);
      Logger.log(`[${FUNC_NAME}] Found existing, valid sheet for ${userEmail}: ID=${existingSheetId}`);
      return createJsonResponse({
        status: "success",
        message: "Sheet already exists.",
        sheetId: existingSheetId,
        sheetUrl: existingSheet.getUrl()
      });
    } catch (openErr) {
      Logger.log(`[${FUNC_NAME}] Stored sheet ID ${existingSheetId} was invalid or inaccessible. Clearing property and creating a new sheet. Error: ${openErr.message}`);
      userProps.deleteProperty('userMjmSheetId');
    }
  }

  // 2. If no valid ID, create a new sheet from the template.
  Logger.log(`[${FUNC_NAME}] No valid sheet found for ${userEmail}. Creating from template...`);
  const templateId = TEMPLATE_SHEET_ID; // From Config.gs
  if (!templateId || templateId.length < 20) {
    return createJsonResponse({ status: 'error', message: 'Server configuration error: Master Template Sheet ID is invalid.' });
  }

  const templateFile = DriveApp.getFileById(templateId);
  const newFileName = `CareerSuite.AI Data`; // From Config.gs -> TARGET_SPREADSHEET_FILENAME
  const newSheetFile = templateFile.makeCopy(newFileName);
  const newSheetId = newSheetFile.getId();
  
  // 3. Run the full project setup on the newly created sheet.
  const newSheet = SpreadsheetApp.openById(newSheetId);
  // const setupResult = runFullProjectInitialSetup(newSheet); // From Main.gs // DEFERRED TO ON_OPEN

  // if (!setupResult.success) { // DEFERRED TO ON_OPEN
  //   Logger.log(`[${FUNC_NAME}] CRITICAL: Initial setup failed on new sheet ${newSheetId}.`);
  //   // Consider deleting the failed sheet to avoid clutter: DriveApp.getFileById(newSheetId).setTrashed(true);
  //   return createJsonResponse({
  //     status: "error",
  //     message: `Failed to initialize the new sheet. Please check script logs for details. Error: ${setupResult.message}`
  //   });
  // }

  // 4. Save the new, successfully initialized sheet ID to user properties.
  userProps.setProperty('userMjmSheetId', newSheetId);
  Logger.log(`[${FUNC_NAME}] New sheet created and initialized for ${userEmail}. ID: ${newSheetId}`);

  return createJsonResponse({
    status: "success",
    message: "Your CareerSuite.AI Data sheet was created and set up successfully!",
    sheetId: newSheetId,
    sheetUrl: newSheet.getUrl()
  });
}

/**
 * Creates a generic HTML response for the OAuth landing page.
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output for the landing page.
 */
function createAuthLandingPage() {
    const htmlOutput = `
      <!DOCTYPE html><html><head><title>CareerSuite.AI Authorization</title>
      <style>body { font-family: sans-serif; margin: 20px; background-color: #f0f4f8; color: #333; text-align: center; } .container { background-color: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); display: inline-block; } h1 { color: #33658A; }</style></head><body><div class="container">
        <h1>CareerSuite.AI</h1><p>Authorization successful!</p><p>You can now close this tab and return to the extension.</p>
      </div></body></html>`;
    return HtmlService.createHtmlOutput(htmlOutput)
      .setTitle("CareerSuite.AI Authorization")
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

/**
 * Handles GET requests for aggregated weekly application data.
 * @param {GoogleAppsScript.Events.DoGet} e The event parameter from the GET request.
 * @returns {GoogleAppsScript.Content.TextOutput} A JSON response.
 */
function doGet_WeeklyApplicationData(e) {
  const FUNC_NAME = "doGet_WeeklyApplicationData";
  try {
    const userMjmSheetId = PropertiesService.getUserProperties().getProperty('userMjmSheetId');
    if (!userMjmSheetId) {
      return createJsonResponse({ 
          success: false, 
          error: "CareerSuite.AI Sheet ID not found. Please complete setup via the extension." 
      });
    }

    let ss;
    try {
        ss = SpreadsheetApp.openById(userMjmSheetId);
    } catch (sheetOpenErr) {
        Logger.log(`[${FUNC_NAME} ERROR] Error opening sheet ID ${userMjmSheetId}: ${sheetOpenErr.message}`);
        PropertiesService.getUserProperties().deleteProperty('userMjmSheetId');
        return createJsonResponse({ 
            success: false, 
            error: `Your saved Sheet ID is no longer accessible. Please re-link your sheet.`
        });
    }
    
    const helperSheet = ss.getSheetByName(HELPER_SHEET_NAME);
    if (!helperSheet) {
      return createJsonResponse({ 
          success: false, 
          error: `Helper data sheet ("${HELPER_SHEET_NAME}") not found. Please run 'Update Dashboard Metrics' from the tools menu in your sheet.`
      });
    }

    const headersRange = helperSheet.getRange("D1:E1").getDisplayValues();
    if (headersRange[0][0] !== "Week Starting" || headersRange[0][1] !== "Applications") {
        return createJsonResponse({ 
            success: false, 
            error: "Helper data format for weekly applications is incorrect."
        });
    }
    
    const lastDataRowInColD = helperSheet.getRange("D1:D").getValues().filter(String).length;
    let weeklyData = [];

    if (lastDataRowInColD > 1) {
        const maxWeeksToShow = 12; // Show up to 12 weeks of data
        const startRowForFetch = Math.max(2, lastDataRowInColD - maxWeeksToShow + 1);
        const numRowsToFetchActual = lastDataRowInColD - startRowForFetch + 1;
        
        if (numRowsToFetchActual > 0) {
            const rangeDataValues = helperSheet.getRange(startRowForFetch, 4, numRowsToFetchActual, 2).getDisplayValues();
            rangeDataValues.forEach(row => {
                if (row[0] && row[1]) {
                     weeklyData.push({ weekStarting: row[0], applications: row[1] });
                }
            });
        }
    }
    
    Logger.log(`[${FUNC_NAME} INFO] Fetched ${weeklyData.length} weekly data points.`);
    return createJsonResponse({ success: true, data: weeklyData });

  } catch (error) {
    Logger.log(`[${FUNC_NAME} ERROR] Error: ${error.toString()}\nStack: ${error.stack}`);
    return createJsonResponse({ 
        success: false, 
        error: `Error fetching weekly application data: ${error.toString()}`
    });
  }
}

/**
 * Handles GET requests to retrieve the stored Gemini API key for the authenticated user.
 * @param {GoogleAppsScript.Events.DoGet} e The event parameter from the GET request.
 * @returns {GoogleAppsScript.Content.TextOutput} A JSON response containing the API key or an error.
 */
function doGet_getApiKeyForScript(e) {
  const FUNC_NAME = "doGet_getApiKeyForScript";
  try {
    const userEmail = Session.getEffectiveUser().getEmail(); 
    Logger.log(`[${FUNC_NAME}] Request from user: ${userEmail} to retrieve API key.`);
    
    // GEMINI_API_KEY_PROPERTY should be globally available if defined in Config.gs
    const apiKey = PropertiesService.getUserProperties().getProperty(GEMINI_API_KEY_PROPERTY); 

    if (apiKey) {
      Logger.log(`[${FUNC_NAME}] API Key retrieved successfully for ${userEmail}.`);
      return createJsonResponse({ success: true, apiKey: apiKey });
    } else {
      Logger.log(`[${FUNC_NAME}] API Key not found for ${userEmail}.`);
      return createJsonResponse({ success: false, error: "API Key not found. Please ensure you have saved your API key in the extension settings." });
    }
  } catch (error) {
    Logger.log(`[${FUNC_NAME} ERROR] Error retrieving API key: ${error.toString()}\nStack: ${error.stack}`);
    return createJsonResponse({ success: false, error: `Server error while retrieving API key: ${error.message}` });
  }
}
