/**
 * @file Contains administrative utility functions for project setup and configuration,
 * such as managing API keys stored in UserProperties.
 */

/**
 * Performs a manual test to check if the template sheet can be accessed and copied.
 * It uses the TEMPLATE_SHEET_ID from Config.js.
 * Displays success or failure messages to the user via the Spreadsheet UI.
 */
function manualTestTemplateAccess() {
  const id = TEMPLATE_SHEET_ID; // From Config.gs
  if (!id || id === "YOUR_ACTUAL_TEMPLATE_MJM_SHEET_ID_HERE") {
    SpreadsheetApp.getUi().alert("Template ID Not Set", "The TEMPLATE_SHEET_ID in Config.js is not set.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  try {
    const file = DriveApp.getFileById(id);
    const copy = file.makeCopy("MANUAL TEST COPY - " + new Date().toLocaleTimeString());
    SpreadsheetApp.getUi().alert("Test Successful", `A test copy named "${copy.getName()}" was successfully created in your Google Drive. You can delete it now.`);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Test Failed", "Error accessing or copying template: " + e.toString());
  }
}
/**
 * Provides a user interface prompt to set the shared Gemini API Key in UserProperties.
 * This function handles both Spreadsheet and non-Spreadsheet environments (fallback to Browser.inputBox).
 * It validates the key's length and format, and confirms overwrites with the user.
 * The property name under which the key is stored is defined by GEMINI_API_KEY_PROPERTY in Config.js.
 */
function setSharedGeminiApiKey_UI() {
  const MIN_KEY_LENGTH = 35; // Minimum expected length for a Gemini API key
  const propertyName = GEMINI_API_KEY_PROPERTY; // From Config.gs (e.g., "CAREERSUITE_GEMINI_API_KEY")
  const userProps = PropertiesService.getUserProperties();
  const currentKey = userProps.getProperty(propertyName);

  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    Logger.log('[ADMIN_UTILS WARN] SpreadsheetApp.getUi() is not available, will use Browser.inputBox/msgBox. Error: ' + e.message);
    ui = null; // Ensure ui is null if it fails
  }

  if (ui) { // --- Use SpreadsheetApp.getUi() ---
    try {
      const initialPromptMessage = `Enter the CareerSuite.AI Gemini API Key.\nThis will be stored in UserProperties under: "${propertyName}".\n(Typically ${MIN_KEY_LENGTH}+ characters).\n${currentKey ? 'An existing key is set and will be overwritten if you enter a new one.' : 'No key is currently set.'}`;
      const response = ui.prompt('Set Gemini API Key', initialPromptMessage, ui.ButtonSet.OK_CANCEL);

      if (response.getSelectedButton() == ui.Button.OK) {
        const apiKey = response.getResponseText().trim();

        if (apiKey && apiKey.length >= MIN_KEY_LENGTH && /^[a-zA-Z0-9_~-]+$/.test(apiKey)) {
          if (currentKey && currentKey !== apiKey) {
            const confirmOverwrite = ui.alert(
              'Confirm Overwrite',
              `An API key already exists. Overwrite it with the new key?\n\nOld (masked): ${currentKey.substring(0,4)}...${currentKey.substring(currentKey.length-4)}\nNew (masked): ${apiKey.substring(0,4)}...${apiKey.substring(apiKey.length-4)}`,
              ui.ButtonSet.YES_NO
            );
            if (confirmOverwrite !== ui.Button.YES) {
              ui.alert('Operation Cancelled', 'Existing API key was NOT overwritten.', ui.ButtonSet.OK);
              return;
            }
          } else if (currentKey && currentKey === apiKey) {
             ui.alert('No Change', 'Entered API key is the same as the current key. No changes made.', ui.ButtonSet.OK);
             return;
          }
          userProps.setProperty(propertyName, apiKey);
          ui.alert('API Key Saved', `Gemini API Key saved successfully for CareerSuite.AI under property: "${propertyName}".`);
        } else if (apiKey) {
          let N = apiKey.length < MIN_KEY_LENGTH ? `Key too short (min ${MIN_KEY_LENGTH} chars). ` : '';
          if(!/^[a-zA-Z0-9_~-]+$/.test(apiKey)) N += `Key contains invalid characters.`;
          ui.alert('API Key Not Saved', `Invalid key. ${N}Please check and retry.`, ui.ButtonSet.OK);
        } else {
          ui.alert('API Key Not Saved', 'No API key was entered.', ui.ButtonSet.OK);
        }
      } else {
        ui.alert('API Key Setup Cancelled', 'API key setup cancelled.', ui.ButtonSet.OK);
      }
    } catch (uiError) {
      Logger.log('[ADMIN_UTILS ERROR] Error during SpreadsheetApp.getUi() interaction: ' + uiError.message);
      // Fallback or just log, as Browser.inputBox might not be ideal if this specific path fails.
    }
  } else { // --- Fallback to Browser.inputBox/msgBox ---
    Logger.log('[ADMIN_UTILS INFO] Attempting Browser.inputBox for API Key input.');
    try {
      const currentKeyInfo = currentKey ? "(Existing key will be overwritten)" : "(No key set)";
      const apiKeyInput = Browser.inputBox(`Set CareerSuite.AI Gemini API Key`, `Enter Gemini API Key for property: ${propertyName}.\n(Min ${MIN_KEY_LENGTH}+ chars).\n${currentKeyInfo}`, Browser.Buttons.OK_CANCEL);

      if (apiKeyInput !== 'cancel' && apiKeyInput !== null) {
        const trimmedApiKey = apiKeyInput.trim();
        if (trimmedApiKey && trimmedApiKey.length >= MIN_KEY_LENGTH && /^[a-zA-Z0-9_~-]+$/.test(trimmedApiKey)) {
          if (currentKey && currentKey !== trimmedApiKey) {
             const confirmOverwriteResponse = Browser.msgBox('Confirm Overwrite', `API key exists. Overwrite?\nOld (masked): ${currentKey.substring(0,4)}... \nNew (masked): ${trimmedApiKey.substring(0,4)}...\n(OK to overwrite, Cancel to keep old)`, Browser.Buttons.OK_CANCEL);
            if (confirmOverwriteResponse !== 'ok') { Browser.msgBox('Operation Cancelled', 'Existing API key NOT overwritten.', Browser.Buttons.OK); return; }
          } else if (currentKey && currentKey === trimmedApiKey) { Browser.msgBox('No Change', 'Entered key is same as current. No changes.', Browser.Buttons.OK); return; }
          userProps.setProperty(propertyName, trimmedApiKey);
          Browser.msgBox('API Key Saved', `Gemini API Key saved successfully as: "${propertyName}".`, Browser.Buttons.OK);
        } else if (trimmedApiKey) {
          let N = trimmedApiKey.length < MIN_KEY_LENGTH ? `Key too short. ` : '';
          if(!/^[a-zA-Z0-9_~-]+$/.test(trimmedApiKey)) N += `Invalid characters.`;
          Browser.msgBox('API Key Not Saved', `Invalid key. ${N}Check and retry.`, Browser.Buttons.OK);
        } else { Browser.msgBox('API Key Not Saved', 'No API key entered.', Browser.Buttons.OK); }
      } else { Browser.msgBox('API Key Setup Cancelled', 'API key setup cancelled.', Browser.Buttons.OK); }
    } catch (e2) {
      Logger.log('[ADMIN_UTILS ERROR] Browser.inputBox/msgBox interaction failed: ' + e2.message);
    }
  }
}


/**
 * Displays all UserProperties set for this script project to the logs.
 * Sensitive values like API keys are partially masked for security.
 * It also displays a confirmation message in the UI.
 */
function showAllUserProperties() {
  const userProps = PropertiesService.getUserProperties().getProperties();
  let logOutput = "Current UserProperties for this CareerSuite.AI script project:\n";
  if (Object.keys(userProps).length === 0) {
    logOutput += "  (No UserProperties are currently set for this project)\n";
  } else {
    for (const key in userProps) {
      let valueToLog = userProps[key];
      if (key.toLowerCase().includes('api') || key.toLowerCase().includes('key') || key.toLowerCase().includes('secret')) {
        if (valueToLog && typeof valueToLog === 'string' && valueToLog.length > 8) {
          valueToLog = valueToLog.substring(0, 4) + "..." + valueToLog.substring(valueToLog.length - 4);
        } else if (valueToLog && typeof valueToLog === 'string') {
            valueToLog = "**** (value too short to fully mask)";
        }
      }
      logOutput += `  ${key}: ${valueToLog}\n`;
    }
  }
  Logger.log(logOutput);
  const alertMsg = "UserProperties logged. Check Apps Script logs (View > Logs/Executions). Sensitive values are partially masked.";
  try { SpreadsheetApp.getUi().alert("User Properties Logged", alertMsg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch(e) { try { Browser.msgBox("User Properties Logged", alertMsg, Browser.Buttons.OK); } catch(e2) {} }
}
