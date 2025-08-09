/**
 * @file Main script file orchestrating setup, email processing, and UI for the grant proposal tracker.
 * @author Francis John LiButti (Originals), AI Integration & Refinements by Assistant
 * @version 10 (FundingFlock.AI Refactor)
 */

/**
 * Checks if the critical configuration variables are set correctly.
 * @returns {boolean} True if the configuration is valid, false otherwise.
 */
function checkConfig() {
    const FUNC_NAME = "checkConfig";
    // MASTER_WEB_APP_URL removed as it's not in the new config and part of a separate system.
    const criticalVars = {
        TEMPLATE_SHEET_ID,
        MASTER_SCRIPT_ID,
        GEMINI_API_KEY_PROPERTY,
        GEMINI_API_ENDPOINT_TEXT_ONLY
    };
    let allClear = true;
    for (const [varName, varValue] of Object.entries(criticalVars)) {
        if (typeof varValue === 'undefined' || varValue === null || varValue === "" || varValue.includes("REPLACE") || varValue.includes("YOUR_")) {
            Logger.log(`[${FUNC_NAME} CRITICAL] Configuration variable ${varName} is not set correctly. Value: "${varValue}"`);
            allClear = false;
        }
    }
    return allClear;
}

/**
 * Runs the complete initial setup for all modules.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [passedSpreadsheet]
 * @returns {{success: boolean, message: string, detailedMessages: string[], sheetId: string|null, sheetUrl: string|null}}
 */
function runFullProjectInitialSetup(passedSpreadsheet) {
    const RUNDATE = new Date().toISOString();
    const FUNC_NAME = "runFullProjectInitialSetup";
    Logger.log(`==== ${FUNC_NAME}: STARTING (FundingFlock.AI v2.0 - ${RUNDATE}) ====`);
    if (!checkConfig()) {
        Logger.log(`[${FUNC_NAME} CRITICAL] Configuration check failed. Aborting setup.`);
        return { success: false, message: "Critical configuration is missing.", detailedMessages: ["Critical configuration is missing. Please check the logs."], sheetId: null, sheetUrl: null };
    }
    let overallSuccess = true;
    let setupMessages = [];
    let activeSS;

    if (passedSpreadsheet && typeof passedSpreadsheet.getId === 'function') {
        activeSS = passedSpreadsheet;
    } else {
        activeSS = SpreadsheetApp.getActiveSpreadsheet();
        if (!activeSS) {
            const { spreadsheet: foundOrCreatedSS } = getOrCreateSpreadsheetAndSheet();
            activeSS = foundOrCreatedSS;
        }
    }

    if (!activeSS) {
        const errorMsg = `CRITICAL [${FUNC_NAME}]: No valid spreadsheet could be determined. Setup aborted.`;
        Logger.log(errorMsg);
        return { success: false, message: errorMsg, detailedMessages: [errorMsg], sheetId: null, sheetUrl: null };
    }

    PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_KEY, activeSS.getId());

    if (typeof TEMPLATE_SHEET_ID !== 'undefined' && TEMPLATE_SHEET_ID !== "" && activeSS.getId() === TEMPLATE_SHEET_ID) {
        const templateMsg = `[${FUNC_NAME} INFO] Target spreadsheet is the TEMPLATE. Setup SKIPPED.`;
        Logger.log(templateMsg);
        if (!passedSpreadsheet) {
            try {
                SpreadsheetApp.getUi().alert('Template Sheet', 'Initial setup is not meant to be run on the template sheet itself. Please make a copy first.', SpreadsheetApp.getUi().ButtonSet.OK);
            } catch (e) { /* UI not available */ }
        }
        return { success: true, message: "Setup skipped: Target is template.", detailedMessages: [templateMsg], sheetId: activeSS.getId(), sheetUrl: activeSS.getUrl() };
    }

    // --- Dashboard & Helper sheets are created FIRST ---
    const dashboardSheet = getOrCreateDashboardSheet(activeSS);
    const helperSheet = getOrCreateHelperSheet(activeSS);
    if (!dashboardSheet || !helperSheet) {
        const errorMsg = "Dashboard/Helper sheet creation failed. Aborting.";
        Logger.log(`[${FUNC_NAME} CRITICAL] ${errorMsg}`);
        setupMessages.push(errorMsg);
        return { success: false, message: errorMsg, detailedMessages: setupMessages, sheetId: activeSS.getId(), sheetUrl: activeSS.getUrl()};
    }

const modules = [
    { name: "Proposal Tracker", setupFunc: initialSetup_LabelsAndSheet }
];

    for (const module of modules) {
        if (typeof module.setupFunc === "function") {
            Logger.log(`\n[${FUNC_NAME} INFO] --- Starting ${module.name} Setup ---`);
            try {
                const result = module.setupFunc(activeSS);
                if (result.messages) setupMessages.push(...result.messages.map(m => `${module.name}: ${m}`));
                if (!result.success) {
                    overallSuccess = false;
                    Logger.log(`[${FUNC_NAME} ERROR] ${module.name} FAILED.`);
                } else {
                    Logger.log(`[${FUNC_NAME} INFO] ${module.name} Success.`);
                }
            } catch (e) {
                Logger.log(`[${FUNC_NAME} CRITICAL ERROR] ${module.name} Exception: ${e.toString()}\n${e.stack}`);
                setupMessages.push(`${module.name}: CRITICAL EXCEPTION - ${e.message}`);
                overallSuccess = false;
            }
        }
    }

    // --- Dashboard formatting and chart updates happen AFTER modules have run (which adds dummy data) ---
    if (overallSuccess) {
        Logger.log(`\n[${FUNC_NAME} INFO] --- Finalizing Dashboard Setup ---`);
        try {
            formatDashboardSheet(dashboardSheet);
            setupHelperSheetFormulas(helperSheet);
            updateDashboardMetrics(dashboardSheet, helperSheet);
            setupMessages.push(`Dashboard Module: Sheet setup and charts configured.`);
        } catch (e) {
             Logger.log(`[${FUNC_NAME} CRITICAL ERROR] Dashboard Finalization Exception: ${e.toString()}\n${e.stack}`);
             setupMessages.push(`Dashboard Module: CRITICAL EXCEPTION - ${e.message}`);
             overallSuccess = false;
        }
    }

    // --- Final cleanup and UI ordering ---
    if (overallSuccess) {
        Logger.log(`[${FUNC_NAME} INFO] Applying final tab order and cleaning up...`);
        try {
            // Clear dummy data from the Proposals sheet now that the dashboard is primed
            const proposalSheet = activeSS.getSheetByName(PROPOSAL_TRACKER_SHEET_TAB_NAME);
            if (proposalSheet && proposalSheet.getLastRow() > 1) {
                proposalSheet.getRange(2, 1, proposalSheet.getLastRow() - 1, proposalSheet.getLastColumn()).clearContent();
                Logger.log(`[${FUNC_NAME} INFO] Cleared dummy data from Proposals sheet.`);
            }

            const tabOrder = [DASHBOARD_TAB_NAME, PROPOSAL_TRACKER_SHEET_TAB_NAME, OPPORTUNITIES_SHEET_TAB_NAME, HELPER_SHEET_NAME];
            tabOrder.forEach((sheetName, index) => {
                const sheetToMove = activeSS.getSheetByName(sheetName);
                if (sheetToMove) {
                    activeSS.setActiveSheet(sheetToMove);
                    activeSS.moveActiveSheet(index + 1);
                }
            });

            if (helperSheet && !helperSheet.isSheetHidden()) {
                helperSheet.hideSheet();
            }
            setupMessages.push("Branding: Tab order & helper data visibility verified.");
        } catch (e) {
            Logger.log(`[${FUNC_NAME} WARN] Error during final cleanup/ordering: ${e.message}`);
        }
    }

    const finalStatusMessage = `FundingFlock.AI Full Setup ${overallSuccess ? "completed" : "had issues"}.`;
    Logger.log(`\n==== ${FUNC_NAME} SUMMARY (SS ID: ${activeSS.getId()}) ====`);
    setupMessages.forEach(msg => Logger.log(`  - ${msg}`));
    Logger.log(`Overall Status: ${overallSuccess ? "SUCCESSFUL" : "ISSUES ENCOUNTERED"}`);

    if (!passedSpreadsheet) {
        try {
            const title = `FundingFlock.AI Setup ${overallSuccess ? "Complete" : "Issues"}`;
            const message = `Setup for "${activeSS.getName()}" ${overallSuccess ? "finished" : "had issues"}.\n\nSummary:\n- ${setupMessages.join('\n- ')}`;
            SpreadsheetApp.getUi().alert(title, message.substring(0, 1000), SpreadsheetApp.getUi().ButtonSet.OK);
        } catch (e) { /* UI not available */ }
    }

    return { success: overallSuccess, message: finalStatusMessage, detailedMessages: setupMessages, sheetId: activeSS.getId(), sheetUrl: activeSS.getUrl() };
}


/**
 * Sets up the core Grant Proposal Tracker module.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} activeSS The spreadsheet object.
 * @returns {{success: boolean, messages: string[]}}
 */
function initialSetup_LabelsAndSheet(activeSS) {
    const trackerConfig = {
        activeSS: activeSS,
        moduleName: "Proposal Tracker",
        sheetTabName: PROPOSAL_TRACKER_SHEET_TAB_NAME,
        sheetHeaders: PROPOSAL_TRACKER_SHEET_HEADERS,
        // Create a generic width mapping based on the number of headers
        columnWidths: PROPOSAL_TRACKER_SHEET_HEADERS.map((h, i) => ({ col: i + 1, width: 150 })),
        bandingTheme: SpreadsheetApp.BandingTheme.BLUE,
        tabColor: BRAND_COLORS.LAPIS_LAZULI,
        gmailLabelParent: MASTER_GMAIL_LABEL_PARENT,
        gmailLabelToProcess: TRACKER_GMAIL_LABEL_TO_PROCESS,
        gmailLabelProcessed: TRACKER_GMAIL_LABEL_PROCESSED,
        gmailLabelManualReview: TRACKER_GMAIL_LABEL_MANUAL_REVIEW,
        gmailFilterQuery: TRACKER_GMAIL_FILTER_QUERY,
        triggerFunctionName: 'processEmails_triggerHandler',
        triggerIntervalHours: 1,
        staleRejectFunctionName: 'markStale_triggerHandler'
    };
    // _setupModule is a generic helper and does not need to be changed.
    return _setupModule(trackerConfig);
}

// START SNIPPET 5A: Replace _processingEngine in Main.js
function _processingEngine(config, ss, scriptProperties) {
    const FUNC_NAME = "_processingEngine";
    const SCRIPT_START_TIME = new Date();
    Logger.log(`\n==== ${FUNC_NAME}: STARTING (${SCRIPT_START_TIME.toLocaleString()}) - ${config.moduleName} ====`);
    const geminiApiKey = scriptProperties.getProperty(GEMINI_API_KEY_PROPERTY);
    if (!geminiApiKey) { /* ... error logging ... */ return; }
    const dataSheet = ss.getSheetByName(config.sheetTabName);
    if (!dataSheet) { /* ... error logging ... */ return; }

    const procLbl = GmailApp.getUserLabelByName(config.gmailLabelToProcess);
    const processedLblObj = GmailApp.getUserLabelByName(config.gmailLabelProcessed);
    const manualLblObj = config.gmailLabelManualReview ? GmailApp.getUserLabelByName(config.gmailLabelManualReview) : processedLblObj;
    if (!procLbl || !processedLblObj || !manualLblObj) { /* ... error logging ... */ return; }

    // --- THIS CACHING LOGIC IS NOW CRITICAL ---
    const allSheetData = dataSheet.getDataRange().getValues();
    const funderIndex = new Map();
    for (let i = 1; i < allSheetData.length; i++) {
        const rowData = allSheetData[i];
        const funderName = rowData[PROP_FUNDER_COL - 1];
        if (funderName && typeof funderName === 'string' && funderName.trim() !== "") {
            const funderKey = funderName.toLowerCase();
            if (!funderIndex.has(funderKey)) { funderIndex.set(funderKey, []); }
            funderIndex.get(funderKey).push({
                row: i + 1, rowData: rowData, emailId: rowData[PROP_EMAIL_ID_COL - 1],
                funder: funderName, title: rowData[PROP_TITLE_COL - 1],
                status: rowData[PROP_STATUS_COL - 1], peakStatus: rowData[PROP_PEAK_STATUS_COL - 1]
            });
        }
    }
    
    const threadsToProcess = procLbl.getThreads(0, 20);
    if (threadsToProcess.length === 0) { /* ... logging ... */ return; }

    const messagesToSort = threadsToProcess.flatMap(thread => thread.getMessages()).sort((a, b) => a.getDate() - b.getDate());
    const dataToUpdate = [];
    const newRowsData = [];
    let threadProcessingOutcomes = {};

    for (const message of messagesToSort) {
        if ((new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000 > 320) break;
        const msgId = message.getId();
        try {
            const geminiResult = config.parserFunction(message.getSubject(), message.getPlainBody(), geminiApiKey);
            const handlerResult = config.dataHandler(geminiResult, message, funderIndex, dataSheet);
            if (handlerResult.updateInfo) { dataToUpdate.push(handlerResult.updateInfo); }
            if (handlerResult.newRowData) { newRowsData.push(...handlerResult.newRowData); }
            threadProcessingOutcomes[message.getThread().getId()] = handlerResult.requiresManualReview ? 'manual' : 'done';
        } catch (e) { /* ... error logging ... */ threadProcessingOutcomes[message.getThread().getId()] = 'manual'; }
    }

    if (dataToUpdate.length > 0) {
        dataToUpdate.forEach(update => dataSheet.getRange(update.row, 1, 1, update.values.length).setValues([update.values]));
    }
    if (newRowsData.length > 0) {
        dataSheet.getRange(dataSheet.getLastRow() + 1, 1, newRowsData.length, newRowsData[0].length).setValues(newRowsData);
    }
    applyFinalLabels(threadProcessingOutcomes, procLbl, processedLblObj, manualLblObj);
    Logger.log(`\n==== ${FUNC_NAME} FINISHED ====`);
}
// END SNIPPET 5A

/**
 * Parser function specific to the Proposal Tracker.
 * @param {string} subject
 * @param {string} body
 * @param {string} key
 * @returns {object}
 */
function _proposalStatusParser(subject, body, key) {
    return callGemini_forProposalStatus(subject, body, key);
}

// START SNIPPET 5B: Replace _proposalDataHandler in Main.js
function _proposalDataHandler(geminiResult, message, funderIndex, dataSheet) {
    if (!geminiResult) return { requiresManualReview: true };

    const { funderName, proposalTitle, submissionStatus } = geminiResult;
    let requiresManualReview = (funderName === MANUAL_REVIEW_NEEDED || proposalTitle === MANUAL_REVIEW_NEEDED);

    let existingRowInfo = null;
    if (!requiresManualReview) {
        const potentialMatches = funderIndex.get(String(funderName).toLowerCase()) || [];
        // Find an existing entry with the same funder AND proposal title
        existingRowInfo = potentialMatches.find(e => e.title && e.title.toLowerCase() === proposalTitle.toLowerCase());
    }

    if (existingRowInfo) {
        // --- UPDATE PATH ---
        const rowDataForSheet = [...existingRowInfo.rowData];
        rowDataForSheet[PROP_LAST_UPDATE_COL - 1] = message.getDate();
        rowDataForSheet[PROP_EMAIL_SUBJ_COL - 1] = message.getSubject();
        rowDataForSheet[PROP_EMAIL_LINK_COL - 1] = `https://mail.google.com/mail/u/0/#inbox/${message.getId()}`;
        rowDataForSheet[PROP_EMAIL_ID_COL - 1] = message.getId();

        const currentStatus = String(rowDataForSheet[PROP_STATUS_COL - 1]).trim() || STATUS_DRAFTING;
        const currentRank = STATUS_HIERARCHY[currentStatus] ?? 0;
        const newRank = STATUS_HIERARCHY[submissionStatus] ?? 0;
        if (newRank >= currentRank) { // Only update status if it's a forward or equal progression
            rowDataForSheet[PROP_STATUS_COL - 1] = submissionStatus;
        }

        const currentPeak = existingRowInfo.peakStatus || currentStatus;
        const peakRank = STATUS_HIERARCHY[currentPeak] ?? 0;
        const finalStatusRank = STATUS_HIERARCHY[rowDataForSheet[PROP_STATUS_COL - 1]] ?? 0;
        if (finalStatusRank > peakRank) {
            rowDataForSheet[PROP_PEAK_STATUS_COL - 1] = rowDataForSheet[PROP_STATUS_COL - 1];
        }
        return { updateInfo: { row: existingRowInfo.row, values: rowDataForSheet }, requiresManualReview };
    } else {
        // --- CREATE NEW ROW PATH ---
        const newRowData = new Array(TOTAL_COLUMNS_IN_PROPOSAL_SHEET).fill("");
        newRowData[PROP_PROC_TS_COL - 1] = new Date();
        newRowData[PROP_SUBMIT_DATE_COL - 1] = message.getDate();
        newRowData[PROP_FUNDER_COL - 1] = funderName;
        newRowData[PROP_TITLE_COL - 1] = proposalTitle;
        newRowData[PROP_STATUS_COL - 1] = submissionStatus || STATUS_SUBMITTED;
        newRowData[PROP_PEAK_STATUS_COL - 1] = submissionStatus || STATUS_SUBMITTED;
        newRowData[PROP_LAST_UPDATE_COL - 1] = message.getDate();
        newRowData[PROP_EMAIL_SUBJ_COL - 1] = message.getSubject();
        newRowData[PROP_EMAIL_LINK_COL - 1] = `https://mail.google.com/mail/u/0/#inbox/${message.getId()}`;
        newRowData[PROP_EMAIL_ID_COL - 1] = message.getId();
        return { newRowData: [newRowData], requiresManualReview };
    }
}
// END SNIPPET 5B

/**
 * Trigger handler for hourly email processing.
 */
function processEmails_triggerHandler() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scriptProperties = PropertiesService.getScriptProperties();
    Logger.log('Hourly email processing trigger started.');
    // The ss and scriptProperties are fetched here and passed down.
    processProposalEmails(ss, scriptProperties);
    Logger.log('Hourly email processing trigger finished.');
}

/**
 * Trigger handler for daily stale application checks.
 */
function markStale_triggerHandler() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
        Logger.log('markStale_triggerHandler: Could not get active spreadsheet. Aborting.');
        return;
    }
    Logger.log('Daily stale check trigger started.');
    markStaleProposals(ss);
    Logger.log('Daily stale check trigger finished.');
}

/**
 * Main "stub" function for processing grant proposal emails.
 * This function is called by the time-driven trigger.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Properties.Properties} scriptProperties
 */
function processProposalEmails(ss, scriptProperties) {
    const proposalProcessingConfig = {
        moduleName: "Proposal Tracker",
        sheetTabName: PROPOSAL_TRACKER_SHEET_TAB_NAME,
        parserFunction: _proposalStatusParser,
        dataHandler: _proposalDataHandler,
        // Pass the required Gmail labels from Config.js
        gmailLabelToProcess: TRACKER_GMAIL_LABEL_TO_PROCESS,
        gmailLabelProcessed: TRACKER_GMAIL_LABEL_PROCESSED,
        gmailLabelManualReview: TRACKER_GMAIL_LABEL_MANUAL_REVIEW
    };
    // The _processingEngine is generic and powerful, so we can reuse it without changes.
    _processingEngine(proposalProcessingConfig, ss, scriptProperties);
}

/**
 * Marks stale proposals as "Declined".
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object.
 */
function markStaleProposals(ss) {
    const FUNC_NAME = "markStaleProposals";
    Logger.log(`\n==== ${FUNC_NAME}: START (${new Date().toLocaleString()}) ====`);
    if (!ss) { /* ... error logging ... */ return; }
    const dataSheet = ss.getSheetByName(PROPOSAL_TRACKER_SHEET_TAB_NAME);
    if (!dataSheet) { /* ... error logging ... */ return; }

    const dataRange = dataSheet.getDataRange();
    const sheetValues = dataRange.getValues();
    const currentDate = new Date();
    const staleThresholdDate = new Date();
    staleThresholdDate.setDate(currentDate.getDate() - (WEEKS_THRESHOLD * 7));

    let updatedProposalsCount = 0;
    for (let i = 1; i < sheetValues.length; i++) {
        const currentStatus = sheetValues[i][PROP_STATUS_COL - 1];
        const lastUpdateDate = new Date(sheetValues[i][PROP_LAST_UPDATE_COL - 1]);

        if (currentStatus && !FINAL_STATUSES_FOR_STALE_CHECK.has(currentStatus) && lastUpdateDate && lastUpdateDate < staleThresholdDate) {
            sheetValues[i][PROP_STATUS_COL - 1] = STATUS_DECLINED;
            sheetValues[i][PROP_LAST_UPDATE_COL - 1] = currentDate;
            sheetValues[i][PROP_NOTES_COL - 1] = (sheetValues[i][PROP_NOTES_COL - 1] + ` (Auto-updated to Declined on ${currentDate.toLocaleDateString()})`).trim();
            updatedProposalsCount++;
        }
    }

    if (updatedProposalsCount > 0) {
        dataRange.setValues(sheetValues);
        Logger.log(`[${FUNC_NAME} INFO] Updated ${updatedProposalsCount} stale proposals to '${STATUS_DECLINED}'.`);
    } else {
        Logger.log(`[${FUNC_NAME} INFO] No stale proposals found.`);
    }
}

// REPLACE the old onOpen with this one
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menuName = CUSTOM_MENU_NAME || '⚙️ FundingFlock.AI Tools';
  const menu = ui.createMenu(menuName);
  
  menu.addItem('🚀 Finalize Project Setup', 'userDrivenFullSetup');
  menu.addSeparator();

  // --- ADDED SECTION ---
  menu.addSubMenu(ui.createMenu('Manual Processing')
      .addItem('📧 Process Proposal Emails', 'processEmails_triggerHandler')
      .addItem('🗑️ Mark Stale Proposals', 'markStale_triggerHandler'));
  menu.addSeparator();
  // --- END ADDED SECTION ---

  menu.addSubMenu(ui.createMenu('Admin & Config')
      .addItem('🔑 Set Gemini API Key', 'setSharedGeminiApiKey_UI')
      .addItem('🔍 Show All User Properties', 'showAllUserProperties'));
  menu.addSeparator();
  menu.addItem('❌ Uninstall Backend', 'uninstall');
  menu.addToUi();

  // Welcome prompt logic remains unchanged
  const scriptProperties = PropertiesService.getScriptProperties();
  const initialSetupDone = scriptProperties.getProperty('initialSetupDone_vFF_1');
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();

  if (!initialSetupDone && activeSS.getId() !== TEMPLATE_SHEET_ID) {
    showWelcomePrompt_FF();
  }
}

// ADD THIS HELPER FUNCTION TO Main.js

/**
 * Displays a one-time welcome and instruction prompt to the user upon first opening the sheet.
 */
function showWelcomePrompt_FF() {
    const ui = SpreadsheetApp.getUi();
    const title = "Welcome to FundingFlock.AI!";
    const message = "Your new Grant Tracker sheet is ready!\n\nTo complete the setup and activate all features (like the dashboard and formatting), please go to the new menu item:\n\n⚙️ FundingFlock.AI Tools > 🚀 Finalize Project Setup\n\nThis is a required one-time step.";
    ui.alert(title, message, ui.ButtonSet.OK);
}

// ADD THIS CORRECTED FUNCTION
function userDrivenFullSetup() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const activeSS = SpreadsheetApp.getActiveSpreadsheet();
  
  // FIX #1: Corrected TEMPLATE_SHEET_ID
  if (typeof TEMPLATE_SHEET_ID !== 'undefined' && activeSS.getId() === TEMPLATE_SHEET_ID) {
      ui.alert('Action Not Allowed on Template', 'This setup function cannot be run on the master template sheet.', ui.ButtonSet.OK);
      return;
  }

  ui.alert('Starting FundingFlock Setup', 'The full project setup will now begin. This will format your sheets, create a dashboard, and prepare your grant tracker.', ui.ButtonSet.OK);
  
  const setupResult = runFullProjectInitialSetup(activeSS);
  
  if (setupResult && setupResult.success) {
      scriptProperties.setProperty('initialSetupDone_vFF_1', 'true');
      
      // FIX #2: Corrected activeSS.getName()
      let finalMessage = `Setup is complete for "${activeSS.getName()}".\n\nYour Grant Tracker is ready. To enable AI features, please set your Gemini API Key using the 'Admin & Config' menu.`;
      
      ui.alert('Setup Complete', finalMessage, ui.ButtonSet.OK);
  } else {
      ui.alert('Setup Issues Encountered', `The project setup had some issues.\n\nPlease check the script logs for more details (Extensions > Apps Script > Executions).`, ui.ButtonSet.OK);
  }
}

/**
 * Uninstalls triggers and the Gmail filter created by the script.
 */
function uninstall() {
    const FUNC_NAME = "uninstall";
    Logger.log(`\n==== ${FUNC_NAME}: STARTING ====`);
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Confirm Full Uninstall',
        'This will remove all triggers AND the Gmail filter that automatically labels grant emails. Are you sure?',
        ui.ButtonSet.YES_NO);

    if (response !== ui.Button.YES) {
        ui.alert('Uninstall Canceled', 'No changes were made.', ui.ButtonSet.OK);
        return;
    }

    // 1. Remove Triggers
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    Logger.log(`[${FUNC_NAME}] All ${triggers.length} triggers removed.`);

    // 2. Remove Gmail Filter
    try {
        const filterQuery = TRACKER_GMAIL_FILTER_QUERY; // From Config.js
        const filters = Gmail.Users.Settings.Filters.list('me').filter;
        let filterIdToRemove = null;

        if (filters && filters.length > 0) {
            const targetFilter = filters.find(f => f.criteria && f.criteria.query === filterQuery);
            if (targetFilter) {
                filterIdToRemove = targetFilter.id;
            }
        }

        if (filterIdToRemove) {
            Gmail.Users.Settings.Filters.remove('me', filterIdToRemove);
            Logger.log(`[${FUNC_NAME}] Successfully removed Gmail filter with query: "${filterQuery}"`);
            ui.alert('Uninstall Complete', `All triggers and the Gmail filter have been removed.`, ui.ButtonSet.OK);
        } else {
            Logger.log(`[${FUNC_NAME}] No matching Gmail filter was found to remove.`);
            ui.alert('Uninstall Complete', 'All triggers have been removed. No matching Gmail filter was found.', ui.ButtonSet.OK);
        }
    } catch (e) {
        Logger.log(`[${FUNC_NAME} ERROR] Could not remove Gmail filter: ${e.message}. Please check your advanced Gmail service is enabled.`);
        ui.alert('Uninstall Partially Complete', `All triggers were removed, but there was an error removing the Gmail filter. Please check logs.`, ui.ButtonSet.OK);
    }
}

/**
 * Placeholder function for a future feature.
 */
function processOpportunities_placeholder() {
    SpreadsheetApp.getUi().alert("This feature is not yet implemented.");
}

function activateAiFeatures() {
  const FUNC_NAME = "activateAiFeatures";
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const userProps = PropertiesService.getUserProperties();

  // This function is now simplified to check for a key set manually via the UI.
  // The web app sync logic has been removed as MASTER_WEB_APP_URL is not part of this project's config.
  Logger.log(`[${FUNC_NAME}] Checking for locally set API key.`);

  const apiKey = userProps.getProperty(GEMINI_API_KEY_PROPERTY);

  if (apiKey && apiKey.trim() !== "" && apiKey.startsWith("AIza") && apiKey.length > 30) {
    scriptProperties.setProperty('aiFeaturesActive', 'true');
    Logger.log(`[${FUNC_NAME}] Valid API key found in UserProperties. AI features marked as active.`);
    ui.alert('AI Features Active', 'A valid Gemini API Key is configured. AI-powered features are enabled.', ui.ButtonSet.OK);
    return true;
  } else {
    scriptProperties.setProperty('aiFeaturesActive', 'false');
    Logger.log(`[${FUNC_NAME}] No valid API key found. AI features are disabled.`);
    ui.alert('AI Features Not Active', "Could not find a valid Gemini API Key. Please use the 'Admin & Config' -> 'Set Gemini API Key' menu to add your key.", ui.ButtonSet.OK);
    return false;
  }
}
