function _setupModule(config) {
  const FUNC_NAME = "_setupModule";
  Logger.log(`\n==== ${FUNC_NAME}: STARTING - ${config.moduleName} Module Setup ====`);
  let messages = [];
  let moduleSuccess = true;
  let dataSh; // This will hold the sheet object for the module

  if (!config.activeSS || typeof config.activeSS.getId !== 'function') {
    const errMsg = "CRITICAL: Invalid spreadsheet object passed.";
    Logger.log(`[${FUNC_NAME} ERROR] ${errMsg}`);
    return { success: false, messages: [errMsg] };
  }
  const activeSS = config.activeSS;
  Logger.log(`[${FUNC_NAME} INFO] Operating on: "${activeSS.getName()}" (ID: ${activeSS.getId()})`);

  // Setup module-specific sheet
  try {
    dataSh = activeSS.getSheetByName(config.sheetTabName);
    if (!dataSh) {
      dataSh = activeSS.insertSheet(config.sheetTabName);
      Logger.log(`[${FUNC_NAME} INFO] Created new sheet: "${config.sheetTabName}".`);
    } else {
      Logger.log(`[${FUNC_NAME} INFO] Found existing sheet: "${config.sheetTabName}".`);
    }
    if (!setupSheetFormatting(dataSh, config.sheetHeaders, config.columnWidths, true, config.bandingTheme)) {
      throw new Error(`Formatting failed for "${config.sheetTabName}".`);
    }
    dataSh.setTabColor(config.tabColor);
    messages.push(`Sheet '${config.sheetTabName}': Setup OK. Color: ${config.tabColor}.`);
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Module sheet setup failed: ${e.toString()}\nStack: ${e.stack}`);
    messages.push(`Module sheet setup FAILED: ${e.message}.`);
    moduleSuccess = false;
  }


  // --- B. Gmail Label & Filter Setup ---
  let trackerToProcessLabelId = null;
  if (moduleSuccess) {
    Logger.log(`[${FUNC_NAME} INFO] Setting up Gmail labels & filters for ${config.moduleName}...`);
    try {
      getOrCreateLabel(config.gmailLabelParent); Utilities.sleep(100);       // From Config.gs
      const toProcessLabelObject = getOrCreateLabel(config.gmailLabelToProcess); Utilities.sleep(100); // From Config.gs
      getOrCreateLabel(config.gmailLabelProcessed); Utilities.sleep(100);   // From Config.gs
      if (config.gmailLabelManualReview) {
        getOrCreateLabel(config.gmailLabelManualReview); Utilities.sleep(100); // From Config.gs
      }

      if (toProcessLabelObject) {
        Utilities.sleep(300);
        const advancedGmailService = Gmail; // Assumes Advanced Gmail API Service is enabled
        if (!advancedGmailService || !advancedGmailService.Users || !advancedGmailService.Users.Labels) {
          throw new Error("Advanced Gmail Service not available/enabled for label ID fetch.");
        }
        const labelsListResponse = advancedGmailService.Users.Labels.list('me');
        if (labelsListResponse.labels && labelsListResponse.labels.length > 0) {
          const targetLabelInfo = labelsListResponse.labels.find(l => l.name === config.gmailLabelToProcess);
          if (targetLabelInfo && targetLabelInfo.id) {
            trackerToProcessLabelId = targetLabelInfo.id;
          } else { Logger.log(`[${FUNC_NAME} WARN] Label "${config.gmailLabelToProcess}" ID not found via Advanced Service.`); }
        } else { Logger.log(`[${FUNC_NAME} WARN} No labels returned by Advanced Gmail Service.`); }
      }
      if (!trackerToProcessLabelId) throw new Error(`CRITICAL: Could not get ID for Gmail label "${config.gmailLabelToProcess}". Filter creation will fail.`);
      messages.push("Tracker Labels & 'To Process' ID: OK.");

      // Filter Creation
      const filterQuery = config.gmailFilterQuery; // from Config.gs
      const gmailApiServiceForFilter = Gmail; // Advanced Gmail Service
      let filterExists = false;
      const existingFiltersResponse = gmailApiServiceForFilter.Users.Settings.Filters.list('me');
      const existingFiltersList = (existingFiltersResponse && existingFiltersResponse.filter && Array.isArray(existingFiltersResponse.filter)) ? existingFiltersResponse.filter : [];

      for (const filterItem of existingFiltersList) {
        if (filterItem.criteria?.query === filterQuery && filterItem.action?.addLabelIds?.includes(trackerToProcessLabelId)) {
          filterExists = true; break;
        }
      }
      if (!filterExists) {
        const filterResource = { criteria: { query: filterQuery }, action: { addLabelIds: [trackerToProcessLabelId] } };
        const createdFilterResponse = gmailApiServiceForFilter.Users.Settings.Filters.create(filterResource, 'me');
        if (!createdFilterResponse || !createdFilterResponse.id) {
          throw new Error(`Gmail filter creation for tracker FAILED or did not return ID. Response: ${JSON.stringify(createdFilterResponse)}`);
        }
        messages.push("Tracker Filter: CREATED.");
      } else { messages.push("Tracker Filter: Exists."); }

    } catch (e) {
      Logger.log(`[${FUNC_NAME} ERROR] Gmail Label/Filter setup: ${e.toString()}`);
      messages.push(`Gmail Label/Filter setup FAILED: ${e.message}.`); moduleSuccess = false;
    }
  }

// START SNIPPET 6A: Replace the dummy data line in ModuleUtils.js
if (moduleSuccess && dataSh && dataSh.getLastRow() <= 1) {
    Logger.log(`[${FUNC_NAME} INFO] Adding dummy data to "${config.sheetTabName}".`);
    try {
        const today = new Date();
        const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
        const dummyRows = [
            [new Date(), weekAgo, "Demo Foundation", "Community Arts Project", STATUS_SUBMITTED, STATUS_SUBMITTED, weekAgo, 5000, 0, "Initial Submission", "http://example.com", "dummy1", "Notes here"],
            [new Date(), today, "Tech Grant Initiative", "STEM Education Program", STATUS_UNDER_REVIEW, STATUS_UNDER_REVIEW, today, 25000, 0, "Follow-up Email", "http://example.com", "dummy2", "Pending response"]
        ];
        const sizedDummyRows = dummyRows.map(r => {
            while (r.length < config.sheetHeaders.length) r.push("");
            return r.slice(0, config.sheetHeaders.length);
        });
        dataSh.getRange(2, 1, sizedDummyRows.length, sizedDummyRows[0].length).setValues(sizedDummyRows);
        messages.push(`Dummy data added for dashboard priming.`);
    } catch (e) {
        Logger.log(`[${FUNC_NAME} WARN] Failed adding dummy data: ${e.message}`);
    }
}
// END SNIPPET 6A

  // --- F. Trigger Verification/Creation ---
  if (moduleSuccess) {
    Logger.log(`[${FUNC_NAME} INFO] Setting up triggers for ${config.moduleName} module...`);
    try { // Assumes createTimeDrivenTrigger & createOrVerifyStaleRejectTrigger are in Triggers.gs
      if (createTimeDrivenTrigger(config.triggerFunctionName, config.triggerIntervalHours)) messages.push(`Trigger '${config.triggerFunctionName}': CREATED.`);
      else messages.push(`Trigger '${config.triggerFunctionName}': Exists/Verified.`);
      if (config.staleRejectFunctionName) {
        if (createOrVerifyStaleProposalTrigger()) messages.push(`Trigger 'markStaleProposals': CREATED.`);
        else messages.push(`Trigger 'markStaleProposals': Exists/Verified.`);
      }
    } catch (e) {
      Logger.log(`[${FUNC_NAME} ERROR] Trigger setup failed: ${e.toString()}`);
      messages.push(`Trigger setup FAILED: ${e.message}.`);
      moduleSuccess = false;
    }
  } else {
    messages.push(`Triggers for ${config.moduleName} Module: SKIPPED due to earlier failures.`);
  }

  Logger.log(`\n==== ${FUNC_NAME} ${moduleSuccess ? "COMPLETED." : "ISSUES."} ====`);
  return { success: moduleSuccess, messages: messages };
}
