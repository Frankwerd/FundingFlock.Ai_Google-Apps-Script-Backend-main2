/**
 * @file Contains functions for interacting with Gmail, primarily label creation and management.
 */

/**
 * Retrieves a Gmail label by name, creating it if it doesn't exist.
 * @param {string} labelName The name of the Gmail label to get or create.
 * @returns {GoogleAppsScript.Gmail.GmailLabel|null} The GmailLabel object or null if an error occurs.
 */
function getOrCreateLabel(labelName) {
  if (!labelName || typeof labelName !== 'string' || labelName.trim() === "") {
    Logger.log(`[GMAIL_UTIL ERROR] Invalid labelName provided to getOrCreateLabel: "${labelName}"`);
    return null;
  }
  let label = null;
  try {
    label = GmailApp.getUserLabelByName(labelName);
  } catch (e) {
    Logger.log(`[GMAIL_UTIL ERROR] Error checking for label "${labelName}": ${e.message}`);
    return null; // Propagate failure
  }

  if (!label) {
    if (DEBUG_MODE) Logger.log(`[GMAIL_UTIL DEBUG] Label "${labelName}" not found. Creating...`);
    try {
      label = GmailApp.createLabel(labelName);
      Logger.log(`[GMAIL_UTIL INFO] Successfully created label: "${labelName}"`);
    } catch (e) {
      Logger.log(`[GMAIL_UTIL ERROR] Failed to create label "${labelName}": ${e.message}\n${e.stack}`);
      return null; // Propagate failure
    }
  } else {
    if (DEBUG_MODE) Logger.log(`[GMAIL_UTIL DEBUG] Label "${labelName}" already exists.`);
  }

  if (label) {
    if (DEBUG_MODE) Logger.log(`[GMAIL_UTIL DEBUG RETURN CHECK] Returning label for "${labelName}". Type: ${typeof label}, Constructor: ${label.constructor ? label.constructor.name : 'N/A'}`);
  } else {
    // This case should ideally be caught by earlier returns if creation failed.
    Logger.log(`[GMAIL_UTIL WARN RETURN CHECK] Returning NULL for "${labelName}" (label object is unexpectedly null after checks).`);
  }
  return label;
}

/**
 * Applies final labels to Gmail threads based on their processing outcome.
 * It removes the "processing" label and adds either a "processed" or "manual review" label.
 * @param {Object.<string, string>} threadOutcomes - An object mapping thread IDs to their outcome ('done' or 'manual').
 * @param {GoogleAppsScript.Gmail.GmailLabel} processingLabel - The label indicating threads are being processed.
 * @param {GoogleAppsScript.Gmail.GmailLabel} processedLabelObj - The label for successfully processed threads.
 * @param {GoogleAppsScript.Gmail.GmailLabel} manualReviewLabelObj - The label for threads requiring manual review.
 */
function applyFinalLabels(threadOutcomes, processingLabel, processedLabelObj, manualReviewLabelObj) {
  const threadIdsToUpdate = Object.keys(threadOutcomes);
  if (threadIdsToUpdate.length === 0) {
    Logger.log("[INFO] LABEL_MGMT: No thread outcomes to process for final labeling.");
    return;
  }
  Logger.log(`[INFO] LABEL_MGMT: Applying final labels for ${threadIdsToUpdate.length} threads.`);
  let successfulLabelChanges = 0;
  let labelErrors = 0;

  // Validate label objects before proceeding
  if (!processingLabel || typeof processingLabel.getName !== 'function') {
    Logger.log(`[ERROR] LABEL_MGMT: Invalid 'processingLabel' object provided. Aborting label application.`);
    return;
  }
  if (!processedLabelObj || typeof processedLabelObj.getName !== 'function') {
    Logger.log(`[ERROR] LABEL_MGMT: Invalid 'processedLabelObj' object provided. Aborting label application.`);
    return;
  }
  if (!manualReviewLabelObj || typeof manualReviewLabelObj.getName !== 'function') {
    Logger.log(`[ERROR] LABEL_MGMT: Invalid 'manualReviewLabelObj' object provided. Aborting label application.`);
    return;
  }

  const toProcessLabelName = processingLabel.getName(); // Get name once

  for (const threadId of threadIdsToUpdate) {
    const outcome = threadOutcomes[threadId]; // 'done' or 'manual'
    const targetLabelToAdd = (outcome === 'manual') ? manualReviewLabelObj : processedLabelObj;
    const targetLabelNameToAdd = targetLabelToAdd.getName(); // Get name once

    try {
      const thread = GmailApp.getThreadById(threadId);
      if (!thread) {
        Logger.log(`[WARN] LABEL_MGMT: Thread ${threadId} not found (may have been deleted). Skipping.`);
        labelErrors++;
        continue;
      }

      const currentThreadLabels = thread.getLabels().map(l => l.getName());
      let labelsActuallyChangedThisThread = false;

      // Remove "To Process" label
      if (currentThreadLabels.includes(toProcessLabelName)) {
        try {
          thread.removeLabel(processingLabel);
          if (DEBUG_MODE) Logger.log(`[DEBUG] LABEL_MGMT: Removed "${toProcessLabelName}" from thread ${threadId}`);
          labelsActuallyChangedThisThread = true;
        } catch (eRem) {
          Logger.log(`[WARN] LABEL_MGMT: Failed to remove "${toProcessLabelName}" from thread ${threadId}: ${eRem.message}`);
          // Continue to attempt adding the target label
        }
      }

      // Add "Processed" or "Manual Review" label
      if (!currentThreadLabels.includes(targetLabelNameToAdd)) {
        try {
          thread.addLabel(targetLabelToAdd);
          Logger.log(`[INFO] LABEL_MGMT: Added "${targetLabelNameToAdd}" to thread ${threadId}`);
          labelsActuallyChangedThisThread = true;
        } catch (eAdd) {
          Logger.log(`[ERROR] LABEL_MGMT: Failed to add "${targetLabelNameToAdd}" to thread ${threadId}: ${eAdd.message}`);
          labelErrors++;
          continue; // Skip to next thread if adding the crucial label fails
        }
      } else {
        if (DEBUG_MODE) Logger.log(`[DEBUG] LABEL_MGMT: Thread ${threadId} already has target label "${targetLabelNameToAdd}". No add needed.`);
      }

      if (labelsActuallyChangedThisThread) {
        successfulLabelChanges++;
        Utilities.sleep(250 + Math.floor(Math.random() * 150)); // Slightly increased sleep
      }

    } catch (eThread) {
      Logger.log(`[ERROR] LABEL_MGMT: General error processing thread ${threadId}: ${eThread.message}. Thread might be inaccessible.`);
      labelErrors++;
    }
  }
  Logger.log(`[INFO] LABEL_MGMT: Finished applying final labels. Success changes/verified: ${successfulLabelChanges}. Errors: ${labelErrors}.`);
}
