/**
 * @file Contains functions for creating, verifying, and managing
 * time-driven triggers for the project.
 */

/**
 * Creates or verifies a time-based trigger for a given function.
 * It checks for the trigger's existence to avoid creating duplicates.
 * @param {string} functionName The name of the function to trigger.
 * @param {number} hours The interval in hours.
 * @returns {boolean} True if a new trigger was created, false if it already existed.
 */
function createTimeDrivenTrigger(functionName, hours) {
  const FUNC_NAME = `createTimeDrivenTrigger for ${functionName}`;
  try {
    const existingTriggers = ScriptApp.getProjectTriggers();
    const triggerExists = existingTriggers.some(t => t.getHandlerFunction() === functionName);

    if (!triggerExists) {
      ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyHours(hours)
        .create();
      Logger.log(`[${FUNC_NAME} INFO] Trigger CREATED successfully.`);
      return true;
    } else {
      Logger.log(`[${FUNC_NAME} INFO] Trigger ALREADY EXISTS.`);
      return false;
    }
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Failed to create or verify trigger: ${e.message}`);
    return false;
  }
}

/**
 * Creates or verifies a daily time-based trigger for marking stale applications.
 * @returns {boolean} True if a new trigger was created, false if it already existed.
 */
function createOrVerifyStaleRejectTrigger() {
  const FUNC_NAME = 'createOrVerifyStaleRejectTrigger';
  const HANDLER_FUNCTION = 'markStaleApplicationsAsRejected';
  try {
    const existingTriggers = ScriptApp.getProjectTriggers();
    const triggerExists = existingTriggers.some(t => t.getHandlerFunction() === HANDLER_FUNCTION);

    if (!triggerExists) {
      ScriptApp.newTrigger(HANDLER_FUNCTION)
        .timeBased()
        .everyDays(1)
        .atHour(2) // Runs around 2 AM in the script's timezone
        .create();
      Logger.log(`[${FUNC_NAME} INFO] Daily stale-check trigger CREATED successfully.`);
      return true;
    } else {
      Logger.log(`[${FUNC_NAME} INFO] Daily stale-check trigger ALREADY EXISTS.`);
      return false;
    }
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Failed to create or verify stale-check trigger: ${e.message}`);
    return false;
  }
}
