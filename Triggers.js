/**
 * @file Manages time-driven triggers for the project.
 */

function createTimeDrivenTrigger(functionName, hours) {
  const FUNC_NAME = `createTimeDrivenTrigger for ${functionName}`;
  try {
    const existingTriggers = ScriptApp.getProjectTriggers();
    const triggerExists = existingTriggers.some(t => t.getHandlerFunction() === functionName);

    if (!triggerExists) {
        if (functionName === 'processEmails_triggerHandler') {
            ScriptApp.newTrigger(functionName)
                .timeBased()
                .everyDays(1) // CHANGED: Runs once per day
                .atHour(1)    // Around 1 AM in the server's timezone
                .create();
            Logger.log(`[${FUNC_NAME} INFO] Daily email processing trigger CREATED successfully.`);
        } else {
             // Fallback for other potential triggers if needed in the future
            ScriptApp.newTrigger(functionName)
                .timeBased()
                .everyHours(hours)
                .create();
            Logger.log(`[${FUNC_NAME} INFO] Hourly trigger CREATED successfully for ${functionName}.`);
        }
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

function createOrVerifyStaleProposalTrigger() {
  const FUNC_NAME = 'createOrVerifyStaleProposalTrigger';
  const HANDLER_FUNCTION = 'markStaleProposals';
  try {
    const existingTriggers = ScriptApp.getProjectTriggers();
    const triggerExists = existingTriggers.some(t => t.getHandlerFunction() === HANDLER_FUNCTION);

    if (!triggerExists) {
      ScriptApp.newTrigger(HANDLER_FUNCTION)
        .timeBased()
        .everyWeeks(1) // CHANGED: Runs once per week
        .onWeekDay(ScriptApp.WeekDay.SUNDAY) // Runs on Sunday
        .atHour(2) // Around 2 AM
        .create();
      Logger.log(`[${FUNC_NAME} INFO] Weekly stale-check trigger for proposals CREATED.`);
      return true;
    } else {
      Logger.log(`[${FUNC_NAME} INFO] Weekly stale-check trigger for proposals ALREADY EXISTS.`);
      return false;
    }
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Failed to create trigger: ${e.message}`);
    return false;
  }
}
