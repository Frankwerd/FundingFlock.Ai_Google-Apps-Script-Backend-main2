/**
 * @file Contains functions dedicated to parsing email content (subject, body, sender)
 * using regular expressions and keyword matching to extract job application details.
 */

/**
 * Parses the email body for keywords to determine the proposal status.
 * It uses predefined keyword lists from `Config.js`.
 * @param {string} plainBody The plain text body of the email.
 * @returns {string|null} The determined status or null if no keywords are matched.
 */
function parseBodyForStatus(plainBody) {
  if (!plainBody || plainBody.length < 10) {
    if (DEBUG_MODE) Logger.log("[PARSING_UTILS] Body too short for status parse.");
    return null;
  }
  // Normalize the text for reliable matching
  let bL = plainBody.toLowerCase().replace(/[.,!?;:()\[\]{}'"“”‘’\-–—]/g, ' ').replace(/\s+/g, ' ').trim();
  
  // Keywords are sourced from your new Config.js
  if (AWARDED_KEYWORDS.some(k => bL.includes(k))) {
    Logger.log(`[PARSING_UTILS] Matched AWARDED_STATUS.`);
    return STATUS_AWARDED;
  }
  if (UNDER_REVIEW_KEYWORDS.some(k => bL.includes(k))) {
    Logger.log(`[PARSING_UTILS] Matched UNDER_REVIEW_STATUS.`);
    return STATUS_UNDER_REVIEW;
  }
  if (DECLINED_KEYWORDS.some(k => bL.includes(k))) {
    Logger.log(`[PARSING_UTILS] Matched DECLINED_STATUS.`);
    return STATUS_DECLINED;
  }

  if (DEBUG_MODE) Logger.log("[PARSING_UTILS] No specific status keywords found by regex.");
  return null; // No specific status matched
}
