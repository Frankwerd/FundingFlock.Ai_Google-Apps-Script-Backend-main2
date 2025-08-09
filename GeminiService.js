/**
 * @file Handles all interactions with the Google Gemini API for
 * AI-powered parsing of email content to extract job application details and job leads.
 */

/**
 * Private function to handle the core Gemini API call with retry logic.
 * @param {string} prompt The complete prompt to send to the API.
 * @param {string} apiKey The user's Gemini API key.
 * @param {object} options Additional options for the API call.
 * @returns {object|null} The parsed JSON response or null on failure.
 * @private
 */
function _callGeminiAPI(prompt, apiKey, options = {}) {
  const { maxAttempts = 2, logContext = "GEMINI_API" } = options;
  const API_ENDPOINT = GEMINI_API_ENDPOINT_TEXT_ONLY + "?key=" + apiKey;

  const payload = {
    "contents": [{"parts": [{"text": prompt}]}],
    "generationConfig": { "temperature": 0.2, "maxOutputTokens": 8192, "topP": 0.95, "topK": 40 },
    "safetySettings": [
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
    ]
  };
  const fetchOptions = {'method':'post', 'contentType':'application/json', 'payload':JSON.stringify(payload), 'muteHttpExceptions':true};

  if(DEBUG_MODE) Logger.log(`[DEBUG] ${logContext}: Calling API. Prompt len (approx): ${prompt.length}`);

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = UrlFetchApp.fetch(API_ENDPOINT, fetchOptions);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      if(DEBUG_MODE) Logger.log(`[DEBUG] ${logContext} (Attempt ${attempt}): RC: ${responseCode}. Body(start): ${responseBody.substring(0,200)}`);

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        if (jsonResponse.candidates && jsonResponse.candidates[0]?.content?.parts?.[0]?.text) {
          let extractedJsonString = jsonResponse.candidates[0].content.parts[0].text.trim();
          if (extractedJsonString.startsWith("```json")) extractedJsonString = extractedJsonString.substring(7).trim();
          if (extractedJsonString.startsWith("```")) extractedJsonString = extractedJsonString.substring(3).trim();
          if (extractedJsonString.endsWith("```")) extractedJsonString = extractedJsonString.substring(0, extractedJsonString.length - 3).trim();

          if(DEBUG_MODE) Logger.log(`[DEBUG] ${logContext}: Cleaned JSON from API: ${extractedJsonString}`);
          try {
            return JSON.parse(extractedJsonString);
          } catch (e) {
            Logger.log(`[ERROR] ${logContext}: Error parsing JSON: ${e.toString()}\nString: >>>${extractedJsonString}<<<`);
            return null;
          }
        } else {
          Logger.log(`[ERROR] ${logContext}: API response structure unexpected. Body (start): ${responseBody.substring(0,500)}`);
          return null;
        }
      } else if (responseCode === 429) {
        Logger.log(`[WARN] ${logContext}: Rate limit (429). Attempt ${attempt}/${maxAttempts}. Waiting...`);
        if (attempt < maxAttempts) {
          Utilities.sleep(5000 + Math.floor(Math.random() * 5000));
        }
      } else {
        Logger.log(`[ERROR] ${logContext}: API HTTP error. Code: ${responseCode}. Body (start): ${responseBody.substring(0,500)}`);
        return null;
      }
    } catch (e) {
      Logger.log(`[ERROR] ${logContext}: Exception during API call (Attempt ${attempt}): ${e.toString()}\nStack: ${e.stack}`);
      if (attempt < maxAttempts) {
        Utilities.sleep(3000);
      }
    }
  }

  Logger.log(`[ERROR] ${logContext}: Failed after ${maxAttempts} attempts.`);
  return null;
}

/**
 * Calls the Gemini API to parse grant proposal details from an email.
 * @param {string} emailSubject The subject of the email.
 * @param {string} emailBody The plain text body of the email.
 * @param {string} apiKey The Gemini API key.
 * @returns {{funderName: string, proposalTitle: string, submissionStatus: string}|null} Parsed details or null on failure.
 */
function callGemini_forProposalStatus(emailSubject, emailBody, apiKey) {
  if (!apiKey || (!emailSubject && !emailBody)) {
    Logger.log("[GEMINI_SERVICE] API Key or email content is empty. Skipping Gemini call.");
    return null;
  }

  const bodySnippet = emailBody ? emailBody.substring(0, 12000) : "";
  const prompt = `${GEMINI_SYSTEM_INSTRUCTION_PROPOSAL_TRACKER}

--- EMAIL TO PROCESS START ---
Subject: ${emailSubject}
Body:
${bodySnippet}
--- EMAIL TO PROCESS END ---

JSON Output:
`;

  const extractedData = _callGeminiAPI(prompt, apiKey, { logContext: "GEMINI_PARSE_PROPOSAL" });

  if (extractedData && typeof extractedData.funderName !== 'undefined' && typeof extractedData.proposalTitle !== 'undefined' && typeof extractedData.submissionStatus !== 'undefined') {
    Logger.log(`[GEMINI_SERVICE] Success. F:"${extractedData.funderName}", T:"${extractedData.proposalTitle}", S:"${extractedData.submissionStatus}"`);
    return {
        funderName: extractedData.funderName || MANUAL_REVIEW_NEEDED,
        proposalTitle: extractedData.proposalTitle || MANUAL_REVIEW_NEEDED,
        submissionStatus: extractedData.submissionStatus || MANUAL_REVIEW_NEEDED
    };
  } else {
    Logger.log(`[GEMINI_SERVICE] WARN: JSON from Gemini missing fields or API call failed. Output: ${JSON.stringify(extractedData)}`);
    return {funderName:MANUAL_REVIEW_NEEDED, proposalTitle:MANUAL_REVIEW_NEEDED, submissionStatus:MANUAL_REVIEW_NEEDED};
  }
}
