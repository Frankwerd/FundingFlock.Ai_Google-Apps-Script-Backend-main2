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
 * Calls the Gemini API to parse job application details from an email.
 * @param {string} emailSubject The subject of the email.
 * @param {string} emailBody The plain text body of the email.
 * @param {string} apiKey The Gemini API key.
 * @returns {{company: string, title: string, status: string}|null} An object with the parsed details or null on failure.
 */
function callGemini_forApplicationDetails(emailSubject, emailBody, apiKey) {
  if (!apiKey || (!emailSubject && !emailBody)) {
    Logger.log("[INFO] GEMINI_PARSE_APP: API Key not provided or email content is empty. Skipping Gemini call.");
    return null;
  }

  const bodySnippet = emailBody ? emailBody.substring(0, 12000) : "";
  const prompt = `You are a highly specialized AI assistant expert in parsing job application-related emails for a tracking system. Your sole purpose is to analyze the provided email Subject and Body, and extract three key pieces of information: "company_name", "job_title", and "status". You MUST return this information ONLY as a single, valid JSON object, with no surrounding text, explanations, apologies, or markdown.

CRITICAL INSTRUCTIONS - READ AND FOLLOW CAREFULLY:

1.  "company_name":
    *   Extract the full, official name of the HIRING COMPANY.
    *   Do NOT extract the name of the ATS (e.g., "Greenhouse") or the job board (e.g., "LinkedIn").
    *   If the company name is genuinely unclear, use the exact string "${MANUAL_REVIEW_NEEDED}".

2.  "job_title":
    *   Extract the SPECIFIC job title the user applied for as mentioned in THIS email.
    *   If the job title is not clearly present, use the exact string "${MANUAL_REVIEW_NEEDED}".

3.  "status":
    *   Determine the current status of the application based on the content of THIS email.
    *   You MUST choose a status ONLY from the following exact list. Do not invent new statuses.
        *   "${DEFAULT_STATUS}" (Use for: Application submitted, application sent, successfully applied, application received)
        *   "${REJECTED_STATUS}" (Use for: Not moving forward, unfortunately, decided not to proceed, position filled)
        *   "${OFFER_STATUS}" (Use for: Offer of employment, pleased to offer, job offer)
        *   "${INTERVIEW_STATUS}" (Use for: Invitation to interview, schedule an interview, interview request)
        *   "${ASSESSMENT_STATUS}" (Use for: Online assessment, coding challenge, technical test, skills test)
        *   "${APPLICATION_VIEWED_STATUS}" (Use for: Application was viewed by recruiter, your profile was viewed for the role)
        *   "Update/Other" (Use for: General updates or if the status is unclear)

**Output Requirements**:
*   **ONLY JSON**: Your entire response must be a single, valid JSON object.
*   **Structure**: {"company_name": "...", "job_title": "...", "status": "..."}
*   **Irrelevant Emails**: If the email is clearly NOT a job application update (e.g., a newsletter, a job alert), your output MUST be: {"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "Not an Application"}

--- EMAIL TO PROCESS START ---
Subject: ${emailSubject}
Body:
${bodySnippet}
--- EMAIL TO PROCESS END ---

JSON Output:
`;

  const extractedData = _callGeminiAPI(prompt, apiKey, { logContext: "GEMINI_PARSE_APP" });

  if (extractedData && typeof extractedData.company_name !== 'undefined' && typeof extractedData.job_title !== 'undefined' && typeof extractedData.status !== 'undefined') {
    Logger.log(`[INFO] GEMINI_PARSE_APP: Success. C:"${extractedData.company_name}", T:"${extractedData.job_title}", S:"${extractedData.status}"`);
    return {
        company: extractedData.company_name || MANUAL_REVIEW_NEEDED,
        title: extractedData.job_title || MANUAL_REVIEW_NEEDED,
        status: extractedData.status || MANUAL_REVIEW_NEEDED
    };
  } else {
    Logger.log(`[WARN] GEMINI_PARSE_APP: JSON from Gemini missing fields or API call failed. Output: ${JSON.stringify(extractedData)}`);
    return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:MANUAL_REVIEW_NEEDED};
  }
}

function callGemini_forJobLeads(emailBody, apiKey) {
    if (typeof emailBody !== 'string') {
        Logger.log(`[GEMINI_LEADS CRITICAL ERR] emailBody not string. Type: ${typeof emailBody}`);
        return { success: false, data: null, error: `emailBody is not a string.` };
    }

    if (!apiKey || apiKey.trim() === '') {
        const errorMsg = "API Key is not set. Please set it in the configuration.";
        Logger.log(`[GEMINI_LEADS ERROR] ${errorMsg}`);
        return { success: false, data: null, error: errorMsg };
    }

    const promptText = `You are an expert AI assistant specializing in extracting job posting details from email content, typically from job alerts or direct emails containing job opportunities.
From the following "Email Content", identify each distinct job posting.

For EACH job posting found, extract the following details:
- "jobTitle": The specific title of the job role (e.g., "Senior Software Engineer", "Product Marketing Manager").
- "company": The name of the hiring company.
- "location": The primary location of the job (e.g., "San Francisco, CA", "Remote", "London, UK", "Hybrid - New York").
- "source": If identifiable from the email content, the origin or job board where this posting was listed (e.g., "LinkedIn Job Alert", "Indeed", "Wellfound", "Company Careers Page" if mentioned). If not explicitly stated, use "N/A".
- "jobUrl": A direct URL link to the job application page or a more detailed job description, if present in the email. If no direct link for *this specific job* is found, use "N/A".
- "notes": Briefly extract 2-3 key requirements, responsibilities, or unique aspects mentioned for this specific job if readily available in the email text (e.g., "Requires Python & AWS; 5+ yrs exp", "Focus on B2B SaaS marketing", "Fast-paced startup environment"). Keep notes concise (max 150 characters). If no specific details are easily extractable for this job, use "N/A".

Strict Formatting Instructions:
- Your entire response MUST be a single, valid JSON array.
- Each element of the array MUST be a JSON object representing one job posting.
- Each JSON object MUST have exactly these keys: "jobTitle", "company", "location", "source", "jobUrl", "notes".
- If a specific field for a job is not found or not applicable, its value MUST be the string "N/A".
- If no job postings at all are found in the email content, return an empty JSON array: [].
- Do NOT include any text, explanations, apologies, or markdown (like \`\`\`json\`\`\`) before or after the JSON array.

--- EXAMPLE OUTPUT START (for an email with two jobs) ---
[
  {
    "jobTitle": "Senior Frontend Developer",
    "company": "Innovatech Solutions",
    "location": "Remote (US)",
    "source": "LinkedIn Job Alert",
    "jobUrl": "https://linkedin.com/jobs/view/12345",
    "notes": "React, TypeScript, Agile environment. 5+ years experience. UI/UX focus."
  },
  {
    "jobTitle": "Data Scientist",
    "company": "Alpha Analytics Co.",
    "location": "Boston, MA",
    "source": "Direct Email from Recruiter",
    "jobUrl": "N/A",
    "notes": "Machine learning, Python, SQL. PhD preferred. Early-stage startup."
  }
]
--- EXAMPLE OUTPUT END ---

Email Content:
---
${emailBody.substring(0, 30000)} 
---
JSON Array Output:`;

    const parsedData = _callGeminiAPI(promptText, apiKey, { logContext: "GEMINI_LEADS" });

    if (parsedData) {
        return { success: true, data: parsedData, error: null };
    } else {
        return { success: false, data: null, error: "Failed to get a valid response from Gemini API." };
    }
}

