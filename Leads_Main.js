/**
 * @file Contains the primary functions for the Job Leads Tracker module,
 * including initial setup of the leads sheet/labels/filters and the
 * ongoing processing of job lead emails.
 */

/**
 * Sets up the Job Leads Tracker module.
 * This function ensures the "Potential Job Leads" sheet exists and is formatted,
 * creates the necessary Gmail labels and filter, and sets up a daily trigger
 * for processing new job leads.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} passedSpreadsheet The spreadsheet object to set up.
 * @returns {{success: boolean, messages: string[]}} An object containing the setup result and messages.
 */
function runInitialSetup_JobLeadsModule(passedSpreadsheet) {
    const leadsConfig = {
        activeSS: passedSpreadsheet,
        moduleName: "Job Leads Tracker",
        sheetTabName: LEADS_SHEET_TAB_NAME,
        sheetHeaders: LEADS_SHEET_HEADERS,
        columnWidths: LEADS_SHEET_COLUMN_WIDTHS,
        bandingTheme: SpreadsheetApp.BandingTheme.YELLOW,
        tabColor: BRAND_COLORS.HUNYADI_YELLOW,
        gmailLabelParent: LEADS_GMAIL_LABEL_PARENT,
        gmailLabelToProcess: LEADS_GMAIL_LABEL_TO_PROCESS,
        gmailLabelProcessed: LEADS_GMAIL_LABEL_PROCESSED,
        gmailFilterQuery: LEADS_GMAIL_FILTER_QUERY,
        triggerFunctionName: 'processJobLeads',
        triggerIntervalHours: 3
    };
    return _setupModule(leadsConfig);
}

/**
 * Processes emails that have been labeled for job lead extraction.
 * This function is designed to be run on a time-driven trigger. It fetches emails,
 * calls the Gemini API to extract job leads, and writes the results to the spreadsheet.
 */
function _leadsParser(subject, body, key) {
    return callGemini_forJobLeads(body, key);
}

// In Leads_Main.js

function _leadsDataHandler(geminiResult, message, companyIndex, dataSheet) {
    const newRows = [];
    let requiresManualReview = false;

    // The 'geminiResult.data' is ALREADY the parsed JavaScript array from _callGeminiAPI
    if (geminiResult && geminiResult.success && Array.isArray(geminiResult.data)) {
        const extractedJobsArray = geminiResult.data;

        if (extractedJobsArray.length > 0) {
            Logger.log(`[_leadsDataHandler INFO] Gemini successfully extracted ${extractedJobsArray.length} job(s) from msg ${message.getId()}.`);

            for (const jobData of extractedJobsArray) {
                if (jobData && jobData.jobTitle && String(jobData.jobTitle).toLowerCase() !== 'n/a') {
                    const newRowData = new Array(LEADS_SHEET_HEADERS.length).fill("");

                    // Map the extracted data to the correct sheet columns
                    newRowData[LEADS_DATE_ADDED_COL - 1] = message.getDate();
                    newRowData[LEADS_COMPANY_COL - 1] = jobData.company || "N/A";
                    newRowData[LEADS_JOB_TITLE_COL - 1] = jobData.jobTitle || "N/A";
                    newRowData[LEADS_LOCATION_COL - 1] = jobData.location || "N/A";
                    newRowData[LEADS_SOURCE_LINK_COL - 1] = jobData.jobUrl || "N/A";
                    newRowData[LEADS_NOTES_COL - 1] = jobData.notes || "";
                    newRowData[LEADS_STATUS_COL - 1] = DEFAULT_LEAD_STATUS;
                    newRowData[LEADS_EMAIL_SUBJECT_COL - 1] = message.getSubject().substring(0, 500);
                    newRowData[LEADS_EMAIL_ID_COL - 1] = message.getId();
                    newRowData[LEADS_PROCESSED_TIMESTAMP_COL - 1] = new Date();

                    newRows.push(newRowData);
                }
            }
        } else {
            Logger.log(`[_leadsDataHandler INFO] Msg ${message.getId()}: Gemini call was successful but no job listings were found in the email (returned empty array).`);
        }
    } else {
        // Handle cases where the API call itself failed
        const errorInfo = {
            moduleName: "Job Leads Tracker",
            errorType: "Gemini API Error",
            details: geminiResult ? String(geminiResult.error) : "Unknown API response error",
            messageSubject: message.getSubject(),
            messageId: message.getId()
        };
        _writeErrorToSheet(dataSheet, errorInfo);
        requiresManualReview = true; // Mark for review if the API call itself fails.
    }

    return { newRowData: newRows, requiresManualReview: requiresManualReview };
}

function processJobLeads() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scriptProperties = PropertiesService.getScriptProperties();
    const leadsProcessingConfig = {
        moduleName: "Job Leads Tracker",
        sheetTabName: LEADS_SHEET_TAB_NAME,
        gmailLabelToProcess: LEADS_GMAIL_LABEL_TO_PROCESS,
        gmailLabelProcessed: LEADS_GMAIL_LABEL_PROCESSED,
        gmailLabelManualReview: null,
        parserFunction: _leadsParser,
        dataHandler: _leadsDataHandler,
    };
    _processingEngine(leadsProcessingConfig, ss, scriptProperties);
}
