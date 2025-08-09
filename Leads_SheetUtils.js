// File: Leads_SheetUtils.gs
// Description: Contains utility functions specific to the "Potential Job Leads" sheet,
// such as header mapping, retrieving processed email IDs, and writing job lead data or errors.

/**
 * Retrieves a specific sheet by name from a given spreadsheet ID and maps its header names to column numbers.
 * @param {string} ssId The ID of the spreadsheet.
 *   If null or undefined, it will attempt to use SpreadsheetApp.getActiveSpreadsheet().
 * @param {string} sheetName The name of the sheet to retrieve (e.g., LEADS_SHEET_TAB_NAME from Config.gs).
 * @return {{sheet: GoogleAppsScript.Spreadsheet.Sheet | null, headerMap: Object}} An object containing the sheet and its header map.
 */
function getSheetAndHeaderMapping_forLeads(ssId, sheetName) {
  let ss;
  try {
    if (ssId) {
      ss = SpreadsheetApp.openById(ssId);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Logger.log(`[LEADS_SHEET_UTIL ERROR] No Spreadsheet ID provided and no active spreadsheet found.`);
        return { sheet: null, headerMap: {} };
      }
      Logger.log(`[LEADS_SHEET_UTIL INFO] No SS ID provided, using active spreadsheet: "${ss.getName()}"`);
    }

    if (!sheetName) {
        Logger.log(`[LEADS_SHEET_UTIL ERROR] Sheet name not provided.`);
        return { sheet: null, headerMap: {} };
    }
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`[LEADS_SHEET_UTIL ERROR] Sheet "${sheetName}" not found in Spreadsheet: "${ss.getName()}" (ID: ${ss.getId()}).`);
      return { sheet: null, headerMap: {} };
    }

    const headerRowValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    if (!headerRowValues || headerRowValues.length === 0) {
        Logger.log(`[LEADS_SHEET_UTIL WARN] Header row is empty or unreadable in sheet "${sheetName}".`);
        return { sheet: sheet, headerMap: {} }; // Return sheet but empty map
    }
    const headers = headerRowValues[0];
    const headerMap = {};
    let actualHeaderCount = 0;
    headers.forEach((h, i) => {
      if (h && h.toString().trim() !== "") {
        headerMap[h.toString().trim()] = i + 1; // 1-based column index
        actualHeaderCount++;
      }
    });

    // Validate if all expected headers from LEADS_SHEET_HEADERS are present
    let missingHeaders = [];
    // LEADS_SHEET_HEADERS is from Config.gs
    LEADS_SHEET_HEADERS.forEach(expectedHeader => {
        if (!headerMap[expectedHeader]) {
            missingHeaders.push(expectedHeader);
        }
    });

    if (missingHeaders.length > 0) {
        Logger.log(`[LEADS_SHEET_UTIL WARN] Sheet "${sheetName}" is missing expected headers: ${missingHeaders.join(", ")}. This may cause issues with data writing.`);
    }
    
    if (actualHeaderCount === 0 && sheet.getLastRow() > 0) {
      Logger.log(`[LEADS_SHEET_UTIL WARN] No headers were mapped in sheet "${sheetName}", but the sheet has content. This could indicate a problem with the header row.`);
    } else if (DEBUG_MODE && actualHeaderCount > 0) {
      // Logger.log(`[LEADS_SHEET_UTIL DEBUG] Header map for "${sheetName}": ${JSON.stringify(headerMap)}`);
    }
    return { sheet: sheet, headerMap: headerMap };

  } catch (e) {
    Logger.log(`[LEADS_SHEET_UTIL ERROR] Error in getSheetAndHeaderMapping_forLeads. SS_ID: ${ssId}, SheetName: ${sheetName}. Error: ${e.toString()}\nStack: ${e.stack}`);
    return { sheet: null, headerMap: {} };
  }
}

/**
 * Retrieves a set of all unique email IDs from the "Source Email ID" column of the leads sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The "Potential Job Leads" sheet object.
 * @param {Object} headerMap An object mapping header names to column numbers for the.
 * @return {Set<string>} A set of processed email IDs.
 */
function getProcessedEmailIdsFromSheet_forLeads(sheet, headerMap) {
  const ids = new Set();
  const emailIdColHeader = "Source Email ID"; // This must match one of the LEADS_SHEET_HEADERS

  if (!sheet || typeof sheet.getName !== 'function') {
    Logger.log('[LEADS_SHEET_UTIL WARN] getProcessedEmailIds: Invalid sheet object provided.');
    return ids;
  }
  if (!headerMap || Object.keys(headerMap).length === 0 || !headerMap[emailIdColHeader]) {
    Logger.log(`[LEADS_SHEET_UTIL WARN] getProcessedEmailIds: "${emailIdColHeader}" not found in headerMap for sheet "${sheet.getName()}" or headerMap is empty. HeaderMap: ${JSON.stringify(headerMap)}`);
    return ids;
  }

  const emailIdColNum = headerMap[emailIdColHeader];
  const lastR = sheet.getLastRow();
  if (lastR < 2) { // No data rows (assuming row 1 is header)
    if (DEBUG_MODE) Logger.log(`[LEADS_SHEET_UTIL DEBUG] getProcessedEmailIds: No data rows in sheet "${sheet.getName()}".`);
    return ids;
  }

  try {
    const rangeToRead = sheet.getRange(2, emailIdColNum, lastR - 1, 1);
    const emailIdValues = rangeToRead.getValues();
    emailIdValues.forEach(row => {
      if (row[0] && row[0].toString().trim() !== "") {
        ids.add(row[0].toString().trim());
      }
    });
    if (DEBUG_MODE) Logger.log(`[LEADS_SHEET_UTIL DEBUG] getProcessedEmailIds: Found ${ids.size} unique email IDs in sheet "${sheet.getName()}".`);
  } catch (e) {
    Logger.log(`[LEADS_SHEET_UTIL ERROR] getProcessedEmailIds: Error reading email IDs from column ${emailIdColNum} in sheet "${sheet.getName()}": ${e.toString()}`);
  }
  return ids;
}

/**
 * Appends a new row with job lead data to the "Potential Job Leads" sheet.
 * Uses LEADS_SHEET_HEADERS from Config.gs as the definitive order and set of columns.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The "Potential Job Leads" sheet object.
 * @param {Object} jobData An object containing the job lead details. Keys should match expected data points.
 * @param {Object} headerMap An object mapping header names to column numbers (used for validation, not direct writing order).
 */
function writeJobDataToSheet_forLeads(sheet, jobData, headerMap) {
  if (!sheet || !jobData || !headerMap || Object.keys(headerMap).length === 0) { /* ... error log ... */ return; }

  const newRow = new Array(LEADS_SHEET_HEADERS.length).fill(""); // From Config.gs

  LEADS_SHEET_HEADERS.forEach((headerName, index) => {
    let value = "";
    switch (headerName) {
        case "Date Added":            value = jobData.dateAdded instanceof Date ? jobData.dateAdded : new Date(); break;
        case "Job Title":             value = jobData.jobTitle || "N/A"; break;
        case "Company":               value = jobData.company || "N/A"; break;
        case "Location":              value = jobData.location || "N/A"; break;
        case "Source":                value = jobData.source || (jobData.sourceEmailSubject ? "Email Subject" : "N/A"); break; // <<< MAPS jobData.source, fallback to indicate subject
        case "Job URL":               value = jobData.jobUrl || "N/A"; break; // <<< MAPS jobData.jobUrl
        case "Status":                value = jobData.status || "New"; break;
        case "Notes":                 value = jobData.notes || ""; break; // <<< MAPS jobData.notes
        case "Applied Date":          value = jobData.appliedDate || ""; break; // Will be blank unless Gemini somehow infers
        case "Follow-up Date":        value = jobData.followUpDate || ""; break; // Will be blank
        case "Source Email ID":       value = jobData.sourceEmailId || ""; break;
        case "Processed Timestamp":   value = jobData.processedTimestamp instanceof Date ? jobData.processedTimestamp : new Date(); break;
        // Added for completeness if sourceEmailSubject is its own column:
        case "Source Email Subject":  value = jobData.sourceEmailSubject || ""; break;
        default:
            if (DEBUG_MODE && !Object.values(jobData).includes(headerName)) { // Avoid logging for standard jobData keys
                 // Logger.log(`[LEADS_SHEET_UTIL WARN] writeJobData: Unhandled header "${headerName}". No direct map from jobData.`);
            }
    }
    newRow[index] = value;
  });

  try {
    sheet.appendRow(newRow);
    // Logger.log ... (reduced logging)
  } catch (e) { /* ... error log ... */ }
}

