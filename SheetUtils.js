
/**
 * @file Contains utility functions for Google Sheets interaction,
 * including sheet creation, formatting, and data access setup.
 */

/**
 * Converts a 1-based column index to its letter representation (e.g., 1 -> A, 27 -> AA).
 * @param {number} column The 1-based column index.
 * @returns {string} The column letter(s).
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Sets up basic formatting for a given sheet: headers, frozen row, column widths, and optional banding.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to format.
 * @param {string[]} headersArray An array of strings for the header row.
 * @param {{col: number, width: number}[]} [columnWidthsArray] Optional. Array of objects specifying column index (1-based) and width.
 * @param {boolean} [applyBandingFlag] Optional. True to apply row banding. Defaults to false.
 * @param {GoogleAppsScript.Spreadsheet.BandingTheme} [bandingThemeEnum] Optional. The banding theme to apply.
 * @returns {boolean} True if formatting was successful, false otherwise.
 */
function setupSheetFormatting(sheet, headersArray, columnWidthsArray, applyBandingFlag, bandingThemeEnum) {
  const FUNC_NAME = "setupSheetFormatting";

  if (!sheet || typeof sheet.getName !== 'function') {
    Logger.log(`[${FUNC_NAME} ERROR] Invalid sheet object passed. Sheet: ${sheet}`);
    return false;
  }
  
  let effectiveHeaderCount = 0;
  if (headersArray && Array.isArray(headersArray) && headersArray.length > 0) {
      effectiveHeaderCount = headersArray.length;
  } else if (sheet.getLastColumn() > 0) {
      effectiveHeaderCount = sheet.getLastColumn();
      Logger.log(`[${FUNC_NAME} WARN] No headersArray for "${sheet.getName()}". Using lastCol (${effectiveHeaderCount}) for effective header count.`);
  } else {
      Logger.log(`[${FUNC_NAME} WARN] Sheet "${sheet.getName()}" empty and no headers. Formatting limited.`);
      // If headersArray is strictly required for other steps, you might return false here.
      // For now, allow it to try setting frozen rows at least.
  }

  Logger.log(`[${FUNC_NAME} INFO] Applying formatting to sheet: "${sheet.getName()}". Effective headers count: ${effectiveHeaderCount}`);
  
  try {
    // --- 0. Clear existing formats on a substantial range to prevent conflicts ---
    // This needs to happen BEFORE setting new headers if they might have old formats.
    const rowsToClearFormat = Math.min(sheet.getMaxRows(), Math.max(200, sheet.getLastRow() > 1 ? sheet.getLastRow() : 200) + 50);
    if (rowsToClearFormat > 0 && effectiveHeaderCount > 0) {
      sheet.getRange(1, 1, rowsToClearFormat, effectiveHeaderCount).clearFormat().removeCheckboxes();
      Logger.log(`[${FUNC_NAME} INFO] Cleared formats on range 1:${rowsToClearFormat}, cols 1:${effectiveHeaderCount} on "${sheet.getName()}".`);
    } else if (rowsToClearFormat === 0 && effectiveHeaderCount === 0 && sheet.getMaxRows() > 0 && sheet.getMaxColumns() > 0) {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearFormat().removeCheckboxes();
      Logger.log(`[${FUNC_NAME} INFO] Sheet "${sheet.getName()}" empty, cleared all formats.`);
    }

    // --- 1. Set Headers (only if headersArray is valid) ---
    if (headersArray && Array.isArray(headersArray) && headersArray.length > 0) {
      const headerRange = sheet.getRange(1, 1, 1, headersArray.length);
      headerRange.setValues([headersArray]);
      headerRange.setFontWeight('bold')
                 .setHorizontalAlignment('center')
                 .setVerticalAlignment('middle')
                 .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      sheet.setRowHeight(1, 45); 
      Logger.log(`[${FUNC_NAME} INFO] Headers set for "${sheet.getName()}".`);
    } else {
        Logger.log(`[${FUNC_NAME} INFO] No headersArray provided, skipping explicit header setting for "${sheet.getName()}".`);
    }

    // --- 2. Freeze Top Row ---
    if (sheet.getFrozenRows() !== 1) {
      try { sheet.setFrozenRows(1); } catch(e) { Logger.log(`[${FUNC_NAME} WARN] Could not set frozen rows on "${sheet.getName()}": ${e.message}`);}
    }

    // --- 3. Set Column Widths ---
    if (columnWidthsArray && Array.isArray(columnWidthsArray) && effectiveHeaderCount > 0) {
      columnWidthsArray.forEach(cw => {
        try {
          if (cw.col > 0 && cw.width > 0 && cw.col <= effectiveHeaderCount) {
            sheet.setColumnWidth(cw.col, cw.width);
          }
        } catch (e) { Logger.log(`[${FUNC_NAME} WARN] Error setting width for col ${cw.col} on "${sheet.getName()}": ${e.message}`); }
      });
      Logger.log(`[${FUNC_NAME} INFO] Column widths applied to "${sheet.getName()}".`);
    }

    // --- 4. General Data Area Formatting ---
    if (effectiveHeaderCount > 0) {
        const lastDataRowForFormat = sheet.getLastRow();
        const numRowsToFormatDataArea = Math.min(sheet.getMaxRows() -1, Math.max(199, lastDataRowForFormat > 1 ? lastDataRowForFormat -1 : 199)); // Format data rows starting from row 2
        if (numRowsToFormatDataArea > 0) { // Check if there's at least one data row to format
            const dataRange = sheet.getRange(2, 1, numRowsToFormatDataArea, effectiveHeaderCount);
            try {
                dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top');
            } catch(e) {Logger.log(`[${FUNC_NAME} WARN] Error applying data area wrap/align on "${sheet.getName()}": ${e.message}`);}
        }
    }
    
    // --- 5. Apply Banding (if requested) ---
    if (applyBandingFlag && effectiveHeaderCount > 0) {
      const lastPopulatedRow = sheet.getLastRow(); // Get current last row with any content
      const defaultVisualRowCountForBanding = 200; // How many rows to show banded if sheet is new/empty

      // Determine how many rows to include in the banding range
      // If sheet has more content than default, use that. Otherwise, use default.
      // Always ensure it's at least 2 (header + 1 data row) for applyRowBanding to have context.
      let bandingTotalRows;
      if (lastPopulatedRow <= 1) { // Sheet is empty or only has header
        bandingTotalRows = defaultVisualRowCountForBanding;
      } else {
        bandingTotalRows = Math.max(lastPopulatedRow, defaultVisualRowCountForBanding);
      }
      bandingTotalRows = Math.max(2, bandingTotalRows); // Ensure at least 2 rows
      bandingTotalRows = Math.min(bandingTotalRows, sheet.getMaxRows()); // Don't exceed max sheet rows

      const rangeForBanding = sheet.getRange(1, 1, bandingTotalRows, effectiveHeaderCount);
      
      try {
        // It's already been cleared comprehensively above, so direct apply.
        // No, still clear bandings directly on THE sheet object before applying to a range
        const existingBandingsOnSheet = sheet.getBandings();
        for (let i = 0; i < existingBandingsOnSheet.length; i++) {
            existingBandingsOnSheet[i].remove();
        }
        Logger.log(`[${FUNC_NAME} INFO] Cleared ALL existing bandings on sheet "${sheet.getName()}". Attempting new banding on ${rangeForBanding.getA1Notation()}`);
        
        const themeToApply = bandingThemeEnum || SpreadsheetApp.BandingTheme.LIGHT_GREY; 
        const banding = rangeForBanding.applyRowBanding(themeToApply, true, false); // Header row true, footer false
        
        Logger.log(`[${FUNC_NAME} INFO] Banding with theme "${themeToApply.toString()}" applied to sheet "${sheet.getName()}", range: ${rangeForBanding.getA1Notation()}.`);
      } catch (eBanding) {
        Logger.log(`[${FUNC_NAME} WARN] BANDING ATTEMPT FAILED on sheet "${sheet.getName()}": ${eBanding.toString()}. Range: ${rangeForBanding.getA1Notation()}. Theme: ${bandingThemeEnum ? bandingThemeEnum.toString() : 'Default'}`);
      }
    }

    // --- 6. Delete Extra Columns ---
    if (effectiveHeaderCount > 0) {
        const maxSheetCols = sheet.getMaxColumns();
        if (maxSheetCols > effectiveHeaderCount) {
            try { 
                sheet.deleteColumns(effectiveHeaderCount + 1, maxSheetCols - effectiveHeaderCount); 
                Logger.log(`[${FUNC_NAME} INFO] Extra columns ${effectiveHeaderCount + 1}-${maxSheetCols} removed from "${sheet.getName()}".`);
            } catch(e){ Logger.log(`[${FUNC_NAME} WARN] Error removing unused columns on "${sheet.getName()}": ${e.message}`);  }
        }
    }
    return true;
  } catch (err) {
    Logger.log(`[${FUNC_NAME} ERROR] Major error during formatting of sheet "${sheet.getName()}": ${err.toString()}\nStack: ${err.stack}`);
    return false;
  }
}

/**
 * Gets or creates the target spreadsheet.
 * It first checks for a stored spreadsheet ID, then by name, and finally creates a new one if none are found.
 * @returns {{spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet}} An object containing the spreadsheet.
 */
function getOrCreateSpreadsheetAndSheet() {
  let ss = null;
  const FUNC_NAME = "getOrCreateSpreadsheetAndSheet";
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_KEY);
  if (spreadsheetId) {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
      Logger.log(`[${FUNC_NAME} INFO] Opened spreadsheet by stored ID: ${spreadsheetId}`);
      return { spreadsheet: ss };
    } catch (e) {
      Logger.log(`[${FUNC_NAME} ERROR] Failed to open spreadsheet by stored ID: ${spreadsheetId}. Error: ${e.message}. Fallback to other methods.`);
    }
  }
  
  if (!ss) {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getName() === TARGET_SPREADSHEET_FILENAME) { 
        ss = active; Logger.log(`[${FUNC_NAME} INFO] Using Active Spreadsheet: "${ss.getName()}".`);
    } else {
        Logger.log(`[${FUNC_NAME} INFO] Finding/creating by name: "${TARGET_SPREADSHEET_FILENAME}".`);
        try {
          const files = DriveApp.getFilesByName(TARGET_SPREADSHEET_FILENAME);
          if (files.hasNext()) {
            ss = SpreadsheetApp.open(files.next()); Logger.log(`[${FUNC_NAME} INFO] Found existing: "${ss.getName()}".`);
          } else {
            ss = SpreadsheetApp.create(TARGET_SPREADSHEET_FILENAME); Logger.log(`[${FUNC_NAME} INFO] Created new: "${ss.getName()}".`);
          }
        } catch (eDrive) { Logger.log(`[${FUNC_NAME} FATAL] Drive/Open/Create Fail: ${eDrive.message}.`); return { spreadsheet: null }; }
    }
  }
  if (!ss) { Logger.log(`[${FUNC_NAME} FATAL] Spreadsheet object is null.`); }
  return { spreadsheet: ss };
}

/**
 * Writes a formatted error row to a specified sheet.
 * This function centralizes error logging to the spreadsheet, providing a consistent format.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to write the error to.
 * @param {object} errorInfo An object containing error details.
 * @param {string} errorInfo.moduleName The name of the module where the error occurred (e.g., "Application Tracker").
 * @param {string} errorInfo.errorType The type of error (e.g., "Gemini API Error").
 * @param {string} errorInfo.details Specific details about the error.
 * @param {string} errorInfo.messageSubject The subject of the email being processed.
 * @param {string} errorInfo.messageId The ID of the email message.
 */
function _writeErrorToSheet(sheet, errorInfo) {
    const FUNC_NAME = "_writeErrorToSheet";
    try {
        let errorRow = [];
        // Use the moduleName to determine which sheet's headers to use for formatting
        if (errorInfo.moduleName === "Application Tracker") {
            errorRow = new Array(APP_TRACKER_SHEET_HEADERS.length).fill("");
            errorRow[PROCESSED_TIMESTAMP_COL - 1] = new Date();
            errorRow[COMPANY_COL - 1] = `ERROR: ${errorInfo.errorType}`;
            errorRow[JOB_TITLE_COL - 1] = "See Notes";
            errorRow[STATUS_COL - 1] = MANUAL_REVIEW_NEEDED;
            errorRow[NOTES_COL - 1] = String(errorInfo.details).substring(0, 500);
            errorRow[EMAIL_SUBJECT_COL - 1] = errorInfo.messageSubject;
            errorRow[EMAIL_ID_COL - 1] = errorInfo.messageId;
        } else if (errorInfo.moduleName === "Job Leads Tracker") {
            errorRow = new Array(LEADS_SHEET_HEADERS.length).fill("");
            errorRow[LEADS_DATE_ADDED_COL - 1] = new Date();
            errorRow[LEADS_COMPANY_COL - 1] = `ERROR: ${errorInfo.errorType}`;
            errorRow[LEADS_JOB_TITLE_COL - 1] = "See Notes";
            errorRow[LEADS_STATUS_COL - 1] = "Error";
            errorRow[LEADS_NOTES_COL - 1] = String(errorInfo.details).substring(0, 500);
            errorRow[LEADS_EMAIL_SUBJECT_COL - 1] = errorInfo.messageSubject;
            errorRow[LEADS_EMAIL_ID_COL - 1] = errorInfo.messageId;
        } else {
            Logger.log(`[${FUNC_NAME}] WARN: Unknown moduleName "${errorInfo.moduleName}". Cannot format error row.`);
            return;
        }
        sheet.appendRow(errorRow);
        Logger.log(`[${FUNC_NAME}] Successfully wrote error entry to sheet for module: ${errorInfo.moduleName}.`);
    } catch (e) {
        Logger.log(`[${FUNC_NAME}] CRITICAL ERROR: Failed to write an error entry to the sheet. Error: ${e.message}`);
    }
}
