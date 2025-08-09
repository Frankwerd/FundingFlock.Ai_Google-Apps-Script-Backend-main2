/**
 * @file Manages the Dashboard and DashboardHelperData sheets,
 * including chart creation and formula setup for helper data.
 */

/**
 * Converts a column index to its letter representation.
 * @param {number} column The column index (1-based).
 * @returns {string} The column letter.
 * @private
 */
function _columnToLetter_DashboardLocal(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Gets or creates the Dashboard sheet and sets its tab color.
 * Final positioning is handled by `runFullProjectInitialSetup` in `Main.js`.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The spreadsheet object.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null} The dashboard sheet or null if an error occurs.
 */
function getOrCreateDashboardSheet(spreadsheet) {
  const FUNC_NAME = "getOrCreateDashboardSheet";
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') { 
    Logger.log(`[${FUNC_NAME} ERROR] Invalid spreadsheet object provided.`); 
    return null; 
  }
  let dashboardSheet = spreadsheet.getSheetByName(DASHBOARD_TAB_NAME); // From Config.gs
  if (!dashboardSheet) {
    try {
      dashboardSheet = spreadsheet.insertSheet(DASHBOARD_TAB_NAME); 
      Logger.log(`[${FUNC_NAME} INFO] Created new dashboard sheet: "${DASHBOARD_TAB_NAME}".`);
    } catch (eCreate) { 
      Logger.log(`[${FUNC_NAME} ERROR] Failed to create dashboard sheet "${DASHBOARD_TAB_NAME}": ${eCreate.message}`); 
      return null; 
    }
  } else { 
    Logger.log(`[${FUNC_NAME} INFO] Found existing dashboard sheet: "${DASHBOARD_TAB_NAME}".`); 
  }
  
  if (dashboardSheet) {
    try { dashboardSheet.setTabColor(BRAND_COLORS.CAROLINA_BLUE); } // From Config.gs
    catch (eTabColor) { Logger.log(`[${FUNC_NAME} WARN] Failed to set tab color for dashboard: ${eTabColor.message}`); }
  }
  return dashboardSheet;
}

/**
 * Formats the dashboard sheet layout, including titles, scorecards, and charts.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet object.
 * @returns {boolean} True if formatting was successful, false otherwise.
 */
function formatDashboardSheet(dashboardSheet) {
  const FUNC_NAME = "formatDashboardSheet";
  if (!dashboardSheet || typeof dashboardSheet.getName !== 'function') { 
    Logger.log(`[${FUNC_NAME} ERROR] Invalid dashboardSheet object.`); 
    return false; 
  }
  Logger.log(`[${FUNC_NAME} INFO] Starting formatting for dashboard: "${dashboardSheet.getName()}".`);

  try {
    dashboardSheet.clear({contentsOnly: false, formatOnly: false, skipConditionalFormatRules: false}); 
    try { dashboardSheet.setConditionalFormatRules([]); } 
    catch (e) { Logger.log(`[${FUNC_NAME} WARN] Could not clear conditional format rules: ${e.message}`);}
    try { dashboardSheet.setHiddenGridlines(true); } 
    catch (e) { Logger.log(`[${FUNC_NAME} WARN] Error hiding gridlines: ${e.toString()}`); }

    const MAIN_TITLE_BG = BRAND_COLORS.LAPIS_LAZULI; const HEADER_TEXT_COLOR = BRAND_COLORS.WHITE;
    const CARD_BG = BRAND_COLORS.PALE_GREY; const CARD_TEXT_COLOR = BRAND_COLORS.CHARCOAL;
    const CARD_BORDER_COLOR = BRAND_COLORS.MEDIUM_GREY_BORDER;
    const PRIMARY_VALUE_COLOR = BRAND_COLORS.PALE_ORANGE; const SECONDARY_VALUE_COLOR = BRAND_COLORS.HUNYADI_YELLOW;
    const METRIC_FONT_SIZE = 15; const METRIC_FONT_WEIGHT = "bold"; const LABEL_FONT_WEIGHT = "bold";
    const spacerColAWidth = 20; const labelWidth = 150; const valueWidth = 75; const spacerS = 15;

    dashboardSheet.getRange("A1:M1").merge().setValue("CareerSuite.AI Job Application Dashboard")
                  .setBackground(MAIN_TITLE_BG).setFontColor(HEADER_TEXT_COLOR).setFontSize(18).setFontWeight("bold")
                  .setHorizontalAlignment("center").setVerticalAlignment("middle");
    dashboardSheet.setRowHeight(1, 45); dashboardSheet.setRowHeight(2, 10); 

    dashboardSheet.getRange("B3").setValue("Key Metrics Overview:").setFontSize(14).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR);
    dashboardSheet.setRowHeight(3, 30); dashboardSheet.setRowHeight(4, 10);

    const appSheetNameForFormula = `'${APP_TRACKER_SHEET_TAB_NAME}'`;
    const companyColLetter = _columnToLetter_DashboardLocal(COMPANY_COL);
    const jobTitleColLetter = _columnToLetter_DashboardLocal(JOB_TITLE_COL);
    const statusColLetter = _columnToLetter_DashboardLocal(STATUS_COL);
    const peakStatusColLetter = _columnToLetter_DashboardLocal(PEAK_STATUS_COL);

    // Scorecard Setup (Formulas direct to Applications sheet)
    // Row 1
    dashboardSheet.getRange("B5").setValue("Total Apps").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("C5").setFormula(`=IFERROR(COUNTA(${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("E5").setValue("Peak Interviews").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("F5").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${INTERVIEW_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("H5").setValue("Interview Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("I5").setFormula(`=IFERROR(F5/C5, 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%").setFontColor(SECONDARY_VALUE_COLOR);
    dashboardSheet.getRange("K5").setValue("Offer Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("L5").setFormula(`=IFERROR(F7/C5, 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%").setFontColor(SECONDARY_VALUE_COLOR);
    dashboardSheet.setRowHeight(5, 40); dashboardSheet.setRowHeight(6, 10);
    // Row 2
    dashboardSheet.getRange("B7").setValue("Active Apps").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    let activeAppsFormula = `=IFERROR(COUNTIFS(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>"&"", ${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>${REJECTED_STATUS}", ${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>${ACCEPTED_STATUS}"), 0)`;
    dashboardSheet.getRange("C7").setFormula(activeAppsFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("E7").setValue("Peak Offers").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("F7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${OFFER_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("H7").setValue("Current Interviews").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("I7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${INTERVIEW_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("K7").setValue("Current Assessments").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("L7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${ASSESSMENT_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.setRowHeight(7, 40); dashboardSheet.setRowHeight(8, 10);
    // Row 3
    dashboardSheet.getRange("B9").setValue("Total Rejections").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("C9").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("E9").setValue("Apps Viewed (Peak)").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    dashboardSheet.getRange("F9").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${APPLICATION_VIEWED_STATUS}"),0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(PRIMARY_VALUE_COLOR);
    dashboardSheet.getRange("H9").setValue("Manual Review").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    const compColManualFormula = `${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}="${MANUAL_REVIEW_NEEDED}"`;
    const titleColManualFormula = `${appSheetNameForFormula}!${jobTitleColLetter}2:${jobTitleColLetter}="${MANUAL_REVIEW_NEEDED}"`;
    const statusColManualFormula = `${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}="${MANUAL_REVIEW_NEEDED}"`;
    const finalManualReviewFormula = `=IFERROR(SUM(ARRAYFORMULA(SIGN((${compColManualFormula})+(${titleColManualFormula})+(${statusColManualFormula})))),0)`;
    dashboardSheet.getRange("I9").setFormula(finalManualReviewFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0").setFontColor(SECONDARY_VALUE_COLOR);
    dashboardSheet.getRange("K9").setValue("Direct Reject Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR).setVerticalAlignment("middle");
    const directRejectFormula = `=IFERROR(COUNTIFS(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${DEFAULT_STATUS}",${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}")/C5, 0)`;
    dashboardSheet.getRange("L9").setFormula(directRejectFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%").setFontColor(SECONDARY_VALUE_COLOR);
    dashboardSheet.setRowHeight(9, 40); dashboardSheet.setRowHeight(10, 15);

    const scorecardRangesToStyle = ["B5:C5", "E5:F5", "H5:I5", "K5:L5", "B7:C7", "E7:F7", "H7:I7", "K7:L7", "B9:C9", "E9:F9", "H9:I9", "K9:L9"];
    scorecardRangesToStyle.forEach(rangeString => {
      const range = dashboardSheet.getRange(rangeString);
      range.setBackground(CARD_BG).setBorder(true, true, true, true, true, true, CARD_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THIN);
    });
    
    const chartSectionTitleRow1 = 11; dashboardSheet.getRange("B"+chartSectionTitleRow1).setValue("Platform & Weekly Trends").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR);
    dashboardSheet.setRowHeight(chartSectionTitleRow1, 25); dashboardSheet.setRowHeight(chartSectionTitleRow1+1, 5);
    const chartSectionTitleRow2 = 28; dashboardSheet.getRange("B"+chartSectionTitleRow2).setValue("Application Funnel Analysis").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR);
    dashboardSheet.setRowHeight(chartSectionTitleRow2, 25); dashboardSheet.setRowHeight(chartSectionTitleRow2+1, 5);

    dashboardSheet.setColumnWidth(1, spacerColAWidth); dashboardSheet.setColumnWidth(2, labelWidth); dashboardSheet.setColumnWidth(3, valueWidth); dashboardSheet.setColumnWidth(4, spacerS);
    dashboardSheet.setColumnWidth(5, labelWidth); dashboardSheet.setColumnWidth(6, valueWidth); dashboardSheet.setColumnWidth(7, spacerS);
    dashboardSheet.setColumnWidth(8, labelWidth); dashboardSheet.setColumnWidth(9, valueWidth); dashboardSheet.setColumnWidth(10, spacerS);
    dashboardSheet.setColumnWidth(11, labelWidth); dashboardSheet.setColumnWidth(12, valueWidth); dashboardSheet.setColumnWidth(13, spacerColAWidth);
    Logger.log(`[${FUNC_NAME} INFO] Dashboard column widths set.`);

    const lastUsedDataColumnOnDashboard = 13; const maxColsDashboard = dashboardSheet.getMaxColumns();
    if (maxColsDashboard > lastUsedDataColumnOnDashboard) dashboardSheet.deleteColumns(lastUsedDataColumnOnDashboard + 1, maxColsDashboard - lastUsedDataColumnOnDashboard);
    const lastUsedDataRowOnDashboard = 45; const maxRowsDashboard = dashboardSheet.getMaxRows();
    if (maxRowsDashboard > lastUsedDataRowOnDashboard) dashboardSheet.deleteRows(lastUsedDataRowOnDashboard + 1, maxRowsDashboard - lastUsedDataRowOnDashboard);
    
    Logger.log(`[${FUNC_NAME} INFO] Formatting concluded for Dashboard visuals and scorecards.`);
    return true;
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Major dashboard formatting error: ${e.toString()}\nStack: ${e.stack}`);
    return false;
  }
}

/**
 * Gets or creates the helper data sheet.
 * Detailed formatting is handled by `initialSetup_LabelsAndSheet` in `Main.js`.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The spreadsheet object.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet | null} The helper sheet or null if an error occurs.
 */
function getOrCreateHelperSheet(spreadsheet) {
  const FUNC_NAME = "getOrCreateHelperSheet";
  if (!spreadsheet || typeof spreadsheet.getSheetByName !== 'function') { Logger.log(`[${FUNC_NAME} ERROR] Invalid spreadsheet.`); return null;}
  let helperSheet = spreadsheet.getSheetByName(HELPER_SHEET_NAME);
  if (!helperSheet) {
    try { helperSheet = spreadsheet.insertSheet(HELPER_SHEET_NAME); Logger.log(`[${FUNC_NAME} INFO] Created: "${HELPER_SHEET_NAME}".`);}
    catch (eCreate) { Logger.log(`[${FUNC_NAME} ERROR] Create Fail for "${HELPER_SHEET_NAME}": ${eCreate.message}`); return null;}
  } else { Logger.log(`[${FUNC_NAME} INFO] Found: "${HELPER_SHEET_NAME}".`); }
  return helperSheet;
}

/**
 * Sets up the formulas in the `DashboardHelperData` sheet. This is called once during initial setup.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet The `DashboardHelperData` sheet object.
 * @returns {boolean} True if formulas were set successfully, false otherwise.
 */
function setupHelperSheetFormulas(helperSheet) {
  const FUNC_NAME = "setupHelperSheetFormulas";
  if (!helperSheet || typeof helperSheet.getName !== 'function') {
    Logger.log(`[${FUNC_NAME} ERROR] Invalid helperSheet object passed.`);
    return false;
  }
  Logger.log(`[${FUNC_NAME} INFO] Setting up formulas in "${helperSheet.getName()}" based on OLD LOGIC structure.`);

  try {
    // Clear existing content to ensure clean state for new formulas
    const maxRows = helperSheet.getMaxRows();
    const maxCols = helperSheet.getMaxColumns();
    if (maxRows > 0 && maxCols > 0) {
        helperSheet.getRange(1, 1, maxRows, maxCols).clearContent().clearNote();
        Logger.log(`[${FUNC_NAME} INFO] Cleared content from helper sheet "${helperSheet.getName()}".`);
    }

    // --- Define sheet references and column letters needed for formulas ---
    // These constants MUST be available from Config.gs
    const appSheetNameForFormula = `'${APP_TRACKER_SHEET_TAB_NAME}'!`; // e.g., "'Applications'!"
    const platformColLetter = _columnToLetter_DashboardLocal(PLATFORM_COL);
    const emailDateColLetter = _columnToLetter_DashboardLocal(EMAIL_DATE_COL);
    const peakStatusColLetter = _columnToLetter_DashboardLocal(PEAK_STATUS_COL);
    const companyColLetter = _columnToLetter_DashboardLocal(COMPANY_COL); // Used for "Total Applications" in funnel

    // --- 1. Platform Distribution Data (Formulas for Helper Columns A:B) ---
    helperSheet.getRange("A1").setValue("Platform");
    helperSheet.getRange("B1").setValue("Count");
    // This QUERY formula gets unique platforms and their counts from Applications sheet's platform column
    const platformQueryFormula = `=IFERROR(QUERY(${appSheetNameForFormula}${platformColLetter}2:${platformColLetter}, "SELECT Col1, COUNT(Col1) WHERE Col1 IS NOT NULL AND Col1 <> '' GROUP BY Col1 ORDER BY COUNT(Col1) DESC LABEL Col1 '', COUNT(Col1) ''", 0), {"No Platform Data",0})`;
    helperSheet.getRange("A2").setFormula(platformQueryFormula);
    Logger.log(`[${FUNC_NAME} INFO] Platform distribution formula set in Helper A2: ${platformQueryFormula}`);

    // --- 2. Data for Applications Over Time (Weekly) Chart ---
    // Intermediate calculation columns (J & K) for weekly data as per your "OLD LOGIC"
    helperSheet.getRange("J1").setValue("RAW_VALID_DATES_FOR_WEEKLY");
    const rawDatesFormula = `=IFERROR(FILTER(${appSheetNameForFormula}${emailDateColLetter}2:${emailDateColLetter}, ISNUMBER(${appSheetNameForFormula}${emailDateColLetter}2:${emailDateColLetter})), "")`;
    helperSheet.getRange("J2").setFormula(rawDatesFormula);
    helperSheet.getRange("J2:J").setNumberFormat("yyyy-mm-dd hh:mm:ss"); // Original format for raw dates

    helperSheet.getRange("K1").setValue("CALCULATED_WEEK_STARTS (Mon)");
    // Formula for Monday as week start: DATE(YEAR(J2), MONTH(J2), DAY(J2) - WEEKDAY(J2,2) + 1)
    const weekStartCalcFormula = `=ARRAYFORMULA(IF(ISBLANK(J2:J), "", DATE(YEAR(J2:J), MONTH(J2:J), DAY(J2:J) - WEEKDAY(J2:J, 2) + 1)))`;
    helperSheet.getRange("K2").setFormula(weekStartCalcFormula);
    helperSheet.getRange("K2:K").setNumberFormat("yyyy-mm-dd"); // Original format for calculated week starts

    // Final aggregated weekly data for chart (Helper Columns D:E)
    helperSheet.getRange("D1").setValue("Week Starting"); // Header for chart data source
    const uniqueWeeksFormula = `=IFERROR(SORT(UNIQUE(FILTER(K2:K, K2:K<>""))), {"No Date Data"})`; // Get unique week start dates from K
    helperSheet.getRange("D2").setFormula(uniqueWeeksFormula);
    helperSheet.getRange("D2:D").setNumberFormat("M/d/yyyy"); // Chart-friendly date format for axis

    helperSheet.getRange("E1").setValue("Applications"); // Header for chart data source
    // Count applications for each unique week start date
    const weeklyCountsFormula = `=ARRAYFORMULA(IF(D2:D="", "", COUNTIF(K2:K, D2:D)))`;
    helperSheet.getRange("E2").setFormula(weeklyCountsFormula);
    helperSheet.getRange("E2:E").setNumberFormat("0");
    Logger.log(`[${FUNC_NAME} INFO] Weekly applications formulas set in Helper (D2, E2 using intermediate J:K).`);

    // --- 3. Data for Application Funnel (Peak Stages) Chart (Helper Columns G:H) ---
    helperSheet.getRange("G1").setValue("Stage");
    helperSheet.getRange("H1").setValue("Count");
    const funnelStagesValues = [DEFAULT_STATUS, APPLICATION_VIEWED_STATUS, ASSESSMENT_STATUS, INTERVIEW_STATUS, OFFER_STATUS]; // From Config.gs

    // Write stage names to column G
    helperSheet.getRange(2, 7, funnelStagesValues.length, 1).setValues(funnelStagesValues.map(stage => [stage]));
    
    // Set formulas for counts in column H
    // First stage (e.g., "Applied") often represents total applications. Your old logic had this for H2:
    helperSheet.getRange("H2").setFormula(`=IFERROR(COUNTA(${appSheetNameForFormula}${companyColLetter}2:${companyColLetter}),0)`);
    // For subsequent stages, count based on Peak Status matching the stage in column G
    for (let i = 1; i < funnelStagesValues.length; i++) { // Starts from the second stage in your array
      // Example: For row 3 (second stage), formula in H3 refers to G3
      helperSheet.getRange(i + 2, 8).setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}${peakStatusColLetter}2:${peakStatusColLetter}, G${i + 2}),0)`);
    }
    Logger.log(`[${FUNC_NAME} INFO] Funnel stage formulas set in Helper G:H.`);
    
    SpreadsheetApp.flush(); // Ensure formulas calculate initially
    return true;
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Error setting formulas in helper sheet: ${e.toString()}\nStack: ${e.stack}`);
    return false;
  }
}


/**
 * Ensures charts on the Dashboard sheet are created or updated.
 * This function relies on the `DashboardHelperData` sheet being populated by formulas.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The "Dashboard" sheet object.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet The "DashboardHelperData" sheet object.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [applicationsSheet] Optional for context.
 */
function updateDashboardMetrics(dashboardSheet, helperSheet, applicationsSheet) {
  const FUNC_NAME = "updateDashboardMetrics";
  Logger.log(`\n==== ${FUNC_NAME}: STARTING - Verifying/Updating Charts (Helper is Formula-Driven) ====`);

  if (!helperSheet || typeof helperSheet.getName !== 'function') { Logger.log(`[${FUNC_NAME} ERROR] Helper sheet invalid.`); return; }
  if (!dashboardSheet && DEBUG_MODE) { Logger.log(`[${FUNC_NAME} WARN] Dashboard sheet invalid. Charts cannot be updated.`); }
  // applicationsSheet parameter is kept for context, though not directly used by THIS version of updateDashboardMetrics for data population.

  Logger.log(`[${FUNC_NAME} INFO] Helper data is formula-driven. Refreshing charts on "${dashboardSheet ? dashboardSheet.getName() : 'N/A'}"...`);
  
  // Force recalculation of all formulas in the spreadsheet
  SpreadsheetApp.flush(); 
  // Add a small delay to allow complex formulas (especially QUERY) to complete calculation
  Utilities.sleep(2000); // Increased pause to 2 seconds

  if (dashboardSheet && helperSheet) {
     Logger.log(`[${FUNC_NAME} INFO] Calling chart creation/update functions...`);
     try {
        updatePlatformDistributionChart(dashboardSheet, helperSheet);
        updateApplicationsOverTimeChart(dashboardSheet, helperSheet);
        updateApplicationFunnelChart(dashboardSheet, helperSheet);
        Logger.log(`[${FUNC_NAME} INFO] Chart update/creation process successfully called.`);
     } catch (e) { 
        Logger.log(`[${FUNC_NAME} ERROR] Chart update calls threw an error: ${e.toString()}\nStack: ${e.stack}`); 
     }
  } else {
      Logger.log(`[${FUNC_NAME} WARN] Skipping chart object updates - dashboardSheet or helperSheet is missing.`);
  }
  Logger.log(`\n==== ${FUNC_NAME} FINISHED ====`);
}

// --- Dashboard Chart Update Functions ---
// (These are updatePlatformDistributionChart, updateApplicationsOverTimeChart, updateApplicationFunnelChart,
//  and BRAND_COLORS_CHART_ARRAY - use the versions from your latest log where the setOption loop was working,
//  or the version I provided just before with those setOption loops.)
// Ensure their function names here EXACTLY match the calls in updateDashboardMetrics above.

/**
 * Updates the platform distribution pie chart on the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet The helper data sheet.
 */
function updatePlatformDistributionChart(dashboardSheet, helperSheet) {
  const FUNC_NAME = "updatePlatformDistributionChart";
  const CHART_TITLE = "Platform Distribution";
  const ANCHOR_ROW = 13, ANCHOR_COL = 2;
  let existingChart = null;
  dashboardSheet.getCharts().forEach(chart => { 
    if (chart.getOptions().get('title') === CHART_TITLE && 
        chart.getContainerInfo().getAnchorColumn() === ANCHOR_COL && 
        chart.getContainerInfo().getAnchorRow() === ANCHOR_ROW) {
      existingChart = chart;
    }
  });

  if (helperSheet.getRange("A1").getValue().toString().trim() !== "Platform") {
      // ... (logging and chart removal logic if header is wrong) ...
      return;
  }
  const dataRange = helperSheet.getRange("A:B");
  Logger.log(`[${FUNC_NAME} INFO] Using data range ${helperSheet.getName()}!A:B for chart "${CHART_TITLE}".`);

  const optionsObject = { 
    title: CHART_TITLE, 
    // legend: { // Complex object commented out
    //     position: Charts.Position.RIGHT,
    //     textStyle: { color: BRAND_COLORS.CHARCOAL, fontSize: 10 }
    // },
    pieHole: 0.4, 
    width: 480, 
    height: 300, 
    sliceVisibilityThreshold: 0, // Show all slices to maximize chance of legend appearing
    is3D: true, 
    colors: BRAND_COLORS_CHART_ARRAY() 
    // Note: The 'legend' key itself is removed from this optionsObject initially.
    // It will be added via setOption below.
  };
  
  try {
    let chartBuilder;
    if (existingChart) { 
      chartBuilder = existingChart.modify().clearRanges().addRange(dataRange).setChartType(Charts.ChartType.PIE); 
    } else { 
      chartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.PIE).addRange(dataRange); 
    }

    // Apply most options individually
    for (const key in optionsObject) { 
      // The 'legend' key is not in optionsObject here, so this loop won't try to set it yet.
      chartBuilder = chartBuilder.setOption(key, optionsObject[key]); 
    } 
    
    // ** Explicitly set legend position using a simple string **
    chartBuilder = chartBuilder.setOption('legend', 'right'); // 'top', 'bottom', 'left', 'right', 'none', 'labeled' (for pie)

    if (existingChart) {
      dashboardSheet.updateChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    } else {
      dashboardSheet.insertChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    }
    Logger.log(`[${FUNC_NAME} INFO] Chart "${CHART_TITLE}" processed.`);
  } catch (e) { 
    Logger.log(`[${FUNC_NAME} ERROR] Build/insert/update "${CHART_TITLE}": ${e.message}`); 
  }


  try {
    // ... (rest of the chart building logic with individual setOption calls) ...
    let chartBuilder;
    if (existingChart) { chartBuilder = existingChart.modify().clearRanges().addRange(dataRange).setChartType(Charts.ChartType.PIE); }
    else { chartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.PIE).addRange(dataRange); }
    for (const key in optionsObject) { chartBuilder = chartBuilder.setOption(key, optionsObject[key]); }

    if (existingChart) dashboardSheet.updateChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    else dashboardSheet.insertChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    Logger.log(`[${FUNC_NAME} INFO] Chart "${CHART_TITLE}" processed.`);
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Build/insert/update "${CHART_TITLE}": ${e.message}`);
  }
}

/**
 * Updates the applications over time line chart on the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet The helper data sheet.
 */
function updateApplicationsOverTimeChart(dashboardSheet, helperSheet) {
  const FUNC_NAME = "updateApplicationsOverTimeChart";
  const CHART_TITLE = "Applications Over Time (Weekly)";
  const ANCHOR_ROW = 13, ANCHOR_COL = 8;
  let existingChart = null;
  dashboardSheet.getCharts().forEach(chart => { if (chart.getOptions().get('title') === CHART_TITLE && chart.getContainerInfo().getAnchorColumn() === ANCHOR_COL && chart.getContainerInfo().getAnchorRow() === ANCHOR_ROW) existingChart = chart; });

  if (helperSheet.getRange("D1").getValue().toString().trim() !== "Week Starting") {
      Logger.log(`[${FUNC_NAME} WARN] Header "Week Starting" missing in Helper D1. Removing chart: ${CHART_TITLE}`);
      if (existingChart) try { dashboardSheet.removeChart(existingChart); } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Remove chart: ${e.message}`); }
      return;
  }
  // **MODIFIED: Use full column range D:E.**
  const dataRange = helperSheet.getRange("D:E");
  Logger.log(`[${FUNC_NAME} INFO] Using data range ${helperSheet.getName()}!D:E for chart "${CHART_TITLE}".`);
  
  const optionsObject = { title:CHART_TITLE, hAxis:{title:'Week Starting',textStyle:{fontSize:10},format:'M/d', gridlines: {color: '#EEE'}}, vAxis:{title:'Applications',textStyle:{fontSize:10},viewWindow:{min:0},gridlines:{count:-1, color: '#CCC'}}, legend:{position:'none'}, colors:[BRAND_COLORS.LAPIS_LAZULI], width:480, height:300, pointSize: 5, lineWidth: 2 };
  try {
    let chartBuilder;
    if (existingChart) { chartBuilder = existingChart.modify().clearRanges().addRange(dataRange).setChartType(Charts.ChartType.LINE); }
    else { chartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.LINE).addRange(dataRange); }
    for (const key in optionsObject) { chartBuilder = chartBuilder.setOption(key, optionsObject[key]); }

    if (existingChart) dashboardSheet.updateChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    else dashboardSheet.insertChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    Logger.log(`[${FUNC_NAME} INFO] Chart "${CHART_TITLE}" processed.`);
  } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Build/insert/update "${CHART_TITLE}": ${e.message}`); }
}

/**
 * Updates the application funnel column chart on the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet The helper data sheet.
 */
function updateApplicationFunnelChart(dashboardSheet, helperSheet) {
  const FUNC_NAME = "updateApplicationFunnelChart";
  const CHART_TITLE = "Application Funnel (Peak Stages)";
  const ANCHOR_ROW = 30, ANCHOR_COL = 2;
  let existingChart = null;
  dashboardSheet.getCharts().forEach(chart => { if (chart.getOptions().get('title') === CHART_TITLE && chart.getContainerInfo().getAnchorColumn() === ANCHOR_COL && chart.getContainerInfo().getAnchorRow() === ANCHOR_ROW) existingChart = chart; });
  
  if (helperSheet.getRange("G1").getValue().toString().trim() !== "Stage") {
      Logger.log(`[${FUNC_NAME} WARN] Header "Stage" missing in Helper G1. Removing chart: ${CHART_TITLE}`);
      if (existingChart) try { dashboardSheet.removeChart(existingChart); } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Remove chart: ${e.message}`); }
      return;
  }
  // **MODIFIED: Use full column range G:H.**
  const dataRange = helperSheet.getRange("G:H");
  Logger.log(`[${FUNC_NAME} INFO] Using data range ${helperSheet.getName()}!G:H for chart "${CHART_TITLE}".`);

  const optionsObject = { title:CHART_TITLE, hAxis:{title:'Application Stage',textStyle:{fontSize:10},slantedText:true,slantedTextAngle:30}, vAxis:{title:'Applications',textStyle:{fontSize:10},viewWindow:{min:0},gridlines:{count:-1, color: '#CCC'}}, legend:{position:'none'}, colors:[BRAND_COLORS.CAROLINA_BLUE], bar:{groupWidth:'60%'}, width:480, height:300 };
  try {
    let chartBuilder;
    if (existingChart) { chartBuilder = existingChart.modify().clearRanges().addRange(dataRange).setChartType(Charts.ChartType.COLUMN); }
    else { chartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.COLUMN).addRange(dataRange); }
    for (const key in optionsObject) { chartBuilder = chartBuilder.setOption(key, optionsObject[key]); }

    if (existingChart) dashboardSheet.updateChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    else dashboardSheet.insertChart(chartBuilder.setPosition(ANCHOR_ROW, ANCHOR_COL, 0, 0).build());
    Logger.log(`[${FUNC_NAME} INFO] Chart "${CHART_TITLE}" processed.`);
  } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Build/insert/update "${CHART_TITLE}": ${e.message}`); }
}

/**
 * Returns an array of brand colors for charts.
 * @returns {string[]} An array of hex color codes.
 */
function BRAND_COLORS_CHART_ARRAY() {
    return [ BRAND_COLORS.LAPIS_LAZULI, BRAND_COLORS.CAROLINA_BLUE, BRAND_COLORS.HUNYADI_YELLOW, BRAND_COLORS.PALE_ORANGE, BRAND_COLORS.CHARCOAL, "#27AE60", "#8E44AD", "#E67E22", "#16A085", "#C0392B" ];
}
