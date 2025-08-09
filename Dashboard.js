/**
 * @file Manages the Dashboard and DashboardHelperData sheets for FundingFlock.AI.
 * This is the definitive version synthesizing the 2x4 scorecard grid with correct chart sizing and placement.
 * @version 8.0 (Final Corrected)
 */

// These three helper functions are all required.
function getOrCreateDashboardSheet(s) {
    if (!s) return null;
    let sheet = s.getSheetByName(DASHBOARD_TAB_NAME);
    if (!sheet) { sheet = s.insertSheet(DASHBOARD_TAB_NAME, 0); }
    sheet.setTabColor(BRAND_COLORS.CAROLINA_BLUE);
    return sheet;
}

function getOrCreateHelperSheet(s) {
    if (!s) return null;
    let sheet = s.getSheetByName(HELPER_SHEET_NAME);
    if (!sheet) { sheet = s.insertSheet(HELPER_SHEET_NAME); }
    return sheet;
}

function _columnToLetter_DashboardLocal(column) {
    let temp, letter = '';
    while (column > 0) {
        temp = (column - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp - 1) / 26;
    }
    return letter;
}

// =========================================================================
// DEFINITIVE DASHBOARD AND CHART CREATION LOGIC
// =========================================================================

/**
 * Formats the dashboard sheet with the 2x4 scorecard grid and chart placeholders.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 */
function formatDashboardSheet(sheet) {
    if (!sheet) return;
    Logger.log(`[formatDashboardSheet] Starting definitive 2x4 grid formatting.`);
    sheet.clear();
    sheet.setHiddenGridlines(true);

    // --- Define Layout Constants ---
    const MAIN_TITLE_BG = BRAND_COLORS.LAPIS_LAZULI, HEADER_TEXT_COLOR = BRAND_COLORS.WHITE;
    const CARD_BG = BRAND_COLORS.PALE_ORANGE, CARD_TEXT_COLOR = BRAND_COLORS.CHARCOAL;
    const CARD_BORDER_COLOR = BRAND_COLORS.MEDIUM_GREY_BORDER;
    const LABEL_FONT_WEIGHT = "bold";

    // --- Main Title & Headers ---
    sheet.getRange("A1:M1").merge().setValue("FundingFlock.AI Grant Proposal Dashboard")
        .setBackground(MAIN_TITLE_BG).setFontColor(HEADER_TEXT_COLOR).setFontSize(18).setFontWeight(LABEL_FONT_WEIGHT)
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange("B3").setValue("Key Metrics Overview:").setFontSize(14).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR);

    // --- Get Sheet & Column References for Formulas ---
    const proposalsRef = `'${PROPOSAL_TRACKER_SHEET_TAB_NAME}'!`;
    const funderColRef = `${proposalsRef}${_columnToLetter_DashboardLocal(PROP_FUNDER_COL)}2:${_columnToLetter_DashboardLocal(PROP_FUNDER_COL)}`;
    const statusColRef = `${proposalsRef}${_columnToLetter_DashboardLocal(PROP_STATUS_COL)}2:${_columnToLetter_DashboardLocal(PROP_STATUS_COL)}`;
    const peakStatusColRef = `${proposalsRef}${_columnToLetter_DashboardLocal(PROP_PEAK_STATUS_COL)}2:${_columnToLetter_DashboardLocal(PROP_PEAK_STATUS_COL)}`;
    const amtReqColRef = `${proposalsRef}${_columnToLetter_DashboardLocal(PROP_AMT_REQ_COL)}2:${_columnToLetter_DashboardLocal(PROP_AMT_REQ_COL)}`;
    const amtAwdColRef = `${proposalsRef}${_columnToLetter_DashboardLocal(PROP_AMT_AWARD_COL)}2:${_columnToLetter_DashboardLocal(PROP_AMT_AWARD_COL)}`;

    // --- Scorecard Creation and Formatting (2x4 Grid) ---
    // Row 1
    _createScorecard(sheet, "B5", "C5", "Total Proposals", `=IFERROR(COUNTA(${funderColRef}), 0)`, "0");
    _createScorecard(sheet, "E5", "F5", "Total Requested", `=IFERROR(SUM(${amtReqColRef}), 0)`, "$#,##0");
    _createScorecard(sheet, "H5", "I5", "Award Rate (#)", `=IFERROR(F7/C5, 0)`, "0.00%");
    _createScorecard(sheet, "K5", "L5", "Award Rate ($)", `=IFERROR(SUMIF(${statusColRef}, "${STATUS_AWARDED}", ${amtAwdColRef})/SUM(${amtReqColRef}), 0)`, "0.00%");
    
    // Row 2
    _createScorecard(sheet, "B7", "C7", "Pending Proposals", `=IFERROR(COUNTIFS(${funderColRef}, "<>", ${statusColRef}, "<>${STATUS_AWARDED}", ${statusColRef}, "<>${STATUS_DECLINED}", ${statusColRef}, "<>${STATUS_WITHDRAWN}"), 0)`, "0");
    _createScorecard(sheet, "E7", "F7", "Total Awarded (#)", `=IFERROR(COUNTIF(${peakStatusColRef},"${STATUS_AWARDED}"), 0)`, "0");
    _createScorecard(sheet, "H7", "I7", "Under Review (#)", `=IFERROR(COUNTIF(${statusColRef},"${STATUS_UNDER_REVIEW}"), 0)`, "0");
    _createScorecard(sheet, "K7", "L7", "Total Awarded ($)", `=IFERROR(SUM(${amtAwdColRef}), 0)`, "$#,##0");

    // --- Chart Section Titles ---
    sheet.getRange("B11").setValue("Funding Funnel").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR);
    sheet.getRange("G11").setValue("Monthly Activity").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(CARD_TEXT_COLOR);

    // --- Layout Sizing and Cleanup ---
    sheet.setColumnWidths(1, 13, 20); // Reset all to a base
    [2,5,8,11].forEach(c => sheet.setColumnWidth(c, 150)); // Labels
    [3,6,9,12].forEach(c => sheet.setColumnWidth(c, 75));  // Values
    [4,7,10].forEach(c => sheet.setColumnWidth(c, 15));   // Spacers
    
    sheet.setRowHeights(1, 45, 30); // Reset all to a base
    [5,7].forEach(r => sheet.setRowHeight(r, 40));
    [2,4,6,8,9,10,12].forEach(r => sheet.setRowHeight(r, 15));
    sheet.setRowHeight(1, 45); sheet.setRowHeight(3, 30); sheet.setRowHeight(11, 25);
    
    if (sheet.getMaxColumns() > 13) { sheet.deleteColumns(14, sheet.getMaxColumns() - 13); }
}

/**
 * Helper to create and format a single scorecard from the 2x4 grid.
 * @private
 */
function _createScorecard(sheet, labelCell, valueCell, labelText, formulaText, numberFormat) {
    sheet.getRange(labelCell).setValue(labelText).setFontWeight("bold").setBackground(BRAND_COLORS.PALE_ORANGE)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(valueCell).setFormula(formulaText).setFontWeight("bold").setFontSize(15).setBackground(BRAND_COLORS.PALE_ORANGE)
      .setNumberFormat(numberFormat).setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(labelCell + ":" + valueCell).setBorder(true, true, true, true, true, true, BRAND_COLORS.MEDIUM_GREY_BORDER, SpreadsheetApp.BorderStyle.SOLID_THIN);
}

/**
 * Sets up all necessary formulas in the hidden helper sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The helper sheet object.
 */
function setupHelperSheetFormulas(sheet) {
    if (!sheet) return;
    sheet.clear();
    const proposalsRef = `'${PROPOSAL_TRACKER_SHEET_TAB_NAME}'!`;
    const funderColLetter = _columnToLetter_DashboardLocal(PROP_FUNDER_COL);
    const statusColLetter = _columnToLetter_DashboardLocal(PROP_STATUS_COL);
    const reqColLetter = _columnToLetter_DashboardLocal(PROP_AMT_REQ_COL);
    const awdColLetter = _columnToLetter_DashboardLocal(PROP_AMT_AWARD_COL);
    const dateColLetter = _columnToLetter_DashboardLocal(PROP_SUBMIT_DATE_COL);

    sheet.getRange("A1:I1").setValues([["Status", "Count", "", "Month", "Count", "", "Funder", "Requested", "Awarded"]]).setFontWeight('bold');
    sheet.getRange("A2").setFormula(`=IFERROR(QUERY(${proposalsRef}${statusColLetter}2:${statusColLetter}, "SELECT Col1, COUNT(Col1) WHERE Col1 IS NOT NULL GROUP BY Col1 LABEL Col1 '', COUNT(Col1) ''"), {"No Data",0})`);
    // This new formula correctly ignores blank rows by checking Col2 instead of Col1.
    sheet.getRange("D2").setFormula(`=IFERROR(QUERY({ARRAYFORMULA(EOMONTH(${proposalsRef}${dateColLetter}2:${dateColLetter}, 0)), ${proposalsRef}${dateColLetter}2:${dateColLetter}}, "SELECT Col1, COUNT(Col2) WHERE Col2 IS NOT NULL GROUP BY Col1 ORDER BY Col1 ASC LABEL Col1 '', COUNT(Col2) ''"), {"No Data",0})`);
    sheet.getRange("D2:D").setNumberFormat("MMM yyyy");
    sheet.getRange("G2").setFormula(`=IFERROR(QUERY(${proposalsRef}${funderColLetter}2:${awdColLetter}, "SELECT Col1, SUM(Col${PROP_AMT_REQ_COL - PROP_FUNDER_COL + 1}), SUM(Col${PROP_AMT_AWARD_COL - PROP_FUNDER_COL + 1}) WHERE Col1 IS NOT NULL GROUP BY Col1 ORDER BY SUM(Col${PROP_AMT_AWARD_COL - PROP_FUNDER_COL + 1}) DESC, SUM(Col${PROP_AMT_REQ_COL - PROP_FUNDER_COL + 1}) DESC LABEL Col1 '', SUM(Col${PROP_AMT_REQ_COL - PROP_FUNDER_COL + 1}) '', SUM(Col${PROP_AMT_AWARD_COL - PROP_FUNDER_COL + 1}) ''"), {"No Data",0,0})`);
}

/**
 * Removes old charts and creates new, correctly positioned charts.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet
 */
function updateDashboardMetrics(dashboardSheet, helperSheet) {
    if (!dashboardSheet || !helperSheet) return;
    SpreadsheetApp.flush();
    Utilities.sleep(2000);
    
    dashboardSheet.getCharts().forEach(c => dashboardSheet.removeChart(c));
    
    // Create new charts with correct positions and dimensions
    updateStatusDistributionChart(dashboardSheet, helperSheet);
    updateSubmissionsOverTimeChart(dashboardSheet, helperSheet);
    updateFundingByFunderChart(dashboardSheet, helperSheet);
}

const BRAND_COLORS_CHART_ARRAY = [
    BRAND_COLORS.LAPIS_LAZULI, BRAND_COLORS.HUNYADI_YELLOW, BRAND_COLORS.CAROLINA_BLUE,
    BRAND_COLORS.PALE_ORANGE, BRAND_COLORS.CHARCOAL
];

function updateStatusDistributionChart(dashboardSheet, helperSheet) {
    const chart = dashboardSheet.newChart().setChartType(Charts.ChartType.PIE).addRange(helperSheet.getRange("A2:B"))
        .setOption('title', 'Proposal Status Distribution').setOption('pieHole', 0.4)
        .setOption('colors', BRAND_COLORS_CHART_ARRAY).setOption('legend', { position: 'right' })
        .setOption('width', 465).setOption('height', 300) // Explicit Sizing
        .setPosition(13, 2, 0, 0).build(); // ANCHOR: Row 13, Col 2 (B)
    dashboardSheet.insertChart(chart);
}

function updateSubmissionsOverTimeChart(dashboardSheet, helperSheet) {
    const chart = dashboardSheet.newChart().setChartType(Charts.ChartType.COLUMN).addRange(helperSheet.getRange("D2:E"))
        .setOption('title', 'Monthly Submissions').setOption('hAxis', { title: 'Month' })
        .setOption('vAxis', { title: 'Count', viewWindow: { min: 0 }})
        .setOption('colors', [BRAND_COLORS.CAROLINA_BLUE]).setOption('legend', { position: 'none' })
        .setOption('width', 465).setOption('height', 300) // Explicit Sizing
        .setPosition(13, 7, 0, 0).build(); // ANCHOR: Row 13, Col 7 (G)
    dashboardSheet.insertChart(chart);
}

function updateFundingByFunderChart(dashboardSheet, helperSheet) {
    const chart = dashboardSheet.newChart().setChartType(Charts.ChartType.BAR)
        .addRange(helperSheet.getRange("G2:G")).addRange(helperSheet.getRange("I2:I")) // Funder vs. Awarded
        .setOption('title', 'Funding Awarded by Funder').setOption('hAxis', { title: 'Amount Awarded ($)' })
        .setOption('vAxis', { title: 'Funder' }).setOption('colors', [BRAND_COLORS.HUNYADI_YELLOW])
        .setOption('legend', { position: 'none' })
        .setOption('width', 938).setOption('height', 300) // Explicit Sizing (Wider Chart)
        .setPosition(24, 2, 0, 0).build(); // ANCHOR: Row 24, Col 2 (B) - Below other charts
    dashboardSheet.insertChart(chart);
}
