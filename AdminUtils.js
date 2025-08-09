// START SNIPPET: Replace the entire content of Config.js with this
/**
 * @file Contains all global configuration constants for FundingFlock.AI.
 * @version 3.0
 */

// --- Core Application Identifiers & Names ---
const APP_NAME = "FundingFlock.AI Grant Tracker";
const CUSTOM_MENU_NAME = "⚙️ FundingFlock.AI Tools";
const TARGET_SPREADSHEET_FILENAME = "FundingFlock.AI Grant Data";

// --- TEMPLATE IDs (CRITICAL - From your existing project) ---
const SPREADSHEET_ID_KEY = 'spreadsheetId';
const TEMPLATE_SHEET_ID = "1WCBvaSUERdZxwabmfgP9aNpW7n3PVk56mGt2ip98cDA";
const MASTER_SCRIPT_ID = "1piNzu4bJCOHdPXeaPvQtHmF93v_EuVut784u8VWiNgG5uT99BMsEMCXr";

// --- Sheet Tab Names ---
const PROPOSAL_TRACKER_SHEET_TAB_NAME = "Proposals";
const OPPORTUNITIES_SHEET_TAB_NAME = "Opportunities"; // For future implementation
const DASHBOARD_TAB_NAME = "Dashboard";
const HELPER_SHEET_NAME = "DashboardHelperData";

// --- Column Configuration for "Proposals" Sheet ---
const PROPOSAL_TRACKER_SHEET_HEADERS = [
  "Processed Timestamp", "Submission Date", "Funder", "RFP Title", "Status",
  "Peak Status", "Last Update", "Amount Requested", "Amount Awarded",
  "Source Email", "Email Link", "Email ID", "Notes"
];
// Column Index Variables (1-based)
const PROP_PROC_TS_COL = 1;
const PROP_SUBMIT_DATE_COL = 2;
const PROP_FUNDER_COL = 3;
const PROP_TITLE_COL = 4;
const PROP_STATUS_COL = 5;
const PROP_PEAK_STATUS_COL = 6;
const PROP_LAST_UPDATE_COL = 7;
const PROP_AMT_REQ_COL = 8;
const PROP_AMT_AWARD_COL = 9;
const PROP_EMAIL_SUBJ_COL = 10;
const PROP_EMAIL_LINK_COL = 11;
const PROP_EMAIL_ID_COL = 12;
const PROP_NOTES_COL = 13;
const TOTAL_COLUMNS_IN_PROPOSAL_SHEET = PROPOSAL_TRACKER_SHEET_HEADERS.length;

// --- Proposal Status Configuration ---
const STATUS_DRAFTING = "Drafting";
const STATUS_SUBMITTED = "Submitted";
const STATUS_UNDER_REVIEW = "Under Review";
const STATUS_AWARDED = "Awarded";
const STATUS_DECLINED = "Declined";
const STATUS_WITHDRAWN = "Withdrawn";
const MANUAL_REVIEW_NEEDED = "Manual Review Needed";

const STATUS_HIERARCHY = {
  [STATUS_DRAFTING]: 1,
  [STATUS_SUBMITTED]: 2,
  [STATUS_UNDER_REVIEW]: 3,
  [STATUS_AWARDED]: 5,
  [STATUS_DECLINED]: 0,
  [STATUS_WITHDRAWN]: -1,
  [MANUAL_REVIEW_NEEDED]: -2
};

const FINAL_STATUSES_FOR_STALE_CHECK = new Set([
  STATUS_AWARDED, STATUS_DECLINED, STATUS_WITHDRAWN, MANUAL_REVIEW_NEEDED
]);
const WEEKS_THRESHOLD = 35; // ~8 months. Grant cycles are long.

// --- Keyword Matching for Status Parsing ---
const AWARDED_KEYWORDS = ["awarded", "pleased to award", "grant has been approved", "funding is approved"];
const UNDER_REVIEW_KEYWORDS = ["under review", "application is being reviewed", "thank you for your submission", "proposal received"];
const DECLINED_KEYWORDS = ["declined", "not selected for funding", "unable to fund", "regret to inform"];

// --- Gmail Configuration ---
const MASTER_GMAIL_LABEL_PARENT = "FundingFlock.AI";
const TRACKER_GMAIL_LABEL_PARENT = `${MASTER_GMAIL_LABEL_PARENT}/Proposals`;
const TRACKER_GMAIL_LABEL_TO_PROCESS = `${TRACKER_GMAIL_LABEL_PARENT}/To Process`;
const TRACKER_GMAIL_LABEL_PROCESSED = `${TRACKER_GMAIL_LABEL_PARENT}/Processed`;
const TRACKER_GMAIL_LABEL_MANUAL_REVIEW = `${TRACKER_GMAIL_LABEL_PARENT}/Manual Review`;

const TRACKER_GMAIL_FILTER_QUERY = `(subject:("your proposal" OR "your application" OR "grant status" OR "funding decision" OR "letter of inquiry" OR "LOI status") OR from:(*foundation*.org OR *@philanthropy.com)) AND -subject:("newsletter" OR "webinar" OR "annual report") AND -list:(newsletter) AND -label:(${TRACKER_GMAIL_LABEL_PROCESSED}) AND -label:(${TRACKER_GMAIL_LABEL_MANUAL_REVIEW})`;

// --- Gemini API Configuration ---
const GEMINI_API_KEY_PROPERTY = 'FUNDINGFLOCK_GEMINI_API_KEY'; // Use a unique property name for safety
const GEMINI_API_ENDPOINT_TEXT_ONLY = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent";

const GEMINI_SYSTEM_INSTRUCTION_PROPOSAL_TRACKER = `
You are an expert assistant parsing emails related to grant proposals for a non-profit.
Your goal is to extract the Funder Name, RFP Title, and Submission Status.
- For "funderName": Extract the name of the funding organization or foundation. If not found, output "N/A".
- For "proposalTitle": Extract the specific title of the grant or RFP. If not found, output "N/A".
- For "submissionStatus": Determine the status. You MUST choose ONLY from this list: "${STATUS_SUBMITTED}", "${STATUS_UNDER_REVIEW}", "${STATUS_AWARDED}", "${STATUS_DECLINED}". If unclear, output "${STATUS_UNDER_REVIEW}".
Output ONLY a valid JSON object with keys "funderName", "proposalTitle", and "submissionStatus".
Example: {"funderName": "The Civic Progress Foundation", "proposalTitle": "Youth Arts & STEM Initiative", "submissionStatus": "Under Review"}
If the email is clearly not a grant proposal update, output: {"funderName": "N/A", "proposalTitle": "N/A", "submissionStatus": "Not a Proposal Update"}
`;

// --- BRANDING COLORS (New FundingFlock Palette) ---
const BRAND_COLORS = {
  // Mapping new brand colors to the existing system keys
  LAPIS_LAZULI: "#96616B",      // Rose Taupe (Primary Professional Color)
  CAROLINA_BLUE: "#37505C",     // Charcoal (Secondary Cool Color)
  CHARCOAL: "#113537",          // Midnight Green (Darkest for Text/BG)
  HUNYADI_YELLOW: "#F76F8E",    // Bright Pink (Primary Accent)
  PALE_ORANGE: "#FFEAD0",        // Champagne (Light Accent/Background)

  // Standard UI Colors
  WHITE: "#FFFFFF",
  PALE_GREY: "#F9E4CB",          // Light Champagne for banding
  MEDIUM_GREY_BORDER: "#E6D8C6"   // Champagne-based border color
};

// --- Development & Debugging ---
const DEBUG_MODE = false; // Set to true for verbose logging during development
// END SNIPPET
