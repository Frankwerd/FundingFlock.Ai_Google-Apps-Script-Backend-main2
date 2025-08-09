/**
 * @file Contains all global configuration constants and user-facing display names for the project.
 * This allows for easy updates to sheet names, labels, statuses, etc., without deep code changes.
 * @author Assistant
 * @version 1.5
 */

/**
 * @description IMPORTANT: Replace YOUR_MASTER_DEPLOYMENT_ID with the actual deployment ID of your master script Web App
 */
const MASTER_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbyYR4sMfRplUw92UC1MiYad9V_0o3jL0j73Om3JJCrxBTIFxpUvEClMUdEVjkKzdDRK/exec'; // Replace with actual URL

// --- Debugging & Development ---
const DEBUG_MODE = true; // Master switch for all debug logging across modules. Set to false for "production".

// --- Core Application Identifiers & Names ---
const APP_NAME = "CareerSuite.AI Job Tracker";
const CUSTOM_MENU_NAME = "⚙️ CareerSuite.AI Tools"; // Recommended: Use a distinct emoji if desired.
const TARGET_SPREADSHEET_FILENAME = "CareerSuite.AI Data"; // Used by WebApp if creating a new sheet.

// --- TEMPLATE IDs (CRITICAL - MUST BE SET MANUALLY BY THE DEPLOYER) ---
// ID of the Master Template Google Sheet (the one users will copy)
const SPREADSHEET_ID_KEY = 'spreadsheetId';
const TEMPLATE_SHEET_ID = "12jj5lTyu_MzA6KBkfD-30mj-KYHaX-BjouFMtPIIzFc"; // REPLACE WITH YOUR TEMPLATE SHEET ID
// ID of the Master Script Project (bound to the template sheet above)
// This is typically found in Project Settings -> Script ID
const MASTER_SCRIPT_ID = "12suq_wdzxKZy7S7MJ9bB2a2-DxiN_Kl5mUVHupR-YAqT-_54eU-gQB8i"; // REPLACE WITH YOUR MASTER SCRIPT ID


// --- Sheet Tab Names (User-Facing) ---
// For Job Application Tracker
const APP_TRACKER_SHEET_TAB_NAME = "Applications";
const DASHBOARD_TAB_NAME = "Dashboard";
const HELPER_SHEET_NAME = "DashboardHelperData"; // This sheet will be hidden by default
// For Job Leads Tracker
const LEADS_SHEET_TAB_NAME = "Potential Job Leads";


// --- Column Configuration for "Applications" Sheet (APP_TRACKER_SHEET_TAB_NAME) ---
// These headers MUST match the order of columns data is written into.
// Changing order here REQUIRES changing column index variables below AND related parsing/writing logic.
const APP_TRACKER_SHEET_HEADERS = [
  "Processed Timestamp", "Email Date", "Platform", "Company", "Job Title", 
  "Status", "Peak Status", "Last Update Date", "Email Subject", 
  "Email Link", "Email ID", "Notes" 
];
// Column Index Variables (1-based for sheet.getRange(), adjust if header order changes)
const PROCESSED_TIMESTAMP_COL = 1;
const EMAIL_DATE_COL = 2;
const PLATFORM_COL = 3;
const COMPANY_COL = 4;
const JOB_TITLE_COL = 5;
const STATUS_COL = 6;
const PEAK_STATUS_COL = 7; // Hidden by default after setup
const LAST_UPDATE_DATE_COL = 8;
const EMAIL_SUBJECT_COL = 9;
const EMAIL_LINK_COL = 10;
const EMAIL_ID_COL = 11;
const NOTES_COL = 12;
const TOTAL_COLUMNS_IN_APP_SHEET = APP_TRACKER_SHEET_HEADERS.length; // Should be 12

// Column Widths for "Applications" Sheet (in pixels) - Array must match header count
const APP_SHEET_COLUMN_WIDTHS = [150, 100, 100, 180, 200, 120, 100, 120, 250, 100, 100, 250];

// --- Column Configuration for "Potential Job Leads" Sheet (LEADS_SHEET_TAB_NAME) ---
const LEADS_SHEET_HEADERS = [
  "Date Added", "Company Name", "Job Title", "Location", "Salary/Pay", 
  "Source/Link", "Notes", "Status", "Follow-up Date", 
  "Source Email Subject", "Source Email ID", "Processed Timestamp"
];
// Column Index Variables (1-based) for Leads Sheet
const LEADS_DATE_ADDED_COL = 1;
const LEADS_COMPANY_COL = 2;
const LEADS_JOB_TITLE_COL = 3;
const LEADS_LOCATION_COL = 4;
const LEADS_SALARY_PAY_COL = 5;
const LEADS_SOURCE_LINK_COL = 6;
const LEADS_NOTES_COL = 7;
const LEADS_STATUS_COL = 8;
const LEADS_FOLLOW_UP_COL = 9;
const LEADS_EMAIL_SUBJECT_COL = 10;
const LEADS_EMAIL_ID_COL = 11;
const LEADS_PROCESSED_TIMESTAMP_COL = 12;
const TOTAL_COLUMNS_IN_LEADS_SHEET = LEADS_SHEET_HEADERS.length;

// Column Widths for "Potential Job Leads" Sheet
const LEADS_SHEET_COLUMN_WIDTHS = [100, 180, 200, 150, 100, 150, 250, 100, 100, 150, 100, 150];
const DEFAULT_LEAD_STATUS = "New"; // Default status for new leads

// --- Application Status Configuration ---
// Define the status options available for job applications.
// The order in STATUS_HIERARCHY defines progression (higher is "better" or "later stage").
const DEFAULT_STATUS = "Applied";
const INTERVIEW_STATUS = "Interviewing";
const OFFER_STATUS = "Offer";
const REJECTED_STATUS = "Rejected";
const KEEP_IN_VIEW_STATUS = "Keep In View"; // For leads or apps to revisit
const WITHDRAWN_STATUS = "Withdrawn";
const ACCEPTED_STATUS = "Accepted Offer"; // New status for accepted offers
const ASSESSMENT_STATUS = "Assessment"; // Added for consistency
const APPLICATION_VIEWED_STATUS = "Application Viewed"; // Added for consistency
const MANUAL_REVIEW_NEEDED = "Manual Review Needed"; // For parsing failures

// Status hierarchy for determining "Peak Status" and for sorting/filtering.
// Higher numbers mean "further along" in the process.
const STATUS_HIERARCHY = {
  [DEFAULT_STATUS]: 1,
  "Screening": 2,
  "Assessment": 3,
  [INTERVIEW_STATUS]: 4, // Generic "Interviewing"
  "Interview 1": 4.1,   // More specific interview stages
  "Interview 2": 4.2,
  "Interview 3+": 4.3,
  "Final Interview": 4.5,
  [OFFER_STATUS]: 5,
  [ACCEPTED_STATUS]: 6, // Highest positive status
  [REJECTED_STATUS]: 0, // Terminal
  [WITHDRAWN_STATUS]: -1, // Terminal, user-initiated
  [KEEP_IN_VIEW_STATUS]: 0.5, 
  [MANUAL_REVIEW_NEEDED]: -2 // Needs attention
};

// Statuses that are considered "final" and should not be automatically changed by the "stale applications" trigger.
const FINAL_STATUSES_FOR_STALE_CHECK = new Set([
  REJECTED_STATUS, 
  OFFER_STATUS, 
  WITHDRAWN_STATUS,
  ACCEPTED_STATUS, // Added
  MANUAL_REVIEW_NEEDED // Should not be marked stale if it needs review
]);
const WEEKS_THRESHOLD = 8; // Number of weeks after which an application is considered stale if not in a final status.

// --- Keyword Matching for Status Parsing ---
const OFFER_KEYWORDS = ["offer", "job offer", "offer of employment", "pleased to offer"];
const INTERVIEW_KEYWORDS = ["interview", "invitation to interview", "schedule an interview", "interview request", "like to speak with you", "let's chat", "connect with you"];
const ASSESSMENT_KEYWORDS = ["assessment", "coding challenge", "technical test", "skills test", "take-home assignment"];
const APPLICATION_VIEWED_KEYWORDS = ["application was viewed", "profile was viewed", "application has been reviewed"];
const REJECTION_KEYWORDS = ["not moving forward", "unfortunately", "decided not to proceed", "other candidates", "regret to inform", "filled the position"];

// --- Gmail Configuration (Job Application Tracker) ---
const MASTER_GMAIL_LABEL_PARENT = "CareerSuite.AI"; // Parent for all app-related labels
const TRACKER_GMAIL_LABEL_PARENT = `${MASTER_GMAIL_LABEL_PARENT}/Applications`;
const TRACKER_GMAIL_LABEL_TO_PROCESS = `${TRACKER_GMAIL_LABEL_PARENT}/To Process`;
const TRACKER_GMAIL_LABEL_PROCESSED = `${TRACKER_GMAIL_LABEL_PARENT}/Processed`;
const TRACKER_GMAIL_LABEL_MANUAL_REVIEW = `${TRACKER_GMAIL_LABEL_PARENT}/Manual Review`;

// Gmail filter query for Application Updates. Example targets common job application sites.
// IMPORTANT: This query is powerful. Test thoroughly.
// It attempts to catch emails that are replies to applications or common platform notifications.
const TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES = `(subject:("re: your application" OR "your application for" OR "update on your application" OR "thank you for applying" OR "regarding your application" OR "application status" OR "next steps" OR "interview invitation") OR from:(notify@linkedin.com OR jobs-noreply@linkedin.com OR no-reply@linkedin.com OR messages-noreply@linkedin.com OR no-reply@indeed.com OR @แจ้งเตือนผ่านindeed.com OR @monster.com OR @ziprecruiter.com OR @message.ziprecruiter.com OR @mail.glassdoor.com OR @otomate.co OR @ripplehire.com OR @icims.com OR @taleo.net OR @successfactors.com OR @myworkdayjobs.com OR @greenhouse.io OR @jobvite.com OR @bamboohr.com OR @lever.co OR @smartrecruiters.com)) AND -subject:("job alert" OR "jobs for you" OR "new jobs" OR "recommended jobs" OR "job recommendations") AND -list:(jobalert) AND -label:(${TRACKER_GMAIL_LABEL_PROCESSED}) AND -label:(${TRACKER_GMAIL_LABEL_MANUAL_REVIEW})`;

// --- Gmail Configuration (Job Leads Tracker) ---
const LEADS_GMAIL_LABEL_PARENT = `${MASTER_GMAIL_LABEL_PARENT}/Leads`;
const LEADS_GMAIL_LABEL_TO_PROCESS = `${LEADS_GMAIL_LABEL_PARENT}/To Process`;
const LEADS_GMAIL_LABEL_PROCESSED = `${LEADS_GMAIL_LABEL_PARENT}/Processed`;
const LEADS_GMAIL_LABEL_MANUAL_REVIEW = `${LEADS_GMAIL_LABEL_PARENT}/Manual Review`;

// User Property keys for storing Leads label IDs (used by processJobLeads)
const LEADS_USER_PROPERTY_TO_PROCESS_LABEL_ID = 'leadsToProcessLabelId';
const LEADS_USER_PROPERTY_PROCESSED_LABEL_ID = 'leadsProcessedLabelId';
const LEADS_USER_PROPERTY_MANUAL_LABEL_ID = 'leadsManualReviewLabelId';

// Gmail filter query for Job Leads. Example targets common job alert emails.
// This should be customized by the user for their specific job alert subscriptions.
const LEADS_GMAIL_FILTER_QUERY = `(subject:("job alert" OR "jobs for you" OR "new jobs" OR "recommended jobs" OR "job recommendations") OR from:(linkedinjobalerts@linkedin.com OR alert@indeed.com OR jobalerts@indeed.com OR info@indeed.com OR email@monster.com OR jobs@ziprecruiter.com OR job δυστυχώς@ziprecruiter.com OR noreply@glassdoor.com)) AND -label:(${LEADS_GMAIL_LABEL_PROCESSED})`;


// --- Platform Detection Keywords (from email body/sender) ---
// Maps keywords found in sender email addresses to platform names.
const IGNORED_DOMAINS = new Set(['.io', '.com', '.net', '.org']);
const PLATFORM_DOMAIN_KEYWORDS = {
  "linkedin.com": "LinkedIn",
  "indeed.com": "Indeed",
  "ziprecruiter.com": "ZipRecruiter",
  "monster.com": "Monster",
  "glassdoor.com": "Glassdoor",
  "google.com": "Google Careers", // Or other Google domains
  " Greenhouse": "Greenhouse", // Often in "via Greenhouse"
  " Lever": "Lever", // Often in "via Lever"
  " Ashby": "Ashby",
  " Workday": "Workday"
};
const DEFAULT_PLATFORM = "Email/Website"; // Default if no specific platform detected

// --- Dashboard Configuration ---
// Headers for the hidden "DashboardHelperData" sheet.
const DASHBOARD_HELPER_HEADERS = [
  "Status", "Count", "Rolling Week Applications", "Week Starting", "Applications",
  "Application Funnel Stage", "Funnel Count", "Platform Name", "Platform Count"
];
const HELPER_SHEET_COLUMN_WIDTHS = [150, 70, 180, 120, 100, 200, 100, 150, 100];

// --- Gemini API Configuration ---
const GEMINI_API_KEY_PROPERTY = 'GEMINI_API_KEY'; // UserProperty key for storing the user's Gemini API key.
const GEMINI_API_ENDPOINT_TEXT_ONLY = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite-preview-06-17:generateContent";

// Default instructions for Gemini (Job Application Parsing)
const GEMINI_SYSTEM_INSTRUCTION_APP_TRACKER = `
You are an expert assistant helping a user parse job application emails.
Your goal is to extract ONLY the Company Name, Job Title, and Application Status.
- For Company Name: Extract the specific name of the company the user applied to. If not found, output "N/A".
- For Job Title: Extract the specific job title. If not found, output "N/A".
- For Application Status: Determine the current status. Examples: "Applied", "Interview Scheduled", "Offer Extended", "Rejected", "Screening", "Assessment", "Update/Other". If no clear status, output "Update/Other".
Do not add any extra words, explanations, or formatting. Output ONLY a valid JSON object with keys "company", "title", and "status".
Example: {"company": "Google", "title": "Software Engineer", "status": "Interview Scheduled"}
If the email is clearly not a job application update (e.g., a newsletter, a job alert, a marketing email), output: {"company": "N/A", "title": "N/A", "status": "Not an Application"}
If any field is truly unidentifiable from the text, use "N/A" for that field's value.
If the email implies the application was received but no other status, use "Applied".
`;

// Default instructions for Gemini (Job Lead Email Parsing)
const GEMINI_SYSTEM_INSTRUCTION_LEADS_PARSER = `
You are an expert assistant helping a user parse job lead emails, typically from job alert services.
Your goal is to extract a list of distinct job opportunities. For each job, extract: Company Name, Job Title, Location (City, State if available, or Remote), and a direct URL link to the job posting if present.
- For Company Name: Extract the specific name of the company. If not found, output "N/A".
- For Job Title: Extract the specific job title. If not found, output "N/A".
- For Location: Extract the location. If remote, indicate "Remote". If not found, output "N/A".
- For URL: Extract the direct URL to the job posting. If multiple URLs seem relevant, pick the one most likely to be the application link. If not found, output "N/A".

Output ONLY a valid JSON array, where each element is an object with keys "company", "title", "location", and "url".
Example: [{"company": "Tech Solutions Inc.", "title": "Frontend Developer", "location": "San Francisco, CA", "url": "https://example.com/job/123"}, {"company": "Innovate Corp", "title": "UX Designer", "location": "Remote", "url": "https://example.com/job/456"}]
If the email contains no job listings (e.g., it's a confirmation of alert settings, a newsletter), output an empty JSON array: [].
If any specific field for a job (company, title, location, url) is truly unidentifiable, use "N/A" for that field's value for that specific job object, but still include the job object if other fields were found.
Prioritize accuracy. Do not invent information.
`;

// --- BRANDING COLORS ---
const BRAND_COLORS = {
  LAPIS_LAZULI: "#33658A",    // Primary (e.g., "Applications" tab, Dashboard Title BG)
  CAROLINA_BLUE: "#86BBD8",   // Secondary (e.g., "Dashboard" tab, Banding Header)
  CHARCOAL: "#2F4858",        // Text, Dark Backgrounds, "HelperData" tab
  HUNYADI_YELLOW: "#F6AE2D",  // Accent 1 (e.g., "Potential Job Leads" tab & its Banding Header)
  PALE_ORANGE: "#F26419",      // Accent 2 (e.g., for some UI elements or specific highlights)
  WHITE: "#FFFFFF",           // Text color on dark backgrounds, First Banding Row
  PALE_GREY: "#F0F4F8",       // Page/Card Light Background (for dashboard cells), Second Banding Row
  MEDIUM_GREY_BORDER: "#DDE4E9" // Borders (for dashboard cells)
};

// Ensure all necessary variables are truly global if they need to be accessed by other .gs files
// (In Apps Script, top-level const/var/let in any .gs file are typically global to the project)
