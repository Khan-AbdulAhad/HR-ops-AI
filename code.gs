/**
 * TURING AI RECRUITER V10 - OPTIMIZED
 * ====================================
 * Changes:
 * - Removed walk-away rate (no longer needed)
 * - Enhanced FAQ handling in AI prompts
 * - Optimized task loading with server-side filtering
 * - Added email thread link support
 * - Better state management
 */

const CONFIG = {
  PROJECT_ID: 'turing-230020',
  DATASET_ID: 'turing-230020',
  EXTERNAL_CONN: 'turing-230020.us.matching-vetting-prod-readonly'
};

const STAGE_CONFIG = {
  'Interested': { type: 'flag', table: 'ms2_job_match_pre_shortlist', condition: 'main.is_interested = 1' },
  'Passed VetSmith': { type: 'flag', table: 'near_term_fulfillment_funnel', condition: 'main.vetsmith_passed = 1' },
  'Pending Review': { type: 'status', table: 'ms2_job_match_pre_shortlist', system_name: 'pending-review' },
  'Completed Testing': { type: 'status', table: 'ms2_job_match', system_name: 'completed-testing' },
  'Developer Backout': { type: 'status', table: 'ms2_job_match', system_name: 'developer-backout' },
  'On Hold - Onboarding': { type: 'status', table: 'ms2_job_match', system_name: 'on-hold-onboarding' },
  'Pending Onboarding': { type: 'status', table: 'ms2_job_match', system_name: 'pending-vetting' },
  'Ready for Selection': { type: 'status', table: 'ms2_job_match', system_name: 'ready-for-selection' },
  'Selected for Trial': { type: 'status', table: 'ms2_job_match', system_name: 'selected-for-trial' }
};

// --- SETUP ---

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Turing AI Recruiter V12')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Get the currently logged-in user's email
 * Used to display who is logged in and track data consumption
 */
function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch(e) {
    console.error("Could not get user email:", e);
    return '';
  }
}

function getStoredSheetUrl() {
  return PropertiesService.getUserProperties().getProperty('LOG_SHEET_URL') || "";
}

// ==========================================
// CACHING SYSTEM FOR FAST DATA LOADING
// ==========================================

// In-memory cache for spreadsheet object (persists during single execution)
let _cachedSpreadsheet = null;
let _cachedSpreadsheetUrl = null;

/**
 * Get spreadsheet with in-memory caching (fast for multiple calls in same execution)
 */
function getCachedSpreadsheet() {
  const url = getStoredSheetUrl();
  if (!url) return null;

  // Return cached spreadsheet if URL matches
  if (_cachedSpreadsheet && _cachedSpreadsheetUrl === url) {
    return _cachedSpreadsheet;
  }

  // Open and cache
  _cachedSpreadsheet = SpreadsheetApp.openByUrl(url);
  _cachedSpreadsheetUrl = url;
  return _cachedSpreadsheet;
}

/**
 * Get sheet data with CacheService caching (fast across multiple executions)
 * @param {string} sheetName - Name of the sheet
 * @param {number} cacheSeconds - How long to cache (default 60 seconds)
 * @returns {Array} Sheet data as 2D array
 */
function getCachedSheetData(sheetName, cacheSeconds = 60) {
  const cache = CacheService.getUserCache();
  const cacheKey = 'sheet_' + sheetName;

  // Try to get from cache first
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {
      // Cache corrupted, fetch fresh
    }
  }

  // Fetch from spreadsheet
  const ss = getCachedSpreadsheet();
  if (!ss) return [];

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();

  // Cache the data (CacheService has 100KB limit per key, so check size)
  try {
    const jsonData = JSON.stringify(data);
    if (jsonData.length < 90000) { // Leave some margin
      cache.put(cacheKey, jsonData, cacheSeconds);
    }
  } catch(e) {
    // Data too large to cache, that's okay
  }

  return data;
}

/**
 * Invalidate cache for a specific sheet (call after writing to sheet)
 */
function invalidateSheetCache(sheetName) {
  const cache = CacheService.getUserCache();
  cache.remove('sheet_' + sheetName);
}

/**
 * Invalidate all sheet caches
 */
function invalidateAllSheetCaches() {
  const cache = CacheService.getUserCache();
  const sheetNames = ['Negotiation_State', 'Negotiation_Tasks', 'Negotiation_Config', 'Negotiation_Completed', 'Negotiation_FAQs'];
  sheetNames.forEach(name => cache.remove('sheet_' + name));
}

function saveLogSheetUrl(url) {
  const cleanUrl = url ? url.trim() : "";
  if(!cleanUrl) throw new Error("Invalid URL");
  try {
    const ss = SpreadsheetApp.openByUrl(cleanUrl);
    ensureSheetsExist(ss);
    PropertiesService.getUserProperties().setProperty('LOG_SHEET_URL', cleanUrl);
    return { success: true };
  } catch (e) { throw new Error("Could not access Sheet."); }
}

function ensureSheetsExist(ss) {
  // Include new sheets: Manual_Sent_Logs, Data_Fetch_Logs, Follow_Up_Queue, Daily_Reports, Unresponsive_Devs
  const sheets = ['Email_Logs', 'Email_Templates', 'Negotiation_Config', 'Negotiation_Tasks', 'Negotiation_State', 'Negotiation_FAQs', 'Negotiation_Completed', 'Rate_Tiers', 'Manual_Sent_Logs', 'Data_Fetch_Logs', 'Follow_Up_Queue', 'Daily_Reports', 'Unresponsive_Devs'];
  sheets.forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  // UPDATED: Removed Walk Away Rate column
  const confSheet = ss.getSheetByName('Negotiation_Config');
  if (confSheet.getLastRow() === 0) confSheet.appendRow(['Job ID', 'Target Rate', 'Max Rate', 'Style', 'Special Rules', 'Job Description', 'Last Updated']);

  const taskSheet = ss.getSheetByName('Negotiation_Tasks');
  if (taskSheet.getLastRow() === 0) taskSheet.appendRow(['Timestamp', 'Job ID', 'Name', 'Email', 'Agreed Rate', 'Status', 'Dev ID', 'Thread ID', 'Region']);

  // UPDATED: Added Region column to track developer's region for rate tiers
  const stateSheet = ss.getSheetByName('Negotiation_State');
  if (stateSheet.getLastRow() === 0) stateSheet.appendRow(['Email', 'Job ID', 'Attempt Count', 'Last Offer', 'Status', 'Last Reply Time', 'Dev ID', 'Name', 'AI Notes', 'Thread ID', 'Region']);

  const faqSheet = ss.getSheetByName('Negotiation_FAQs');
  if (faqSheet.getLastRow() === 0) faqSheet.appendRow(['Question', 'Answer']);

  const compSheet = ss.getSheetByName('Negotiation_Completed');
  if (compSheet.getLastRow() === 0) compSheet.appendRow(['Timestamp', 'Job ID', 'Email', 'Name', 'Final Status', 'Notes', 'Dev ID', 'Region']);

  // Rate Tiers sheet - for region-based rate management
  const rateTiersSheet = ss.getSheetByName('Rate_Tiers');
  if (rateTiersSheet.getLastRow() === 0) {
    rateTiersSheet.appendRow(['Job ID', 'Region', 'Target Rate', 'Max Rate', 'Notes']);
    // Add example data for reference
    rateTiersSheet.appendRow(['EXAMPLE', 'US/Canada', 35, 45, 'Tier 1 - High cost regions']);
    rateTiersSheet.appendRow(['EXAMPLE', 'Europe', 30, 40, 'Tier 2 - Medium-high cost']);
    rateTiersSheet.appendRow(['EXAMPLE', 'LATAM', 20, 28, 'Tier 3 - Medium cost']);
    rateTiersSheet.appendRow(['EXAMPLE', 'APAC', 18, 25, 'Tier 4 - Lower cost']);
    rateTiersSheet.appendRow(['EXAMPLE', 'India', 15, 22, 'Tier 5 - Lowest cost']);
    rateTiersSheet.appendRow(['EXAMPLE', 'Default', 25, 35, 'Fallback for unknown regions']);
  }

  // Note: Job-specific details sheets (Job_XXX_Details) are created dynamically
  // when outreach emails are sent - see getOrCreateJobDetailsSheet()

  // Email_Templates sheet - for saving/loading email templates
  const templatesSheet = ss.getSheetByName('Email_Templates');
  if (templatesSheet.getLastRow() === 0) templatesSheet.appendRow(['Template Name', 'Subject', 'Body', 'Job ID', 'Created Date']);

  // Manual_Sent_Logs sheet - for tracking manually sent emails outside this system
  const manualSentSheet = ss.getSheetByName('Manual_Sent_Logs');
  if (manualSentSheet.getLastRow() === 0) manualSentSheet.appendRow(['Timestamp', 'Job ID', 'Developer ID', 'Email', 'Name', 'Note', 'Marked By']);

  // Data_Fetch_Logs sheet - for tracking data consumption with user info
  const dataFetchSheet = ss.getSheetByName('Data_Fetch_Logs');
  if (dataFetchSheet.getLastRow() === 0) dataFetchSheet.appendRow(['Timestamp', 'Source', 'Context', 'Data Size (Bytes)', 'Duration (ms)', 'Details', 'User']);

  // Follow_Up_Queue sheet - for tracking automated follow-ups (12hr and 28hr)
  // Column 11 (K) = Manual Override - when TRUE, prevents automatic status changes from processor
  const followUpSheet = ss.getSheetByName('Follow_Up_Queue');
  if (followUpSheet.getLastRow() === 0) followUpSheet.appendRow(['Email', 'Job ID', 'Thread ID', 'Name', 'Dev ID', 'Initial Send Time', 'Follow Up 1 Sent', 'Follow Up 2 Sent', 'Status', 'Last Updated', 'Manual Override']);

  // Daily_Reports sheet - for storing daily activity reports
  const dailyReportsSheet = ss.getSheetByName('Daily_Reports');
  if (dailyReportsSheet.getLastRow() === 0) dailyReportsSheet.appendRow(['Report Date', 'Job ID', 'AI Replies Succeeded', 'Human Negotiations Sent', 'Data Gathering Emails', 'First Follow-Ups Sent', 'Second Follow-Ups Sent', 'Total Outreach', 'Generated At', 'Sent To']);

  // Unresponsive_Devs sheet - for tracking developers who didn't respond after all follow-ups
  const unresponsiveSheet = ss.getSheetByName('Unresponsive_Devs');
  if (unresponsiveSheet.getLastRow() === 0) unresponsiveSheet.appendRow(['Email', 'Job ID', 'Name', 'Dev ID', 'Thread ID', 'Initial Send Time', 'Follow Up 1 Time', 'Follow Up 2 Time', 'Marked Unresponsive', 'Days Since Initial']);
}

// --- DATA CONSUMPTION LOGGING ---

/**
 * Log data consumption for tracking/auditing purposes
 * @param {string} source - Where data came from (BigQuery, OpenAI, Gmail, etc.)
 * @param {string} context - What the data was used for (Job-51000, Negotiation, etc.)
 * @param {number} byteSize - Size of data in bytes
 * @param {number} durationMs - How long the operation took in milliseconds
 * @param {string} details - Additional details about the operation
 */
function logDataConsumption(source, context, byteSize, durationMs, details) {
  const url = getStoredSheetUrl();
  if(!url) return;

  try {
    const ss = SpreadsheetApp.openByUrl(url);

    // Get current user email
    let userEmail = 'System/Automated';
    try {
      const email = Session.getActiveUser().getEmail();
      if (email) userEmail = email;
    } catch (e) {
      // Fallback if permission issues or running in a context without a user
    }

    let sheet = ss.getSheetByName('Data_Fetch_Logs');
    if (!sheet) {
      sheet = ss.insertSheet('Data_Fetch_Logs');
      sheet.appendRow(['Timestamp', 'Source', 'Context', 'Data Size (Bytes)', 'Duration (ms)', 'Details', 'User']);
    }

    sheet.appendRow([
      new Date(),
      source,
      context,
      byteSize,
      durationMs,
      details || '',
      userEmail
    ]);
  } catch (e) {
    console.error("Failed to log data consumption:", e);
  }
}

// --- MANUAL SENT LOGS ---

/**
 * Get manual sent logs for a job
 * @param {number} jobId - The Job ID to get logs for
 * @returns {Map} Map of developer IDs to their manual sent info
 */
function getManualSentLogs(jobId) {
  const url = getStoredSheetUrl();
  if (!url) return new Map();

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName("Manual_Sent_Logs");
    if (!sheet) return new Map();

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return new Map();

    const data = sheet.getDataRange().getValues();
    const logMap = new Map();

    for (let i = 1; i < data.length; i++) {
      // Column B (index 1) = Job ID, Column C (index 2) = Developer ID
      if (String(data[i][1]) === String(jobId)) {
        const devId = String(data[i][2]);
        const note = data[i][5] || '';

        if (!logMap.has(devId)) {
          logMap.set(devId, { count: 0, note: note });
        }
        logMap.get(devId).count++;
        // Keep the most recent note
        if (note) {
          logMap.get(devId).note = note;
        }
      }
    }

    console.log(`Found ${logMap.size} manual sent entries for Job ${jobId}`);
    return logMap;
  } catch(e) {
    console.error("Error reading Manual_Sent_Logs:", e);
    return new Map();
  }
}

/**
 * Mark developers as manually sent (sent outside this system)
 * @param {Array} developerIds - Array of developer IDs to mark
 * @param {string} jobId - The Job ID
 * @param {string} note - Optional note about why/how they were sent
 */
function markAsManualSent(developerIds, jobId, note) {
  const url = getStoredSheetUrl();
  if (!url) return { success: false, error: "No config URL set" };

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    ensureSheetsExist(ss);

    const sheet = ss.getSheetByName("Manual_Sent_Logs");
    if (!sheet) {
      return { success: false, error: "Manual_Sent_Logs sheet not found" };
    }

    const markedBy = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date();
    let marked = 0;

    developerIds.forEach(devId => {
      // Schema: Timestamp, Job ID, Developer ID, Email, Name, Note, Marked By
      sheet.appendRow([
        timestamp,
        jobId,
        devId,
        '',  // Email - optional, can be filled later
        '',  // Name - optional
        note || 'Manually marked as sent',
        markedBy
      ]);
      marked++;
    });

    SpreadsheetApp.flush();

    console.log(`Marked ${marked} developers as manually sent for Job ${jobId}`);
    return { success: true, marked: marked };

  } catch (e) {
    console.error("Error in markAsManualSent:", e);
    return { success: false, error: e.message };
  }
}

// --- EMAIL TEMPLATES ---

/**
 * Get email templates for a specific job
 * @param {string} jobId - The Job ID to get templates for
 * @returns {Array} Array of template objects
 */
function getJobTemplates(jobId) {
  const url = getStoredSheetUrl();
  if(!url) {
    console.log("getJobTemplates: No sheet URL configured");
    return [];
  }

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName("Email_Templates");

    if(!sheet) {
      console.log("getJobTemplates: Email_Templates sheet not found");
      return [];
    }

    const lastRow = sheet.getLastRow();
    if(lastRow <= 1) {
      console.log("getJobTemplates: No templates found (empty sheet or only headers)");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const templates = [];
    const requestedJobId = String(jobId).trim();

    for(let i = 1; i < data.length; i++) {
      const templateJobId = String(data[i][3] || '').trim();

      // Include templates that match the job ID or have no job ID (global templates)
      if(templateJobId === requestedJobId || templateJobId === '') {
        templates.push({
          name: String(data[i][0] || `Template ${i}`).trim(),
          subject: String(data[i][1] || '').trim(),
          body: String(data[i][2] || ''),
          jobId: templateJobId,
          createdDate: data[i][4] ? new Date(data[i][4]).toLocaleString() : 'Unknown'
        });
      }
    }

    console.log(`getJobTemplates: Found ${templates.length} templates for Job ${requestedJobId}`);
    return templates;

  } catch(e) {
    console.error("getJobTemplates Error:", e);
    return [];
  }
}

/**
 * Get all email templates
 * @returns {Array} Array of all template objects
 */
function getAllTemplates() {
  const url = getStoredSheetUrl();
  if(!url) return [];

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName("Email_Templates");

    if(!sheet) return [];

    const lastRow = sheet.getLastRow();
    if(lastRow <= 1) return [];

    const data = sheet.getDataRange().getValues();
    const templates = [];

    for(let i = 1; i < data.length; i++) {
      if(data[i][0]) {
        templates.push({
          name: String(data[i][0]).trim(),
          subject: String(data[i][1] || '').trim(),
          body: String(data[i][2] || ''),
          jobId: String(data[i][3] || '').trim(),
          createdDate: data[i][4] ? new Date(data[i][4]).toLocaleString() : 'Unknown'
        });
      }
    }

    return templates;
  } catch(e) {
    console.error("getAllTemplates Error:", e);
    return [];
  }
}

/**
 * Save an email template
 * @param {string} templateName - Name of the template
 * @param {string} subject - Email subject
 * @param {string} body - Email body HTML
 * @param {string} jobId - Job ID (optional - leave empty for global templates)
 * @returns {Object} Result object with success status
 */
function saveEmailTemplate(templateName, subject, body, jobId) {
  const url = getStoredSheetUrl();
  if(!url) return { success: false, error: "No config URL" };

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    ensureSheetsExist(ss);

    const sheet = ss.getSheetByName("Email_Templates");
    if(!sheet) return { success: false, error: "Templates sheet not found" };

    const data = sheet.getDataRange().getValues();
    const cleanName = String(templateName).trim();
    const cleanJobId = String(jobId || '').trim();

    // Check if template with same name and job ID exists
    let existingRow = -1;
    for(let i = 1; i < data.length; i++) {
      const existingName = String(data[i][0]).trim();
      const existingJobId = String(data[i][3] || '').trim();

      if(existingName === cleanName && existingJobId === cleanJobId) {
        existingRow = i + 1;
        break;
      }
    }

    if(existingRow > 0) {
      // Update existing template
      sheet.getRange(existingRow, 2, 1, 2).setValues([[subject, body]]);
      sheet.getRange(existingRow, 5).setValue(new Date());
      return { success: true, message: "Template updated", isUpdate: true };
    } else {
      // Add new template
      sheet.appendRow([cleanName, subject, body, cleanJobId, new Date()]);
      return { success: true, message: "Template saved", isUpdate: false };
    }

  } catch(e) {
    console.error("saveEmailTemplate Error:", e);
    return { success: false, error: e.message };
  }
}

// --- NEGOTIATION CONFIGURATION (UPDATED - No Walk Away) ---

function saveNegotiationConfig(jobId, config) {
  const url = getStoredSheetUrl();
  if(!url) return;
  const ss = getCachedSpreadsheet();
  const sheet = ss.getSheetByName('Negotiation_Config');
  const data = sheet.getDataRange().getValues();

  let found = false;
  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(jobId)) {
      const row = i+1;
      // Updated: Removed walkAwayRate
      sheet.getRange(row, 2, 1, 6).setValues([[config.targetRate, config.maxRate, config.style, config.specialRules, config.jobDescription || '', new Date()]]);
      found = true;
      break;
    }
  }

  if(!found) {
    sheet.appendRow([jobId, config.targetRate, config.maxRate, config.style, config.specialRules, config.jobDescription || '', new Date()]);
  }

  // Invalidate cache after save
  invalidateSheetCache('Negotiation_Config');
}

function getNegotiationConfig(jobId) {
  const url = getStoredSheetUrl();
  if(!url) return null;

  // Use caching for faster loading
  const data = getCachedSheetData('Negotiation_Config', 60); // 60 second cache
  if(!data || data.length === 0) return null;

  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(jobId)) {
      return {
        targetRate: data[i][1],
        maxRate: data[i][2],
        style: data[i][3],
        specialRules: data[i][4],
        jobDescription: data[i][5] || ''
      };
    }
  }
  return null;
}

function getJobConfigList() {
  const url = getStoredSheetUrl();
  if(!url) return [];

  // Use caching for faster loading
  const data = getCachedSheetData('Negotiation_Config', 60); // 60 second cache
  if(!data || data.length === 0) return [];

  const jobs = [];
  for(let i=1; i<data.length; i++) {
    if(data[i][0]) jobs.push(String(data[i][0]));
  }
  return jobs;
}

// --- ENHANCED FAQ HELPER ---

function getFAQs() {
  const url = getStoredSheetUrl();
  if(!url) return "";

  // Use caching for faster loading
  const data = getCachedSheetData('Negotiation_FAQs', 120); // 2 minute cache (FAQs change rarely)
  if(!data || data.length <= 1) return "No specific FAQs available.";

  let faqText = "";
  for(let i=1; i<data.length; i++) {
    if(data[i][0]) {
      faqText += `Q: ${data[i][0]}\nA: ${data[i][1]}\n---\n`;
    }
  }
  return faqText;
}

// --- RATE TIERS MANAGEMENT (Region-Based Pricing) ---

/**
 * Get all rate tiers for a specific job
 * Returns array of tier objects: { region, targetRate, maxRate, notes }
 */
function getRateTiersForJob(jobId) {
  const url = getStoredSheetUrl();
  if(!url) return [];

  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName('Rate_Tiers');
  if(!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const tiers = [];

  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(jobId)) {
      tiers.push({
        region: data[i][1],
        targetRate: Number(data[i][2]) || 0,
        maxRate: Number(data[i][3]) || 0,
        notes: data[i][4] || ''
      });
    }
  }

  return tiers;
}

/**
 * Country-to-Region mapping for flexible matching
 * Maps common country names/codes to their tier regions
 */
const COUNTRY_TO_REGION_MAP = {
  // India
  'india': 'India',
  'in': 'India',
  'ind': 'India',

  // US/Canada (North America Tier 1)
  'us': 'US/Canada',
  'usa': 'US/Canada',
  'united states': 'US/Canada',
  'america': 'US/Canada',
  'canada': 'US/Canada',
  'ca': 'US/Canada',

  // Europe
  'uk': 'Europe',
  'united kingdom': 'Europe',
  'england': 'Europe',
  'germany': 'Europe',
  'de': 'Europe',
  'france': 'Europe',
  'fr': 'Europe',
  'spain': 'Europe',
  'es': 'Europe',
  'italy': 'Europe',
  'it': 'Europe',
  'netherlands': 'Europe',
  'nl': 'Europe',
  'poland': 'Europe',
  'pl': 'Europe',
  'portugal': 'Europe',
  'sweden': 'Europe',
  'norway': 'Europe',
  'denmark': 'Europe',
  'finland': 'Europe',
  'ireland': 'Europe',
  'austria': 'Europe',
  'switzerland': 'Europe',
  'belgium': 'Europe',
  'europe': 'Europe',
  'eu': 'Europe',

  // LATAM (Latin America)
  'mexico': 'LATAM',
  'mx': 'LATAM',
  'brazil': 'LATAM',
  'br': 'LATAM',
  'argentina': 'LATAM',
  'ar': 'LATAM',
  'colombia': 'LATAM',
  'co': 'LATAM',
  'chile': 'LATAM',
  'cl': 'LATAM',
  'peru': 'LATAM',
  'pe': 'LATAM',
  'venezuela': 'LATAM',
  'ecuador': 'LATAM',
  'uruguay': 'LATAM',
  'costa rica': 'LATAM',
  'panama': 'LATAM',
  'latam': 'LATAM',
  'latin america': 'LATAM',
  'south america': 'LATAM',

  // APAC (Asia-Pacific excluding India)
  'china': 'APAC',
  'cn': 'APAC',
  'japan': 'APAC',
  'jp': 'APAC',
  'korea': 'APAC',
  'south korea': 'APAC',
  'kr': 'APAC',
  'australia': 'APAC',
  'au': 'APAC',
  'new zealand': 'APAC',
  'nz': 'APAC',
  'singapore': 'APAC',
  'sg': 'APAC',
  'malaysia': 'APAC',
  'my': 'APAC',
  'indonesia': 'APAC',
  'id': 'APAC',
  'philippines': 'APAC',
  'ph': 'APAC',
  'vietnam': 'APAC',
  'vn': 'APAC',
  'thailand': 'APAC',
  'th': 'APAC',
  'taiwan': 'APAC',
  'tw': 'APAC',
  'hong kong': 'APAC',
  'hk': 'APAC',
  'apac': 'APAC',
  'asia': 'APAC',
  'asia pacific': 'APAC',

  // Eastern Europe (often separate tier)
  'ukraine': 'Eastern Europe',
  'ua': 'Eastern Europe',
  'russia': 'Eastern Europe',
  'ru': 'Eastern Europe',
  'romania': 'Eastern Europe',
  'ro': 'Eastern Europe',
  'bulgaria': 'Eastern Europe',
  'bg': 'Eastern Europe',
  'czech': 'Eastern Europe',
  'czech republic': 'Eastern Europe',
  'cz': 'Eastern Europe',
  'hungary': 'Eastern Europe',
  'hu': 'Eastern Europe',
  'serbia': 'Eastern Europe',
  'croatia': 'Eastern Europe',
  'eastern europe': 'Eastern Europe',

  // Africa
  'nigeria': 'Africa',
  'ng': 'Africa',
  'kenya': 'Africa',
  'ke': 'Africa',
  'south africa': 'Africa',
  'za': 'Africa',
  'egypt': 'Africa',
  'eg': 'Africa',
  'ghana': 'Africa',
  'morocco': 'Africa',
  'africa': 'Africa',

  // Middle East
  'israel': 'Middle East',
  'il': 'Middle East',
  'uae': 'Middle East',
  'dubai': 'Middle East',
  'saudi arabia': 'Middle East',
  'sa': 'Middle East',
  'pakistan': 'Middle East',
  'pk': 'Middle East',
  'bangladesh': 'Middle East',
  'bd': 'Middle East',
  'middle east': 'Middle East'
};

/**
 * Normalize region/country to a standard tier name
 */
function normalizeRegion(regionInput) {
  if(!regionInput) return '';

  const cleanInput = String(regionInput).toLowerCase().trim();

  // Direct mapping lookup
  if(COUNTRY_TO_REGION_MAP[cleanInput]) {
    return COUNTRY_TO_REGION_MAP[cleanInput];
  }

  // Return as-is if it's already a standard tier name
  const standardTiers = ['india', 'us/canada', 'europe', 'latam', 'apac', 'eastern europe', 'africa', 'middle east', 'default'];
  if(standardTiers.includes(cleanInput)) {
    // Capitalize properly
    return cleanInput.split('/').map(s => s.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ')).join('/');
  }

  // Return original input for custom regions
  return regionInput;
}

/**
 * Get rate tier for a specific region within a job
 * Falls back to 'Default' tier if region not found, then to job's base config
 * Now supports country names like "India", "US", "Mexico" etc.
 */
function getRateForRegion(jobId, region, ss) {
  if(!ss) {
    const url = getStoredSheetUrl();
    if(!url) return null;
    ss = SpreadsheetApp.openByUrl(url);
  }

  const sheet = ss.getSheetByName('Rate_Tiers');
  if(!sheet) return null;

  const data = sheet.getDataRange().getValues();

  // Normalize the input region to a standard tier name
  const normalizedRegion = normalizeRegion(region);
  const cleanRegion = String(normalizedRegion || '').toLowerCase().trim();
  const cleanJobId = String(jobId);

  let exactMatch = null;
  let defaultMatch = null;

  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === cleanJobId) {
      const tierRegion = String(data[i][1] || '').toLowerCase().trim();

      // Check for exact match (case-insensitive)
      if(tierRegion === cleanRegion) {
        exactMatch = {
          region: data[i][1],
          targetRate: Number(data[i][2]) || 0,
          maxRate: Number(data[i][3]) || 0,
          notes: data[i][4] || ''
        };
        break;
      }

      // Check for partial match (region contains or is contained)
      if(!exactMatch && (tierRegion.includes(cleanRegion) || cleanRegion.includes(tierRegion))) {
        exactMatch = {
          region: data[i][1],
          targetRate: Number(data[i][2]) || 0,
          maxRate: Number(data[i][3]) || 0,
          notes: data[i][4] || ''
        };
      }

      // Capture default tier
      if(tierRegion === 'default') {
        defaultMatch = {
          region: 'Default',
          targetRate: Number(data[i][2]) || 0,
          maxRate: Number(data[i][3]) || 0,
          notes: data[i][4] || ''
        };
      }
    }
  }

  // Return exact match, or default, or null
  return exactMatch || defaultMatch || null;
}

/**
 * Save or update a rate tier for a job
 */
function saveRateTier(jobId, region, targetRate, maxRate, notes) {
  const url = getStoredSheetUrl();
  if(!url) return { success: false, message: "No config URL" };

  const ss = SpreadsheetApp.openByUrl(url);
  ensureSheetsExist(ss);

  const sheet = ss.getSheetByName('Rate_Tiers');
  const data = sheet.getDataRange().getValues();

  const cleanJobId = String(jobId);
  const cleanRegion = String(region || '').trim();

  // Check if this tier already exists
  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === cleanJobId &&
       String(data[i][1]).toLowerCase() === cleanRegion.toLowerCase()) {
      // Update existing row
      sheet.getRange(i+1, 3, 1, 3).setValues([[targetRate, maxRate, notes || '']]);
      return { success: true, message: "Updated existing tier", isUpdate: true };
    }
  }

  // Add new row
  sheet.appendRow([jobId, region, targetRate, maxRate, notes || '']);
  return { success: true, message: "Added new tier", isUpdate: false };
}

/**
 * Delete a rate tier
 */
function deleteRateTier(jobId, region) {
  const url = getStoredSheetUrl();
  if(!url) return { success: false, message: "No config URL" };

  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName('Rate_Tiers');
  if(!sheet) return { success: false, message: "Rate_Tiers sheet not found" };

  const data = sheet.getDataRange().getValues();
  const cleanJobId = String(jobId);
  const cleanRegion = String(region || '').toLowerCase().trim();

  for(let i=data.length-1; i>=1; i--) {
    if(String(data[i][0]) === cleanJobId &&
       String(data[i][1]).toLowerCase() === cleanRegion) {
      sheet.deleteRow(i+1);
      return { success: true, message: "Tier deleted" };
    }
  }

  return { success: false, message: "Tier not found" };
}

/**
 * Copy rate tiers from one job to another (useful for similar jobs)
 */
function copyRateTiers(sourceJobId, targetJobId) {
  const tiers = getRateTiersForJob(sourceJobId);
  if(tiers.length === 0) return { success: false, message: "No tiers found for source job" };

  let copied = 0;
  tiers.forEach(tier => {
    const result = saveRateTier(targetJobId, tier.region, tier.targetRate, tier.maxRate, tier.notes);
    if(result.success) copied++;
  });

  return { success: true, message: `Copied ${copied} tiers to job ${targetJobId}` };
}

/**
 * Get all unique regions used across all jobs (for UI dropdown)
 */
function getAllRegions() {
  const url = getStoredSheetUrl();
  if(!url) return [];

  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName('Rate_Tiers');
  if(!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const regions = new Set();

  for(let i=1; i<data.length; i++) {
    if(data[i][1] && data[i][0] !== 'EXAMPLE') {
      regions.add(data[i][1]);
    }
  }

  return Array.from(regions).sort();
}

// --- CANDIDATE DETAILS EXTRACTION (JOB-SPECIFIC SHEETS) ---

/**
 * Analyze the outreach email to extract questions being asked to candidates
 * Returns an array of question objects with short headers and full questions
 */
function analyzeOutreachForQuestions(emailBody, jobId) {
  const prompt = `
You are analyzing a recruitment outreach email to identify what information/questions are being asked from candidates.

EMAIL CONTENT:
"${emailBody}"

TASK:
Identify ALL questions or information requests in this email. For each one, provide:
1. A short header (2-4 words, suitable for a spreadsheet column)
2. The full question or request

Common things recruiters ask about:
- Expected/desired rate
- Start date / availability
- Notice period
- Work hours availability
- Location/timezone
- Equipment (laptop, etc.)
- Current employment status
- Years of experience
- Specific skills or certifications

Return a JSON array of objects like this:
[
  {"header": "Expected Rate", "question": "What is your expected hourly rate?"},
  {"header": "Start Date", "question": "When can you start?"},
  {"header": "Weekly Hours", "question": "How many hours per week can you commit?"}
]

RULES:
- Only include questions that are EXPLICITLY asked in the email
- Keep headers short and clear (max 4 words)
- If the email asks about rate negotiation, include "Expected Rate" as a header
- Always include basic fields: "Expected Rate", "Start Date" even if not explicitly asked
- Maximum 10 questions/headers

Return ONLY the JSON array, no other text.
`;

  try {
    const response = callAI(prompt);
    let cleanResponse = response.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const questions = JSON.parse(cleanResponse);

    // Always ensure we have basic required fields
    const headers = questions.map(q => q.header);
    if (!headers.includes("Expected Rate")) {
      questions.unshift({header: "Expected Rate", question: "What is your expected hourly rate?"});
    }
    if (!headers.includes("Start Date")) {
      questions.push({header: "Start Date", question: "When can you start?"});
    }

    return questions;
  } catch(e) {
    console.error("Failed to analyze outreach for questions:", e);
    // Return default questions if AI fails
    return [
      {header: "Expected Rate", question: "What is your expected hourly rate?"},
      {header: "Start Date", question: "When can you start?"},
      {header: "Notice Period", question: "What is your notice period?"},
      {header: "Weekly Hours", question: "How many hours per week can you work?"},
      {header: "Has Equipment", question: "Do you have your own laptop?"}
    ];
  }
}

/**
 * Get or create a job-specific details sheet
 * Sheet name format: "Job_{jobId}_Details"
 */
function getOrCreateJobDetailsSheet(ss, jobId, emailBody) {
  const sheetName = `Job_${jobId}_Details`;
  let sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    return { sheet: sheet, isNew: false };
  }

  // Create new sheet
  sheet = ss.insertSheet(sheetName);

  // Analyze outreach email to get questions and create headers
  const questions = analyzeOutreachForQuestions(emailBody, jobId);

  // Store questions metadata in first row as JSON (hidden later or in a config)
  // Build headers: fixed columns + dynamic question columns + status columns
  const fixedHeaders = ['Timestamp', 'Email', 'Name', 'Dev ID', 'Thread ID', 'Region'];
  const questionHeaders = questions.map(q => q.header);
  const statusHeaders = ['Negotiation Notes', 'Status', 'Agreed Rate'];

  const allHeaders = [...fixedHeaders, ...questionHeaders, ...statusHeaders];

  // Set headers
  sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);

  // Store the questions metadata in script properties for this job
  const questionsKey = `JOB_${jobId}_QUESTIONS`;
  PropertiesService.getScriptProperties().setProperty(questionsKey, JSON.stringify(questions));

  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, allHeaders.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // Freeze header row
  sheet.setFrozenRows(1);

  return { sheet: sheet, isNew: true, questions: questions };
}

/**
 * Get the questions configured for a job
 */
function getJobQuestions(jobId) {
  const questionsKey = `JOB_${jobId}_QUESTIONS`;
  const stored = PropertiesService.getScriptProperties().getProperty(questionsKey);

  if (stored) {
    try {
      return JSON.parse(stored);
    } catch(e) {
      console.error("Failed to parse stored questions:", e);
    }
  }

  // Return default questions if not found
  return [
    {header: "Expected Rate", question: "What is your expected hourly rate?"},
    {header: "Start Date", question: "When can you start?"},
    {header: "Notice Period", question: "What is your notice period?"},
    {header: "Weekly Hours", question: "How many hours per week can you work?"},
    {header: "Has Equipment", question: "Do you have your own laptop?"}
  ];
}

/**
 * Extract answers from candidate's response based on the questions asked
 */
function extractAnswersFromResponse(candidateMessage, questions, candidateName) {
  const questionsList = questions.map((q, i) => `${i+1}. "${q.header}": ${q.question}`).join('\n');

  const prompt = `
You are an intelligent data extraction assistant analyzing a candidate's email response to extract answers to specific questions.

CANDIDATE'S MESSAGE:
"${candidateMessage}"

CANDIDATE NAME: ${candidateName}

QUESTIONS TO EXTRACT ANSWERS FOR:
${questionsList}

TASK:
For each question, extract the candidate's answer from their message using semantic understanding.

EXTRACTION STRATEGY:
1. Understand the INTENT of each question, not just the literal words
2. Match candidate's response to questions based on MEANING, not exact keywords
3. Candidates often answer questions indirectly or use different phrasing - recognize these as valid answers
4. If the candidate provides information that answers a question, extract it even if they didn't directly address that question

COMMON SEMANTIC EQUIVALENTS (apply similar logic to any question type):
- Start Date / Availability: "available for/from [date]", "free on [date]", "can join [date]", "ready by [date]"
- Location: "based in [place]", "living in [place]", "from [place]", "I'm in [place]", "currently in [place]"
- Rate / Salary: "expecting [amount]", "looking for [amount]", "my rate is [amount]", "charge [amount]"
- Education: "studied at [school]", "graduated from [school]", "have a [degree]", "completed [degree]"
- Experience: "[X] years in [field]", "working as [role] for [time]", "been doing [skill] since [year]"
- Age: "I am [X] years old", "[X] y/o", "born in [year]"
- Time Overlap / Schedule: "can work [hours]", "available during [time]", "my working hours are [time]"

Return a JSON object where:
- Keys are the exact header names from the questions
- Values are the answers extracted from the message

Use these special values when appropriate:
- "NOT_PROVIDED" - if the candidate didn't mention this at all
- "NEGOTIATING" - if they're vague or want to discuss later (e.g., "depends on the project", "let's discuss")
- The actual answer if they provided one

Example response format:
{
  "Expected Rate": "$45/hr",
  "Start Date": "12th December",
  "Location": "Mumbai, India",
  "Education": "NOT_PROVIDED",
  "Weekly Hours": "NEGOTIATING",
  "is_negotiating": true,
  "negotiation_notes": "Candidate wants to negotiate rate, asking for $45 but willing to discuss"
}

IMPORTANT:
- Always include "is_negotiating" (true/false) and "negotiation_notes" fields
- Extract the value the candidate actually mentions, even if different from what was asked
  Example: Asked "Can you start December 5th?" → Candidate says "available for 12th dec" → Extract "12th dec"
- If they give a range, include the full range (e.g., "$40-50/hr", "25-30 years old")
- Capture any concerns, conditions, or counter-proposals in negotiation_notes
- When in doubt about whether something answers a question, extract it rather than marking NOT_PROVIDED

Return ONLY the JSON object, no other text.
`;

  try {
    const response = callAI(prompt);
    let cleanResponse = response.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    return JSON.parse(cleanResponse);
  } catch(e) {
    console.error("Failed to extract answers:", e);
    // Return empty answers with error note
    const errorResult = {
      is_negotiating: false,
      negotiation_notes: "Failed to parse response"
    };
    questions.forEach(q => {
      errorResult[q.header] = "PARSE_ERROR";
    });
    return errorResult;
  }
}

/**
 * Save or update candidate details in the job-specific sheet
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} jobId - Job ID
 * @param {string} candidateEmail - Candidate's email
 * @param {string} candidateName - Candidate's name
 * @param {string} devId - Developer ID
 * @param {string} threadId - Gmail thread ID
 * @param {Object} answers - Extracted answers object
 * @param {string} status - Current status
 * @param {string} region - Developer's region for rate tier (optional)
 */
function saveJobCandidateDetails(ss, jobId, candidateEmail, candidateName, devId, threadId, answers, status, region) {
  const sheetName = `Job_${jobId}_Details`;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.error(`Job details sheet not found: ${sheetName}`);
    return { success: false, message: "Job details sheet not found" };
  }

  const cleanEmail = String(candidateEmail).toLowerCase().trim();
  const questions = getJobQuestions(jobId);

  // Get headers from sheet first
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Fixed column headers that are NOT question columns
  const fixedHeaders = ['Timestamp', 'Email', 'Name', 'Dev ID', 'Thread ID', 'Region', 'Negotiation Notes', 'Status', 'Agreed Rate'];

  // Log for debugging
  console.log(`saveJobCandidateDetails: jobId=${jobId}, email=${cleanEmail}, questions=${questions.length}, answers=${JSON.stringify(answers)}`);

  // If no questions configured or questions don't match sheet, auto-detect from sheet headers
  let effectiveQuestions = questions;
  if (!questions || questions.length === 0) {
    console.warn(`No questions configured for job ${jobId}, auto-detecting from sheet headers`);
    effectiveQuestions = [];
    headers.forEach(h => {
      if (h && !fixedHeaders.includes(h)) {
        effectiveQuestions.push({ header: h, question: h });
      }
    });
    console.log(`Auto-detected ${effectiveQuestions.length} question columns: ${effectiveQuestions.map(q => q.header).join(', ')}`);
  }

  // Find fixed column indices
  const emailColIdx = headers.indexOf('Email');
  const timestampColIdx = headers.indexOf('Timestamp');
  const nameColIdx = headers.indexOf('Name');
  const devIdColIdx = headers.indexOf('Dev ID');
  const threadIdColIdx = headers.indexOf('Thread ID');
  const regionColIdx = headers.indexOf('Region');
  const notesColIdx = headers.indexOf('Negotiation Notes');
  const statusColIdx = headers.indexOf('Status');
  const agreedRateColIdx = headers.indexOf('Agreed Rate');

  // Check if candidate already exists in sheet
  const data = sheet.getDataRange().getValues();
  let existingRowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][emailColIdx]).toLowerCase() === cleanEmail) {
      existingRowIndex = i + 1;
      break;
    }
  }

  // Build row data
  const rowData = new Array(headers.length).fill('');
  rowData[timestampColIdx] = new Date();
  rowData[emailColIdx] = candidateEmail;
  rowData[nameColIdx] = candidateName;
  rowData[devIdColIdx] = devId || 'N/A';
  rowData[threadIdColIdx] = threadId || '';
  if(regionColIdx !== -1) rowData[regionColIdx] = region || '';
  rowData[notesColIdx] = answers.negotiation_notes || '';

  // Fill in question answers and track data completeness
  let totalQuestions = 0;
  let answeredQuestions = 0;
  let pendingQuestions = [];

  effectiveQuestions.forEach(q => {
    const colIdx = headers.indexOf(q.header);
    if (colIdx !== -1) {
      totalQuestions++;
      const answer = answers[q.header];
      if (answer && answer !== 'NOT_PROVIDED' && answer !== 'PARSE_ERROR') {
        answeredQuestions++;
        rowData[colIdx] = answer;
      } else {
        pendingQuestions.push(q.header);
      }
    }
  });

  // Determine data gathering status automatically
  let dataGatheringStatus = status; // Use passed status if provided (e.g., 'Human Escalation')

  if (!status || status === 'Negotiating' || status === 'Details Provided' || status === 'Pending' || status === 'Data Complete') {
    if (answers.is_negotiating) {
      dataGatheringStatus = 'Negotiating';
    } else if (totalQuestions > 0 && answeredQuestions === totalQuestions) {
      dataGatheringStatus = 'Data Complete';
    } else if (answeredQuestions > 0) {
      dataGatheringStatus = 'Pending';
    } else {
      dataGatheringStatus = 'Pending';
    }
  }

  rowData[statusColIdx] = dataGatheringStatus;

  if (existingRowIndex > -1) {
    // Merge with existing data - keep non-empty existing values
    const existingRow = data[existingRowIndex - 1];
    let mergedAnsweredCount = 0;

    // Build a set of question column indices for accurate merging
    const questionColIndices = new Set();
    effectiveQuestions.forEach(q => {
      const idx = headers.indexOf(q.header);
      if (idx !== -1) questionColIndices.add(idx);
    });

    for (let col = 0; col < headers.length; col++) {
      // For question columns, keep existing if new is empty/NOT_PROVIDED
      if (questionColIndices.has(col)) {
        if ((!rowData[col] || rowData[col] === 'NOT_PROVIDED') &&
            existingRow[col] && existingRow[col] !== 'NOT_PROVIDED') {
          rowData[col] = existingRow[col];
        }
        // Count answered questions after merge
        if (rowData[col] && rowData[col] !== 'NOT_PROVIDED' && rowData[col] !== 'PARSE_ERROR') {
          mergedAnsweredCount++;
        }
      }
      // Keep existing region if new one is empty
      if (col === regionColIdx && !rowData[col] && existingRow[col]) {
        rowData[col] = existingRow[col];
      }
    }

    // Recalculate status after merge (unless it's a special status like 'Offer Accepted' or 'Human Escalation')
    // Note: Human-Negotiation and Pending Escalation are also preserved to not lose escalation info
    const preserveStatuses = ['Offer Accepted', 'Human Escalation', 'Human-Negotiation', 'Escalated', 'Completed', 'Pending Escalation', 'Rate Agreed'];
    const existingStatus = existingRow[statusColIdx];

    if (!preserveStatuses.some(s => String(existingStatus).includes(s)) &&
        !preserveStatuses.some(s => String(status).includes(s))) {
      if (totalQuestions > 0 && mergedAnsweredCount === totalQuestions) {
        rowData[statusColIdx] = 'Data Complete';
      } else if (mergedAnsweredCount > 0 && !answers.is_negotiating) {
        rowData[statusColIdx] = 'Pending';
      }
    } else if (preserveStatuses.some(s => String(existingStatus).includes(s))) {
      // Keep existing special status
      rowData[statusColIdx] = existingStatus;
    }

    sheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
    return { success: true, message: "Updated existing candidate", isUpdate: true, dataComplete: mergedAnsweredCount === totalQuestions };
  } else {
    // Add new row
    sheet.appendRow(rowData);
    return { success: true, message: "Added new candidate", isUpdate: false, dataComplete: answeredQuestions === totalQuestions };
  }
}

/**
 * Update candidate status in job details sheet (for completion/acceptance)
 */
function updateJobCandidateStatus(ss, jobId, candidateEmail, status, agreedRate) {
  const sheetName = `Job_${jobId}_Details`;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) return { success: false, message: "Sheet not found" };

  const cleanEmail = String(candidateEmail).toLowerCase().trim();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const emailColIdx = headers.indexOf('Email');
  const statusColIdx = headers.indexOf('Status');
  const agreedRateColIdx = headers.indexOf('Agreed Rate');
  const timestampColIdx = headers.indexOf('Timestamp');

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][emailColIdx]).toLowerCase() === cleanEmail) {
      const rowIndex = i + 1;
      sheet.getRange(rowIndex, statusColIdx + 1).setValue(status);
      sheet.getRange(rowIndex, timestampColIdx + 1).setValue(new Date());
      if (agreedRate && agreedRateColIdx !== -1) {
        sheet.getRange(rowIndex, agreedRateColIdx + 1).setValue(agreedRate);
      }
      return { success: true };
    }
  }

  return { success: false, message: "Candidate not found" };
}

/**
 * Manually set the agreed rate for a candidate (for human negotiators)
 * This updates the job-specific details sheet with the negotiated rate
 */
function setAgreedRateForCandidate(email, jobId, agreedRate) {
  const url = getStoredSheetUrl();
  if(!url) return { success: false, message: "No config URL" };

  const ss = SpreadsheetApp.openByUrl(url);

  // Format the rate consistently
  let formattedRate = agreedRate;
  if (agreedRate) {
    // Remove any existing $ or /hr to normalize, then reformat
    const rateNum = String(agreedRate).replace(/[$,\/hr\s]/gi, '');
    if (!isNaN(parseFloat(rateNum))) {
      formattedRate = `$${rateNum}/hr`;
    }
  }

  try {
    const result = updateJobCandidateStatus(ss, jobId, email, 'Rate Agreed (Human)', formattedRate);
    if (result.success) {
      return { success: true, message: `Agreed rate ${formattedRate} saved for ${email}` };
    }
    return result;
  } catch(e) {
    console.error("Failed to set agreed rate:", e);
    return { success: false, message: "Failed to update: " + e.message };
  }
}

/**
 * Complete a human negotiation with an agreed rate
 * This combines moveToCompleted with rate setting
 */
function completeHumanNegotiation(email, jobId, agreedRate, finalStatus) {
  const url = getStoredSheetUrl();
  if(!url) return { success: false, message: "No config URL" };

  const ss = SpreadsheetApp.openByUrl(url);

  // Format the rate
  let formattedRate = null;
  if (agreedRate) {
    const rateNum = String(agreedRate).replace(/[$,\/hr\s]/gi, '');
    if (!isNaN(parseFloat(rateNum))) {
      formattedRate = `$${rateNum}/hr`;
    }
  }

  // First, update the job details sheet with the rate and status
  try {
    const status = finalStatus || (agreedRate ? 'Offer Accepted (Human)' : 'Completed (Human)');
    updateJobCandidateStatus(ss, jobId, email, status, formattedRate);
  } catch(e) {
    console.error("Failed to update job details:", e);
  }

  // Then move to completed
  const result = moveToCompleted(email, finalStatus || 'Human Negotiation Complete', jobId);
  return result;
}

/**
 * Get the outreach email body for a job (from Email_Logs or stored config)
 */
function getJobOutreachEmail(ss, jobId) {
  // First try to get from Negotiation_Config (if we store it there)
  const configSheet = ss.getSheetByName('Negotiation_Config');
  if (configSheet) {
    const configData = configSheet.getDataRange().getValues();
    for (let i = 1; i < configData.length; i++) {
      if (String(configData[i][0]) === String(jobId)) {
        // Check if there's an outreach template stored (we'll add this column)
        if (configData[i][7]) { // Column 8 = Outreach Template
          return configData[i][7];
        }
      }
    }
  }

  // Fallback: return empty (will use default questions)
  return '';
}


// --- OPTIMIZED TASK LIST MANAGEMENT ---

function getAllTasks(filters) {
  const url = getStoredSheetUrl();
  if(!url) return { tasks: [], jobIds: [], stats: { total: 0, active: 0, human: 0, accepted: 0 } };

  // If forceRefresh is true, invalidate caches first
  if (filters?.forceRefresh) {
    invalidateSheetCache('Negotiation_State');
    invalidateSheetCache('Negotiation_Tasks');
  }

  // Use caching for faster loading
  const tasks = [];
  const jobIdSet = new Set();

  // Stats counters
  let statActive = 0, statHuman = 0, statAccepted = 0;

  // Apply filters if provided
  const jobFilter = filters?.jobId || 'all';
  const statusFilter = filters?.status || 'all';

  // 1. Get Active Negotiations (State) - with caching
  const stateData = getCachedSheetData('Negotiation_State', 30); // 30 second cache
  if(!stateData || stateData.length === 0) return { tasks: [], jobIds: [], stats: { total: 0, active: 0, human: 0, accepted: 0 } };
  
  for(let i=1; i<stateData.length; i++) {
    if(!stateData[i][0]) continue;

    const jobId = String(stateData[i][1]);
    const status = stateData[i][4] || 'Active';
    const attempts = Number(stateData[i][2]) || 0;

    // Collect all job IDs for filter dropdown
    jobIdSet.add(jobId);

    // Count stats (always, regardless of filters)
    if(status === 'Human-Negotiation') {
      statHuman++;
    } else {
      statActive++;
    }

    // Apply filters server-side
    if(jobFilter !== 'all' && jobId !== jobFilter) continue;
    if(statusFilter !== 'all' && status !== statusFilter) continue;

    let tag = '';
    if(status === 'Human-Negotiation') {
      tag = 'Human-Negotiation';
    } else if(status === 'Initial Outreach' || attempts === 0) {
      tag = 'Initial Outreach';
    } else {
      tag = `AI-Attempt-${attempts}/2`;
    }

    tasks.push({
      email: stateData[i][0],
      jobId: jobId,
      devId: stateData[i][6] || 'N/A',
      name: stateData[i][7] || 'Unknown',
      status: status,
      attempts: attempts,
      tags: tag,
      type: 'Negotiating',
      lastReply: stateData[i][5] ? new Date(stateData[i][5]).toLocaleString() : 'N/A',
      aiNotes: stateData[i][8] || '',
      threadId: stateData[i][9] || ''
    });
  }

  // 2. Get Accepted Offers (Tasks) - with caching
  const taskData = getCachedSheetData('Negotiation_Tasks', 30); // 30 second cache
  if(taskData && taskData.length > 0) {
    for(let i=1; i<taskData.length; i++) {
      if(!taskData[i][3]) continue;
      if(taskData[i][5] === 'Archived') continue;

      const jobId = String(taskData[i][1]);
      jobIdSet.add(jobId);

      // Count for stats (always)
      statAccepted++;

      // Apply filters for display
      if(statusFilter !== 'all' && statusFilter !== 'Offer Accepted') continue;
      if(jobFilter !== 'all' && jobId !== jobFilter) continue;

      tasks.push({
        email: taskData[i][3],
        jobId: jobId,
        devId: taskData[i][6] || 'N/A',
        name: taskData[i][2] || 'Unknown',
        status: 'Offer Accepted',
        attempts: 'N/A',
        tags: 'Completed',
        type: 'Accepted',
        agreedRate: taskData[i][4] || 'N/A',
        lastReply: taskData[i][0] ? new Date(taskData[i][0]).toLocaleString() : 'N/A',
        threadId: taskData[i][7] || ''
      });
    }
  }

  return {
    tasks: tasks,
    jobIds: Array.from(jobIdSet).sort(),
    stats: {
      total: statActive + statHuman + statAccepted,
      active: statActive,
      human: statHuman,
      accepted: statAccepted
    }
  };
}

// Bulk complete multiple tasks
function bulkComplete(emails, finalStatus) {
  const results = { success: 0, failed: 0, errors: [] };
  
  emails.forEach(email => {
    try {
      const result = moveToCompleted(email, finalStatus);
      if(result.success) {
        results.success++;
      } else {
        results.failed++;
        results.errors.push(`${email}: ${result.message}`);
      }
    } catch(e) {
      results.failed++;
      results.errors.push(`${email}: ${e.message}`);
    }
  });
  
  return results;
}

function moveToCompleted(email, finalStatus, jobIdFilter) {
  const url = getStoredSheetUrl();
  if(!url) return { success: false, message: "No config URL" };

  const ss = SpreadsheetApp.openByUrl(url);
  ensureSheetsExist(ss);

  const compSheet = ss.getSheetByName('Negotiation_Completed');
  const stateSheet = ss.getSheetByName('Negotiation_State');
  const taskSheet = ss.getSheetByName('Negotiation_Tasks');

  if(!compSheet || !stateSheet || !taskSheet) {
    return { success: false, message: "Required sheets not found" };
  }

  let moved = false;
  let taskInfo = { jobId: '', name: '', devId: '', threadId: '', agreedRate: null };
  const cleanEmail = String(email).toLowerCase();

  // First, find info from Task Sheet
  const taskData = taskSheet.getDataRange().getValues();
  for(let i=taskData.length-1; i>=1; i--) {
    if(String(taskData[i][3]).toLowerCase() === cleanEmail) {
      // If jobIdFilter is provided, only match if job ID matches
      if(jobIdFilter && String(taskData[i][1]) !== String(jobIdFilter)) continue;

      // Column 4 contains the agreed rate from Task sheet
      const rate = taskData[i][4];
      taskInfo = {
        jobId: taskData[i][1],
        name: taskData[i][2],
        devId: taskData[i][6] || 'N/A',
        threadId: taskData[i][7] || '',
        agreedRate: rate ? `$${rate}/hr` : null
      };
      compSheet.appendRow([new Date(), taskData[i][1], email, taskData[i][2], finalStatus || "Accepted", "Moved from Task List", taskData[i][6] || 'N/A']);
      taskSheet.deleteRow(i+1);
      moved = true;
      break;
    }
  }

  // Also remove from State Sheet if exists
  const stateData = stateSheet.getDataRange().getValues();
  for(let i=stateData.length-1; i>=1; i--) {
    if(String(stateData[i][0]).toLowerCase() === cleanEmail) {
      // If jobIdFilter is provided, only match if job ID matches
      if(jobIdFilter && String(stateData[i][1]) !== String(jobIdFilter)) continue;

      if(!moved) {
        taskInfo = {
          jobId: stateData[i][1],
          name: stateData[i][7] || 'Unknown',
          devId: stateData[i][6] || 'N/A',
          threadId: stateData[i][9] || ''
        };
        compSheet.appendRow([new Date(), stateData[i][1], email, stateData[i][7] || 'Unknown', finalStatus || stateData[i][4], "Moved from State List", stateData[i][6] || 'N/A']);
        moved = true;
      }
      // Capture threadId if not already captured
      if(!taskInfo.threadId) {
        taskInfo.threadId = stateData[i][9] || '';
      }
      // Capture jobId if not already captured
      if(!taskInfo.jobId) {
        taskInfo.jobId = stateData[i][1];
      }
      stateSheet.deleteRow(i+1);
      break;
    }
  }

  // UPDATE JOB-SPECIFIC DETAILS SHEET - Mark as completed with agreed rate if available
  if(moved && taskInfo.jobId) {
    try {
      updateJobCandidateStatus(ss, taskInfo.jobId, email, `Completed: ${finalStatus || 'Done'}`, taskInfo.agreedRate);
    } catch(e) {
      console.error("Failed to update job details sheet:", e);
    }
  }

  // ADD COMPLETED LABEL TO GMAIL THREAD - This saves processing cost!
  if(moved && taskInfo.threadId) {
    try {
      const thread = GmailApp.getThreadById(taskInfo.threadId);
      if(thread) {
        const completedLabel = GmailApp.getUserLabelByName("Completed") || GmailApp.createLabel("Completed");
        thread.addLabel(completedLabel);

        // Also remove Human-Negotiation label if present
        const humanLabel = GmailApp.getUserLabelByName("Human-Negotiation");
        if(humanLabel) {
          thread.removeLabel(humanLabel);
        }
      }
    } catch(e) {
      console.error("Failed to add Completed label to Gmail:", e);
      // Don't fail the whole operation just because labeling failed
    }
  }

  // Invalidate caches after modifying data
  if(moved) {
    invalidateSheetCache('Negotiation_State');
    invalidateSheetCache('Negotiation_Tasks');
    invalidateSheetCache('Negotiation_Completed');
  }

  return { success: moved, message: moved ? "Moved to completed" : "Email not found", jobId: taskInfo.jobId };
}

// ==============================================================
// ===  FIXED getDevelopers - WITH AGENCY/SUBCONTRACTOR SUPPORT ===
// ==============================================================

function getDevelopers(jobId, selectedStages) {
  if (!jobId || !selectedStages || selectedStages.length === 0) throw new Error("Missing inputs");
  const cleanJobId = Number(jobId);

  // Get both regular logs AND manual sent logs (guard with empty Map)
  const logs = getEmailLogs(cleanJobId) || new Map();
  const manualLogs = getManualSentLogs(cleanJobId) || new Map();

  // Build per-stage query chunks (use single quotes for SQL string literals)
  let queryChunks = [];
  selectedStages.forEach(stage => {
    const config = STAGE_CONFIG[stage];
    if (config) {
      const q = config.type === 'flag'
        ? `SELECT DISTINCT main.developer_id, '${stage}' as stage_label FROM ${config.table} main WHERE main.job_id = ${cleanJobId} AND ${config.condition}`
        : `SELECT DISTINCT main.developer_id, '${stage}' as stage_label FROM ${config.table} main LEFT JOIN ms2_job_match_status s ON main.job_match_status_id = s.id WHERE main.job_id = ${cleanJobId} AND s.system_name = '${config.system_name}'`;
      queryChunks.push(q);
    } else {
      console.warn(`getDevelopers: Unknown stage requested: ${stage}`);
    }
  });

  if (queryChunks.length === 0) {
    console.log("getDevelopers: No valid stages selected");
    return [];
  }

  const unionQuery = queryChunks.join(' UNION ALL ');

  // Use GROUP_CONCAT + SUBSTRING_INDEX to pick the highest-priority stage per developer
  // Priority (1 = highest) encoded in ORDER BY CASE expression below:
  const innerQuery = `
    WITH target_devs AS (
      ${unionQuery}
    ),
    -- For each developer, order their stages by our priority and take the first
    unique_devs AS (
      SELECT
        td.developer_id,
        SUBSTRING_INDEX(
          GROUP_CONCAT(DISTINCT td.stage_label
            ORDER BY CASE td.stage_label
              WHEN 'Selected for Trial' THEN 1
              WHEN 'Ready for Selection' THEN 2
              WHEN 'On Hold - Onboarding' THEN 3
              WHEN 'Developer Backout' THEN 4
              WHEN 'Completed Testing' THEN 5
              WHEN 'Pending Onboarding' THEN 6
              WHEN 'Pending Review' THEN 7
              WHEN 'Passed VetSmith' THEN 8
              WHEN 'Interested' THEN 9
              ELSE 99 END
            SEPARATOR ','
          ),
          ',', 1
        ) AS stage_label
      FROM target_devs td
      GROUP BY td.developer_id
    ),
    unique_ids AS (
      SELECT DISTINCT developer_id FROM unique_devs
    ),
    -- FIXED: Properly aggregate agency data with developer_type and review_status filter
    agency_info AS (
      SELECT
        dev_id,
        IF(MAX(IF(developer_type = 'sub_contractor', 1, 0)) = 1, 'Sub Contractor', 'Contractor') AS agency_sub_con,
        MAX(a.name) AS agency_name
      FROM ms2_agency_devs ad
      LEFT JOIN ms2_agencies a ON ad.agency_id = a.id
      LEFT JOIN ms2_agency_devs_applications ada ON ad.id = ada.agency_dev_id
      WHERE review_status = 'approved'
      GROUP BY dev_id
    ),
    dev_details AS (
      SELECT
        d.id,
        d.full_name,
        d.email,
        CASE
          WHEN ai.agency_sub_con = 'Sub Contractor' THEN 'Agency Sub-Contractor'
          WHEN ai.agency_sub_con = 'Contractor' THEN 'Agency Contractor'
          ELSE 'Independent'
        END AS candidate_status,
        COALESCE(ai.agency_name, '') AS agency_name
      FROM user_list_v4 d
      LEFT JOIN agency_info ai ON d.id = ai.dev_id
      WHERE d.id IN (SELECT developer_id FROM unique_ids)
    )
    SELECT ud.developer_id, d.full_name, d.email, ud.stage_label AS status, d.candidate_status, d.agency_name
    FROM unique_devs ud
    JOIN dev_details d ON ud.developer_id = d.id
  `;

  const finalSql = `SELECT * FROM EXTERNAL_QUERY("${CONFIG.EXTERNAL_CONN}", """${innerQuery}""")`;

  try {
    const startTime = new Date().getTime();

    let queryResults = BigQuery.Jobs.query({ query: finalSql, useLegacySql: false }, CONFIG.PROJECT_ID);
    let job = queryResults.jobReference;
    while (!queryResults.jobComplete) {
      Utilities.sleep(500);
      queryResults = BigQuery.Jobs.getQueryResults(CONFIG.PROJECT_ID, job.jobId);
    }

    const endTime = new Date().getTime();
    const rows = queryResults.rows || [];
    const dataSizeBytes = JSON.stringify(rows).length;

    // Log data consumption
    try {
      logDataConsumption('BigQuery', `Job-${cleanJobId}`, dataSizeBytes, endTime - startTime, `Rows returned: ${rows.length}`);
    } catch (le) {
      console.warn("logDataConsumption call failed:", le);
    }

    // Map results into developer objects (ensure uniqueness by developer_id as a safety net)
    const devMap = new Map();
    (rows || []).forEach(row => {
      // BigQuery EXTERNAL_QUERY returns row.f[].v shape
      const devId = row.f[0] && row.f[0].v ? String(row.f[0].v) : null;
      if (!devId) return;

      const fullName = row.f[1] && row.f[1].v ? row.f[1].v : '';
      const email = row.f[2] && row.f[2].v ? row.f[2].v : '';
      const status = row.f[3] && row.f[3].v ? row.f[3].v : '';
      const candidateStatus = row.f[4] && row.f[4].v ? row.f[4].v : 'Independent';
      const agencyName = row.f[5] && row.f[5].v ? row.f[5].v : '';

      // Defensive uniqueness: if we already have the dev, merge conservatively
      if (!devMap.has(devId)) {
        const emailLower = String(email).toLowerCase().trim();
        const log = logs ? logs.get(emailLower) : null;
        const manualLog = manualLogs ? manualLogs.get(String(devId)) : null;

        devMap.set(devId, {
          developer_id: devId,
          full_name: fullName,
          email: email,
          status: status,
          is_sent: !!log,
          sent_count: log ? log.count : 0,
          is_manual_sent: !!manualLog,
          manual_sent_count: manualLog ? manualLog.count : 0,
          manual_sent_note: manualLog ? manualLog.note : null,
          candidate_status: candidateStatus,
          agency_name: agencyName
        });
      } else {
        // Shouldn't usually happen now, but merge counts safely
        const existing = devMap.get(devId);
        const emailLower = String(email).toLowerCase().trim();
        const log = logs ? logs.get(emailLower) : null;
        const manualLog = manualLogs ? manualLogs.get(String(devId)) : null;

        existing.is_sent = existing.is_sent || !!log;
        existing.sent_count = Math.max(existing.sent_count || 0, log ? log.count : 0);
        existing.is_manual_sent = existing.is_manual_sent || !!manualLog;
        existing.manual_sent_count = Math.max(existing.manual_sent_count || 0, manualLog ? manualLog.count : 0);
        if (!existing.agency_name && agencyName) existing.agency_name = agencyName;
      }
    });

    const result = Array.from(devMap.values());
    console.log(`getDevelopers: Returning ${result.length} unique developers for Job ${cleanJobId}`);
    return result;

  } catch (e) {
    console.error("getDevelopers Error:", e);
    throw new Error(e.toString());
  }
}

function getEmailLogs(jobId) {
  const url = getStoredSheetUrl();
  if (!url) return null;
  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName("Email_Logs");
  if (!sheet) return new Map();
  const data = sheet.getDataRange().getValues();
  const logMap = new Map();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(jobId)) {
      const email = String(data[i][2]).toLowerCase();
      if (!logMap.has(email)) logMap.set(email, {count: 0, type: data[i][5]});
      logMap.get(email).count++;
    }
  }
  return logMap;
}

// UPDATED: Send Email with progress callback support + Job Details Sheet creation
function sendBulkEmails(recipients, senderName, subject, htmlBody, jobId, opts) {
  const url = getStoredSheetUrl();
  if(!url) return {success: false, sent: 0, errors: ["No config URL set. Please configure in Settings."]};

  const ss = SpreadsheetApp.openByUrl(url);
  ensureSheetsExist(ss);

  const logSheet = ss.getSheetByName("Email_Logs");
  const stateSheet = ss.getSheetByName("Negotiation_State");

  if(!logSheet || !stateSheet) {
    return {success: false, sent: 0, errors: ["Required sheets not found. Please re-save your config."]};
  }

  // CREATE JOB-SPECIFIC DETAILS SHEET if not exists
  // This analyzes the outreach email to determine what questions are being asked
  try {
    const sheetResult = getOrCreateJobDetailsSheet(ss, jobId, htmlBody);
    if(sheetResult.isNew) {
      console.log(`Created new job details sheet: Job_${jobId}_Details with ${sheetResult.questions?.length || 0} question columns`);
    }
  } catch(sheetError) {
    console.error("Failed to create job details sheet:", sheetError);
    // Don't fail the whole operation, just log the error
  }

  // Build existing emails set to check for duplicates
  const stateData = stateSheet.getDataRange().getValues();
  const existingEmails = new Set();
  for(let i=1; i<stateData.length; i++) {
    if(stateData[i][0]) {
      existingEmails.add(String(stateData[i][0]).toLowerCase() + '_' + String(stateData[i][1]));
    }
  }

  const labelName = `Job-${jobId}`;
  const labelId = getOrCreateLabelId(labelName);

  let count = 0;
  let errors = [];
  let skipped = 0;
  const total = recipients.length;

  recipients.forEach((r, index) => {
    try {
      const emailKey = String(r.email).toLowerCase() + '_' + String(jobId);

      // Check if already in system
      if(existingEmails.has(emailKey)) {
        skipped++;
        return;
      }

      const body = htmlBody.replace(/{{name}}/gi, r.name.split(' ')[0]);
      const rawMessage = createMimeMessage(senderName, r.email, subject, body);
      const message = Gmail.Users.Messages.send({ raw: rawMessage }, 'me');
      const threadId = message.threadId;

      if (labelId) {
        Gmail.Users.Threads.modify({ addLabelIds: [labelId] }, 'me', threadId);
      }

      // Log the email
      logSheet.appendRow([new Date(), jobId, r.email, r.name, threadId, "Initial"]);

      // Add to state with thread ID and Region (column 11)
      const region = r.region || '';
      stateSheet.appendRow([r.email, jobId, 0, "Initial Sent", "Initial Outreach", new Date(), r.devId || "N/A", r.name, "", threadId, region]);
      existingEmails.add(emailKey);

      // Add candidate to job details sheet with initial status and region
      try {
        const initialAnswers = { is_negotiating: false, negotiation_notes: '' };
        saveJobCandidateDetails(ss, jobId, r.email, r.name, r.devId || 'N/A', threadId, initialAnswers, 'Outreach Sent', region);
      } catch(detailsError) {
        console.error("Failed to add candidate to details sheet:", detailsError);
      }

      // Add to follow-up queue for automated 12hr and 28hr follow-ups
      try {
        addToFollowUpQueue(r.email, jobId, threadId, r.name, r.devId || 'N/A');
      } catch(followUpError) {
        console.error("Failed to add to follow-up queue:", followUpError);
      }

      count++;
    } catch(e) {
      console.error(e);
      errors.push(`Failed for ${r.email}: ${e.message}`);
    }
  });

  // Log data consumption for tracking
  try {
    const dataSize = JSON.stringify(recipients).length;
    logDataConsumption('Gmail', `Job-${jobId}`, dataSize, 0, `Sent ${count} emails, skipped ${skipped}`);
  } catch(logError) {
    console.error("Failed to log data consumption:", logError);
  }

  // Invalidate cache after adding new tasks
  if(count > 0) {
    invalidateSheetCache('Negotiation_State');
  }

  return {success: true, sent: count, skipped: skipped, total: total, errors: errors};
}

function createMimeMessage(senderName, recipientEmail, subject, htmlBody) {
  const userEmail = Session.getActiveUser().getEmail();
  const nl = "\r\n";
  const encodedSubject = "=?utf-8?B?" + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + "?=";
  const encodedSender = "=?utf-8?B?" + Utilities.base64Encode(senderName, Utilities.Charset.UTF_8) + "?=";
  let mime = `From: ${encodedSender} <${userEmail}>${nl}`;
  mime += `To: ${recipientEmail}${nl}`;
  mime += `Subject: ${encodedSubject}${nl}`;
  mime += `MIME-Version: 1.0${nl}`;
  mime += `Content-Type: text/html; charset=UTF-8${nl}${nl}`;
  mime += `${htmlBody}${nl}`;
  return Utilities.base64EncodeWebSafe(mime, Utilities.Charset.UTF_8);
}

function getOrCreateLabelId(name) {
  try {
    const list = Gmail.Users.Labels.list('me');
    const existing = list.labels.find(l => l.name === name);
    if(existing) return existing.id;
    const created = Gmail.Users.Labels.create({ name: name }, 'me');
    return created.id;
  } catch(e) { return null; }
}

/**
 * Send a reply to a thread with custom sender name
 * This ensures AI replies show as "Recruiter" not the actual email
 */
function sendReplyWithSenderName(thread, replyBody, senderName) {
  try {
    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    const subject = lastMessage.getSubject();
    const threadId = thread.getId();
    
    // Get recipient email (the person we're replying to)
    const fromHeader = lastMessage.getFrom();
    const emailMatch = fromHeader.match(/<([^>]+)>/);
    const recipientEmail = emailMatch ? emailMatch[1] : fromHeader.replace(/.*<|>.*/g, '');
    
    // Get Message-ID for threading
    const messageId = lastMessage.getHeader('Message-ID');
    
    const userEmail = Session.getActiveUser().getEmail();
    const nl = "\r\n";
    
    // Encode sender name and subject for UTF-8
    const encodedSender = "=?utf-8?B?" + Utilities.base64Encode(senderName || 'Recruiter', Utilities.Charset.UTF_8) + "?=";
    const replySubject = subject.startsWith('Re:') ? subject : 'Re: ' + subject;
    const encodedSubject = "=?utf-8?B?" + Utilities.base64Encode(replySubject, Utilities.Charset.UTF_8) + "?=";
    
    // Build MIME message with proper threading headers
    let mime = `From: ${encodedSender} <${userEmail}>${nl}`;
    mime += `To: ${recipientEmail}${nl}`;
    mime += `Subject: ${encodedSubject}${nl}`;
    mime += `In-Reply-To: ${messageId}${nl}`;
    mime += `References: ${messageId}${nl}`;
    mime += `MIME-Version: 1.0${nl}`;
    mime += `Content-Type: text/html; charset=UTF-8${nl}${nl}`;
    mime += `${replyBody.replace(/\n/g, '<br>')}${nl}`;
    
    const rawMessage = Utilities.base64EncodeWebSafe(mime, Utilities.Charset.UTF_8);
    
    // Send as reply to existing thread
    Gmail.Users.Messages.send({
      raw: rawMessage,
      threadId: threadId
    }, 'me');
    
    return true;
  } catch(e) {
    console.error("Failed to send reply with sender name:", e);
    // Fallback to regular reply
    thread.reply(replyBody);
    return false;
  }
}

// ======================================================
// ===    AUTOMATED NEGOTIATION AGENT (ENHANCED)      ===
// ======================================================

function runAutoNegotiator() {
  const url = getStoredSheetUrl();
  if(!url) return {status: "Error", message: "No Config URL. Please set your Google Sheets URL in Config.", stats: null, log: []};
  
  const ss = SpreadsheetApp.openByUrl(url);
  ensureSheetsExist(ss);
  
  const configSheet = ss.getSheetByName('Negotiation_Config');
  if(!configSheet) {
    return {status: "Error", message: "Negotiation_Config sheet not found", stats: null, log: []};
  }
  
  const configs = configSheet.getDataRange().getValues();
  const faqContent = getFAQs();
  
  let stats = { replied: 0, escalated: 0, accepted: 0, skipped: 0, processed: 0, synced: 0, cleaned: 0, detailsExtracted: 0 };
  let log = [];
  
  // STEP 1: Sync completed items from Gmail first (saves processing cost!)
  log.push({type: 'info', message: 'Syncing completed items from Gmail...'});
  try {
    const syncResult = syncCompletedFromGmail();
    stats.synced = syncResult.synced;
    stats.cleaned = syncResult.cleaned;
    if(syncResult.synced > 0) {
      log.push({type: 'success', message: `Synced ${syncResult.synced} completed items from Gmail`});
    }
    if(syncResult.cleaned > 0) {
      log.push({type: 'info', message: `Cleaned ${syncResult.cleaned} conflicting labels`});
    }
  } catch(e) {
    log.push({type: 'warning', message: 'Gmail sync skipped: ' + e.message});
  }
  
  // STEP 2: Cleanup any conflicting labels (Completed + Human-Negotiation)
  try {
    const cleanupResult = cleanupConflictingLabels();
    if(cleanupResult.cleaned > 0) {
      stats.cleaned += cleanupResult.cleaned;
      log.push({type: 'info', message: `Removed Human-Negotiation label from ${cleanupResult.cleaned} completed threads`});
    }
  } catch(e) {}
  
  if(configs.length <= 1) {
    return {status: "Error", message: "No job configurations found. Please configure at least one job.", stats: stats, log: log};
  }
  
  // STEP 3: Process negotiations for each job
  for(let i=1; i<configs.length; i++) {
    const jobId = configs[i][0];
    if(!jobId) continue; 
    
    log.push({type: 'info', message: `Processing Job ${jobId}...`});
    
    // UPDATED: Removed walkaway from rules
    const rules = {
      target: configs[i][1],
      max: configs[i][2],
      style: configs[i][3],
      special: configs[i][4],
      jobDescription: configs[i][5] || ''
    };
    
    let jobResult = processJobNegotiations(jobId, rules, ss, faqContent);
    
    stats.replied += jobResult.replied;
    stats.escalated += jobResult.escalated;
    stats.accepted += jobResult.accepted;
    stats.skipped += jobResult.skipped;
    stats.processed += jobResult.processed;
    stats.detailsExtracted += jobResult.detailsExtracted || 0;

    jobResult.log.forEach(l => log.push(l));
    log.push({type: 'success', message: `Job ${jobId} complete: ${jobResult.processed} threads processed, ${jobResult.detailsExtracted || 0} details extracted`});
  }
  
  return {status: "Success", stats: stats, log: log};
}

function processJobNegotiations(jobId, rules, ss, faqContent) {
  const query = `label:Job-${jobId}`;
  let threads = [];

  try {
    threads = GmailApp.search(query, 0, 50);
  } catch(e) {
    return {replied:0, escalated:0, accepted:0, skipped:0, processed:0, log:[{type:'error', message:`Gmail search failed for Job ${jobId}: ${e.message}`}]};
  }

  // Warn if no threads found - may indicate job ID mismatch
  if (threads.length === 0) {
    return {
      replied:0, escalated:0, accepted:0, skipped:0, processed:0, detailsExtracted:0,
      log:[{type:'warning', message:`No email threads found with label "Job-${jobId}". Check that the Job ID in Negotiation_Config matches the Gmail labels.`}]
    };
  }

  const stateSheet = ss.getSheetByName('Negotiation_State');
  const taskSheet = ss.getSheetByName('Negotiation_Tasks');

  let jobStats = {replied:0, escalated:0, accepted:0, skipped:0, processed:0, detailsExtracted:0, log:[]};
  
  // Cache state data for efficiency
  const stateData = stateSheet.getDataRange().getValues();
  const stateMap = new Map();
  for(let r=1; r<stateData.length; r++) {
    const key = String(stateData[r][0]).toLowerCase() + '_' + String(stateData[r][1]);
    stateMap.set(key, {
      rowIndex: r + 1,
      attempts: Number(stateData[r][2]) || 0,
      status: stateData[r][4],
      name: stateData[r][7] || 'Unknown',
      devId: stateData[r][6] || 'N/A',
      region: stateData[r][10] || '' // Column 11 = Region
    });
  }
  
  // Get current user email
  let myEmail = '';
  try {
    myEmail = Session.getActiveUser().getEmail();
    if(!myEmail) {
      myEmail = Gmail.Users.getProfile('me').emailAddress;
    }
  } catch(e) {
    try {
      myEmail = Gmail.Users.getProfile('me').emailAddress;
    } catch(e2) {
      jobStats.log.push({type: 'error', message: 'Could not determine user email'});
    }
  }
  myEmail = (myEmail || '').toLowerCase();
  
  jobStats.log.push({type: 'info', message: `Processing as: ${myEmail || 'unknown'}`});
  
  threads.forEach(thread => {
    jobStats.processed++;

    const labels = thread.getLabels().map(l => l.getName());

    if (labels.includes("Completed")) {
      jobStats.skipped++;
      jobStats.log.push({type: 'info', message: `Skipped thread: Already completed`});
      return;
    }

    const msgs = thread.getMessages();
    const lastMsg = msgs[msgs.length - 1];
    const lastSender = lastMsg.getFrom().toLowerCase();

    if (myEmail && myEmail.length > 3 && lastSender.indexOf(myEmail) > -1) {
      jobStats.skipped++;
      jobStats.log.push({type: 'info', message: `Skipped thread: Waiting for candidate reply (last message from ${myEmail})`});
      return;
    }

    // Also check for common sender names used by our system
    const ourSenderNames = ['recruiter', 'turing recruitment', 'turing team'];
    const isSentByUs = ourSenderNames.some(name => lastSender.includes(name));
    if (isSentByUs) {
      jobStats.skipped++;
      jobStats.log.push({type: 'info', message: `Skipped thread - last message from our system`});
      return;
    }

    // Prevent duplicate replies: Check if we sent a message in the last 5 minutes
    if (msgs.length > 1) {
      const recentMessages = msgs.slice(-3); // Check last 3 messages
      const fiveMinutesAgo = new Date(Date.now() - 5 * 60 * 1000);

      for (const msg of recentMessages) {
        const msgSender = msg.getFrom().toLowerCase();
        const msgDate = msg.getDate();
        const wasFromUs = (myEmail && msgSender.indexOf(myEmail) > -1) ||
                          ourSenderNames.some(name => msgSender.includes(name));

        if (wasFromUs && msgDate > fiveMinutesAgo) {
          jobStats.skipped++;
          jobStats.log.push({type: 'info', message: `Skipped thread - we replied recently (${Math.round((Date.now() - msgDate.getTime()) / 1000)}s ago)`});
          return;
        }
      }
    }
    
    // Extract candidate email
    const candidateEmailMatch = lastMsg.getFrom().match(/<([^>]+)>/);
    const candidateEmail = candidateEmailMatch ? candidateEmailMatch[1] : lastMsg.getFrom().replace(/.*<|>.*/g, '');
    const cleanCandidateEmail = candidateEmail.toLowerCase().trim();
    
    // Get state from cache
    const stateKey = cleanCandidateEmail + '_' + String(jobId);
    const state = stateMap.get(stateKey);

    let stateRowIndex = state ? state.rowIndex : -1;
    let attempts = state ? state.attempts : 0;
    let candidateName = state ? state.name : 'Unknown';
    let devId = state ? state.devId : 'N/A';
    let candidateRegion = state ? state.region : '';

    // Build conversation history FIRST (needed for escalation and data extraction)
    const recentMsgs = msgs.slice(-5);
    const conversationHistory = recentMsgs.map(m => {
      const from = m.getFrom();
      const isMe = from.toLowerCase().indexOf(myEmail) > -1;
      return `[${isMe ? 'ME' : 'CANDIDATE'}]: ${m.getPlainBody().substring(0, 400)}`;
    }).join("\n---\n");

    // ALWAYS extract and save candidate details BEFORE any early returns
    // This ensures we capture candidate data even during ongoing negotiation or escalation
    const candidateLatestMessage = lastMsg.getPlainBody();
    try {
      // Get the questions configured for this job
      const questions = getJobQuestions(jobId);

      // Extract answers from candidate's message based on the questions
      const answers = extractAnswersFromResponse(candidateLatestMessage, questions, candidateName);

      // Determine the appropriate status based on current state
      let dataStatus = 'Negotiating';
      if (state && state.status === 'Human-Negotiation') {
        // Preserve human negotiation status - data still gets updated
        dataStatus = 'Human-Negotiation';
      } else if (attempts >= 2) {
        // Will be escalated, but still save the data first
        dataStatus = 'Pending Escalation';
      }

      // Save to job-specific details sheet
      const saveResult = saveJobCandidateDetails(ss, jobId, candidateEmail, candidateName, devId, thread.getId(), answers, dataStatus);
      if(saveResult.success) {
        jobStats.detailsExtracted++;
        jobStats.log.push({type: 'info', message: `${candidateEmail} - Details extracted: ${answers.is_negotiating ? 'Negotiating' : 'Provided'}`});
      }
    } catch(detailsError) {
      // Don't block negotiation if details extraction fails
      console.error("Details extraction error for " + candidateEmail + ":", detailsError);
    }

    // Now handle special cases AFTER data extraction
    if (state && state.status === 'Human-Negotiation') {
      jobStats.skipped++;
      jobStats.log.push({type: 'info', message: `Skipped AI negotiation for ${cleanCandidateEmail}: Already escalated to human (data was extracted)`});
      return;
    }

    // Check attempt limit (2 AI attempts max)
    if (attempts >= 2) {
      // Generate AI summary notes before escalating
      const aiSummaryNotes = generateAISummaryNotes(conversationHistory, candidateEmail, '', attempts, "Max AI attempts reached - candidate did not agree to offered rate");

      escalateToHuman(thread, "Max AI attempts reached", candidateName, `We've had ${attempts} negotiation rounds. ${aiSummaryNotes}`);
      if(stateRowIndex > -1) {
        stateSheet.getRange(stateRowIndex, 5).setValue("Human-Negotiation");
        stateSheet.getRange(stateRowIndex, 9).setValue(aiSummaryNotes);
      }

      // Update job details sheet with escalation status (data already saved above)
      try {
        updateJobCandidateStatus(ss, jobId, candidateEmail, 'Escalated: Max AI attempts', null);
      } catch(detailsErr) {
        console.error("Failed to update job details sheet:", detailsErr);
      }

      jobStats.escalated++;
      jobStats.log.push({type: 'warning', message: `${candidateEmail} escalated: Max attempts reached (data was extracted)`});
      return;
    }

    const isFirstResponse = attempts === 0;

    // Calculate offer amounts based on attempt
    // Use region-specific rates if available, otherwise fall back to job config rates
    let targetRate = Number(rules.target) || 25;
    let maxRate = Number(rules.max) || 30;

    if (candidateRegion) {
      const regionRates = getRateForRegion(jobId, candidateRegion, ss);
      if (regionRates) {
        targetRate = regionRates.targetRate || targetRate;
        maxRate = regionRates.maxRate || maxRate;
        jobStats.log.push({type: 'info', message: `${candidateEmail} - Using ${regionRates.region || candidateRegion} region rates: target=$${targetRate}, max=$${maxRate}`});
      }
    }

    // Progressive offer strategy:
    // Attempt 1 (attempts=0): 80% of target rate
    // Attempt 2 (attempts=1): 100% of target rate
    // Attempt 3+ (attempts>=2): Human escalation
    const firstOfferRate = Math.round(targetRate * 0.8); // 80% of target for first offer
    const secondOfferRate = targetRate; // 100% of target for second offer
    const currentOfferRate = attempts === 0 ? firstOfferRate : secondOfferRate;

    // AI-POWERED RATE ANALYSIS: Use AI to intelligently analyze developer's message and decide action
    const rateAnalysisPrompt = `
You are analyzing a candidate's message to determine their rate expectation and recommend an action.

CANDIDATE'S MESSAGE:
"${candidateLatestMessage}"

OUR RATE PARAMETERS:
- Target Rate: $${targetRate}/hr (maximum we want to pay)
- First Offer Rate: $${firstOfferRate}/hr (80% of target - for first attempt)
- Current Attempt: ${attempts + 1}

TASK:
Analyze the candidate's message and determine:
1. What hourly rate are they proposing, expecting, or asking for? (extract the exact number)
2. Are they accepting a previous offer we made?
3. What action should we take?

IMPORTANT INTERPRETATION RULES:
- "Can I get $X?" = They are asking for $X (treat as a rate proposal of $X)
- "I would like $X" = They are proposing $X
- "I expect $X" = They are proposing $X
- "I want $X" = They are proposing $X
- "How about $X?" = They are counter-proposing $X
- ANY mention of a specific dollar amount by the candidate = Their proposed/expected rate
- Questions about rate ("can I get...?", "is $X possible?") should be treated as rate proposals

DECISION RULES:
- If candidate mentions/asks for a rate AT OR BELOW $${targetRate}/hr → ACTION: AUTO_ACCEPT (accept their rate!)
- If candidate explicitly ACCEPTS an offer we made → ACTION: AUTO_ACCEPT
- If candidate mentions/asks for a rate ABOVE $${targetRate}/hr → ACTION: COUNTER (we will counter with target rate)
- If candidate is very firm on a rate ABOVE $${targetRate}/hr after multiple attempts → ACTION: ESCALATE
- If no clear rate mentioned → ACTION: COUNTER

CRITICAL:
- If they ask "can I get $41?" and our target is $50, that means $41 <= $50, so AUTO_ACCEPT at $41
- If they ask "can I get $41?" and our target is $40, that means $41 > $40, so COUNTER
- Always extract the EXACT rate number they mentioned

RESPONSE FORMAT (JSON only):
{
  "proposed_rate": <number - the rate they mentioned/asked for, or null if none>,
  "is_accepting_offer": <true/false>,
  "action": "<AUTO_ACCEPT|COUNTER|ESCALATE>",
  "reason": "<brief explanation including the comparison: their rate vs our target>",
  "candidate_flexibility": "<flexible|firm|unclear>"
}

Return ONLY the JSON object, no other text.
`;

    let rateAnalysis = null;
    try {
      const analysisResponse = callAI(rateAnalysisPrompt);
      const cleanResponse = analysisResponse.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      rateAnalysis = JSON.parse(cleanResponse);
      jobStats.log.push({type: 'info', message: `${candidateEmail} - AI rate analysis: ${rateAnalysis.action} (rate: $${rateAnalysis.proposed_rate || 'not specified'}, reason: ${rateAnalysis.reason})`});
    } catch(e) {
      console.error("Failed to parse rate analysis:", e);
      // Continue with normal negotiation if analysis fails
    }

    // If AI recommends AUTO_ACCEPT (rate at or below target, or accepting our offer)
    if (rateAnalysis && rateAnalysis.action === 'AUTO_ACCEPT') {
      const rate = rateAnalysis.proposed_rate || targetRate;
      jobStats.log.push({type: 'success', message: `${candidateEmail} - AI recommends AUTO-ACCEPT at $${rate}/hr (target: $${targetRate}/hr)`});

      // Update job-specific details sheet with accepted status and rate
      try {
        updateJobCandidateStatus(ss, jobId, candidateEmail, 'Offer Accepted', `$${rate}/hr`);
      } catch(detailsErr) {
        console.error("Failed to update job details sheet:", detailsErr);
      }

      // Extract job details for the acceptance email
      const jobDescription = rules.jobDescription || '';

      // Send detailed acceptance confirmation email with project info
      const acceptPrompt = `
You are a recruiter at Turing. Write a professional acceptance confirmation email to ${candidateName.split(' ')[0]}.

CANDIDATE NAME: ${candidateName.split(' ')[0]}
AGREED RATE: $${rate}/hr

JOB DESCRIPTION FOR CONTEXT:
${jobDescription}

TASK:
Write an email that includes ALL of the following points:
1. Thank them for sharing their rate alignment
2. Confirm you are sharing all the details with the team
3. Acknowledge the project details from the JD (mention if it's remote, hours per week if specified, etc.)
4. Inform them their profile is currently under client review
5. If approved and selected, you will reach out to confirm the onboarding date and provide further details along with contract specifics

IMPORTANT:
- Extract work arrangement details from the JD (remote/hybrid, hours per week, etc.)
- If hours not specified in JD, mention "full-time commitment" or "as per project requirements"
- Keep the tone warm and professional
- DO NOT make up details not in the JD

FORMAT:
Hi [First Name],

Thank you for sharing your alignment on the rate. I am sharing all the details with the team.

Please acknowledge that the project your profile is being considered for is [extract from JD: remote/location, hours per week if mentioned].

Please be aware that your profile is currently under client review. If approved and selected, we will reach out to confirm the onboarding date and provide further details along with contract specifics.

Best regards,
Turing Recruitment Team

Write ONLY the email, nothing else.
`;
      const acceptEmail = callAI(acceptPrompt);
      sendReplyWithSenderName(thread, acceptEmail, 'Recruiter');
      markCompleted(thread);

      // Record directly in Negotiation_Completed (auto-completed, not pending)
      const compSheet = ss.getSheetByName('Negotiation_Completed');
      if (compSheet) {
        compSheet.appendRow([
          new Date(),
          jobId,
          candidateEmail,
          candidateName,
          `Offer Accepted at $${rate}/hr`,
          'AI Auto-Accepted - Email sent',
          devId,
          candidateRegion || ''
        ]);
      }

      // Remove from state sheet
      if(stateRowIndex > -1) {
        stateSheet.deleteRow(stateRowIndex);
        stateMap.delete(stateKey);
      }

      // Invalidate caches so UI reflects the change immediately
      invalidateSheetCache('Negotiation_State');
      invalidateSheetCache('Negotiation_Completed');

      jobStats.accepted++;
      jobStats.log.push({type: 'success', message: `${candidateEmail} AUTO-ACCEPTED at $${rate}/hr - Completed`});
      return; // Skip the rest of negotiation logic
    }

    // If AI recommends ESCALATE (candidate firm on rate above target)
    if (rateAnalysis && rateAnalysis.action === 'ESCALATE' && attempts >= 1) {
      const escalationReason = rateAnalysis.reason || 'Candidate firm on rate above target';
      jobStats.log.push({type: 'warning', message: `${candidateEmail} - AI recommends ESCALATE: ${escalationReason}`});

      // Generate summary for human handoff
      const aiSummaryNotes = generateAISummaryNotes(conversationHistory, candidateEmail, '', attempts, escalationReason);

      // Escalate to human with detailed handoff
      escalateToHuman(thread, escalationReason, candidateName, aiSummaryNotes);

      if(stateRowIndex > -1) {
        stateSheet.getRange(stateRowIndex, 5).setValue("Human-Negotiation");
        stateSheet.getRange(stateRowIndex, 9).setValue(aiSummaryNotes);
      }

      // Update job-specific details sheet with escalation status
      try {
        updateJobCandidateStatus(ss, jobId, candidateEmail, `Escalated: ${escalationReason}`, null);
      } catch(detailsErr) {
        console.error("Failed to update job details sheet:", detailsErr);
      }

      jobStats.escalated++;
      jobStats.log.push({type: 'warning', message: `${candidateEmail} escalated: ${escalationReason}`});
      return; // Skip the rest of negotiation logic
    }

    // CRITICAL SAFEGUARD: Never offer more than what the candidate is asking for
    // If candidate asks for $41 and our scheduled offer is $50, we should accept $41 (not offer $50!)
    const candidateProposedRate = rateAnalysis ? rateAnalysis.proposed_rate : null;
    if (candidateProposedRate !== null && candidateProposedRate < currentOfferRate) {
      // Candidate is asking for LESS than what we were going to offer - accept their rate!
      jobStats.log.push({type: 'success', message: `${candidateEmail} - Candidate asking $${candidateProposedRate}/hr which is LESS than our offer of $${currentOfferRate}/hr - accepting their lower rate!`});

      const rate = candidateProposedRate;
      taskSheet.appendRow([new Date(), jobId, candidateName, candidateEmail, rate, "Pending Archive", devId, thread.getId()]);

      try {
        updateJobCandidateStatus(ss, jobId, candidateEmail, 'Offer Accepted', `$${rate}/hr`);
      } catch(detailsErr) {
        console.error("Failed to update job details sheet:", detailsErr);
      }

      const jobDescription = rules.jobDescription || '';
      const acceptPrompt = `
You are a recruiter at Turing. Write a professional acceptance confirmation email to ${candidateName.split(' ')[0]}.

CANDIDATE NAME: ${candidateName.split(' ')[0]}
AGREED RATE: $${rate}/hr

JOB DESCRIPTION FOR CONTEXT:
${jobDescription}

TASK:
Write an email that includes ALL of the following points:
1. Thank them for sharing their rate alignment
2. Confirm you are sharing all the details with the team
3. Acknowledge the project details from the JD (mention if it's remote, hours per week if specified, etc.)
4. Inform them their profile is currently under client review
5. If approved and selected, you will reach out to confirm the onboarding date and provide further details along with contract specifics

FORMAT:
Hi [First Name],

Thank you for sharing your alignment on the rate. I am sharing all the details with the team.

Please acknowledge that the project your profile is being considered for is [extract from JD: remote/location, hours per week if mentioned].

Please be aware that your profile is currently under client review. If approved and selected, we will reach out to confirm the onboarding date and provide further details along with contract specifics.

Best regards,
Turing Recruitment Team

Write ONLY the email, nothing else.
`;
      const acceptEmail = callAI(acceptPrompt);
      sendReplyWithSenderName(thread, acceptEmail, 'Recruiter');
      markCompleted(thread);

      if(stateRowIndex > -1) {
        stateSheet.deleteRow(stateRowIndex);
        stateMap.delete(stateKey);
      }

      jobStats.accepted++;
      jobStats.log.push({type: 'success', message: `${candidateEmail} ACCEPTED at $${rate}/hr (candidate asked less than our offer)`});
      return;
    }

    // ENHANCED AI PROMPT with better negotiation strategy (NEVER reveal target!)
    const prompt = `
You are a recruiter at Turing negotiating a rate for Job ID ${jobId}.

=== ABOUT TURING ===
Turing is one of the world's fastest-growing AI companies accelerating the advancement and deployment of powerful AI systems.
Turing helps customers in two ways: Working with the world's leading AI labs to advance frontier model capabilities in thinking, reasoning, coding, agentic behavior, multimodality, multilinguality, STEM and frontier knowledge; and leveraging that work to build real-world AI systems that solve mission-critical priorities for companies.

Perks of Freelancing With Turing:
- Work in a fully remote environment
- Opportunity to work on cutting-edge AI projects with leading LLM companies
- Flexible freelance arrangement

=== JOB DESCRIPTION ===
${rules.jobDescription || 'No specific job description provided.'}

=== NEGOTIATION RULES ===
${attempts === 0 ? `
**FIRST ATTEMPT - Start at 80% of Target:**
- YOUR OFFER: $${firstOfferRate}/hr - This is what you MUST offer (80% of our target)
- DO NOT offer anything higher than $${firstOfferRate}/hr on this first attempt, no matter what rate they ask for
- Even if they request $100/hr, you respond with YOUR offer of $${firstOfferRate}/hr
- Be confident and direct: "We can offer $${firstOfferRate}/hr for this role"
- NEVER mention any "target rate" or "budget" or "aim for" - just state your offer confidently
` : `
**SECOND ATTEMPT - Full Target Rate:**
- YOUR OFFER: $${secondOfferRate}/hr - This is what you MUST offer (our full target rate)
- DO NOT offer anything higher than $${secondOfferRate}/hr, no matter what
- If they don't accept $${secondOfferRate}/hr, that's okay - we will escalate to human
- Be confident and direct: "We can offer $${secondOfferRate}/hr for this role"
- NEVER mention any "target rate" or "budget" - just state what you can offer
- DO NOT offer max rate - only humans can approve rates above $${secondOfferRate}/hr
`}

- Negotiation Style: ${rules.style}
- Special Instructions: ${rules.special || 'None'}
- Current Attempt: ${attempts + 1} of 2 (human escalation after attempt 2)

=== CRITICAL RULES - READ CAREFULLY ===
1. **NEVER reveal internal numbers**: Do not say "we aim for", "our target is", "our budget is", or "we're looking at". Just state your offer directly.
2. **Be confident**: Say "We can offer $X/hr for this role" - not "We're hoping for" or "We'd like to offer"
3. **This is FREELANCE**: Never mention full-time benefits, team culture, or long-term employment
4. **Answer questions on NEW LINES**: If answering multiple questions, put each answer on a separate line for readability

=== FREQUENTLY ASKED QUESTIONS (Reference Only) ===
${faqContent}

**IMPORTANT FAQ INSTRUCTIONS:**
- ONLY use these FAQs if the candidate EXPLICITLY asks a question
- Do NOT proactively volunteer information - only answer what they ask
- If they didn't ask any questions, do NOT include any FAQ answers in your response
- "Let me know if you need anything else" is NOT a question - it's a polite closing
- If they DO ask a question that matches an FAQ:
  1. Use the provided answer as your source of truth
  2. Paraphrase naturally in your own words
  3. Put each answer on a SEPARATE LINE for better readability

=== CONVERSATION HISTORY ===
${conversationHistory}

=== EMAIL FORMATTING RULES ===
1. Start with a warm greeting: "Hi [First Name],"
2. Keep the main message to 2-3 short paragraphs
3. If answering multiple questions, format like this:
   
   Regarding your questions:
   
   • [Answer to first question]
   
   • [Answer to second question]
   
4. End with a clear call to action
5. Sign off professionally

=== RESPONSE INSTRUCTIONS ===
${isFirstResponse ? `
THIS IS THE FIRST RESPONSE - Offer $${firstOfferRate}/hr (MANDATORY - 80% of target)
- You MUST offer exactly $${firstOfferRate}/hr - this is non-negotiable for the first attempt
- Present this rate confidently without justification
- If they asked for a higher rate (even $100/hr), acknowledge their experience but offer $${firstOfferRate}/hr
- ONLY answer questions they EXPLICITLY asked - do NOT volunteer information they didn't request
- NEVER say "we aim for" or "our target is" or reveal any internal numbers
- NEVER use ACTION: ESCALATE on first attempt
` : `
THIS IS ATTEMPT ${attempts + 1} - Offer $${secondOfferRate}/hr (MANDATORY - full target rate)
- You MUST offer exactly $${secondOfferRate}/hr - this is your final AI offer
- DO NOT offer anything higher - max rate ($${maxRate}/hr) can only be approved by humans
- If they don't accept $${secondOfferRate}/hr, that's fine - we will escalate to human
- ONLY answer questions they EXPLICITLY asked - do NOT volunteer information they didn't request
- You may escalate if they explicitly refuse $${secondOfferRate}/hr
`}

=== RESPONSE FORMAT ===
1. If they clearly ACCEPT an offer at or below $${secondOfferRate}/hr:
   Reply with: ACTION: ACCEPT [$RATE]

2. If this is attempt 2 AND they refuse $${secondOfferRate}/hr or demand higher:
   Reply with: ACTION: ESCALATE [REASON: brief reason here]

3. Otherwise, write a professional email:
   
   Hi [First Name],

   [Opening - acknowledge their message/questions]

   [Your offer or counter-offer - state confidently without revealing targets]

   [If answering questions, put on separate lines]

   [Call to action]

   Best regards,
   Turing Recruitment Team

Respond with ONLY the email text OR the ACTION code. No other explanations.
    `;
    
    const aiResponse = callAI(prompt);
    
    jobStats.log.push({type: 'info', message: `${candidateEmail} - AI response: ${aiResponse.substring(0, 50)}...`});
    
    // Extract escalation reason
    let escalationReason = "";
    const reasonMatch = aiResponse.match(/\[REASON:\s*([^\]]+)\]/i);
    if(reasonMatch) {
      escalationReason = reasonMatch[1].trim();
    }
    
    if(!escalationReason && aiResponse.includes("ACTION: ESCALATE")) {
      const afterEscalate = aiResponse.split("ACTION: ESCALATE")[1];
      if(afterEscalate && afterEscalate.trim().length > 0) {
        escalationReason = afterEscalate.trim().substring(0, 100);
      } else {
        escalationReason = "Rate above budget or complex question";
      }
    }
    
    // Handle AI response
    if (aiResponse.includes("ACTION: ESCALATE")) {
      if(attempts < 2) {
        // Force negotiation on early attempts
        jobStats.log.push({type: 'warning', message: `${candidateEmail} - Attempt ${attempts + 1}/2: AI wanted to escalate, forcing negotiation`});

        const candidateMessage = lastMsg.getPlainBody().substring(0, 500);

        // Use already-calculated region-specific rates (targetRate and maxRate are set above)
        // Always use target rate - don't start lower
        const currentOffer = targetRate;

        const retryPrompt = `
You are a recruiter at Turing. You MUST write a negotiation email - escalation is NOT an option.

CANDIDATE'S ACTUAL MESSAGE:
"${candidateMessage}"

CANDIDATE NAME: ${candidateName}

YOUR OFFER FOR THIS ATTEMPT: $${currentOffer}/hr (this is MANDATORY - do not offer higher)
${attempts === 0 ? `This is the first attempt (80% of target) - you MUST offer exactly $${currentOffer}/hr, no higher.` : `This is the second attempt (full target) - you MUST offer exactly $${currentOffer}/hr. DO NOT offer max rate - only humans can approve higher rates.`}

JOB CONTEXT:
${rules.jobDescription || 'Freelance AI/Tech role'}

FAQs (ONLY use if candidate explicitly asks a question):
${faqContent}

CRITICAL RULES:
1. NEVER say "we aim for", "our target is", "our budget is" - just state your offer confidently
2. Say "We can offer $${currentOffer}/hr for this role" - be direct and confident
3. ONLY answer questions the candidate EXPLICITLY asked - do NOT volunteer information
4. "Let me know if you need anything else" is NOT a question - it's a polite closing
5. This is FREELANCE - never mention full-time benefits

TASK:
1. Read the candidate's message above carefully
2. Note the ACTUAL rate they mentioned (if any)
3. ONLY answer questions they EXPLICITLY asked (not implied or anticipated)
4. Make your offer of $${currentOffer}/hr confidently
5. Keep the tone ${rules.style}

FORMAT:
Hi [First Name],

[Acknowledge their message]

[Your offer - state directly without revealing any targets]

[If they asked questions, answer each on a new line like:
• Answer to question 1
• Answer to question 2]

[Call to action - ask if they can proceed]

Best regards,
Turing Recruitment Team

Write ONLY the email, nothing else.
`;
        
        const retryResponse = callAI(retryPrompt);
        sendReplyWithSenderName(thread, retryResponse, 'Recruiter');
        
        const newAttemptCount = attempts + 1;
        
        // Generate contextual note about this negotiation attempt
        const attemptNotePrompt = `
Summarize this negotiation attempt in one short sentence (under 20 words):
- Candidate said: "${candidateMessage.substring(0, 200)}"
- We responded with a counter-offer
- This is attempt ${newAttemptCount} of 2

Write a brief note like: "Candidate asking $X, we countered with $Y" or "Dev wants higher rate, negotiating"
Write ONLY the note, nothing else.
`;
        let noteText = `Attempt ${newAttemptCount}: AI negotiated`;
        try {
          const aiNote = callAI(attemptNotePrompt);
          noteText = `Attempt ${newAttemptCount}: ${aiNote.substring(0, 100)}`;
        } catch(e) {}
        
        if(stateRowIndex > -1) {
          stateSheet.getRange(stateRowIndex, 3).setValue(newAttemptCount);
          stateSheet.getRange(stateRowIndex, 4).setValue("Counter Offer Sent");
          stateSheet.getRange(stateRowIndex, 5).setValue("Active");
          stateSheet.getRange(stateRowIndex, 6).setValue(new Date());
          stateSheet.getRange(stateRowIndex, 9).setValue(noteText);
        } else {
          stateSheet.appendRow([cleanCandidateEmail, jobId, newAttemptCount, "Counter Offer", "Active", new Date(), devId, candidateName, noteText, thread.getId()]);
        }
        
        jobStats.replied++;
        jobStats.log.push({type: 'info', message: `${candidateEmail} - AI negotiated (attempt ${newAttemptCount}/2)`});
        return;
      }
      
      // Allow escalation after 2 attempts - generate detailed summary
      const finalReason = escalationReason || "Candidate did not agree after 2 negotiation attempts";
      const aiSummaryNotes = generateAISummaryNotes(conversationHistory, candidateEmail, '', attempts, finalReason);

      escalateToHuman(thread, finalReason, candidateName, aiSummaryNotes);
      if(stateRowIndex > -1) {
        stateSheet.getRange(stateRowIndex, 5).setValue("Human-Negotiation");
        stateSheet.getRange(stateRowIndex, 9).setValue(aiSummaryNotes);
      }

      // Update job-specific details sheet with escalation status
      try {
        updateJobCandidateStatus(ss, jobId, candidateEmail, `Escalated: ${finalReason}`, null);
      } catch(detailsErr) {
        console.error("Failed to update job details sheet:", detailsErr);
      }

      jobStats.escalated++;
      jobStats.log.push({type: 'warning', message: `${candidateEmail} escalated: ${finalReason}`});
    }
    else if (aiResponse.includes("ACTION: ACCEPT")) {
      const rateMatch = aiResponse.match(/\[([^\]]+)\]/);
      const rate = rateMatch ? rateMatch[1].replace('$','').replace('/hr','').trim() : rules.target;

      // Update job-specific details sheet with accepted status and rate
      try {
        updateJobCandidateStatus(ss, jobId, candidateEmail, 'Offer Accepted', `$${rate}/hr`);
      } catch(detailsErr) {
        console.error("Failed to update job details sheet:", detailsErr);
      }

      // Extract job details for the acceptance email
      const jobDescription = rules.jobDescription || '';

      // Send detailed acceptance confirmation email with project info
      const acceptPrompt = `
You are a recruiter at Turing. Write a professional acceptance confirmation email to ${candidateName.split(' ')[0]}.

CANDIDATE NAME: ${candidateName.split(' ')[0]}
AGREED RATE: $${rate}/hr

JOB DESCRIPTION FOR CONTEXT:
${jobDescription}

TASK:
Write an email that includes ALL of the following points:
1. Thank them for sharing their rate alignment
2. Confirm you are sharing all the details with the team
3. Acknowledge the project details from the JD (mention if it's remote, hours per week if specified, etc.)
4. Inform them their profile is currently under client review
5. If approved and selected, you will reach out to confirm the onboarding date and provide further details along with contract specifics

IMPORTANT:
- Extract work arrangement details from the JD (remote/hybrid, hours per week, etc.)
- If hours not specified in JD, mention "full-time commitment" or "as per project requirements"
- Keep the tone warm and professional
- DO NOT make up details not in the JD

FORMAT:
Hi [First Name],

Thank you for sharing your alignment on the rate. I am sharing all the details with the team.

Please acknowledge that the project your profile is being considered for is [extract from JD: remote/location, hours per week if mentioned].

Please be aware that your profile is currently under client review. If approved and selected, we will reach out to confirm the onboarding date and provide further details along with contract specifics.

Best regards,
Turing Recruitment Team

Write ONLY the email, nothing else.
`;
      const acceptEmail = callAI(acceptPrompt);
      sendReplyWithSenderName(thread, acceptEmail, 'Recruiter');
      markCompleted(thread);

      // Record directly in Negotiation_Completed (auto-completed, not pending)
      const compSheet = ss.getSheetByName('Negotiation_Completed');
      if (compSheet) {
        compSheet.appendRow([
          new Date(),
          jobId,
          candidateEmail,
          candidateName,
          `Offer Accepted at $${rate}/hr`,
          'AI Accepted - Email sent',
          devId,
          candidateRegion || ''
        ]);
      }

      // Remove from state sheet
      if(stateRowIndex > -1) {
        stateSheet.deleteRow(stateRowIndex);
        stateMap.delete(stateKey);
      }

      // Invalidate caches so UI reflects the change immediately
      invalidateSheetCache('Negotiation_State');
      invalidateSheetCache('Negotiation_Completed');

      jobStats.accepted++;
      jobStats.log.push({type: 'success', message: `${candidateEmail} ACCEPTED at $${rate}/hr - Completed`});
    } 
    else {
      sendReplyWithSenderName(thread, aiResponse, 'Recruiter');
      
      const newAttemptCount = attempts + 1;
      
      // Generate contextual note about what happened in this exchange
      const candidateMessage = lastMsg.getPlainBody().substring(0, 400);
      const notePrompt = `
Analyze this negotiation exchange and write a specific, actionable note for the recruiter.

CANDIDATE'S MESSAGE:
"${candidateMessage}"

OUR AI RESPONSE (summary):
"${aiResponse.substring(0, 250)}"

This is negotiation attempt ${newAttemptCount} of 2.
Our offer in this round: approximately $${currentOfferRate}/hr

TASK:
Write a brief note (1-2 sentences, under 40 words) that captures:
1. The specific rate the candidate requested (if mentioned)
2. What we offered them
3. Any key concerns or questions they raised

EXAMPLES OF GOOD NOTES:
- "Candidate requested $35/hr. We offered $15/hr. They asked about work hours and laptop provision."
- "Dev wants $40/hr citing 5 years experience. Countered with $18/hr. Interested but firm on rate."
- "Candidate asked about project details before discussing rate. We offered $15/hr with role info."

EXAMPLES OF BAD NOTES (too vague):
- "AI negotiated"
- "Candidate responded"
- "Rate discussion ongoing"

Write ONLY the note, nothing else. Be specific about numbers.
`;
      
      let noteText = `Attempt ${newAttemptCount}: Offered $${currentOfferRate}/hr`;
      try {
        const aiNote = callAI(notePrompt);
        if(aiNote && aiNote.length > 10 && !aiNote.includes("ACTION:")) {
          noteText = `Attempt ${newAttemptCount}: ${aiNote.substring(0, 200)}`;
        }
      } catch(e) {
        // Keep default note on error
      }
      
      if(stateRowIndex > -1) {
        stateSheet.getRange(stateRowIndex, 3).setValue(newAttemptCount);
        stateSheet.getRange(stateRowIndex, 4).setValue("Counter Offer Sent");
        stateSheet.getRange(stateRowIndex, 5).setValue("Active");
        stateSheet.getRange(stateRowIndex, 6).setValue(new Date());
        stateSheet.getRange(stateRowIndex, 9).setValue(noteText);
      } else {
        stateSheet.appendRow([cleanCandidateEmail, jobId, newAttemptCount, "Counter Offer", "Active", new Date(), devId, candidateName, noteText, thread.getId()]);
      }
      
      jobStats.replied++;
      jobStats.log.push({type: 'info', message: `${candidateEmail} - AI negotiated (attempt ${newAttemptCount}/2)`});
    }
  });
  
  return jobStats;
}

function escalateToHuman(thread, reason, candidateName, conversationContext) {
  try {
    const label = GmailApp.getUserLabelByName("Human-Negotiation") || GmailApp.createLabel("Human-Negotiation");
    thread.addLabel(label);

    // Generate handoff message - inform candidate about talent ops taking over
    const firstName = candidateName ? candidateName.split(' ')[0] : 'there';

    const handoffPrompt = `
You are a recruiter at Turing. You need to write a handoff email to a candidate whose rate negotiation needs human attention.

CANDIDATE NAME: ${firstName}
CONTEXT: ${conversationContext || 'Rate negotiation needs human attention'}

TASK:
Write a professional email that clearly communicates:
1. Thank them for sharing their rate expectation
2. You have shared their message with the Talent Operations team member
3. The team will get back to them with an update on the rates
4. Keep a positive, professional tone

IMPORTANT POINTS TO INCLUDE:
- Acknowledge you received their message about rates
- Explicitly mention you are sharing this with a Talent Operations team member
- They will receive an update on the rate discussion shortly
- Thank them for their patience

FORMAT:
Hi ${firstName},

Thank you for sharing your rate expectation. I have shared your message with a member of our Talent Operations team, and they will get back to you shortly with an update on the rates.

We appreciate your patience and interest in this opportunity.

Best regards,
Turing Recruitment Team

Write ONLY the email, nothing else. Keep it concise (3-4 sentences).
`;

    const handoffMessage = callAI(handoffPrompt);
    sendReplyWithSenderName(thread, handoffMessage, 'Recruiter');

  } catch(e) {
    console.error("Failed to escalate to human:", e);
    // Fallback to simple reply if AI fails
    try {
      const fallbackMsg = `Hi ${candidateName ? candidateName.split(' ')[0] : 'there'},\n\nThank you for sharing your rate expectation. I have shared your message with a member of our Talent Operations team, and they will get back to you shortly with an update on the rates.\n\nWe appreciate your patience and interest in this opportunity.\n\nBest regards,\nTuring Recruitment Team`;
      sendReplyWithSenderName(thread, fallbackMsg, 'Recruiter');
    } catch(e2) {}
  }
}

// Generate contextual AI summary notes based on conversation
function generateAISummaryNotes(conversationHistory, candidateEmail, lastOffer, attempts, escalationReason) {
  const prompt = `
You are analyzing a recruitment negotiation conversation to write a brief, specific handoff note for a human recruiter.

CONVERSATION HISTORY:
${conversationHistory}

CONTEXT:
- Candidate Email: ${candidateEmail}
- Number of AI negotiation attempts: ${attempts}
- Escalation Reason: ${escalationReason || 'Max attempts reached'}

TASK:
Write a 2-3 sentence summary note that tells the human recruiter exactly what happened. Be SPECIFIC about:
1. What rate the candidate is asking for (exact number if mentioned)
2. What offers were made and rejected
3. Any specific concerns or questions the candidate raised
4. The candidate's tone/flexibility (firm, open to negotiation, frustrated, etc.)

DO NOT use generic phrases like "candidate did not agree" - instead say exactly what they asked for.

Examples of GOOD notes:
- "Dev is firm on $35/hr, rejected our $28 offer twice. Mentioned they have competing offers. Seems open to discussion on project scope."
- "Candidate asking for $40/hr citing 8 years experience. Showed interest in the role but won't budge below $38. Asked about benefits."
- "Dev concerned about hourly commitment (wants max 20hrs/week). Rate negotiation stalled at $30 vs our $25 offer."

Examples of BAD notes (too generic):
- "Candidate did not agree to rate"
- "Escalated after 2 attempts"
- "Rate negotiation unsuccessful"

Write ONLY the summary note, nothing else. Keep it under 50 words.
`;

  try {
    const response = callAI(prompt);
    // Clean up response - remove any quotes or extra formatting
    return response.replace(/^["']|["']$/g, '').trim();
  } catch(e) {
    // Fallback to basic note if AI fails
    return `Escalated after ${attempts} AI attempts. Reason: ${escalationReason || 'Max attempts reached'}`;
  }
}

function markCompleted(thread) {
  try {
    const label = GmailApp.getUserLabelByName("Completed") || GmailApp.createLabel("Completed");
    thread.addLabel(label);
    
    // Remove Human-Negotiation label if present
    const humanLabel = GmailApp.getUserLabelByName("Human-Negotiation");
    if(humanLabel) {
      try { thread.removeLabel(humanLabel); } catch(e) {}
    }
  } catch(e) {
    console.error("Failed to add Completed label:", e);
  }
}

/**
 * GMAIL → APP SYNC
 * Scans for threads marked "Completed" in Gmail and syncs them to sheets
 * This allows recruiters to mark complete directly in Gmail
 * Call this from runAutoNegotiator or set up as a separate trigger
 */
function syncCompletedFromGmail() {
  const url = getStoredSheetUrl();
  if(!url) return { synced: 0, cleaned: 0 };

  const ss = SpreadsheetApp.openByUrl(url);
  ensureSheetsExist(ss);

  const stateSheet = ss.getSheetByName('Negotiation_State');
  const taskSheet = ss.getSheetByName('Negotiation_Tasks');
  const compSheet = ss.getSheetByName('Negotiation_Completed');

  if(!stateSheet || !taskSheet || !compSheet) return { synced: 0, cleaned: 0 };
  
  let syncedCount = 0;
  let cleanedCount = 0;
  
  // Get all job labels to search
  const configSheet = ss.getSheetByName('Negotiation_Config');
  const configs = configSheet.getDataRange().getValues();
  
  for(let i=1; i<configs.length; i++) {
    const jobId = configs[i][0];
    if(!jobId) continue;
    
    // Search for threads with both Job label AND Completed label
    const query = `label:Job-${jobId} label:Completed`;
    let threads = [];
    
    try {
      threads = GmailApp.search(query, 0, 100);
    } catch(e) {
      continue;
    }
    
    threads.forEach(thread => {
      const msgs = thread.getMessages();
      if(msgs.length === 0) return;
      
      // Extract candidate email from thread
      let candidateEmail = '';
      let candidateName = '';
      
      // Find the first non-me sender
      const myEmail = Session.getActiveUser().getEmail().toLowerCase();
      for(let m of msgs) {
        const from = m.getFrom();
        if(from.toLowerCase().indexOf(myEmail) === -1) {
          const emailMatch = from.match(/<([^>]+)>/);
          candidateEmail = emailMatch ? emailMatch[1] : from.replace(/.*<|>.*/g, '');
          // Try to get name from "Name <email>" format
          const nameMatch = from.match(/^([^<]+)</);
          candidateName = nameMatch ? nameMatch[1].trim() : '';
          break;
        }
      }
      
      if(!candidateEmail) return;
      candidateEmail = candidateEmail.toLowerCase().trim();
      
      // Check if this email exists in State sheet (not yet synced)
      const stateData = stateSheet.getDataRange().getValues();
      for(let r=stateData.length-1; r>=1; r--) {
        if(String(stateData[r][0]).toLowerCase() === candidateEmail && String(stateData[r][1]) === String(jobId)) {
          // Found in state sheet - move to completed
          const devId = stateData[r][6] || 'N/A';
          const name = stateData[r][7] || candidateName || 'Unknown';
          const lastStatus = stateData[r][4] || 'Completed via Gmail';
          const aiNotes = stateData[r][8] || '';
          
          // Add to completed sheet
          compSheet.appendRow([
            new Date(),
            jobId,
            candidateEmail,
            name,
            'Completed (Gmail Sync)',
            aiNotes || 'Marked complete directly in Gmail',
            devId
          ]);

          // Update job-specific details sheet
          try {
            updateJobCandidateStatus(ss, jobId, candidateEmail, 'Completed (Gmail Sync)', null);
          } catch(detailsErr) {
            console.error("Failed to update job details sheet:", detailsErr);
          }

          // Remove from state sheet
          stateSheet.deleteRow(r+1);
          syncedCount++;
          break;
        }
      }

      // Also check Task sheet (accepted offers)
      const taskData = taskSheet.getDataRange().getValues();
      for(let r=taskData.length-1; r>=1; r--) {
        if(String(taskData[r][3]).toLowerCase() === candidateEmail && String(taskData[r][1]) === String(jobId)) {
          if(taskData[r][5] === 'Archived') continue; // Already archived

          const devId = taskData[r][6] || 'N/A';
          const name = taskData[r][2] || candidateName || 'Unknown';
          const agreedRate = taskData[r][4] || 'N/A';

          // Add to completed sheet
          compSheet.appendRow([
            new Date(),
            jobId,
            candidateEmail,
            name,
            `Accepted at $${agreedRate}/hr (Gmail Sync)`,
            'Marked complete directly in Gmail',
            devId
          ]);

          // Update job-specific details sheet
          try {
            updateJobCandidateStatus(ss, jobId, candidateEmail, 'Accepted (Gmail Sync)', `$${agreedRate}/hr`);
          } catch(detailsErr) {
            console.error("Failed to update job details sheet:", detailsErr);
          }

          // Remove from task sheet
          taskSheet.deleteRow(r+1);
          syncedCount++;
          break;
        }
      }

      // CLEANUP: Remove Human-Negotiation label if thread has Completed label
      const humanLabel = GmailApp.getUserLabelByName("Human-Negotiation");
      if(humanLabel) {
        const threadLabels = thread.getLabels().map(l => l.getName());
        if(threadLabels.includes("Human-Negotiation")) {
          try {
            thread.removeLabel(humanLabel);
            cleanedCount++;
          } catch(e) {}
        }
      }
    });
  }
  
  return { synced: syncedCount, cleaned: cleanedCount };
}

/**
 * Cleanup function: Remove Human-Negotiation label from any thread that has Completed label
 * Can be run manually or as part of the auto-negotiator
 */
function cleanupConflictingLabels() {
  let cleaned = 0;
  
  try {
    // Find threads with both labels
    const query = 'label:Completed label:Human-Negotiation';
    const threads = GmailApp.search(query, 0, 100);
    
    const humanLabel = GmailApp.getUserLabelByName("Human-Negotiation");
    if(!humanLabel) return { cleaned: 0 };
    
    threads.forEach(thread => {
      try {
        thread.removeLabel(humanLabel);
        cleaned++;
      } catch(e) {}
    });
  } catch(e) {
    console.error("Cleanup error:", e);
  }
  
  return { cleaned: cleaned };
}

function callAI(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) return "ACTION: ESCALATE (API Key missing)";

  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-4o",
    messages: [
      { role: "system", content: "You are a helpful recruitment negotiation assistant. Be concise and professional. Always check FAQs before answering candidate questions." },
      { role: "user", content: prompt }
    ],
    temperature: 0.7,
    max_tokens: 400
  };
  
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { "Authorization": `Bearer ${apiKey}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if(json.error) {
      console.error("OpenAI Error:", json.error);
      return "ACTION: ESCALATE (AI Error)";
    }
    
    return json.choices[0].message.content.trim();
  } catch (e) {
    console.error("AI call failed:", e);
    return "ACTION: ESCALATE (AI Error)";
  }
}

// --- UTILITY FUNCTIONS ---

function getTaskStats() {
  const url = getStoredSheetUrl();
  if(!url) return { total: 0, active: 0, human: 0, accepted: 0 };
  
  const ss = SpreadsheetApp.openByUrl(url);
  ensureSheetsExist(ss);
  
  const stateSheet = ss.getSheetByName('Negotiation_State');
  const taskSheet = ss.getSheetByName('Negotiation_Tasks');
  
  if(!stateSheet || !taskSheet) return { total: 0, active: 0, human: 0, accepted: 0 };
  
  const stateData = stateSheet.getDataRange().getValues();
  const taskData = taskSheet.getDataRange().getValues();
  
  let active = 0, human = 0;
  for(let i=1; i<stateData.length; i++) {
    if(stateData[i][0]) {
      if(stateData[i][4] === 'Human-Negotiation') human++;
      else active++;
    }
  }
  
  let accepted = 0;
  for(let i=1; i<taskData.length; i++) {
    if(taskData[i][3] && taskData[i][5] !== 'Archived') accepted++;
  }
  
  return {
    total: active + human + accepted,
    active: active,
    human: human,
    accepted: accepted
  };
}

// Get Gmail thread URL for viewing
function getThreadUrl(threadId) {
  if(!threadId) return null;
  return `https://mail.google.com/mail/u/0/#inbox/${threadId}`;
}

// ======================================================
// ===       AUTOMATED FOLLOW-UP EMAIL SYSTEM         ===
// ======================================================

/**
 * Follow-up timing configuration
 * FOLLOW_UP_1_HOURS: Hours after initial email to send first follow-up (12 hours)
 * FOLLOW_UP_2_HOURS: Hours after initial email to send second follow-up (28 hours)
 */
const FOLLOW_UP_CONFIG = {
  FOLLOW_UP_1_HOURS: 12,  // First follow-up after 12 hours
  FOLLOW_UP_2_HOURS: 28,  // Second follow-up after 28 hours (third total reachout)
  UNRESPONSIVE_HOURS: 76  // Mark as unresponsive 48 hours after 2nd follow-up (28 + 48 = 76 hours)
};

// Gmail labels for follow-up tracking
const FOLLOW_UP_LABELS = {
  AWAITING_RESPONSE: 'Awaiting-Response',
  FOLLOW_UP_1_SENT: 'Follow-Up-1-Sent',
  FOLLOW_UP_2_SENT: 'Follow-Up-2-Sent',
  UNRESPONSIVE: 'Unresponsive'
};

/**
 * Add a sent email to the follow-up queue
 * Called automatically when emails are sent via sendBulkEmails
 */
function addToFollowUpQueue(email, jobId, threadId, name, devId) {
  const url = getStoredSheetUrl();
  if(!url) return { success: false };

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    let sheet = ss.getSheetByName('Follow_Up_Queue');

    if(!sheet) {
      sheet = ss.insertSheet('Follow_Up_Queue');
      sheet.appendRow(['Email', 'Job ID', 'Thread ID', 'Name', 'Dev ID', 'Initial Send Time', 'Follow Up 1 Sent', 'Follow Up 2 Sent', 'Status', 'Last Updated', 'Manual Override']);
    }

    // Check if already in queue
    const data = sheet.getDataRange().getValues();
    const cleanEmail = String(email).toLowerCase().trim();

    for(let i = 1; i < data.length; i++) {
      if(String(data[i][0]).toLowerCase().trim() === cleanEmail && String(data[i][1]) === String(jobId)) {
        console.log(`Email ${email} already in follow-up queue for Job ${jobId}`);
        return { success: true, message: "Already in queue" };
      }
    }

    // Add new entry
    sheet.appendRow([
      email,
      jobId,
      threadId,
      name,
      devId || 'N/A',
      new Date(),        // Initial Send Time
      false,             // Follow Up 1 Sent
      false,             // Follow Up 2 Sent
      'Pending',         // Status
      new Date()         // Last Updated
    ]);

    // Add "Awaiting-Response" Gmail label to the thread
    if(threadId) {
      try {
        const thread = GmailApp.getThreadById(threadId);
        if(thread) {
          const label = GmailApp.getUserLabelByName(FOLLOW_UP_LABELS.AWAITING_RESPONSE) ||
                        GmailApp.createLabel(FOLLOW_UP_LABELS.AWAITING_RESPONSE);
          thread.addLabel(label);
        }
      } catch(labelError) {
        console.error("Error adding Awaiting-Response label:", labelError);
      }
    }

    return { success: true, message: "Added to queue" };
  } catch(e) {
    console.error("Error adding to follow-up queue:", e);
    return { success: false, error: e.message };
  }
}

/**
 * Check and process follow-up emails
 * This should be called periodically (e.g., every hour via time-based trigger)
 * Sends follow-up emails if no response received after 12 hours (first) and 28 hours (second)
 * Marks as unresponsive after 48 hours past 2nd follow-up (76 hours total)
 */
function processFollowUpQueue() {
  const url = getStoredSheetUrl();
  if(!url) return { processed: 0, followUp1Sent: 0, followUp2Sent: 0, unresponsiveMarked: 0, log: [] };

  const log = [];
  let followUp1Sent = 0;
  let followUp2Sent = 0;
  let unresponsiveMarked = 0;

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    ensureSheetsExist(ss);
    const sheet = ss.getSheetByName('Follow_Up_Queue');

    if(!sheet) {
      log.push({ type: 'warning', message: 'Follow_Up_Queue sheet not found' });
      return { processed: 0, followUp1Sent: 0, followUp2Sent: 0, unresponsiveMarked: 0, log: log };
    }

    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) {
      log.push({ type: 'info', message: 'No items in follow-up queue' });
      return { processed: 0, followUp1Sent: 0, followUp2Sent: 0, unresponsiveMarked: 0, log: log };
    }

    // Get negotiation state data to check if candidates are already in active negotiations
    const stateSheet = ss.getSheetByName('Negotiation_State');
    const stateData = stateSheet ? stateSheet.getDataRange().getValues() : [];
    const activeNegotiations = new Set();
    for(let j = 1; j < stateData.length; j++) {
      const stateEmail = String(stateData[j][0]).toLowerCase().trim();
      const stateJobId = String(stateData[j][1]);
      if(stateEmail && stateJobId) {
        activeNegotiations.add(`${stateEmail}|${stateJobId}`);
      }
    }

    // Also check Negotiation_Completed and Negotiation_Tasks for devs who already responded/completed
    const completedSheet = ss.getSheetByName('Negotiation_Completed');
    const completedData = completedSheet ? completedSheet.getDataRange().getValues() : [];
    const completedNegotiations = new Set();
    for(let j = 1; j < completedData.length; j++) {
      const compEmail = String(completedData[j][2]).toLowerCase().trim(); // Email is column 3
      const compJobId = String(completedData[j][1]);
      if(compEmail && compJobId) {
        completedNegotiations.add(`${compEmail}|${compJobId}`);
      }
    }

    const tasksSheet = ss.getSheetByName('Negotiation_Tasks');
    const tasksData = tasksSheet ? tasksSheet.getDataRange().getValues() : [];
    const acceptedOffers = new Set();
    for(let j = 1; j < tasksData.length; j++) {
      const taskEmail = String(tasksData[j][3]).toLowerCase().trim(); // Email is column 4
      const taskJobId = String(tasksData[j][1]);
      if(taskEmail && taskJobId) {
        acceptedOffers.add(`${taskEmail}|${taskJobId}`);
      }
    }

    const now = new Date();
    let processed = 0;
    const myEmail = Session.getActiveUser().getEmail().toLowerCase();
    const rowsToDelete = []; // Track rows to delete (for moving to unresponsive)

    for(let i = 1; i < data.length; i++) {
      const email = String(data[i][0]).toLowerCase().trim();
      const jobId = String(data[i][1]);
      const threadId = data[i][2];
      const name = data[i][3];
      const devId = data[i][4];
      const initialSendTime = new Date(data[i][5]);
      const followUp1Done = data[i][6] === true || data[i][6] === 'TRUE';
      const followUp2Done = data[i][7] === true || data[i][7] === 'TRUE';
      const status = data[i][8];
      const manualOverride = data[i][10] === true || data[i][10] === 'TRUE'; // Column 11 (index 10)

      // Skip already processed items
      if(status === 'Responded' || status === 'Unresponsive') {
        continue;
      }

      // Check if candidate is already in active negotiation, completed, or has accepted offer
      // BUT skip this auto-marking if Manual Override is set (user manually reset this entry)
      const negotiationKey = `${email}|${jobId}`;
      if(!manualOverride && (activeNegotiations.has(negotiationKey) || completedNegotiations.has(negotiationKey) || acceptedOffers.has(negotiationKey))) {
        // Mark as responded since they're in negotiation/completed/accepted
        sheet.getRange(i + 1, 9).setValue('Responded');
        sheet.getRange(i + 1, 10).setValue(new Date());
        updateFollowUpLabels(threadId, 'responded');
        const reason = activeNegotiations.has(negotiationKey) ? 'in active negotiation' :
                       completedNegotiations.has(negotiationKey) ? 'in completed negotiations' : 'has accepted offer';
        log.push({ type: 'success', message: `${email} ${reason} - marked as responded` });
        processed++;
        continue;
      }

      // IMPROVED: Check Gmail thread for candidate response by verifying sender matches candidate email
      // Skip auto-marking if Manual Override is set (user manually reset this entry)
      if(threadId && !manualOverride) {
        try {
          const thread = GmailApp.getThreadById(threadId);
          if(thread) {
            const messages = thread.getMessages();
            let candidateHasResponded = false;

            // Check ALL messages to see if candidate has replied at any point
            // We need to verify the message is FROM the candidate's email, not just "not from us"
            for(let m = 1; m < messages.length; m++) {
              const msg = messages[m];
              const sender = msg.getFrom().toLowerCase();

              // Extract just the email address from sender (handles "Name <email@domain.com>" format)
              const senderEmailMatch = sender.match(/<([^>]+)>/) || [null, sender.replace(/.*<|>.*/g, '').trim()];
              const senderEmail = senderEmailMatch[1] || sender.trim();

              // Check if this message is FROM the candidate (not just "not from us")
              // This prevents false positives from automated replies, other system messages, etc.
              if(senderEmail.includes(email) || email.includes(senderEmail)) {
                candidateHasResponded = true;
                log.push({ type: 'info', message: `${email} found response from sender: ${senderEmail}` });
                break;
              }
            }

            if(candidateHasResponded) {
              // Candidate has responded at some point - mark as responded
              sheet.getRange(i + 1, 9).setValue('Responded');
              sheet.getRange(i + 1, 10).setValue(new Date());
              updateFollowUpLabels(threadId, 'responded');
              log.push({ type: 'success', message: `${email} has responded - marked in queue` });
              processed++;
              continue;
            }
          }
        } catch(threadError) {
          console.error(`Error checking thread for ${email}:`, threadError);
        }
      }

      // Calculate hours since initial send
      const hoursSinceSend = (now - initialSendTime) / (1000 * 60 * 60);

      // Check if should be marked as unresponsive (76 hours = 28 + 48 hours after 2nd follow-up)
      // Skip all status changes if Manual Override is set - user wants to keep this entry as-is
      if(followUp1Done && followUp2Done && hoursSinceSend >= FOLLOW_UP_CONFIG.UNRESPONSIVE_HOURS) {
        // If Manual Override is set, skip all automatic status changes for this entry
        if(manualOverride) {
          log.push({ type: 'info', message: `${email} has Manual Override set - skipping automatic status change` });
          continue;
        }

        // SAFETY CHECK: Triple verify dev hasn't responded before marking unresponsive
        // This is a final safeguard to prevent marking responded/negotiating devs as unresponsive
        if(activeNegotiations.has(negotiationKey) || completedNegotiations.has(negotiationKey) || acceptedOffers.has(negotiationKey)) {
          sheet.getRange(i + 1, 9).setValue('Responded');
          sheet.getRange(i + 1, 10).setValue(new Date());
          updateFollowUpLabels(threadId, 'responded');
          log.push({ type: 'success', message: `${email} found in negotiations (safety check) - marked as responded` });
          processed++;
          continue;
        }

        // Move to Unresponsive_Devs sheet
        const moveResult = moveToUnresponsive(ss, email, jobId, name, devId, threadId, initialSendTime, data[i]);
        if(moveResult.success) {
          sheet.getRange(i + 1, 9).setValue('Unresponsive');
          sheet.getRange(i + 1, 10).setValue(new Date());
          updateFollowUpLabels(threadId, 'unresponsive');
          unresponsiveMarked++;
          log.push({ type: 'warning', message: `${email} marked unresponsive (${hoursSinceSend.toFixed(1)}hrs since initial)` });
        }
        processed++;
        continue;
      }

      // Check if first follow-up is due (12 hours)
      if(!followUp1Done && hoursSinceSend >= FOLLOW_UP_CONFIG.FOLLOW_UP_1_HOURS) {
        const result = sendFollowUpEmail(email, jobId, threadId, name, 1);
        if(result.success) {
          sheet.getRange(i + 1, 7).setValue(true); // Mark Follow Up 1 Sent
          sheet.getRange(i + 1, 10).setValue(new Date());
          updateFollowUpLabels(threadId, 'followup1');
          followUp1Sent++;
          log.push({ type: 'success', message: `Sent 1st follow-up to ${email} (${hoursSinceSend.toFixed(1)}hrs)` });
        } else {
          log.push({ type: 'error', message: `Failed 1st follow-up to ${email}: ${result.error}` });
        }
        processed++;
      }
      // Check if second follow-up is due (28 hours)
      else if(followUp1Done && !followUp2Done && hoursSinceSend >= FOLLOW_UP_CONFIG.FOLLOW_UP_2_HOURS) {
        const result = sendFollowUpEmail(email, jobId, threadId, name, 2);
        if(result.success) {
          sheet.getRange(i + 1, 8).setValue(true); // Mark Follow Up 2 Sent
          sheet.getRange(i + 1, 10).setValue(new Date());
          updateFollowUpLabels(threadId, 'followup2');
          followUp2Sent++;
          log.push({ type: 'success', message: `Sent 2nd follow-up to ${email} (${hoursSinceSend.toFixed(1)}hrs)` });
        } else {
          log.push({ type: 'error', message: `Failed 2nd follow-up to ${email}: ${result.error}` });
        }
        processed++;
      }
    }

    SpreadsheetApp.flush();

    log.push({ type: 'info', message: `Processed ${processed} items. 1st: ${followUp1Sent}, 2nd: ${followUp2Sent}, Unresponsive: ${unresponsiveMarked}` });

    return {
      processed: processed,
      followUp1Sent: followUp1Sent,
      followUp2Sent: followUp2Sent,
      unresponsiveMarked: unresponsiveMarked,
      log: log
    };

  } catch(e) {
    console.error("Error processing follow-up queue:", e);
    return { processed: 0, followUp1Sent: 0, followUp2Sent: 0, unresponsiveMarked: 0, log: [{ type: 'error', message: e.message }] };
  }
}

/**
 * Update Gmail labels for follow-up status changes
 * @param {string} threadId - Gmail thread ID
 * @param {string} newStatus - 'followup1', 'followup2', 'responded', or 'unresponsive'
 */
function updateFollowUpLabels(threadId, newStatus) {
  if(!threadId) return;

  try {
    const thread = GmailApp.getThreadById(threadId);
    if(!thread) return;

    // Get or create all follow-up labels
    const awaitingLabel = GmailApp.getUserLabelByName(FOLLOW_UP_LABELS.AWAITING_RESPONSE) ||
                          GmailApp.createLabel(FOLLOW_UP_LABELS.AWAITING_RESPONSE);
    const followUp1Label = GmailApp.getUserLabelByName(FOLLOW_UP_LABELS.FOLLOW_UP_1_SENT) ||
                           GmailApp.createLabel(FOLLOW_UP_LABELS.FOLLOW_UP_1_SENT);
    const followUp2Label = GmailApp.getUserLabelByName(FOLLOW_UP_LABELS.FOLLOW_UP_2_SENT) ||
                           GmailApp.createLabel(FOLLOW_UP_LABELS.FOLLOW_UP_2_SENT);
    const unresponsiveLabel = GmailApp.getUserLabelByName(FOLLOW_UP_LABELS.UNRESPONSIVE) ||
                              GmailApp.createLabel(FOLLOW_UP_LABELS.UNRESPONSIVE);

    // Remove all follow-up labels first
    try { thread.removeLabel(awaitingLabel); } catch(e) {}
    try { thread.removeLabel(followUp1Label); } catch(e) {}
    try { thread.removeLabel(followUp2Label); } catch(e) {}
    try { thread.removeLabel(unresponsiveLabel); } catch(e) {}

    // Add appropriate label based on new status
    switch(newStatus) {
      case 'pending':
        thread.addLabel(awaitingLabel);
        break;
      case 'followup1':
        thread.addLabel(followUp1Label);
        break;
      case 'followup2':
        thread.addLabel(followUp2Label);
        break;
      case 'unresponsive':
        thread.addLabel(unresponsiveLabel);
        break;
      case 'responded':
        // No label needed for responded - all labels removed
        break;
    }
  } catch(e) {
    console.error("Error updating follow-up labels:", e);
  }
}

/**
 * Move a candidate to the Unresponsive_Devs sheet
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} email - Candidate email
 * @param {string} jobId - Job ID
 * @param {string} name - Candidate name
 * @param {string} devId - Developer ID
 * @param {string} threadId - Gmail thread ID
 * @param {Date} initialSendTime - Initial email send time
 * @param {Array} rowData - Original row data from Follow_Up_Queue
 * @returns {Object} Result with success status
 */
function moveToUnresponsive(ss, email, jobId, name, devId, threadId, initialSendTime, rowData) {
  try {
    let unresponsiveSheet = ss.getSheetByName('Unresponsive_Devs');

    if(!unresponsiveSheet) {
      unresponsiveSheet = ss.insertSheet('Unresponsive_Devs');
      unresponsiveSheet.appendRow(['Email', 'Job ID', 'Name', 'Dev ID', 'Thread ID', 'Initial Send Time', 'Follow Up 1 Time', 'Follow Up 2 Time', 'Marked Unresponsive', 'Days Since Initial']);
    }

    // Calculate days since initial send
    const now = new Date();
    const daysSinceInitial = Math.round((now - initialSendTime) / (1000 * 60 * 60 * 24) * 10) / 10;

    // Get follow-up timestamps from original data if available
    const followUp1Time = rowData[6] === true || rowData[6] === 'TRUE' ? (rowData[9] || 'N/A') : 'N/A';
    const followUp2Time = rowData[7] === true || rowData[7] === 'TRUE' ? (rowData[9] || 'N/A') : 'N/A';

    // Add to Unresponsive_Devs sheet
    unresponsiveSheet.appendRow([
      email,
      jobId,
      name,
      devId || 'N/A',
      threadId || 'N/A',
      initialSendTime,
      followUp1Time,
      followUp2Time,
      now,
      daysSinceInitial
    ]);

    return { success: true };
  } catch(e) {
    console.error("Error moving to unresponsive:", e);
    return { success: false, error: e.message };
  }
}

/**
 * Send a follow-up email using AI to generate contextual content
 * @param {string} email - Recipient email
 * @param {string} jobId - Job ID
 * @param {string} threadId - Gmail thread ID to reply to
 * @param {string} name - Recipient name
 * @param {number} followUpNumber - 1 for first follow-up, 2 for second
 */
function sendFollowUpEmail(email, jobId, threadId, name, followUpNumber) {
  try {
    // Get job config for context
    const jobConfig = getNegotiationConfig(jobId);
    const jobDescription = jobConfig ? jobConfig.jobDescription : '';
    const firstName = name ? name.split(' ')[0] : 'there';

    // Generate follow-up email content using AI
    const prompt = `
You are a recruiter at Turing sending a follow-up email. This is follow-up #${followUpNumber} to a candidate who hasn't responded.

CONTEXT:
- Candidate Name: ${name}
- Job ID: ${jobId}
${jobDescription ? `- Job Description: ${jobDescription.substring(0, 500)}...` : ''}
- This is follow-up ${followUpNumber} of 2

FOLLOW-UP GUIDELINES:
${followUpNumber === 1 ? `
- This is the FIRST follow-up (sent 12 hours after initial email)
- Be friendly and casual
- Briefly remind them of the opportunity
- Ask if they have any questions
- Keep it short (3-4 sentences max)
` : `
- This is the SECOND and FINAL follow-up (sent 28 hours after initial email)
- Express that you want to make sure they saw your previous messages
- Mention this is a time-sensitive opportunity
- Keep it professional but slightly more urgent
- This is the last outreach, so be clear about that
- Keep it short (3-4 sentences max)
`}

Write ONLY the email body (no subject line). Start with "Hi ${firstName}," and end with appropriate sign-off.
`;

    const emailBody = callAI(prompt);

    // Get the thread and send reply
    if(threadId) {
      const thread = GmailApp.getThreadById(threadId);
      if(thread) {
        // Use the custom sender reply function
        sendReplyWithSenderName(thread, emailBody, 'Turing Recruitment Team');

        // Log the follow-up
        const url = getStoredSheetUrl();
        if(url) {
          const ss = SpreadsheetApp.openByUrl(url);
          const logSheet = ss.getSheetByName('Email_Logs');
          if(logSheet) {
            logSheet.appendRow([new Date(), jobId, email, name, threadId, `Follow-up ${followUpNumber}`]);
          }
        }

        return { success: true };
      }
    }

    // If no thread ID, send new email (fallback)
    const messages = GmailApp.search(`to:${email}`);
    if(messages && messages.length > 0) {
      const thread = messages[0];
      thread.reply(emailBody);
      return { success: true };
    }

    return { success: false, error: "Could not find thread to reply to" };

  } catch(e) {
    console.error(`Error sending follow-up email to ${email}:`, e);
    return { success: false, error: e.message };
  }
}

/**
 * Manual trigger to run follow-up processing
 * Can be called from the UI or set up as a time-based trigger
 */
function runFollowUpProcessor() {
  console.log("Starting follow-up processor...");
  const result = processFollowUpQueue();
  console.log(`Follow-up processing complete. Results:`, result);
  return result;
}

/**
 * Reset follow-up status for a specific email or all incorrectly marked "Responded" entries
 * This function re-verifies each entry by checking if the candidate actually responded
 * @param {string} emailToReset - Optional: specific email to reset. If not provided, checks all.
 * @returns {Object} { success: boolean, reset: number, log: Array }
 */
function resetFollowUpStatus(emailToReset) {
  const url = getStoredSheetUrl();
  if(!url) return { success: false, reset: 0, log: ['No sheet URL configured'] };

  const log = [];
  let resetCount = 0;

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName('Follow_Up_Queue');

    if(!sheet) {
      return { success: false, reset: 0, log: ['Follow_Up_Queue sheet not found'] };
    }

    const data = sheet.getDataRange().getValues();
    const myEmail = Session.getActiveUser().getEmail().toLowerCase();

    for(let i = 1; i < data.length; i++) {
      const email = String(data[i][0]).toLowerCase().trim();
      const status = data[i][8];
      const threadId = data[i][2];

      // Skip if looking for specific email and this isn't it
      if(emailToReset && email !== emailToReset.toLowerCase().trim()) continue;

      // Only process "Responded" entries (to re-verify them)
      if(status !== 'Responded') continue;

      // Re-check Gmail thread to verify if candidate actually responded
      let actuallyResponded = false;
      if(threadId) {
        try {
          const thread = GmailApp.getThreadById(threadId);
          if(thread) {
            const messages = thread.getMessages();

            // Check if any message is actually FROM the candidate
            for(let m = 1; m < messages.length; m++) {
              const msg = messages[m];
              const sender = msg.getFrom().toLowerCase();

              // Extract sender email
              const senderEmailMatch = sender.match(/<([^>]+)>/) || [null, sender.replace(/.*<|>.*/g, '').trim()];
              const senderEmail = senderEmailMatch[1] || sender.trim();

              // Verify it's from the candidate
              if(senderEmail.includes(email) || email.includes(senderEmail)) {
                actuallyResponded = true;
                log.push(`${email}: Verified - candidate did respond (from: ${senderEmail})`);
                break;
              }
            }
          }
        } catch(threadError) {
          log.push(`${email}: Error checking thread - ${threadError.message}`);
        }
      }

      // If candidate did NOT actually respond, reset to Pending
      if(!actuallyResponded) {
        sheet.getRange(i + 1, 9).setValue('Pending'); // Reset Status
        sheet.getRange(i + 1, 10).setValue(new Date()); // Update Last Updated
        sheet.getRange(i + 1, 11).setValue(true); // Set Manual Override flag to prevent auto-reprocessing

        // Determine correct label based on follow-up state
        const followUp1Done = data[i][6] === true || data[i][6] === 'TRUE';
        const followUp2Done = data[i][7] === true || data[i][7] === 'TRUE';

        let labelStatus = 'pending';
        if(followUp2Done) {
          labelStatus = 'followup2';
        } else if(followUp1Done) {
          labelStatus = 'followup1';
        }

        // Update Gmail label based on follow-up state
        updateFollowUpLabels(threadId, labelStatus);
        resetCount++;
        log.push(`${email}: RESET to Pending with Manual Override - no actual response found in thread (label: ${labelStatus})`);
      }
    }

    return { success: true, reset: resetCount, log: log };

  } catch(e) {
    console.error("Error resetting follow-up status:", e);
    return { success: false, reset: resetCount, log: [...log, `Error: ${e.message}`] };
  }
}

/**
 * Get follow-up queue statistics
 */
function getFollowUpStats() {
  const url = getStoredSheetUrl();
  if(!url) return { pending: 0, followUp1Done: 0, followUp2Done: 0, responded: 0, unresponsive: 0, jobIds: [] };

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName('Follow_Up_Queue');

    if(!sheet) return { pending: 0, followUp1Done: 0, followUp2Done: 0, responded: 0, unresponsive: 0, jobIds: [] };

    const data = sheet.getDataRange().getValues();

    let pending = 0, followUp1Done = 0, followUp2Done = 0, responded = 0, unresponsive = 0;
    const jobIdSet = new Set();

    for(let i = 1; i < data.length; i++) {
      const f1Done = data[i][6] === true || data[i][6] === 'TRUE';
      const f2Done = data[i][7] === true || data[i][7] === 'TRUE';
      const status = data[i][8];
      const jobId = String(data[i][1]);

      if(jobId) jobIdSet.add(jobId);

      if(status === 'Responded') {
        responded++;
      } else if(status === 'Unresponsive') {
        unresponsive++;
      } else if(f2Done) {
        followUp2Done++;
      } else if(f1Done) {
        followUp1Done++;
      } else {
        pending++;
      }
    }

    return { pending, followUp1Done, followUp2Done, responded, unresponsive, jobIds: Array.from(jobIdSet).sort() };

  } catch(e) {
    console.error("Error getting follow-up stats:", e);
    return { pending: 0, followUp1Done: 0, followUp2Done: 0, responded: 0, unresponsive: 0, jobIds: [] };
  }
}

/**
 * Get follow-up queue data for the UI table with filters
 * @param {Object} filters - Optional filters { jobId: 'all'|jobId, status: 'all'|status }
 * @returns {Object} Queue data with items and jobIds for filter dropdowns
 */
function getFollowUpQueueData(filters) {
  const url = getStoredSheetUrl();
  if(!url) return { items: [], jobIds: [] };

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName('Follow_Up_Queue');

    if(!sheet) return { items: [], jobIds: [] };

    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return { items: [], jobIds: [] };

    const jobIdFilter = filters?.jobId || 'all';
    const statusFilter = filters?.status || 'all';

    const items = [];
    const jobIdSet = new Set();

    for(let i = 1; i < data.length; i++) {
      const email = data[i][0];
      const jobId = String(data[i][1]);
      const threadId = data[i][2];
      const name = data[i][3];
      const devId = data[i][4];
      const initialSendTime = data[i][5];
      const f1Done = data[i][6] === true || data[i][6] === 'TRUE';
      const f2Done = data[i][7] === true || data[i][7] === 'TRUE';
      const status = data[i][8] || 'Pending';
      const lastUpdated = data[i][9];

      // Collect job IDs for filter dropdown
      if(jobId) jobIdSet.add(jobId);

      // Determine display status
      let displayStatus = status;
      if(status !== 'Responded' && status !== 'Unresponsive') {
        if(f2Done) {
          displayStatus = '2nd Follow-Up Sent';
        } else if(f1Done) {
          displayStatus = '1st Follow-Up Sent';
        } else {
          displayStatus = 'Awaiting Response';
        }
      }

      // Apply filters
      if(jobIdFilter !== 'all' && jobId !== jobIdFilter) continue;
      if(statusFilter !== 'all') {
        if(statusFilter === 'pending' && displayStatus !== 'Awaiting Response') continue;
        if(statusFilter === 'followup1' && displayStatus !== '1st Follow-Up Sent') continue;
        if(statusFilter === 'followup2' && displayStatus !== '2nd Follow-Up Sent') continue;
        if(statusFilter === 'responded' && displayStatus !== 'Responded') continue;
        if(statusFilter === 'unresponsive' && displayStatus !== 'Unresponsive') continue;
      }

      items.push({
        email: email,
        jobId: jobId,
        name: name || 'Unknown',
        devId: devId || 'N/A',
        threadId: threadId || '',
        initialSendTime: initialSendTime ? new Date(initialSendTime).toLocaleString() : 'N/A',
        followUp1Sent: f1Done,
        followUp2Sent: f2Done,
        status: displayStatus,
        lastUpdated: lastUpdated ? new Date(lastUpdated).toLocaleString() : 'N/A'
      });
    }

    return {
      items: items,
      jobIds: Array.from(jobIdSet).sort()
    };

  } catch(e) {
    console.error("Error getting follow-up queue data:", e);
    return { items: [], jobIds: [] };
  }
}

// ==========================================
// DAILY REPORT SYSTEM
// ==========================================

/**
 * Generate daily report data for all Job IDs
 * Gathers stats from Email_Logs, Negotiation_Tasks, Follow_Up_Queue
 * @param {Date} reportDate - The date to generate report for (defaults to today)
 * @returns {Object} Report data with stats per Job ID
 */
function generateDailyReport(reportDate) {
  const url = getStoredSheetUrl();
  if (!url) return { success: false, error: "No sheet URL configured" };

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    ensureSheetsExist(ss);

    const targetDate = reportDate ? new Date(reportDate) : new Date();
    targetDate.setHours(0, 0, 0, 0);
    const nextDay = new Date(targetDate);
    nextDay.setDate(nextDay.getDate() + 1);

    const reportByJobId = {};

    // Helper to check if timestamp is within target date
    function isWithinDate(timestamp) {
      if (!timestamp) return false;
      const d = new Date(timestamp);
      return d >= targetDate && d < nextDay;
    }

    // 1. Get Email_Logs for initial outreach (data gathering emails)
    const emailLogsSheet = ss.getSheetByName('Email_Logs');
    if (emailLogsSheet && emailLogsSheet.getLastRow() > 1) {
      const emailData = emailLogsSheet.getDataRange().getValues();
      // Headers: Timestamp, Job ID, Email, Name, Thread ID, Type (optional)
      for (let i = 1; i < emailData.length; i++) {
        const timestamp = emailData[i][0];
        const jobId = String(emailData[i][1] || '').trim();
        const emailType = String(emailData[i][5] || 'Data Gathering').trim();

        if (jobId && isWithinDate(timestamp)) {
          if (!reportByJobId[jobId]) {
            reportByJobId[jobId] = {
              aiRepliesSucceeded: 0,
              humanNegotiationsSent: 0,
              dataGatheringEmails: 0,
              firstFollowUps: 0,
              secondFollowUps: 0,
              totalOutreach: 0
            };
          }

          if (emailType.toLowerCase().includes('follow-up 1') || emailType.toLowerCase().includes('follow up 1')) {
            reportByJobId[jobId].firstFollowUps++;
          } else if (emailType.toLowerCase().includes('follow-up 2') || emailType.toLowerCase().includes('follow up 2')) {
            reportByJobId[jobId].secondFollowUps++;
          } else {
            reportByJobId[jobId].dataGatheringEmails++;
          }
          reportByJobId[jobId].totalOutreach++;
        }
      }
    }

    // 2. Get Negotiation_Tasks for human negotiations and AI replies
    const tasksSheet = ss.getSheetByName('Negotiation_Tasks');
    if (tasksSheet && tasksSheet.getLastRow() > 1) {
      const taskData = tasksSheet.getDataRange().getValues();
      // Headers: Timestamp, Job ID, Name, Email, Agreed Rate, Status, Dev ID, Thread ID, Region
      for (let i = 1; i < taskData.length; i++) {
        const timestamp = taskData[i][0];
        const jobId = String(taskData[i][1] || '').trim();
        const status = String(taskData[i][5] || '').toLowerCase();

        if (jobId && isWithinDate(timestamp)) {
          if (!reportByJobId[jobId]) {
            reportByJobId[jobId] = {
              aiRepliesSucceeded: 0,
              humanNegotiationsSent: 0,
              dataGatheringEmails: 0,
              firstFollowUps: 0,
              secondFollowUps: 0,
              totalOutreach: 0
            };
          }

          // Check if it's human intervention or AI reply
          if (status.includes('human') || status.includes('manual')) {
            reportByJobId[jobId].humanNegotiationsSent++;
          } else if (status.includes('active') || status.includes('accepted') || status.includes('succeeded') || status.includes('counter offer') || status.includes('counter-offer') || status.includes('pending')) {
            reportByJobId[jobId].aiRepliesSucceeded++;
          }
        }
      }
    }

    // 3. Get Negotiation_Completed for successful AI negotiations
    const completedSheet = ss.getSheetByName('Negotiation_Completed');
    if (completedSheet && completedSheet.getLastRow() > 1) {
      const completedData = completedSheet.getDataRange().getValues();
      // Headers: Timestamp, Job ID, Email, Name, Final Status, Notes, Dev ID, Region
      for (let i = 1; i < completedData.length; i++) {
        const timestamp = completedData[i][0];
        const jobId = String(completedData[i][1] || '').trim();
        const finalStatus = String(completedData[i][4] || '').toLowerCase();

        if (jobId && isWithinDate(timestamp)) {
          if (!reportByJobId[jobId]) {
            reportByJobId[jobId] = {
              aiRepliesSucceeded: 0,
              humanNegotiationsSent: 0,
              dataGatheringEmails: 0,
              firstFollowUps: 0,
              secondFollowUps: 0,
              totalOutreach: 0
            };
          }

          if (finalStatus.includes('accepted') || finalStatus.includes('success')) {
            reportByJobId[jobId].aiRepliesSucceeded++;
          }
        }
      }
    }

    // 4. Get Follow_Up_Queue for follow-up counts (check Last Updated date)
    const followUpSheet = ss.getSheetByName('Follow_Up_Queue');
    if (followUpSheet && followUpSheet.getLastRow() > 1) {
      const followUpData = followUpSheet.getDataRange().getValues();
      // Headers: Email, Job ID, Thread ID, Name, Dev ID, Initial Send Time, Follow Up 1 Sent, Follow Up 2 Sent, Status, Last Updated
      for (let i = 1; i < followUpData.length; i++) {
        const jobId = String(followUpData[i][1] || '').trim();
        const lastUpdated = followUpData[i][9];
        const f1Sent = followUpData[i][6];
        const f2Sent = followUpData[i][7];

        if (jobId && isWithinDate(lastUpdated)) {
          if (!reportByJobId[jobId]) {
            reportByJobId[jobId] = {
              aiRepliesSucceeded: 0,
              humanNegotiationsSent: 0,
              dataGatheringEmails: 0,
              firstFollowUps: 0,
              secondFollowUps: 0,
              totalOutreach: 0
            };
          }

          // These may already be counted in Email_Logs, but let's add from queue for accuracy
          if (f1Sent === true || f1Sent === 'TRUE') {
            // Follow-up 1 was sent - already counted if logged
          }
          if (f2Sent === true || f2Sent === 'TRUE') {
            // Follow-up 2 was sent - already counted if logged
          }
        }
      }
    }

    // Calculate totals
    let totalAiReplies = 0, totalHuman = 0, totalDataGathering = 0, totalF1 = 0, totalF2 = 0, totalOutreach = 0;
    const jobIds = Object.keys(reportByJobId);

    jobIds.forEach(jobId => {
      const stats = reportByJobId[jobId];
      totalAiReplies += stats.aiRepliesSucceeded;
      totalHuman += stats.humanNegotiationsSent;
      totalDataGathering += stats.dataGatheringEmails;
      totalF1 += stats.firstFollowUps;
      totalF2 += stats.secondFollowUps;
      totalOutreach += stats.totalOutreach;
    });

    return {
      success: true,
      reportDate: targetDate.toISOString().split('T')[0],
      jobStats: reportByJobId,
      totals: {
        aiRepliesSucceeded: totalAiReplies,
        humanNegotiationsSent: totalHuman,
        dataGatheringEmails: totalDataGathering,
        firstFollowUps: totalF1,
        secondFollowUps: totalF2,
        totalOutreach: totalOutreach
      },
      jobCount: jobIds.length
    };

  } catch (e) {
    console.error("Error generating daily report:", e);
    return { success: false, error: e.message };
  }
}

/**
 * Save daily report to Daily_Reports sheet
 * @param {Object} reportData - Report data from generateDailyReport
 * @param {string} sentTo - Email address report was sent to
 */
function saveDailyReportToSheet(reportData, sentTo) {
  const url = getStoredSheetUrl();
  if (!url || !reportData || !reportData.success) return;

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName('Daily_Reports');
    if (!sheet) return;

    const now = new Date();
    const jobIds = Object.keys(reportData.jobStats);

    // Save a row for each Job ID
    jobIds.forEach(jobId => {
      const stats = reportData.jobStats[jobId];
      sheet.appendRow([
        reportData.reportDate,
        jobId,
        stats.aiRepliesSucceeded,
        stats.humanNegotiationsSent,
        stats.dataGatheringEmails,
        stats.firstFollowUps,
        stats.secondFollowUps,
        stats.totalOutreach,
        now,
        sentTo || ''
      ]);
    });

    // Also save totals row
    sheet.appendRow([
      reportData.reportDate,
      'TOTAL',
      reportData.totals.aiRepliesSucceeded,
      reportData.totals.humanNegotiationsSent,
      reportData.totals.dataGatheringEmails,
      reportData.totals.firstFollowUps,
      reportData.totals.secondFollowUps,
      reportData.totals.totalOutreach,
      now,
      sentTo || ''
    ]);

  } catch (e) {
    console.error("Error saving daily report to sheet:", e);
  }
}

/**
 * Send daily report email with stats in HTML table format
 * @param {string} recipientEmail - Email to send report to (defaults to current user)
 * @param {Date} reportDate - Date to generate report for (defaults to today)
 * @returns {Object} Result of the operation
 */
function sendDailyReportEmail(recipientEmail, reportDate) {
  try {
    const email = recipientEmail || Session.getActiveUser().getEmail();
    if (!email) {
      return { success: false, error: "No recipient email specified" };
    }

    const report = generateDailyReport(reportDate);
    if (!report.success) {
      return { success: false, error: report.error || "Failed to generate report" };
    }

    const jobIds = Object.keys(report.jobStats);
    if (jobIds.length === 0) {
      return { success: false, error: "No activity found for the specified date" };
    }

    // Build HTML email
    const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        .container { max-width: 800px; margin: 0 auto; padding: 20px; }
        h1 { color: #2563eb; border-bottom: 2px solid #2563eb; padding-bottom: 10px; }
        h2 { color: #1e40af; margin-top: 30px; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th { background: #2563eb; color: white; padding: 12px 8px; text-align: left; font-weight: 600; }
        td { padding: 10px 8px; border-bottom: 1px solid #e5e7eb; }
        tr:nth-child(even) { background: #f9fafb; }
        tr:hover { background: #f3f4f6; }
        .total-row { background: #dbeafe !important; font-weight: bold; }
        .total-row td { border-top: 2px solid #2563eb; }
        .stat-card { display: inline-block; background: #f3f4f6; padding: 15px 20px; margin: 5px; border-radius: 8px; text-align: center; }
        .stat-value { font-size: 24px; font-weight: bold; color: #2563eb; }
        .stat-label { font-size: 12px; color: #6b7280; text-transform: uppercase; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #e5e7eb; color: #6b7280; font-size: 12px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>Daily Activity Report</h1>
        <p><strong>Date:</strong> ${report.reportDate}</p>
        <p><strong>Generated:</strong> ${new Date().toLocaleString()}</p>

        <h2>Summary</h2>
        <div style="margin: 20px 0;">
          <div class="stat-card">
            <div class="stat-value">${report.totals.totalOutreach}</div>
            <div class="stat-label">Total Outreach</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">${report.totals.aiRepliesSucceeded}</div>
            <div class="stat-label">AI Replies</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">${report.totals.humanNegotiationsSent}</div>
            <div class="stat-label">Human Negotiations</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">${report.totals.dataGatheringEmails}</div>
            <div class="stat-label">Data Gathering</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">${report.totals.firstFollowUps}</div>
            <div class="stat-label">1st Follow-ups</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">${report.totals.secondFollowUps}</div>
            <div class="stat-label">2nd Follow-ups</div>
          </div>
        </div>

        <h2>Activity by Job ID</h2>
        <table>
          <thead>
            <tr>
              <th>Job ID</th>
              <th>AI Replies</th>
              <th>Human Negotiations</th>
              <th>Data Gathering</th>
              <th>1st Follow-ups</th>
              <th>2nd Follow-ups</th>
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            ${jobIds.map(jobId => {
              const s = report.jobStats[jobId];
              return `<tr>
                <td><strong>${jobId}</strong></td>
                <td>${s.aiRepliesSucceeded}</td>
                <td>${s.humanNegotiationsSent}</td>
                <td>${s.dataGatheringEmails}</td>
                <td>${s.firstFollowUps}</td>
                <td>${s.secondFollowUps}</td>
                <td><strong>${s.totalOutreach}</strong></td>
              </tr>`;
            }).join('')}
            <tr class="total-row">
              <td>TOTAL</td>
              <td>${report.totals.aiRepliesSucceeded}</td>
              <td>${report.totals.humanNegotiationsSent}</td>
              <td>${report.totals.dataGatheringEmails}</td>
              <td>${report.totals.firstFollowUps}</td>
              <td>${report.totals.secondFollowUps}</td>
              <td><strong>${report.totals.totalOutreach}</strong></td>
            </tr>
          </tbody>
        </table>

        <div class="footer">
          <p>This report was automatically generated by Turing AI Recruiter V11.</p>
          <p>Total Jobs with Activity: ${report.jobCount}</p>
        </div>
      </div>
    </body>
    </html>`;

    // Send the email
    GmailApp.sendEmail(email, `Daily Activity Report - ${report.reportDate}`, '', {
      htmlBody: html,
      name: 'Turing AI Recruiter'
    });

    // Save to sheet
    saveDailyReportToSheet(report, email);

    return {
      success: true,
      message: `Report sent to ${email}`,
      stats: report.totals,
      jobCount: report.jobCount
    };

  } catch (e) {
    console.error("Error sending daily report email:", e);
    return { success: false, error: e.message };
  }
}

/**
 * Trigger function to run daily report automatically
 * Set up a time-based trigger to run this function daily
 * Sends YESTERDAY's report (so 8 AM trigger sends previous day's full activity)
 * NOTE: If there's no activity for the day, NO email is sent (by design)
 */
function runDailyReportTrigger() {
  const userEmail = Session.getActiveUser().getEmail();
  if (userEmail) {
    // Get yesterday's date for the report
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const result = sendDailyReportEmail(userEmail, yesterday);

    // Log the result for debugging
    if (result.success) {
      console.log(`Daily report sent to ${userEmail} for ${yesterday.toDateString()}`);
    } else {
      console.log(`Daily report NOT sent: ${result.error}`);
    }
  }
}

/**
 * Get historical daily reports from the Daily_Reports sheet
 * @param {number} limit - Number of days to retrieve (default 7)
 * @returns {Array} Array of report summaries
 */
function getDailyReportHistory(limit) {
  const url = getStoredSheetUrl();
  if (!url) return [];

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName('Daily_Reports');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const data = sheet.getDataRange().getValues();
    const reportsByDate = {};

    // Group by date and aggregate TOTAL rows
    for (let i = 1; i < data.length; i++) {
      const date = data[i][0];
      const jobId = String(data[i][1]);

      if (jobId === 'TOTAL') {
        const dateKey = typeof date === 'object' ? date.toISOString().split('T')[0] : String(date);
        reportsByDate[dateKey] = {
          date: dateKey,
          aiReplies: data[i][2] || 0,
          humanNegotiations: data[i][3] || 0,
          dataGathering: data[i][4] || 0,
          firstFollowUps: data[i][5] || 0,
          secondFollowUps: data[i][6] || 0,
          totalOutreach: data[i][7] || 0,
          generatedAt: data[i][8],
          sentTo: data[i][9]
        };
      }
    }

    // Sort by date descending and limit
    const sortedDates = Object.keys(reportsByDate).sort().reverse();
    const maxResults = limit || 7;

    return sortedDates.slice(0, maxResults).map(date => reportsByDate[date]);

  } catch (e) {
    console.error("Error getting daily report history:", e);
    return [];
  }
}

// ==========================================
// TRIGGER MANAGEMENT SYSTEM
// ==========================================

/**
 * Configuration for all required triggers
 * Each trigger has: functionName, type (hourly/daily), description
 */
const REQUIRED_TRIGGERS = [
  {
    functionName: 'runAutoNegotiator',
    type: 'hourly',
    description: 'Processes Gmail negotiations automatically (sync, cleanup, AI replies)',
    everyHours: 1
  },
  {
    functionName: 'runFollowUpProcessor',
    type: 'hourly',
    description: 'Processes follow-up emails (1st at 12hrs, 2nd at 28hrs)',
    everyHours: 1
  },
  {
    functionName: 'runDailyReportTrigger',
    type: 'daily',
    description: 'Sends daily activity report email',
    atHour: 8 // 8 AM
  }
];

/**
 * Get status of all triggers - which exist and which are missing
 * @returns {Object} Status of each required trigger
 */
function getTriggerStatus() {
  try {
    const allTriggers = ScriptApp.getProjectTriggers();
    const status = {
      triggers: [],
      existingCount: 0,
      missingCount: 0,
      totalRequired: REQUIRED_TRIGGERS.length
    };

    REQUIRED_TRIGGERS.forEach(required => {
      const existingTrigger = allTriggers.find(t =>
        t.getHandlerFunction() === required.functionName
      );

      const triggerInfo = {
        functionName: required.functionName,
        type: required.type,
        description: required.description,
        exists: !!existingTrigger,
        triggerId: existingTrigger ? existingTrigger.getUniqueId() : null
      };

      if (existingTrigger) {
        triggerInfo.lastRun = null; // Can't get this from ScriptApp
        status.existingCount++;
      } else {
        status.missingCount++;
      }

      status.triggers.push(triggerInfo);
    });

    return { success: true, status };

  } catch (e) {
    console.error("Error getting trigger status:", e);
    return { success: false, error: e.message };
  }
}

/**
 * Create a single trigger by function name
 * @param {string} functionName - The function to create trigger for
 * @returns {Object} Result of trigger creation
 */
function createTrigger(functionName) {
  try {
    const config = REQUIRED_TRIGGERS.find(t => t.functionName === functionName);
    if (!config) {
      return { success: false, error: `Unknown trigger function: ${functionName}` };
    }

    // Check if trigger already exists
    const existingTriggers = ScriptApp.getProjectTriggers();
    const existing = existingTriggers.find(t => t.getHandlerFunction() === functionName);
    if (existing) {
      return { success: true, message: 'Trigger already exists', alreadyExisted: true };
    }

    // Create the appropriate trigger type
    let trigger;
    if (config.type === 'hourly') {
      trigger = ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyHours(config.everyHours || 1)
        .create();
    } else if (config.type === 'daily') {
      trigger = ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyDays(1)
        .atHour(config.atHour || 8)
        .create();
    }

    return {
      success: true,
      message: `Trigger created for ${functionName}`,
      triggerId: trigger.getUniqueId(),
      alreadyExisted: false
    };

  } catch (e) {
    console.error(`Error creating trigger for ${functionName}:`, e);
    return { success: false, error: e.message };
  }
}

/**
 * Create all missing triggers at once
 * @returns {Object} Results of trigger creation
 */
function createAllMissingTriggers() {
  try {
    const results = {
      created: [],
      alreadyExisted: [],
      failed: []
    };

    REQUIRED_TRIGGERS.forEach(config => {
      const result = createTrigger(config.functionName);
      if (result.success) {
        if (result.alreadyExisted) {
          results.alreadyExisted.push(config.functionName);
        } else {
          results.created.push(config.functionName);
        }
      } else {
        results.failed.push({ functionName: config.functionName, error: result.error });
      }
    });

    return {
      success: true,
      results,
      message: `Created ${results.created.length} triggers, ${results.alreadyExisted.length} already existed, ${results.failed.length} failed`
    };

  } catch (e) {
    console.error("Error creating triggers:", e);
    return { success: false, error: e.message };
  }
}

/**
 * Delete a specific trigger by function name
 * @param {string} functionName - The function whose trigger to delete
 * @returns {Object} Result of deletion
 */
function deleteTrigger(functionName) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const trigger = triggers.find(t => t.getHandlerFunction() === functionName);

    if (!trigger) {
      return { success: false, error: 'Trigger not found' };
    }

    ScriptApp.deleteTrigger(trigger);
    return { success: true, message: `Trigger deleted for ${functionName}` };

  } catch (e) {
    console.error(`Error deleting trigger for ${functionName}:`, e);
    return { success: false, error: e.message };
  }
}

/**
 * Delete all project triggers (use with caution)
 * @returns {Object} Result of deletion
 */
function deleteAllTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deleted = 0;

    triggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    });

    return { success: true, message: `Deleted ${deleted} triggers` };

  } catch (e) {
    console.error("Error deleting triggers:", e);
    return { success: false, error: e.message };
  }
}

/**
 * Get detailed information about all project triggers
 * @returns {Array} List of all triggers with details
 */
function getAllTriggersInfo() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    return triggers.map(t => ({
      id: t.getUniqueId(),
      functionName: t.getHandlerFunction(),
      eventType: t.getEventType().toString(),
      triggerSource: t.getTriggerSource().toString()
    }));
  } catch (e) {
    console.error("Error getting triggers info:", e);
    return [];
  }
}

// ==========================================
// FOLLOW-UP TABLE FUNCTIONS (UI Support)
// ==========================================

/**
 * Get follow-up table data for the UI with filters
 * Returns developers currently under follow-up with their details
 * @param {Object} filters - Optional filters { jobId: 'all'|jobId, status: 'all'|'Pending'|'Follow-Up-1-Sent'|'Follow-Up-2-Sent' }
 * @returns {Object} { data: [...], jobIds: [...] }
 */
function getFollowUpTableData(filters) {
  const url = getStoredSheetUrl();
  if (!url) return { data: [], jobIds: [] };

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName('Follow_Up_Queue');

    if (!sheet || sheet.getLastRow() <= 1) return { data: [], jobIds: [] };

    const data = sheet.getDataRange().getValues();
    const jobIdFilter = filters?.jobId || 'all';
    const statusFilter = filters?.status || 'all';

    const items = [];
    const jobIdSet = new Set();

    // Headers: Email, Job ID, Thread ID, Name, Dev ID, Initial Send Time, Follow Up 1 Sent, Follow Up 2 Sent, Status, Last Updated
    for (let i = 1; i < data.length; i++) {
      const email = data[i][0];
      const jobId = String(data[i][1] || '');
      const threadId = data[i][2];
      const name = data[i][3];
      const devId = data[i][4];
      const initialSendTime = data[i][5];
      const f1Done = data[i][6] === true || data[i][6] === 'TRUE';
      const f2Done = data[i][7] === true || data[i][7] === 'TRUE';
      const status = data[i][8] || '';
      const lastUpdated = data[i][9];

      // Skip responded or already marked unresponsive
      if (status === 'Responded' || status === 'Unresponsive') continue;

      // Collect job IDs for filter dropdown
      if (jobId) jobIdSet.add(jobId);

      // Determine follow-up status for display
      let followUpStatus = 'Pending';
      if (f2Done) {
        followUpStatus = 'Follow-Up-2-Sent';
      } else if (f1Done) {
        followUpStatus = 'Follow-Up-1-Sent';
      }

      // Apply job ID filter
      if (jobIdFilter !== 'all' && jobId !== jobIdFilter) continue;

      // Apply status filter
      if (statusFilter !== 'all' && followUpStatus !== statusFilter) continue;

      // Determine last reached out time
      let lastReachedOut = initialSendTime;
      if (lastUpdated && (f1Done || f2Done)) {
        lastReachedOut = lastUpdated;
      }

      items.push({
        email: email,
        jobId: jobId,
        name: name || 'Unknown',
        devId: devId || '',
        threadId: threadId || '',
        followUpStatus: followUpStatus,
        lastReachedOut: lastReachedOut ? new Date(lastReachedOut).toISOString() : null,
        initialSendTime: initialSendTime ? new Date(initialSendTime).toISOString() : null
      });
    }

    // Sort by last reached out (most recent first)
    items.sort((a, b) => {
      const timeA = a.lastReachedOut ? new Date(a.lastReachedOut) : new Date(0);
      const timeB = b.lastReachedOut ? new Date(b.lastReachedOut) : new Date(0);
      return timeB - timeA;
    });

    return {
      data: items,
      jobIds: Array.from(jobIdSet).sort()
    };

  } catch (e) {
    console.error("Error getting follow-up table data:", e);
    return { data: [], jobIds: [] };
  }
}

/**
 * Mark a developer as unresponsive and remove from follow-up queue
 * Moves the entry to Unresponsive_Devs sheet and removes from Follow_Up_Queue
 * @param {string} email - Developer email to mark as unresponsive
 * @returns {Object} { success: boolean, message: string }
 */
function markFollowUpAsUnresponsive(email) {
  if (!email) return { success: false, message: "Email is required" };

  const url = getStoredSheetUrl();
  if (!url) return { success: false, message: "No sheet URL configured" };

  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName('Follow_Up_Queue');

    if (!sheet) return { success: false, message: "Follow_Up_Queue sheet not found" };

    const data = sheet.getDataRange().getValues();

    // Find the row with this email
    let rowIndex = -1;
    let rowData = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        rowIndex = i + 1; // Sheet rows are 1-indexed
        rowData = data[i];
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "Email not found in follow-up queue" };
    }

    // Extract data from the row
    const jobId = rowData[1];
    const threadId = rowData[2];
    const name = rowData[3];
    const devId = rowData[4];
    const initialSendTime = rowData[5];

    // Move to Unresponsive_Devs sheet
    const moveResult = moveToUnresponsive(ss, email, jobId, name, devId, threadId, initialSendTime, rowData);

    if (!moveResult.success) {
      return { success: false, message: "Failed to move to unresponsive: " + (moveResult.error || "Unknown error") };
    }

    // Delete the row from Follow_Up_Queue
    sheet.deleteRow(rowIndex);

    // Invalidate cache
    invalidateSheetCache('Follow_Up_Queue');

    return { success: true, message: "Developer marked as unresponsive and removed from follow-up queue" };

  } catch (e) {
    console.error("Error marking as unresponsive:", e);
    return { success: false, message: "Error: " + e.message };
  }
}
