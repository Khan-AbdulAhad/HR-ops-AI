/**
 * TURING AI RECRUITER V9 - LIGHT VERSION
 * =================================================================
 * Changes:
 * 1. REMOVED Negotiation Sheets (Config, Tasks, State, FAQs, Completed).
 * 2. PREVENTED auto-creation of those sheets.
 * 3. DISABLED all negotiation logic in sendBulkEmails and runAutoNegotiator.
 * 4. KEPT Sourcing, Manual Logging, and Email Logs.
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

// UPDATED: Removed Negotiation sheets to prevent auto-creation
const SHEET_SCHEMAS = {
  'Email_Logs': ['Timestamp', 'Job ID', 'Email', 'Name', 'Thread ID', 'Type'],
  'Email_Templates': ['Template Name', 'Subject', 'Body', 'Job ID', 'Created Date'],
  'Manual_Sent_Logs': ['Timestamp', 'Job ID', 'Developer ID', 'Email', 'Name', 'Note', 'Marked By'],
  'Data_Fetch_Logs': ['Timestamp', 'Source', 'Context', 'Data Size (Bytes)', 'Duration', 'Details', 'User']
};

// --- SETUP ---

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Turing AI Recruiter V9')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getStoredSheetUrl() {
  return PropertiesService.getUserProperties().getProperty('LOG_SHEET_URL') || "";
}

function saveLogSheetUrl(url) {
  const cleanUrl = url ? url.trim() : "";
  if(!cleanUrl) throw new Error("Invalid URL");
  try {
    const ss = SpreadsheetApp.openByUrl(cleanUrl);
    const result = validateAndFixSheets(ss);
    PropertiesService.getUserProperties().setProperty('LOG_SHEET_URL', cleanUrl);
    return { success: true, message: result.message };
  } catch (e) { 
    throw new Error("Could not access Sheet: " + e.message); 
  }
}

function validateAndFixSheets(ss) {
  if(!ss) {
    const url = getStoredSheetUrl();
    if(!url) return { success: false, message: "No sheet URL configured" };
    ss = SpreadsheetApp.openByUrl(url);
  }
  
  let issues = [];
  let fixes = [];
  
  console.log("Starting sheet validation...");
  
  Object.keys(SHEET_SCHEMAS).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      fixes.push(`Created missing sheet: ${sheetName}`);
      console.log(`Created sheet: ${sheetName}`);
    }
    
    const expectedHeaders = SHEET_SCHEMAS[sheetName];
    const lastRow = sheet.getLastRow();
    
    if (lastRow === 0) {
      sheet.appendRow(expectedHeaders);
      fixes.push(`Added headers to: ${sheetName}`);
      console.log(`Added headers to: ${sheetName}`);
    } else {
      const actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
      let headersMismatch = false;
      
      for(let i = 0; i < expectedHeaders.length; i++) {
        if(actualHeaders[i] !== expectedHeaders[i]) {
          headersMismatch = true;
          break;
        }
      }
      
      if(headersMismatch) {
        issues.push(`Headers mismatch in: ${sheetName}`);
        sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
        fixes.push(`Fixed headers in: ${sheetName}`);
        console.log(`Fixed headers in: ${sheetName}`);
      }
    }
  });
  
  SpreadsheetApp.flush();
  
  const message = fixes.length > 0 
    ? `Fixed ${fixes.length} issue(s): ${fixes.join(', ')}` 
    : "All sheets validated successfully";
  
  console.log("Validation complete: " + message);
  
  return {
    success: true,
    issues: issues,
    fixes: fixes,
    message: message
  };
}

function manualValidateSheets() {
  const url = getStoredSheetUrl();
  if(!url) {
    console.error("No sheet URL configured");
    return { success: false, message: "Please configure your Google Sheets URL first" };
  }
  
  try {
    const ss = SpreadsheetApp.openByUrl(url);
    const result = validateAndFixSheets(ss);
    console.log("Manual validation result:", result);
    return result;
  } catch(e) {
    console.error("Error during manual validation:", e);
    return { success: false, message: "Error: " + e.message };
  }
}

function ensureSheetsExist(ss) {
  validateAndFixSheets(ss);
}

// --- LOGGING HELPER ---

function logDataConsumption(source, context, byteSize, durationMs, details) {
  const url = getStoredSheetUrl();
  if(!url) return;
  
  try {
    const ss = SpreadsheetApp.openByUrl(url);
    
    let userEmail = 'System/Automated';
    try {
        const email = Session.getActiveUser().getEmail();
        if (email) userEmail = email;
    } catch (e) {}

    let sheet = ss.getSheetByName('Data_Fetch_Logs');
    if (!sheet) {
        sheet = ss.insertSheet('Data_Fetch_Logs');
        sheet.appendRow(SHEET_SCHEMAS['Data_Fetch_Logs']);
    }
    const timeValue = durationMs / 86400000;
    sheet.appendRow([
        new Date(),
        source,
        context,
        byteSize,
        timeValue,
        details || '',
        userEmail
    ]);
  } catch (e) {
      console.error("Failed to log data consumption:", e);
  }
}

// --- NEGOTIATION CONFIGURATION (DISABLED) ---

function saveNegotiationConfig(jobId, config) {
  // DISABLED: Negotiation_Config sheet removed
  console.log("Negotiation features disabled. Config not saved.");
  return;
}

function getNegotiationConfig(jobId) {
  // DISABLED: Negotiation_Config sheet removed
  return null;
}

function getJobConfigList() {
  // DISABLED: Negotiation_Config sheet removed
  return [];
}

// --- FAQ HELPER (DISABLED) ---

function getFAQs() {
  // DISABLED: Negotiation_FAQs sheet removed
  return "Negotiation FAQ features disabled.";
}

// --- TASK LIST MANAGEMENT (DISABLED) ---

function getAllTasks() {
  // DISABLED: Negotiation_State and Tasks sheets removed
  return [];
}

function moveToCompleted(email, finalStatus) {
  // DISABLED: Negotiation_Completed sheet removed
  return { success: false, message: "Negotiation features are disabled in this version." };
}

// ==============================================================
// ===  GET DEVELOPERS - AGENCY/SUBCONTRACTOR SUPPORT ACTIVE  ===
// ==============================================================

function getDevelopers(jobId, selectedStages) {
  if (!jobId || !selectedStages || selectedStages.length === 0) throw new Error("Missing inputs");
  const cleanJobId = Number(jobId);

  const logs = getEmailLogs(cleanJobId) || new Map();
  const manualLogs = getManualSentLogs(cleanJobId) || new Map();

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

  const innerQuery = `
    WITH target_devs AS (
      ${unionQuery}
    ),
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

    if (typeof logDataConsumption === 'function') {
      try {
        logDataConsumption('BigQuery', `Job-${cleanJobId}`, dataSizeBytes, endTime - startTime, `Rows returned: ${rows.length}`);
      } catch (le) {
        console.warn("logDataConsumption call failed:", le);
      }
    }

    const devMap = new Map();
    (rows || []).forEach(row => {
      const devId = row.f[0] && row.f[0].v ? String(row.f[0].v) : null;
      if (!devId) return;

      const fullName = row.f[1] && row.f[1].v ? row.f[1].v : '';
      const email = row.f[2] && row.f[2].v ? row.f[2].v : '';
      const status = row.f[3] && row.f[3].v ? row.f[3].v : '';
      const candidateStatus = row.f[4] && row.f[4].v ? row.f[4].v : 'Independent';
      const agencyName = row.f[5] && row.f[5].v ? row.f[5].v : '';

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
  
  validateAndFixSheets(ss);
  
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
      if (String(data[i][1]) === String(jobId)) {
        const devId = String(data[i][2]);
        const note = data[i][5] || '';
        
        if (!logMap.has(devId)) {
          logMap.set(devId, { count: 0, note: note });
        }
        logMap.get(devId).count++;
        if (note) {
          logMap.get(devId).note = note;
        }
      }
    }
    return logMap;
  } catch(e) {
    console.error("Error reading Manual_Sent_Logs:", e);
    return new Map();
  }
}

function markAsManualSent(developerIds, jobId, note) {
  const url = getStoredSheetUrl();
  if (!url) return { success: false, error: "No config URL set" };
  
  try {
    const ss = SpreadsheetApp.openByUrl(url);
    validateAndFixSheets(ss);
    
    const sheet = ss.getSheetByName("Manual_Sent_Logs");
    if (!sheet) {
      return { success: false, error: "Manual_Sent_Logs sheet not found" };
    }
    
    const markedBy = Session.getActiveUser().getEmail() || 'Unknown';
    const timestamp = new Date();
    let marked = 0;
    
    developerIds.forEach(devId => {
      sheet.appendRow([
        timestamp,
        jobId,
        devId,
        '', 
        '', 
        note || 'Manually marked as sent',
        markedBy
      ]);
      marked++;
    });
    
    SpreadsheetApp.flush();
    return { success: true, marked: marked };
    
  } catch (e) {
    console.error("Error in markAsManualSent:", e);
    return { success: false, error: e.message };
  }
}

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
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const templates = [];
    
    const requestedJobId = String(jobId).trim();
    
    for(let i = 1; i < data.length; i++) {
      const rawJobId = data[i][3];
      const templateJobId = String(rawJobId === undefined || rawJobId === null ? '' : rawJobId).trim();
      
      if(templateJobId === requestedJobId) {
        templates.push({
          name: String(data[i][0] || `Template ${i}`).trim(),
          subject: String(data[i][1] || '').trim(),
          body: String(data[i][2] || ''),
          jobId: templateJobId,
          createdDate: data[i][4] ? new Date(data[i][4]).toLocaleString() : 'Unknown'
        });
      }
    }
    return templates;
    
  } catch(e) {
    console.error("getJobTemplates Error:", e);
    return [];
  }
}

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

function testTemplateLoading() {
  const testJobId = "51000";
  return { 
    forJob: getJobTemplates(testJobId), 
    all: getAllTemplates(),
    jobIdTested: testJobId
  };
}

/**
 * UPDATED sendBulkEmails
 * Removed logic that wrote to 'Negotiation_State'.
 * Now only sends emails and logs to 'Email_Logs'.
 */
function sendBulkEmails(recipients, senderName, subject, htmlBody, jobId, opts) {
  const url = getStoredSheetUrl();
  if(!url) return {success: false, sent: 0, errors: ["No config URL set. Please configure in Settings."]};
  
  try {
    const ss = SpreadsheetApp.openByUrl(url);
    validateAndFixSheets(ss);
    
    const logSheet = ss.getSheetByName("Email_Logs");
    const templateSheet = ss.getSheetByName("Email_Templates");
    
    if(!logSheet) {
      return {success: false, sent: 0, errors: ["Email_Logs sheet not found. Please re-save your config."]};
    }
    
    // ============ SAVE TEMPLATE IF REQUESTED ============
    if(opts && opts.shouldSave && opts.templateName && templateSheet) {
      try {
        const templateName = opts.templateName.trim();
        const existingData = templateSheet.getDataRange().getValues();
        let existingRow = -1;
        
        for(let i = 1; i < existingData.length; i++) {
          const existingName = String(existingData[i][0]).trim();
          const existingJobId = String(existingData[i][3]).trim();
          
          if(existingName === templateName && existingJobId === String(jobId)) {
            existingRow = i + 1;
            break;
          }
        }
        
        if(existingRow > 0) {
          templateSheet.getRange(existingRow, 2, 1, 2).setValues([[subject, htmlBody]]);
          templateSheet.getRange(existingRow, 5).setValue(new Date());
        } else {
          templateSheet.appendRow([templateName, subject, htmlBody, jobId, new Date()]);
        }
      } catch(templateError) {
        console.error("Error saving template:", templateError);
      }
    }
    
    // ============ EMAIL SENDING LOGIC (NO NEGOTIATION TRACKING) ============
    const labelName = `Job-${jobId}`;
    const labelId = getOrCreateLabelId(labelName);
    
    let count = 0;
    let errors = [];
    let skipped = 0;

    recipients.forEach(r => {
      try {
        const body = htmlBody.replace(/{{name}}/gi, r.name.split(' ')[0]);
        const rawMessage = createMimeMessage(senderName, r.email, subject, body);
        const message = Gmail.Users.Messages.send({ raw: rawMessage }, 'me');
        const threadId = message.threadId;
        
        if (labelId) {
          Gmail.Users.Threads.modify({ addLabelIds: [labelId] }, 'me', threadId);
        }
        
        logSheet.appendRow([new Date(), jobId, r.email, r.name, threadId, "Initial"]);
        
        // REMOVED: Writing to Negotiation_State
        
        count++;
      } catch(e) { 
        console.error("Send error:", e); 
        errors.push(`Failed for ${r.email}: ${e.message}`);
      }
    });
    
    SpreadsheetApp.flush();
    
    return {success: true, sent: count, skipped: skipped, errors: errors};
    
  } catch(e) {
    console.error("Bulk email error:", e);
    return {success: false, sent: 0, errors: [e.message]};
  }
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
  } catch(e) { 
    console.error("Label error:", e);
    return null; 
  }
}

// ======================================================
// ===          AUTOMATED NEGOTIATION AGENT           ===
// ======================================================

function runAutoNegotiator() {
  // DISABLED: All negotiation logic removed because supporting sheets were removed.
  return {
    status: "Disabled", 
    message: "Negotiation features have been disabled in this version.", 
    stats: null, 
    log: [{type: 'info', message: 'Negotiation system is inactive.'}]
  };
}

// Helper function kept but unused/disabled
function processJobNegotiations(jobId, rules, ss, faqContent) {
  return {
    replied:0, 
    escalated:0, 
    accepted:0, 
    skipped:0, 
    processed:0, 
    log:[{type:'info', message:'Disabled'}]
  }; 
}

// --- AI PROMPT BUILDERS (DISABLED) ---
function buildNegotiationPrompt() { return ""; }
function buildForceNegotiationPrompt() { return ""; }
function buildAcceptancePrompt() { return ""; }
function escalateToHuman() {}
function markCompleted() {}
function callAI() { return ""; }

// --- UTILITY FUNCTIONS ---

function getTaskStats() {
  // DISABLED: Stats rely on Negotiation sheets
  return { total: 0, active: 0, human: 0, accepted: 0 };
}
