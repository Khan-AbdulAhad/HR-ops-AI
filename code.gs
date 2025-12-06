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
      .setTitle('Turing AI Recruiter V10')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getStoredSheetUrl() {
  return PropertiesService.getUserProperties().getProperty('LOG_SHEET_URL') || "";
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
  const sheets = ['Email_Logs', 'Email_Templates', 'Negotiation_Config', 'Negotiation_Tasks', 'Negotiation_State', 'Negotiation_FAQs', 'Negotiation_Completed', 'Candidate_Details'];
  sheets.forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  // UPDATED: Removed Walk Away Rate column
  const confSheet = ss.getSheetByName('Negotiation_Config');
  if (confSheet.getLastRow() === 0) confSheet.appendRow(['Job ID', 'Target Rate', 'Max Rate', 'Style', 'Special Rules', 'Job Description', 'Last Updated']);

  const taskSheet = ss.getSheetByName('Negotiation_Tasks');
  if (taskSheet.getLastRow() === 0) taskSheet.appendRow(['Timestamp', 'Job ID', 'Name', 'Email', 'Agreed Rate', 'Status', 'Dev ID', 'Thread ID']);

  const stateSheet = ss.getSheetByName('Negotiation_State');
  if (stateSheet.getLastRow() === 0) stateSheet.appendRow(['Email', 'Job ID', 'Attempt Count', 'Last Offer', 'Status', 'Last Reply Time', 'Dev ID', 'Name', 'AI Notes', 'Thread ID']);

  const faqSheet = ss.getSheetByName('Negotiation_FAQs');
  if (faqSheet.getLastRow() === 0) faqSheet.appendRow(['Question', 'Answer']);

  const compSheet = ss.getSheetByName('Negotiation_Completed');
  if (compSheet.getLastRow() === 0) compSheet.appendRow(['Timestamp', 'Job ID', 'Email', 'Name', 'Final Status', 'Notes', 'Dev ID']);

  // Candidate Details sheet - stores details shared by candidates (start dates, etc.)
  const detailsSheet = ss.getSheetByName('Candidate_Details');
  if (detailsSheet.getLastRow() === 0) {
    detailsSheet.appendRow([
      'Timestamp', 'Job ID', 'Email', 'Name', 'Dev ID', 'Thread ID',
      'Start Date', 'Notice Period', 'Expected Rate', 'Current Company',
      'Years of Experience', 'Availability (hrs/week)', 'Location/Timezone',
      'Has Own Equipment', 'Other Details', 'Negotiation Status', 'Raw Response'
    ]);
  }
}

// --- NEGOTIATION CONFIGURATION (UPDATED - No Walk Away) ---

function saveNegotiationConfig(jobId, config) {
  const url = getStoredSheetUrl();
  if(!url) return;
  const ss = SpreadsheetApp.openByUrl(url);
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
}

function getNegotiationConfig(jobId) {
  const url = getStoredSheetUrl();
  if(!url) return null;
  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName('Negotiation_Config');
  const data = sheet.getDataRange().getValues();
  
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
  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName('Negotiation_Config');
  const data = sheet.getDataRange().getValues();
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
  const ss = SpreadsheetApp.openByUrl(url);
  const sheet = ss.getSheetByName('Negotiation_FAQs');
  const data = sheet.getDataRange().getValues();

  if(data.length <= 1) return "No specific FAQs available.";

  let faqText = "";
  for(let i=1; i<data.length; i++) {
    if(data[i][0]) {
      faqText += `Q: ${data[i][0]}\nA: ${data[i][1]}\n---\n`;
    }
  }
  return faqText;
}

// --- CANDIDATE DETAILS EXTRACTION ---

/**
 * Extract candidate details from their email response using AI
 * Returns a structured object with all extracted details
 */
function extractCandidateDetails(candidateMessage, candidateName, candidateEmail, jobId) {
  const prompt = `
You are analyzing a candidate's email response to extract specific details they've shared.

CANDIDATE'S MESSAGE:
"${candidateMessage}"

CANDIDATE INFO:
- Name: ${candidateName}
- Email: ${candidateEmail}
- Job ID: ${jobId}

TASK:
Extract any of the following details the candidate has shared. If a detail is not mentioned, use "NOT_PROVIDED".
If the candidate is negotiating or being vague about a detail (e.g., "I'll let you know later", "depends on the offer"), mark it as "NEGOTIATING".

Return a JSON object with these exact fields:
{
  "start_date": "The date or timeframe they can start (e.g., 'Jan 15, 2025', '2 weeks notice', 'immediately')",
  "notice_period": "Current notice period if mentioned (e.g., '2 weeks', '1 month', 'none')",
  "expected_rate": "Rate they're asking for (e.g., '$40/hr', '$50-60/hr')",
  "current_company": "Their current employer if mentioned",
  "years_experience": "Years of experience if mentioned (e.g., '5 years', '10+')",
  "availability_hours": "Hours per week they can work (e.g., '40 hrs', '20-30 hrs', 'full-time')",
  "location_timezone": "Their location or timezone if mentioned",
  "has_equipment": "Whether they have their own laptop/equipment (e.g., 'yes', 'no', 'needs laptop')",
  "other_details": "Any other relevant details shared (skills, certifications, preferences)",
  "is_negotiating": true/false (true if they're negotiating rate or being vague about key details),
  "negotiation_notes": "Brief note about what they're negotiating or unclear about"
}

IMPORTANT:
- Only extract what is EXPLICITLY stated in the message
- Do not infer or assume details that aren't clearly mentioned
- If they give a range (e.g., "$40-50/hr"), capture the full range
- If they're asking questions instead of providing details, note that in negotiation_notes

Return ONLY the JSON object, no other text.
`;

  try {
    const response = callAI(prompt);
    // Clean up the response - remove any markdown code blocks
    let cleanResponse = response.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const details = JSON.parse(cleanResponse);
    return details;
  } catch(e) {
    console.error("Failed to extract candidate details:", e);
    return {
      start_date: "PARSE_ERROR",
      notice_period: "PARSE_ERROR",
      expected_rate: "PARSE_ERROR",
      current_company: "PARSE_ERROR",
      years_experience: "PARSE_ERROR",
      availability_hours: "PARSE_ERROR",
      location_timezone: "PARSE_ERROR",
      has_equipment: "PARSE_ERROR",
      other_details: "Failed to parse response",
      is_negotiating: false,
      negotiation_notes: ""
    };
  }
}

/**
 * Save or update candidate details in the Candidate_Details sheet
 * Will update existing row if candidate already exists for this job, otherwise creates new row
 */
function saveCandidateDetails(ss, candidateEmail, jobId, candidateName, devId, threadId, details, rawMessage) {
  const detailsSheet = ss.getSheetByName('Candidate_Details');
  if(!detailsSheet) return { success: false, message: "Candidate_Details sheet not found" };

  const cleanEmail = String(candidateEmail).toLowerCase().trim();
  const cleanJobId = String(jobId);

  // Check if this candidate already has details for this job
  const data = detailsSheet.getDataRange().getValues();
  let existingRowIndex = -1;

  for(let i=1; i<data.length; i++) {
    if(String(data[i][2]).toLowerCase() === cleanEmail && String(data[i][1]) === cleanJobId) {
      existingRowIndex = i + 1;
      break;
    }
  }

  // Prepare the row data
  const negotiationStatus = details.is_negotiating ?
    `Negotiating: ${details.negotiation_notes || 'Rate/details under discussion'}` :
    'Details Provided';

  const rowData = [
    new Date(),
    jobId,
    candidateEmail,
    candidateName,
    devId || 'N/A',
    threadId || '',
    details.start_date || 'NOT_PROVIDED',
    details.notice_period || 'NOT_PROVIDED',
    details.expected_rate || 'NOT_PROVIDED',
    details.current_company || 'NOT_PROVIDED',
    details.years_experience || 'NOT_PROVIDED',
    details.availability_hours || 'NOT_PROVIDED',
    details.location_timezone || 'NOT_PROVIDED',
    details.has_equipment || 'NOT_PROVIDED',
    details.other_details || '',
    negotiationStatus,
    rawMessage.substring(0, 500) // Truncate raw message to prevent sheet overflow
  ];

  if(existingRowIndex > -1) {
    // Update existing row - merge new details with existing (keep non-empty values)
    const existingRow = data[existingRowIndex - 1];
    for(let col=6; col<15; col++) { // Columns 7-15 are the detail fields
      if(rowData[col] === 'NOT_PROVIDED' && existingRow[col] && existingRow[col] !== 'NOT_PROVIDED') {
        rowData[col] = existingRow[col]; // Keep existing value if new one is empty
      }
    }
    detailsSheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
    return { success: true, message: "Updated existing candidate details", isUpdate: true };
  } else {
    // Add new row
    detailsSheet.appendRow(rowData);
    return { success: true, message: "Added new candidate details", isUpdate: false };
  }
}

/**
 * Get candidate details from the sheet
 */
function getCandidateDetails(email, jobId) {
  const url = getStoredSheetUrl();
  if(!url) return null;

  const ss = SpreadsheetApp.openByUrl(url);
  const detailsSheet = ss.getSheetByName('Candidate_Details');
  if(!detailsSheet) return null;

  const cleanEmail = String(email).toLowerCase().trim();
  const cleanJobId = String(jobId);

  const data = detailsSheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(String(data[i][2]).toLowerCase() === cleanEmail && String(data[i][1]) === cleanJobId) {
      return {
        timestamp: data[i][0],
        jobId: data[i][1],
        email: data[i][2],
        name: data[i][3],
        devId: data[i][4],
        threadId: data[i][5],
        startDate: data[i][6],
        noticePeriod: data[i][7],
        expectedRate: data[i][8],
        currentCompany: data[i][9],
        yearsExperience: data[i][10],
        availabilityHours: data[i][11],
        locationTimezone: data[i][12],
        hasEquipment: data[i][13],
        otherDetails: data[i][14],
        negotiationStatus: data[i][15],
        rawResponse: data[i][16]
      };
    }
  }
  return null;
}

/**
 * Get all candidate details for a specific job
 */
function getAllCandidateDetailsForJob(jobId) {
  const url = getStoredSheetUrl();
  if(!url) return [];

  const ss = SpreadsheetApp.openByUrl(url);
  const detailsSheet = ss.getSheetByName('Candidate_Details');
  if(!detailsSheet) return [];

  const cleanJobId = String(jobId);
  const data = detailsSheet.getDataRange().getValues();
  const results = [];

  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === cleanJobId) {
      results.push({
        timestamp: data[i][0],
        jobId: data[i][1],
        email: data[i][2],
        name: data[i][3],
        devId: data[i][4],
        threadId: data[i][5],
        startDate: data[i][6],
        noticePeriod: data[i][7],
        expectedRate: data[i][8],
        currentCompany: data[i][9],
        yearsExperience: data[i][10],
        availabilityHours: data[i][11],
        locationTimezone: data[i][12],
        hasEquipment: data[i][13],
        otherDetails: data[i][14],
        negotiationStatus: data[i][15]
      });
    }
  }
  return results;
}

// --- OPTIMIZED TASK LIST MANAGEMENT ---

function getAllTasks(filters) {
  const url = getStoredSheetUrl();
  if(!url) return { tasks: [], jobIds: [] };
  
  const ss = SpreadsheetApp.openByUrl(url);
  ensureSheetsExist(ss);
  
  const tasks = [];
  const jobIdSet = new Set();

  // Apply filters if provided
  const jobFilter = filters?.jobId || 'all';
  const statusFilter = filters?.status || 'all';

  // 1. Get Active Negotiations (State)
  const stateSheet = ss.getSheetByName('Negotiation_State');
  if(!stateSheet) return { tasks: [], jobIds: [] };
  const stateData = stateSheet.getDataRange().getValues();
  
  for(let i=1; i<stateData.length; i++) {
    if(!stateData[i][0]) continue;
    
    const jobId = String(stateData[i][1]);
    const status = stateData[i][4] || 'Active';
    const attempts = Number(stateData[i][2]) || 0;
    
    // Collect all job IDs for filter dropdown
    jobIdSet.add(jobId);
    
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

  // 2. Get Accepted Offers (Tasks) - only if status filter allows
  if(statusFilter === 'all' || statusFilter === 'Offer Accepted') {
    const taskSheet = ss.getSheetByName('Negotiation_Tasks');
    const taskData = taskSheet.getDataRange().getValues();
    
    for(let i=1; i<taskData.length; i++) {
      if(!taskData[i][3]) continue;
      if(taskData[i][5] === 'Archived') continue;
      
      const jobId = String(taskData[i][1]);
      jobIdSet.add(jobId);
      
      // Apply job filter
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
    jobIds: Array.from(jobIdSet).sort() 
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
  const detailsSheet = ss.getSheetByName('Candidate_Details');

  if(!compSheet || !stateSheet || !taskSheet) {
    return { success: false, message: "Required sheets not found" };
  }

  let moved = false;
  let taskInfo = { jobId: '', name: '', devId: '', threadId: '' };
  const cleanEmail = String(email).toLowerCase();

  // First, find info from Task Sheet
  const taskData = taskSheet.getDataRange().getValues();
  for(let i=taskData.length-1; i>=1; i--) {
    if(String(taskData[i][3]).toLowerCase() === cleanEmail) {
      // If jobIdFilter is provided, only match if job ID matches
      if(jobIdFilter && String(taskData[i][1]) !== String(jobIdFilter)) continue;

      taskInfo = {
        jobId: taskData[i][1],
        name: taskData[i][2],
        devId: taskData[i][6] || 'N/A',
        threadId: taskData[i][7] || ''
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

  // UPDATE CANDIDATE_DETAILS SHEET - Mark as completed
  if(moved && detailsSheet && taskInfo.jobId) {
    try {
      const detailsData = detailsSheet.getDataRange().getValues();
      for(let i=1; i<detailsData.length; i++) {
        if(String(detailsData[i][2]).toLowerCase() === cleanEmail && String(detailsData[i][1]) === String(taskInfo.jobId)) {
          // Update the negotiation status column (column 16 = index 15)
          detailsSheet.getRange(i+1, 16).setValue(`Completed: ${finalStatus || 'Done'}`);
          break;
        }
      }
    } catch(e) {
      console.error("Failed to update Candidate_Details:", e);
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

  return { success: moved, message: moved ? "Moved to completed" : "Email not found", jobId: taskInfo.jobId };
}

// --- OUTREACH FETCHING ---

function getDevelopers(jobId, selectedStages) {
  if (!jobId || !selectedStages || selectedStages.length === 0) throw new Error("Missing inputs");
  const cleanJobId = Number(jobId);
  const logs = getEmailLogs(cleanJobId); 
  let queryChunks = [];
  selectedStages.forEach(stage => {
    const config = STAGE_CONFIG[stage];
    if (config) {
      const q = config.type === 'flag' 
        ? `SELECT DISTINCT main.developer_id, "${stage}" as stage_label FROM ${config.table} main WHERE main.job_id = ${cleanJobId} AND ${config.condition}`
        : `SELECT DISTINCT main.developer_id, "${stage}" as stage_label FROM ${config.table} main LEFT JOIN ms2_job_match_status s ON main.job_match_status_id = s.id WHERE main.job_id = ${cleanJobId} AND s.system_name = "${config.system_name}"`;
      queryChunks.push(q);
    }
  });
  const unionQuery = queryChunks.join(' UNION ALL ');
  const innerQuery = `
    WITH target_devs AS (${unionQuery}),
    unique_ids AS (SELECT DISTINCT developer_id FROM target_devs),
    dev_details AS (SELECT id, full_name, email FROM user_list_v4 WHERE id IN (SELECT developer_id FROM unique_ids))
    SELECT td.developer_id, d.full_name, d.email, td.stage_label
    FROM target_devs td JOIN dev_details d ON td.developer_id = d.id
  `;
  const finalSql = `SELECT * FROM EXTERNAL_QUERY("${CONFIG.EXTERNAL_CONN}", """${innerQuery}""")`;
  try {
    let queryResults = BigQuery.Jobs.query({ query: finalSql, useLegacySql: false }, CONFIG.PROJECT_ID);
    let job = queryResults.jobReference;
    while (!queryResults.jobComplete) {
      Utilities.sleep(500);
      queryResults = BigQuery.Jobs.getQueryResults(CONFIG.PROJECT_ID, job.jobId);
    }
    return (queryResults.rows || []).map(row => {
      const email = row.f[2].v;
      const log = logs ? logs.get(email.toLowerCase()) : null;
      return {
        developer_id: row.f[0].v,
        full_name: row.f[1].v,
        email: email,
        status: row.f[3].v,
        is_sent: !!log,
        sent_count: log ? log.count : 0
      };
    });
  } catch (e) {
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

// UPDATED: Send Email with progress callback support
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
      
      // Add to state with thread ID
      stateSheet.appendRow([r.email, jobId, 0, "Initial Sent", "Initial Outreach", new Date(), r.devId || "N/A", r.name, "", threadId]);
      existingEmails.add(emailKey);
      
      count++;
    } catch(e) { 
      console.error(e); 
      errors.push(`Failed for ${r.email}: ${e.message}`);
    }
  });
  
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
      devId: stateData[r][6] || 'N/A'
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
      return;
    }

    const msgs = thread.getMessages();
    const lastMsg = msgs[msgs.length - 1];
    const lastSender = lastMsg.getFrom().toLowerCase();
    
    if (myEmail && myEmail.length > 3 && lastSender.indexOf(myEmail) > -1) {
      jobStats.skipped++;
      return;
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
    
    if (state && state.status === 'Human-Negotiation') {
      jobStats.skipped++;
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
      jobStats.escalated++;
      jobStats.log.push({type: 'warning', message: `${candidateEmail} escalated: Max attempts reached`});
      return;
    }
    
    // Build conversation history
    const recentMsgs = msgs.slice(-5);
    const conversationHistory = recentMsgs.map(m => {
      const from = m.getFrom();
      const isMe = from.toLowerCase().indexOf(myEmail) > -1;
      return `[${isMe ? 'ME' : 'CANDIDATE'}]: ${m.getPlainBody().substring(0, 400)}`;
    }).join("\n---\n");

    // Extract and save candidate details from their latest message
    const candidateLatestMessage = lastMsg.getPlainBody();
    try {
      const extractedDetails = extractCandidateDetails(candidateLatestMessage, candidateName, candidateEmail, jobId);
      const saveResult = saveCandidateDetails(ss, candidateEmail, jobId, candidateName, devId, thread.getId(), extractedDetails, candidateLatestMessage);
      if(saveResult.success) {
        jobStats.detailsExtracted++;
        jobStats.log.push({type: 'info', message: `${candidateEmail} - Details extracted: ${extractedDetails.is_negotiating ? 'Negotiating' : 'Provided'}`});
      }
    } catch(detailsError) {
      // Don't block negotiation if details extraction fails
      console.error("Details extraction error for " + candidateEmail + ":", detailsError);
    }

    const isFirstResponse = attempts === 0;
    
    // Calculate offer amounts based on attempt
    const targetRate = Number(rules.target) || 25;
    const maxRate = Number(rules.max) || 30;
    const firstOfferRate = Math.round(targetRate * 0.7); // 70% of target for first offer
    const secondOfferRate = targetRate; // 100% of target for second offer
    const currentOfferRate = attempts === 0 ? firstOfferRate : secondOfferRate;
    
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
**FIRST ATTEMPT - Start Low:**
- Offer Rate: $${firstOfferRate}/hr (this is your opening offer)
- You can go up to $${secondOfferRate}/hr if they push back strongly
- NEVER mention any "target rate" or "budget" or "aim for" - just state your offer confidently
` : `
**SECOND ATTEMPT - Final Offer:**
- Offer Rate: $${secondOfferRate}/hr (this is your best offer)
- Maximum you can approve: $${maxRate}/hr (only if they're very firm and qualified)
- NEVER mention any "target rate" or "budget" - just state what you can offer
`}

- Negotiation Style: ${rules.style}
- Special Instructions: ${rules.special || 'None'}
- Current Attempt: ${attempts + 1} of 2

=== CRITICAL RULES - READ CAREFULLY ===
1. **NEVER reveal internal numbers**: Do not say "we aim for", "our target is", "our budget is", or "we're looking at". Just state your offer directly.
2. **Be confident**: Say "We can offer $X/hr for this role" - not "We're hoping for" or "We'd like to offer"
3. **This is FREELANCE**: Never mention full-time benefits, team culture, or long-term employment
4. **Answer questions on NEW LINES**: If answering multiple questions, put each answer on a separate line for readability

=== FREQUENTLY ASKED QUESTIONS ===
${faqContent}

**IMPORTANT FAQ INSTRUCTIONS:**
If the candidate asks ANY question that matches or is similar to the FAQs above:
1. Use the provided answer as your source of truth
2. Paraphrase naturally in your own words
3. Put each answer on a SEPARATE LINE for better readability
4. If their question isn't covered by FAQs, answer based on the job description

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
THIS IS THE FIRST RESPONSE - Offer $${firstOfferRate}/hr
- Present this rate confidently without justification
- If they asked for a higher rate, acknowledge their experience but present your offer
- Answer any questions they have (on separate lines if multiple)
- NEVER say "we aim for" or "our target is" or reveal any internal numbers
- NEVER use ACTION: ESCALATE on first attempt
` : `
THIS IS ATTEMPT ${attempts + 1} - Offer up to $${secondOfferRate}/hr
- You can now offer $${secondOfferRate}/hr as your "best offer"
- If they're highly qualified and firm, you can go up to $${maxRate}/hr maximum
- Be more flexible but still professional
- You may escalate ONLY if they demand significantly above $${maxRate}/hr
`}

=== RESPONSE FORMAT ===
1. If they clearly ACCEPT an offer at or below $${maxRate}/hr: 
   Reply with: ACTION: ACCEPT [$RATE]

2. If this is attempt 2 AND they refuse to negotiate reasonably (demanding way above $${maxRate}/hr):
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
        
        // Calculate offer for this attempt
        const targetRate = Number(rules.target) || 25;
        const currentOffer = attempts === 0 ? Math.round(targetRate * 0.7) : targetRate;
        
        const retryPrompt = `
You are a recruiter at Turing. You MUST write a negotiation email - escalation is NOT an option.

CANDIDATE'S ACTUAL MESSAGE:
"${candidateMessage}"

CANDIDATE NAME: ${candidateName}

YOUR OFFER FOR THIS ATTEMPT: $${currentOffer}/hr
(You can go slightly higher if they push back, up to $${rules.max}/hr maximum)

JOB CONTEXT:
${rules.jobDescription || 'Freelance AI/Tech role'}

FAQs (use these to answer any questions):
${faqContent}

CRITICAL RULES:
1. NEVER say "we aim for", "our target is", "our budget is" - just state your offer confidently
2. Say "We can offer $${currentOffer}/hr for this role" - be direct and confident
3. If answering multiple questions, put each answer on a SEPARATE LINE
4. This is FREELANCE - never mention full-time benefits

TASK: 
1. Read the candidate's message above carefully
2. Note the ACTUAL rate they mentioned (if any)
3. Answer any questions they asked (on separate lines)
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

      // Update Candidate_Details sheet with escalation status
      const detailsSheet = ss.getSheetByName('Candidate_Details');
      if(detailsSheet) {
        try {
          const detailsData = detailsSheet.getDataRange().getValues();
          for(let d=1; d<detailsData.length; d++) {
            if(String(detailsData[d][2]).toLowerCase() === cleanCandidateEmail && String(detailsData[d][1]) === String(jobId)) {
              detailsSheet.getRange(d+1, 16).setValue(`Escalated: ${finalReason}`);
              break;
            }
          }
        } catch(detailsErr) {
          console.error("Failed to update details sheet:", detailsErr);
        }
      }

      jobStats.escalated++;
      jobStats.log.push({type: 'warning', message: `${candidateEmail} escalated: ${finalReason}`});
    } 
    else if (aiResponse.includes("ACTION: ACCEPT")) {
      const rateMatch = aiResponse.match(/\[([^\]]+)\]/);
      const rate = rateMatch ? rateMatch[1].replace('$','').replace('/hr','').trim() : rules.target;

      taskSheet.appendRow([new Date(), jobId, candidateName, candidateEmail, rate, "Pending Archive", devId, thread.getId()]);

      // Update Candidate_Details sheet with accepted status
      const detailsSheet = ss.getSheetByName('Candidate_Details');
      if(detailsSheet) {
        try {
          const detailsData = detailsSheet.getDataRange().getValues();
          for(let d=1; d<detailsData.length; d++) {
            if(String(detailsData[d][2]).toLowerCase() === cleanCandidateEmail && String(detailsData[d][1]) === String(jobId)) {
              detailsSheet.getRange(d+1, 16).setValue(`Offer Accepted at $${rate}/hr`);
              break;
            }
          }
        } catch(detailsErr) {
          console.error("Failed to update details sheet:", detailsErr);
        }
      }

      const acceptPrompt = `
Write a brief, warm confirmation email to ${candidateName.split(' ')[0]} confirming they've accepted the rate of $${rate}/hr for a freelance position at Turing.

Keep it short (3-4 sentences). Mention:
- Thank them for accepting
- Confirm the rate
- Say next steps/contract details will follow shortly

FORMAT:
Hi [Name],

[Your message]

Best regards,
Turing Recruitment Team
`;
      const acceptEmail = callAI(acceptPrompt);
      sendReplyWithSenderName(thread, acceptEmail, 'Recruiter');
      markCompleted(thread);

      if(stateRowIndex > -1) {
        stateSheet.deleteRow(stateRowIndex);
        stateMap.delete(stateKey);
      }

      jobStats.accepted++;
      jobStats.log.push({type: 'success', message: `${candidateEmail} ACCEPTED at $${rate}/hr`});
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
    
    // Generate handoff message using AI
    const firstName = candidateName ? candidateName.split(' ')[0] : 'there';
    
    const handoffPrompt = `
You are a recruiter at Turing. You need to write a brief, warm handoff email to a candidate.

CANDIDATE NAME: ${firstName}
CONTEXT: ${conversationContext || 'Rate negotiation needs human attention'}

TASK:
Write a short, professional email (3-4 sentences) that:
1. Thanks them for their response and patience
2. Lets them know a member of the Talent Operations team will take over
3. Assures them they'll hear back shortly about the final rate and next steps
4. Sounds natural and warm, not robotic

DO NOT:
- Apologize excessively
- Mention "AI" or "automated"
- Sound like a template
- Be too formal or stiff

FORMAT:
Hi ${firstName},

[Your message - 3-4 natural sentences]

Best regards,
Turing Recruitment Team

Write ONLY the email, nothing else.
`;
    
    const handoffMessage = callAI(handoffPrompt);
    sendReplyWithSenderName(thread, handoffMessage, 'Recruiter');
    
  } catch(e) {
    console.error("Failed to escalate to human:", e);
    // Fallback to simple reply if AI fails
    try {
      const fallbackMsg = `Hi ${candidateName ? candidateName.split(' ')[0] : 'there'},\n\nThank you for your response. A member of our Talent Operations team will be in touch shortly to continue this discussion.\n\nBest regards,\nTuring Recruitment Team`;
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
  const detailsSheet = ss.getSheetByName('Candidate_Details');

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

          // Update Candidate_Details sheet
          if(detailsSheet) {
            try {
              const detailsData = detailsSheet.getDataRange().getValues();
              for(let d=1; d<detailsData.length; d++) {
                if(String(detailsData[d][2]).toLowerCase() === candidateEmail && String(detailsData[d][1]) === String(jobId)) {
                  detailsSheet.getRange(d+1, 16).setValue('Completed (Gmail Sync)');
                  break;
                }
              }
            } catch(detailsErr) {
              console.error("Failed to update details sheet:", detailsErr);
            }
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

          // Update Candidate_Details sheet
          if(detailsSheet) {
            try {
              const detailsData = detailsSheet.getDataRange().getValues();
              for(let d=1; d<detailsData.length; d++) {
                if(String(detailsData[d][2]).toLowerCase() === candidateEmail && String(detailsData[d][1]) === String(jobId)) {
                  detailsSheet.getRange(d+1, 16).setValue(`Accepted at $${agreedRate}/hr (Gmail Sync)`);
                  break;
                }
              }
            } catch(detailsErr) {
              console.error("Failed to update details sheet:", detailsErr);
            }
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
