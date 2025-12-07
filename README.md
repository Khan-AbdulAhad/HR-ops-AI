# Turing AI Recruiter V12

A Google Apps Script-based recruitment automation platform for managing developer outreach, AI-powered rate negotiation, and automated follow-up emails.

---

## Table of Contents

1. [Features](#features)
2. [Prerequisites](#prerequisites)
3. [Setup Guide](#setup-guide)
   - [Step 1: Create Google Sheet](#step-1-create-google-sheet)
   - [Step 2: Deploy Google Apps Script](#step-2-deploy-google-apps-script)
   - [Step 3: Configure API Keys](#step-3-configure-api-keys)
   - [Step 4: Set Up Triggers](#step-4-set-up-triggers)
   - [Step 5: First Run Configuration](#step-5-first-run-configuration)
4. [Usage Guide](#usage-guide)
   - [Outreach Manager](#outreach-manager)
   - [Negotiation AI](#negotiation-ai)
   - [Task List](#task-list)
   - [Follow-Ups](#follow-ups)
5. [Sheet Structure](#sheet-structure)
6. [Important Notes & Best Practices](#important-notes--best-practices)
7. [Troubleshooting](#troubleshooting)

---

## Features

### Core Features
- **Developer Fetching**: Query BigQuery for candidates by Job ID and pipeline stage
- **Bulk Email Sending**: Send personalized outreach emails with `{{name}}` placeholder
- **Gmail Integration**: Automatic labeling, thread tracking, and reply management
- **Email Templates**: Save and reuse email templates (global or job-specific)

### AI-Powered Negotiation
- **Automated Rate Negotiation**: 2-attempt strategy with configurable target/max rates
- **Region-Based Pricing**: Different rates for US/Canada, Europe, LATAM, APAC, India
- **Negotiation Styles**: Friendly, Professional, Empathetic, or High-pressure
- **Smart Escalation**: Auto-escalates to human after 2 failed attempts with AI-generated summary

### Follow-Up System
- **Automated Follow-Ups**:
  - 1st follow-up: 12 hours after initial email (if no response)
  - 2nd follow-up: 28 hours after initial email (if no response)
- **Response Detection**: Automatically stops follow-ups when candidate responds
- **AI-Generated Content**: Contextual follow-up messages based on job description

### Tracking & Logging
- **User Email Display**: Shows logged-in user in the header
- **Data Consumption Logging**: Tracks BigQuery usage, email sends, and API calls
- **Manual Sent Logs**: Mark candidates as contacted outside this system
- **Agency/Subcontractor Detection**: Identifies independent vs agency developers

### Daily Reports
- **Automated Daily Reports**: Get activity summaries emailed to you
- **Per-Job Statistics**: AI replies, human negotiations, data gathering, follow-ups per Job ID
- **HTML Table Format**: Beautiful email reports with stats cards and tables
- **Historical Tracking**: All reports saved to Daily_Reports sheet for reference
- **Trigger Support**: Set up daily triggers to receive automated reports

---

## Prerequisites

Before setting up, ensure you have:

1. **Google Account** with access to:
   - Google Apps Script
   - Google Sheets
   - Gmail API
   - BigQuery (with access to Turing's dataset)

2. **OpenAI API Key** for AI-powered negotiation

3. **BigQuery Access** with:
   - Project ID configured
   - External connection to Turing's MySQL database

---

## Setup Guide

### Step 1: Create Google Sheet

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Name it something like `HR-Ops-AI-Data`
3. Copy the spreadsheet URL (you'll need this later)
4. **Important**: Share the sheet with yourself (the same Google account you'll use for the script)

> The required sheets will be automatically created when you configure the app.

### Step 2: Deploy Google Apps Script

1. Go to [Google Apps Script](https://script.google.com)
2. Click **New Project**
3. Delete any default code in `Code.gs`
4. Copy the entire contents of `code.gs` from this repository and paste it
5. Create a new HTML file:
   - Click **+** next to Files → **HTML**
   - Name it `index` (without .html extension)
   - Delete any default content
   - Copy the entire contents of `index.html` from this repository and paste it
6. Save both files (Ctrl+S or Cmd+S)

#### Enable Required APIs

1. In Apps Script, click **Services** (+ icon on the left sidebar)
2. Add the following services:
   - **Gmail API** (v1)
   - **BigQuery API** (v2)
3. Click **Add** for each

#### Deploy as Web App

1. Click **Deploy** → **New deployment**
2. Click the gear icon → **Web app**
3. Configure:
   - **Description**: `HR-Ops-AI V11`
   - **Execute as**: `Me (your email)`
   - **Who has access**: `Only myself` (or your organization)
4. Click **Deploy**
5. **Authorize** the app when prompted (review and allow all permissions)
6. Copy the **Web app URL** - this is your app's access link

### Step 3: Configure API Keys

1. In Apps Script, go to **Project Settings** (gear icon)
2. Scroll down to **Script Properties**
3. Click **Add script property**
4. Add the following property:

| Property | Value |
|----------|-------|
| `OPENAI_API_KEY` | Your OpenAI API key (sk-...) |

5. Click **Save script properties**

#### Update BigQuery Configuration

In `code.gs`, update the `CONFIG` object at the top of the file:

```javascript
const CONFIG = {
  PROJECT_ID: "your-gcp-project-id",           // Your Google Cloud Project ID
  EXTERNAL_CONN: "your-project.dataset.connection"  // Your BigQuery external connection
};
```

### Step 4: Set Up Triggers (Automatic)

Triggers automate the follow-up emails and daily reports. **The app can now automatically create these triggers for you!**

#### Automatic Setup (Recommended)
1. Open your deployed Web App URL
2. Click the **Config** button (gear icon)
3. Look at the **Automated Triggers** section
4. If any triggers are missing (shown in yellow), click **"Create Missing Triggers"**
5. The app will automatically create:
   - **Follow-Up Processor** - Runs every hour to send follow-up emails
   - **Daily Report** - Runs daily at 8 AM to send activity reports

#### Manual Setup (Alternative)
If you prefer to set up triggers manually:

1. In Apps Script, click **Triggers** (clock icon on the left)
2. Click **+ Add Trigger**

| Trigger | Function | Type | Interval |
|---------|----------|------|----------|
| Follow-Up Processor | `runFollowUpProcessor` | Time-driven (Hour) | Every hour |
| Daily Report | `runDailyReportTrigger` | Time-driven (Day) | 8 AM |

3. Click **Save** for each trigger
4. Authorize when prompted

#### Trigger Status
You can check trigger status anytime:
- Open Config modal to see which triggers are active
- Green = Active and running
- Yellow = Missing (needs to be created)
- Click "Refresh" to update status

### Step 5: First Run Configuration

1. Open your deployed Web App URL
2. Click the **Config** button (gear icon in top-right)
3. Paste your Google Sheet URL
4. Click **Save**

The app will automatically create all required sheets with proper headers.

---

## Usage Guide

### Outreach Manager

The first tab for fetching developers and sending outreach emails.

#### Fetching Developers

1. Enter a **Job ID** (e.g., `51000`)
2. Select one or more **pipeline stages**:
   - Interested
   - Passed VetSmith
   - Pending Review
   - Completed Testing
   - Ready for Selection
   - etc.
3. Click **Fetch**
4. Review the results table showing:
   - Developer ID, Name, Email
   - **Type**: Independent / Agency / Sub-Con
   - **Stage**: Current pipeline stage
   - **History**: New / Sent / Manual

#### Sending Emails

1. Select candidates using checkboxes
2. Click **Compose Email**
3. (Optional) Select a saved template or create new
4. Enter **Subject** and **Body**
   - Use `{{name}}` to personalize with first name
5. Click **Send Emails**
6. (Optional) Click the save icon to save as template

#### Mark as Manually Sent

For candidates you've contacted outside this system:
1. Select the candidates
2. Click **Mark Sent**
3. Add an optional note
4. They'll show "Manual" in the History column

### Negotiation AI

Configure how the AI negotiates rates for each job.

#### Basic Configuration

1. Select or enter a **Job ID**
2. Set **Target Rate** ($/hr) - AI's initial offer (70% of this)
3. Set **Max Rate** ($/hr) - Highest rate AI can offer
4. Choose **Negotiation Style**:
   - Friendly but firm
   - Professional and direct
   - Empathetic
   - High-pressure / Urgent
5. Add **Special Rules** (optional)
6. Add **Job Description** (helps AI answer questions)
7. Click **Save Configuration**

#### Region-Based Rate Tiers

For different rates by region:

1. Scroll to **Rate Tiers** section
2. Click **Add Tier**
3. Select **Region**: US/Canada, Europe, LATAM, APAC, India, or Default
4. Set **Target** and **Max** rates for that region
5. Add optional **Notes**
6. Click **Save**

> If a candidate's region matches a tier, those rates override the default.

### Task List

Monitor and manage all active negotiations.

#### Stats Dashboard
- **Total Active**: All ongoing negotiations
- **AI Negotiating**: Being handled by AI
- **Needs Human**: Escalated after 2 attempts
- **Offer Accepted**: Candidates who agreed

#### Filtering Tasks
- Filter by **Job ID**
- Filter by **Status**: Initial, AI Active, Human Required, Accepted

#### Running AI Manually
Click **Run AI Now** to trigger negotiation processing immediately.

#### Viewing Details
- Click the **notes icon** to view AI conversation summary
- Click the **envelope icon** to open Gmail thread

#### Completing Tasks
1. Select tasks with checkboxes
2. Click **Complete Selected**
3. Choose final status

### Follow-Ups

Manage automated follow-up emails.

#### Stats Cards
- **Awaiting Response**: Initial email sent, waiting
- **1st Follow-up Sent**: 12-hour follow-up sent
- **2nd Follow-up Sent**: 28-hour follow-up sent
- **Responded**: Candidate replied (removed from queue)

#### Manual Processing
Click **Process Follow-Ups Now** to:
- Check for candidate responses
- Send due follow-ups immediately
- View processing log

> With hourly trigger set up, this runs automatically.

---

## Sheet Structure

The app automatically creates these sheets:

| Sheet | Purpose |
|-------|---------|
| `Email_Logs` | All sent emails with timestamps and thread IDs |
| `Email_Templates` | Saved email templates |
| `Negotiation_Config` | Job-specific negotiation settings |
| `Negotiation_State` | Current state of each negotiation |
| `Negotiation_Tasks` | Accepted offers and completed negotiations |
| `Negotiation_Completed` | Archive of finished negotiations |
| `Negotiation_FAQs` | Q&A database for AI responses |
| `Rate_Tiers` | Region-based rate configurations |
| `Manual_Sent_Logs` | Candidates marked as manually contacted |
| `Data_Fetch_Logs` | API usage and data consumption tracking |
| `Follow_Up_Queue` | Automated follow-up status tracking |
| `Daily_Reports` | Historical daily activity reports per Job ID |
| `Job_XXXXX_Details` | Per-job candidate details (created dynamically) |

### Adding FAQs

To help AI answer common questions:
1. Open your Google Sheet
2. Go to `Negotiation_FAQs` sheet
3. Add rows with:
   - Column A: Question (e.g., "What are the working hours?")
   - Column B: Answer (e.g., "Flexible hours with 4-hour overlap with US timezone")

---

## Important Notes & Best Practices

### Gmail Labels
- The app creates labels like `Job-51000` for each job
- `Human-Negotiation` label marks escalated threads
- `Completed` label marks finished negotiations

### Rate Negotiation Strategy
1. **Attempt 1**: AI offers 70% of target rate
2. **Attempt 2**: AI offers 100% of target rate
3. **Escalation**: After 2 rejections, thread is marked for human follow-up

### Data Security
- Never commit `OPENAI_API_KEY` to version control
- Keep your Google Sheet URL private
- Review BigQuery access permissions regularly

### Performance Tips
- Use specific pipeline stages to reduce query size
- Set up triggers during off-peak hours
- Monitor `Data_Fetch_Logs` for usage patterns

### Email Best Practices
- Always test with a small batch first
- Check Gmail's sending limits (500/day for regular, 2000/day for Workspace)
- Review sent emails in Gmail to verify formatting

### Follow-Up Timing
- 1st follow-up: 12 hours after initial send
- 2nd follow-up: 28 hours after initial send
- Modify timing in `code.gs` → `FOLLOW_UP_CONFIG` object

---

## Troubleshooting

### "No config URL set"
- Click Config (gear icon) and enter your Google Sheet URL
- Make sure the URL is for the spreadsheet, not a specific sheet tab

### "Could not access Sheet"
- Verify the Sheet URL is correct
- Ensure you have edit access to the spreadsheet
- Check if the spreadsheet was deleted or moved

### BigQuery Errors
- Verify `CONFIG.PROJECT_ID` is correct
- Check BigQuery API is enabled in your GCP project
- Ensure external connection is properly configured

### Emails Not Sending
- Check Gmail API is enabled in Apps Script
- Verify you haven't hit Gmail's daily sending limit
- Look for errors in Apps Script execution log (View → Executions)

### AI Not Responding
- Verify `OPENAI_API_KEY` is set in Script Properties
- Check your OpenAI account has available credits
- Review the execution log for API errors

### Follow-Ups Not Processing
- Ensure `runFollowUpProcessor` trigger is set up
- Check the trigger is authorized
- Verify `Follow_Up_Queue` sheet exists

### Missing Developer Data
- Verify BigQuery connection is working
- Check STAGE_CONFIG matches your database schema
- Review execution logs for SQL errors

---

## Support

For issues or feature requests, please open an issue on the GitHub repository.

---

## Version History

- **V11**: Added follow-up automation, templates, user tracking, agency support
- **V10**: Added region-based rate tiers, improved negotiation flow
- **V9**: Initial release with core features

---

*Built for Turing recruitment operations*
