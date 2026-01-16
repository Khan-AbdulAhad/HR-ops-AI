# Turing AI Recruiter V2

A Google Apps Script-based recruitment automation platform for managing developer outreach, AI-powered rate negotiation, automated follow-up emails, analytics, and learning systems.

---

## Table of Contents

1. [Features](#features)
2. [Pipeline Stages](#pipeline-stages)
3. [Prerequisites](#prerequisites)
4. [Setup Guide](#setup-guide)
   - [Step 1: Create Google Sheet](#step-1-create-google-sheet)
   - [Step 2: Deploy Google Apps Script](#step-2-deploy-google-apps-script)
   - [Step 3: Configure API Keys](#step-3-configure-api-keys)
   - [Step 4: Set Up Triggers](#step-4-set-up-triggers)
   - [Step 5: First Run Configuration](#step-5-first-run-configuration)
5. [Usage Guide](#usage-guide)
   - [Outreach Manager](#outreach-manager)
   - [Negotiation AI](#negotiation-ai)
   - [Task List](#task-list)
   - [Follow-Ups](#follow-ups)
   - [Analytics Dashboard](#analytics-dashboard)
   - [Learning System](#learning-system)
   - [AI Testing](#ai-testing)
6. [Sheet Structure](#sheet-structure)
7. [Important Notes & Best Practices](#important-notes--best-practices)
8. [Troubleshooting](#troubleshooting)

---

## Features

### Core Features
- **Developer Fetching**: Query BigQuery for candidates by Job ID and pipeline stage
- **Bulk Email Sending**: Send personalized outreach emails with `{{name}}` placeholder
- **Gmail Integration**: Automatic labeling, thread tracking, and reply management
- **Email Templates**: Save and reuse email templates (global or job-specific)
- **Manual Entry**: Add candidates manually for testing or external data entry
- **Dark Mode**: Toggle between light and dark themes

### AI-Powered Negotiation
- **Automated Rate Negotiation**: 2-attempt strategy with configurable target/max rates
- **Region-Based Pricing**: Different rates for US/Canada, Europe, LATAM, APAC, India
- **Negotiation Styles**: Friendly, Professional, Empathetic, or High-pressure
- **Smart Escalation**: Auto-escalates to human after 2 failed attempts with AI-generated summary
- **AI Notes & Summaries**: Comprehensive AI-generated summaries for each candidate

### Follow-Up System
- **Automated Follow-Ups**:
  - 1st follow-up: 12 hours after initial email (if no response)
  - 2nd follow-up: 28 hours after initial email (if no response)
- **Response Detection**: Automatically stops follow-ups when candidate responds
- **AI-Generated Content**: Contextual follow-up messages based on job description
- **Unresponsive Tracking**: Auto-marks candidates as unresponsive after 76 hours

### Data Gathering
- **Missing Information Detection**: AI identifies missing candidate data
- **Data Gathering Emails**: Automated requests for missing information
- **Completeness Tracking**: Track data gathering progress per candidate

### Export & Filtering
- **CSV Export**: Download candidate data with all fields (ID, Name, Email, Phone, Country, Status, Type)
- **Copy to Clipboard**: Quick copy selected rows to clipboard
- **Country Filtering**: Filter candidates by country before export
- **Paste & Filter**: Bulk filter by pasting email list or developer IDs
- **Status Filters**: Filter by New, Sent, Manual Sent, Sub-Con, Agency, Direct

### Analytics Dashboard
- **Comprehensive Metrics**: Track outreach, negotiations, acceptances, escalations
- **Role-Based Access**: Different views for Admin, TL, TM, TA, Manager roles
- **Date Range Filtering**: Custom date ranges with presets (Today, 7 days, 30 days, All)
- **User Activity Tracking**: Per-user analytics for team management
- **Job-Specific Stats**: Filter analytics by Job ID

### Learning System
- **AI Learning Extraction**: Extract learnings from successful human escalations
- **Learning Case Management**: Approve/reject/edit learning cases
- **FAQ Consolidation**: Automatically add approved learnings to FAQ database
- **Weekly Consolidation**: Optional automated learning consolidation

### Daily Reports
- **Automated Daily Reports**: Get activity summaries emailed to you
- **Per-Job Statistics**: AI replies, human negotiations, data gathering, follow-ups per Job ID
- **HTML Table Format**: Beautiful email reports with stats cards and tables
- **Historical Tracking**: All reports saved to Daily_Reports sheet for reference

### User Management
- **Role-Based Access Control**: Admin, TL, TM, TA, Manager, Other roles
- **Page Access Control**: Configure which roles can access which features
- **New User Notifications**: Admin notifications for new system users
- **Analytics Viewers**: Manage who can view analytics data

### Notification System
- **Notification Center**: Centralized notifications in the header
- **Task Notifications**: Alerts for escalations, new tasks, etc.
- **Mark All as Read**: Quick notification management

---

## Pipeline Stages

The system supports the following pipeline stages for candidate filtering:

| Stage | System Name | Description |
|-------|-------------|-------------|
| Interested | `is_interested` | Candidates who expressed interest |
| Passed VetSmith | `vetsmith_passed` | Passed VetSmith screening |
| Passed Internal Interviews | `passed-internal-interviews` | Passed internal interview process |
| Pending Review | `pending-review` | Awaiting review |
| Completed Testing | `completed-testing` | Completed technical testing |
| Developer Backout | `developer-backout` | Developer withdrew from process |
| On Hold - Onboarding | `on-hold-onboarding` | Onboarding paused |
| Pending Onboarding | `pending-vetting` | Awaiting onboarding |
| Ready for Selection | `ready-for-selection` | Ready for client selection |
| Selected for Trial | `selected-for-trial` | Selected for trial period |

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
   - **Description**: `HR-Ops-AI V12`
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

Triggers automate the follow-up emails, AI negotiation, and daily reports. **The app can now automatically create these triggers for you!**

#### Automatic Setup (Recommended)
1. Open your deployed Web App URL
2. Click the **Config** button (gear icon)
3. Look at the **Automated Triggers** section
4. If any triggers are missing (shown in yellow), click **"Create Missing Triggers"**
5. The app will automatically create:
   - **AI Negotiation Processor** - Runs every hour to process negotiations
   - **Follow-Up Processor** - Runs every hour to send follow-up emails
   - **Daily Report** - Runs daily at 8 AM to send activity reports
   - **Weekly Learning Consolidation** - Optional weekly trigger for learning system

#### Manual Setup (Alternative)
If you prefer to set up triggers manually:

1. In Apps Script, click **Triggers** (clock icon on the left)
2. Click **+ Add Trigger**

| Trigger | Function | Type | Interval |
|---------|----------|------|----------|
| AI Processor | `runHourlyAITrigger` | Time-driven (Hour) | Every hour |
| Follow-Up Processor | `runFollowUpProcessor` | Time-driven (Hour) | Every hour |
| Daily Report | `runDailyReportTrigger` | Time-driven (Day) | 8 AM |
| Learning Consolidation | `runWeeklyLearningConsolidation` | Time-driven (Week) | Weekly |

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
   - Passed Internal Interviews
   - Pending Review
   - Completed Testing
   - Developer Backout
   - On Hold - Onboarding
   - Pending Onboarding
   - Ready for Selection
   - Selected for Trial
3. Click **Fetch Data**
4. Review the results table showing:
   - Developer ID, Name, Email, Phone, Country
   - **Type**: Independent / Agency / Sub-Con
   - **Stage**: Current pipeline stage
   - **History**: New / Sent / Manual

#### Filtering Results

- **Status Filter Chips**: Click to filter by New, Sent (Followup), M.Sent (Manual), Sub-Con, Agency, Direct
- **Search Bar**: Type to search by name, email, or Dev ID
- **Country Filter**: Click to filter by specific countries
- **Paste & Filter**: Paste a list of emails or Dev IDs to bulk filter

#### Sending Emails

1. Select candidates using checkboxes
2. Click **Compose Email** (or **Send Initial Outreach**)
3. (Optional) Select a saved template or create new
4. Enter **Subject** and **Body**
   - Use `{{name}}` to personalize with first name
5. Click **Send Emails**
6. (Optional) Click the save icon to save as template

#### Manual Entry

For adding candidates manually or testing:
1. Click **Manual Entry** button
2. Enter candidate details (email, name, region - one per line)
3. Click **Load Data**
4. Candidates appear in the table for email composition

#### Export Options

- **Copy**: Copy selected rows to clipboard (includes all fields)
- **CSV**: Download selected rows as CSV file
- **Select All / Deselect**: Bulk selection controls

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
2. Set **Target Rate** ($/hr) - AI's initial offer (80% of this)
3. Set **Max Rate** ($/hr) - Highest rate AI can offer
4. Choose **Negotiation Style**:
   - Friendly but firm
   - Professional and direct
   - Empathetic
   - High-pressure / Urgent
5. Add **Special Rules** (optional)
6. Add **Job Description** (helps AI answer questions)
7. Add **Available Start Dates** (optional)
8. Click **Save Configuration**

#### Feature Toggles

Each job can have these features enabled/disabled:
- **Negotiation**: Enable AI rate negotiation
- **Follow-Up**: Enable automated follow-up emails
- **Data Gathering**: Enable AI to request missing information

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
- Search by email or Dev ID

#### Running AI Manually
Click **Run AI Now** to trigger negotiation processing immediately.

#### Processing Human Escalations
Click **Process Human** to import completed human escalations back into the system.

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
- **Unresponsive**: No response after all follow-ups

#### Manual Processing
Click **Process Follow-Ups Now** to:
- Check for candidate responses
- Send due follow-ups immediately
- View processing log

#### Reset Statuses
Click **Reset Follow-Up Statuses** to reset statuses for reprocessing.

> With hourly trigger set up, this runs automatically.

### Analytics Dashboard

Comprehensive reporting and metrics.

#### Metrics Overview
- Total Outreach count
- Active candidates count
- Negotiation replies count
- Accepted offers count
- Escalations count
- Follow-up pending count
- Unresponsive count
- Data gathering count

#### Filtering
- **User Filter**: Filter by user (admin only, or view own data)
- **Job ID Filter**: Filter by specific job
- **Date Range**: Select custom range or use presets (Today, 7 days, 30 days, All)

#### Role-Based Access
- **Admin**: Full access to all data and user management
- **TL (Team Lead)**: Full access to all data
- **TM/TA/Manager**: Analytics only, own data only
- **Other**: Full operational access, own analytics only

#### Manage Viewers
Admins can manage who has access to analytics via the **Manage Viewers** button.

### Learning System

Extract learnings from human escalations to improve AI.

#### Learning Tab
- View pending learning cases extracted from human escalations
- Each case shows the scenario and proposed learning
- Approve, edit, or reject each learning case

#### Consolidation
- Approved learnings are consolidated into the FAQ database
- Run manually or set up weekly automatic consolidation
- Consolidated learnings improve future AI responses

### AI Testing

Test AI responses before deployment (Admin only).

#### Testing Modes
- **Negotiation Test**: Test rate negotiation responses
- **Follow-Up Test**: Test follow-up email generation
- **Data Gathering Test**: Test data gathering email generation

#### Interactive Testing
- Enter candidate scenario
- Get AI response preview
- Iterate and refine before production use

---

## Sheet Structure

The app automatically creates these sheets:

| Sheet | Purpose |
|-------|---------|
| `Email_Logs` | All sent emails with timestamps, thread IDs, country/region |
| `Email_Templates` | Saved email templates |
| `Email_Mismatch_Reports` | Candidates who replied from different email |
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
| `Unresponsive_Devs` | Candidates marked as unresponsive |
| `Learning_Cases` | AI learning cases from human escalations |
| `Analytics_Viewers` | Role-based analytics access control |
| `Page_Access` | Role-based page access configuration |
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
- `AI-Managed` label prevents interference with personal emails
- `Awaiting-Response` labels track follow-up status

### Rate Negotiation Strategy
1. **Attempt 1**: AI offers 80% of target rate
2. **Attempt 2**: AI offers 100% of target rate
3. **Escalation**: After 2 rejections, thread is marked for human follow-up

### Data Security
- Never commit `OPENAI_API_KEY` to version control
- Keep your Google Sheet URL private
- Review BigQuery access permissions regularly
- AI-Managed Gmail label prevents system from accessing personal emails

### Performance Tips
- Use specific pipeline stages to reduce query size
- Set up triggers during off-peak hours
- Monitor `Data_Fetch_Logs` for usage patterns
- In-memory caching improves response times

### Email Best Practices
- Always test with a small batch first
- Check Gmail's sending limits (500/day for regular, 2000/day for Workspace)
- Review sent emails in Gmail to verify formatting
- Use Manual Entry to test email templates

### Follow-Up Timing
- 1st follow-up: 12 hours after initial send
- 2nd follow-up: 28 hours after initial send
- Unresponsive: 76 hours after initial send
- Modify timing in `code.gs` → `FOLLOW_UP_CONFIG` object

### Dark Mode
- Toggle dark mode using the moon/sun icon in the header
- Preference is saved in browser localStorage
- All UI components support dark mode

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

### Pipeline Stage Not Working
- Verify stage is in STAGE_CONFIG in code.gs
- Check system_name matches database values
- Ensure UI chip exists in index.html

### Analytics Not Loading
- Check your role in Analytics_Viewers sheet
- Verify you have access to view analytics
- Contact admin if you need elevated access

---

## Support

For issues or feature requests, please open an issue on the GitHub repository.

---

## Version History

- **V12**: Added Passed Internal Interviews stage, Manual Entry (renamed from Test Mode), enhanced analytics, learning system
- **V11**: Added follow-up automation, templates, user tracking, agency support, dark mode
- **V10**: Added region-based rate tiers, improved negotiation flow
- **V9**: Initial release with core features

---

*Built for Turing recruitment operations*
