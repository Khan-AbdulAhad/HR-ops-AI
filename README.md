# Turing AI Recruiter V2

A Google Apps Script-based recruitment automation platform for managing developer outreach, AI-powered rate negotiation, automated follow-up emails, centralized team analytics, status reconciliation, and a learning system for AI improvement.

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
   - [Email Settings](#email-settings)
   - [AI Testing](#ai-testing)
6. [Sheet Structure](#sheet-structure)
7. [Critical System Behaviors](#critical-system-behaviors)
   - [Candidate Responds After Being Marked Unresponsive](#candidate-responds-after-being-marked-unresponsive)
   - [Candidate Responds After Rate is Already Accepted](#candidate-responds-after-rate-is-already-accepted)
   - [What Triggers AI Escalation to Human](#what-triggers-ai-escalation-to-human)
   - [What Happens When Both Negotiation Attempts Are Rejected](#what-happens-when-both-negotiation-attempts-are-rejected)
   - [Ambiguous / Non-Committal Replies](#ambiguous--non-committal-replies)
   - [Email Reply from a Different Address (Mismatch)](#email-reply-from-a-different-address-mismatch)
   - [Duplicate Send Prevention](#duplicate-send-prevention)
   - [AI Email Rate Limit](#ai-email-rate-limit-50-per-trigger-run)
   - [Data Gathering — Full Flow](#data-gathering--full-flow)
   - [Follow-Up & Unresponsive Timing](#follow-up--unresponsive-timing-exact-values)
   - [Gmail Label Full Lifecycle](#gmail-label-full-lifecycle)
   - [WhatsApp Outreach](#whatsapp-outreach)
   - [BigQuery Returns No Results](#bigquery-returns-no-results)
8. [Important Notes & Best Practices](#important-notes--best-practices)
9. [Troubleshooting](#troubleshooting)

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

### User Management
- **Role-Based Access Control**: Admin, TL, TM, TA, Manager, Other roles
- **Page Access Control**: Configure which roles can access which features
- **New User Notifications**: Admin notifications for new system users
- **Analytics Viewers**: Manage who can view analytics data

#### Analytics Visibility Matrix

| Role | Sees |
|---|---|
| **Admin** | All users, all teams |
| **TM** | All users, all teams (job stakeholder — full cross-team visibility) |
| **TA** | All users, all teams (job stakeholder — full cross-team visibility) |
| **TL** | Their team members + self |
| **Manager** | Their team members + self |
| **Other / TOS** | Own data only |

Enforced server-side in every analytics reader via `seesAllAnalyticsData_()` (`code.gs:18641`).

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
| Selected for Internal Interviews | `selected-for-internal-interviews` | Selected for internal interview process |
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
   - **Description**: `HR-Ops-AI V2`
   - **Execute as**: `Me (your email)`
   - **Who has access**: `Only myself` (or your organization)
4. Click **Deploy**
5. **Authorize** the app when prompted (review and allow all permissions)
6. Copy the **Web app URL** - this is your app's access link

### Step 3: Configure API Keys

1. In Apps Script, go to **Project Settings** (gear icon)
2. Scroll down to **Script Properties**
3. Click **Add script property**
4. Add the following properties:

| Property | Value | Required |
|----------|-------|----------|
| `OPENAI_API_KEY` | Your OpenAI API key (sk-...) | Yes |
| `PROJECT_ID` | Your Google Cloud Project ID | Yes |
| `DATASET_ID` | Your BigQuery dataset ID | Yes |
| `EXTERNAL_CONN` | Your BigQuery external connection (e.g. `project.region.connection`) | Yes |
| `ANALYTICS_SHEET_ID` | Central analytics spreadsheet ID (optional override) | No |
| `EMAIL_SENDER_NAME_CUSTOM` | Custom sender display name override | No |
| `EMAIL_SIGNATURE_CUSTOM` | Custom email signature override | No |

5. Click **Save script properties**

> **Note:** `PROJECT_ID`, `DATASET_ID`, and `EXTERNAL_CONN` were previously hardcoded in the `CONFIG` object. They are now configured via Script Properties, which is more secure and easier to update without redeploying.

### Step 4: Set Up Triggers (Automatic)

Triggers automate the follow-up emails and AI negotiation. **The app can now automatically create these triggers for you!**

#### Automatic Setup (Recommended)
1. Open your deployed Web App URL
2. Click the **Config** button (gear icon)
3. Look at the **Automated Triggers** section
4. If any triggers are missing (shown in yellow), click **"Create Missing Triggers"**
5. The app will automatically create:
   - **AI Negotiation Processor** - Runs every hour to process negotiations
   - **Follow-Up Processor** - Runs every hour to send follow-up emails
   - **Weekly Learning Consolidation** - Optional weekly trigger for learning system

#### Manual Setup (Alternative)
If you prefer to set up triggers manually:

1. In Apps Script, click **Triggers** (clock icon on the left)
2. Click **+ Add Trigger**

| Trigger | Function | Type | Interval |
|---------|----------|------|----------|
| AI Processor | `runAutoNegotiator` | Time-driven (Hour) | Every hour |
| Follow-Up Processor | `runFollowUpProcessor` | Time-driven (Hour) | Every hour |
| Onboarding Scan | `runOnboardingIssueScan` | Time-driven (6 hours) | Every 6 hours |
| Learning Consolidation | `consolidateApprovedLearnings` | Time-driven (Week) | Weekly, Sunday 2 AM |

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
   - Selected for Internal Interviews
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

### Email Settings

Customize how outgoing emails appear to recipients.

1. Open the **Config** modal (gear icon)
2. Navigate to the **Email Settings** section
3. Configure:
   - **Sender Name**: Display name shown to email recipients (overrides the default `Turing Recruitment Team`)
   - **Email Signature**: Signature appended to outgoing emails (overrides the default `Turing | Talent Operations`)
4. Click **Save**

> Settings are stored in Script Properties and persist across deployments.

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
| `Unresponsive_Devs` | Candidates marked as unresponsive |
| `Reconciliation_Log` | Audit trail for every status change made by the reconciler |
| `Response_Times` | Centralized team-wide response-time analytics (write-through projection) |
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

## Critical System Behaviors

This section covers important behaviors that happen automatically in the background. Understanding these will help you interpret the data in your sheets and avoid confusion when things don't behave as expected.

---

### Candidate Responds After Being Marked Unresponsive

If a candidate was marked as **Unresponsive** (76 hours with no reply) but later replies in the same Gmail thread, the system **automatically detects this on the next trigger run**:

- Status is changed back to **Responded** in `Follow_Up_Queue`
- Gmail labels are updated (`Unresponsive` removed)
- If negotiation is active for that job, the AI resumes processing their reply
- The candidate is **not re-sent** follow-up emails — their reply is treated as a live response

> This means unresponsive is not permanent. Check `Follow_Up_Queue` and `Negotiation_State` if you see unexpected re-activation.

---

### Candidate Responds After Rate is Already Accepted

If a candidate sends another message after their acceptance was recorded:

- The system checks whether **data gathering is still incomplete** for that candidate
- **If data is still missing:** The candidate is NOT moved to `Negotiation_Completed` yet. The AI continues gathering the outstanding information (start date, notice period, equipment, etc.)
- **If all data is already collected:** The extra response is ignored for processing purposes — no new email is sent, and the candidate stays in `Negotiation_Completed`

> Log note you may see: *"Rate agreed at $X/hr but data gathering incomplete. Will NOT mark as Completed."*

---

### What Triggers AI Escalation to Human

The AI escalates to a human (applies `Human-Negotiation` Gmail label) when any of the following occur:

1. Candidate **rejects the final max-rate offer** (2nd attempt)
2. Candidate asks about **internal processes**, rate structures, other candidates, or system identifiers
3. Candidate asks a question **not covered in the FAQ database**
4. Candidate explicitly asks to **speak with a human**

**What happens on escalation:**
- The `Human-Negotiation` Gmail label is added to the thread — no email is sent to the candidate at this point
- The AI writes a **summary of the conversation and the candidate's concerns** into the AI Notes column in `Negotiation_State`
- The thread appears in the Task List under **"Needs Human"**
- A human can reply directly in Gmail — the system detects this and marks the escalation as handled

> After a human completes the negotiation, click **Process Human** in the Task List to import the result back into the system.

**Status tag sync:** If you manually apply the `Human-Negotiation` label in Gmail, the Task List status tag will update to **Human-Negotiation** as soon as you click **Refresh** (previously required waiting up to 1 hour for the hourly trigger). The AI will immediately stop sending emails to that candidate.

---

### What Happens When Both Negotiation Attempts Are Rejected

- **Attempt 1** (80% of target rate) rejected → AI sends Attempt 2
- **Attempt 2** (100% of max rate) rejected → AI triggers **escalation to human** (not auto-archive)
- The negotiation is **never automatically closed or archived** after two rejections — a human must review and take action
- Only after a human marks the outcome (e.g., "Rejected by Candidate") is the record moved to `Negotiation_Completed`

---

### Ambiguous / Non-Committal Replies

If a candidate responds with something unclear (e.g., *"I'll think about it"*, *"Let me check with my family"*, *"Maybe"*):

- The AI takes a **SOFT_HOLD** action — no acceptance or rejection is recorded
- An acknowledging reply is sent to the candidate
- The candidate's status in `Negotiation_State` **remains unchanged** (still active)
- They are **not marked unresponsive** — the 76-hour timer resets from their last reply
- They will remain in active state until they give a clear answer or the system eventually times out

---

### Email Reply from a Different Address (Mismatch)

When a candidate replies from a different email address than the one the system sent to (e.g., sent to `john.doe@gmail.com`, replied from `johndoe@gmail.com`):

- The system uses **Gmail thread matching** to identify the candidate regardless of the reply address
- For Gmail/Googlemail addresses, dots are ignored in matching (`john.doe@gmail.com` = `johndoe@gmail.com`)
- For non-Gmail addresses, dots **are significant** — `john.doe@company.com` ≠ `johndoe@company.com`
- The mismatch is logged in `Email_Mismatch_Reports` with:
  - Expected email, actual reply email, Dev ID, Thread ID, timestamp
  - `Requires Review: Yes` flag for manual investigation
- Processing continues normally — the candidate is not blocked

> Regularly check `Email_Mismatch_Reports` if you notice candidates in unexpected states.

---

### Duplicate Send Prevention

Before sending any email, the system checks whether the candidate is already in the system for the same job:

- Checks `Negotiation_State`, `Follow_Up_Queue`, and `Job_*_Details` for a matching email + Job ID combination
- Gmail dot-variants are treated as the same email for all Gmail addresses
- If a duplicate is found, **the email is silently skipped** and counted as already queued — no error is shown, no duplicate is sent
- This applies to both bulk sends from the UI and AI trigger-based sends

---

### AI Email Rate Limit (50 per trigger run)

The system caps AI-generated emails at **50 per trigger execution** (covers negotiations + follow-ups + data gathering combined). This prevents runaway sending if something is misconfigured.

- **Manual bulk sends from the UI are not affected** by this limit
- If the limit is reached mid-run, remaining candidates are skipped and processed on the next hourly trigger
- No error is shown to the user — it silently defers to the next run
- If your queue is large, candidates will be processed in batches across hourly runs

---

### Data Gathering — Full Flow

Data gathering is triggered when the AI detects that a candidate's response is missing key information (start date, notice period, weekly hours, equipment preferences, etc.).

**Timing (separate from follow-up timers):**

| Step | Trigger |
|------|---------|
| Data follow-up 1 | 12 hours after data gap detected |
| Data follow-up 2 | 28 hours after data gap detected |
| Data follow-up 3 (final) | 48 hours after data gap detected |
| Marked "Incomplete Data - Unresponsive" | 72 hours with no response to final follow-up |

**After candidate provides all data:**
- Status updated to "Data Complete" in `Negotiation_State`
- If rate negotiation is also resolved, candidate is moved to `Negotiation_Completed`
- A confirmation email is sent to the candidate
- If data gathering is complete but negotiation is pending, only the negotiation remains open

---

### Follow-Up & Unresponsive Timing (Exact Values)

| Event | Hours After Initial Send |
|-------|--------------------------|
| Follow-up 1 sent | 12 hours |
| Follow-up 2 sent | 28 hours |
| Marked as Unresponsive | 76 hours (both follow-ups must be sent first) |
| Negotiation silent follow-up 1 | 24 hours after last AI reply (no candidate response) |
| Negotiation silent follow-up 2 | 48 hours |
| Negotiation marked Unresponsive | 96 hours |

> These values can be overridden by setting `FOLLOW_UP_TIMING_CONFIG` in Script Properties (JSON format).

**When marked Unresponsive:**
1. Candidate is added to `Unresponsive_Devs` sheet
2. `Follow_Up_Queue` status is set to `Unresponsive` (row is retained for recovery)
3. Gmail labels: adds `Unresponsive`, removes `Awaiting-Response`, `Follow-Up-1-Sent`, `Follow-Up-2-Sent`
4. No further emails are sent unless the candidate responds (see above)
5. `Negotiation_State` status is updated to `Unresponsive` immediately if the row exists — if it doesn't, clicking **Refresh** will recover the row automatically

---

### Gmail Label Full Lifecycle

| Event | Labels Added | Labels Removed |
|-------|-------------|----------------|
| Email sent via app | `AI-Managed`, `Job-XXXXX` | — |
| Added to follow-up queue | `Awaiting-Response` | — |
| Follow-up 1 sent | `Follow-Up-1-Sent` | — |
| Follow-up 2 sent | `Follow-Up-2-Sent` | `Follow-Up-1-Sent` |
| Marked unresponsive | `Unresponsive` | `Awaiting-Response`, `Follow-Up-1-Sent`, `Follow-Up-2-Sent` |
| Escalated to human | `Human-Negotiation` | — |
| Human completes escalation | `Human-Negotiation-Completed` | `Human-Negotiation` |
| Negotiation accepted/completed | `Completed` | `Human-Negotiation` (if present) |
| Data gathering follow-up sent | `Data-Follow-Up-1-Sent` | — |
| Negotiation silent follow-up | `Neg-Follow-Up-1-Sent` | — |

> **Critical:** The AI will **only process email threads that have the `AI-Managed` label**. If you manually send an email outside the system, the AI will not touch that thread.

---

### WhatsApp Outreach

**WhatsApp is not implemented.** All outreach and response detection is email-only via the Gmail API. There is no WhatsApp integration in the current version.

---

### Status Reconciliation System

The system runs an automatic cross-sheet status reconciliation every time the **Refresh** button is clicked and on every hourly AI trigger run. This ensures the Task List status tag always converges with the actual data in Gmail, `Follow_Up_Queue`, and `Unresponsive_Devs`.

**What it fixes automatically:**

| Scenario | What reconciler does |
|---|---|
| Candidate in `Negotiation_Completed` but still in `Negotiation_State` | Removes duplicate State row |
| AI notes indicate acceptance but status tag is still "Follow Up" | Moves candidate to `Negotiation_Completed`, adds `Completed` Gmail label |
| Candidate moved to `Unresponsive_Devs` but status tag shows "Follow Up" | Updates status tag to `Unresponsive` |
| Candidate sent a follow-up but status tag still shows "Initial Outreach" | Updates status tag to `Follow Up` |

**Reconciliation_Log sheet:** Every change made by the reconciler is written to `Reconciliation_Log` with: timestamp, source (Refresh or AI trigger), action type, job ID, email, candidate name, previous status, new status, and the evidence that triggered the change. Use this to audit unexpected status flips.

**Missing candidates (recovery):** If a candidate exists in `Follow_Up_Queue` or `Unresponsive_Devs` but is absent from `Negotiation_State` (e.g., due to a sheet error or manual deletion), clicking **Refresh** will automatically re-create their `Negotiation_State` row with the correct status. Previously this recovery only ran on the hourly trigger.

---

### BigQuery Returns No Results

If a Fetch returns 0 candidates:

- A user-facing message is shown: *"No candidates found matching the selected criteria"*
- No emails are sent
- If the query times out (>300 seconds): *"BigQuery query timed out — please click Refresh to retry"*
- The system does not automatically retry — you must click Fetch again
- Check your Job ID, pipeline stage selection, and BigQuery connection if this happens unexpectedly

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
- Uses explicit Tailwind `dark:` classes (not relying on global CSS overrides alone)

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

- **V2** *(current)*: Full-featured release. Redesigned dashboard UI (deep indigo / slate / emerald). Complete dark mode support with explicit Tailwind `dark:` classes. Status reconciliation system (`reconcileCandidateStatuses`) with full audit trail in `Reconciliation_Log`. Centralized response-time analytics (`Response_Times` sheet) with privacy-safe hashed keys. TM/TA full cross-team analytics visibility. Regional Rate Tiers with per-country overrides. AI Learning System. Role-based access control. BigQuery config via Script Properties. Email sender name/signature customization. Human-Negotiation Gmail label syncs on Refresh. Missing candidates auto-recovered from `Follow_Up_Queue` and `Unresponsive_Devs`.
- **V1**: Initial release — core outreach, AI negotiation, follow-up automation, templates, dark mode, region-based rate tiers, learning system foundation.

---

*Built for Turing recruitment operations*
