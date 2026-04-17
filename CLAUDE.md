# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Turing AI Recruiter V12** — a Google Apps Script (GAS) application that automates developer recruitment outreach, AI-powered rate negotiation, follow-up emails, analytics, and a learning system for AI improvement. It integrates Gmail, Google Sheets, BigQuery, and OpenAI.

There is no build step, no package manager, and no test framework. Code runs directly inside Google Apps Script (server-side `.gs` files + an `index.html` frontend).

---

## Repository Structure

| File | Purpose |
|------|---------|
| `code.gs` | ~24k-line GAS backend — all server-side logic |
| `index.html` | ~16k-line frontend — Tailwind CSS, jQuery, Chart.js |

Everything lives in two files. There are no subdirectories, no npm dependencies, and no compiled assets.

---

## Development Workflow

### Making Changes
1. Edit `code.gs` (server-side) or `index.html` (client-side) locally.
2. Deploy by pasting content into the Google Apps Script editor, or push via `clasp` if configured.
3. Test via the deployed Web App URL.

### Frontend ↔ Backend Communication
All UI calls use the GAS `google.script.run` pattern — there is no REST API or fetch():
```javascript
google.script.run
  .withSuccessHandler(callback)
  .withFailureHandler(errorHandler)
  .serverFunctionName(params);
```
Server functions must be top-level functions in `code.gs` (not nested).

### Script Properties (Secrets & Config)
Stored in GAS Project Settings → Script Properties, never in code:
```
PROJECT_ID         = turing-230020
DATASET_ID         = turing-230020
EXTERNAL_CONN      = turing-230020.us.matching-vetting-prod-readonly
OPENAI_API_KEY     = sk-...
ANALYTICS_SHEET_ID = (optional override)
EMAIL_SENDER_NAME_CUSTOM = (optional)
EMAIL_SIGNATURE_CUSTOM   = (optional)
FOLLOW_UP_TIMING_CONFIG  = (optional JSON)
```

### User Properties (Per-User)
Stored via `PropertiesService.getUserProperties()`:
- `LOG_SHEET_URL` — Primary Google Sheet URL (main database)
- `JOBS_SHEET_URL` — Jobs detail spreadsheet URL

---

## Architecture: Key Systems

### 1. BigQuery → Candidate Fetching
`fetchData()` queries BigQuery via an external connection, wrapping an inner SQL query in `EXTERNAL_QUERY(...)`. Results are polled up to 300 seconds. All queries are logged in the `Data_Fetch_Logs` sheet via `logDataConsumption()`.

### 2. AI Negotiation Engine
- **Hourly trigger:** `runAutoNegotiator()` → `processJobNegotiations(jobId, rules, ss, faqContent)`
- **2-attempt strategy:** Attempt 1 = 80% of target rate; Attempt 2 = 100% of max rate (or region-specific via `getRateForRegion()`)
- **Escalation:** After 2 rejections, or when candidate asks about internal identifiers, AI labels thread `Human-Negotiation` and stops
- **Rate limit:** `MAX_AI_EMAILS_PER_TRIGGER_RUN = 50` (AI emails only; manual sends are unlimited)
- **Concurrency:** `LockService.getScriptLock()` prevents duplicate trigger runs

### 3. Follow-Up Automation
- **Hourly trigger:** `runFollowUpProcessor()` → `processFollowUpQueue()`
- **Timing defaults** (configurable via `FOLLOW_UP_TIMING_CONFIG` Script Property):
  - Follow-up 1: 12 hours after initial email
  - Follow-up 2: 28 hours after initial email
  - Marked Unresponsive: 76 hours after initial email
- Data-gathering follow-ups (missing info): 12h / 28h / 48h / 72h → "Incomplete Data - Unresponsive"

### 4. Email & Thread Management
- **Normalization:** `normalizeEmail()` strips Gmail dot-variants so `john.doe@gmail.com` == `johndoe@gmail.com`
- **Mismatch detection:** `logEmailMismatch()` logs when replies come from a different address
- **Sanitization:** `sanitizeEmailContent()` and `validateEmailForSending()` strip internal identifiers before every send
- **Gmail label lifecycle:** `AI-Managed` → `Follow-Up-1-Sent` → `Follow-Up-2-Sent` → `Unresponsive` or `Completed` / `Human-Negotiation`

### 5. Google Sheets as Database
The primary spreadsheet (URL from `LOG_SHEET_URL`) holds all persistent state. Key sheets:

| Sheet | Purpose |
|-------|---------|
| `Email_Logs` | Every outreach email sent |
| `Negotiation_State` | In-progress negotiations |
| `Negotiation_Tasks` | Accepted offers |
| `Negotiation_Completed` | Archived finished negotiations |
| `Follow_Up_Queue` | Active follow-up tracking |
| `Unresponsive_Devs` | Candidates who never replied |
| `Rate_Tiers` | Region-based rate config per job |
| `AI_Learning_Cases` | Learnings extracted from human escalations |
| `Negotiation_FAQs` | Q&A used in AI prompts |
| `Reconciliation_Log` | Audit trail for automatic status fixes |
| `Job_XXXXX_Details` | Dynamic per-job candidate detail sheets |

Sheet access is cached: `getCachedSheetData()` (default 60s TTL). The spreadsheet object itself is cached in `_cachedSpreadsheet`.

### 6. Role-Based Access
Roles: `Admin`, `TL` (full access), `TM`/`TA`/`Manager` (own data only), `Other`. Controlled via the `Page_Access` sheet and enforced server-side in every analytics function.

### 7. Learning System
Human escalations are analyzed by `extractLearningCaseFromEscalation()`, stored in `AI_Learning_Cases`, reviewed via the Learning tab, then consolidated into `Negotiation_FAQs` via `consolidateLearningsToFAQs()`. FAQ content is injected into every AI negotiation prompt.

### 8. Trigger System
Three time-based GAS triggers must exist:
| Function | Frequency |
|----------|-----------|
| `runAutoNegotiator()` | Every 1 hour |
| `runFollowUpProcessor()` | Every 1 hour |
| `runOnboardingIssueScan()` | Every 6 hours |

Manage via Config tab → Automated Triggers, or call `createAllMissingTriggers()` / `getTriggerStatus()` directly.

---

## Critical Conventions

- **No nested server functions called from UI** — `google.script.run` only resolves top-level GAS functions.
- **Always sanitize before sending** — call `sanitizeEmailContent()` + `validateEmailForSending()` before any programmatic send.
- **Duplicate prevention is multi-layered** — check `Negotiation_State`, `Follow_Up_Queue`, and `Job_XXXXX_Details` before inserting a new record; rely on `normalizeEmail()` for address matching.
- **Script Properties for secrets, never hardcoded** — API keys and project IDs are read via `PropertiesService.getScriptProperties().getProperty(key)`.
- **50-email cap per trigger run** — the `MAX_AI_EMAILS_PER_TRIGGER_RUN` constant applies to all AI-generated emails in a single trigger execution.
- **LockService guards critical sections** — acquire `LockService.getScriptLock()` in any function that reads-then-writes negotiation or follow-up state.
