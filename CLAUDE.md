# HR-ops-AI — Claude Code Reference

## Project Structure

Single-file Google Apps Script project:
- `code.gs` — entire backend (~24 500 lines)
- `index.html` — frontend SPA (~17 300 lines)

No build step. Deploy via Google Apps Script editor → Deploy → Web App.

---

## Sheet Schemas (Column Indices)

### Negotiation_State
`[0]Email [1]JobID [2]Attempts [3]LastOffer [4]Status [5]LastReply [6]DevID [7]Name [8]AINotes [9]ThreadID [10]Region [11]? [12]?`

Status is the authoritative runtime field for the task list display.

### Follow_Up_Queue
`[0]Email [1]Job ID [2]Thread ID [3]Name [4]Dev ID [5]Initial Send Time [6]Follow Up 1 Sent [7]Follow Up 2 Sent [8]Status [9]Last Response Time ...`

Status values: `''` / `'Responded'` / `'Unresponsive'`
f1Done/f2Done are booleans (TRUE or true string).

### Unresponsive_Devs
`[0]Email [1]Job ID [2]Name [3]Dev ID [4]Thread ID [5]Initial Send Time [6]Follow Up 1 Time [7]Follow Up 2 Time [8]Marked Unresponsive [9]Days Since Initial`

### Email_Logs
`[0]Timestamp [1]Job ID [2]Email [3]Name [4]Thread ID [5]Type [6]Country`

### Negotiation_Completed
`[0]Timestamp [1]Job ID [2]Email [3]Name [4]Final Status [5]Notes [6]Dev ID [7]Region [8]Thread ID`

Thread ID column was added in migration (ensureSheetsExist backfills from Email_Logs primary, Follow_Up_Queue fallback).

### Negotiation_Tasks
`[0]Timestamp [1]Job ID [2]? [3]Email [4]Name [5]Status [6]? [7]Thread ID`

### Reconciliation_Log
`[0]Timestamp [1]Source [2]Action [3]Job ID [4]Email [5]Name [6]From Status [7]To Status [8]Basis [9]Evidence`

Lazy-created by `logReconcileAction()`. Never pre-created in ensureSheetsExist.

---

## Key Functions

### Trigger Entry Points
| Function | Trigger | Cadence |
|---|---|---|
| `runAutoNegotiator()` | Time-based | Every 1 hour |
| `runFollowUpProcessor()` | Time-based | Every 1 hour |
| `runOnboardingIssueScan()` | Time-based | Every 6 hours |
| `consolidateApprovedLearnings()` | Time-based | Weekly, Sunday 2 AM |

### Refresh Button Path
`getAllTasks({ forceRefresh: true })` → in order:
1. Invalidate all sheet caches
2. `syncHumanNegotiationFromGmail()` — Gmail label → Negotiation_State promotion
3. `enrichNegotiationStateData(ss)` — recover missing candidates from FQ + Unresponsive_Devs
4. `reconcileCandidateStatuses(ss, 'refresh')` — cross-sheet status convergence

### runAutoNegotiator() Pipeline (key steps)
1. `syncCompletedFromGmail()` — sync Completed label → Negotiation_Completed
2. `syncHumanNegotiationFromGmail()` — Gmail H-N label → Negotiation_State
3. `processCompletedHumanEscalations()` — H-N + Completed threads → extract AI outcome
4. `generateMissingSummaries(ss)` — fill blank AI notes
5. `enrichNegotiationStateData(ss)` — enrich + recover missing state rows
6. `reconcileCandidateStatuses(ss, 'runAutoNegotiator')` — status convergence
7. Main AI loop — send negotiation emails

### Status Reconciliation (`reconcileCandidateStatuses`)
Walks `Negotiation_State`. Five cases per row:

| Case | Condition | Action |
|---|---|---|
| `DEDUP_COMPLETED` | Key in Negotiation_Completed | Delete state row |
| `DEDUP_TASKS` | Key in Negotiation_Tasks (non-archived) | Delete state row |
| `MOVED_TO_COMPLETED` | noteIndicatesAcceptance(aiNotes) && !terminal && !dataPending | Append to Negotiation_Completed; add Completed label; remove H-N label |
| `SYNCED_UNRESPONSIVE` | In Unresponsive_Devs or FQ.status=Unresponsive && status is syncable | Set status='Unresponsive' |
| `SYNCED_FOLLOW_UP` | fq.f1Done && status='Initial Outreach' | Set status='Follow Up' |

Terminal statuses (never touched): Completed, Data Complete, Offer Accepted, Not Interested, Human-Negotiation, Rate Agreed, Escalated, Pending Escalation.

Source logged to Reconciliation_Log (`'refresh'` or `'runAutoNegotiator'`).

### Recovery in `enrichNegotiationStateData`
Two passes (both inside one try-catch sharing stateEmails/completedEmails sets):
1. **Follow_Up_Queue pass** — recovers candidates in FQ but missing from State. Sets status based on fqStatus: Unresponsive → 'Unresponsive', Responded → 'Active', f1Done → 'Follow Up', else → 'Initial Outreach'.
2. **Unresponsive_Devs pass** — recovers candidates in Unresponsive_Devs but missing from State. Always sets status='Unresponsive'. Dedupes against stateEmails updated by pass 1 to prevent double-adds.

Both passes use `lookupCandidateDetails` for Name/DevID/Region enrichment.

### Gmail Label Sync
`syncHumanNegotiationFromGmail()`:
- Query: `label:Human-Negotiation -label:Completed label:<AI_MANAGED_LABEL>`
- For each thread: extract candidate email, extract Job-XXXXX label → build key
- If key found in Negotiation_State with non-H-N status → update to 'Human-Negotiation', set AI notes
- If not in State at all → append new row as 'Human-Negotiation'
- Also marks FQ.status='Responded'

---

## Identity & Matching

### Key Format
All cross-sheet lookups use composite key: `normalizeEmail(email) + '|' + String(jobId)`

Exception: `enrichNegotiationStateData` recovery uses `_` separator (not `|`). Both are consistent within each function.

### normalizeEmail(email)
- Lowercases
- For gmail.com / googlemail.com: strips dots from local part
- Example: `John.Doe@Gmail.com` → `johndoe@gmail.com`

### Thread ID
Primary identity for Gmail operations. Present in:
- Negotiation_State col [9]
- Follow_Up_Queue col [2]
- Unresponsive_Devs col [4]
- Email_Logs col [4]
- Negotiation_Completed col [8] (added via migration + all appendRow sites)

---

## Caching

```javascript
getCachedSheetData(sheetName, ttlSeconds)  // read-through cache
invalidateSheetCache(sheetName)            // manual invalidation
```

Cache is in-memory per execution. After any write to a sheet, always call `invalidateSheetCache(sheetName)` so subsequent reads in the same execution see fresh data.

---

## Schema Migration Pattern

In `ensureSheetsExist()`:
1. Check if sheet exists; if not, create with full header row.
2. If sheet exists, read headers and check for missing columns.
3. If missing: append column header, then backfill historical data via cross-sheet lookup.
4. Call `invalidateSheetCache` after backfill.

Example: Thread ID migration for Negotiation_Completed — backfills from Email_Logs (primary) then Follow_Up_Queue (fallback) using normalizedEmail|jobId key.

---

## Known Fragile Points

- `moveToUnresponsive()` silently warns if no Negotiation_State row exists (line 14951). The candidate ends up in Unresponsive_Devs without a State row. Fixed by `enrichNegotiationStateData`'s UD recovery pass running on both hourly trigger and Refresh.
- Display-layer recovery in `getAllTasks` (Email_Logs section, ~line 4350) builds a synthetic task from Email_Logs + Follow_Up_Queue but does NOT check Unresponsive_Devs. If a candidate is recovered this way instead of from State, status may show 'Follow Up' instead of 'Unresponsive'. Correct fix is ensuring State row is populated (done by `enrichNegotiationStateData` on Refresh).
- Human-Negotiation safety: the Gmail label is a signal; the State.status field is the authoritative gate. The reconciler's `isTerminal` set guards on State.status. The Gmail→State sync (`syncHumanNegotiationFromGmail`) must run before the reconciler so legitimate escalations are promoted before the terminal check.

---

## Centralized Analytics (Response_Times)

Per-user `UserProperties` mean one user's session cannot read another user's `LOG_SHEET_URL`. Team-wide analytics therefore require a **write-through projection**: each user's hourly trigger writes their derived data to a shared sheet, and viewers read from that sheet with RBAC.

### Response_Times sheet
Columns: `[0]Timestamp_Outreach [1]Timestamp_Response [2]Hours_To_Respond [3]User_Email [4]Job_ID [5]Candidate_Email_Hash [6]Dedup_Key (hashed) [7]Sync_Time`

- **Writer:** `syncResponseTimesToCentral()` — called from `runAutoNegotiator()` step 3.7, and manually from the UI "Backfill my history" button.
- **Reader:** `getTimeToResponseMetrics()` — reads the shared sheet, never the personal one, so pure TL/Manager viewers see team aggregates.
- **Backfill gate:** `hasUnsyncedResponseTimes()` returns true only when the current user has personal pairs missing from central. UI uses it to decide whether to show the Backfill button.

### Rules when adding shared analytics sheets
1. **Dedup by a hashed, stable key** — never by raw candidate email or raw Gmail thread ID. Use `sha256Hex_('tid:' + threadId)` (defined at code.gs:17488) and `hashCandidateEmail_(email)` (code.gs:17498). Raw identifiers in a shared sheet are a privacy leak.
2. **Lock the read-seen-then-append section** with `LockService.getScriptLock().tryLock(10000)`. The hourly trigger and a UI button can interleave and double-append otherwise.
3. **Cache expensive "unsynced?" checks** in `CacheService.getUserCache()` with a short TTL (15 min). Invalidate after successful sync via a dedicated helper (`invalidateUnsyncedCache_`).
4. **Synthesize a dedup key** when thread ID is missing (pre-instrumentation rows): `sha256('synth:' + user + '|' + candidate|job + '|' + outreachMs)`. Dropping those rows silently hides history.
5. **Backward-compat for hash migrations:** detect legacy raw-identifier rows via `/^[0-9a-f]{64}$/` and hash them on-the-fly into the seen-set so an upgraded sync doesn't re-append.
6. **Earliest-wins dedup at read time** — when `(candidateHash, jobId)` appears for multiple users (job transfer), keep the row with the earliest outreach timestamp; that's the real first-contact.
7. **Never run expensive syncs inside getters** called on every chart render. The writer runs on the hourly trigger; the reader only reads.
8. **RBAC in the reader, not the writer** — each user writes their own rows; the reader filters by `allowedEmails` (null=admin, teamMembers=TL, [self]=IC).

### Helpers (all at code.gs:17483-17518)
- `sha256Hex_(value)` — generic SHA-256 hex digest
- `hashCandidateEmail_(email)` — normalizes then hashes
- `isValidDate_(d)` — guards every `new Date()` parse from `NaN` cascade
- `invalidateUnsyncedCache_(userEmail)` — clears cached backfill-banner state

---

## Coding Guardrails (lessons from QA)

These patterns have bitten us; apply proactively.

- **`filterJobId` must be coerced** — UI may pass numbers. Use `const jobFilter = filterJobId ? String(filterJobId) : '';` then compare. Direct `!==` against string sheet values silently drops all rows.
- **Every `new Date(x)` needs `isValidDate_()`** — bad cells in any sheet produce Invalid Date which cascades into NaN hours, broken `<` comparisons, and skewed metrics.
- **`parseInt(x, 10)` + `Number.isFinite(v) && v > 0 ? v : fallback`** — never trust numeric columns to be numeric. Blank/NaN silently inflates sums.
- **Banner/UI visibility calls belong outside renderers** — call them from the orchestrator (e.g. `loadAnalyticsCharts`), not from inside a chart renderer that only runs on the success path. Otherwise the `else`/error path leaves stale UI.
- **Don't mitigate with retry/fallback logic** — fix the root cause. Adding `try { sync() } catch {}` everywhere masks the design error that sync shouldn't run there in the first place.
- **Dark mode: always use explicit `dark:` Tailwind classes** — Do NOT rely solely on the global `:where(.dark) .class` CSS overrides in the `<style>` block. Tailwind CDN injects its stylesheet last (after the custom `<style>` block), so its plain-class rules win the cascade and silently override any `:where(.dark)` fallback. Every element that needs dark-mode-specific styling must carry an explicit `dark:bg-slate-800`, `dark:text-slate-400`, etc. Dynamically generated HTML (via `innerHTML`/`createElement`) must also include `dark:` classes since the stylesheet injection has already happened.

---

## Commit Process (required)

Before every commit:
1. **Self-read the diff end-to-end** — `git diff HEAD` and read every hunk. Don't rely on what you think you wrote.
2. **Syntax-check** — `cp code.gs /tmp/code.js && node --check /tmp/code.js`. GAS runs ES2019 and Node's parser catches most issues despite the extension mismatch.
3. **Spawn a QA agent for non-trivial changes** — use `Agent` with a prompt that asks for CRITICAL/HIGH/MEDIUM/LOW findings across correctness, concurrency, privacy, performance, and edge cases. Apply all actionable findings before committing.
4. **Verify QA findings aren't false positives** — trace the actual code path (e.g. the user asked to verify that `logAnalytics('email_sent', ...)` is initial-only before "fixing" it as a response-rate bug; it was, and the fix was dropped).
5. **Commit message lists each finding addressed**, grouped by severity. The commit is the audit trail.

---

## Development Notes

- No TypeScript, no bundler, no npm. Pure GAS (ES2019 subset, no modules).
- All changes go in `code.gs` and/or `index.html`.
- Test by deploying a new version in Apps Script editor, or use "Run" for individual functions.
- `debugLog()` wraps `console.log` and can be toggled via `DEBUG_LOGGING` property.
- `ensureSheetsExist(ss)` is called at startup; lazy schema migrations live there.
- `REQUIRED_TRIGGERS` array drives the Config modal's trigger health check.

---

## Branch

Active development branch: `claude/update-docs-dark-mode-fix-8ZnkE`
# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Turing AI Recruiter V12** — a Google Apps Script (GAS) application that automates developer recruitment outreach, AI-powered rate negotiation, follow-up emails, analytics, and a learning system for AI improvement. It integrates Gmail, Google Sheets, BigQuery, and OpenAI.

There is no build step, no package manager, and no test framework. Code runs directly inside Google Apps Script (server-side `.gs` files + an `index.html` frontend).

---

## Repository Structure

| File | Purpose |
|------|---------|
| `code.gs` | ~24 500-line GAS backend — all server-side logic |
| `index.html` | ~17 300-line frontend — Tailwind CSS, jQuery, Chart.js |

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
