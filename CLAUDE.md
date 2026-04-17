# HR-ops-AI — Claude Code Reference

## Project Structure

Single-file Google Apps Script project:
- `code.gs` — entire backend (~17 000 lines)
- `index.html` — frontend SPA served as a web app

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

## Development Notes

- No TypeScript, no bundler, no npm. Pure GAS (ES2019 subset, no modules).
- All changes go in `code.gs` and/or `index.html`.
- Test by deploying a new version in Apps Script editor, or use "Run" for individual functions.
- `debugLog()` wraps `console.log` and can be toggled via `DEBUG_LOGGING` property.
- `ensureSheetsExist(ss)` is called at startup; lazy schema migrations live there.
- `REQUIRED_TRIGGERS` array drives the Config modal's trigger health check.

---

## Branch

Active development branch: `claude/fix-candidate-completion-duplicates-UNc1f`
