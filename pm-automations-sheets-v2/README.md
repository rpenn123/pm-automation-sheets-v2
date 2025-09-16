# Project Pipeline with Dependency Enforcement - Google Sheets Edition

A robust, scalable migration of your Notion-based project pipeline into Google Sheets, with strict phase-gate validation, automatic status reverts, structured error logging, and Slack + Salesforce integrations via Zapier.

This repo gives you:
- A single Google Sheet with 5 tabs: `Project Pipeline`, `Tasks`, `Upcoming`, `Framing`, `Ops Inbox`.
- Data validation for every select and multi-select field using canonical lists.
- Rollups and analytics via formulas to replace Notion relations and rollups.
- A powerful `onEdit` Apps Script that enforces gates, reverts invalid status changes, annotates the cell with a block reason, logs the event to `Ops Inbox`, and stamps first-time transitions.
- Read-only views for `Upcoming` and `Framing` populated by `QUERY` formulas from the pipeline.
- An Executive & Ops Dashboard built from `QUERY`, `COUNTIFS`, `SUMIFS`, `SPARKLINE`, and native charts.
- Zapier workflows to post Slack updates and sync from Salesforce Closed-Won with idempotency.

> Keep option strings identical across tabs and Zaps. Changing an option requires updating data validation and any filter/logic that references it.

---

## 1) Make the Google Sheet

1) Create a new Google Sheet named **Project Pipeline**.

2) Create these tabs (exact names):
   - `Project Pipeline`
   - `Tasks`
   - `Upcoming`
   - `Framing`
   - `Ops Inbox`
   - `Lists` (helper tab for dropdown options)

3) On `Lists`, paste the canonical lists into separate columns with these headers in row 1:
   - **Project Status**: `Scheduled, Permitting, Done, Canceled, On Hold, Stuck, Inspections, Overdue`
   - **Permits**: `Not Started, In Progress, Approved, N/A, Resubmission`
   - **Task Status**: `Todo, In Progress, Blocked, Waiting External, Skipped, Done`
   - **Task Phase**: `P1 Kickoff, P2 Plans & Permits, P3 Construction, P4 Installation, P5 Closeout`
   - **Probability**: `100, 90, 75, 50, 25`
   - **Inspection Performed By**: `By Mobility123, DCA, 3rd party`
   - **Equipment** (multi-select): `Duo Alta, Trio Alta, Multilift, V-1504, Eclipse, Bathroom, Trio Thru, Duo Thru, Trio Classic, Telecab, Stairlift, VPL-3200, Vuelift, Trio Alta Plus, Symmetry - SHE, Cibes A5000, Ramp, Symmetry - VPL, IPL, Cibes Primo`
   - **Architect**: `Sidrane, KVD, Ingrid, Mclaughlin, CBO, Other / Unknown`

4) Create **named ranges** for each list. Example: select the column of values under `Project Status` and set Named range = `List_ProjectStatus`. Do the same for all lists.

---

## 2) Column schemas

Create columns in each tab exactly as listed here and keep this order. Do not insert extra columns in the middle later. Use **Data validation** to bind each select to the proper named range. Multi-selects in Sheets are entered as comma-separated text.

### 2.1 Project Pipeline (writable)
```
A: Name
B: SFID
C: Slack Channel ID
D: external_id
E: import_batch_id
F: Drive Folder URL
G: Slack Team ID
H: Slack Channel URL (formula)
I: Project Status (select)
J: Permits (select)
K: Priority (select: Standard, High)
L: Probability (select from List_Probability)
M: pipeline_last_transfer_status (select: pending, success, skipped-duplicate, error)
N: Inspection Performed By (select)
O: Status (Last Valid) (select - same options as Project Status)
P: Source (select: Notion, Salesforce)
Q: Assigned to
R: Deadline (date)
S: ts_permits_approved (datetime)
T: ts_entered_permitting (datetime)
U: ts_added_to_upcoming (datetime)
V: ts_added_to_framing (datetime)
W: pipeline_last_transfer_ts (datetime)
X: last_validated_ts (datetime)
Y: Deposit received (checkbox)
Z: Final payment received (checkbox)
AA: Final payment date (date)
AB: Deposit received date (date)
AC: Blocked since (datetime)
AD: Override until (datetime)
AE: ts_first_scheduled (datetime)
AF: ts_marked_done (datetime)
AG: opportunity_closed_date (date)
AH: Permit application submitted (checkbox)
AI: Permit artifacts in Drive (checkbox)
AJ: Change orders approved (checkbox)
AK: Equipment received in warehouse (checkbox)
AL: Site prep checklist complete (checkbox)
AM: Rough inspections passed (checkbox)
AN: Final inspection passed (checkbox)
AO: Duplicate SFID (checkbox)
AP: Override: Allow Advance (checkbox)
AQ: Equipment (multi-select text, comma-separated)
AR: Architect (select)
AS: Revenue (number)
AT: COGS (number)
AU: last_upcoming_notified_ts (datetime)
AV: last_framing_notified_ts (datetime)
AW: last_block_notified_ts (datetime)
AX: last_escalation_notified_ts (datetime)
AY: Gross Margin % (formula)
AZ: Week of (Monday) (formula)
BA: open_tasks_count (formula)
BB: overdue_tasks_count (formula)
BC: completed_tasks_count (formula)
BD: total_blocking_tasks (formula)
BE: completed_blocking_tasks (formula)
BF: task_progress_% (formula)
BG: can_advance_globally (formula)
BH: can_advance_to_Permitting (formula)
BI: can_advance_to_Scheduled (formula)
BJ: can_advance_to_Inspections (formula)
BK: can_advance_to_Done (formula)
BL: Advance block reason (custom function)
BM: last_edit_relative (optional formula - informational)
BN: escalate_ready (formula)
BO: Month (Deadline) (formula)
BP: Created Month (formula)
BQ: days_in_permitting (formula)
BR: days_to_schedule (formula)
BS: lead_time_days (formula)
BT: Revenue Weighted (formula)
BU: docs_required_but_missing (formula)
BV: aging_days_since_edit (formula)
BW: is_active_backlog (formula)
BX: blocked_hours (optional formula)
BY: staleness_flag (optional formula)
```

### 2.2 Tasks (subtasks)
```
A: Name
B: Project SFID
C: Phase (select from List_TaskPhase)
D: Status (select from List_TaskStatus)
E: Type (select: Permitting, Framing, Install, Inspection, Procurement, Site Prep, QA, Admin, Other)
F: Assigned to
G: Due Date (date)
H: Completed Date (date)
I: Effort hours (number)
J: Depends on (free text or SFID of prerequisite task)
K: Counts toward completion (formula)
L: Completed % (formula)
```

### 2.3 Upcoming (read-only via QUERY in A1)
Columns are generated by query. No manual entry.

### 2.4 Framing (read-only via QUERY in A1)
Columns are generated by query. No manual entry.

### 2.5 Ops Inbox
```
A: Name
B: Source SFID
C: Type (select: automation_error, duplicate_sfid, data_missing)
D: Resolved (checkbox)
E: Details
F: Timestamp (datetime)
```

---

## 3) Add formulas

Open `formulas/sheet_formulas.md` from this repo and paste the formulas into the correct columns. That file is organized by sheet and column. Do not modify formula columns by hand later.

---

## 4) Protect critical ranges

In Google Sheets: Data → Protect sheets and ranges.
- Protect all formula columns in `Project Pipeline` and `Tasks`.
- Protect entire `Upcoming` and `Framing` tabs since they are view-only.
- Optionally protect `Ops Inbox` except for Ops leads.

---

## 5) Install the Apps Script

1) Extensions → Apps Script.
2) Replace all code with the contents of `src/main.gs`.
3) Save. File → Project properties → set script timezone to your business timezone.
4) Triggers (alarm-clock icon on left) → Add Trigger:
   - Function: `onEditHandler`
   - Deployment: Head
   - Event source: From spreadsheet
   - Event type: On edit
5) Review authorization and approve.

> Why installable onEdit: it gives you more consistent `e.oldValue` for single cell edits and avoids simple-trigger limits.

---

## 6) Build read-only views

Open `dashboard/dashboard_queries.md` and paste the `QUERY` for `Upcoming` into `Upcoming!A1`, and the `QUERY` for `Framing` into `Framing!A1`. These are live views sourced from `Project Pipeline`.

---

## 7) Create the Dashboard

Create a new tab named `Dashboard`. Paste the KPI and table queries from `dashboard/dashboard_queries.md`. Add native charts as noted in that file.

---

## 8) Zapier integrations

Open `docs/zapier_guide.md`. It covers:
- Google Sheets triggers using “New or Updated Spreadsheet Row” with a Last Updated helper.
- Slack posting with idempotency using `last_*_notified_ts`.
- Salesforce Closed-Won sync with “Find or Create by external_id”.
- Escalation and blocked notifications.

Follow every step and test each Zap with a sample row.

---

## 9) Governance

- Edit only in `Project Pipeline` and `Tasks`.
- `Upcoming` and `Framing` are read-only.
- Admin-only override window is 24 hours. The script stamps and clears it as required.
- Keep strings identical to the canonical options on `Lists`. If you must change options, update data validation and any formula that references that option.

---

## 10) Test plan

- Try to move a row to `Scheduled` without permits approved or missing artifacts. The script should:
  - Revert the status.
  - Note the cell with the reason.
  - Append an issue to `Ops Inbox`.
- Mark `Permits` to `Approved`. The script stamps `ts_permits_approved` if empty.
- Move to `Permitting`. The script stamps `ts_entered_permitting` if empty.
- First valid `Scheduled` sets `ts_first_scheduled`. First valid `Done` sets `ts_marked_done`.
- Payment guard. If `Final payment received` is unchecked, attempting `Done` should revert.
- Duplicate SFID. Typing a duplicate SFID flags `Duplicate SFID` and logs a duplicate issue.
- Zapier. Confirm Slack messages only post once per event per row.
