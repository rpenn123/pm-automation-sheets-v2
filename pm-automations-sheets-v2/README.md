# Project Pipeline with Dependency Enforcement - Google Sheets Edition

A robust, scalable migration of your Notion-based project pipeline into Google Sheets, with strict phase-gate validation, automatic status reverts, structured error logging, and Slack + Salesforce integrations via Zapier.

This solution has been refactored for improved performance, maintainability, and ease of setup.

## Key Features

- **Automated Phase-Gate Enforcement**: A powerful Apps Script enforces business rules, reverts invalid status changes, adds explanatory notes to cells, and logs issues to a dedicated `Ops Inbox`.
- **Automated Setup**: A custom menu (`Mobility123 PM`) automates the most error-prone parts of the setup, creating required sheets and installing the necessary script trigger.
- **Dynamic & Robust**: The script is resilient to changes in column order and gracefully ignores routine edits from Zapier to prevent unnecessary processing.
- **Configurable Admins**: A simple UI allows you to define which users can access the "override" functionality without needing to edit any code.
- **Comprehensive Views**: `QUERY`-based tabs provide read-only views for `Upcoming` and `Framing` projects, and a powerful dashboard provides high-level KPIs and analytics.
- **External Integrations**: A detailed guide (`docs/zapier_guide.md`) explains how to connect the sheet to Slack and Salesforce.

---

## 1. Setup Instructions

The setup process has been significantly simplified.

### Step 1: Create the Google Sheet & Install the Script
1.  Create a new Google Sheet. You can name it whatever you like (e.g., "Project Pipeline").
2.  Open the script editor by going to **Extensions → Apps Script**.
3.  Copy the entire contents of `src/main.gs` from this repository and paste it into the script editor, replacing any existing code.
4.  Save the script project (File → Save).

### Step 2: Run the Automated Setup
1.  Refresh the Google Sheet in your browser. After a moment, a new menu named **Mobility123 PM** should appear.
2.  Click **Mobility123 PM → Verify & Setup Sheet**.
3.  The script will check for required sheets and install the necessary `onEdit` trigger. It will show you a report of the actions it took.
4.  You will be asked to authorize the script. Please grant it the required permissions.

### Step 3: Populate Sheet Content
1.  The setup script will have created a `Lists` tab. Go to this tab and paste in the canonical lists for your dropdown menus (Project Status, Permits, etc.), with each list in its own column and a header in the first row.
2.  Go to the `Project Pipeline` and `Tasks` tabs and paste in the column headers exactly as defined in the **Column Schemas** section below.
3.  Follow the instructions in `formulas/sheet_formulas.md` and `dashboard/dashboard_queries.md` to paste the required formulas into the sheet.
4.  Set up your data validation and named ranges. For example, select the values under the `Project Status` header in the `Lists` sheet, go to **Data → Named ranges**, and name it `List_ProjectStatus`. Then, in the `Project Pipeline` sheet, select the `Project Status` column, go to **Data → Data validation**, and set the criteria to be "Dropdown (from a range)" with the value `List_ProjectStatus`. Repeat for all dropdowns.

---

## 2. Configuration

### Admin Users
The "Override: Allow Advance" checkbox can only be used by authorized administrators. To configure who is an admin:
1.  Click **Mobility123 PM → Configure Admins**.
2.  Enter a comma-separated list of user email addresses in the prompt.
3.  Click OK. The list is saved securely and can be updated anytime.

---

## 3. Column Schemas

Create columns in each tab exactly as listed here. While the script is resilient to column reordering, the sheet-level formulas are not. Keep the columns in this order.

### 3.1 Project Pipeline (writable)
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

### 3.2 Tasks (subtasks)
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

### 3.3 Upcoming (read-only via QUERY in A1)
Columns are generated by query. No manual entry.

### 3.4 Framing (read-only via QUERY in A1)
Columns are generated by query. No manual entry.

### 3.5 Ops Inbox
```
A: Name
B: Source SFID
C: Type (select: automation_error, duplicate_sfid, data_missing)
D: Resolved (checkbox)
E: Details
F: Timestamp (datetime)
```

---

## 4. Zapier Integrations

Open `docs/zapier_guide.md`. It covers the detailed steps for connecting Google Sheets to Slack and Salesforce.

---

## 5. Governance

- Edit only in `Project Pipeline` and `Tasks`.
- `Upcoming` and `Framing` are read-only.
- The admin override window is 24 hours. The script manages this automatically.
- Keep dropdown option strings identical to the canonical options on the `Lists` sheet.

---

## 6. Test Plan

- Try to move a row to `Scheduled` without permits approved or missing artifacts. The script should revert the status and add a note to the cell.
- Mark `Permits` to `Approved`. The script should stamp `ts_permits_approved`.
- Move to `Permitting`. The script should stamp `ts_entered_permitting`.
- Use the "Override: Allow Advance" checkbox as a non-admin (it should fail) and then as an admin (it should work).
- Check that payment guards and duplicate SFID detection are working.
- Confirm Zapier messages only post once per event per row.
