# Zapier Integration Guide

This guide wires Google Sheets to Slack and Salesforce with idempotency. Use “New or Updated Spreadsheet Row” where possible. Use filters to prevent duplicate posts and timestamp write-backs as idempotency keys.

## 0) Sheet prep for Zaps
- Make sure your `Project Pipeline` sheet contains:
  - `Slack Channel ID`, `Slack Team ID`, `Slack Channel URL` (auto formula)
  - `last_upcoming_notified_ts`, `last_framing_notified_ts`, `last_block_notified_ts`, `last_escalation_notified_ts`
- These fields are updated by Zaps after posting to Slack.

---

## Zap 1 - Upcoming created → project channel

**Trigger**  
- App: Google Sheets  
- Event: New or Updated Spreadsheet Row  
- Spreadsheet: your Project Pipeline sheet  
- Trigger Column: `ts_permits_approved` (or use “All” with a Filter)  

**Filter**  
- `Permits` = `Approved`  
- `Slack Channel ID` is not empty  
- `ts_added_to_upcoming` is empty OR use your own condition  
- Optional: `OR( ISBLANK(last_upcoming_notified_ts), ts_permits_approved > last_upcoming_notified_ts )`

**Action**  
- App: Slack  
- Event: Send Channel Message  
- Channel: map `Slack Channel ID`  
- Message:
  ```
  :new: Upcoming created — SFID {{SFID}} · {{Name}}
  Deadline: {{Deadline}}
  Sheet: {{Slack Channel URL}}
  Drive: {{Drive Folder URL}}
  ```

**Action**  
- App: Google Sheets  
- Event: Update Spreadsheet Row  
- Set `last_upcoming_notified_ts` = Zap Meta Timestamp

---

## Zap 2 - Framing created → project channel

**Trigger**  
- Google Sheets: New or Updated Row  
- Filter: `Project Status` = `Permitting`, `Slack Channel ID` present  
- Optional idempotency: `OR( ISBLANK(last_framing_notified_ts), NOW() > last_framing_notified_ts )`

**Action**  
- Slack: Send Channel Message  
- Message:
  ```
  :triangular_ruler: Framing created — SFID {{SFID}} · {{Name}}
  Architect: {{Architect}} · Deadline: {{Deadline}}
  ```

**Action**  
- Google Sheets: Update Row -> `last_framing_notified_ts` = Zap Meta Timestamp

---

## Zap 3 - Status blocked → project channel

**Trigger**  
- Google Sheets: New or Updated Row  
- Filter:
  - `Blocked since` is not empty
  - `Slack Channel ID` present
  - `OR( ISBLANK(last_block_notified_ts), Blocked since > last_block_notified_ts )`

**Action**  
- Slack: Send Channel Message  
- Message:
  ```
  :warning: Status change blocked — SFID {{SFID}} · {{Name}}
  Reason(s): {{Advance block reason}}
  Owner: {{Assigned to}}
  ```

**Action**  
- Google Sheets: Update Row -> `last_block_notified_ts` = Zap Meta Timestamp

---

## Zap 4 - 24h escalation

**Trigger**  
- Google Sheets: New or Updated Row  
- Filter:
  - `escalate_ready` = TRUE
  - `Slack Channel ID` present
  - `OR( ISBLANK(last_escalation_notified_ts), Blocked since > last_escalation_notified_ts )`

**Action**  
- Slack: Send Channel Message  
- Message:
  ```
  :rotating_light: Blocked for 24h — SFID {{SFID}} · {{Name}}
  Still failing: {{Advance block reason}}
  Please resolve or request Admin override
  ```

**Action**  
- Google Sheets: Update Row -> `last_escalation_notified_ts` = Zap Meta Timestamp

---

## Zap 5 (optional) - Create Slack channel on new project, write back ID

**Trigger**  
- Google Sheets: New Row  
- Filter: `Slack Channel ID` is empty

**Action**  
- Slack: Create Conversation  
- Use a normalized name pattern like `proj-{{SFID}}-{{Name}}`  
- Invite PMs, post a welcome note

**Action**  
- Google Sheets: Update Row -> write back `Slack Channel ID`

---

## Zap 6 (optional) - DM assignee on key tasks or due soon

**Trigger**  
- Google Sheets: New or Updated Row on `Tasks`  
- Filter: `Status` = `Todo` or `In Progress` and Due soon

**Actions**  
- Slack: Find User by Email  
- Slack: Send DM with context and link back to the sheet row

---

## Zap 7 (optional) - Final payment received → channel and finance

**Trigger**  
- Google Sheets: Updated Row  
- Filter: `Final payment received` = TRUE

**Action**  
- Slack: Send Channel Message to project channel  
- Slack: Optionally message finance channel with revenue summary

---

## Zap 8 (optional) - Daily ops digest at 8am ET

**Trigger**  
- Schedule by Zapier: Every day at 8:00 AM Eastern

**Actions**  
- Google Sheets: Lookups for active backlog, blocked count, ready-to-schedule, weighted revenue this month  
- Slack: Post a formatted summary

---

## Zap 9 - Salesforce Closed-Won → upsert project

**Trigger**  
- Salesforce: New or Updated Opportunity  
- Filter: StageName = `Closed Won`

**Action**  
- Google Sheets: Find Row by `external_id` = Opportunity Id  
  - If not found: Create Row  
  - If found: Update Row

**Mappings**
- Name = Opportunity Name  
- external_id = Opportunity Id  
- Source = Salesforce  
- SFID = your Salesforce “Project SFID” or fallback to Opportunity Id  
- Equipment = from products or custom field  
- Revenue = Amount  
- Probability = 100  
- Project Status = Scheduled (or leave blank for PM)  
- Permits = Not Started  
- opportunity_closed_date = ClosedDate  
- Priority = Standard  
- Slack Channel ID = leave empty  
- Drive Folder URL = leave empty

Idempotency: using Find or Create by `external_id` avoids duplicates.
