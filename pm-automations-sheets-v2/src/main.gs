/**
 * Project Pipeline with Dependency Enforcement - Google Sheets Edition
 * Apps Script
 */

const CFG = {
  TZ: Session.getScriptTimeZone() || 'America/New_York',

  SHEETS: {
    PIPELINE: 'Project Pipeline',
    TASKS: 'Tasks',
    UPCOMING: 'Upcoming',
    FRAMING: 'Framing',
    OPS_INBOX: 'Ops Inbox',
  },

  PIPELINE_HEADERS: [
    'Name','SFID','Slack Channel ID','external_id','import_batch_id','Drive Folder URL','Slack Team ID',
    'Slack Channel URL','Project Status','Permits','Priority','Probability','pipeline_last_transfer_status',
    'Inspection Performed By','Status (Last Valid)','Source','Assigned to','Deadline','ts_permits_approved',
    'ts_entered_permitting','ts_added_to_upcoming','ts_added_to_framing','pipeline_last_transfer_ts',
    'last_validated_ts','Deposit received','Final payment received','Final payment date','Deposit received date',
    'Blocked since','Override until','ts_first_scheduled','ts_marked_done','opportunity_closed_date',
    'Permit application submitted','Permit artifacts in Drive','Change orders approved',
    'Equipment received in warehouse','Site prep checklist complete','Rough inspections passed',
    'Final inspection passed','Duplicate SFID','Override: Allow Advance','Equipment','Architect','Revenue','COGS',
    'last_upcoming_notified_ts','last_framing_notified_ts','last_block_notified_ts','last_escalation_notified_ts',
    'Gross Margin %','Week of','open_tasks_count','overdue_tasks_count','completed_tasks_count',
    'total_blocking_tasks','completed_blocking_tasks','task_progress_%','can_advance_globally',
    'can_advance_to_Permitting','can_advance_to_Scheduled','can_advance_to_Inspections','can_advance_to_Done',
    'Advance block reason','last_edit_relative','escalate_ready','Month (Deadline)','Created Month',
    'days_in_permitting','days_to_schedule','lead_time_days','Revenue Weighted','docs_required_but_missing',
    'aging_days_since_edit','is_active_backlog','blocked_hours','staleness_flag'
  ],

  TASK_HEADERS: [
    'Name','Project SFID','Phase','Status','Type','Assigned to','Due Date','Completed Date',
    'Effort hours','Depends on','Counts toward completion','Completed %'
  ],

  OPS_HEADERS: ['Name','Source SFID','Type','Resolved','Details','Timestamp'],

  STATUS_VALUES: ['Scheduled','Permitting','Done','Canceled','On Hold','Stuck','Inspections','Overdue'],

  OVERRIDE_ADMINS: ['Amber','Ryan','Cassidy'],
};

function onEditHandler(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const sheetName = sh.getName();
    if (sheetName === CFG.SHEETS.PIPELINE) return handlePipelineEdit_(e);
    if (sheetName === CFG.SHEETS.TASKS) return handleTasksEdit_(e);
  } catch (err) {
    logOpsInbox_('script_error', '', `Unhandled error: ${err && err.message ? err.message : err}`);
  }
}

function handlePipelineEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 1) return;

  const headerMap = getHeaderMap_(sh, CFG.PIPELINE_HEADERS);
  if (!headerMap) return;

  const rowVals = sh.getRange(row, 1, 1, headerMap._width).getValues()[0];
  const colIdx = (h) => headerMap[h];

  if (col === colIdx('Permits')) {
    const permits = rowVals[colIdx('Permits') - 1];
    if (permits === 'Approved') stampIfEmpty_(sh, row, 'ts_permits_approved', headerMap);
  }

  if (col === colIdx('Project Status')) enforceStatusChange_(e, sh, row, rowVals, headerMap);

  const status = rowVals[colIdx('Project Status') - 1];
  if (status === 'Permitting') stampIfEmpty_(sh, row, 'ts_entered_permitting', headerMap);

  if (col === colIdx('SFID')) detectDuplicateSFID_(sh, row, headerMap);

  if (col === colIdx('Override: Allow Advance')) {
    const overrideChecked = !!rowVals[colIdx('Override: Allow Advance') - 1];
    if (overrideChecked) setOverrideWindow_(sh, row, headerMap);
  }

  if (col === colIdx('Final payment received') || col === colIdx('Project Status')) {
    paymentGuard_(e, sh, row, rowVals, headerMap);
  }

  sh.getRange(row, colIdx('last_validated_ts')).setValue(new Date());
}

function enforceStatusChange_(e, sh, row, rowVals, headerMap) {
  const colIdx = (h) => headerMap[h];
  const statusNew = rowVals[colIdx('Project Status') - 1];
  const statusLastValid = rowVals[colIdx('Status (Last Valid)') - 1] || e.oldValue || '';
  const sfid = rowVals[colIdx('SFID') - 1] || '';

  const overrideChecked = !!rowVals[colIdx('Override: Allow Advance') - 1];
  const overrideUntil = rowVals[colIdx('Override until') - 1];
  if (overrideChecked && isOverrideActive_(overrideUntil)) {
    sh.getRange(row, colIdx('Status (Last Valid)')).setValue(statusNew);
    clearBlocked_(sh, row, headerMap);
    return;
    }

  const canGlobal = toBool_(rowVals[colIdx('can_advance_globally') - 1]);
  const canPerm = toBool_(rowVals[colIdx('can_advance_to_Permitting') - 1]);
  const canSched = toBool_(rowVals[colIdx('can_advance_to_Scheduled') - 1]);
  const canInspect = toBool_(rowVals[colIdx('can_advance_to_Inspections') - 1]);
  const canDone = toBool_(rowVals[colIdx('can_advance_to_Done') - 1]);

  let valid = true;
  if (!canGlobal) valid = false;
  if (statusNew === 'Permitting' && !canPerm) valid = false;
  if (statusNew === 'Scheduled' && !canSched) valid = false;
  if (statusNew === 'Inspections' && !canInspect) valid = false;
  if (statusNew === 'Done' && !canDone) valid = false;

  if (!valid) {
    const revertTo = statusLastValid || e.oldValue || '';
    if (revertTo) sh.getRange(row, colIdx('Project Status')).setValue(revertTo);
    const reason = GET_ADVANCE_BLOCK_REASON_(statusNew, canGlobal, canPerm, canSched, canInspect, canDone);
    sh.getRange(row, colIdx('Project Status')).setNote(`Advance blocked. Reasons: ${reason}`);
    if (!sh.getRange(row, colIdx('Blocked since')).getValue()) {
      sh.getRange(row, colIdx('Blocked since')).setValue(new Date());
    }
    logOpsInbox_('data_missing', sfid, `Blocked changing status to ${statusNew}. ${reason}`);
    return;
  }

  sh.getRange(row, colIdx('Status (Last Valid)')).setValue(statusNew);
  clearBlocked_(sh, row, headerMap);

  if (statusNew === 'Scheduled') stampIfEmpty_(sh, row, 'ts_first_scheduled', headerMap);
  if (statusNew === 'Done') stampIfEmpty_(sh, row, 'ts_marked_done', headerMap);

  if (overrideChecked) {
    sh.getRange(row, colIdx('Override: Allow Advance')).setValue(false);
    sh.getRange(row, colIdx('Override until')).clearContent();
  }
}

function paymentGuard_(e, sh, row, rowVals, headerMap) {
  const colIdx = (h) => headerMap[h];
  const status = rowVals[colIdx('Project Status') - 1];
  const paid = toBool_(rowVals[colIdx('Final payment received') - 1]);
  if (status === 'Done' && !paid) {
    const lastValid = rowVals[colIdx('Status (Last Valid)') - 1] || e.oldValue || '';
    if (lastValid) sh.getRange(row, colIdx('Project Status')).setValue(lastValid);
    sh.getRange(row, colIdx('Project Status')).setNote('Payment must precede Done.');
    const sfid = rowVals[colIdx('SFID') - 1] || '';
    logOpsInbox_('data_missing', sfid, 'Attempted to set Done without Final payment received.');
  }
}

function detectDuplicateSFID_(pipelineSheet, row, headerMap) {
  const colIdx = (h) => headerMap[h];
  const sfid = pipelineSheet.getRange(row, colIdx('SFID')).getValue();
  if (!sfid) return;
  const lastRow = pipelineSheet.getLastRow();
  if (lastRow < 2) return;
  const sfids = pipelineSheet.getRange(2, colIdx('SFID'), lastRow - 1, 1).getValues().flat();
  const dupCount = sfids.filter(v => v === sfid).length;
  const isDup = dupCount > 1;
  pipelineSheet.getRange(row, colIdx('Duplicate SFID')).setValue(isDup);
  if (isDup) logOpsInbox_('duplicate_sfid', sfid, `Detected duplicate SFID: ${sfid}`);
}

function setOverrideWindow_(sh, row, headerMap) {
  const colIdx = (h) => headerMap[h];
  const until = new Date(Date.now() + 24 * 60 * 60 * 1000);
  sh.getRange(row, colIdx('Override until')).setValue(until);
}

function clearBlocked_(sh, row, headerMap) {
  const colIdx = (h) => headerMap[h];
  sh.getRange(row, colIdx('Blocked since')).clearContent();
}

function handleTasksEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  if (row === 1) return;
  const headerMap = getHeaderMap_(sh, CFG.TASK_HEADERS);
  if (!headerMap) return;
  const colIdx = (h) => headerMap[h];
  if (e.range.getColumn() === colIdx('Status')) {
    const status = sh.getRange(row, colIdx('Status')).getValue();
    const completedDate = sh.getRange(row, colIdx('Completed Date')).getValue();
    if (status === 'Done' && !completedDate) sh.getRange(row, colIdx('Completed Date')).setValue(new Date());
  }
}

function logOpsInbox_(type, sfid, details) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.SHEETS.OPS_INBOX);
  if (!sh) return;
  const headers = getHeaderMap_(sh, CFG.OPS_HEADERS);
  const row = [
    sfid ? `SFID ${sfid} ${type}` : `pipeline ${type}`,
    sfid || '',
    type,
    false,
    details || '',
    new Date()
  ];
  sh.appendRow(row);
}

function GET_ADVANCE_BLOCK_REASON(status, canGlobal, canPerm, canSched, canInspect, canDone) {
  return GET_ADVANCE_BLOCK_REASON_(status, canGlobal, canPerm, canSched, canInspect, canDone);
}

function GET_ADVANCE_BLOCK_REASON_(status, canGlobal, canPerm, canSched, canInspect, canDone) {
  const reasons = [];
  if (!toBool_(canGlobal)) reasons.push('Missing global data or overdue tasks or duplicate SFID');
  if (status === 'Permitting' && !toBool_(canPerm)) reasons.push('Gate to Permitting not met: deposit, permit app, or tasks incomplete');
  if (status === 'Scheduled' && !toBool_(canSched)) reasons.push('Gate to Scheduled not met: permits approved, artifacts in Drive, revenue+probability, or tasks incomplete');
  if (status === 'Inspections' && !toBool_(canInspect)) reasons.push('Gate to Inspections not met: equipment, site prep, or rough inspections');
  if (status === 'Done' && !toBool_(canDone)) reasons.push('Gate to Done not met: final inspection, payment, artifacts in Drive, or tasks incomplete');
  return reasons.join(' • ');
}

function getHeaderMap_(sheet, expectedHeaders) {
  const width = expectedHeaders.length;
  const row = sheet.getRange(1, 1, 1, width).getValues()[0];
  const map = {};
  for (let i = 0; i < expectedHeaders.length; i++) {
    if ((row[i] || '').toString().trim() !== expectedHeaders[i]) return null;
    map[expectedHeaders[i]] = i + 1;
  }
  map._width = width;
  return map;
}

function stampIfEmpty_(sh, row, headerName, headerMap) {
  const col = headerMap[headerName];
  if (!col) return;
  const cell = sh.getRange(row, col);
  if (!cell.getValue()) cell.setValue(new Date());
}

function isOverrideActive_(overrideUntil) {
  if (!overrideUntil) return false;
  try {
    const until = new Date(overrideUntil);
    return Date.now() <= until.getTime();
  } catch (e) { return false; }
}

function toBool_(v) {
  if (v === true || v === false) return !!v;
  if (typeof v === 'string') return v.toLowerCase() === 'true';
  if (typeof v === 'number') return v !== 0;
  return !!v;
}
