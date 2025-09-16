/**
 * Project Pipeline with Dependency Enforcement - Google Sheets Edition
 * Apps Script
 *
 * Refactored for robustness and maintainability.
 */

// --- Configuration --- //
const CFG = {
  TZ: Session.getScriptTimeZone() || 'America/New_York',

  SHEETS: {
    PIPELINE: 'Project Pipeline',
    TASKS: 'Tasks',
    UPCOMING: 'Upcoming',
    FRAMING: 'Framing',
    OPS_INBOX: 'Ops Inbox',
  },

  // Headers the script REQUIRES to be present. The script is resilient to column order or new columns being added.
  REQUIRED_PIPELINE_HEADERS: [
    'Project Status', 'Permits', 'SFID', 'Override: Allow Advance', 'Final payment received', 'Status (Last Valid)',
    'can_advance_globally', 'can_advance_to_Permitting', 'can_advance_to_Scheduled', 'can_advance_to_Inspections',
    'can_advance_to_Done', 'Blocked since', 'Override until', 'ts_permits_approved', 'ts_entered_permitting',
    'ts_first_scheduled', 'ts_marked_done', 'Duplicate SFID', 'last_validated_ts'
  ],

  REQUIRED_TASK_HEADERS: ['Status', 'Completed Date'],

  REQUIRED_OPS_HEADERS: ['Name', 'Source SFID', 'Type', 'Resolved', 'Details', 'Timestamp'],

  STATUS_VALUES: ['Scheduled', 'Permitting', 'Done', 'Canceled', 'On Hold', 'Stuck', 'Inspections', 'Overdue'],

  // Admin management is now handled via Script Properties for security and ease of management.
  // See getAdmins() and setAdmins(). A UI will be added later to manage this list.

  // Columns that are written to by Zapier and should not trigger validation logic.
  ZAPIER_ONLY_COLUMNS: [
    'last_upcoming_notified_ts',
    'last_framing_notified_ts',
    'last_block_notified_ts',
    'last_escalation_notified_ts',
    'pipeline_last_transfer_ts',
    'pipeline_last_transfer_status'
  ],
};

// --- Menu & Setup --- //

/**
 * Runs when the spreadsheet is opened. Adds a custom menu to the UI.
 * @param {Object} e The event object for the onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Mobility123 PM')
    .addItem('Verify & Setup Sheet', 'verifyAndSetupSheet')
    .addSeparator()
    .addItem('Configure Admins', 'showAdminConfigUi')
    .addToUi();
}

/**
 * Provides a simple UI prompt to configure the list of admin users.
 */
function showAdminConfigUi() {
  const ui = SpreadsheetApp.getUi();
  const currentAdmins = getAdmins_().join(', ');

  const result = ui.prompt(
    'Configure Admins',
    'Enter a comma-separated list of admin email addresses. These users will be able to use the "Override: Allow Advance" checkbox.',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const newAdminsText = result.getResponseText();
    const newAdmins = newAdminsText.split(',').map(s => s.trim()).filter(Boolean);
    setAdmins_(newAdmins);
    ui.alert('Success', `Admins updated to: ${newAdmins.join(', ')}`, ui.ButtonSet.OK);
  }
}

/**
 * Verifies the spreadsheet has the necessary components (sheets, triggers) and creates them if missing.
 * This function automates the most critical and error-prone parts of the setup.
 */
function verifyAndSetupSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = Object.values(CFG.SHEETS);
  const existingSheets = ss.getSheets().map(s => s.getName());
  const report = ['Setup & Verification Report:'];

  // 1. Check for missing sheets
  let createdSheets = [];
  requiredSheets.forEach(sheetName => {
    if (!existingSheets.includes(sheetName)) {
      ss.insertSheet(sheetName);
      createdSheets.push(sheetName);
    }
  });
  if (createdSheets.length > 0) {
    report.push(`✅ Created missing sheets: ${createdSheets.join(', ')}.`);
  } else {
    report.push('✅ All required sheets are present.');
  }

  // 2. Verify the onEdit trigger is installed
  const triggers = ScriptApp.getProjectTriggers();
  const hasOnEditTrigger = triggers.some(t =>
    t.getEventType() === ScriptApp.EventType.ON_EDIT &&
    t.getHandlerFunction() === 'onEditHandler'
  );

  if (!hasOnEditTrigger) {
    ScriptApp.newTrigger('onEditHandler')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
    report.push('✅ Installed the required onEdit trigger.');
  } else {
    report.push('✅ The onEdit trigger is correctly installed.');
  }

  // 3. Remind user to configure other items
  report.push('\nNext Steps:');
  report.push('1. Populate the `Lists` sheet with your dropdown options.');
  report.push('2. Copy the column headers and formulas from the documentation into the appropriate sheets.');
  report.push('3. Use "Data -> Named ranges" and "Data -> Data validation" to hook up the dropdowns.');
  report.push('4. Use "Configure Admins" menu to set who can use the override feature.');

  ui.alert(report.join('\n'));
}


// --- Main Trigger --- //

/**
 * Main function that runs on any edit in the spreadsheet.
 * Configured as an installable trigger.
 * @param {Object} e The event object from the `onEdit` trigger.
 */
function onEditHandler(e) {
  try {
    if (!e || !e.range) return; // Exit if no event object or range
    const sh = e.range.getSheet();
    const sheetName = sh.getName();

    // Route edit events to the appropriate handler
    if (sheetName === CFG.SHEETS.PIPELINE) {
      return handlePipelineEdit_(e);
    }
    if (sheetName === CFG.SHEETS.TASKS) {
      return handleTasksEdit_(e);
    }
  } catch (err) {
    // Log any unhandled errors to the Ops Inbox for visibility
    logOpsInbox_('script_error', '', `Unhandled error: ${err && err.message ? err.message : err} Stack: ${err.stack}`);
  }
}

// --- Edit Handlers --- //

/**
 * Handles edits on the main 'Project Pipeline' sheet using a batch-processing model.
 * @param {Object} e The event object from the `onEdit` trigger.
 */
function handlePipelineEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 1) return; // Ignore header row edits

  const headerMap = getHeaderMap_(sh, CFG.REQUIRED_PIPELINE_HEADERS);
  if (!headerMap) return;

  // --- BATCH READ ---
  // Read all values and notes for the row at once.
  const range = sh.getRange(row, 1, 1, headerMap._width);
  const initialValues = range.getValues()[0];
  const initialNotes = range.getNotes()[0];

  // Create a state object to hold changes in memory.
  const state = new RowState(headerMap, initialValues, initialNotes);

  // --- PROCESS IN MEMORY ---
  // Pass the state object to all logic functions. They will modify the state in memory.
  const colName = state.getHeader(col);

  // --- Zapier Guard Clause ---
  // Exit gracefully if the edit is on a column that only Zapier should modify.
  // This prevents unnecessary validation runs and potential race conditions.
  if (CFG.ZAPIER_ONLY_COLUMNS.includes(colName)) {
    return;
  }

  if (colName === 'Permits') {
    if (state.getValue('Permits') === 'Approved') {
      stampIfEmpty_(state, 'ts_permits_approved');
    }
  }

  if (colName === 'Project Status') {
    enforceStatusChange_(e, state);
  }

  if (state.getValue('Project Status') === 'Permitting') {
    stampIfEmpty_(state, 'ts_entered_permitting');
  }

  if (colName === 'SFID') {
    detectDuplicateSFID_(sh, row, state);
  }

  if (colName === 'Override: Allow Advance') {
    handleOverrideToggle_(e, state);
  }

  if (colName === 'Final payment received' || colName === 'Project Status') {
    paymentGuard_(e, state);
  }

  // Always stamp the last validation time
  state.setValue('last_validated_ts', new Date());

  // --- BATCH WRITE ---
  // Commit the final state back to the sheet, only writing if changes were made.
  commitRowState_(range, state);
}


/**
 * Handles the logic for the 'Override: Allow Advance' checkbox.
 * Checks if the user is an admin before allowing the override.
 * @param {Object} e The event object.
 * @param {RowState} state The state object for the current row.
 */
function handleOverrideToggle_(e, state) {
  if (state.getValue('Override: Allow Advance')) {
    const userEmail = Session.getActiveUser().getEmail();
    const admins = getAdmins_();

    if (userEmail && admins.includes(userEmail)) {
      setOverrideWindow_(state);
    } else {
      state.revertValue(e, 'Override: Allow Advance');
      state.setNote('Project Status', `Override is only available for admins: ${admins.join(', ')}`);
    }
  }
}

/**
 * Enforces the rules for changing a project's status.
 * @param {Object} e The event object.
 * @param {RowState} state The state object for the current row.
 */
function enforceStatusChange_(e, state) {
  const statusNew = state.getValue('Project Status');
  const sfid = state.getValue('SFID');

  if (state.getValue('Override: Allow Advance') && isOverrideActive_(state.getValue('Override until'))) {
    state.setValue('Status (Last Valid)', statusNew);
    clearBlocked_(state);
    return;
  }

  const canGlobal = toBool_(state.getValue('can_advance_globally'));
  const canPerm = toBool_(state.getValue('can_advance_to_Permitting'));
  const canSched = toBool_(state.getValue('can_advance_to_Scheduled'));
  const canInspect = toBool_(state.getValue('can_advance_to_Inspections'));
  const canDone = toBool_(state.getValue('can_advance_to_Done'));

  let valid = true;
  if (!canGlobal) valid = false;
  if (statusNew === 'Permitting' && !canPerm) valid = false;
  if (statusNew === 'Scheduled' && !canSched) valid = false;
  if (statusNew === 'Inspections' && !canInspect) valid = false;
  if (statusNew === 'Done' && !canDone) valid = false;

  if (!valid) {
    state.revertValue(e, 'Project Status');
    const reason = GET_ADVANCE_BLOCK_REASON(statusNew, canGlobal, canPerm, canSched, canInspect, canDone);
    state.setNote('Project Status', `Advance blocked. Reasons: ${reason}`);
    if (!state.getValue('Blocked since')) {
      state.setValue('Blocked since', new Date());
    }
    logOpsInbox_('data_missing', sfid, `Blocked changing status to ${statusNew}. ${reason}`);
    return;
  }

  state.setValue('Status (Last Valid)', statusNew);
  clearBlocked_(state);

  if (statusNew === 'Scheduled') stampIfEmpty_(state, 'ts_first_scheduled');
  if (statusNew === 'Done') stampIfEmpty_(state, 'ts_marked_done');

  if (state.getValue('Override: Allow Advance')) {
    state.setValue('Override: Allow Advance', false);
    state.setValue('Override until', '');
  }
}

/**
 * Prevents setting status to 'Done' if final payment is not received.
 * @param {Object} e The event object.
 * @param {RowState} state The state object for the current row.
 */
function paymentGuard_(e, state) {
  const status = state.getValue('Project Status');
  const paid = toBool_(state.getValue('Final payment received'));
  if (status === 'Done' && !paid) {
    state.revertValue(e, 'Project Status');
    state.setNote('Project Status', 'Payment must precede Done.');
    const sfid = state.getValue('SFID');
    logOpsInbox_('data_missing', sfid, 'Attempted to set Done without Final payment received.');
  }
}

/**
 * Detects if an SFID is a duplicate and updates the row state.
 * @param {Sheet} pipelineSheet The sheet object for the pipeline.
 * @param {number} row The current row number.
 * @param {RowState} state The state object for the current row.
 */
function detectDuplicateSFID_(pipelineSheet, row, state) {
  const sfid = state.getValue('SFID');
  if (!sfid) return;
  const lastRow = pipelineSheet.getLastRow();
  if (lastRow < 2) return;

  // This is one operation that still needs to read from the sheet.
  const sfidCol = state.getCol('SFID');
  const sfids = pipelineSheet.getRange(2, sfidCol, lastRow - 1, 1).getValues().flat();

  const dupCount = sfids.filter(v => v === sfid).length;
  const isDup = dupCount > 1;
  state.setValue('Duplicate SFID', isDup);
  if (isDup) logOpsInbox_('duplicate_sfid', sfid, `Detected duplicate SFID: ${sfid}`);
}

/**
 * Sets the override window in the row state.
 * @param {RowState} state The state object for the current row.
 */
function setOverrideWindow_(state) {
  const until = new Date(Date.now() + 24 * 60 * 60 * 1000);
  state.setValue('Override until', until);
}

/**
 * Clears the 'Blocked since' field in the row state.
 * @param {RowState} state The state object for the current row.
 */
function clearBlocked_(state) {
  state.setValue('Blocked since', '');
  // Also clear the note on the status column if it was previously blocked
  if ((state.getNote('Project Status') || '').startsWith('Advance blocked')) {
     state.setNote('Project Status', '');
  }
}

/**
 * Handles edits on the 'Tasks' sheet using a batch-processing model.
 * @param {Object} e The event object from the `onEdit` trigger.
 */
function handleTasksEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  if (row === 1) return;
  const headerMap = getHeaderMap_(sh, CFG.REQUIRED_TASK_HEADERS);
  if (!headerMap) return;

  const range = sh.getRange(row, 1, 1, headerMap._width);
  const state = new RowState(headerMap, range.getValues()[0], range.getNotes()[0]);

  if (state.getHeader(e.range.getColumn()) === 'Status') {
    if (state.getValue('Status') === 'Done' && !state.getValue('Completed Date')) {
      state.setValue('Completed Date', new Date());
    }
  }

  commitRowState_(range, state);
}

function logOpsInbox_(type, sfid, details) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CFG.SHEETS.OPS_INBOX);
    if (!sh) return; // Exit if Ops Inbox sheet doesn't exist

    // This function is critical, so we get headers without validation
    // to avoid circular error logging if the Ops Inbox is misconfigured.
    const headerMap = getHeaderMap_(sh);
    if (!headerMap || !headerMap['Timestamp']) {
       // Cannot log if the sheet is fundamentally broken.
       console.error(`Could not log to Ops Inbox. Sheet or headers are misconfigured.`);
       return;
    }

    const rowData = {
      'Name': sfid ? `SFID ${sfid} ${type}` : `pipeline ${type}`,
      'Source SFID': sfid || '',
      'Type': type,
      'Resolved': false,
      'Details': details || '',
      'Timestamp': new Date()
    };

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const row = headers.map(h => rowData[h] || '');

    sh.appendRow(row);
  } catch(err) {
    // Fallback to console logging if Ops Inbox logging fails
    console.error(`Failed to write to Ops Inbox. Type: ${type}, SFID: ${sfid}, Details: ${details}. Error: ${err.message}`);
  }
}

/**
 * Generates a human-readable string explaining why a status advance is blocked.
 * This function is exposed as a custom formula in the sheet and is also called internally by the script.
 * @param {string} status The target status.
 * @param {boolean} canGlobal Global advancement prerequisite.
 * @param {boolean} canPerm Permitting advancement prerequisite.
 * @param {boolean} canSched Scheduled advancement prerequisite.
 * @param {boolean} canInspect Inspections advancement prerequisite.
 * @param {boolean} canDone Done advancement prerequisite.
 * @returns {string} A concatenated string of block reasons.
 */
function GET_ADVANCE_BLOCK_REASON(status, canGlobal, canPerm, canSched, canInspect, canDone) {
  const reasons = [];
  if (!toBool_(canGlobal)) reasons.push('Missing global data or overdue tasks or duplicate SFID');
  if (status === 'Permitting' && !toBool_(canPerm)) reasons.push('Gate to Permitting not met: deposit, permit app, or tasks incomplete');
  if (status === 'Scheduled' && !toBool_(canSched)) reasons.push('Gate to Scheduled not met: permits approved, artifacts in Drive, revenue+probability, or tasks incomplete');
  if (status === 'Inspections' && !toBool_(canInspect)) reasons.push('Gate to Inspections not met: equipment, site prep, or rough inspections');
  if (status === 'Done' && !toBool_(canDone)) reasons.push('Gate to Done not met: final inspection, payment, artifacts in Drive, or tasks incomplete');
  return reasons.join(' • ');
}

// --- Utility Functions --- //

/**
 * Gets a map of header names to their column numbers (1-indexed).
 * This is now resilient to column order and extra columns.
 * Caching could be added here for performance in very large sheets.
 * @param {Sheet} sheet The sheet object to get headers from.
 * @param {string[]} requiredHeaders An optional array of headers that MUST exist.
 * @returns {Object|null} A map of {headerName: columnNumber} or null if validation fails.
 */
function getHeaderMap_(sheet, requiredHeaders) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = { _width: headerRow.length };

  headerRow.forEach((header, index) => {
    if (header) {
      map[header.toString().trim()] = index + 1;
    }
  });

  if (requiredHeaders) {
    for (const h of requiredHeaders) {
      if (!map[h]) {
        logOpsInbox_('script_error', '', `Configuration error in sheet '${sheet.getName()}': Missing required header '${h}'.`);
        return null;
      }
    }
  }
  return map;
}

function stampIfEmpty_(state, headerName) {
  if (!state.getValue(headerName)) {
    state.setValue(headerName, new Date());
  }
}

function isOverrideActive_(overrideUntil) {
  if (!overrideUntil) return false;
  try {
    const until = new Date(overrideUntil);
    return Date.now() <= until.getTime();
  } catch (e) {
    return false;
  }
}

function toBool_(v) {
  if (v === true || v === false) return v;
  if (typeof v === 'string') {
    const lower = v.toLowerCase();
    if (lower === 'true') return true;
    if (lower === 'false') return false;
  }
  if (typeof v === 'number') return v !== 0;
  return !!v; // Coerce other truthy/falsy values
}

// --- State Management & Commit Logic --- //

/**
 * A helper class to manage the state of a row in memory.
 * All read/write operations during a script run should go through this object.
 */
class RowState {
  constructor(headerMap, initialValues, initialNotes) {
    this.headerMap = headerMap;
    this.initialValues = [...initialValues];
    this.finalValues = [...initialValues];
    this.initialNotes = [...initialNotes];
    this.finalNotes = [...initialNotes];

    // Create a reverse map for col number to header name
    this.colToHeader = {};
    for (const header in headerMap) {
      if (header !== '_width') {
        this.colToHeader[headerMap[header]] = header;
      }
    }
  }

  getHeader(colNumber) { return this.colToHeader[colNumber]; }
  getCol(headerName) { return this.headerMap[headerName]; }

  getValue(headerName) { return this.finalValues[this.getCol(headerName) - 1]; }
  getNote(headerName) { return this.finalNotes[this.getCol(headerName) - 1]; }

  setValue(headerName, value) { this.finalValues[this.getCol(headerName) - 1] = value; }
  setNote(headerName, note) { this.finalNotes[this.getCol(headerName) - 1] = note; }

  revertValue(e, headerName) {
    const oldValue = e.oldValue !== undefined ? e.oldValue : this.initialValues[this.getCol(headerName) - 1];
    this.setValue(headerName, oldValue);
  }
}

/**
 * Compares the final state with the initial state and writes changes back to the sheet.
 * This is the only place where `range.setValues` or `range.setNotes` should be called.
 * @param {Range} range The range object for the entire row.
 * @param {RowState} state The final state object for the row.
 */
function commitRowState_(range, state) {
  // Compare final values to initial values and write if changed.
  if (JSON.stringify(state.initialValues) !== JSON.stringify(state.finalValues)) {
    range.setValues([state.finalValues]);
  }
  // Compare final notes to initial notes and write if changed.
  if (JSON.stringify(state.initialNotes) !== JSON.stringify(state.finalNotes)) {
    range.setNotes([state.finalNotes]);
  }
}

// --- Admin and Properties Management --- //

/**
 * Retrieves the list of admin users from Script Properties.
 * On first run, it will initialize with a default placeholder.
 * @returns {string[]} An array of admin email addresses.
 */
function getAdmins_() {
  const properties = PropertiesService.getScriptProperties();
  let admins = properties.getProperty('OVERRIDE_ADMINS');
  if (!admins) {
    // Initialize with a placeholder if no admins are set.
    // The user will be instructed to configure this via a menu.
    const defaultAdmin = Session.getEffectiveUser().getEmail() || 'jules@example.com';
    properties.setProperty('OVERRIDE_ADMINS', defaultAdmin);
    admins = defaultAdmin;
  }
  return admins.split(',').map(s => s.trim()).filter(Boolean);
}

/**
 * Stores a list of admin users in Script Properties.
 * @param {string[]} adminsArray An array of admin email addresses.
 */
function setAdmins_(adminsArray) {
  const adminString = adminsArray.join(',');
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('OVERRIDE_ADMINS', adminString);
}
