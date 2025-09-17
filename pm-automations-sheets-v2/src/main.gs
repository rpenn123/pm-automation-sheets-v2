/**
 * Project Pipeline with Dependency Enforcement - Google Sheets Edition
 * Apps Script
 *
 * Refactored for robustness, maintainability, and automated setup.
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
    DASHBOARD: 'Dashboard',
    LISTS: 'Lists',
  },

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

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Mobility123 PM')
    .addItem('Verify & Setup Sheet', 'runFullSetup')
    .addSeparator()
    .addItem('Configure Admins', 'showAdminConfigUi')
    .addToUi();
}

function showAdminConfigUi() {
  const ui = SpreadsheetApp.getUi();
  const currentAdmins = getAdmins_().join(', ');
  const result = ui.prompt(
    'Configure Admins',
    'Enter a comma-separated list of admin email addresses. These users can use the "Override: Allow Advance" checkbox.',
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
 * Runs the full, automated setup of the Google Sheet.
 */
function runFullSetup() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirm Setup',
    'This will set up all required sheets, headers, and formulas. Existing data in sheets with the same names may be overwritten. Are you sure you want to continue?',
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    ui.alert('Setup canceled.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = ['Setup & Verification Report:'];

  try {
    SpreadsheetApp.getUi().showModalDialog(
       HtmlService.createHtmlOutput('<p>Setup in progress, please wait...</p>').setWidth(300).setHeight(100),
      'Setup'
    );

    // 1. Create sheets
    createSheets_(ss, report);

    // 2. Set headers and formulas
    setupSheetContents_(ss, report);

    // 3. Install trigger
    installTrigger_(report);

    report.push('\n✅✅✅ Setup Complete! ✅✅✅');
    report.push('\nNext Steps:');
    report.push('1. Populate the `Lists` sheet with your dropdown options.');
    report.push('2. Use "Data -> Named ranges" and "Data -> Data validation" to hook up the dropdowns.');
    report.push('3. Use the "Configure Admins" menu to set who can use the override feature.');

  } catch (e) {
    report.push(`❌ An error occurred: ${e.message}`);
    report.push(`Stack: ${e.stack}`);
  } finally {
    ui.alert(report.join('\n'));
  }
}

function createSheets_(ss, report) {
  const requiredSheets = [
      CFG.SHEETS.PIPELINE, CFG.SHEETS.TASKS, CFG.SHEETS.OPS_INBOX,
      CFG.SHEETS.UPCOMING, CFG.SHEETS.FRAMING, CFG.SHEETS.DASHBOARD, CFG.SHEETS.LISTS
  ];
  const existingSheets = ss.getSheets().map(s => s.getName());
  let created = [];

  requiredSheets.forEach(name => {
    if (!existingSheets.includes(name)) {
      ss.insertSheet(name);
      created.push(name);
    }
  });

  if (created.length > 0) {
    report.push(`✅ Created missing sheets: ${created.join(', ')}.`);
  } else {
    report.push('✅ All required sheets are present.');
  }
}

function setupSheetContents_(ss, report) {
    const allHeaders = SHEET_SETUP_CONFIG.HEADERS;
    for (const sheetName in allHeaders) {
        const sh = ss.getSheetByName(sheetName);
        if (!sh) {
            report.push(`⚠️ Could not find sheet: ${sheetName}. Skipping content setup.`);
            continue;
        }

        // Set Headers
        const headers = allHeaders[sheetName];
        if (headers && headers.length > 0) {
            sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
            report.push(`✅ Set headers for '${sheetName}'.`);
        }

        // Set Array Formulas
        const arrayFormulas = SHEET_SETUP_CONFIG.ARRAY_FORMULAS[sheetName];
        if (arrayFormulas) {
            const headerMap = getHeaderMap_(sh);
            for (const header in arrayFormulas) {
                const col = headerMap[header];
                if (col) {
                    sh.getRange(2, col).setFormula(arrayFormulas[header]);
                } else {
                    report.push(`⚠️ Could not find column '${header}' in '${sheetName}' to apply formula.`);
                }
            }
            report.push(`✅ Applied array formulas to '${sheetName}'.`);
        }

        // Set Static Formulas (for dashboards, etc.)
        const staticFormulas = SHEET_SETUP_CONFIG.STATIC_FORMULAS[sheetName];
        if (staticFormulas) {
            for (const cell in staticFormulas) {
                sh.getRange(cell).setFormula(staticFormulas[cell]);
            }
            report.push(`✅ Applied static formulas to '${sheetName}'.`);
        }
    }
}


function installTrigger_(report) {
  const triggers = ScriptApp.getProjectTriggers();
  const hasOnEditTrigger = triggers.some(t =>
    t.getEventType() === ScriptApp.EventType.ON_EDIT &&
    t.getHandlerFunction() === 'onEditHandler'
  );

  if (!hasOnEditTrigger) {
    ScriptApp.newTrigger('onEditHandler')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    report.push('✅ Installed the required onEdit trigger.');
  } else {
    report.push('✅ The onEdit trigger is correctly installed.');
  }
}

// --- Main Trigger --- //

function onEditHandler(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const sheetName = sh.getName();

    if (sheetName === CFG.SHEETS.PIPELINE) {
      return handlePipelineEdit_(e);
    }
    if (sheetName === CFG.SHEETS.TASKS) {
      return handleTasksEdit_(e);
    }
  } catch (err) {
    logOpsInbox_('script_error', '', `Unhandled error: ${err.message} Stack: ${err.stack}`);
  }
}

// --- Edit Handlers --- //

function handlePipelineEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 1) return;

  const headerMap = getHeaderMap_(sh, SHEET_SETUP_CONFIG.HEADERS[CFG.SHEETS.PIPELINE]);
  if (!headerMap) return;

  const range = sh.getRange(row, 1, 1, headerMap._width);
  const state = new RowState(headerMap, range.getValues()[0], range.getNotes()[0]);
  const colName = state.getHeader(col);

  // --- New Row Detection ---
  // If last_validated_ts is empty, this is the first time the script is running on this row.
  if (!state.getValue('last_validated_ts')) {
    const now = new Date();
    state.setValue('Created Date', now);
    state.setValue('Created Month', Utilities.formatDate(now, CFG.TZ, 'yyyy-MM'));
  }
  // --- End New Row Detection ---

  if (CFG.ZAPIER_ONLY_COLUMNS.includes(colName)) {
    return;
  }

  if (colName === 'Permits' && state.getValue('Permits') === 'Approved') {
    stampIfEmpty_(state, 'ts_permits_approved');
  }

  if (colName === 'Project Status') {
    enforceStatusChange_(e, state);
    paymentGuard_(e, state);
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

  state.setValue('last_validated_ts', new Date());
  commitRowState_(range, state);
}

function handleTasksEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  if (row === 1) return;

  const headerMap = getHeaderMap_(sh, SHEET_SETUP_CONFIG.HEADERS[CFG.SHEETS.TASKS]);
  if (!headerMap) return;

  const range = sh.getRange(row, 1, 1, headerMap._width);
  const state = new RowState(headerMap, range.getValues()[0], range.getNotes()[0]);

  if (state.getHeader(e.range.getColumn()) === 'Status' && state.getValue('Status') === 'Done') {
    stampIfEmpty_(state, 'Completed Date');
  }

  commitRowState_(range, state);
}

// --- Core Logic Functions (largely unchanged) --- //

function handleOverrideToggle_(e, state) {
    if (state.getValue('Override: Allow Advance')) {
        const userEmail = Session.getActiveUser().getEmail();
        const admins = getAdmins_();
        if (userEmail && admins.includes(userEmail)) {
            const until = new Date(Date.now() + 24 * 60 * 60 * 1000);
            state.setValue('Override until', until);
        } else {
            state.revertValue(e, 'Override: Allow Advance');
            state.setNote('Project Status', `Override is only available for admins: ${admins.join(', ')}`);
        }
    }
}

function enforceStatusChange_(e, state) {
  const statusNew = state.getValue('Project Status');

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
    logOpsInbox_('data_missing', state.getValue('SFID'), `Blocked changing status to ${statusNew}. ${reason}`);
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

function paymentGuard_(e, state) {
  if (state.getValue('Project Status') === 'Done' && !toBool_(state.getValue('Final payment received'))) {
    state.revertValue(e, 'Project Status');
    state.setNote('Project Status', 'Payment must precede Done.');
    logOpsInbox_('data_missing', state.getValue('SFID'), 'Attempted to set Done without Final payment received.');
  }
}

function detectDuplicateSFID_(pipelineSheet, row, state) {
  const sfid = state.getValue('SFID');
  if (!sfid) return;
  const lastRow = pipelineSheet.getLastRow();
  if (lastRow < 2) return;

  const sfidCol = state.getCol('SFID');
  const sfids = pipelineSheet.getRange(2, sfidCol, lastRow - 1, 1).getValues().flat();
  const isDup = sfids.filter(v => v === sfid).length > 1;

  state.setValue('Duplicate SFID', isDup);
  if (isDup) logOpsInbox_('duplicate_sfid', sfid, `Detected duplicate SFID: ${sfid}`);
}

function clearBlocked_(state) {
  state.setValue('Blocked since', '');
  if ((state.getNote('Project Status') || '').startsWith('Advance blocked')) {
     state.setNote('Project Status', '');
  }
}

function logOpsInbox_(type, sfid, details) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.OPS_INBOX);
    if (!sh) return;

    const headers = SHEET_SETUP_CONFIG.HEADERS[CFG.SHEETS.OPS_INBOX];
    if (!headers || headers.length === 0) return;

    const rowData = {
      'Name': sfid ? `SFID ${sfid} ${type}` : `pipeline ${type}`,
      'Source SFID': sfid || '',
      'Type': type,
      'Resolved': false,
      'Details': details || '',
      'Timestamp': new Date()
    };
    const row = headers.map(h => rowData[h] || '');
    sh.appendRow(row);
  } catch(err) {
    console.error(`Failed to write to Ops Inbox. Type: ${type}, SFID: ${sfid}, Details: ${details}. Error: ${err.message}`);
  }
}

function GET_ADVANCE_BLOCK_REASON(status, canGlobal, canPerm, canSched, canInspect, canDone) {
  const reasons = [];
  if (!toBool_(canGlobal)) reasons.push('Missing global data or overdue tasks or duplicate SFID');
  if (status === 'Permitting' && !toBool_(canPerm)) reasons.push('Gate to Permitting not met');
  if (status === 'Scheduled' && !toBool_(canSched)) reasons.push('Gate to Scheduled not met');
  if (status === 'Inspections' && !toBool_(canInspect)) reasons.push('Gate to Inspections not met');
  if (status === 'Done' && !toBool_(canDone)) reasons.push('Gate to Done not met');
  return reasons.join(' • ');
}

// --- Utility Functions --- //

function getHeaderMap_(sheet, requiredHeaders) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = { _width: headerRow.length };
  headerRow.forEach((header, index) => {
    if (header) map[header.toString().trim()] = index + 1;
  });

  if (requiredHeaders) {
    for (const h of requiredHeaders) {
      if (!map[h]) {
        logOpsInbox_('script_error', '', `Config error in '${sheet.getName()}': Missing header '${h}'.`);
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
  return overrideUntil && new Date(overrideUntil).getTime() >= Date.now();
}

function toBool_(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'string') return v.toLowerCase() === 'true';
  if (typeof v === 'number') return v !== 0;
  return !!v;
}

// --- State Management Class --- //

class RowState {
  constructor(headerMap, initialValues, initialNotes) {
    this.headerMap = headerMap;
    this.initialValues = [...initialValues];
    this.finalValues = [...initialValues];
    this.initialNotes = [...initialNotes];
    this.finalNotes = [...initialNotes];
    this.colToHeader = Object.fromEntries(Object.entries(headerMap).map(([k, v]) => [v, k]));
  }
  getHeader(col) { return this.colToHeader[col]; }
  getCol(header) { return this.headerMap[header]; }
  getValue(header) { return this.finalValues[this.getCol(header) - 1]; }
  getNote(header) { return this.finalNotes[this.getCol(header) - 1]; }
  setValue(header, value) { this.finalValues[this.getCol(header) - 1] = value; }
  setNote(header, note) { this.finalNotes[this.getCol(header) - 1] = note; }
  revertValue(e, header) {
    const oldValue = e.oldValue !== undefined ? e.oldValue : this.initialValues[this.getCol(header) - 1];
    this.setValue(header, oldValue);
  }
}

function commitRowState_(range, state) {
  if (JSON.stringify(state.initialValues) !== JSON.stringify(state.finalValues)) {
    range.setValues([state.finalValues]);
  }
  if (JSON.stringify(state.initialNotes) !== JSON.stringify(state.finalNotes)) {
    range.setNotes([state.finalNotes]);
  }
}

// --- Admin and Properties Management --- //

function getAdmins_() {
  const properties = PropertiesService.getScriptProperties();
  let admins = properties.getProperty('OVERRIDE_ADMINS');
  if (!admins) {
    const defaultAdmin = Session.getEffectiveUser().getEmail() || 'jules@example.com';
    properties.setProperty('OVERRIDE_ADMINS', defaultAdmin);
    return [defaultAdmin];
  }
  return admins.split(',').map(s => s.trim()).filter(Boolean);
}

function setAdmins_(adminsArray) {
  const adminString = adminsArray.join(',');
  PropertiesService.getScriptProperties().setProperty('OVERRIDE_ADMINS', adminString);
}
