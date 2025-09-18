/**
 * @file main.gs
 * @description This script automates a project management pipeline in Google Sheets.
 * It handles automated sheet setup, enforces business rules for project status changes,
 * logs errors, and provides utility functions for sheet manipulation.
 *
 * @license MIT
 * @version 2.0.0
 */

// --- Global Configuration --- //

/**
 * CFG is a configuration object holding global constants for the script.
 * This centralization makes it easier to manage and update key settings.
 * @const
 */
const CFG = {
  /**
   * The script's time zone. Used for consistent date/time formatting.
   * Defaults to the script's time zone but can be overridden.
   * @type {string}
   */
  TZ: Session.getScriptTimeZone() || 'America/New_York',

  /**
   * An enumeration of sheet names used throughout the script.
   * Using this object prevents hardcoding strings and reduces errors from typos.
   * @enum {string}
   */
  SHEETS: {
    PIPELINE: 'Project Pipeline',
    TASKS: 'Tasks',
    UPCOMING: 'Upcoming',
    FRAMING: 'Framing',
    OPS_INBOX: 'Ops Inbox',
    DASHBOARD: 'Dashboard',
    LISTS: 'Lists',
    SETTINGS: 'Settings',
  },

  /**
   * A list of column headers that are expected to be updated by external services
   * like Zapier. The onEdit trigger will ignore edits in these columns to prevent
   * unnecessary script executions or infinite loops.
   * @type {string[]}
   */
  ZAPIER_ONLY_COLUMNS: [
    'last_upcoming_notified_ts',
    'last_framing_notified_ts',
    'last_block_notified_ts',
    'last_escalation_notified_ts',
    'pipeline_last_transfer_ts',
    'pipeline_last_transfer_status'
  ],
};

// --- Menu & Setup Functions --- //

/**
 * A simple trigger that runs when the spreadsheet is opened.
 * It creates a custom menu for accessing the script's features.
 * @param {Object} e The onOpen event object.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Mobility123 PM')
    .addItem('Verify & Setup Sheet', 'runFullSetup')
    .addSeparator()
    .addItem('Configure Admins', 'showAdminConfigUi')
    .addItem('Manually Update Timestamp', 'updateTimestamp')
    .addToUi();
}

/**
 * Displays a UI prompt to configure the list of administrative users.
 * Admins have special permissions, such as using the override checkbox.
 */
function showAdminConfigUi() {
  const ui = SpreadsheetApp.getUi();
  const currentAdmins = getAdmins_().join(', ');
  const result = ui.prompt(
    'Configure Admins',
    'Enter a comma-separated list of admin email addresses. These users can use the "Override: Allow Advance" checkbox.',
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response
  if (result.getSelectedButton() == ui.Button.OK) {
    const newAdminsText = result.getResponseText();
    const newAdmins = newAdminsText.split(',').map(s => s.trim()).filter(Boolean);
    setAdmins_(newAdmins);
    ui.alert('Success', `Admins updated to: ${newAdmins.join(', ')}`, ui.ButtonSet.OK);
  }
}

/**
 * Orchestrates the entire setup process for the spreadsheet.
 * This function is called from the custom menu. It's designed to be idempotent,
 * meaning it can be run multiple times without causing issues.
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
    // Show a modal dialog to inform the user that the setup is running.
    SpreadsheetApp.getUi().showModalDialog(
       HtmlService.createHtmlOutput('<p>Setup in progress, please wait...</p>').setWidth(300).setHeight(100),
      'Setup'
    );

    // Execute setup steps
    createSheets_(ss, report);
    setupSheetContents_(ss, report);
    setupDataValidation_(ss, report); // New function call
    installTrigger_(report);

    report.push('\n✅✅✅ Setup Complete! ✅✅✅');
    report.push('\nNext Steps:');
    report.push('1. Review the `Lists` sheet to customize dropdown options.');
    report.push('2. Use the "Configure Admins" menu to set who can use the override feature.');

  } catch (e) {
    // Log any errors that occur during setup for easier debugging.
    report.push(`❌ An error occurred: ${e.message}`);
    report.push(`Stack: ${e.stack}`);
  } finally {
    // Display the final report to the user.
    ui.alert(report.join('\n'));
  }
}

/**
 * Creates any required sheets that are missing from the spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {string[]} report An array to log setup actions for the user.
 * @private
 */
function createSheets_(ss, report) {
  const requiredSheets = [
      CFG.SHEETS.PIPELINE, CFG.SHEETS.TASKS, CFG.SHEETS.OPS_INBOX,
      CFG.SHEETS.UPCOMING, CFG.SHEETS.FRAMING, CFG.SHEETS.DASHBOARD, CFG.SHEETS.LISTS,
      CFG.SHEETS.SETTINGS
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

/**
 * Sets the headers and formulas for all managed sheets based on SHEET_SETUP_CONFIG.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {string[]} report An array to log setup actions for the user.
 * @private
 */
function setupSheetContents_(ss, report) {
    const allHeaders = SHEET_SETUP_CONFIG.HEADERS;
    for (const sheetName in allHeaders) {
        const sh = ss.getSheetByName(sheetName);
        if (!sh) {
            report.push(`⚠️ Could not find sheet: ${sheetName}. Skipping content setup.`);
            continue;
        }

        // Set Headers if defined in config
        const headers = allHeaders[sheetName];
        if (headers && headers.length > 0) {
            sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
            report.push(`✅ Set headers for '${sheetName}'.`);
        }

        // Set Array Formulas if defined in config
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

        // Set Static Formulas (e.g., for dashboards) if defined in config
        const staticFormulas = SHEET_SETUP_CONFIG.STATIC_FORMULAS[sheetName];
        if (staticFormulas) {
            for (const cell in staticFormulas) {
                sh.getRange(cell).setFormula(staticFormulas[cell]);
            }
            report.push(`✅ Applied static formulas to '${sheetName}'.`);
        }
    }

    // --- Special Setup for Settings Sheet ---
    const settingsSheet = ss.getSheetByName(CFG.SHEETS.SETTINGS);
    if (settingsSheet && settingsSheet.getLastRow() === 1) { // Only populate if it's a fresh sheet
        const settingsData = [
            ['Current Time', new Date(), 'Automatically updated timestamp used by formulas to avoid volatile functions like NOW() and TODAY(). Updated by a time-based trigger.'],
            ['Upcoming Days Threshold', 7, 'Number of days to look ahead for "upcoming" items on the dashboard.'],
            ['Staleness Days Threshold', 7, 'Number of days without an edit before a project is flagged as stale.']
        ];
        settingsSheet.getRange(2, 1, 3, 3).setValues(settingsData);
        settingsSheet.getRange('B2').setNumberFormat('yyyy-mm-dd hh:mm:ss');
        settingsSheet.getRange('C:C').setWrap(true);
        report.push(`✅ Populated initial data in '${CFG.SHEETS.SETTINGS}'.`);
    }
}

/**
 * Sets up data validation rules (dropdowns) for the sheets.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {string[]} report An array to log setup actions for the user.
 * @private
 */
function setupDataValidation_(ss, report) {
    const listSheet = ss.getSheetByName(CFG.SHEETS.LISTS);
    if (!listSheet) {
        report.push(`⚠️ Could not find sheet: ${CFG.SHEETS.LISTS}. Skipping dropdown setup.`);
        return;
    }

    // --- 1. Populate Lists Sheet with Default Data ---
    if (listSheet.getLastRow() < 2) { // Only populate if it seems empty
        const listData = SHEET_SETUP_CONFIG.LIST_SHEET_DATA;
        const headers = Object.keys(listData);
        listSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

        headers.forEach((header, colIndex) => {
            const values = listData[header].map(v => [v]); // Convert to 2D array
            if (values.length > 0) {
                listSheet.getRange(2, colIndex + 1, values.length, 1).setValues(values);
            }
        });
        report.push(`✅ Populated default options in '${CFG.SHEETS.LISTS}'.`);
    }

    // --- 2. Create Named Ranges from Lists Sheet ---
    const listHeaders = listSheet.getRange(1, 1, 1, listSheet.getLastColumn()).getValues()[0];
    listHeaders.forEach((header, colIndex) => {
        if (!header) return;
        const lastRow = listSheet.getLastRow();
        const range = listSheet.getRange(2, colIndex + 1, lastRow -1);
        const namedRangeName = `List_${header.replace(/ /g, '')}`;
        ss.setNamedRange(namedRangeName, range);
    });
    report.push(`✅ Created named ranges for all lists.`);


    // --- 3. Apply Data Validation Rules ---
    const rules = SHEET_SETUP_CONFIG.DATA_VALIDATION_RULES;
    for (const sheetName in rules) {
        const sh = ss.getSheetByName(sheetName);
        if (!sh) {
            report.push(`⚠️ Could not find sheet '${sheetName}' to apply dropdowns.`);
            continue;
        }

        const headerMap = getHeaderMap_(sh);
        const sheetRules = rules[sheetName];

        for (const header in sheetRules) {
            const col = headerMap[header];
            const namedRangeName = sheetRules[header];
            const namedRange = ss.getRangeByName(namedRangeName);

            if (col && namedRange) {
                const rule = SpreadsheetApp.newDataValidation()
                    .requireValueInRange(namedRange)
                    .setAllowInvalid(false) // Disallow values not in the list
                    .setHelpText(`Select a valid ${header}.`)
                    .build();
                sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(rule);
            } else {
                report.push(`⚠️ Could not apply dropdown for '${header}' in '${sheetName}'. Column or Named Range missing.`);
            }
        }
        report.push(`✅ Applied dropdowns to '${sheetName}'.`);
    }
}

/**
 * Installs the onEdit trigger for the script if it doesn't already exist.
 * @param {string[]} report An array to log setup actions for the user.
 * @private
 */
function installTrigger_(report) {
  const triggers = ScriptApp.getProjectTriggers();

  // Check for onEdit trigger
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

  // Check for time-based trigger for updateTimestamp
  const hasTimeTrigger = triggers.some(t =>
    t.getEventType() === ScriptApp.EventType.CLOCK &&
    t.getHandlerFunction() === 'updateTimestamp'
  );

  if (!hasTimeTrigger) {
    ScriptApp.newTrigger('updateTimestamp')
      .timeBased()
      .everyMinutes(1)
      .create();
    report.push('✅ Installed time-based trigger for timestamp updates.');
  } else {
    report.push('✅ The time-based trigger is correctly installed.');
  }
}

// --- Time-based Trigger Functions --- //

/**
 * Updates the 'Current Time' value in the 'Settings' sheet.
 * This function is designed to be run on a time-based trigger (e.g., every minute)
 * to provide a non-volatile timestamp for formulas, improving sheet performance.
 *
 * To set up the trigger:
 * 1. Open the Apps Script editor.
 * 2. Go to Edit > Current project's triggers.
 * 3. Click "+ Add Trigger".
 * 4. Choose function to run: "updateTimestamp".
 * 5. Choose which deployment should run: "Head".
 * 6. Select event source: "Time-driven".
 * 7. Select type of time-based trigger: "Minutes timer".
 * 8. Select minute interval: "Every minute".
 * 9. Click "Save".
 */
function updateTimestamp() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(CFG.SHEETS.SETTINGS);
    if (!settingsSheet) {
      // Log to Ops Inbox if settings sheet is missing
      logOpsInbox_('script_error', '', `Critical Error: The '${CFG.SHEETS.SETTINGS}' sheet is missing. Cannot update timestamp.`);
      return;
    }

    // Find the 'Current Time' setting row. This is more robust than hardcoding a cell.
    const settingNames = settingsSheet.getRange('A2:A').getValues().flat();
    const rowIndex = settingNames.indexOf('Current Time');

    if (rowIndex === -1) {
      // Log to Ops Inbox if the setting is missing
      logOpsInbox_('script_error', '', `Critical Error: 'Current Time' setting not found in '${CFG.SHEETS.SETTINGS}' sheet.`);
      return;
    }

    // Update the timestamp in column B of the found row (A is 1-based, rowIndex is 0-based, so row is rowIndex + 2)
    const targetCell = settingsSheet.getRange(rowIndex + 2, 2);
    targetCell.setValue(new Date());

  } catch (e) {
    // Log any unexpected errors to the Ops Inbox for review.
    const errorMessage = `Error in updateTimestamp: ${e.message}. Stack: ${e.stack}`;
    logOpsInbox_('script_error', '', errorMessage);
    // Also log to the console for immediate debugging.
    console.error(errorMessage);
  }
}

// --- Main Trigger Handler --- //

/**
 * The main onEdit trigger handler. This function is called by Google Apps Script
 * whenever a user edits the spreadsheet.
 * @param {Object} e The onEdit event object.
 */
function onEditHandler(e) {
  try {
    // Exit if the event object is invalid or doesn't have a range.
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const sheetName = sh.getName();

    // Delegate to the appropriate handler based on the edited sheet.
    if (sheetName === CFG.SHEETS.PIPELINE) {
      return handlePipelineEdit_(e);
    }
    if (sheetName === CFG.SHEETS.TASKS) {
      return handleTasksEdit_(e);
    }
  } catch (err) {
    // Catch any unhandled errors and log them to the Ops Inbox.
    logOpsInbox_('script_error', '', `Unhandled error: ${err.message} Stack: ${err.stack}`);
  }
}

// --- Sheet-Specific Edit Handlers --- //

/**
 * Handles edits made to the 'Project Pipeline' sheet.
 * @param {Object} e The onEdit event object.
 * @private
 */
function handlePipelineEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 1) return; // Ignore edits to the header row.

  // Get a map of headers to column numbers for easy data access.
  const headerMap = getHeaderMap_(sh, SHEET_SETUP_CONFIG.HEADERS[CFG.SHEETS.PIPELINE]);
  if (!headerMap) return; // Exit if headers are not configured correctly.

  // Use the RowState class to efficiently manage reads and writes for the edited row.
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

  // Ignore edits to columns that are managed by Zapier.
  if (CFG.ZAPIER_ONLY_COLUMNS.includes(colName)) {
    return;
  }

  // --- Business Logic Checks ---

  // Timestamp when a permit is first marked 'Approved'.
  if (colName === 'Permits' && state.getValue('Permits') === 'Approved') {
    stampIfEmpty_(state, 'ts_permits_approved');
  }

  // Handle status changes, which triggers core validation logic.
  if (colName === 'Project Status') {
    enforceStatusChange_(e, state);
    paymentGuard_(e, state);
  }

  // Timestamp when a project first enters the 'Permitting' status.
  if (state.getValue('Project Status') === 'Permitting') {
    stampIfEmpty_(state, 'ts_entered_permitting');
  }

  // Check for duplicate Salesforce IDs.
  if (colName === 'SFID') {
    detectDuplicateSFID_(sh, row, state);
  }

  // Handle the admin override checkbox.
  if (colName === 'Override: Allow Advance') {
    handleOverrideToggle_(e, state);
  }

  // Set the timestamp for the last validation run on this row.
  state.setValue('last_validated_ts', new Date());

  // Commit all changes made to the row state back to the sheet.
  commitRowState_(range, state);
}

/**
 * Handles edits made to the 'Tasks' sheet.
 * @param {Object} e The onEdit event object.
 * @private
 */
function handleTasksEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  if (row === 1) return; // Ignore header row

  const headerMap = getHeaderMap_(sh, SHEET_SETUP_CONFIG.HEADERS[CFG.SHEETS.TASKS]);
  if (!headerMap) return;

  const range = sh.getRange(row, 1, 1, headerMap._width);
  const state = new RowState(headerMap, range.getValues()[0], range.getNotes()[0]);

  // If a task's status is changed to 'Done', timestamp the completion date.
  if (state.getHeader(e.range.getColumn()) === 'Status' && state.getValue('Status') === 'Done') {
    stampIfEmpty_(state, 'Completed Date');
  }

  commitRowState_(range, state);
}

// --- Core Logic Functions --- //

/**
 * Handles the logic for the "Override: Allow Advance" checkbox.
 * Only allows admins to check the box and sets a 24-hour override window.
 * @param {Object} e The onEdit event object.
 * @param {RowState} state The state object for the edited row.
 */
function handleOverrideToggle_(e, state) {
    if (state.getValue('Override: Allow Advance')) {
        const userEmail = Session.getActiveUser().getEmail();
        const admins = getAdmins_();
        if (userEmail && admins.includes(userEmail)) {
            // Set the override expiration to 24 hours from now.
            const until = new Date(Date.now() + 24 * 60 * 60 * 1000);
            state.setValue('Override until', until);
        } else {
            // If user is not an admin, revert the checkbox and add a note.
            state.revertValue(e, 'Override: Allow Advance');
            state.setNote('Project Status', `Override is only available for admins: ${admins.join(', ')}`);
        }
    }
}

/**
 * Enforces the phase-gate rules for changing a project's status.
 * It checks various formula-driven 'can_advance' columns.
 * @param {Object} e The onEdit event object.
 * @param {RowState} state The state object for the edited row.
 */
function enforceStatusChange_(e, state) {
  const statusNew = state.getValue('Project Status');

  // If a valid override is active, allow any status change and exit.
  if (state.getValue('Override: Allow Advance') && isOverrideActive_(state.getValue('Override until'))) {
    state.setValue('Status (Last Valid)', statusNew);
    clearBlocked_(state);
    return;
  }

  // Read the boolean values from the 'can_advance' helper columns.
  const canGlobal = toBool_(state.getValue('can_advance_globally'));
  const canPerm = toBool_(state.getValue('can_advance_to_Permitting'));
  const canSched = toBool_(state.getValue('can_advance_to_Scheduled'));
  const canInspect = toBool_(state.getValue('can_advance_to_Inspections'));
  const canDone = toBool_(state.getValue('can_advance_to_Done'));

  // Determine if the status change is valid based on the 'can_advance' flags.
  let valid = true;
  if (!canGlobal) valid = false;
  if (statusNew === 'Permitting' && !canPerm) valid = false;
  if (statusNew === 'Scheduled' && !canSched) valid = false;
  if (statusNew === 'Inspections' && !canInspect) valid = false;
  if (statusNew === 'Done' && !canDone) valid = false;

  // If the change is not valid, revert it and provide feedback.
  if (!valid) {
    state.revertValue(e, 'Project Status');
    const reason = GET_ADVANCE_BLOCK_REASON(statusNew, canGlobal, canPerm, canSched, canInspect, canDone);
    state.setNote('Project Status', `Advance blocked. Reasons: ${reason}`);
    // If not already blocked, set the 'Blocked since' timestamp.
    if (!state.getValue('Blocked since')) {
      state.setValue('Blocked since', new Date());
    }
    logOpsInbox_('data_missing', state.getValue('SFID'), `Blocked changing status to ${statusNew}. ${reason}`);
    return;
  }

  // If the change is valid, update the 'Last Valid Status' and clear any blocks.
  state.setValue('Status (Last Valid)', statusNew);
  clearBlocked_(state);

  // Timestamp key project milestones.
  if (statusNew === 'Scheduled') stampIfEmpty_(state, 'ts_first_scheduled');
  if (statusNew === 'Done') stampIfEmpty_(state, 'ts_marked_done');

  // Uncheck the override box after a successful advance.
  if (state.getValue('Override: Allow Advance')) {
    state.setValue('Override: Allow Advance', false);
    state.setValue('Override until', '');
  }
}

/**
 * Prevents a project from being marked as 'Done' if final payment has not been received.
 * @param {Object} e The onEdit event object.
 * @param {RowState} state The state object for the edited row.
 */
function paymentGuard_(e, state) {
  if (state.getValue('Project Status') === 'Done' && !toBool_(state.getValue('Final payment received'))) {
    state.revertValue(e, 'Project Status');
    state.setNote('Project Status', 'Payment must precede Done.');
    logOpsInbox_('data_missing', state.getValue('SFID'), 'Attempted to set Done without Final payment received.');
  }
}

/**
 * Detects if the SFID entered in a row is a duplicate of another row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} pipelineSheet The 'Project Pipeline' sheet object.
 * @param {number} row The row number that was edited.
 * @param {RowState} state The state object for the edited row.
 */
function detectDuplicateSFID_(pipelineSheet, row, state) {
  const sfid = state.getValue('SFID');
  if (!sfid) return;
  const lastRow = pipelineSheet.getLastRow();
  if (lastRow < 2) return;

  const sfidCol = state.getCol('SFID');
  // Get all SFIDs from the sheet to check for duplicates.
  const sfids = pipelineSheet.getRange(2, sfidCol, lastRow - 1, 1).getValues().flat();
  const isDup = sfids.filter(v => v === sfid).length > 1;

  state.setValue('Duplicate SFID', isDup);
  if (isDup) logOpsInbox_('duplicate_sfid', sfid, `Detected duplicate SFID: ${sfid}`);
}

/**
 * Clears the 'Blocked since' timestamp and any related notes.
 * @param {RowState} state The state object for the row.
 */
function clearBlocked_(state) {
  state.setValue('Blocked since', '');
  // Only clear the note if it's a block-related note.
  if ((state.getNote('Project Status') || '').startsWith('Advance blocked')) {
     state.setNote('Project Status', '');
  }
}

/**
 * Logs an issue to the 'Ops Inbox' sheet for manual review.
 * @param {string} type The type of issue (e.g., 'data_missing', 'script_error').
 * @param {string} sfid The Salesforce ID related to the issue, if any.
 * @param {string} details A description of the issue.
 */
function logOpsInbox_(type, sfid, details) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.OPS_INBOX);
    if (!sh) return;

    const headers = SHEET_SETUP_CONFIG.HEADERS[CFG.SHEETS.OPS_INBOX];
    if (!headers || headers.length === 0) return;

    // Construct the row data based on headers to ensure correct column order.
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

    // After logging to the sheet, send an email notification.
    sendErrorNotification_(type, sfid, details);
  } catch(err) {
    // Fallback to console logging if writing to the sheet fails.
    console.error(`Failed to write to Ops Inbox. Type: ${type}, SFID: ${sfid}, Details: ${details}. Error: ${err.message}`);
  }
}

/**
 * A custom function intended to be called from a spreadsheet formula.
 * It generates a human-readable reason why a project cannot advance.
 * @param {string} status The target status.
 * @param {boolean} canGlobal Global advancement check.
 * @param {boolean} canPerm Permitting advancement check.
 * @param {boolean} canSched Scheduled advancement check.
 * @param {boolean} canInspect Inspections advancement check.
 * @param {boolean} canDone Done advancement check.
 * @returns {string} A concatenated string of reasons for the block.
 * @customfunction
 */
function GET_ADVANCE_BLOCK_REASON(status, canGlobal, canPerm, canSched, canInspect, canDone) {
  const reasons = [];
  if (!toBool_(canGlobal)) reasons.push('Missing global data or overdue tasks or duplicate SFID');
  if (status === 'Permitting' && !toBool_(canPerm)) reasons.push('Gate to Permitting not met');
  if (status === 'Scheduled' && !toBool_(canSched)) reasons.push('Gate to Scheduled not met');
  if (status === 'Inspections' && !toBool_(canInspect)) reasons.push('Gate to Inspections not met');
  if (status === 'Done' && !toBool_(canDone)) reasons.push('Gate to Done not met');
  return reasons.join(' • ');
}

// --- Notification Functions --- //

/**
 * Sends an email notification to admins when an error is logged.
 * @param {string} type The type of error.
 * @param {string} sfid The SFID related to the error.
 * @param {string} details The detailed error message.
 * @private
 */
function sendErrorNotification_(type, sfid, details) {
  try {
    const admins = getAdmins_();
    if (!admins || admins.length === 0) {
      console.error("Cannot send error notification: No admin emails are configured.");
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const subject = `[Error] PM Automations Script: ${type}`;
    const body = `
      An error was logged in the Project Management Automations spreadsheet.

      Sheet Name: ${ss.getName()}
      Sheet URL: ${ss.getUrl()}

      Error Details:
      - Type: ${type}
      - Related SFID: ${sfid || 'N/A'}
      - Timestamp: ${new Date().toUTCString()}
      - Full Details: ${details}

      Please review the 'Ops Inbox' sheet for more information.
    `;

    MailApp.sendEmail(admins.join(','), subject, body.trim());

  } catch (e) {
    // Log a failure to send the email itself to the console.
    // We don't want to trigger another error log and create a loop.
    console.error(`Failed to send error notification email. Error: ${e.message}`);
  }
}

// --- Utility Functions --- //

/**
 * Creates a map of header names to their column numbers (1-indexed).
 * This is a crucial utility for making the script resilient to column reordering.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to get headers from.
 * @param {string[]} [requiredHeaders] An optional array of headers to validate.
 * @returns {Object|null} A map of headers to column numbers, or null if validation fails.
 */
function getHeaderMap_(sheet, requiredHeaders) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = { _width: headerRow.length };
  headerRow.forEach((header, index) => {
    if (header) map[header.toString().trim()] = index + 1;
  });

  // If required headers are specified, ensure they all exist.
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

/**
 * Sets a cell's value to the current date and time, but only if the cell is empty.
 * @param {RowState} state The state object for the row.
 * @param {string} headerName The header of the column to stamp.
 */
function stampIfEmpty_(state, headerName) {
  if (!state.getValue(headerName)) {
    state.setValue(headerName, new Date());
  }
}

/**
 * Checks if an override timestamp is still active (i.e., in the future).
 * @param {Date|string} overrideUntil The timestamp when the override expires.
 * @returns {boolean} True if the override is active.
 */
function isOverrideActive_(overrideUntil) {
  return overrideUntil && new Date(overrideUntil).getTime() >= Date.now();
}

/**
 * A robust utility to convert various "truthy" values to a boolean.
 * Handles boolean, string, and number types.
 * @param {*} v The value to convert.
 * @returns {boolean} The boolean representation of the value.
 */
function toBool_(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'string') return v.toLowerCase() === 'true';
  if (typeof v === 'number') return v !== 0;
  return !!v;
}

// --- State Management Class --- //

/**
 * A class to manage the state of a single row during an onEdit event.
 * This approach minimizes expensive spreadsheet read/write operations by batching
 * all changes and writing them back to the sheet only once, if needed.
 */
class RowState {
  /**
   * @param {Object} headerMap A map of header names to column numbers.
   * @param {Array<*>} initialValues The initial values of the row.
   * @param {Array<string>} initialNotes The initial notes of the row.
   */
  constructor(headerMap, initialValues, initialNotes) {
    this.headerMap = headerMap;
    this.initialValues = [...initialValues];
    this.finalValues = [...initialValues];
    this.initialNotes = [...initialNotes];
    this.finalNotes = [...initialNotes];
    // Create a reverse map for getting header names from column numbers.
    this.colToHeader = Object.fromEntries(Object.entries(headerMap).map(([k, v]) => [v, k]));
  }
  getHeader(col) { return this.colToHeader[col]; }
  getCol(header) { return this.headerMap[header]; }
  getValue(header) { return this.finalValues[this.getCol(header) - 1]; }
  getNote(header) { return this.finalNotes[this.getCol(header) - 1]; }
  setValue(header, value) { this.finalValues[this.getCol(header) - 1] = value; }
  setNote(header, note) { this.finalNotes[this.getCol(header) - 1] = note; }

  /**
   * Reverts a value back to its original state before the edit.
   * @param {Object} e The onEdit event object.
   * @param {string} header The header of the column to revert.
   */
  revertValue(e, header) {
    const oldValue = e.oldValue !== undefined ? e.oldValue : this.initialValues[this.getCol(header) - 1];
    this.setValue(header, oldValue);
  }
}

/**
 * Writes the final values and notes from a RowState object back to the spreadsheet.
 * It only performs a write operation if the values or notes have actually changed.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range object for the row.
 * @param {RowState} state The state object for the row.
 */
function commitRowState_(range, state) {
  // Only write values if they have changed.
  if (JSON.stringify(state.initialValues) !== JSON.stringify(state.finalValues)) {
    range.setValues([state.finalValues]);
  }
  // Only write notes if they have changed.
  if (JSON.stringify(state.initialNotes) !== JSON.stringify(state.finalNotes)) {
    range.setNotes([state.finalNotes]);
  }
}

// --- Admin and Properties Management --- //

/**
 * Retrieves the list of admin email addresses from Script Properties.
 * If no admins are set, it defaults to the effective user.
 * @returns {string[]} An array of admin email addresses.
 * @private
 */
function getAdmins_() {
  const properties = PropertiesService.getScriptProperties();
  let admins = properties.getProperty('OVERRIDE_ADMINS');
  if (!admins) {
    // Set a default admin if none exists.
    const defaultAdmin = Session.getEffectiveUser().getEmail() || 'jules@example.com';
    properties.setProperty('OVERRIDE_ADMINS', defaultAdmin);
    return [defaultAdmin];
  }
  return admins.split(',').map(s => s.trim()).filter(Boolean);
}

/**
 * Saves the list of admin email addresses to Script Properties.
 * @param {string[]} adminsArray An array of admin email addresses.
 * @private
 */
function setAdmins_(adminsArray) {
  const adminString = adminsArray.join(',');
  PropertiesService.getScriptProperties().setProperty('OVERRIDE_ADMINS', adminString);
}
