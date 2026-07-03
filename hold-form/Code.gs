// =============================================================================
// MATHNASIUM — STUDENT HOLD WORKFLOW
// Google Apps Script  |  One deployment per center account
// =============================================================================
//
// QUICK START:
//   1. Open Google Sheet in the center's Google account
//   2. Extensions > Apps Script — paste this file as Code.gs, then Save
//   3. Run  setupWorkbook()  from the ⚙ Hold Workflow menu (authorize when prompted)
//   4. Review and update the SETTINGS tab
//   5. Deploy as Web App:
//        Extensions > Apps Script > Deploy > New deployment
//        Type: Web app | Execute as: Me | Who has access: Anyone
//        Copy the deployment URL → paste into SETTINGS as WEB_APP_URL
//   6. Run  installTriggers()  from the ⚙ Hold Workflow menu
//   7. Share the sheet with Matt's account as Editor
//
// =============================================================================


// =============================================================================
// SECTION 0 — CONSTANTS
// =============================================================================

// ── Tab names ─────────────────────────────────────────────────────────────
const TAB_INPUT     = 'INPUT';
const TAB_MASTER    = 'MASTER_HOLDS';
const TAB_SETTINGS  = 'SETTINGS';
const TAB_DASHBOARD = 'DASHBOARD';
const TAB_ERROR_LOG = 'ERROR_LOG';

// ── Brand colors ──────────────────────────────────────────────────────────
const CLR_RED       = '#EE3124';  // Mathnasium red
const CLR_NAVY      = '#1C2B6E';  // Mathnasium navy
const CLR_RED_LIGHT = '#FDEDEC';  // light red for detail boxes
const CLR_RED_MUTED = '#C0392B';  // darker red for muted headers
const CLR_WHITE     = '#FFFFFF';

// ── INPUT tab column indices (0-based) ────────────────────────────────────
// CD edits columns A–L  (indices 0–11)
// Script auto-fills M–U (indices 12–20)  ← greyed out
const IC = {
  STUDENT_NAME:    0,   // A
  PARENT_NAME:     1,   // B
  PARENT_EMAIL:    2,   // C
  HOLD_START:      3,   // D
  EXPECTED_RETURN: 4,   // E
  BILLING_RESUME:  5,   // F
  NOTES:           6,   // G
  IN_CONTRACT:     7,   // H  ← YES/NO — is student still in initial contract?
  CHARGED_MONTHS:  8,   // I  ← free text, e.g. "June and July" (contract only)
  CREDIT_MONTHS:   9,   // J  ← free text, e.g. "August and September" (contract only)
  SEND_EMAIL:      10,  // K  ← YES to trigger
  RADIUS_UPDATED:  11,  // L  ← CD marks YES after updating Radius
  // ── auto-filled (greyed) ──────────────────────
  STATUS:          12,  // M
  FORM_LINK:       13,  // N
  EMAIL_SENT_TS:   14,  // O
  SIGNED:          15,  // P
  SIGNATURE_TS:    16,  // Q
  REMINDER_1:      17,  // R
  REMINDER_2:      18,  // S
  REMINDER_3:      19,  // T
  HOLD_ID:         20   // U
};

// ── MASTER_HOLDS column indices (0-based) ─────────────────────────────────
const MC = {
  HOLD_ID:          0,
  CENTER:           1,
  SOURCE_TAB:       2,
  SOURCE_ROW:       3,
  STUDENT_NAME:     4,
  PARENT_NAME:      5,
  PARENT_EMAIL:     6,
  HOLD_START:       7,
  EXPECTED_RETURN:  8,
  BILLING_RESUME:   9,
  NOTES:            10,
  IN_CONTRACT:      11,  // ← NEW
  CHARGED_MONTHS:   12,  // ← NEW
  CREDIT_MONTHS:    13,  // ← NEW
  FORM_LINK:        14,
  EMAIL_SENT:       15,
  EMAIL_SENT_TS:    16,
  SIGNED:           17,
  SIGNATURE_TS:     18,
  REMINDER_1:       19,
  REMINDER_2:       20,
  REMINDER_3:       21,
  RADIUS_UPDATED:   22,
  STATUS:           23
};

// ── Reminder timing (hours after initial email) ───────────────────────────
const REMINDER_1_HOURS = 24;   // 1 day
const REMINDER_2_HOURS = 72;   // 3 days
const REMINDER_3_HOURS = 168;  // 7 days (escalation — CCs admin)

// ── Rows to pre-format in INPUT ───────────────────────────────────────────
const INPUT_FORMAT_ROWS = 500;


// =============================================================================
// SECTION 1 — MENU
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙ Hold Workflow')
    .addItem('▶  Process New Rows (send emails)',  'processDirectorInputs')
    .addItem('🔔  Check Reminder Statuses',         'checkReminderStatuses')
    .addItem('📊  Refresh Dashboard',               'buildDashboard')
    .addSeparator()
    .addItem('🔧  Setup Workbook (first-time)',     'setupWorkbook')
    .addItem('🔄  Migrate Columns (run once after update)', 'migrateColumnsV2')
    .addItem('⏱  Install Triggers',               'installTriggers')
    .addItem('🗑  Remove All Triggers',             'removeTriggers')
    .addToUi();
}

/**
 * One-time migration: inserts the 3 new in-contract columns into existing sheets.
 * Run this ONCE after pasting the updated Code.gs into an existing workbook.
 * Safe to run only once — running it a second time will insert duplicate columns.
 */
function migrateColumnsV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── INPUT tab ──────────────────────────────────────────────────────────────
  // Insert 3 cols at H (col 8), pushing Send Email? and Radius Updated to K and L
  const inputSh = ss.getSheetByName(TAB_INPUT);
  if (inputSh) {
    // Safety check: abort if col H header is already "In Contract?"
    if (inputSh.getRange(1, 8).getValue() === 'In Contract?') {
      ui.alert('Already migrated', 'Column H is already "In Contract?" — migration was already run. No changes made.', ui.ButtonSet.OK);
      return;
    }
    inputSh.insertColumnsBefore(8, 3);
    inputSh.getRange(1, 8).setValue('In Contract?');
    inputSh.getRange(1, 9).setValue('Charged Month(s)');
    inputSh.getRange(1, 10).setValue('Credit Month(s)');
    formatInputTab(ss);
  }

  // ── MASTER_HOLDS tab ───────────────────────────────────────────────────────
  // Insert 3 cols at col 12 (after Notes at col 11)
  const masterSh = ss.getSheetByName(TAB_MASTER);
  if (masterSh) {
    masterSh.insertColumnsBefore(12, 3);
    masterSh.getRange(1, 12).setValue('In Contract?');
    masterSh.getRange(1, 13).setValue('Charged Month(s)');
    masterSh.getRange(1, 14).setValue('Credit Month(s)');
    formatMasterTab(ss);
  }

  ui.alert(
    '✓ Migration Complete',
    'New columns added:\n\n' +
    'INPUT tab:        In Contract? (H) | Charged Month(s) (I) | Credit Month(s) (J)\n' +
    'MASTER_HOLDS tab: In Contract? | Charged Month(s) | Credit Month(s) — after Notes\n\n' +
    'Next: Deploy > Manage deployments → New version → Deploy',
    ui.ButtonSet.OK
  );
}


// =============================================================================
// SECTION 2 — SETTINGS HELPERS
// =============================================================================

function getSettings() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_SETTINGS);
  if (!sheet) throw new Error('SETTINGS tab not found. Run setupWorkbook() first.');
  const data = sheet.getDataRange().getValues();
  const out  = {};
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0]).trim();
    if (k) out[k] = data[i][1];
  }
  return out;
}

function setSetting(key, value) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_SETTINGS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}


// =============================================================================
// SECTION 3 — SETUP
// =============================================================================

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  createTabIfNeeded(ss, TAB_INPUT,     getInputHeaders());
  createTabIfNeeded(ss, TAB_MASTER,    getMasterHeaders());
  createTabIfNeeded(ss, TAB_SETTINGS,  ['Key', 'Value']);
  createTabIfNeeded(ss, TAB_ERROR_LOG, ['Timestamp', 'Function', 'Error', 'Context']);
  createTabIfNeeded(ss, TAB_DASHBOARD, null);

  initSettingsDefaults(ss);
  formatInputTab(ss);
  formatMasterTab(ss);
  buildDashboard();

  ui.alert(
    '✓ Setup Complete',
    'Next steps:\n\n' +
    '1. Update SETTINGS tab (PHONE, CENTER_NAME, etc.)\n\n' +
    '2. Deploy as Web App:\n' +
    '   Extensions > Apps Script > Deploy > New deployment\n' +
    '   Type: Web app | Execute as: Me | Who has access: Anyone\n' +
    '   → Paste the URL into SETTINGS as WEB_APP_URL\n\n' +
    '3. ⚙ Hold Workflow > Install Triggers',
    ui.ButtonSet.OK
  );
}

function createTabIfNeeded(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground(CLR_RED)
        .setFontColor(CLR_WHITE);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getInputHeaders() {
  return [
    'Student Name',            // A  IC.STUDENT_NAME
    'Parent Name',             // B
    'Parent Email',            // C
    'Hold Start Date',         // D
    'Expected Return Date',    // E
    'Enrollment Resume Month', // F
    'Notes',                   // G
    'In Contract?',            // H  ← YES if student is within initial contract
    'Charged Month(s)',        // I  ← month(s) billed during hold (contract only)
    'Credit Month(s)',         // J  ← month(s) credited at contract end (contract only)
    'Send Email?',             // K  ← YES to trigger
    'Radius Updated',          // L  ← CD marks YES after Radius update
    // ── auto-filled ─────────────────────
    'Status',                  // M
    'Acknowledgment Link',     // N
    'Email Sent',              // O
    'Signed?',                 // P
    'Signed On',               // Q
    'Reminder 1 Sent',         // R
    'Reminder 2 Sent',         // S
    'Reminder 3 Sent',         // T
    'Hold ID'                  // U
  ];
}

function getMasterHeaders() {
  return [
    'Hold ID', 'Center', 'Source Tab', 'Source Row',
    'Student Name', 'Parent Name', 'Parent Email',
    'Hold Start Date', 'Expected Return Date', 'Enrollment Resume Month',
    'Notes', 'In Contract?', 'Charged Month(s)', 'Credit Month(s)',
    'Acknowledgment Link', 'Email Sent', 'Email Sent Date',
    'Signed?', 'Signed On',
    'Reminder 1 Sent', 'Reminder 2 Sent', 'Reminder 3 Sent',
    'Radius Updated', 'Status'
  ];
}

function initSettingsDefaults(ss) {
  const sheet = ss.getSheetByName(TAB_SETTINGS);
  if (sheet.getLastRow() > 1) return;

  const rows = [
    ['CENTER_NAME',  'Mathnasium of Teaneck'],
    ['CENTER_EMAIL', 'teaneck@mathnasium.com'],
    ['CD_EMAIL',     'teaneck@mathnasium.com'],
    ['FROM_NAME',    'Mathnasium of Teaneck'],
    ['PHONE',        '(201) XXX-XXXX'],
    ['ADMIN_CC',     'matt.diamond@mathnasium.com'],
    ['BCC_EMAIL',    'bccradius@mathnasium.com'],
    ['WEB_APP_URL',  '']
  ];
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  sheet.getRange(rows.length + 1, 1, 1, 2).setBackground('#FFF9C4'); // yellow highlight for WEB_APP_URL
}

/**
 * Formats INPUT tab:
 *  • A–L  (cols 1–12): white, CD-editable. Includes In Contract?, Charged/Credit months, Radius Updated.
 *  • M–U  (cols 13–21): grey, auto-filled by script.
 *  • Billing Resume Month column forced to plain text so Sheets never
 *    misreads "July 2026" as a date object.
 */
function formatInputTab(ss) {
  const sheet    = ss.getSheetByName(TAB_INPUT);
  if (!sheet) return;
  const dataRows = INPUT_FORMAT_ROWS;

  // CD-editable zone A–L: white
  sheet.getRange(2, 1, dataRows, 12).setBackground(CLR_WHITE).setFontColor('#222222');

  // Auto-fill zone M–U: grey
  sheet.getRange(2, 13, dataRows, 9).setBackground('#EFEFEF').setFontColor('#888888');

  // Muted header color for auto-fill columns M–U
  sheet.getRange(1, 13, 1, 9).setBackground(CLR_RED_MUTED).setFontColor(CLR_WHITE);

  // Thick border between L and M to mark the boundary
  sheet.getRange(1, 12, dataRows + 1, 1)
    .setBorder(null, null, null, true, null, null, CLR_RED, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Notes on auto-fill headers (M–U)
  const autoNotes = [
    'Auto-filled: Sent / Signed',
    'Auto-filled: acknowledgment page URL',
    'Auto-filled: date initial email sent',
    'Auto-filled: YES when parent signs',
    'Auto-filled: date parent signed',
    'Auto-filled: date Reminder 1 sent',
    'Auto-filled: date Reminder 2 sent',
    'Auto-filled: date Reminder 3 sent',
    'Auto-filled: unique Hold ID'
  ];
  autoNotes.forEach((note, i) => sheet.getRange(1, 13 + i).setNote(note));

  // Notes on CD-editable special headers
  sheet.getRange(1, IC.IN_CONTRACT + 1)
    .setNote('YES if student is still within their initial enrollment contract. Select NO for standard holds.');
  sheet.getRange(1, IC.CHARGED_MONTHS + 1)
    .setNote('Only used when In Contract? = YES. Enter the month(s) billing continues, e.g. "June and July".');
  sheet.getRange(1, IC.CREDIT_MONTHS + 1)
    .setNote('Only used when In Contract? = YES. Enter the month(s) credit will be applied, e.g. "August and September".');
  sheet.getRange(1, IC.RADIUS_UPDATED + 1)
    .setNote('Mark YES here after you have updated Radius for this hold.');

  // YES/NO dropdown on In Contract? (H)
  const yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['YES', 'NO'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, IC.IN_CONTRACT + 1, dataRows, 1).setDataValidation(yesNoRule);

  // YES/NO dropdown on Send Email? (K)
  const sendEmailRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['YES', 'NO'], true)
    .setAllowInvalid(false)
    .setHelpText('Select YES to send the acknowledgment email to the parent.')
    .build();
  sheet.getRange(2, IC.SEND_EMAIL + 1, dataRows, 1).setDataValidation(sendEmailRule);

  // YES/NO dropdown on Radius Updated (L)
  sheet.getRange(2, IC.RADIUS_UPDATED + 1, dataRows, 1).setDataValidation(yesNoRule);

  // Force Enrollment Resume Month (F) to plain text so Sheets never
  // converts "July 2026" into a Date object
  sheet.getRange(2, IC.BILLING_RESUME + 1, dataRows, 1).setNumberFormat('@');

  // Date-only format on date columns
  const dateFmt = 'M/d/yyyy';
  [IC.HOLD_START, IC.EXPECTED_RETURN, IC.EMAIL_SENT_TS,
   IC.SIGNATURE_TS, IC.REMINDER_1, IC.REMINDER_2, IC.REMINDER_3].forEach(col => {
    sheet.getRange(2, col + 1, dataRows, 1).setNumberFormat(dateFmt);
  });

  sheet.autoResizeColumns(1, 21);
}

function formatMasterTab(ss) {
  const sheet = ss.getSheetByName(TAB_MASTER);
  if (!sheet) return;
  const dateFmt = 'M/d/yyyy';
  [MC.HOLD_START, MC.EXPECTED_RETURN, MC.EMAIL_SENT_TS,
   MC.SIGNATURE_TS, MC.REMINDER_1, MC.REMINDER_2, MC.REMINDER_3].forEach(col => {
    sheet.getRange(2, col + 1, INPUT_FORMAT_ROWS, 1).setNumberFormat(dateFmt);
  });
  // Force billing column to plain text in master too
  sheet.getRange(2, MC.BILLING_RESUME + 1, INPUT_FORMAT_ROWS, 1).setNumberFormat('@');
}


// =============================================================================
// SECTION 4 — MAIN PROCESSING
// =============================================================================

function processDirectorInputs() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings();
  const inputSh  = ss.getSheetByName(TAB_INPUT);
  const masterSh = ss.getSheetByName(TAB_MASTER);

  if (!inputSh)  { logError('processDirectorInputs', 'INPUT tab missing', '');        return; }
  if (!masterSh) { logError('processDirectorInputs', 'MASTER_HOLDS tab missing', ''); return; }

  const webAppUrl = String(settings['WEB_APP_URL'] || '').trim();
  if (!webAppUrl) {
    logError('processDirectorInputs', 'WEB_APP_URL not set in SETTINGS',
             'Deploy the script as a Web App and paste the URL into SETTINGS.');
    return;
  }

  const lastRow = inputSh.getLastRow();
  if (lastRow < 2) return;

  const data = inputSh.getRange(2, 1, lastRow - 1, Object.keys(IC).length).getValues();

  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const sheetRow = i + 2;
    const sendEmail = String(row[IC.SEND_EMAIL]).trim().toUpperCase();
    const status    = String(row[IC.STATUS]).trim();

    if (sendEmail !== 'YES' || status !== '') continue;

    const studentName    = String(row[IC.STUDENT_NAME]).trim();
    const parentEmail    = String(row[IC.PARENT_EMAIL]).trim();
    const holdStart      = row[IC.HOLD_START];
    const expReturn      = row[IC.EXPECTED_RETURN];

    if (!studentName) {
      logError('processDirectorInputs', 'Missing Student Name',            `Row ${sheetRow}`); continue;
    }
    if (!parentEmail || !parentEmail.includes('@')) {
      logError('processDirectorInputs', 'Missing or invalid Parent Email', `Row ${sheetRow} — ${studentName}`); continue;
    }
    if (!holdStart) {
      logError('processDirectorInputs', 'Missing Hold Start Date',         `Row ${sheetRow} — ${studentName}`); continue;
    }
    if (!expReturn) {
      logError('processDirectorInputs', 'Missing Expected Return Date',    `Row ${sheetRow} — ${studentName}`); continue;
    }

    const inContract    = String(row[IC.IN_CONTRACT]).trim().toUpperCase();
    const chargedMonths = String(row[IC.CHARGED_MONTHS]).trim();
    const creditMonths  = String(row[IC.CREDIT_MONTHS]).trim();

    if (inContract !== 'YES' && inContract !== 'NO') {
      logError('processDirectorInputs', 'In Contract? must be YES or NO — please select from the dropdown', `Row ${sheetRow} — ${studentName}`); continue;
    }
    if (inContract === 'YES' && (!chargedMonths || !creditMonths)) {
      logError('processDirectorInputs', 'Charged Month(s) and Credit Month(s) are required when In Contract? = YES', `Row ${sheetRow} — ${studentName}`); continue;
    }

    try {
      const holdId   = generateHoldId(masterSh);
      const holdData = {
        holdId,
        center:         settings['CENTER_NAME'],
        sourceTab:      TAB_INPUT,
        sourceRow:      sheetRow,
        studentName,
        parentName:     String(row[IC.PARENT_NAME]).trim(),
        parentEmail,
        holdStart,
        expectedReturn: expReturn,
        billingResume:  String(row[IC.BILLING_RESUME]).trim(),
        notes:          String(row[IC.NOTES]).trim(),
        inContract,
        chargedMonths,
        creditMonths
      };

      holdData.ackLink = generateWebAppLink(holdId, webAppUrl);
      sendInitialEmail(holdData, settings);

      const now = new Date();
      appendToMaster(masterSh, holdData, now);

      inputSh.getRange(sheetRow, IC.HOLD_ID       + 1).setValue(holdId);
      inputSh.getRange(sheetRow, IC.STATUS        + 1).setValue('Sent');
      inputSh.getRange(sheetRow, IC.FORM_LINK     + 1).setValue(holdData.ackLink);
      inputSh.getRange(sheetRow, IC.EMAIL_SENT_TS + 1).setValue(now);

      Logger.log(`processDirectorInputs ✓  Row ${sheetRow}: ${studentName} — ${holdId}`);
    } catch (err) {
      logError('processDirectorInputs', err.message, `Row ${sheetRow} — ${studentName}`);
    }
  }
}

function generateHoldId(masterSh) {
  const tz     = Session.getScriptTimeZone();
  const today  = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const prefix = `HOLD-${today}-`;
  const data   = masterSh.getDataRange().getValues();
  let maxN     = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][MC.HOLD_ID]);
    if (id.startsWith(prefix)) {
      const n = parseInt(id.slice(prefix.length), 10);
      if (n > maxN) maxN = n;
    }
  }
  return prefix + String(maxN + 1).padStart(4, '0');
}

function appendToMaster(masterSh, holdData, timestamp) {
  masterSh.appendRow([
    holdData.holdId, holdData.center, holdData.sourceTab, holdData.sourceRow,
    holdData.studentName, holdData.parentName, holdData.parentEmail,
    holdData.holdStart, holdData.expectedReturn, holdData.billingResume,
    holdData.notes,
    holdData.inContract, holdData.chargedMonths, holdData.creditMonths,
    holdData.ackLink,
    'YES', timestamp,
    '', '', '', '', '', '', 'Sent'
  ]);
}


// =============================================================================
// SECTION 5 — WEB APP LINK GENERATION
// =============================================================================

function generateWebAppLink(holdId, webAppUrl) {
  return `${webAppUrl}?id=${encodeURIComponent(holdId)}`;
}


// =============================================================================
// SECTION 6 — EMAIL TEMPLATES
// ─────────────────────────────────────────────────────────────────────────────
// EDIT THIS SECTION to change email wording or layout.
//
// Available variables:
//   holdData.studentName, holdData.parentName, holdData.parentEmail
//   holdData.holdStart, holdData.expectedReturn, holdData.billingResume
//   holdData.notes, holdData.ackLink
//   settings['FROM_NAME'], settings['PHONE']
// =============================================================================

function getInitialEmailSubject(holdData) {
  return `Please Acknowledge Enrollment Hold Details — ${holdData.studentName}`;
}

function getInitialEmailBody(holdData, settings) {
  const firstName  = holdData.parentName.split(' ')[0];
  const isContract = holdData.inContract === 'YES';

  const notesRow = holdData.notes
    ? `<tr>
         <td style="padding:6px 20px 6px 0;font-weight:bold;color:#333;vertical-align:top;">Notes:</td>
         <td style="padding:6px 0;">${holdData.notes}</td>
       </tr>`
    : '';

  const billingTableRow = isContract
    ? ''
    : `<tr style="background:${CLR_RED_LIGHT};">
         <td style="padding:8px 16px;font-weight:bold;color:#333;border:1px solid #f5c6c2;">Enrollment Billing Resumes</td>
         <td style="padding:8px 16px;border:1px solid #f5c6c2;">${formatBillingMonth(holdData.billingResume)}</td>
       </tr>`;

  const billingNotice = isContract
    ? `<p style="background:${CLR_RED_LIGHT};border:1px solid #f5c6c2;border-radius:4px;
                padding:12px 16px;margin:0 0 24px;font-size:14px;line-height:1.6;">
         As <strong>${holdData.studentName}</strong> is currently within their enrollment contract,
         monthly billing will continue during the hold period. They will be charged for
         <strong>${holdData.chargedMonths}</strong>, and a corresponding credit will be applied to
         <strong>${holdData.creditMonths}</strong> at the end of the initial contract.
       </p>`
    : '';

  const disclaimerText = isContract
    ? `By signing, you acknowledge that monthly enrollment billing will continue during the hold period
       and that a credit will be applied to ${holdData.creditMonths} at the end of the initial contract.`
    : `By signing, you acknowledge that recurring enrollment billing will resume beginning
       with the return month listed above unless alternative arrangements are made with
       the center before billing is processed.`;

  return `
<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;line-height:1.6;">

  <!-- Header -->
  <div style="background-color:${CLR_RED};padding:20px 24px;border-radius:6px 6px 0 0;">
    <p style="color:${CLR_WHITE};font-size:18px;font-weight:bold;margin:0;">
      Enrollment Hold Acknowledgment
    </p>
  </div>

  <!-- Body -->
  <div style="border:1px solid #ddd;border-top:none;border-radius:0 0 6px 6px;padding:24px;">
    <p>Hi ${firstName},</p>

    <p>We have entered the following temporary hold details for
    <strong>${holdData.studentName}</strong>:</p>

    <table style="border-collapse:collapse;margin:16px 0 24px;width:100%;max-width:480px;">
      <tr style="background:${CLR_RED_LIGHT};">
        <td style="padding:8px 16px;font-weight:bold;color:#333;border:1px solid #f5c6c2;">Hold Start Date</td>
        <td style="padding:8px 16px;border:1px solid #f5c6c2;">${formatDateReadable(holdData.holdStart)}</td>
      </tr>
      <tr>
        <td style="padding:8px 16px;font-weight:bold;color:#333;border:1px solid #f5c6c2;">Expected Return Date</td>
        <td style="padding:8px 16px;border:1px solid #f5c6c2;">${formatDateReadable(holdData.expectedReturn)}</td>
      </tr>
      ${billingTableRow}
      ${notesRow}
    </table>

    ${billingNotice}

    <p>Please review the details above and click the button below to acknowledge and sign:</p>

    <p style="margin:24px 0;">
      <a href="${holdData.ackLink}"
         style="display:inline-block;background-color:${CLR_RED};color:${CLR_WHITE};
                padding:12px 28px;text-decoration:none;border-radius:4px;
                font-weight:bold;font-size:14px;">
        Review &amp; Acknowledge
      </a>
    </p>

    <p style="font-size:13px;color:#666;border-top:1px solid #eee;padding-top:14px;">
      ${disclaimerText}
    </p>

    <p>Thank you,<br>
    <strong>${settings['FROM_NAME']}</strong><br>
    ${settings['PHONE'] || ''}</p>
  </div>

</div>`;
}

function getReminderEmailSubject(holdData, reminderNum) {
  return `Reminder ${reminderNum}: Please Acknowledge Your Enrollment Hold — ${holdData.studentName}`;
}

function getReminderEmailBody(holdData, settings, reminderNum) {
  const firstName  = holdData.parentName.split(' ')[0];
  const isContract = holdData.inContract === 'YES';

  const notesRow = holdData.notes
    ? `<tr>
         <td style="padding:6px 20px 6px 0;font-weight:bold;color:#333;vertical-align:top;">Notes:</td>
         <td style="padding:6px 0;">${holdData.notes}</td>
       </tr>`
    : '';

  const billingTableRow = isContract
    ? ''
    : `<tr style="background:${CLR_RED_LIGHT};">
         <td style="padding:8px 16px;font-weight:bold;color:#333;border:1px solid #f5c6c2;">Enrollment Billing Resumes</td>
         <td style="padding:8px 16px;border:1px solid #f5c6c2;">${formatBillingMonth(holdData.billingResume)}</td>
       </tr>`;

  const billingNotice = isContract
    ? `<p style="background:${CLR_RED_LIGHT};border:1px solid #f5c6c2;border-radius:4px;
                padding:12px 16px;margin:0 0 24px;font-size:14px;line-height:1.6;">
         As <strong>${holdData.studentName}</strong> is currently within their enrollment contract,
         monthly billing will continue during the hold period. They will be charged for
         <strong>${holdData.chargedMonths}</strong>, and a corresponding credit will be applied to
         <strong>${holdData.creditMonths}</strong> at the end of the initial contract.
       </p>`
    : '';

  const disclaimerText = isContract
    ? `By signing, you acknowledge that monthly enrollment billing will continue during the hold period
       and that a credit will be applied to ${holdData.creditMonths} at the end of the initial contract.`
    : `By signing, you acknowledge that recurring enrollment billing will resume beginning
       with the return month listed above unless alternative arrangements are made with
       the center before billing is processed.`;

  const urgencyBlock = (reminderNum === 3)
    ? `<p style="color:${CLR_RED};font-weight:bold;border:2px solid ${CLR_RED};
                 padding:10px 14px;border-radius:4px;margin:16px 0;">
         This is your final reminder. Please complete the acknowledgment as soon as possible.
       </p>`
    : '';

  return `
<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;line-height:1.6;">

  <div style="background-color:${CLR_RED};padding:20px 24px;border-radius:6px 6px 0 0;">
    <p style="color:${CLR_WHITE};font-size:18px;font-weight:bold;margin:0;">
      Enrollment Hold Acknowledgment — Reminder ${reminderNum}
    </p>
  </div>

  <div style="border:1px solid #ddd;border-top:none;border-radius:0 0 6px 6px;padding:24px;">
    <p>Hi ${firstName},</p>

    <p>This is Reminder ${reminderNum} — we have not yet received your acknowledgment
    for the temporary hold on <strong>${holdData.studentName}'s</strong> account.</p>

    ${urgencyBlock}

    <table style="border-collapse:collapse;margin:16px 0 24px;width:100%;max-width:480px;">
      <tr style="background:${CLR_RED_LIGHT};">
        <td style="padding:8px 16px;font-weight:bold;color:#333;border:1px solid #f5c6c2;">Hold Start Date</td>
        <td style="padding:8px 16px;border:1px solid #f5c6c2;">${formatDateReadable(holdData.holdStart)}</td>
      </tr>
      <tr>
        <td style="padding:8px 16px;font-weight:bold;color:#333;border:1px solid #f5c6c2;">Expected Return Date</td>
        <td style="padding:8px 16px;border:1px solid #f5c6c2;">${formatDateReadable(holdData.expectedReturn)}</td>
      </tr>
      ${billingTableRow}
      ${notesRow}
    </table>

    ${billingNotice}

    <p>Please click the button below to acknowledge and sign:</p>

    <p style="margin:24px 0;">
      <a href="${holdData.ackLink}"
         style="display:inline-block;background-color:${CLR_RED};color:${CLR_WHITE};
                padding:12px 28px;text-decoration:none;border-radius:4px;
                font-weight:bold;font-size:14px;">
        Review &amp; Acknowledge
      </a>
    </p>

    <p style="font-size:13px;color:#666;border-top:1px solid #eee;padding-top:14px;">
      ${disclaimerText}
    </p>

    <p>Thank you,<br>
    <strong>${settings['FROM_NAME']}</strong><br>
    ${settings['PHONE'] || ''}</p>
  </div>

</div>`;
}


// =============================================================================
// SECTION 7 — EMAIL SENDING
// =============================================================================

function sendInitialEmail(holdData, settings) {
  GmailApp.sendEmail(
    holdData.parentEmail,
    getInitialEmailSubject(holdData),
    '',
    {
      htmlBody: getInitialEmailBody(holdData, settings),
      name:     settings['FROM_NAME'] || 'Mathnasium',
      replyTo:  settings['CENTER_EMAIL'] || '',
      bcc:      settings['BCC_EMAIL'] || ''
    }
  );
  Logger.log(`sendInitialEmail → ${holdData.parentEmail}  (${holdData.studentName})`);
}

function sendReminderEmail(holdData, settings, reminderNum) {
  const cdEmail    = settings['CD_EMAIL']  || '';
  const bccEmail   = settings['BCC_EMAIL'] || '';
  const adminEmail = settings['ADMIN_CC']  || '';
  let cc = cdEmail;
  if (reminderNum === 3 && adminEmail) cc = [cdEmail, adminEmail].filter(Boolean).join(',');

  GmailApp.sendEmail(
    holdData.parentEmail,
    getReminderEmailSubject(holdData, reminderNum),
    '',
    {
      htmlBody: getReminderEmailBody(holdData, settings, reminderNum),
      name:     settings['FROM_NAME'] || 'Mathnasium',
      replyTo:  settings['CENTER_EMAIL'] || '',
      cc,
      bcc:      bccEmail
    }
  );
  Logger.log(`sendReminderEmail (${reminderNum}) → ${holdData.parentEmail}  (${holdData.studentName})`);
}


// =============================================================================
// SECTION 8 — REMINDER CHECKER
// =============================================================================

function checkReminderStatuses() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const masterSh = ss.getSheetByName(TAB_MASTER);
  const settings = getSettings();
  const now      = new Date();

  if (!masterSh || masterSh.getLastRow() < 2) return;

  // Hold IDs still present in INPUT — deleted rows are treated as cancelled
  const activeIds = getActiveInputHoldIds(ss);

  const data = masterSh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const signed = String(row[MC.SIGNED]).trim().toUpperCase();
    const sentTs = row[MC.EMAIL_SENT_TS];

    if (signed === 'YES' || !sentTs) continue;
    if (!activeIds.has(String(row[MC.HOLD_ID]).trim())) continue; // deleted from INPUT

    const hrs = (now - new Date(sentTs)) / 3600000;
    const r1  = row[MC.REMINDER_1];
    const r2  = row[MC.REMINDER_2];
    const r3  = row[MC.REMINDER_3];
    const hd  = masterRowToHoldData(row);
    let num   = null;

    if      (!r1 && hrs >= REMINDER_1_HOURS)         num = 1;
    else if (r1 && !r2 && hrs >= REMINDER_2_HOURS)   num = 2;
    else if (r2 && !r3 && hrs >= REMINDER_3_HOURS)   num = 3;
    if (!num) continue;

    try {
      sendReminderEmail(hd, settings, num);
      const col = [null, MC.REMINDER_1, MC.REMINDER_2, MC.REMINDER_3][num];
      masterSh.getRange(i + 1, col + 1).setValue(now);
      syncReminderToInput(ss, hd.holdId, ['', 'REMINDER_1', 'REMINDER_2', 'REMINDER_3'][num], now);
    } catch (err) {
      logError('checkReminderStatuses', err.message, `Hold ID: ${hd.holdId}`);
    }
  }
}

function getActiveInputHoldIds(ss) {
  const inputSh = ss.getSheetByName(TAB_INPUT);
  if (!inputSh || inputSh.getLastRow() < 2) return new Set();
  const ids = inputSh.getRange(2, IC.HOLD_ID + 1, inputSh.getLastRow() - 1, 1).getValues();
  return new Set(ids.map(r => String(r[0]).trim()).filter(Boolean));
}

function masterRowToHoldData(row) {
  return {
    holdId:         row[MC.HOLD_ID],
    center:         row[MC.CENTER],
    sourceTab:      row[MC.SOURCE_TAB],
    sourceRow:      row[MC.SOURCE_ROW],
    studentName:    row[MC.STUDENT_NAME],
    parentName:     row[MC.PARENT_NAME],
    parentEmail:    row[MC.PARENT_EMAIL],
    holdStart:      row[MC.HOLD_START],
    expectedReturn: row[MC.EXPECTED_RETURN],
    billingResume:  row[MC.BILLING_RESUME],
    notes:          String(row[MC.NOTES] || ''),
    inContract:     String(row[MC.IN_CONTRACT] || '').toUpperCase(),
    chargedMonths:  String(row[MC.CHARGED_MONTHS] || ''),
    creditMonths:   String(row[MC.CREDIT_MONTHS] || ''),
    ackLink:        row[MC.FORM_LINK]
  };
}


// =============================================================================
// SECTION 9 — WEB APP HANDLER
// =============================================================================

function doGet(e) {
  const holdId = (e && e.parameter && e.parameter.id) ? e.parameter.id.trim() : '';

  if (!holdId) {
    return HtmlService.createHtmlOutput(
      `<body style="font-family:Arial;padding:40px;">
         <p>No hold ID provided. Please use the link from your email.</p>
       </body>`
    );
  }

  const hold = getHoldById(holdId);

  if (!hold) {
    return HtmlService.createHtmlOutput(
      `<body style="font-family:Arial;padding:40px;">
         <p>Hold not found. Please contact the center directly.</p>
       </body>`
    );
  }

  if (String(hold.signed || '').toUpperCase() === 'YES') {
    return HtmlService.createHtmlOutput(
      `<body style="font-family:Arial;padding:40px;max-width:500px;margin:auto;">
         <h2 style="color:#2E7D32;">✓ Already Acknowledged</h2>
         <p>This hold for <strong>${hold.studentName}</strong> has already been acknowledged. Thank you!</p>
       </body>`
    );
  }

  const tz       = Session.getScriptTimeZone();
  const template = HtmlService.createTemplate(getAcknowledgmentPageHtml());
  template.holdId        = holdId;
  template.studentName   = hold.studentName;
  template.centerName    = hold.center || 'Mathnasium';
  template.startDate     = formatDateReadable(hold.holdStart);
  template.returnDate    = formatDateReadable(hold.expectedReturn);
  template.billingResume = formatBillingMonth(hold.billingResume);
  template.notes         = hold.notes || '';
  template.inContract    = hold.inContract || '';
  template.chargedMonths = hold.chargedMonths || '';
  template.creditMonths  = hold.creditMonths || '';
  template.today         = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  return template.evaluate()
    .setTitle('Hold Acknowledgment — ' + hold.studentName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processWebAcknowledgment(holdId, signature, dateStr) {
  try {
    if (!holdId || !signature) return { error: 'Missing required fields.' };

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const masterSh = ss.getSheetByName(TAB_MASTER);
    const data     = masterSh.getDataRange().getValues();
    const now      = new Date();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][MC.HOLD_ID]).trim() !== holdId) continue;
      masterSh.getRange(i + 1, MC.SIGNED       + 1).setValue('YES');
      masterSh.getRange(i + 1, MC.SIGNATURE_TS + 1).setValue(now);
      masterSh.getRange(i + 1, MC.STATUS       + 1).setValue('Signed');
      syncStatusBackToInputTab(ss, holdId, 'YES', now);
      Logger.log(`processWebAcknowledgment ✓  ${holdId}  sig: ${signature}`);
      return { success: true };
    }

    logError('processWebAcknowledgment', `Hold ID not found: ${holdId}`, '');
    return { error: 'Hold not found. Please contact the center.' };
  } catch (err) {
    logError('processWebAcknowledgment', err.message, `Hold ID: ${holdId}`);
    return { error: err.message };
  }
}

function getHoldById(holdId) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const masterSh = ss.getSheetByName(TAB_MASTER);
  if (!masterSh || masterSh.getLastRow() < 2) return null;
  const data = masterSh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][MC.HOLD_ID]).trim() !== holdId) continue;
    const row = data[i];
    return {
      holdId:         String(row[MC.HOLD_ID]),
      center:         String(row[MC.CENTER]),
      studentName:    String(row[MC.STUDENT_NAME]),
      parentName:     String(row[MC.PARENT_NAME]),
      parentEmail:    String(row[MC.PARENT_EMAIL]),
      holdStart:      row[MC.HOLD_START],
      expectedReturn: row[MC.EXPECTED_RETURN],
      billingResume:  String(row[MC.BILLING_RESUME] || ''),
      notes:          String(row[MC.NOTES] || ''),
      inContract:     String(row[MC.IN_CONTRACT] || '').toUpperCase(),
      chargedMonths:  String(row[MC.CHARGED_MONTHS] || ''),
      creditMonths:   String(row[MC.CREDIT_MONTHS] || ''),
      signed:         String(row[MC.SIGNED] || '')
    };
  }
  return null;
}


// =============================================================================
// SECTION 10 — ACKNOWLEDGMENT PAGE HTML
// Edit the styles or copy below to change the look of the parent-facing page.
// Colors use the Mathnasium brand palette defined in Section 0.
// =============================================================================

function getAcknowledgmentPageHtml() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Hold Acknowledgment</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; background: #f5f5f5; color: #222; }
    .wrapper { max-width: 600px; margin: 32px auto; padding: 0 16px 48px; }

    /* Header — Mathnasium red */
    .header {
      background: #EE3124;
      color: #fff;
      padding: 20px 24px;
      border-radius: 8px 8px 0 0;
    }
    .header .center-name { font-size: 13px; opacity: 0.85; margin-bottom: 4px; }
    .header h1 { font-size: 20px; font-weight: bold; }

    /* Card */
    .card {
      background: #fff;
      border-radius: 0 0 8px 8px;
      padding: 28px;
      box-shadow: 0 2px 8px rgba(0,0,0,.08);
    }

    /* Details table */
    .details-box {
      border: 1px solid #f5c6c2;
      border-radius: 6px;
      overflow: hidden;
      margin-bottom: 20px;
    }
    .details-box table { width: 100%; border-collapse: collapse; }
    .details-box td {
      padding: 9px 16px;
      border-bottom: 1px solid #f5c6c2;
      font-size: 14px;
      vertical-align: top;
    }
    .details-box tr:last-child td { border-bottom: none; }
    .details-box .label { font-weight: bold; color: #333; width: 45%; background: #FDEDEC; }
    .details-box .value { background: #fff; }

    /* Notes */
    .notes-box {
      background: #fff8e1;
      border-left: 4px solid #EE3124;
      padding: 12px 16px;
      margin-bottom: 20px;
      border-radius: 0 6px 6px 0;
      font-size: 14px;
    }

    hr { border: none; border-top: 1px solid #eee; margin: 22px 0; }
    .instruction { font-size: 13px; color: #555; margin-bottom: 20px; }

    /* Acknowledgment checkbox */
    .ack-block { display: flex; gap: 12px; margin-bottom: 24px; align-items: flex-start; }
    .ack-block input[type=checkbox] {
      width: 20px; height: 20px; margin-top: 2px; flex-shrink: 0;
      accent-color: #EE3124; cursor: pointer;
    }
    .ack-block label { font-size: 14px; line-height: 1.6; cursor: pointer; }

    /* Fields */
    .field { margin-bottom: 20px; }
    .field label { display: block; font-weight: bold; font-size: 14px; margin-bottom: 6px; }
    .field input {
      width: 100%; padding: 10px 12px; font-size: 14px;
      border: 1px solid #ccc; border-radius: 4px;
    }
    .field input:focus { outline: 2px solid #EE3124; border-color: #EE3124; }

    /* Submit button — Mathnasium red */
    .submit-btn {
      width: 100%; padding: 14px; font-size: 16px; font-weight: bold;
      background: #EE3124; color: #fff;
      border: none; border-radius: 4px; cursor: pointer; margin-top: 4px;
    }
    .submit-btn:hover    { background: #C9291E; }
    .submit-btn:disabled { background: #aaa; cursor: not-allowed; }

    .error-msg { color: #EE3124; font-size: 13px; margin-top: 10px; display: none; }
    .spinner   { text-align: center; color: #666; padding: 10px; display: none; }

    /* Success state */
    .success-box {
      background: #e8f5e9; border: 1px solid #a5d6a7;
      border-radius: 8px; padding: 32px; text-align: center;
    }
    .success-box h2 { color: #2E7D32; font-size: 22px; margin-bottom: 12px; }
    .success-box p  { color: #444; font-size: 14px; line-height: 1.7; }
  </style>
</head>
<body>
<div class="wrapper">

  <div class="header">
    <div class="center-name"><?= centerName ?></div>
    <h1>Enrollment Hold Acknowledgment<br><?= studentName ?></h1>
  </div>

  <div class="card" id="mainCard">

    <!-- Hold details — read-only -->
    <div class="details-box">
      <table>
        <tr>
          <td class="label">Hold Start Date</td>
          <td class="value"><?= startDate ?></td>
        </tr>
        <tr>
          <td class="label">Expected Return Date</td>
          <td class="value"><?= returnDate ?></td>
        </tr>
        <? if (inContract !== 'YES') { ?>
        <tr>
          <td class="label">Enrollment Billing Resumes</td>
          <td class="value"><?= billingResume ?></td>
        </tr>
        <? } ?>
      </table>
    </div>

    <? if (inContract === 'YES') { ?>
    <div style="background:#FDEDEC;border:1px solid #f5c6c2;border-radius:4px;
                padding:12px 16px;margin-bottom:20px;font-size:14px;line-height:1.6;">
      As <strong><?= studentName ?></strong> is currently within their enrollment contract,
      monthly billing will continue during the hold period. They will be charged for
      <strong><?= chargedMonths ?></strong>, and a corresponding credit will be applied to
      <strong><?= creditMonths ?></strong> at the end of the initial contract.
    </div>
    <? } ?>

    <? if (notes) { ?>
    <div class="notes-box"><strong>Notes:</strong> <?= notes ?></div>
    <? } ?>

    <hr>
    <p class="instruction">
      Please review the details above, then complete the acknowledgment below.
    </p>

    <!-- Acknowledgment -->
    <div class="ack-block">
      <input type="checkbox" id="ackCheck">
      <? if (inContract === 'YES') { ?>
      <label for="ackCheck">
        I understand that <?= studentName ?>'s monthly enrollment fee will continue
        during the hold period, and that a credit will be applied to
        <?= creditMonths ?> at the end of the initial contract.
      </label>
      <? } else { ?>
      <label for="ackCheck">
        I acknowledge the hold details listed above and understand that recurring
        enrollment billing will resume beginning with the return month listed above,
        unless alternative arrangements are made with the center before billing
        is processed.
      </label>
      <? } ?>
    </div>

    <div class="field">
      <label for="sigField">Parent Full Name <span style="font-weight:normal;">(Typed Signature)</span></label>
      <input type="text" id="sigField" placeholder="Type your full name">
    </div>

    <div class="field">
      <label for="dateField">Date</label>
      <input type="date" id="dateField" value="<?= today ?>">
    </div>

    <button class="submit-btn" id="submitBtn" onclick="submitForm()">
      Submit Acknowledgment
    </button>
    <div class="spinner" id="spinner">Submitting…</div>
    <div class="error-msg" id="errMsg"></div>

  </div><!-- /card -->
</div><!-- /wrapper -->

<script>
  var HOLD_ID = '<?= holdId ?>';

  function submitForm() {
    var ack  = document.getElementById('ackCheck').checked;
    var sig  = document.getElementById('sigField').value.trim();
    var date = document.getElementById('dateField').value;
    hideError();
    if (!ack)  { showError('Please check the acknowledgment box.');       return; }
    if (!sig)  { showError('Please type your full name as a signature.'); return; }
    if (!date) { showError("Please enter today's date.");                 return; }

    document.getElementById('submitBtn').disabled = true;
    document.getElementById('spinner').style.display = 'block';

    google.script.run
      .withSuccessHandler(onSuccess)
      .withFailureHandler(onFailure)
      .processWebAcknowledgment(HOLD_ID, sig, date);
  }

  function onSuccess(result) {
    document.getElementById('spinner').style.display = 'none';
    if (result && result.error) { onFailure(result.error); return; }
    document.getElementById('mainCard').innerHTML =
      '<div class="success-box">' +
        '<h2>&#10003; Acknowledgment Received</h2>' +
        '<p>Thank you — your acknowledgment has been recorded.<br>' +
        'Enrollment billing will resume as indicated unless you contact us before that date.</p>' +
      '</div>';
  }

  function onFailure(err) {
    document.getElementById('submitBtn').disabled = false;
    document.getElementById('spinner').style.display = 'none';
    showError('Something went wrong. Please try again or contact the center directly.');
    console.error(err);
  }

  function showError(msg) {
    var el = document.getElementById('errMsg');
    el.textContent = msg; el.style.display = 'block';
  }
  function hideError() {
    document.getElementById('errMsg').style.display = 'none';
  }
</script>
</body>
</html>`;
}


// =============================================================================
// SECTION 11 — STATUS SYNC
// =============================================================================

function syncStatusBackToInputTab(ss, holdId, signed, signatureTs) {
  const inputSh = ss.getSheetByName(TAB_INPUT);
  if (!inputSh) { logError('syncStatusBackToInputTab', 'INPUT tab missing', holdId); return; }
  const lastRow = inputSh.getLastRow();
  if (lastRow < 2) return;
  const holdIds = inputSh.getRange(2, IC.HOLD_ID + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < holdIds.length; i++) {
    if (String(holdIds[i][0]).trim() !== holdId) continue;
    const row = i + 2;
    inputSh.getRange(row, IC.SIGNED       + 1).setValue(signed);
    inputSh.getRange(row, IC.SIGNATURE_TS + 1).setValue(signatureTs);
    inputSh.getRange(row, IC.STATUS       + 1).setValue('Signed');
    Logger.log(`syncStatusBackToInputTab ✓  Row ${row}  ${holdId}`);
    return;
  }
  logError('syncStatusBackToInputTab', `Hold ID not found in INPUT: ${holdId}`, '');
}

function syncReminderToInput(ss, holdId, reminderKey, timestamp) {
  const colMap = { REMINDER_1: IC.REMINDER_1, REMINDER_2: IC.REMINDER_2, REMINDER_3: IC.REMINDER_3 };
  const col    = colMap[reminderKey];
  if (col === undefined) return;
  const inputSh = ss.getSheetByName(TAB_INPUT);
  if (!inputSh) return;
  const lastRow = inputSh.getLastRow();
  if (lastRow < 2) return;
  const holdIds = inputSh.getRange(2, IC.HOLD_ID + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < holdIds.length; i++) {
    if (String(holdIds[i][0]).trim() === holdId) {
      inputSh.getRange(i + 2, col + 1).setValue(timestamp); return;
    }
  }
}


// =============================================================================
// SECTION 12 — DASHBOARD
// =============================================================================

function buildDashboard() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const settings   = getSettings();
  const centerName = settings['CENTER_NAME'] || 'Mathnasium';
  const now        = new Date();

  let dashSh = ss.getSheetByName(TAB_DASHBOARD);
  if (!dashSh) dashSh = ss.insertSheet(TAB_DASHBOARD);
  dashSh.clearContents();
  dashSh.clearFormats();

  let r = 1;

  function titleRow(text) {
    dashSh.getRange(r, 1, 1, 6).merge()
      .setValue(text).setFontWeight('bold').setFontSize(13)
      .setBackground(CLR_NAVY).setFontColor(CLR_WHITE);
    r++;
  }
  function sectionHeader(text) {
    dashSh.getRange(r, 1, 1, 6).merge()
      .setValue(text).setFontWeight('bold').setFontSize(11)
      .setBackground(CLR_RED).setFontColor(CLR_WHITE);
    r++;
  }
  function colHeaders(headers) {
    dashSh.getRange(r, 1, 1, headers.length)
      .setValues([headers]).setFontWeight('bold')
      .setBackground(CLR_RED_LIGHT).setFontColor('#333');
    r++;
  }
  function dataRow(values) {
    dashSh.getRange(r, 1, 1, values.length).setValues([values]); r++;
  }
  function emptyRow(msg, color) {
    dashSh.getRange(r, 1).setValue(msg)
      .setFontColor(color || '#888').setFontStyle('italic'); r++;
  }

  titleRow(`${centerName} — Hold Status Dashboard`);
  dashSh.getRange(r, 1).setValue(`Last refreshed: ${formatDateReadable(now)}`)
    .setFontStyle('italic').setFontColor('#888');
  r += 2;

  const masterSh = ss.getSheetByName(TAB_MASTER);
  if (!masterSh || masterSh.getLastRow() < 2) {
    emptyRow('No hold records yet.', '#444');
    dashSh.autoResizeColumns(1, 6);
    return;
  }

  const rows = masterSh.getDataRange().getValues().slice(1);

  // 1. Unsigned
  const unsigned = rows.filter(r =>
    String(r[MC.SIGNED]).toUpperCase() !== 'YES' && String(r[MC.STATUS]) === 'Sent'
  );
  sectionHeader(`⚠  Unsigned Acknowledgments (${unsigned.length})`);
  if (unsigned.length) {
    colHeaders(['Hold ID', 'Student', 'Parent Email', 'Email Sent', 'Reminder 1', 'Reminder 2']);
    unsigned.forEach(row => dataRow([
      row[MC.HOLD_ID], row[MC.STUDENT_NAME], row[MC.PARENT_EMAIL],
      formatDateReadable(row[MC.EMAIL_SENT_TS]),
      row[MC.REMINDER_1] ? formatDateReadable(row[MC.REMINDER_1]) : '—',
      row[MC.REMINDER_2] ? formatDateReadable(row[MC.REMINDER_2]) : '—'
    ]));
  } else { emptyRow('All acknowledgments received ✓', '#2E7D32'); }
  r++;

  // 2. Starting soon
  const startingSoon = rows.filter(row => {
    const d = (new Date(row[MC.HOLD_START]) - now) / 86400000;
    return d >= 0 && d <= 14;
  });
  sectionHeader(`📅  Holds Starting Soon — Next 14 Days (${startingSoon.length})`);
  if (startingSoon.length) {
    colHeaders(['Hold ID', 'Student', 'Hold Start', 'Expected Return', 'Signed?', 'Status']);
    startingSoon.forEach(row => dataRow([
      row[MC.HOLD_ID], row[MC.STUDENT_NAME],
      formatDateReadable(row[MC.HOLD_START]), formatDateReadable(row[MC.EXPECTED_RETURN]),
      row[MC.SIGNED], row[MC.STATUS]
    ]));
  } else { emptyRow('None'); }
  r++;

  // 3. Ending soon
  const endingSoon = rows.filter(row => {
    const d = (new Date(row[MC.EXPECTED_RETURN]) - now) / 86400000;
    return d >= 0 && d <= 14;
  });
  sectionHeader(`🔔  Holds Ending Soon — Next 14 Days (${endingSoon.length})`);
  if (endingSoon.length) {
    colHeaders(['Hold ID', 'Student', 'Expected Return', 'Enrollment Billing Resumes', 'Signed?', 'Radius Updated']);
    endingSoon.forEach(row => dataRow([
      row[MC.HOLD_ID], row[MC.STUDENT_NAME],
      formatDateReadable(row[MC.EXPECTED_RETURN]),
      formatBillingMonth(row[MC.BILLING_RESUME]),
      row[MC.SIGNED], row[MC.RADIUS_UPDATED] || '—'
    ]));
  } else { emptyRow('None'); }
  r++;

  // 4. Signed / Radius not updated
  const needsRadius = rows.filter(row =>
    String(row[MC.SIGNED]).toUpperCase() === 'YES' && !row[MC.RADIUS_UPDATED]
  );
  sectionHeader(`📋  Signed — Radius Not Yet Updated (${needsRadius.length})`);
  if (needsRadius.length) {
    colHeaders(['Hold ID', 'Student', 'Signed On', 'Hold Start', 'Expected Return', 'Enrollment Billing Resumes']);
    needsRadius.forEach(row => dataRow([
      row[MC.HOLD_ID], row[MC.STUDENT_NAME],
      formatDateReadable(row[MC.SIGNATURE_TS]),
      formatDateReadable(row[MC.HOLD_START]),
      formatDateReadable(row[MC.EXPECTED_RETURN]),
      formatBillingMonth(row[MC.BILLING_RESUME])
    ]));
  } else { emptyRow('All signed holds updated in Radius ✓', '#2E7D32'); }
  r++;

  // 5. Summary
  sectionHeader('📊  Summary');
  const total   = rows.length;
  const signedN = rows.filter(r => String(r[MC.SIGNED]).toUpperCase() === 'YES').length;
  const activeN = rows.filter(r =>
    new Date(r[MC.HOLD_START]) <= now && now <= new Date(r[MC.EXPECTED_RETURN])
  ).length;
  [['Total Holds', total], ['Signed', signedN],
   ['Awaiting Signature', total - signedN], ['Currently Active', activeN]].forEach(([label, val]) => {
    dashSh.getRange(r, 1).setValue(label).setFontWeight('bold');
    dashSh.getRange(r, 2).setValue(val); r++;
  });

  dashSh.autoResizeColumns(1, 6);
}


// =============================================================================
// SECTION 13 — TRIGGERS
// =============================================================================

function installTriggers() {
  removeTriggers();
  ScriptApp.newTrigger('processDirectorInputs').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('checkReminderStatuses').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('buildDashboard').timeBased().everyDays(1).atHour(8).create();

  SpreadsheetApp.getUi().alert(
    '✓ Triggers Installed',
    '• Process new rows — every hour\n' +
    '• Check reminders  — daily at 9am\n' +
    '• Refresh dashboard — daily at 8am\n\n' +
    'View: Extensions > Apps Script > Triggers (clock icon)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}


// =============================================================================
// SECTION 14 — UTILITIES
// =============================================================================

/**
 * Formats any date value as "Month D, YYYY" (e.g. "June 1, 2026").
 * Handles Date objects, date strings, and serial numbers from Sheets.
 */
function formatDateReadable(date) {
  if (!date) return '';
  try {
    const d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d.getTime())) return String(date);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  } catch (_) { return String(date); }
}

/**
 * Formats the Enrollment Resume Month field.
 *
 * Google Sheets sometimes auto-detects "July 2026" as a Date object.
 * This function safely converts it back to a readable month string.
 * If the value is already a plain string (correctly entered), it passes through.
 */
function formatBillingMonth(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'MMMM yyyy');
  }
  const str = String(value).trim();
  // If it looks like it might have been serialized from a Date, try parsing
  if (/^\w+ \d{4}$/.test(str) || str === '') return str; // already clean
  const d = new Date(str);
  if (!isNaN(d.getTime()) && str.match(/\d{4}/)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM yyyy');
  }
  return str;
}


// =============================================================================
// SECTION 15 — ERROR LOGGING
// =============================================================================

function logError(functionName, errorMessage, context) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSh = ss.getSheetByName(TAB_ERROR_LOG);
    if (!logSh) {
      logSh = ss.insertSheet(TAB_ERROR_LOG);
      logSh.getRange(1, 1, 1, 4)
        .setValues([['Timestamp', 'Function', 'Error', 'Context']])
        .setFontWeight('bold').setBackground(CLR_RED).setFontColor(CLR_WHITE);
    }
    logSh.appendRow([new Date(), functionName, errorMessage, String(context)]);
    Logger.log(`[ERROR] ${functionName}: ${errorMessage} | ${context}`);
  } catch (e) {
    Logger.log(`logError failed: ${e.message}`);
  }
}
