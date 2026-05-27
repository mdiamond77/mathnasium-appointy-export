# Hold Form — In-Contract Student Handling — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add In Contract?, Charged Month(s), and Credit Month(s) fields to the hold workflow so that in-contract students receive a different email and acknowledgment page explaining billing continues with a credit at contract end.

**Architecture:** All changes are in a single file (`hold-form/Code.gs`). Tasks follow the file top-to-bottom: constants first, then headers, then formatting, then processing, then email templates, then web app. Each task shows the exact replacement code. After all code tasks, one manual deployment task redeploys the Web App.

**Tech Stack:** Google Apps Script, Google Sheets, GmailApp, HtmlService (scriptlet templates)

---

### Task 1: Update IC and MC constants

**Files:**
- Modify: `hold-form/Code.gs` (Section 0, lines ~42–87)

Three new INPUT columns are inserted at H/I/J (indices 7/8/9). Send Email? shifts from index 7→10, Radius Updated from 8→11, and all auto-fill columns shift +3. MASTER_HOLDS gets three new fields after NOTES (index 10).

- [ ] **Step 1: Replace the IC block**

Find the entire `const IC = { ... };` block and replace it with:

```javascript
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
```

- [ ] **Step 2: Replace the MC block**

Find the entire `const MC = { ... };` block and replace it with:

```javascript
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
```

- [ ] **Step 3: Save** (Ctrl+S in Apps Script editor)

---

### Task 2: Update INPUT and MASTER_HOLDS headers

**Files:**
- Modify: `hold-form/Code.gs` (functions `getInputHeaders`, `getMasterHeaders`)

- [ ] **Step 1: Replace `getInputHeaders()`**

```javascript
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
```

- [ ] **Step 2: Replace `getMasterHeaders()`**

```javascript
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
```

- [ ] **Step 3: Save**

---

### Task 3: Update `formatInputTab()` and `formatMasterTab()`

**Files:**
- Modify: `hold-form/Code.gs` (functions `formatInputTab`, `formatMasterTab`)

The CD-editable zone expands from 9 columns (A–I) to 12 (A–L). The auto-fill zone shifts to M–U (cols 13–21). The dividing border moves to the right of column L. A new YES/NO dropdown is added for In Contract? (H). Auto-resize updates from 18 to 21 columns.

- [ ] **Step 1: Replace `formatInputTab()`**

```javascript
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
    .setNote('YES if student is still within their initial enrollment contract. Leave NO for standard holds.');
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
```

- [ ] **Step 2: Replace `formatMasterTab()`**

```javascript
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
```

- [ ] **Step 3: Save**

---

### Task 4: Update `processDirectorInputs()`, `appendToMaster()`, and `masterRowToHoldData()`

**Files:**
- Modify: `hold-form/Code.gs` (Section 4)

Three new fields are read from the INPUT row, validated, added to `holdData`, written to MASTER_HOLDS, and returned by `masterRowToHoldData()`.

- [ ] **Step 1: In `processDirectorInputs()`, add reads and validation after the existing `expReturn` validation block**

Find this block (after the `if (!expReturn)` check, before the `try {`):

```javascript
    if (!expReturn) {
      logError('processDirectorInputs', 'Missing Expected Return Date',    `Row ${sheetRow} — ${studentName}`); continue;
    }

    try {
```

Replace it with:

```javascript
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
```

- [ ] **Step 2: In `processDirectorInputs()`, add the three new fields to the `holdData` object**

Find the `holdData` object literal (inside the `try {` block):

```javascript
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
        notes:          String(row[IC.NOTES]).trim()
      };
```

Replace it with:

```javascript
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
```

- [ ] **Step 3: Replace `appendToMaster()`**

```javascript
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
```

- [ ] **Step 4: Replace `masterRowToHoldData()`**

```javascript
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
```

- [ ] **Step 5: Save**

---

### Task 5: Update email templates

**Files:**
- Modify: `hold-form/Code.gs` (Section 6 — `getInitialEmailBody`, `getReminderEmailBody`)

When `holdData.inContract === 'YES'`, the "Enrollment Billing Resumes" table row is omitted, a billing notice paragraph replaces it, and the footer disclaimer text changes. The reminder email gets the same treatment.

- [ ] **Step 1: Replace `getInitialEmailBody()`**

```javascript
function getInitialEmailBody(holdData, settings) {
  const firstName = holdData.parentName.split(' ')[0];
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
```

- [ ] **Step 2: Replace `getReminderEmailBody()`**

```javascript
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
```

- [ ] **Step 3: Save**

---

### Task 6: Update `doGet()` and `getHoldById()`

**Files:**
- Modify: `hold-form/Code.gs` (Section 9)

`getHoldById()` needs to return the three new fields. `doGet()` needs to pass them to the template.

- [ ] **Step 1: Replace `getHoldById()`**

```javascript
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
```

- [ ] **Step 2: In `doGet()`, add the three new template variable assignments**

Find this block in `doGet()`:

```javascript
  template.holdId        = holdId;
  template.studentName   = hold.studentName;
  template.centerName    = hold.center || 'Mathnasium';
  template.startDate     = formatDateReadable(hold.holdStart);
  template.returnDate    = formatDateReadable(hold.expectedReturn);
  template.billingResume = formatBillingMonth(hold.billingResume);
  template.notes         = hold.notes || '';
  template.today         = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
```

Replace it with:

```javascript
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
```

- [ ] **Step 3: Save**

---

### Task 7: Update the acknowledgment page HTML

**Files:**
- Modify: `hold-form/Code.gs` (Section 10 — `getAcknowledgmentPageHtml`)

The details table conditionally shows or hides the "Enrollment Billing Resumes" row. A billing notice block appears when in-contract. The checkbox label text changes based on contract status.

- [ ] **Step 1: Replace the details table and notes section inside `getAcknowledgmentPageHtml()`**

Find this block in the HTML template (inside the `<!-- Hold details — read-only -->` section):

```html
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
        <tr>
          <td class="label">Enrollment Billing Resumes</td>
          <td class="value"><?= billingResume ?></td>
        </tr>
      </table>
    </div>

    <? if (notes) { ?>
    <div class="notes-box"><strong>Notes:</strong> <?= notes ?></div>
    <? } ?>
```

Replace it with:

```html
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
```

- [ ] **Step 2: Replace the acknowledgment checkbox label**

Find this block:

```html
    <!-- Acknowledgment -->
    <div class="ack-block">
      <input type="checkbox" id="ackCheck">
      <label for="ackCheck">
        I acknowledge the hold details listed above and understand that recurring
        enrollment billing will resume beginning with the return month listed above,
        unless alternative arrangements are made with the center before billing
        is processed.
      </label>
    </div>
```

Replace it with:

```html
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
```

- [ ] **Step 3: Save**

---

### Task 8: Re-run setup and re-deploy

This task has no code changes. It applies the updated column structure to the live sheet and publishes the updated web app.

- [ ] **Step 1: Re-run `formatInputTab` to apply the new column layout**

In the Apps Script editor:
- Select `formatInputTab` — but it requires a `ss` argument. Instead, run `setupWorkbook` from the ⚙ Hold Workflow menu to re-apply formatting to both INPUT and MASTER_HOLDS.

> ⚠ `setupWorkbook` uses `createTabIfNeeded` — it will NOT overwrite existing tabs or clear data. It is safe to run on a sheet that already has data. It will re-run `formatInputTab()` and `formatMasterTab()` on the existing tabs.

- [ ] **Step 2: Re-deploy the Web App**

In Apps Script:
1. Click **Deploy > Manage deployments**
2. Click the pencil (edit) icon on the existing deployment
3. Change **Version** to **New version**
4. Click **Deploy**
5. The URL stays the same — no need to update SETTINGS

- [ ] **Step 3: Test a standard hold (In Contract? = NO)**

In the INPUT tab:
- Add a test row with your own email, set `In Contract? = NO`, leave Charged/Credit months blank
- Set `Send Email? = YES`
- Run ⚙ Hold Workflow > Process New Rows
- Verify: email received shows "Enrollment Billing Resumes" row, NOT the billing notice
- Click the acknowledgment link, verify: "Enrollment Billing Resumes" row visible, standard checkbox text

- [ ] **Step 4: Test an in-contract hold (In Contract? = YES)**

In the INPUT tab:
- Add a test row with your own email, set `In Contract? = YES`, enter e.g. `June and July` for Charged Month(s), `August and September` for Credit Month(s)
- Set `Send Email? = YES`
- Run ⚙ Hold Workflow > Process New Rows
- Verify: email received shows billing notice ("...charged for **June and July**...credit...to **August and September**..."), "Enrollment Billing Resumes" row NOT shown
- Click the acknowledgment link, verify: billing notice visible, checkbox reads "...credit will be applied to August and September at the end of the initial contract"
- Submit the form, verify: MASTER_HOLDS updated (Signed? = YES), INPUT tab updated

- [ ] **Step 5: Test validation — missing In Contract?**

- Add a row with Send Email? = YES but leave In Contract? blank
- Run Process New Rows
- Verify: row skipped, error appears in ERROR_LOG tab

- [ ] **Step 6: Test validation — In Contract? = YES but missing months**

- Add a row with In Contract? = YES, Send Email? = YES, but leave Charged Month(s) blank
- Run Process New Rows
- Verify: row skipped, error appears in ERROR_LOG tab

- [ ] **Step 7: Repeat steps 3–6 for the Englewood sheet** (same process, different account)
