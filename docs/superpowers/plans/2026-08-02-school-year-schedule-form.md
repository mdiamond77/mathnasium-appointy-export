# School Year Schedule Preference Form — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build `school-year-schedule-form/`, a standalone Google Sheets + Apps Script system (one deployment per center) that collects school-year day/time preferences from two audiences — students returning from a summer hold, and continuing students — via a Mathnasium-branded parent web page, with reminders, a dashboard, and a slot-demand capacity grid.

**Architecture:** Everything lives in one file, `school-year-schedule-form/Code.gs`, following the same single-file convention as the existing `hold-form/Code.gs` (this codebase deploys Apps Script by pasting one file into the online editor — no `clasp`, no build step). Tasks build the file top-to-bottom in dependency order: pure time/slot math first (with real local Node tests, since that logic has no Google Apps Script dependency), then sheet scaffolding, then processing, email, the web app, dashboard, and slot demand. Everything that touches `SpreadsheetApp`/`GmailApp`/`HtmlService` can't be unit-tested locally (those globals only exist inside Google's runtime) — those tasks are verified with `node --check` for syntax correctness as the file grows, and with a full manual test pass against a real deployed sheet in the final task, matching how `hold-form` itself was verified (see `docs/superpowers/plans/2026-05-27-hold-form-in-contract.md`).

**Tech Stack:** Google Apps Script, Google Sheets, GmailApp, HtmlService (scriptlet templates), Node.js (local-only, for testing pure logic — not deployed)

**Spec:** `docs/superpowers/specs/2026-08-02-school-year-schedule-form-design.md`

---

## File Structure

- `school-year-schedule-form/Code.gs` — the entire Apps Script implementation, built incrementally, organized into 18 numbered sections (0–17)
- `school-year-schedule-form/test/slot-logic.test.js` — local Node test for the pure time/slot math in Section 0 (`node school-year-schedule-form/test/slot-logic.test.js`) — dev-only, never pasted into Apps Script
- `school-year-schedule-form/SETUP.md` — setup/reference guide for deploying to each center's Google account, written the same way as `hold-form/SETUP.md`

---

### Task 1: Write failing tests for the pure time/slot logic

**Files:**
- Create: `school-year-schedule-form/test/slot-logic.test.js`

This logic (dropdown time options, the 1-hour-before-close cutoff, the "occupies two half-hour blocks" rule for SLOT_DEMAND, and Request ID formatting) is the trickiest math in the whole system and has zero Google Apps Script dependency — so it gets real automated tests, run locally with plain Node (no framework needed).

- [ ] **Step 1: Write the test file**

```javascript
const assert = require('assert');
const {
  parseTimeToMinutes, minutesToTimeLabel, dayHours,
  selectableStartTimes, slotDemandRows, occupiedSlots,
  buildRequestId, isPrefPairValid
} = require('../Code.gs');

// parseTimeToMinutes
assert.strictEqual(parseTimeToMinutes('10:00'), 600, 'parseTimeToMinutes 10:00');
assert.strictEqual(parseTimeToMinutes('15:30'), 930, 'parseTimeToMinutes 15:30');
assert.strictEqual(parseTimeToMinutes('19:30'), 1170, 'parseTimeToMinutes 19:30');

// minutesToTimeLabel
assert.strictEqual(minutesToTimeLabel(600), '10:00 AM', 'minutesToTimeLabel 600');
assert.strictEqual(minutesToTimeLabel(930), '3:30 PM', 'minutesToTimeLabel 930');
assert.strictEqual(minutesToTimeLabel(0), '12:00 AM', 'minutesToTimeLabel midnight');
assert.strictEqual(minutesToTimeLabel(720), '12:00 PM', 'minutesToTimeLabel noon');

// selectableStartTimes — Englewood Sunday (10:00-13:00): stop 1hr before close
assert.deepStrictEqual(
  selectableStartTimes('10:00', '13:00'),
  [600, 630, 660, 690, 720],
  'Englewood Sunday start times'
);

// selectableStartTimes — Englewood weekday (15:30-19:00)
assert.deepStrictEqual(
  selectableStartTimes('15:30', '19:00'),
  [930, 960, 990, 1020, 1050, 1080],
  'Englewood weekday start times'
);

// selectableStartTimes — Teaneck Sunday (10:00-14:00)
assert.deepStrictEqual(
  selectableStartTimes('10:00', '14:00'),
  [600, 630, 660, 690, 720, 750, 780],
  'Teaneck Sunday start times'
);

// selectableStartTimes — Teaneck weekday (15:30-19:30)
assert.deepStrictEqual(
  selectableStartTimes('15:30', '19:30'),
  [930, 960, 990, 1020, 1050, 1080, 1110],
  'Teaneck weekday start times'
);

// slotDemandRows — rows run one block past the last selectable start
assert.deepStrictEqual(
  slotDemandRows('15:30', '19:00'),
  [930, 960, 990, 1020, 1050, 1080, 1110],
  'Englewood weekday demand rows (3:30 through 6:30)'
);
assert.deepStrictEqual(
  slotDemandRows('10:00', '14:00'),
  [600, 630, 660, 690, 720, 750, 780, 810],
  'Teaneck Sunday demand rows (10:00 through 1:30)'
);

// occupiedSlots — a 3:30 preference occupies 3:30 and 4:00
assert.deepStrictEqual(occupiedSlots(930), [930, 960], 'occupiedSlots 3:30 -> [3:30, 4:00]');

// dayHours
const settings = { SUN_OPEN: '10:00', SUN_CLOSE: '14:00', WEEKDAY_OPEN: '15:30', WEEKDAY_CLOSE: '19:30' };
assert.deepStrictEqual(dayHours('Sunday', settings), { open: '10:00', close: '14:00' }, 'dayHours Sunday');
assert.deepStrictEqual(dayHours('Wednesday', settings), { open: '15:30', close: '19:30' }, 'dayHours weekday');

// buildRequestId
assert.strictEqual(buildRequestId('20260915', 3), 'SCHED-20260915-0003', 'buildRequestId zero-pads to 4 digits');
assert.strictEqual(buildRequestId('20260915', 12), 'SCHED-20260915-0012', 'buildRequestId two-digit seq');

// isPrefPairValid
assert.strictEqual(isPrefPairValid('Monday', '930'), true, 'both given -> valid');
assert.strictEqual(isPrefPairValid('', ''), true, 'both blank -> valid (optional pref)');
assert.strictEqual(isPrefPairValid('Monday', ''), false, 'day only -> invalid');
assert.strictEqual(isPrefPairValid('', '930'), false, 'time only -> invalid');

console.log('All slot-logic tests passed.');
```

- [ ] **Step 2: Run it and confirm it fails**

Run: `node school-year-schedule-form/test/slot-logic.test.js`
Expected: `Error: Cannot find module '../Code.gs'` (the file doesn't exist yet)

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/test/slot-logic.test.js
git commit -m "test: add failing tests for school-year-schedule-form slot logic"
```

---

### Task 2: Implement the pure time/slot logic (Section 0)

**Files:**
- Create: `school-year-schedule-form/Code.gs`

- [ ] **Step 1: Write Section 0**

```javascript
// =============================================================================
// MATHNASIUM — SCHOOL YEAR SCHEDULE PREFERENCE FORM
// Google Apps Script  |  One deployment per center account
// =============================================================================
//
// QUICK START:
//   1. Open Google Sheet in the center's Google account
//   2. Extensions > Apps Script — paste this file as Code.gs, then Save
//   3. Run  setupWorkbook()  from the ⚙ School Year Schedule menu (authorize when prompted)
//   4. Review and update the SETTINGS tab
//   5. Deploy as Web App:
//        Extensions > Apps Script > Deploy > New deployment
//        Type: Web app | Execute as: Me | Who has access: Anyone
//        Copy the deployment URL → paste into SETTINGS as WEB_APP_URL
//   6. Run  installTriggers()  from the ⚙ School Year Schedule menu
//   7. Share the sheet with Matt's account as Editor
//
// =============================================================================


// =============================================================================
// SECTION 0 — TIME / SLOT LOGIC (pure functions, no GAS dependencies)
// Testable locally: node school-year-schedule-form/test/slot-logic.test.js
// =============================================================================

const SLOT_INCREMENT_MIN = 30;
const SLOT_CUTOFF_BEFORE_CLOSE_MIN = 60;
const OFFERED_DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday'];

function parseTimeToMinutes(hhmm) {
  const parts = String(hhmm).split(':');
  return Number(parts[0]) * 60 + Number(parts[1]);
}

function minutesToTimeLabel(mins) {
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  const ap = h >= 12 ? 'PM' : 'AM';
  let h12 = h % 12;
  if (h12 === 0) h12 = 12;
  return h12 + ':' + String(m).padStart(2, '0') + ' ' + ap;
}

function dayHours(day, settings) {
  if (day === 'Sunday') return { open: settings['SUN_OPEN'], close: settings['SUN_CLOSE'] };
  return { open: settings['WEEKDAY_OPEN'], close: settings['WEEKDAY_CLOSE'] };
}

// Selectable start times for the parent-facing dropdown: 30-min increments
// from open, stopping 1 hour before close so a full session fits before close.
function selectableStartTimes(openHHMM, closeHHMM) {
  const start  = parseTimeToMinutes(openHHMM);
  const cutoff = parseTimeToMinutes(closeHHMM) - SLOT_CUTOFF_BEFORE_CLOSE_MIN;
  const out = [];
  for (let t = start; t <= cutoff; t += SLOT_INCREMENT_MIN) out.push(t);
  return out;
}

// All half-hour rows a center occupies for a day in the SLOT_DEMAND grid:
// from open through one block past the last selectable start time, so the
// occupancy spillover (start + 30) always has a row to land in.
function slotDemandRows(openHHMM, closeHHMM) {
  const starts = selectableStartTimes(openHHMM, closeHHMM);
  if (!starts.length) return [];
  const last = starts[starts.length - 1];
  const openMin = parseTimeToMinutes(openHHMM);
  const rows = [];
  for (let t = openMin; t <= last + SLOT_INCREMENT_MIN; t += SLOT_INCREMENT_MIN) rows.push(t);
  return rows;
}

// A preference occupies its start slot and the following half-hour slot.
function occupiedSlots(startMinutes) {
  return [startMinutes, startMinutes + SLOT_INCREMENT_MIN];
}

function buildRequestId(dateStr, seq) {
  return 'SCHED-' + dateStr + '-' + String(seq).padStart(4, '0');
}

// A preference pair (day + time) must be given together or both left blank.
function isPrefPairValid(day, time) {
  const dayGiven  = day !== null && day !== undefined && String(day).trim() !== '';
  const timeGiven = time !== null && time !== undefined && String(time).trim() !== '';
  return (dayGiven && timeGiven) || (!dayGiven && !timeGiven);
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseTimeToMinutes, minutesToTimeLabel, dayHours,
    selectableStartTimes, slotDemandRows, occupiedSlots,
    buildRequestId, isPrefPairValid
  };
}
```

- [ ] **Step 2: Run the tests and confirm they pass**

Run: `node school-year-schedule-form/test/slot-logic.test.js`
Expected: `All slot-logic tests passed.`

- [ ] **Step 3: Syntax-check the full file**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output (silent success)

- [ ] **Step 4: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add pure time/slot logic for school-year-schedule-form"
```

---

### Task 3: Sheet/tab constants (Section 1)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append Section 1 to the end of the file**

```javascript


// =============================================================================
// SECTION 1 — SHEET / TAB CONSTANTS
// =============================================================================

const TAB_INPUT_HOLD       = 'INPUT_RETURNING_FROM_HOLD';
const TAB_INPUT_CONTINUING = 'INPUT_CONTINUING_STUDENTS';
const TAB_MASTER           = 'MASTER_REQUESTS';
const TAB_SETTINGS         = 'SETTINGS';
const TAB_DASHBOARD        = 'DASHBOARD';
const TAB_SLOT_DEMAND      = 'SLOT_DEMAND';
const TAB_ERROR_LOG        = 'ERROR_LOG';

const TYPE_HOLD       = 'Returning from Hold';
const TYPE_CONTINUING = 'Continuing Student';

// ── Brand colors ──────────────────────────────────────────────────────────
const CLR_RED       = '#EE3124';
const CLR_NAVY      = '#1C2B6E';
const CLR_RED_LIGHT = '#FDEDEC';
const CLR_RED_MUTED = '#C0392B';
const CLR_WHITE     = '#FFFFFF';

// ── INPUT tab column indices (0-based) — identical layout for both tabs ───
const IC = {
  STUDENT_NAME:   0,  // A
  PARENT_NAME:    1,  // B
  PARENT_EMAIL:   2,  // C
  SEND_EMAIL:     3,  // D  ← YES to trigger
  STOP_SENDING:   4,  // E  ← YES halts all future emails, for any reason
  STOP_REASON:    5,  // F  ← CD context only, never shown to parent
  // ── auto-filled (greyed) ─────────────────────
  STATUS:         6,  // G
  PREF_LINK:      7,  // H
  EMAIL_SENT_TS:  8,  // I
  RESPONDED:      9,  // J
  RESPONDED_TS:   10, // K
  REMINDER_1:     11, // L
  REMINDER_2:     12, // M
  REMINDER_3:     13, // N
  REQUEST_ID:     14  // O
};

// ── MASTER_REQUESTS column indices (0-based) ───────────────────────────────
const MC = {
  REQUEST_ID:      0,
  CENTER:          1,
  TYPE:            2,
  SOURCE_TAB:      3,
  SOURCE_ROW:      4,
  STUDENT_NAME:    5,
  PARENT_NAME:     6,
  PARENT_EMAIL:    7,
  PREF_LINK:       8,
  EMAIL_SENT:      9,
  EMAIL_SENT_TS:   10,
  RESPONDED:       11,
  RESPONDED_TS:    12,
  REMINDER_1:      13,
  REMINDER_2:      14,
  REMINDER_3:      15,
  PREF1_DAY:       16,
  PREF1_TIME:      17,
  PREF1_NOTE:      18,
  PREF2_DAY:       19,
  PREF2_TIME:      20,
  PREF2_NOTE:      21,
  STATUS:          22
};

// ── Reminder timing (hours after initial email) ────────────────────────────
const REMINDER_1_HOURS = 48;   // 2 days
const REMINDER_2_HOURS = 120;  // 5 days
const REMINDER_3_HOURS = 192;  // 8 days

const INPUT_FORMAT_ROWS = 500;
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Re-run the slot-logic tests to confirm nothing broke**

Run: `node school-year-schedule-form/test/slot-logic.test.js`
Expected: `All slot-logic tests passed.`

- [ ] **Step 4: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add sheet/tab constants for school-year-schedule-form"
```

---

### Task 4: Menu and settings helpers (Sections 2–3)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append Sections 2 and 3**

```javascript


// =============================================================================
// SECTION 2 — MENU
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙ School Year Schedule')
    .addItem('▶  Process New Rows (send emails)',  'processDirectorInputs')
    .addItem('🔔  Check Reminder Statuses',         'checkReminderStatuses')
    .addItem('📊  Refresh Dashboard',               'refreshDashboardAndSlotDemand')
    .addSeparator()
    .addItem('🔧  Setup Workbook (first-time)',     'setupWorkbook')
    .addItem('⏱  Install Triggers',                'installTriggers')
    .addItem('🗑  Remove All Triggers',              'removeTriggers')
    .addToUi();
}

function refreshDashboardAndSlotDemand() {
  buildDashboard();
  buildSlotDemand();
}


// =============================================================================
// SECTION 3 — SETTINGS HELPERS
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

function initSettingsDefaults(ss) {
  const sheet = ss.getSheetByName(TAB_SETTINGS);
  if (sheet.getLastRow() > 1) return;

  const rows = [
    ['CENTER_NAME',    'Mathnasium of Teaneck'],
    ['CENTER_EMAIL',   'teaneck@mathnasium.com'],
    ['CD_EMAIL',       'teaneck@mathnasium.com'],
    ['FROM_NAME',      'Mathnasium of Teaneck'],
    ['PHONE',          '(201) XXX-XXXX'],
    ['SUN_OPEN',       '10:00'],
    ['SUN_CLOSE',      '14:00'],
    ['WEEKDAY_OPEN',   '15:30'],
    ['WEEKDAY_CLOSE',  '19:30'],
    ['WEB_APP_URL',    '']
  ];
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  sheet.getRange(rows.length + 1, 1, 1, 2).setBackground('#FFF9C4'); // yellow highlight for WEB_APP_URL
}
```

> Defaults above are Teaneck's hours (Sunday close 14:00, weekday close 19:30). SETUP.md will call out updating these to Englewood's (13:00 / 19:00) on that center's copy.

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add menu and settings helpers for school-year-schedule-form"
```

---

### Task 5: Setup — tabs and headers (Section 4, part A)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append the first half of Section 4**

```javascript


// =============================================================================
// SECTION 4 — SETUP
// =============================================================================

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  createTabIfNeeded(ss, TAB_INPUT_HOLD,       getInputHeaders());
  createTabIfNeeded(ss, TAB_INPUT_CONTINUING, getInputHeaders());
  createTabIfNeeded(ss, TAB_MASTER,           getMasterHeaders());
  createTabIfNeeded(ss, TAB_SETTINGS,         ['Key', 'Value']);
  createTabIfNeeded(ss, TAB_ERROR_LOG,        ['Timestamp', 'Function', 'Error', 'Context']);
  createTabIfNeeded(ss, TAB_DASHBOARD,        null);
  createTabIfNeeded(ss, TAB_SLOT_DEMAND,      null);

  initSettingsDefaults(ss);
  formatInputTab(ss, TAB_INPUT_HOLD);
  formatInputTab(ss, TAB_INPUT_CONTINUING);
  formatMasterTab(ss);
  buildDashboard();
  buildSlotDemand();

  ui.alert(
    '✓ Setup Complete',
    'Next steps:\n\n' +
    '1. Update SETTINGS tab (PHONE, CENTER_NAME, hours, etc.)\n\n' +
    '2. Deploy as Web App:\n' +
    '   Extensions > Apps Script > Deploy > New deployment\n' +
    '   Type: Web app | Execute as: Me | Who has access: Anyone\n' +
    '   → Paste the URL into SETTINGS as WEB_APP_URL\n\n' +
    '3. ⚙ School Year Schedule > Install Triggers',
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
    'Student Name',       // A  IC.STUDENT_NAME
    'Parent Name',        // B
    'Parent Email',       // C
    'Send Email?',        // D  ← YES to trigger
    'Stop Sending?',      // E  ← YES halts all future emails
    'Stop Reason',        // F  ← CD context only
    // ── auto-filled ─────────────────────
    'Status',             // G
    'Preference Link',    // H
    'Email Sent',         // I
    'Responded?',         // J
    'Responded On',       // K
    'Reminder 1 Sent',    // L
    'Reminder 2 Sent',    // M
    'Reminder 3 Sent',    // N
    'Request ID'          // O
  ];
}

function getMasterHeaders() {
  return [
    'Request ID', 'Center', 'Type', 'Source Tab', 'Source Row',
    'Student Name', 'Parent Name', 'Parent Email',
    'Preference Link', 'Email Sent', 'Email Sent Date',
    'Responded?', 'Responded On',
    'Reminder 1 Sent', 'Reminder 2 Sent', 'Reminder 3 Sent',
    'Pref 1 Day', 'Pref 1 Time', 'Pref 1 Note',
    'Pref 2 Day', 'Pref 2 Time', 'Pref 2 Note',
    'Status'
  ];
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add setup workbook, tab creation, and headers"
```

---

### Task 6: Setup — formatting and dropdown validation (Section 4, part B)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append the rest of Section 4**

```javascript

/**
 * Formats an INPUT tab (works for either TAB_INPUT_HOLD or TAB_INPUT_CONTINUING):
 *  • A–F (cols 1–6): white, CD-editable.
 *  • G–O (cols 7–15): grey, auto-filled by script.
 */
function formatInputTab(ss, tabName) {
  const sheet    = ss.getSheetByName(tabName);
  if (!sheet) return;
  const dataRows = INPUT_FORMAT_ROWS;

  // CD-editable zone A–F: white
  sheet.getRange(2, 1, dataRows, 6).setBackground(CLR_WHITE).setFontColor('#222222');

  // Auto-fill zone G–O: grey
  sheet.getRange(2, 7, dataRows, 9).setBackground('#EFEFEF').setFontColor('#888888');

  // Muted header color for auto-fill columns G–O
  sheet.getRange(1, 7, 1, 9).setBackground(CLR_RED_MUTED).setFontColor(CLR_WHITE);

  // Thick border between F and G to mark the boundary
  sheet.getRange(1, 6, dataRows + 1, 1)
    .setBorder(null, null, null, true, null, null, CLR_RED, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Notes on auto-fill headers (G–O)
  const autoNotes = [
    'Auto-filled: Sent / Responded',
    'Auto-filled: preference page URL',
    'Auto-filled: date initial email sent',
    'Auto-filled: YES when parent submits',
    'Auto-filled: date parent submitted',
    'Auto-filled: date Reminder 1 sent',
    'Auto-filled: date Reminder 2 sent',
    'Auto-filled: date Reminder 3 sent',
    'Auto-filled: unique Request ID'
  ];
  autoNotes.forEach((note, i) => sheet.getRange(1, 7 + i).setNote(note));

  // Note on Stop Sending
  sheet.getRange(1, IC.STOP_SENDING + 1)
    .setNote('YES halts all future emails for this row (initial send if not yet sent, or any remaining reminders) — for any reason, e.g. not returning.');

  // YES/NO dropdown on Send Email? (D)
  const sendEmailRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['YES', 'NO'], true)
    .setAllowInvalid(false)
    .setHelpText('Select YES to send the schedule preference email to the parent.')
    .build();
  sheet.getRange(2, IC.SEND_EMAIL + 1, dataRows, 1).setDataValidation(sendEmailRule);

  // YES/NO dropdown on Stop Sending? (E)
  const stopSendingRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['YES', 'NO'], true)
    .setAllowInvalid(false)
    .setHelpText('Select YES to permanently stop all emails for this row.')
    .build();
  sheet.getRange(2, IC.STOP_SENDING + 1, dataRows, 1).setDataValidation(stopSendingRule);

  // Date/time format on timestamp columns
  const dateFmt = 'M/d/yyyy h:mm am/pm';
  [IC.EMAIL_SENT_TS, IC.RESPONDED_TS, IC.REMINDER_1, IC.REMINDER_2, IC.REMINDER_3].forEach(col => {
    sheet.getRange(2, col + 1, dataRows, 1).setNumberFormat(dateFmt);
  });

  sheet.autoResizeColumns(1, 15);
}

function formatMasterTab(ss) {
  const sheet = ss.getSheetByName(TAB_MASTER);
  if (!sheet) return;
  const dateFmt = 'M/d/yyyy h:mm am/pm';
  [MC.EMAIL_SENT_TS, MC.RESPONDED_TS, MC.REMINDER_1, MC.REMINDER_2, MC.REMINDER_3].forEach(col => {
    sheet.getRange(2, col + 1, INPUT_FORMAT_ROWS, 1).setNumberFormat(dateFmt);
  });
  sheet.autoResizeColumns(1, 23);
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add INPUT/MASTER tab formatting and dropdown validation"
```

---

### Task 7: Main processing loop (Section 5, part A)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append the processing loop**

```javascript


// =============================================================================
// SECTION 5 — MAIN PROCESSING
// =============================================================================

function processDirectorInputs() {
  processInputTab(TAB_INPUT_HOLD, TYPE_HOLD);
  processInputTab(TAB_INPUT_CONTINUING, TYPE_CONTINUING);
}

function processInputTab(tabName, type) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings();
  const inputSh  = ss.getSheetByName(tabName);
  const masterSh = ss.getSheetByName(TAB_MASTER);

  if (!inputSh)  { logError('processInputTab', tabName + ' tab missing', '');       return; }
  if (!masterSh) { logError('processInputTab', 'MASTER_REQUESTS tab missing', ''); return; }

  const webAppUrl = String(settings['WEB_APP_URL'] || '').trim();
  if (!webAppUrl) {
    logError('processInputTab', 'WEB_APP_URL not set in SETTINGS',
             'Deploy the script as a Web App and paste the URL into SETTINGS.');
    return;
  }

  const lastRow = inputSh.getLastRow();
  if (lastRow < 2) return;

  const data = inputSh.getRange(2, 1, lastRow - 1, Object.keys(IC).length).getValues();

  for (let i = 0; i < data.length; i++) {
    const row       = data[i];
    const sheetRow  = i + 2;
    const sendEmail = String(row[IC.SEND_EMAIL]).trim().toUpperCase();
    const stopSend  = String(row[IC.STOP_SENDING]).trim().toUpperCase();
    const status    = String(row[IC.STATUS]).trim();

    if (sendEmail !== 'YES' || status !== '' || stopSend === 'YES') continue;

    const studentName = String(row[IC.STUDENT_NAME]).trim();
    const parentEmail = String(row[IC.PARENT_EMAIL]).trim();

    if (!studentName) {
      logError('processInputTab', 'Missing Student Name', `${tabName} Row ${sheetRow}`); continue;
    }
    if (!parentEmail || !parentEmail.includes('@')) {
      logError('processInputTab', 'Missing or invalid Parent Email', `${tabName} Row ${sheetRow} — ${studentName}`); continue;
    }

    try {
      const requestId = generateRequestId(masterSh);
      const reqData = {
        requestId,
        center:     settings['CENTER_NAME'],
        type,
        sourceTab:  tabName,
        sourceRow:  sheetRow,
        studentName,
        parentName: String(row[IC.PARENT_NAME]).trim(),
        parentEmail
      };

      reqData.prefLink = generateWebAppLink(requestId, webAppUrl);
      sendInitialEmail(reqData, settings);

      const now = new Date();
      appendToMaster(masterSh, reqData, now);

      inputSh.getRange(sheetRow, IC.REQUEST_ID    + 1).setValue(requestId);
      inputSh.getRange(sheetRow, IC.STATUS        + 1).setValue('Sent');
      inputSh.getRange(sheetRow, IC.PREF_LINK     + 1).setValue(reqData.prefLink);
      inputSh.getRange(sheetRow, IC.EMAIL_SENT_TS + 1).setValue(now);

      Logger.log(`processInputTab ✓  ${tabName} Row ${sheetRow}: ${studentName} — ${requestId}`);
    } catch (err) {
      logError('processInputTab', err.message, `${tabName} Row ${sheetRow} — ${studentName}`);
    }
  }
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add processDirectorInputs for both INPUT tabs"
```

---

### Task 8: Request ID generation and MASTER append (Section 5, part B)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append**

```javascript

function generateRequestId(masterSh) {
  const tz     = Session.getScriptTimeZone();
  const today  = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const prefix = 'SCHED-' + today + '-';
  const data   = masterSh.getDataRange().getValues();
  let maxN     = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][MC.REQUEST_ID]);
    if (id.startsWith(prefix)) {
      const n = parseInt(id.slice(prefix.length), 10);
      if (n > maxN) maxN = n;
    }
  }
  return buildRequestId(today, maxN + 1);
}

function appendToMaster(masterSh, reqData, timestamp) {
  masterSh.appendRow([
    reqData.requestId, reqData.center, reqData.type, reqData.sourceTab, reqData.sourceRow,
    reqData.studentName, reqData.parentName, reqData.parentEmail,
    reqData.prefLink,
    'YES', timestamp,
    '', '',
    '', '', '',
    '', '', '',
    '', '', '',
    'Sent'
  ]);
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add generateRequestId and appendToMaster"
```

---

### Task 9: Web app link + initial email — Returning from Hold (Section 6, Section 7 part A)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append**

```javascript


// =============================================================================
// SECTION 6 — WEB APP LINK GENERATION
// =============================================================================

function generateWebAppLink(requestId, webAppUrl) {
  return `${webAppUrl}?id=${encodeURIComponent(requestId)}`;
}


// =============================================================================
// SECTION 7 — EMAIL TEMPLATES
// ─────────────────────────────────────────────────────────────────────────────
// EDIT THIS SECTION to change email wording or layout.
// Available variables: reqData.studentName, reqData.parentName, reqData.parentEmail,
//   reqData.prefLink, reqData.type, reqData.center; settings['FROM_NAME'], settings['PHONE']
// =============================================================================

function getDisclaimerHtml() {
  return `We'll do our best to lock in exactly what you choose. If scheduling gets tight and we
    need to offer a slightly different day or time, we'll reach out personally to find something
    that works — nothing is finalized without talking to you first.`;
}

function getCtaButtonHtml(link) {
  return `<p style="margin:24px 0;">
    <a href="${link}"
       style="display:inline-block;background-color:${CLR_RED};color:${CLR_WHITE};
              padding:12px 28px;text-decoration:none;border-radius:4px;
              font-weight:bold;font-size:14px;">
      Tell Us Your Preferred Schedule
    </a>
  </p>`;
}

function getInitialEmailSubjectHold(reqData, settings) {
  return `Welcome Back! Help Us Plan ${reqData.studentName}'s Fall Schedule — ${settings['CENTER_NAME']}`;
}

function getInitialEmailBodyHold(reqData, settings) {
  const firstName = reqData.parentName.split(' ')[0];
  return `
<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;line-height:1.6;">
  <div style="background-color:${CLR_RED};padding:20px 24px;border-radius:6px 6px 0 0;">
    <p style="color:${CLR_WHITE};font-size:18px;font-weight:bold;margin:0;">
      Let's Plan ${reqData.studentName}'s Fall Schedule
    </p>
  </div>
  <div style="border:1px solid #ddd;border-top:none;border-radius:0 0 6px 6px;padding:24px;">
    <p>Hi ${firstName},</p>
    <p>Welcome back! We hope ${reqData.studentName} had a great summer, and we're excited to
    have them back at ${settings['FROM_NAME']}.</p>
    <p>We're putting together our schedule for the upcoming school year, and spots are filling
    up quickly. To help us reserve a great time for ${reqData.studentName}, could you let us
    know your preferred day(s) and time(s)?</p>

    ${getCtaButtonHtml(reqData.prefLink)}

    <p style="font-size:13px;color:#666;border-top:1px solid #eee;padding-top:14px;">
      ${getDisclaimerHtml()}
    </p>

    <p>Thank you,<br>
    <strong>${settings['FROM_NAME']}</strong><br>
    ${settings['PHONE'] || ''}<br>
    ${settings['CENTER_EMAIL'] || ''}</p>
  </div>
</div>`;
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add web app link generation and returning-from-hold email template"
```

---

### Task 10: Initial email — Continuing Student + template dispatcher (Section 7, part B)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append**

```javascript

function getInitialEmailSubjectContinuing(reqData, settings) {
  return `Let's Set ${reqData.studentName}'s School-Year Schedule — ${settings['CENTER_NAME']}`;
}

function getInitialEmailBodyContinuing(reqData, settings) {
  const firstName = reqData.parentName.split(' ')[0];
  return `
<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;line-height:1.6;">
  <div style="background-color:${CLR_RED};padding:20px 24px;border-radius:6px 6px 0 0;">
    <p style="color:${CLR_WHITE};font-size:18px;font-weight:bold;margin:0;">
      Let's Set ${reqData.studentName}'s School-Year Schedule
    </p>
  </div>
  <div style="border:1px solid #ddd;border-top:none;border-radius:0 0 6px 6px;padding:24px;">
    <p>Hi ${firstName},</p>
    <p>As summer winds down, we're shifting gears from our summer schedule to our school-year
    schedule! ${reqData.studentName}'s current time slot may not carry over automatically, so
    we want to make sure we lock in a great time before the year gets underway.</p>
    <p>We're putting together our schedule for the upcoming school year, and spots are filling
    up quickly. Could you let us know your preferred day(s) and time(s)?</p>

    ${getCtaButtonHtml(reqData.prefLink)}

    <p style="font-size:13px;color:#666;border-top:1px solid #eee;padding-top:14px;">
      ${getDisclaimerHtml()}
    </p>

    <p>Thank you,<br>
    <strong>${settings['FROM_NAME']}</strong><br>
    ${settings['PHONE'] || ''}<br>
    ${settings['CENTER_EMAIL'] || ''}</p>
  </div>
</div>`;
}

function getInitialEmailSubject(reqData, settings) {
  return reqData.type === TYPE_CONTINUING
    ? getInitialEmailSubjectContinuing(reqData, settings)
    : getInitialEmailSubjectHold(reqData, settings);
}

function getInitialEmailBody(reqData, settings) {
  return reqData.type === TYPE_CONTINUING
    ? getInitialEmailBodyContinuing(reqData, settings)
    : getInitialEmailBodyHold(reqData, settings);
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add continuing-student email template and subject/body dispatcher"
```

---

### Task 11: Reminder email + CD notification templates (Section 7, part C)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append**

```javascript

function getReminderEmailSubject(reqData) {
  return `Reminder: We Still Need ${reqData.studentName}'s Schedule Preference — ${reqData.center}`;
}

function getReminderEmailBody(reqData, settings) {
  const firstName = reqData.parentName.split(' ')[0];
  return `
<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;line-height:1.6;">
  <div style="background-color:${CLR_RED};padding:20px 24px;border-radius:6px 6px 0 0;">
    <p style="color:${CLR_WHITE};font-size:18px;font-weight:bold;margin:0;">
      We Still Need ${reqData.studentName}'s Schedule Preference
    </p>
  </div>
  <div style="border:1px solid #ddd;border-top:none;border-radius:0 0 6px 6px;padding:24px;">
    <p>Hi ${firstName},</p>
    <p>We haven't heard back yet about ${reqData.studentName}'s preferred day and time for the
    school year. Spots are filling up quickly, so we'd love to hear from you soon.</p>

    ${getCtaButtonHtml(reqData.prefLink)}

    <p style="font-size:13px;color:#666;border-top:1px solid #eee;padding-top:14px;">
      ${getDisclaimerHtml()}
    </p>

    <p>Thank you,<br>
    <strong>${settings['FROM_NAME']}</strong><br>
    ${settings['PHONE'] || ''}<br>
    ${settings['CENTER_EMAIL'] || ''}</p>
  </div>
</div>`;
}

function getCdNotificationSubject(reqData) {
  return `Schedule Preference Received — ${reqData.studentName}`;
}

function getCdNotificationBody(reqData) {
  const pref2Html = reqData.pref2Day
    ? `<tr>
         <td style="padding:6px 20px 6px 0;font-weight:bold;color:#333;vertical-align:top;">Preference 2:</td>
         <td style="padding:6px 0;">${reqData.pref2Day} at ${minutesToTimeLabel(Number(reqData.pref2Time))}${reqData.pref2Note ? ' — ' + reqData.pref2Note : ''}</td>
       </tr>`
    : '';
  return `
<div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;line-height:1.6;">
  <p><strong>${reqData.studentName}</strong> (${reqData.type}) submitted a schedule preference:</p>
  <table style="border-collapse:collapse;margin:12px 0;">
    <tr>
      <td style="padding:6px 20px 6px 0;font-weight:bold;color:#333;vertical-align:top;">Preference 1:</td>
      <td style="padding:6px 0;">${reqData.pref1Day} at ${minutesToTimeLabel(Number(reqData.pref1Time))}${reqData.pref1Note ? ' — ' + reqData.pref1Note : ''}</td>
    </tr>
    ${pref2Html}
  </table>
  <p style="color:#666;font-size:13px;">Parent: ${reqData.parentName} (${reqData.parentEmail})</p>
</div>`;
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add reminder and CD notification email templates"
```

---

### Task 12: Email sending functions (Section 8)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append**

```javascript


// =============================================================================
// SECTION 8 — EMAIL SENDING
// =============================================================================

function sendInitialEmail(reqData, settings) {
  GmailApp.sendEmail(
    reqData.parentEmail,
    getInitialEmailSubject(reqData, settings),
    '',
    {
      htmlBody: getInitialEmailBody(reqData, settings),
      name:     settings['FROM_NAME'] || 'Mathnasium',
      replyTo:  settings['CENTER_EMAIL'] || ''
    }
  );
  Logger.log(`sendInitialEmail → ${reqData.parentEmail}  (${reqData.studentName})`);
}

function sendReminderEmail(reqData, settings) {
  const cdEmail = settings['CD_EMAIL'] || '';
  GmailApp.sendEmail(
    reqData.parentEmail,
    getReminderEmailSubject(reqData),
    '',
    {
      htmlBody: getReminderEmailBody(reqData, settings),
      name:     settings['FROM_NAME'] || 'Mathnasium',
      replyTo:  settings['CENTER_EMAIL'] || '',
      cc:       cdEmail
    }
  );
  Logger.log(`sendReminderEmail → ${reqData.parentEmail}  (${reqData.studentName})`);
}

function sendCdNotification(reqData, settings) {
  const cdEmail = settings['CD_EMAIL'] || '';
  if (!cdEmail) return;
  GmailApp.sendEmail(
    cdEmail,
    getCdNotificationSubject(reqData),
    '',
    {
      htmlBody: getCdNotificationBody(reqData),
      name:     settings['FROM_NAME'] || 'Mathnasium'
    }
  );
  Logger.log(`sendCdNotification → ${cdEmail}  (${reqData.studentName})`);
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add email sending functions"
```

---

### Task 13: Reminder checker with live Stop Sending re-check (Section 9)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

`isStopSendingOrDeleted` re-reads the current Stop Sending value from the row's own source INPUT tab on every run — not a cached snapshot — so toggling it to YES after the initial email still stops all remaining reminders. If the row can no longer be found there (deleted), it's also treated as stopped.

- [ ] **Step 1: Append**

```javascript


// =============================================================================
// SECTION 9 — REMINDER CHECKER
// =============================================================================

function checkReminderStatuses() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const masterSh = ss.getSheetByName(TAB_MASTER);
  const settings = getSettings();
  const now      = new Date();

  if (!masterSh || masterSh.getLastRow() < 2) return;

  const data = masterSh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row       = data[i];
    const responded = String(row[MC.RESPONDED]).trim().toUpperCase();
    const sentTs    = row[MC.EMAIL_SENT_TS];

    if (responded === 'YES' || !sentTs) continue;

    const requestId = String(row[MC.REQUEST_ID]).trim();
    const sourceTab = String(row[MC.SOURCE_TAB]).trim();
    if (isStopSendingOrDeleted(ss, sourceTab, requestId)) continue;

    const hrs = (now - new Date(sentTs)) / 3600000;
    const r1  = row[MC.REMINDER_1];
    const r2  = row[MC.REMINDER_2];
    const r3  = row[MC.REMINDER_3];
    const rd  = masterRowToRequestData(row);
    let num   = null;

    if      (!r1 && hrs >= REMINDER_1_HOURS)       num = 1;
    else if (r1 && !r2 && hrs >= REMINDER_2_HOURS) num = 2;
    else if (r2 && !r3 && hrs >= REMINDER_3_HOURS) num = 3;
    if (!num) continue;

    try {
      sendReminderEmail(rd, settings);
      const col = [null, MC.REMINDER_1, MC.REMINDER_2, MC.REMINDER_3][num];
      masterSh.getRange(i + 1, col + 1).setValue(now);
      syncReminderToInput(ss, sourceTab, requestId, ['', 'REMINDER_1', 'REMINDER_2', 'REMINDER_3'][num], now);
    } catch (err) {
      logError('checkReminderStatuses', err.message, `Request ID: ${requestId}`);
    }
  }
}

// Re-reads the current Stop Sending value from the row's own source INPUT tab
// every time this runs. Returns true (halt reminders) if Stop Sending = YES,
// or if the row can no longer be found there (treated as cancelled).
function isStopSendingOrDeleted(ss, sourceTab, requestId) {
  const inputSh = ss.getSheetByName(sourceTab);
  if (!inputSh || inputSh.getLastRow() < 2) return true;
  const lastRow = inputSh.getLastRow();
  const ids     = inputSh.getRange(2, IC.REQUEST_ID + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === requestId) {
      const stopVal = inputSh.getRange(i + 2, IC.STOP_SENDING + 1).getValue();
      return String(stopVal).trim().toUpperCase() === 'YES';
    }
  }
  return true; // not found — row was deleted from INPUT
}

function masterRowToRequestData(row) {
  return {
    requestId:   row[MC.REQUEST_ID],
    center:      row[MC.CENTER],
    type:        row[MC.TYPE],
    sourceTab:   row[MC.SOURCE_TAB],
    sourceRow:   row[MC.SOURCE_ROW],
    studentName: row[MC.STUDENT_NAME],
    parentName:  row[MC.PARENT_NAME],
    parentEmail: row[MC.PARENT_EMAIL],
    prefLink:    row[MC.PREF_LINK]
  };
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add reminder checker with live Stop Sending re-check"
```

---

### Task 14: Web app GET handler (Section 10)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append**

```javascript


// =============================================================================
// SECTION 10 — WEB APP HANDLER
// =============================================================================

function doGet(e) {
  const requestId = (e && e.parameter && e.parameter.id) ? e.parameter.id.trim() : '';

  if (!requestId) {
    return HtmlService.createHtmlOutput(
      `<body style="font-family:Arial;padding:40px;">
         <p>No request ID provided. Please use the link from your email.</p>
       </body>`
    );
  }

  const req = getRequestById(requestId);

  if (!req) {
    return HtmlService.createHtmlOutput(
      `<body style="font-family:Arial;padding:40px;">
         <p>Request not found. Please contact the center directly.</p>
       </body>`
    );
  }

  if (String(req.responded || '').toUpperCase() === 'YES') {
    return HtmlService.createHtmlOutput(
      `<body style="font-family:Arial;padding:40px;max-width:500px;margin:auto;">
         <h2 style="color:#2E7D32;">✓ Preference Already Received</h2>
         <p>We've already received a schedule preference for <strong>${req.studentName}</strong>. Thank you!</p>
       </body>`
    );
  }

  const settings = getSettings();
  const template = HtmlService.createTemplate(getPreferencePageHtml());
  template.requestId       = requestId;
  template.studentName     = req.studentName;
  template.centerName      = req.center || 'Mathnasium';
  template.type            = req.type;
  template.offeredDaysJson = JSON.stringify(buildOfferedSlotsJson(settings));

  return template.evaluate()
    .setTitle('Schedule Preference — ' + req.studentName)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getRequestById(requestId) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const masterSh = ss.getSheetByName(TAB_MASTER);
  if (!masterSh || masterSh.getLastRow() < 2) return null;
  const data = masterSh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][MC.REQUEST_ID]).trim() !== requestId) continue;
    const row = data[i];
    return {
      requestId:   String(row[MC.REQUEST_ID]),
      center:      String(row[MC.CENTER]),
      type:        String(row[MC.TYPE]),
      studentName: String(row[MC.STUDENT_NAME]),
      parentName:  String(row[MC.PARENT_NAME]),
      parentEmail: String(row[MC.PARENT_EMAIL]),
      responded:   String(row[MC.RESPONDED] || '')
    };
  }
  return null;
}

// Builds { Sunday: [{value, label}, ...], Monday: [...], ... } for the
// client-side day/time dropdowns, computed server-side from SETTINGS.
function buildOfferedSlotsJson(settings) {
  const out = {};
  OFFERED_DAYS.forEach(day => {
    const hours = dayHours(day, settings);
    out[day] = selectableStartTimes(hours.open, hours.close).map(mins => ({
      value: mins,
      label: minutesToTimeLabel(mins)
    }));
  });
  return out;
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add web app doGet handler and offered-slots builder"
```

---

### Task 15: Preference page HTML (Section 11)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

The day dropdown drives the time dropdown client-side using `offeredDaysJson` (computed server-side in Task 14 and embedded via the unescaped `<?!= ?>` scriptlet, since it must stay valid JSON rather than being HTML-escaped). Submission calls `google.script.run.processWebSubmission(...)` directly — no separate `doPost` endpoint, matching how `hold-form`'s acknowledgment page already works.

- [ ] **Step 1: Append**

```javascript


// =============================================================================
// SECTION 11 — PREFERENCE PAGE HTML
// Edit the styles or copy below to change the look of the parent-facing page.
// =============================================================================

function getPreferencePageHtml() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Schedule Preference</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; background: #f5f5f5; color: #222; }
    .wrapper { max-width: 600px; margin: 32px auto; padding: 0 16px 48px; }

    .header { background: #EE3124; color: #fff; padding: 20px 24px; border-radius: 8px 8px 0 0; }
    .header .center-name { font-size: 13px; opacity: 0.85; margin-bottom: 4px; }
    .header h1 { font-size: 20px; font-weight: bold; }

    .card { background: #fff; border-radius: 0 0 8px 8px; padding: 28px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
    .intro { font-size: 14px; line-height: 1.65; margin-bottom: 22px; }

    .pref-block { border: 1px solid #f5c6c2; border-radius: 8px; padding: 18px 18px 6px; margin-bottom: 18px; }
    .pref-title { font-size: 13px; font-weight: bold; color: #C0392B; text-transform: uppercase; letter-spacing: .04em; margin-bottom: 14px; }
    .pref-title .opt { color: #999; font-weight: normal; text-transform: none; letter-spacing: 0; }

    .field { margin-bottom: 14px; }
    .field label { display: block; font-size: 13px; font-weight: bold; margin-bottom: 5px; }
    select, input[type=text] { width: 100%; padding: 9px 10px; font-size: 14px; border: 1px solid #ccc; border-radius: 5px; font-family: inherit; background: #fff; color: #222; }
    select:focus, input:focus { outline: 2px solid #EE3124; border-color: #EE3124; }
    select:disabled { background: #f3f3f3; color: #999; }

    .disclaimer { background: #FDEDEC; border: 1px solid #f5c6c2; border-radius: 6px; padding: 12px 14px; font-size: 12.5px; line-height: 1.6; color: #7a3b35; margin: 6px 0 22px; }

    .submit-btn { width: 100%; padding: 13px; font-size: 15px; font-weight: bold; background: #EE3124; color: #fff; border: none; border-radius: 5px; cursor: pointer; font-family: inherit; }
    .submit-btn:hover    { background: #C9291E; }
    .submit-btn:disabled { background: #aaa; cursor: not-allowed; }

    .error-msg { color: #EE3124; font-size: 12.5px; margin-top: 10px; display: none; }
    .spinner   { text-align: center; color: #666; padding: 10px; display: none; }

    .success-box { background: #e8f5e9; border: 1px solid #a5d6a7; border-radius: 8px; padding: 32px; text-align: center; }
    .success-box h2 { color: #2E7D32; font-size: 22px; margin-bottom: 12px; }
    .success-box p  { color: #444; font-size: 14px; line-height: 1.7; }
  </style>
</head>
<body>
<div class="wrapper">

  <div class="header">
    <div class="center-name"><?= centerName ?></div>
    <h1>Schedule Preference — <?= studentName ?></h1>
  </div>

  <div class="card" id="mainCard">
    <p class="intro">
      <? if (type === 'Continuing Student') { ?>
        As we shift from our summer schedule to our school-year schedule, <?= studentName ?>'s
        current time slot may not carry over automatically. Pick your preferred day and time
        below — most students attend twice a week, so feel free to fill in a second preference too.
      <? } else { ?>
        Welcome back, <?= studentName ?>! We're building our school-year schedule and want to
        reserve a great time for you. Pick your preferred day and time below — most students
        attend twice a week, so feel free to fill in a second preference too.
      <? } ?>
    </p>

    <div class="pref-block">
      <p class="pref-title">Preference 1</p>
      <div class="field"><label for="day1">Preferred Day</label><select id="day1"></select></div>
      <div class="field"><label for="time1">Preferred Time</label><select id="time1"></select></div>
      <div class="field"><label for="note1">Anything else about your flexibility on this day? <span style="font-weight:normal;color:#999;">(optional)</span></label>
        <input type="text" id="note1" placeholder="e.g. anytime after 4:30 works too"></div>
    </div>

    <div class="pref-block">
      <p class="pref-title">Preference 2 <span class="opt">(optional — most students attend twice a week)</span></p>
      <div class="field"><label for="day2">Preferred Day</label><select id="day2"></select></div>
      <div class="field"><label for="time2">Preferred Time</label><select id="time2"></select></div>
      <div class="field"><label for="note2">Anything else about your flexibility on this day? <span style="font-weight:normal;color:#999;">(optional)</span></label>
        <input type="text" id="note2" placeholder="e.g. anytime after 4:30 works too"></div>
    </div>

    <div class="disclaimer">
      We'll do our best to lock in exactly what you choose above. If scheduling gets tight and we
      need to offer a slightly different day or time, we'll reach out personally to find something
      that works — nothing is finalized without talking to you first.
    </div>

    <button class="submit-btn" id="submitBtn" onclick="submitForm()">Submit My Preference</button>
    <div class="spinner" id="spinner">Submitting…</div>
    <div class="error-msg" id="errMsg"></div>

  </div>
</div>

<script>
  var REQUEST_ID    = '<?= requestId ?>';
  var OFFERED_SLOTS = <?!= offeredDaysJson ?>;
  var DAYS          = Object.keys(OFFERED_SLOTS);

  function buildDaySelect(sel) {
    sel.innerHTML = '<option value="">Choose a day…</option>' +
      DAYS.map(function(d) { return '<option value="' + d + '">' + d + '</option>'; }).join('');
  }

  function buildDaySelectOptional(sel) {
    sel.innerHTML = '<option value="">No second preference</option>' +
      DAYS.map(function(d) { return '<option value="' + d + '">' + d + '</option>'; }).join('');
  }

  function populateTimes(daySel, timeSel) {
    if (!daySel.value) {
      timeSel.innerHTML = '<option value="">Pick a day first</option>';
      timeSel.disabled = true;
      return;
    }
    timeSel.disabled = false;
    var opts = OFFERED_SLOTS[daySel.value] || [];
    timeSel.innerHTML = '<option value="">Choose a time…</option>' +
      opts.map(function(o) { return '<option value="' + o.value + '">' + o.label + '</option>'; }).join('');
  }

  function wirePrefBlock(dayId, timeId, optional) {
    var daySel = document.getElementById(dayId);
    var timeSel = document.getElementById(timeId);
    if (optional) buildDaySelectOptional(daySel); else buildDaySelect(daySel);
    populateTimes(daySel, timeSel);
    daySel.addEventListener('change', function() { populateTimes(daySel, timeSel); });
  }

  wirePrefBlock('day1', 'time1', false);
  wirePrefBlock('day2', 'time2', true);

  function submitForm() {
    var day1  = document.getElementById('day1').value;
    var time1 = document.getElementById('time1').value;
    var note1 = document.getElementById('note1').value.trim();
    var day2  = document.getElementById('day2').value;
    var time2 = document.getElementById('time2').value;
    var note2 = document.getElementById('note2').value.trim();

    hideError();
    if (!day1 || !time1) { showError('Please choose a day and time for Preference 1.'); return; }
    if ((day2 && !time2) || (!day2 && time2)) { showError('Please choose both a day and time for Preference 2, or leave both blank.'); return; }

    document.getElementById('submitBtn').disabled = true;
    document.getElementById('spinner').style.display = 'block';

    google.script.run
      .withSuccessHandler(onSuccess)
      .withFailureHandler(onFailure)
      .processWebSubmission(REQUEST_ID, day1, time1, note1, day2, time2, note2);
  }

  function onSuccess(result) {
    document.getElementById('spinner').style.display = 'none';
    if (result && result.error) { onFailure(result.error); return; }
    document.getElementById('mainCard').innerHTML =
      '<div class="success-box">' +
        '<h2>&#10003; Thank You!</h2>' +
        '<p>Your schedule preference has been received.<br>' +
        'We will follow up if we need to offer a different day or time — otherwise, see you this fall!</p>' +
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
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add preference page HTML"
```

---

### Task 16: Submission handling and status sync (Section 12)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

`syncResponseBackToInput` and `syncReminderToInput` both search their target INPUT tab by Request ID rather than trusting the `SOURCE_ROW` number stored in MASTER_REQUESTS, so they stay correct even if rows are inserted or deleted above the original row.

- [ ] **Step 1: Append**

```javascript


// =============================================================================
// SECTION 12 — SUBMISSION HANDLING + STATUS SYNC
// =============================================================================

function processWebSubmission(requestId, day1, time1, note1, day2, time2, note2) {
  try {
    if (!requestId || !day1 || !time1) return { error: 'Missing required fields.' };
    if (!isPrefPairValid(day2, time2)) return { error: 'Preference 2 needs both a day and a time, or leave both blank.' };

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const masterSh = ss.getSheetByName(TAB_MASTER);
    const data     = masterSh.getDataRange().getValues();
    const now      = new Date();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][MC.REQUEST_ID]).trim() !== requestId) continue;

      masterSh.getRange(i + 1, MC.RESPONDED    + 1).setValue('YES');
      masterSh.getRange(i + 1, MC.RESPONDED_TS + 1).setValue(now);
      masterSh.getRange(i + 1, MC.PREF1_DAY    + 1).setValue(day1);
      masterSh.getRange(i + 1, MC.PREF1_TIME   + 1).setValue(time1);
      masterSh.getRange(i + 1, MC.PREF1_NOTE   + 1).setValue(note1 || '');
      masterSh.getRange(i + 1, MC.PREF2_DAY    + 1).setValue(day2 || '');
      masterSh.getRange(i + 1, MC.PREF2_TIME   + 1).setValue(time2 || '');
      masterSh.getRange(i + 1, MC.PREF2_NOTE   + 1).setValue(note2 || '');
      masterSh.getRange(i + 1, MC.STATUS       + 1).setValue('Responded');

      const sourceTab = String(data[i][MC.SOURCE_TAB]).trim();
      syncResponseBackToInput(ss, sourceTab, requestId, now);

      const settings = getSettings();
      const refreshedRow = masterSh.getRange(i + 1, 1, 1, Object.keys(MC).length).getValues()[0];
      const rd = masterRowToRequestData(refreshedRow);
      rd.pref1Day = day1; rd.pref1Time = time1; rd.pref1Note = note1;
      rd.pref2Day = day2; rd.pref2Time = time2; rd.pref2Note = note2;
      sendCdNotification(rd, settings);

      Logger.log(`processWebSubmission ✓  ${requestId}`);
      return { success: true };
    }

    logError('processWebSubmission', `Request ID not found: ${requestId}`, '');
    return { error: 'Request not found. Please contact the center.' };
  } catch (err) {
    logError('processWebSubmission', err.message, `Request ID: ${requestId}`);
    return { error: err.message };
  }
}

function syncResponseBackToInput(ss, sourceTab, requestId, timestamp) {
  const inputSh = ss.getSheetByName(sourceTab);
  if (!inputSh) { logError('syncResponseBackToInput', sourceTab + ' tab missing', requestId); return; }
  const lastRow = inputSh.getLastRow();
  if (lastRow < 2) return;
  const ids = inputSh.getRange(2, IC.REQUEST_ID + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() !== requestId) continue;
    const row = i + 2;
    inputSh.getRange(row, IC.RESPONDED    + 1).setValue('YES');
    inputSh.getRange(row, IC.RESPONDED_TS + 1).setValue(timestamp);
    inputSh.getRange(row, IC.STATUS       + 1).setValue('Responded');
    Logger.log(`syncResponseBackToInput ✓  ${sourceTab} Row ${row}  ${requestId}`);
    return;
  }
  logError('syncResponseBackToInput', `Request ID not found in ${sourceTab}: ${requestId}`, '');
}

function syncReminderToInput(ss, sourceTab, requestId, reminderKey, timestamp) {
  const colMap = { REMINDER_1: IC.REMINDER_1, REMINDER_2: IC.REMINDER_2, REMINDER_3: IC.REMINDER_3 };
  const col    = colMap[reminderKey];
  if (col === undefined) return;
  const inputSh = ss.getSheetByName(sourceTab);
  if (!inputSh) return;
  const lastRow = inputSh.getLastRow();
  if (lastRow < 2) return;
  const ids = inputSh.getRange(2, IC.REQUEST_ID + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === requestId) {
      inputSh.getRange(i + 2, col + 1).setValue(timestamp); return;
    }
  }
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add web submission handler and status sync"
```

---

### Task 17: Dashboard (Section 13)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append**

```javascript


// =============================================================================
// SECTION 13 — DASHBOARD
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
    dashSh.getRange(r, 1, 1, 7).merge()
      .setValue(text).setFontWeight('bold').setFontSize(13)
      .setBackground(CLR_NAVY).setFontColor(CLR_WHITE);
    r++;
  }
  function sectionHeader(text) {
    dashSh.getRange(r, 1, 1, 7).merge()
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

  titleRow(`${centerName} — School Year Schedule Dashboard`);
  dashSh.getRange(r, 1).setValue(`Last refreshed: ${formatDateReadable(now)}`)
    .setFontStyle('italic').setFontColor('#888');
  r += 2;

  const masterSh = ss.getSheetByName(TAB_MASTER);
  if (!masterSh || masterSh.getLastRow() < 2) {
    emptyRow('No requests yet.', '#444');
    dashSh.autoResizeColumns(1, 7);
    return;
  }

  const rows = masterSh.getDataRange().getValues().slice(1);

  // 1. Summary counts
  sectionHeader('📊  Summary');
  function countFor(typeFilter) {
    const subset = typeFilter ? rows.filter(r => r[MC.TYPE] === typeFilter) : rows;
    return {
      sent:      subset.length,
      responded: subset.filter(r => String(r[MC.RESPONDED]).toUpperCase() === 'YES').length
    };
  }
  const totals = countFor(null);
  const holdC  = countFor(TYPE_HOLD);
  const contC  = countFor(TYPE_CONTINUING);
  colHeaders(['', 'Sent', 'Responded', 'Awaiting Response']);
  dataRow(['Combined', totals.sent, totals.responded, totals.sent - totals.responded]);
  dataRow([TYPE_HOLD, holdC.sent, holdC.responded, holdC.sent - holdC.responded]);
  dataRow([TYPE_CONTINUING, contC.sent, contC.responded, contC.sent - contC.responded]);
  r++;

  // 2. Schedule Preferences Received
  const responded = rows.filter(row => String(row[MC.RESPONDED]).toUpperCase() === 'YES');
  sectionHeader(`✓  Schedule Preferences Received (${responded.length})`);
  if (responded.length) {
    colHeaders(['Student', 'Type', 'Pref 1', 'Pref 1 Note', 'Pref 2', 'Pref 2 Note', 'Responded On']);
    responded.forEach(row => {
      const pref1 = row[MC.PREF1_DAY] ? `${row[MC.PREF1_DAY]} ${minutesToTimeLabel(Number(row[MC.PREF1_TIME]))}` : '—';
      const pref2 = row[MC.PREF2_DAY] ? `${row[MC.PREF2_DAY]} ${minutesToTimeLabel(Number(row[MC.PREF2_TIME]))}` : '—';
      dataRow([
        row[MC.STUDENT_NAME], row[MC.TYPE], pref1, row[MC.PREF1_NOTE] || '—',
        pref2, row[MC.PREF2_NOTE] || '—', formatDateReadable(row[MC.RESPONDED_TS])
      ]);
    });
  } else { emptyRow('No responses yet.'); }
  r++;

  // 3. Awaiting Response
  const awaiting = rows.filter(row => String(row[MC.RESPONDED]).toUpperCase() !== 'YES');
  sectionHeader(`⏳  Awaiting Response (${awaiting.length})`);
  if (awaiting.length) {
    colHeaders(['Student', 'Type', 'Parent', 'Email Sent', 'Days Elapsed']);
    awaiting.forEach(row => {
      const sentTs = row[MC.EMAIL_SENT_TS];
      const days = sentTs ? Math.floor((now - new Date(sentTs)) / 86400000) : '—';
      dataRow([row[MC.STUDENT_NAME], row[MC.TYPE], row[MC.PARENT_NAME], formatDateReadable(sentTs), days]);
    });
  } else { emptyRow('Everyone has responded ✓', '#2E7D32'); }
  r++;

  dashSh.autoResizeColumns(1, 7);
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add DASHBOARD tab builder"
```

---

### Task 18: Slot demand grid (Section 14)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

Uses the pure `slotDemandRows`/`occupiedSlots` functions from Task 2, already covered by the local Node tests — this task wires them into the sheet.

- [ ] **Step 1: Append**

```javascript


// =============================================================================
// SECTION 14 — SLOT DEMAND
// =============================================================================

function buildSlotDemand() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings();
  const now      = new Date();

  let sh = ss.getSheetByName(TAB_SLOT_DEMAND);
  if (!sh) sh = ss.insertSheet(TAB_SLOT_DEMAND);
  sh.clearContents();
  sh.clearFormats();

  sh.getRange(1, 1, 1, OFFERED_DAYS.length + 1).merge()
    .setValue(`${settings['CENTER_NAME'] || 'Mathnasium'} — Slot Demand (half-hour blocks occupied)`)
    .setFontWeight('bold').setFontSize(13)
    .setBackground(CLR_NAVY).setFontColor(CLR_WHITE);
  sh.getRange(2, 1).setValue(`Last refreshed: ${formatDateReadable(now)}`)
    .setFontStyle('italic').setFontColor('#888');

  // Union of every offered day's demand rows, sorted ascending
  const rowSet = new Set();
  const dayRows = {};
  OFFERED_DAYS.forEach(day => {
    const hours = dayHours(day, settings);
    const demandRows = slotDemandRows(hours.open, hours.close);
    dayRows[day] = new Set(demandRows);
    demandRows.forEach(m => rowSet.add(m));
  });
  const allRows = Array.from(rowSet).sort((a, b) => a - b);

  const headerRow = 4;
  sh.getRange(headerRow, 1).setValue('Time').setFontWeight('bold')
    .setBackground(CLR_RED_LIGHT).setFontColor('#333');
  OFFERED_DAYS.forEach((day, i) => {
    sh.getRange(headerRow, i + 2).setValue(day).setFontWeight('bold')
      .setBackground(CLR_RED_LIGHT).setFontColor('#333');
  });

  // Count occupancy from MASTER_REQUESTS: every responded row's Pref 1 and
  // (if given) Pref 2 occupies its start slot and the following half-hour.
  const counts = {}; // "day|slotMinutes" -> count
  function addOccupancy(day, startMinutes) {
    if (!day || isNaN(startMinutes)) return;
    occupiedSlots(startMinutes).forEach(slot => {
      const key = day + '|' + slot;
      counts[key] = (counts[key] || 0) + 1;
    });
  }

  const masterSh = ss.getSheetByName(TAB_MASTER);
  if (masterSh && masterSh.getLastRow() >= 2) {
    const rows = masterSh.getDataRange().getValues().slice(1);
    rows.forEach(row => {
      if (String(row[MC.RESPONDED]).toUpperCase() !== 'YES') return;
      addOccupancy(String(row[MC.PREF1_DAY]), Number(row[MC.PREF1_TIME]));
      if (row[MC.PREF2_DAY]) addOccupancy(String(row[MC.PREF2_DAY]), Number(row[MC.PREF2_TIME]));
    });
  }

  allRows.forEach((minutes, idx) => {
    const rowNum = headerRow + 1 + idx;
    sh.getRange(rowNum, 1).setValue(minutesToTimeLabel(minutes));
    OFFERED_DAYS.forEach((day, i) => {
      const cell = sh.getRange(rowNum, i + 2);
      if (!dayRows[day].has(minutes)) {
        cell.setValue('').setBackground('#f3f3f3');
        return;
      }
      const count = counts[day + '|' + minutes] || 0;
      cell.setValue(count);
      if (count > 0) cell.setBackground(CLR_RED_LIGHT);
    });
  });

  sh.autoResizeColumns(1, OFFERED_DAYS.length + 1);
}
```

- [ ] **Step 2: Syntax-check**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add SLOT_DEMAND tab builder"
```

---

### Task 19: Triggers, utilities, error logging (Sections 15–17)

**Files:**
- Modify: `school-year-schedule-form/Code.gs` (append)

- [ ] **Step 1: Append**

```javascript


// =============================================================================
// SECTION 15 — TRIGGERS
// =============================================================================

function installTriggers() {
  removeTriggers();
  ScriptApp.newTrigger('processDirectorInputs').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('checkReminderStatuses').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('refreshDashboardAndSlotDemand').timeBased().everyDays(1).atHour(8).create();

  SpreadsheetApp.getUi().alert(
    '✓ Triggers Installed',
    '• Process new rows — every hour\n' +
    '• Check reminders  — daily at 9am\n' +
    '• Refresh dashboard & slot demand — daily at 8am\n\n' +
    'View: Extensions > Apps Script > Triggers (clock icon)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}


// =============================================================================
// SECTION 16 — UTILITIES
// =============================================================================

function formatDateReadable(date) {
  if (!date) return '';
  try {
    const d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d.getTime())) return String(date);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMMM d, yyyy h:mm a');
  } catch (_) { return String(date); }
}


// =============================================================================
// SECTION 17 — ERROR LOGGING
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
```

- [ ] **Step 2: Syntax-check the complete file**

Run: `node --input-type=commonjs --check < school-year-schedule-form/Code.gs`
Expected: no output

- [ ] **Step 3: Re-run the slot-logic tests one final time**

Run: `node school-year-schedule-form/test/slot-logic.test.js`
Expected: `All slot-logic tests passed.`

- [ ] **Step 4: Commit**

```bash
git add school-year-schedule-form/Code.gs
git commit -m "feat: add triggers, date formatting utility, and error logging"
```

---

### Task 20: Write SETUP.md

**Files:**
- Create: `school-year-schedule-form/SETUP.md`

- [ ] **Step 1: Write the file**

```markdown
# School Year Schedule Preference Form — Setup & Reference Guide

---

## What This Is

A Google Sheets + Apps Script system that:
- Lets center directors (CDs) enter students returning from a summer hold, or students who attended through the summer, into two separate INPUT tabs
- Automatically emails each parent a link to a schedule-preference page (wording differs by which tab the student came from)
- Sends reminder emails at 2, 5, and 8 days if the parent hasn't responded
- Updates the sheet automatically when the parent submits
- Shows a live dashboard of preferences received and a SLOT_DEMAND grid of half-hour capacity, for building the fall schedule

Each center (Teaneck, Englewood) has its own separate Google Sheet deployed in that center's Google account. Emails send automatically from that account.

---

## 1. Sheet Structure

| Tab | Purpose |
|---|---|
| **INPUT_RETURNING_FROM_HOLD** | CD enters students coming back from a summer hold |
| **INPUT_CONTINUING_STUDENTS** | CD enters students who attended through the summer |
| **MASTER_REQUESTS** | Auto-filled canonical record of all requests from both INPUT tabs |
| **SETTINGS** | Center config, email addresses, center hours, web app URL |
| **DASHBOARD** | Auto-refreshed summary + working list of preferences received |
| **SLOT_DEMAND** | Auto-refreshed grid of how many students land in each half-hour slot |
| **ERROR_LOG** | Script errors logged here for troubleshooting |

---

## 2. First-Time Setup (do this once per center)

### Step 1 — Create the Google Sheet

1. Log in to the center's Google account (e.g. `teaneck@mathnasium.com`)
2. Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet
3. Name it: **School Year Schedule — Teaneck** (or Englewood)

### Step 2 — Open the Apps Script editor

1. In the spreadsheet: **Extensions > Apps Script**
2. Delete any existing code in `Code.gs`
3. Paste the entire contents of `Code.gs` from this repository
4. Click **Save**
5. Name the project: **School Year Schedule**

### Step 3 — Run setupWorkbook()

1. In the Apps Script editor, select `setupWorkbook` from the function dropdown
2. Click **Run ▶**
3. Click **Review permissions** → **Allow** (click "Advanced" → "Go to School Year Schedule (unsafe)" on the first run if prompted)
4. Wait for the script to complete — it creates all tabs, writes default SETTINGS, and builds the initial DASHBOARD and SLOT_DEMAND views

### Step 4 — Update the SETTINGS tab

| Key | Teaneck | Englewood |
|---|---|---|
| `CENTER_NAME` | `Mathnasium of Teaneck` | `Mathnasium of Englewood` |
| `CENTER_EMAIL` | `teaneck@mathnasium.com` | `englewood@mathnasium.com` |
| `CD_EMAIL` | `teaneck@mathnasium.com` | `englewood@mathnasium.com` |
| `FROM_NAME` | `Mathnasium of Teaneck` | `Mathnasium of Englewood` |
| `PHONE` | *(update with actual number)* | `201-431-5510` |
| `SUN_OPEN` | `10:00` | `10:00` |
| `SUN_CLOSE` | `14:00` | `13:00` |
| `WEEKDAY_OPEN` | `15:30` | `15:30` |
| `WEEKDAY_CLOSE` | `19:30` | `19:00` |

> `WEB_APP_URL` is filled in manually in Step 5 below — do not edit it until then.
> Hours use 24-hour `HH:MM` format. "Weekday" means Monday–Thursday; neither center offers Friday or Saturday, so those days never appear on the parent-facing form.

### Step 5 — Deploy as Web App

1. In Apps Script: **Deploy > New deployment**
2. Type: **Web app** | Execute as: **Me** | Who has access: **Anyone**
3. Click **Deploy**, copy the deployment URL
4. Paste it into SETTINGS as `WEB_APP_URL`

### Step 6 — Install triggers

Click **⚙ School Year Schedule > Install Triggers** from the top menu.

### Step 7 — Share with Matt

**File > Share > Share with people** — add `mdiamond77@gmail.com` with **Editor** access.

### Step 8 — Repeat for the other center

Log into the other center's Google account and repeat steps 1–7 with that center's SETTINGS values.

---

## 3. Trigger Reference

| Trigger | Function | Schedule |
|---|---|---|
| Process new rows (both INPUT tabs) | `processDirectorInputs` | Every 1 hour |
| Check reminders | `checkReminderStatuses` | Daily at 9am |
| Refresh dashboard & slot demand | `refreshDashboardAndSlotDemand` | Daily at 8am |

To reinstall (e.g. after changing the schedule): **Remove All Triggers**, then **Install Triggers**.

---

## 4. Day-to-Day Workflow

### Adding a student

1. Open the correct INPUT tab — `INPUT_RETURNING_FROM_HOLD` or `INPUT_CONTINUING_STUDENTS`
2. Fill in a new row: Student Name, Parent Name, Parent Email
3. Set **Send Email?** = `YES`
4. Within the hour, the script generates a Request ID, emails the parent, and fills in Status = `Sent`, Preference Link, Email Sent

> **Do not edit Status, Preference Link, Responded?, Responded On, Reminder columns, or Request ID manually** — these are managed by the script.

### Stopping emails for a student

If a student isn't returning, gave a wrong email, or should stop receiving emails for any other reason: set **Stop Sending?** = `YES` and note why in **Stop Reason**. This halts the initial email if not yet sent, or any remaining reminders — checked fresh every time the reminder trigger runs, so it works even after the first email already went out.

### Checking preferences received

Open **DASHBOARD** for the working list of who's responded, with their preferences, plus who's still awaiting a response. Open **SLOT_DEMAND** for a half-hour-by-half-hour capacity view across the week.

### Running manually

Any function is available from the **⚙ School Year Schedule** menu without waiting for the scheduled trigger.

---

## 5. Email Routing Summary

| Email | From | To | CC |
|---|---|---|---|
| Initial preference request | Center Gmail account | Parent | — |
| Reminders 1/2/3 | Center Gmail account | Parent | CD |
| Submission notification | Center Gmail account | CD | — |

---

## 6. How to Edit Email Wording

All email templates live in **Section 7** of `Code.gs`.

- Returning-from-hold initial email: `getInitialEmailSubjectHold()` / `getInitialEmailBodyHold()`
- Continuing-student initial email: `getInitialEmailSubjectContinuing()` / `getInitialEmailBodyContinuing()`
- Reminder email (shared): `getReminderEmailSubject()` / `getReminderEmailBody()`
- CD submission notification: `getCdNotificationSubject()` / `getCdNotificationBody()`
- Shared disclaimer text: `getDisclaimerHtml()`

The parent-facing page copy lives in **Section 11** (`getPreferencePageHtml()`), including the intro paragraph (which also differs by type) and the same disclaimer.

After editing, save the script (Ctrl+S). No trigger reinstallation needed. If you change field IDs in the HTML form, you do need a **new deployment version** (Deploy > Manage deployments > pencil icon > New version > Deploy) for changes to take effect on the live link.

---

## 7. Testing Checklist

### A — Sheet setup
- [ ] All 7 tabs exist: INPUT_RETURNING_FROM_HOLD, INPUT_CONTINUING_STUDENTS, MASTER_REQUESTS, SETTINGS, DASHBOARD, SLOT_DEMAND, ERROR_LOG
- [ ] SETTINGS has all keys filled in, including hours and (after Step 5) WEB_APP_URL

### B — Returning-from-hold email test
- [ ] Add a test row to INPUT_RETURNING_FROM_HOLD with your own email, Send Email? = YES
- [ ] Run **⚙ School Year Schedule > Process New Rows**
- [ ] Confirm: Request ID filled in, Status = Sent, Preference Link is a valid URL
- [ ] Confirm: you received the "Welcome Back!" email with the correct subject and center name
- [ ] Confirm: a row was added to MASTER_REQUESTS with Type = "Returning from Hold"

### C — Continuing-student email test
- [ ] Repeat B in INPUT_CONTINUING_STUDENTS
- [ ] Confirm: the email reads "Let's Set [Student]'s School-Year Schedule" with the shifting-gears framing, not the welcome-back framing

### D — Preference page test
- [ ] Click the Preference Link from either test row
- [ ] Confirm: the day dropdown for Preference 1 only offers Sunday–Thursday
- [ ] Pick a day, confirm the time dropdown updates to that center's actual hours, stopping 1 hour before close
- [ ] Try submitting with Preference 1 blank — confirm it's blocked with an inline error
- [ ] Fill in Preference 1 only, leave Preference 2 blank, submit — confirm success message appears
- [ ] Confirm: MASTER_REQUESTS row updated (Responded? = YES, Pref 1 Day/Time/Note filled in)
- [ ] Confirm: INPUT tab row updated (Responded? = YES, Responded On filled in)
- [ ] Confirm: the CD_EMAIL address received a "Schedule Preference Received" notification

### E — Stop Sending test
- [ ] Add a row with Send Email? = YES and Stop Sending? = YES at the same time
- [ ] Run Process New Rows — confirm no email was sent and the row was skipped
- [ ] Add a second row, send it normally, then set Stop Sending? = YES before any reminder is due
- [ ] Manually change Email Sent in MASTER_REQUESTS to 3 days ago, run **Check Reminder Statuses**
- [ ] Confirm: no reminder was sent (Stop Sending halted it)

### F — Reminder test
- [ ] Add a row, let the initial email send, don't respond
- [ ] In MASTER_REQUESTS, temporarily change Email Sent Date to 49 hours ago
- [ ] Run **Check Reminder Statuses**
- [ ] Confirm: Reminder 1 sent to parent, CD CC'd, MASTER_REQUESTS and INPUT Reminder 1 columns both stamped
- [ ] Restore the Email Sent Date

### G — Dashboard and Slot Demand test
- [ ] Run **⚙ School Year Schedule > Refresh Dashboard**
- [ ] Confirm: Summary counts match what you've sent/responded in testing
- [ ] Confirm: your responded test row appears in "Schedule Preferences Received" with the correct day/time
- [ ] Confirm: SLOT_DEMAND shows a count of 1 in both the chosen start slot and the following half-hour slot for that day, and blank cells outside operating hours

### H — Triggers test
- [ ] Extensions > Apps Script > Triggers — confirm 3 triggers exist: processDirectorInputs (hourly), checkReminderStatuses (daily 9am), refreshDashboardAndSlotDemand (daily 8am)

---

## 8. Troubleshooting

| Problem | Check |
|---|---|
| No email sent after setting Send Email? = YES | Wait up to 1 hour, or run "Process New Rows" manually. Check ERROR_LOG for validation failures. Check that Stop Sending? isn't set to YES. |
| Preference Link says "Request not found" | The Request ID doesn't exist in MASTER_REQUESTS — re-run setupWorkbook or check the row wasn't deleted before Process New Rows ran. |
| Time dropdown is empty | Check SUN_OPEN/SUN_CLOSE/WEEKDAY_OPEN/WEEKDAY_CLOSE in SETTINGS are valid `HH:MM` 24-hour values. |
| Reminders not sending | Check the checkReminderStatuses trigger is installed and that Stop Sending? isn't YES for that row. |
| SLOT_DEMAND shows zero everywhere | Confirm at least one row has Responded? = YES in MASTER_REQUESTS, then re-run Refresh Dashboard. |
| Errors in ERROR_LOG | The error message and context column describe which row and function failed. Fix the data issue and re-run manually. |

---

## 9. Request ID Format

`SCHED-YYYYMMDD-NNNN`, e.g. `SCHED-20260915-0003` = the 3rd request generated (across both INPUT tabs, combined) on September 15, 2026. Unique per center. Do not edit manually.
```

- [ ] **Step 2: Commit**

```bash
git add school-year-schedule-form/SETUP.md
git commit -m "docs: add SETUP.md for school-year-schedule-form"
```

---

### Task 21: Deploy and run the end-to-end test checklist

This task has no code changes. It deploys the finished `Code.gs` to a real sheet and walks through `SETUP.md`'s testing checklist (Section 7) against it, for one center first.

- [ ] **Step 1: Follow SETUP.md Sections 2–3** — create the Teaneck sheet, paste `Code.gs`, run `setupWorkbook()`, update SETTINGS with Teaneck's real values

- [ ] **Step 2: Follow SETUP.md Section 2, Steps 5–7** — deploy as Web App, paste the URL into `WEB_APP_URL`, install triggers, share with Matt

- [ ] **Step 3: Run through SETUP.md Section 7, Testing Checklist A–H** in order, using your own email address for test rows

- [ ] **Step 4: Fix any issues found**, following the Troubleshooting table in SETUP.md Section 8; re-run the affected checklist item after each fix

- [ ] **Step 5: Repeat Steps 1–4 for the Englewood sheet**, using Englewood's SETTINGS values (`SUN_CLOSE` = `13:00`, `WEEKDAY_CLOSE` = `19:00`, phone `201-431-5510`)

- [ ] **Step 6: Report back** — confirm both centers are deployed, both testing checklists pass, and note the two Web App URLs for the record
