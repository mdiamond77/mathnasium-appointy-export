# Student Hold Workflow — Setup & Reference Guide

---

## What This Is

A Google Sheets + Google Forms + Apps Script system that:
- Lets center directors (CDs) enter hold details in a spreadsheet
- Automatically emails parents a prefilled acknowledgment form
- Sends reminder emails at 24h, 3 days, and 7 days if the form isn't signed
- Updates the sheet automatically when the parent signs
- Shows a live dashboard of unsigned holds, upcoming start/end dates, and Radius to-dos

Each center (Teaneck, Englewood) has its own separate Google Sheet deployed in that center's Google account. Emails send automatically from that account.

---

## 1. Sheet Structure

Each sheet has these tabs:

| Tab | Purpose |
|---|---|
| **INPUT** | CD enters hold info here; set "Send Email?" = YES to trigger |
| **MASTER_HOLDS** | Auto-filled canonical record of all holds |
| **SETTINGS** | Center config, email addresses, form IDs |
| **DASHBOARD** | Auto-refreshed summary view |
| **FORM_RESPONSES** | Auto-filled with Google Form submissions |
| **ERROR_LOG** | Script errors logged here for troubleshooting |

---

## 2. First-Time Setup (do this once per center)

### Step 1 — Create the Google Sheet

1. Log in to the center's Google account (e.g. `teaneck@mathnasium.com`)
2. Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet
3. Name it: **Student Hold Workflow — Teaneck** (or Englewood)

### Step 2 — Open the Apps Script editor

1. In the spreadsheet: **Extensions > Apps Script**
2. Delete any existing code in `Code.gs`
3. Paste the entire contents of `Code.gs` from this repository
4. Click **Save** (floppy disk icon or Ctrl+S)
5. Name the project: **Hold Workflow**

### Step 3 — Run setupWorkbook()

1. In the Apps Script editor, select `setupWorkbook` from the function dropdown
2. Click **Run ▶**
3. When prompted, click **Review permissions** → **Allow**
   - You may need to click "Advanced" → "Go to Hold Workflow (unsafe)" on the first run
4. Wait for the script to complete — it will:
   - Create all tabs with formatted headers
   - Write default SETTINGS values
   - Create the Google Form and link it to the sheet
   - Store all form item IDs in SETTINGS
   - Build the initial dashboard
5. A dialog will appear confirming setup is complete

### Step 4 — Update the SETTINGS tab

Open the SETTINGS tab and verify/update these values:

| Key | Teaneck | Englewood |
|---|---|---|
| `CENTER_NAME` | `Mathnasium of Teaneck` | `Mathnasium of Englewood` |
| `CENTER_EMAIL` | `teaneck@mathnasium.com` | `englewood@mathnasium.com` |
| `CD_EMAIL` | `teaneck@mathnasium.com` | `englewood@mathnasium.com` |
| `FROM_NAME` | `Mathnasium of Teaneck` | `Mathnasium of Englewood` |
| `PHONE` | *(update with actual number)* | `201-431-5510` |
| `ADMIN_CC` | `matt.diamond@mathnasium.com` | `matt.diamond@mathnasium.com` |
| `BCC_EMAIL` | `bccradius@mathnasium.com` | `bccradius@mathnasium.com` |

> The `FORM_ID`, `FORM_URL`, and all `ITEM_*` rows are filled automatically by `setupWorkbook()` — do not edit them.

### Step 5 — Install triggers

1. Go back to the spreadsheet
2. Click **⚙ Hold Workflow > Install Triggers** from the top menu
3. A confirmation dialog will appear listing all active triggers

### Step 6 — Share with Matt

1. In the spreadsheet: **File > Share > Share with people**
2. Add `mdiamond77@gmail.com` (or Matt's Mathnasium account) with **Editor** access

### Step 7 — Repeat for the other center

Log into `englewood@mathnasium.com` and repeat steps 1–6. On step 4, update SETTINGS with Englewood's values.

---

## 3. Google Form Structure

The form is created automatically by `setupWorkbook()`. You do not need to build it manually.

**Form title:** Temporary Hold Acknowledgment — [Center Name]

**Fields created (in order):**

| Field | Type | Prefilled? | Required? |
|---|---|---|---|
| Hold ID | Short answer | ✅ Yes | ✅ Yes |
| Student Name | Short answer | ✅ Yes | ✅ Yes |
| Parent Name | Short answer | ✅ Yes | ✅ Yes |
| Center | Short answer | ✅ Yes | ✅ Yes |
| Hold Start Date | Short answer | ✅ Yes | ✅ Yes |
| Expected Return Date | Short answer | ✅ Yes | ✅ Yes |
| Tuition Billing Resume Month | Short answer | ✅ Yes | ✅ Yes |
| Acknowledgment checkbox | Checkbox | ❌ No | ✅ Yes |
| Parent Full Name (Typed Signature) | Short answer | ❌ No | ✅ Yes |
| Today's Date | Date | ❌ No | ✅ Yes |
| Optional Comments | Paragraph | ❌ No | ❌ No |

> The form URL for each parent is generated with fields pre-filled so they only need to check the box, type their name, and submit.

To view the form:
- Go to SETTINGS tab → find `FORM_URL`
- Or: **⚙ Hold Workflow > Show Form Item IDs (diagnostic)**

---

## 4. Trigger Setup

Triggers are installed by running **⚙ Hold Workflow > Install Triggers**.

| Trigger | Function | Schedule |
|---|---|---|
| Process new rows | `processDirectorInputs` | Every 1 hour |
| Check reminders | `checkReminderStatuses` | Daily at 9am |
| Refresh dashboard | `buildDashboard` | Daily at 8am |
| Form submission | `processFormSubmission` | On form submit |

To view or edit triggers manually:
**Extensions > Apps Script > Triggers** (clock icon on the left sidebar)

To reinstall (e.g., after changing the form):
Run **⚙ Hold Workflow > Remove All Triggers**, then **Install Triggers**.

---

## 5. Day-to-Day Workflow

### Adding a hold

1. Open the **INPUT** tab in the center's sheet
2. Fill in a new row:
   - Student Name, Parent Name, Parent Email
   - Hold Start Date, Expected Return Date, Billing Resume Month
   - Notes (optional)
3. In the **"Send Email?"** column, type `YES`
4. The script will pick it up within the hour and:
   - Generate a unique Hold ID
   - Email the parent with a prefilled acknowledgment link
   - Fill in Status = `Sent`, Form Link, Email Sent Timestamp, Hold ID

> **Do not change Status, Form Link, or Hold ID manually** — these are managed by the script.

### Marking Radius as updated

After you update Radius for a signed hold, mark the **Radius Updated** column in the INPUT tab. The dashboard will stop showing that hold in the "Signed — Radius Not Yet Updated" section.

### Running manually

You can trigger any function manually from the **⚙ Hold Workflow** menu without waiting for the scheduled trigger.

---

## 6. Email Routing Summary

| Email | From | To | CC | BCC |
|---|---|---|---|---|
| Initial acknowledgment | Center Gmail account | Parent | — | bccradius |
| Reminder 1 (24h) | Center Gmail account | Parent | CD | bccradius |
| Reminder 2 (3 days) | Center Gmail account | Parent | CD | bccradius |
| Reminder 3 (7 days) | Center Gmail account | Parent | CD + Admin | bccradius |

---

## 7. Testing Checklist

Run through this after setup to confirm everything works.

### A — Sheet setup
- [ ] All 6 tabs exist: INPUT, MASTER_HOLDS, SETTINGS, DASHBOARD, FORM_RESPONSES, ERROR_LOG
- [ ] SETTINGS tab has all keys filled in, including FORM_ID, FORM_URL, and all ITEM_* rows
- [ ] DASHBOARD tab shows the title and "No hold records yet" message

### B — End-to-end email test
- [ ] Add a test row to INPUT with your own email as Parent Email
- [ ] Set Send Email? = YES
- [ ] Run **⚙ Hold Workflow > Process New Rows** manually
- [ ] Confirm: Hold ID filled in, Status = Sent, Form Link is a valid URL, Email Sent Timestamp set
- [ ] Confirm: a row was added to MASTER_HOLDS
- [ ] Confirm: you received the initial email from the center's Gmail address
- [ ] Confirm: bccradius@mathnasium.com received the BCC copy

### C — Form submission test
- [ ] Click the Form Link from the test row
- [ ] Verify the form opens with Student Name, Parent Name, Center, and hold dates pre-filled
- [ ] Fill in the acknowledgment checkbox, typed name, and date — submit
- [ ] Confirm: MASTER_HOLDS row updated (Signed? = YES, Signature Timestamp, Status = Signed)
- [ ] Confirm: INPUT row updated (Signed? = YES, Signature Timestamp, Status = Signed)
- [ ] Confirm: FORM_RESPONSES tab has the submission

### D — Reminder test
- [ ] Add a second test row with Send Email? = YES, let it send
- [ ] In MASTER_HOLDS, temporarily change Email Sent Timestamp to 25 hours ago
- [ ] Run **⚙ Hold Workflow > Check Reminder Statuses** manually
- [ ] Confirm: Reminder 1 sent to parent, CD CC'd, bccradius BCC'd
- [ ] Confirm: MASTER_HOLDS Reminder 1 Sent column has a timestamp
- [ ] Confirm: INPUT tab Reminder 1 Sent column has a timestamp
- [ ] Restore the Email Sent Timestamp

### E — Dashboard test
- [ ] Run **⚙ Hold Workflow > Refresh Dashboard**
- [ ] Confirm: unsigned hold appears in "Unsigned Acknowledgments" section
- [ ] Confirm: signed hold does NOT appear there
- [ ] Confirm: summary counts are correct

### F — Error handling test
- [ ] Add a row with Send Email? = YES but no Parent Email
- [ ] Run Process New Rows
- [ ] Confirm: row was skipped (Status still blank), error appears in ERROR_LOG tab

### G — Triggers test
- [ ] Go to Extensions > Apps Script > Triggers
- [ ] Confirm 4 triggers exist: processDirectorInputs (hourly), checkReminderStatuses (daily 9am), buildDashboard (daily 8am), processFormSubmission (on form submit)

---

## 8. How to Edit Email Wording

All email templates live in **Section 6** of `Code.gs`. You do not need to touch any other section to change what emails say.

**To edit the initial email:**
- Subject line → `getInitialEmailSubject()`
- Body → `getInitialEmailBody()`

**To edit reminder emails:**
- Subject line → `getReminderEmailSubject()`
- Body → `getReminderEmailBody()`
  - The urgency block (red warning text) in `getReminderEmailBody()` only renders for Reminder 3

**Available template variables** (use inside the body functions):

| Variable | Value |
|---|---|
| `holdData.studentName` | Student's full name |
| `holdData.parentName` | Parent's full name |
| `holdData.parentEmail` | Parent's email |
| `holdData.holdStart` | Hold start date (raw — wrap in `formatDateReadable()`) |
| `holdData.expectedReturn` | Expected return date |
| `holdData.billingResume` | Billing resume month (e.g. "July 2026") |
| `holdData.formLink` | Full prefilled form URL |
| `settings['FROM_NAME']` | Center display name |
| `settings['PHONE']` | Center phone number |
| `settings['CENTER_EMAIL']` | Center email address |

**Example:** To add a phone number to the initial email, find `getInitialEmailBody()` and edit the sign-off block:

```javascript
<p>
  Thank you,<br>
  <strong>${settings['FROM_NAME']}</strong><br>
  ${settings['PHONE']}<br>
  ${settings['CENTER_EMAIL']}
</p>
```

After editing, save the script (Ctrl+S). No trigger reinstallation needed.

---

## 9. Reminder Timing Reference

| Reminder | Fires when... | CC |
|---|---|---|
| Reminder 1 | 24 hours after initial email, if not signed | CD only |
| Reminder 2 | 72 hours (3 days) after initial email, if not signed | CD only |
| Reminder 3 | 168 hours (7 days) after initial email, if not signed | CD + Admin |

Reminders fire sequentially — at most one per hold per daily check. If the daily trigger was missed for several days, they will still send in order (R1, then R2 the next day, then R3).

---

## 10. Troubleshooting

| Problem | Check |
|---|---|
| No email sent after setting YES | Wait up to 1 hour (trigger runs hourly). Or run "Process New Rows" manually. Check ERROR_LOG for validation failures. |
| Form link says "FORM_ID missing" | Run setupWorkbook() again — the form wasn't created. |
| Form responses not updating MASTER_HOLDS | Check that the form submit trigger is installed (Extensions > Apps Script > Triggers). Re-run installTriggers() if missing. |
| Reminders not sending | Check that checkReminderStatuses trigger is installed. Verify Email Sent Timestamp is filled in for the hold. |
| "Form Responses 1" tab instead of "FORM_RESPONSES" | Rename it manually to `FORM_RESPONSES`. The auto-rename occasionally fails on the first run. |
| Errors in ERROR_LOG | The error message and context column describe which row and function failed. Fix the data issue and re-run manually. |
| Duplicate emails | Check that installTriggers() wasn't run multiple times without removing first. Run "Remove All Triggers" then "Install Triggers". |

---

## 11. Hold ID Format

Hold IDs are auto-generated in the format: `HOLD-YYYYMMDD-NNNN`

Example: `HOLD-20260601-0003` = the 3rd hold entered on June 1, 2026.

Hold IDs are unique per center and are used to match form submissions back to the correct MASTER_HOLDS row. **Do not edit Hold IDs manually.**
