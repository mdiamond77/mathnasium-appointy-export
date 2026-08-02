# School Year Schedule Preference Form

## Goal

Build a new Google Sheets + Apps Script system, modeled on `hold-form/`, that collects preferred day/time for the upcoming school year from two groups of parents:

1. **Students returning from a summer hold** — welcomed back, asked to help set their schedule.
2. **Continuing students** who attended through the summer — told the center is shifting from its summer schedule to its school-year schedule and asked for their preference too.

Both groups get a link to the same parent-facing preference page, where they pick up to 2 preferred day/time slots within that center's hours. Reminders chase non-responders. A dashboard gives each center director (CD) a working list of preferences to build the fall schedule from.

## Background

The existing `hold-form/` system (see `hold-form/SETUP.md` and `hold-form/Code.gs`) already establishes the pattern this reuses: a Sheet with a CD-editable INPUT tab, a script-managed MASTER tab, a SETTINGS tab, a web app for the parent-facing page (not a native Google Form, despite what `hold-form/SETUP.md` says — the actual implementation is a custom `doGet`/`doPost` HTML page), and hourly/daily time-driven triggers. This project follows that same architecture but is **entirely standalone** — it does not read from or write to the hold-form sheet or its data.

## Deployment

One Sheet + Apps Script Web App per center, deployed the same way as hold-form: created and run from that center's own Google account (`teaneck@mathnasium.com`, `englewood@mathnasium.com`), shared with Matt as Editor. Project name: **`school-year-schedule-form/`** (repo directory, mirrors `hold-form/`), containing `Code.gs` and a `SETUP.md` written the same way as `hold-form/SETUP.md`.

## Sheet Structure

| Tab | Purpose |
|---|---|
| **INPUT_RETURNING_FROM_HOLD** | CD enters students coming back from a summer hold |
| **INPUT_CONTINUING_STUDENTS** | CD enters students who attended through the summer |
| **MASTER_REQUESTS** | Auto-filled canonical record of all requests from both INPUT tabs |
| **SETTINGS** | Center config, email addresses, center hours, form/web app URL |
| **DASHBOARD** | Auto-refreshed summary + working list of preferences received |
| **ERROR_LOG** | Script errors logged here for troubleshooting |

## INPUT Tabs

Both `INPUT_RETURNING_FROM_HOLD` and `INPUT_CONTINUING_STUDENTS` share an identical column layout:

| Col | Field | Type | Notes |
|---|---|---|---|
| A | Student Name | Text | Required |
| B | Parent Name | Text | Required |
| C | Parent Email | Text | Required |
| D | Send Email? | YES/NO dropdown | CD sets YES to trigger the initial email |
| E | Stop Sending? | YES/NO dropdown | See "Stop Sending" below |
| F | Stop Reason | Free text | CD context only (e.g. "Not returning," "Wrong email") — never shown to the parent |
| G | Status | Auto-filled, grey | blank → `Sent` → `Responded` |
| H | Preference Link | Auto-filled, grey | Prefilled web app URL |
| I | Email Sent | Auto-filled, grey | Timestamp of initial email |
| J | Responded? | Auto-filled, grey | YES once parent submits |
| K | Responded On | Auto-filled, grey | Timestamp of submission |
| L | Reminder 1 Sent | Auto-filled, grey | Timestamp |
| M | Reminder 2 Sent | Auto-filled, grey | Timestamp |
| N | Reminder 3 Sent | Auto-filled, grey | Timestamp |
| O | Request ID | Auto-filled, grey | Format `SCHED-YYYYMMDD-NNNN`, unique per center |

No preference detail columns in INPUT — full answers (day/time/notes) live in MASTER_REQUESTS and DASHBOARD, keeping INPUT focused on status only. This mirrors how `hold-form` keeps Signed?/Signed On in INPUT but full hold detail only in MASTER_HOLDS.

### Stop Sending

`Stop Sending? = YES` halts all future emails for that row, for any reason (not returning, wrong contact info, handled manually, etc.):

- `processDirectorInputs` skips rows where Stop Sending = YES before sending the initial email.
- `checkReminderStatuses` re-reads the current Stop Sending value from the source INPUT tab for every row before sending any reminder — so setting it to YES after the initial email has already gone out still stops all remaining reminders. This is a re-check every run (not a one-time snapshot), matching how `hold-form` treats deleted INPUT rows as cancelled via `getActiveInputHoldIds`.

## MASTER_REQUESTS Columns

```
REQUEST_ID, CENTER, TYPE, SOURCE_TAB, SOURCE_ROW,
STUDENT_NAME, PARENT_NAME, PARENT_EMAIL,
PREFERENCE_LINK, EMAIL_SENT, EMAIL_SENT_TS,
RESPONDED, RESPONDED_TS,
REMINDER_1, REMINDER_2, REMINDER_3,
PREF1_DAY, PREF1_TIME, PREF1_NOTE,
PREF2_DAY, PREF2_TIME, PREF2_NOTE,
STATUS
```

`TYPE` is `"Returning from Hold"` or `"Continuing Student"`, set from which INPUT tab the row came from (`SOURCE_TAB`). This drives which initial email template is used and is shown in the dashboard.

## SETTINGS

Same keys as `hold-form` (`CENTER_NAME`, `CENTER_EMAIL`, `CD_EMAIL`, `FROM_NAME`, `PHONE`, `BCC_EMAIL`, `WEB_APP_URL`), plus new hours keys used to compute the parent-facing day/time dropdowns:

| Key | Teaneck | Englewood |
|---|---|---|
| `SUN_OPEN` | `10:00` | `10:00` |
| `SUN_CLOSE` | `14:00` | `13:00` |
| `WEEKDAY_OPEN` | `15:30` | `15:30` |
| `WEEKDAY_CLOSE` | `19:30` | `19:00` |

`WEEKDAY` means Monday–Thursday for both centers. Neither center offers Friday or Saturday, so those days are never presented. Storing hours in SETTINGS (rather than hardcoding) means a future hours change is a spreadsheet edit, not a code change.

## Parent-Facing Web App Page

Custom `doGet`/`doPost` HTML page (styled like the hold acknowledgment page — Mathnasium red/navy branding), reached via a prefilled link (`WEB_APP_URL?id=REQUEST_ID`) unique per request.

Shows: center name, student name, a short intro (varies by TYPE — see Email Templates), then:

- **Preference 1 (required):**
  - Day dropdown — options are Sunday, Monday, Tuesday, Wednesday, Thursday (only days that center is open)
  - Time dropdown — populated based on the chosen day, computed server-side from SETTINGS: 30-minute increments from that day's open time up to **1 hour before** that day's close time (e.g. Englewood Sunday: 10:00, 10:30, 11:00, 11:30, 12:00 — closes at 1:00, so 12:00 is the last option)
  - Optional free-text note: "Anything else about your flexibility on this day?"
- **Preference 2 (optional):** same three fields, parent may leave entirely blank
- A disclaimer near the fields and in both emails: *these are preferences, not guarantees — if we need to offer a different day or time, we'll be in touch.*

On submit (`processWebSubmission`):
- Validates Preference 1 is filled in (day + time); Preference 2 fields are optional but if either day or time is given, both are required together.
- Writes preferences + `RESPONDED = YES` + timestamp to MASTER_REQUESTS and back to the source INPUT tab's Responded?/Responded On columns (same sync pattern as `syncStatusBackToInputTab` in hold-form).
- Shows a thank-you confirmation on the page.
- Sends a short notification email to `CD_EMAIL`: student name, type, and both preferences (day/time/note) as submitted.

### Computed Day/Time Table (for reference)

| Center | Day | Time options |
|---|---|---|
| Englewood | Sunday | 10:00, 10:30, 11:00, 11:30, 12:00 |
| Englewood | Mon–Thu | 3:30, 4:00, 4:30, 5:00, 5:30, 6:00 |
| Teaneck | Sunday | 10:00, 10:30, 11:00, 11:30, 12:00, 12:30, 1:00 |
| Teaneck | Mon–Thu | 3:30, 4:00, 4:30, 5:00, 5:30, 6:00, 6:30 |

## Email Templates

All emails use the same Mathnasium red/navy visual style as `hold-form`'s email templates (Section 6 of its `Code.gs`).

### Initial Email — Returning from Hold

- Subject: `Welcome Back! Help Us Plan [Student]'s Fall Schedule — Mathnasium of [Center]`
- Body: welcomes the family back from summer, notes the center is building its school-year schedule and filling up quickly, asks for their preferred day(s)/time(s) via the link, includes the "preferences aren't guaranteed" disclaimer.

### Initial Email — Continuing Student

- Subject: `Let's Set [Student]'s School-Year Schedule — Mathnasium of [Center]`
- Body: notes the center is shifting from its summer schedule to its school-year schedule, explains the current summer time slot may not carry over automatically, asks for their preferred day(s)/time(s) via the link, same disclaimer.

Both templates share the same details/CTA structure (button linking to the preference page) — only the opening framing paragraph and subject line differ, selected by `holdData.type` (or equivalent) at send time.

### Reminder Email (shared, both types)

- Subject: `Reminder: We Still Need [Student]'s Schedule Preference — Mathnasium of [Center]`
- Body: generic "we haven't heard back yet about your schedule preference" — no hold/continuing-specific framing, same CTA button and disclaimer.
- Fires per request if not yet responded and Stop Sending ≠ YES at the time the daily reminder check runs:

| Reminder | Fires | CC |
|---|---|---|
| Reminder 1 | 2 days after initial email | CD |
| Reminder 2 | 5 days after initial email | CD |
| Reminder 3 | 8 days after initial email | CD |

No admin CC on any reminder (unlike `hold-form`'s Reminder 3, which CCs admin).

### CD Submission Notification

- Subject: `Schedule Preference Received — [Student]`
- Body: student name, type, Preference 1 (day/time/note), Preference 2 if given, link to the row in MASTER_REQUESTS or a summary — sent to `CD_EMAIL` only, no BCC needed since this isn't parent-facing.

## Dashboard

Single combined `DASHBOARD` tab (not split into separate tabs per type), refreshed daily via trigger and on-demand via menu:

1. **Summary counts** — Total Sent, Responded, Awaiting Response, Stopped, broken out by Type (Returning from Hold / Continuing Student) and combined.
2. **Schedule Preferences Received** — the working table a CD uses to build the fall schedule: Student, Type, Pref 1 Day/Time, Pref 1 Note, Pref 2 Day/Time, Pref 2 Note, Responded On.
3. **Awaiting Response** — Student, Type, Parent, Email Sent date, days elapsed — so a CD can see who to chase or call directly.

## Triggers

Same pattern as `hold-form`, installed via **⚙ School Year Schedule > Install Triggers**:

| Trigger | Function | Schedule |
|---|---|---|
| Process new rows (both INPUT tabs) | `processDirectorInputs` | Every 1 hour |
| Check reminders | `checkReminderStatuses` | Daily |
| Refresh dashboard | `buildDashboard` | Daily |
| Form submission | `doPost` / `processWebSubmission` | On web app submit |

Menu (`onOpen`): **⚙ School Year Schedule** with Process New Rows, Check Reminder Statuses, Refresh Dashboard, Setup Workbook (first-time), Install Triggers, Remove All Triggers — same structure as `hold-form`'s menu.

## Validation

- Student Name, Parent Name, Parent Email required before Send Email? = YES is processed; missing/invalid rows are skipped and logged to ERROR_LOG (same as hold-form).
- Parent Email must contain `@`.
- On the parent-facing page: Preference 1 day + time both required; Preference 2 day/time are optional as a pair (if one is given, both must be).

## Out of Scope

- No integration with Radius (this is purely a Sheets/email/web-app workflow, no scraping).
- No integration with the existing `hold-form/` sheet or its MASTER_HOLDS data — a student appearing in both systems is coincidental and not cross-referenced.
- No SMS/text reminders — email only, matching `hold-form`.

## Files

- `school-year-schedule-form/Code.gs` — full Apps Script implementation
- `school-year-schedule-form/SETUP.md` — setup/reference guide, written the same way as `hold-form/SETUP.md`
