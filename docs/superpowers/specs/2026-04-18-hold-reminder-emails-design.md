# Hold Reminder Emails — Design Spec

## Goal

Automatically email each center director a list of students whose hold ends this month, so they can confirm each student's status before the billing cycle rolls over. Also reminds them to add any students going on hold next month.

## Background

Students at Mathnasium can be placed "on hold," pausing their billing temporarily. Center directors need to review holds nearing their end date before the month closes to confirm they're still accurate and to reach out to families as needed.

## Schedule

Emails fire on the **last Monday, Tuesday, and Thursday of each month.**

Implementation: GitHub Actions cron `0 15 * * 1,2,4` fires every Mon/Tue/Thu at 11am ET (3pm UTC). The script checks at runtime whether today falls within the last 7 days of the month. If not, it exits silently. This means up to 3 sends per month (whichever of the last Mon/Tue/Thu exist that month).

Also supports `workflow_dispatch` for manual testing at any time.

## Data Source

**URL:** `https://radius.mathnasium.com/Holds/HoldsReport`

**Filter:** Rows where `Hold End Date` month == current calendar month.

**Columns used:**
| Column | Purpose |
|--------|---------|
| `Student Name` | Display in email |
| `Hold End Date` | Filter + display (format: `M/D/YYYY`) |
| `Guardian Name` | Contact info in email |
| `Guardian Phone` | Contact info in email |
| `Guardian Email` | Contact info in email |
| `Center Name` | Split into per-center emails (`"Englewood"` / `"Teaneck"`) |

## Email Format

**Recipients:** Each center director receives their own email (students from their center only). Matt is CC'd on both.

**Subject:** `Hold Reminders — [Center Name] — [Month Year]`
Example: `Hold Reminders — Teaneck — May 2026`

**Body:**

> Hi [Center Name] Team,
>
> The following students are scheduled to come off hold this month and return to billing for the next cycle. Please confirm each one is still correct before the end of the month.
>
> | Student | Hold End Date | Guardian | Phone | Email |
> |---------|--------------|----------|-------|-------|
> | ... | ... | ... | ... | ... |
>
> **Reminder:** Please also add any students to next month's holds if they have notified you that they are going on hold.
>
> *This email sends automatically on the last Monday, Tuesday, and Thursday of each month.*

**Empty state:** If a center has no students with hold end dates this month, no email is sent for that center.

## Architecture

New repo: `mathnasium-hold-reminders`

```
mathnasium-hold-reminders/
├── main.py           # Orchestrator: download → process → deliver → log
├── download.py       # Playwright: login to Radius, export Holds Report Excel
├── process.py        # pandas: filter by end date month, split by center
├── deliver.py        # HTML email builder + Gmail SMTP (SSL port 465)
├── config.py         # Recipients, URLs, column names, center config
├── run_log.py        # Append run result to run_log.json (dashboard-compatible)
├── requirements.txt
└── .github/
    └── workflows/
        └── hold_reminders.yml
```

### module responsibilities

**`download.py`**
- Playwright (headless Chromium) logs into `https://radius.mathnasium.com`
- Navigates to `/Holds/HoldsReport`
- Selects both centers via the Kendo MultiSelect (`#AllCenterListMultiSelect`)
- Clicks search, then exports to Excel
- Saves file to `input/holds_YYYY-MM.xlsx`

**`process.py`**
- Reads Excel with pandas
- Parses `Hold End Date` as datetime
- Filters: `hold_end_date.month == target_month`
- Splits into per-center dicts: `{"Englewood": [...], "Teaneck": [...]}`
- Each entry: `{name, hold_end_date, guardian_name, guardian_phone, guardian_email}`

**`deliver.py`**
- Builds HTML email per center
- Skips centers with no matching students
- Sends via `smtplib` + SSL on port 465 (Gmail App Password)

**`main.py`**
- Checks if today is in last 7 days of month (when trigger == "auto"); skips if not
- Accepts `--trigger manual` to bypass date check
- Orchestrates download → process → deliver → log

**`config.py`**
```python
RADIUS_LOGIN_URL = "https://radius.mathnasium.com"
HOLDS_REPORT_URL = "https://radius.mathnasium.com/Holds/HoldsReport"

CENTERS = {
    "Englewood": {"radius_id": "2428", "recipient": "englewood@mathnasium.com"},
    "Teaneck":   {"radius_id": "2871", "recipient": "teaneck@mathnasium.com"},
}
CC_RECIPIENT = "matt.diamond@mathnasium.com"

# Column names (verified from real Radius export 4/18/2026)
COL_STUDENT_NAME   = "Student Name"
COL_HOLD_END_DATE  = "Hold End Date"
COL_GUARDIAN_NAME  = "Guardian Name"
COL_GUARDIAN_PHONE = "Guardian Phone"
COL_GUARDIAN_EMAIL = "Guardian Email"
COL_CENTER         = "Center Name"
```

## GitHub Actions Workflow

```yaml
name: Hold Reminder Emails

on:
  schedule:
    - cron: "0 15 * * 1,2,4"   # Mon/Tue/Thu at 11am ET (3pm UTC)
  workflow_dispatch:

jobs:
  run:
    runs-on: ubuntu-latest
    permissions:
      contents: write

    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      - name: Install dependencies
        run: |
          pip install -r requirements.txt
          playwright install chromium
      - name: Run automation
        env:
          RADIUS_USERNAME: ${{ secrets.RADIUS_USERNAME }}
          RADIUS_PASSWORD: ${{ secrets.RADIUS_PASSWORD }}
          SMTP_USER: ${{ secrets.SMTP_USER }}
          SMTP_PASSWORD: ${{ secrets.SMTP_PASSWORD }}
        run: python main.py
      - name: Commit run log
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add run_log.json
          git diff --cached --quiet || git commit -m "chore: update run log [skip ci]"
          git push
```

## Dashboard Integration

- New entry in `automation-dashboard/server.py` SCRIPTS dict: `"hold-reminders"`
- New entry in REPORTS list with `run_log_path` pointing to `run_log.json`
- Entry in `future_projects.json` updated to `"status": "Live"` once deployed

## GitHub Secrets Required

Same secrets as other automations (already configured on the org/repo):
- `RADIUS_USERNAME`, `RADIUS_PASSWORD`
- `SMTP_USER`, `SMTP_PASSWORD`

## Error Handling

- If Radius download fails: logs error, exits non-zero (Actions marks run as failed)
- If a center has zero matching students: skips that center's email silently
- If SMTP fails: logs error, exits non-zero
- All runs (success and failure) appended to `run_log.json`
