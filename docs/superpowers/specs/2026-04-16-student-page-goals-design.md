# Student Page Goals Automation — Design Spec

**Date:** 2026-04-16  
**Status:** Approved

---

## Overview

A monthly automation that downloads the Digital Workout Plan Report from Radius, calculates personalized page goals for each student, delivers the formatted Excel report by email and Google Drive, and logs run history to a reports dashboard on the existing automation dashboard.

---

## Repository

New repo: `mathnasium-page-goals`  
Pattern mirrors `radius-cc-lists` and `radius-morning-briefing`.

---

## Architecture

### Modules

**`download.py`**  
Uses Playwright (headless Chromium) to log into Radius and download the Digital Workout Plan Report for the previous calendar month. Reuses the same login approach as `radius-cc-lists` (username/password, no OAuth). Saves the file to `input/Digital_Workout_Plan_YYYY_MM.xlsx`.

**`process.py`**  
Reads the downloaded Excel file with pandas. Extracts four columns: `Date`, `Student Name`, `Pages Completed`, `Center`. Applies calculation logic (see below). Outputs a formatted `Student_Page_Goals_YYYY_MM.xlsx` to `output/` using openpyxl.

**`deliver.py`**  
- Emails the output file to all recipients as an attachment via Gmail SMTP (SSL, port 465), reusing the `SMTP_USER`/`SMTP_PASSWORD` secrets from `radius-morning-briefing`.
- Uploads the output file to the shared Google Drive folder using a Google service account.
- Returns the public Drive link for storage in the run log.

**`main.py`**  
Orchestrates download → process → deliver in sequence. Accepts a `--month YYYY-MM` flag to reprocess any past month (defaults to previous month). On completion (success or failure), writes a record to `run_log.json` and commits it back to the repo.

---

## Calculation Logic

For each student:

1. Filter sessions: keep only rows where `Pages Completed > 0` and is not null.
2. Sort by date descending; take the most recent 10 qualifying sessions.
3. Average = sum of session pages ÷ number of sessions (rounded to 2 decimal places).
4. Page Goal = `MIN(Average × 1.20, MAX(session_pages) − 0.01)` (rounded to 2 decimal places).
5. Students with 0 qualifying sessions are excluded from output.
6. Students with fewer than 10 sessions use all available qualifying sessions.

Center assignment: determined by the `Center` field. Values containing "Englewood" → Englewood sheet; values containing "Teaneck" → Teaneck sheet.

---

## Output Format

**Filename:** `Student_Page_Goals_YYYY_MM.xlsx`

**Two worksheets:** Englewood, Teaneck — each sorted alphabetically by student name.

**14 columns:**

| # | Column | Width | Alignment |
|---|---|---|---|
| 1 | Student Name | 30 | Left |
| 2 | Number of Sessions | 18 | Center |
| 3–12 | Session 1–10 (oldest → most recent) | 10 each | Center |
| 13 | Average Pages | 14 | Center |
| 14 | Page Goal | 12 | Center |

**Formatting:**
- Header row: blue background (`#4472C4`), white text, bold, centered
- Session columns (3–12): light gray background (`#E7E6E6`)
- Average/Goal columns (13–14): light yellow background (`#FFF2CC`)
- Freeze header row and first column

---

## Delivery

**Email recipients:**
- Matt Diamond — matt.diamond@mathnasium.com
- Elizabeth Anacleto (Teaneck Center Director)
- Samba Taha (Englewood Center Director)

**Google Drive folder:** `1NcVaeoFtyJkJfy6-GtLrlMxoN5cT_gyY`  
Folder is shared with recipients. Upload uses a Google service account; the resulting Drive link is stored in `run_log.json`.

---

## Run Log

`run_log.json` in the repo root. Array of run records; newest appended last.

```json
[
  {
    "timestamp": "2026-04-01T06:12:43Z",
    "trigger": "auto",
    "month": "2026-03",
    "status": "success",
    "output_file": "Student_Page_Goals_2026_03.xlsx",
    "drive_link": "https://drive.google.com/...",
    "error": null
  }
]
```

`trigger` is `"auto"` (GitHub Actions cron) or `"manual"` (dashboard button).  
On failure, `status` is `"error"` and `error` contains the exception message.  
After each GitHub Actions run, the workflow commits the updated `run_log.json` back to the repo.

---

## Scheduling

**GitHub Actions cron:** 1st of every month at 11:00 AM UTC (6:00 AM ET).  
**`workflow_dispatch`:** enabled so the workflow can be triggered manually from GitHub if needed.  
**Dashboard button:** runs the script locally (same `main.py`), streams live output in the browser output panel, writes to `run_log.json` directly.

---

## Reports Dashboard

New `/reports` page on the existing automation dashboard (`automation-dashboard/`). Linked from the main header.

Each registered report appears as a row:

| Report | Schedule | Next Run | Last Auto Run | Last Manual Run | Delivery | |
|---|---|---|---|---|---|---|
| Student Page Goals | 1st of month | May 1, 2026 | Apr 1 ✓ | Apr 14 ✓ | 📁 Drive · 📧 Emailed | ▶ Run Now |

- ✓ / ✗ indicates success or failure
- Drive link opens the file in Google Drive
- "Run Now" triggers the script locally with live output streaming
- The dashboard reads `run_log.json` from the local repo path
- New reports are added by registering them in a config in `server.py`

---

## GitHub Actions Secrets

All secrets are added to the `mathnasium-page-goals` repo:

| Secret | Source |
|---|---|
| `RADIUS_USERNAME` | Copy from existing repos |
| `RADIUS_PASSWORD` | Copy from existing repos |
| `SMTP_USER` | Copy from `radius-morning-briefing` |
| `SMTP_PASSWORD` | Copy from `radius-morning-briefing` |
| `GOOGLE_DRIVE_CREDENTIALS` | New — Google service account JSON |
| `REPORT_RECIPIENTS` | Comma-separated email list |
| `GH_TOKEN` | For committing run_log.json back to repo |

---

## One-Time Setup Steps

1. **Google Drive folder:** Already created at the link above. Share with the service account email after creating it.
2. **Google service account:** Create in Google Cloud Console, enable Drive API, download credentials JSON, store as `GOOGLE_DRIVE_CREDENTIALS` secret.
3. **Gmail App Password:** Already set up in `radius-morning-briefing` — copy `SMTP_USER` and `SMTP_PASSWORD` secrets.
4. **Radius Digital Workout Plan URL:** Navigate to the report in Radius during implementation to capture the URL and UI selectors for Playwright.

---

## Error Handling

- If the Radius download fails, log the error and exit without sending email.
- If processing fails for a student, log a warning and continue with remaining students.
- If Drive upload fails, still send the email with the attachment.
- All errors written to `run_log.json` and printed to stdout (visible in dashboard output panel and GitHub Actions logs).

---

## Out of Scope

- Dashboard view of trends over time
- Alerts for students significantly above/below goals
- Historical comparison (month-over-month changes)
- Automated parent emails
