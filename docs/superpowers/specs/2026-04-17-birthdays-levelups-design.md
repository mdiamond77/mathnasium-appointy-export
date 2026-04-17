# Birthdays & Level Ups Automation — Design Spec

**Date:** 2026-04-17  
**Status:** Approved

---

## Overview

A monthly automation that pulls two reports from Radius — the Enrollment Report (for level-ups) and the Birthdays Report (for birthdays) — and emails a formatted, printable HTML summary to each center's director on the 1st of every month. No spreadsheet attachment, no Drive upload.

---

## Repository

New repo: `mathnasium-birthdays-levelups`  
Pattern mirrors `mathnasium-page-goals`.

---

## Architecture

### Modules

**`download.py`**  
Uses Playwright (headless Chromium) to log into Radius and download two reports, each covering both centers (Englewood + Teaneck selected in the center MultiSelect):
1. **Enrollment Report** — used for level-up calculation via the "student length of stay" column.
2. **Birthdays Report** — filtered to enrolled students only.

Files saved to `input/`. The Center column in each report is used downstream to split data per center.

**`process.py`**  
Reads downloaded files with pandas. Produces two data structures per center:
- **Level-up data** — current month and last month (see Level-Up Logic below).
- **Birthday data** — name, birthday date, age.

**`deliver.py`**  
Sends one HTML email per center via Gmail SMTP (SSL, port 465):
- Englewood email → Englewood CD
- Teaneck email → Teaneck CD
- Matt copied on both

No attachment. No Drive upload.

**`main.py`**  
Orchestrates download → process → deliver. Accepts a `--month YYYY-MM` flag (defaults to current month). Writes a record to `run_log.json` on completion (success or failure) and commits it back to the repo via Actions.

**`config.py`**  
Recipients, Radius URLs, center IDs, file paths.

---

## Data Sources

| Report | Purpose | Radius filter |
|---|---|---|
| Enrollment Report | Student length of stay → level-ups | Default (enrolled students) |
| Birthdays Report | Student name, birthday, age | Filter: Enrolled |

---

## Level-Up Logic

Levels are determined by months enrolled:

| Months enrolled | Level |
|---|---|
| 0–11 | 1 |
| 12–23 | 2 |
| 24–35 | 3 |
| 36–47 | 4 |
| 48+ | 5 (cap — stays at 5 forever) |

The "student length of stay" column contains a decimal value (e.g., `12.1` months).

**Current month level-ups:** `floor(length_of_stay)` in `{12, 24, 36, 48}`  
**Last month level-ups:** `floor(length_of_stay)` in `{13, 25, 37, 49}`

A student whose `floor` value matches the current month set just crossed a level boundary this month. Last month's set confirms the prior month's crossings.

> **Build note:** Verify the exact format of the length of stay column from a real Radius export before finalizing the parsing step. It may be a float, an integer, or a formatted string — adjust accordingly.

---

## Email Format

### Subject
`Mathnasium [Center] — Birthdays & Level Ups: [Month Year]`

### Body (HTML, printable)

---

Hi [Center] Center Directors,

Please find this month's Birthdays & Level Ups below. As a reminder:
- Please have an instructor **update the student binders** for any level ups.
- Please **add this month's birthdays to the whiteboard** so we can celebrate with our students!

---

**🎂 Birthdays — [Month Year]**

| Name | Birthday | Age |
|---|---|---|
| Alex Johnson | Apr 3 | 9 |

---

**⭐ Level Ups This Month**

| Name | Old Level | New Level |
|---|---|---|
| Jordan Kim | Level 1 | Level 2 |

---

**✅ Last Month's Level Ups — Please Confirm Binders Were Updated**

The following students leveled up last month. Please confirm their binders have been updated.

| Name | Old Level | New Level |
|---|---|---|
| Casey Lee | Level 2 | Level 3 |

---

*This email was generated automatically. Questions? Contact matt.diamond@mathnasium.com.*

---

If any section has no entries, it displays "None this month." rather than being omitted.

---

## Delivery & Scheduling

- **Transport:** Gmail SMTP, SSL, port 465
- **Credentials:** `SMTP_USER` / `SMTP_PASSWORD` (same secrets as page goals)
- **Recipients:**
  - Englewood: englewood@mathnasium.com (+ matt.diamond@mathnasium.com)
  - Teaneck: teaneck@mathnasium.com (+ matt.diamond@mathnasium.com)
- **Schedule:** GitHub Actions cron `0 15 1 * *` (1st of month, 11am ET)
- **Manual trigger:** `workflow_dispatch` for testing
- **Re-run:** `--month YYYY-MM` flag on `main.py`

---

## Run Logging

Same pattern as page goals: `run_log.json` in repo root, committed back by Actions after each run. Each entry records: timestamp, month processed, success/failure, any error message.

---

## GitHub Secrets Required

| Secret | Purpose |
|---|---|
| `RADIUS_USERNAME` | Radius login |
| `RADIUS_PASSWORD` | Radius login |
| `SMTP_USER` | Gmail sender address |
| `SMTP_PASSWORD` | Gmail App Password |

---

## Error Handling

- If either Radius download fails, the run aborts and logs the error — no partial email sent.
- If SMTP fails, the error is logged with full traceback.
- All credentials read inside functions via `os.environ.get()` (never at module level).

---

## Out of Scope

- Google Drive upload (email only)
- Spreadsheet attachment
- Dashboard integration (can be added later)
