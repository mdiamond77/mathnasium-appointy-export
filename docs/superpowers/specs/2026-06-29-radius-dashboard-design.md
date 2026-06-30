# Radius Dashboard — Design Spec

**Date:** 2026-06-29
**Status:** Approved

## Overview

A recurring Google Sheets dashboard pulling data from two Radius reports and publishing two views per center: a lead funnel report and an enrollment/MRR report. The sheet is shared with center directors via a single permanent link and fully rewritten on each run.

---

## Delivery

- **Format:** Google Sheet, one workbook, 6 tabs
- **Sharing:** One permanent link shared with center directors
- **Update cadence:** Fully rewritten on each run — no incremental append logic

**Tab names:**
1. Funnel - Combined
2. Funnel - Teaneck
3. Funnel - Englewood
4. Enrollments - Combined
5. Enrollments - Teaneck
6. Enrollments - Englewood

---

## Schedule

- GitHub Actions cron: 1st and 15th of each month (9am ET)
- Manual trigger: `workflow_dispatch` for on-demand runs

---

## Data Sources

Both exports are driven by Playwright (login → navigate → set filters → export CSV), consistent with existing automations.

### 1. Leads Tracking Report
- **Scope:** Per center (Teaneck, Englewood exported separately)
- **Date range:** Covers 2 completed calendar years + current year to date
- **Used for:** Lead counts, assessed flags, enrolled flags

### 2. Enrollment Report
- **Scope:** Per center (Teaneck, Englewood exported separately)
- **Date range:** January 1, 2018 to today (full history needed for accurate assessed backfill)
- **Used for:** Enrollment type, MRR per enrollment, enrollment date, student name

---

## Data Transform Logic

### Lead filtering
- Exclude leads whose source is an event type, **unless** that lead enrolled (appears in the enrollment report as a non-private enrollment)
- Event leads who never converted are dropped entirely from all funnel counts

### Assessed backfill
- A student is counted as assessed if:
  - The leads report marks them as assessed, **or**
  - They appear in the enrollment report with a non-private enrollment type
- This means enrolled students are always counted as assessed, even if the leads report does not flag them

### Private Tutoring
- Enrollments where type = Private Tutoring are excluded from all main enrollment counts and MRR
- They are tracked separately in muted columns at the right of the enrollment summary

### Date grouping
- All tables display newest first (reverse chronological)
- Rows are grouped by year with a year-header row separating sections
- Display range: current year to date + 2 completed calendar years

---

## Funnel Tabs (Combined, Teaneck, Englewood)

One row per month. Columns:

| Column | Description |
|--------|-------------|
| Month | e.g. Jun 2026 |
| Leads | Count of leads (event leads excluded unless enrolled) |
| Assessed | Count assessed (leads report flag OR enrolled non-private) |
| Enrolled | Count enrolled (non-private) |
| % leads assessed | Assessed / Leads |
| % assessments converted | Enrolled / Assessed |
| % leads converted | Enrolled / Leads |

Combined tab aggregates both centers. Per-center tabs filter to that center only.

---

## Enrollment Tabs (Combined, Teaneck, Englewood)

Each tab has two sections:

### Section 1 — Monthly summary (all months in display range)

One row per month, newest first. Columns:

| Column | Description |
|--------|-------------|
| Month | e.g. Jun 2026 |
| Total enrollments | Non-private enrollments |
| [Type A], [Type B], … | One column per non-private enrollment type (dynamic — reflects actual Radius types) |
| MRR | Sum of MRR for non-private enrollments |
| Private Tutoring | Count of PT enrollments (muted) |
| PT MRR | Sum of MRR for PT enrollments (muted) |

### Section 2 — Individual enrollment detail (recent months only)

One row per enrollment, grouped by month. Covers: past 3 completed months + current month (4 months total).

| Column | Description |
|--------|-------------|
| Month | e.g. Jun 2026 |
| Student name | From enrollment report |
| Center | Teaneck or Englewood |
| Enrollment type | From enrollment report |
| MRR | Per-enrollment amount |
| Enrollment date | From enrollment report |

Private Tutoring enrollments are excluded from this detail section.

Combined tab shows all centers. Per-center tabs filter to that center only.

---

## Code Structure

New GitHub repo (e.g. `radius-dashboard`), following the same pattern as existing automations:

```
main.py              — orchestrator; --date override for testing
radius_export.py     — Playwright: login, export Leads Tracking + Enrollment reports per center
transform.py         — pandas: filter, join, aggregate, calculate rates
sheets.py            — Google Sheets API: authenticate via service account, write all 6 tabs
config.py            — constants: date windows, sheet ID, center names, Radius selectors
```

### GitHub Actions
- Two scheduled triggers: `0 14 1 * *` and `0 14 15 * *` (9am ET)
- `workflow_dispatch` for manual runs
- Secrets: `RADIUS_USERNAME`, `RADIUS_PASSWORD`, `GOOGLE_SERVICE_ACCOUNT_JSON`, `SPREADSHEET_ID`

### Google Sheets authentication
- Service account (JSON key) — no OAuth flow required for automated runs
- Sheet is shared with the service account email; directors access via the normal shared link

---

## Selector Discovery

Exact Radius report URLs and export selectors for the Leads Tracking and Enrollment reports are unknown at design time. A `discover_selectors.py` script (consistent with the existing automation pattern) will be used to identify and validate selectors before full implementation.

---

## Out of Scope

- Charts or visualizations within the sheet (directors can add these manually)
- Email notification when the sheet is updated
- Historical data beyond 2 completed years + YTD in the display
