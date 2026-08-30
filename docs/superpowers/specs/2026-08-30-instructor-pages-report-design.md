# Instructor Pages Report — Design Spec

**Date:** 2026-08-30
**Status:** Draft — awaiting user review

---

## Overview

A report (ultimately a weekly automation) that downloads the Digital Workout Plan (DWP) report from Radius,
calculates how many pages the average student completes **per instructor**, and delivers
a formatted Excel workbook by email. The workbook shows each instructor's numbers for a
set of recent calendar-aligned windows (the "Snapshot") plus full week-by-week and
month-by-month history back to April.

## Phasing

**Phase 1 — this build.** Repo scaffold, `download.py`, `process.py`, the Excel workbook,
and a local `main.py` run. Deliverable: one workbook covering **April 1 → today**, handed
to Matt. No email automation, no cron, no dashboard row yet. Includes the live Radius
check to confirm the instructor column. Matt reviews the workbook and confirms the
numbers look right before Phase 2.

**Phase 2 — after Matt approves the Phase 1 output.** `deliver.py` (weekly email to Matt),
GitHub Actions cron (Mondays 13:00 UTC), and the `/reports` dashboard row + "Run Now"
button. This spec already describes all three; they just don't get built until the
output is trusted.

**Phase 3 — later, separate spec.** A compact per-instructor block for the "Instructors"
box of the daily center summary email.

---

## Repository

New repo: `mathnasium-instructor-pages`
Pattern mirrors `mathnasium-page-goals` exactly (`download` / `process` / `deliver` /
`main` + `run_log` + GitHub Actions + `/reports` dashboard row).

The DWP download is the **same Radius report** page-goals already pulls, so `download.py`
is a near-copy with a wider, parameterized date window.

---

## Architecture

### Modules

**`config.py`**
Constants: `RADIUS_LOGIN_URL`, `RADIUS_DWP_URL`
(`https://radius.mathnasium.com/DigitalWorkoutPlan/Report`), center IDs
(`2428` Englewood, `2871` Teaneck), recipients, paths, `SEASON_START` (`2026-04-01`).

**`download.py`**
Playwright (headless Chromium). Copy of the page-goals downloader with one change: the
date window is `--start` / `--end` (defaults: start = `SEASON_START`, end = today) instead
of a fixed 90-day trailing window. Selects both centers via the Kendo MultiSelect
(`#AllCenterListMultiSelect`), sets `#dwpFromDate` / `#dwpToDate` Kendo date pickers,
clicks `#btnsearch`, waits for `#dwpExcelBtn`, downloads to
`input/Digital_Workout_Plan_<start>_<end>.xlsx`.

**`process.py`**
Reads the DWP Excel with pandas. Applies the filtering and bucketing logic below.
Writes `output/Instructor Pages <run-date>.xlsx` with openpyxl (4 analysis tabs + raw
Data tab). Pure functions for all calculation, so the logic is unit-tested without a
real download.

**`deliver.py`**
Emails the output workbook as an attachment via Gmail SMTP (SSL 465), reusing
`SMTP_USER` / `SMTP_PASSWORD`. Success email → Matt only. Failure email → Matt only.
No Google Drive upload in this version.

**`main.py`**
Orchestrates download → process → deliver. Flags:
- `--start YYYY-MM-DD` / `--end YYYY-MM-DD` — override the date window (default April 1 → today)
- `--trigger auto|manual` — recorded in the run log
- `--skip-download` — reuse the latest file already in `input/` (for iterating on formatting)

Writes a record to `run_log.json` on success or failure; GitHub Actions commits it back.

**`run_log.py`**
Identical to the page-goals module (`read_log`, `append_run`, `get_last_run`).

---

## Source Data

DWP export, one row per student session. Columns the report needs:

| Field | Use |
|---|---|
| `Date` | Session date → week / month bucketing |
| `Student Name` | Distinct-student counts, per-student averages |
| `Pages Completed` | The metric |
| `Center` | Split Teaneck vs Englewood |
| *Instructor column* | **Exact header TBD** — confirmed during implementation by opening a real download (candidates: "Instructor", "Coach", "Staff"). |

### Center assignment

Same rule as page-goals: `Center` value containing "Englewood" → Englewood;
else containing "Teaneck" → Teaneck; else the row is dropped. Englewood is checked
first so a combined value like "Englewood, Teaneck Virtual" routes to Englewood.

### Instructor edge cases

- **Blank / missing instructor** → attributed to a row labeled `(Unassigned)` rather than
  silently dropped, so page totals still reconcile.
- **Multiple instructors in one field** (separator-delimited) → behavior confirmed on the
  live check. Default plan: split on the separator and attribute the session to **each**
  named instructor (the session is counted once under each). If co-taught sessions turn
  out not to exist in the data, this code path is dead and harmless.

---

## Calculation Logic

### Session classification

For each row, `Pages Completed` is coerced to a number:

| Class | Rule |
|---|---|
| **Productive session** | numeric and `> 0` |
| **Zero / blank session** | null, non-numeric, or `== 0` |
| **Low session (≤3)** | productive **and** `pages <= 3` (a subset of productive) |

Averages use **productive sessions only**. Zero/blank sessions are excluded from every
average but their count is reported (Detail tab, and as a footnote on the Snapshot).

### Per-instructor, per-period metrics

Given an instructor, a center, and a date range:

1. **Pages/Session** — `sum(pages) / count` over productive sessions in range, 2 dp.
2. **Sessions** — count of productive sessions.
3. **≤3 pg** — count of low sessions.
4. **≤3 pg %** — `low sessions / productive sessions`, shown as a percent, 0 dp.
5. **Zero/blank** — count of zero/blank sessions (Detail tab only + Snapshot footnote).
6. **Students** — count of distinct students with ≥1 productive session (Detail tab).

An instructor with 0 productive sessions in a period shows blank cells for that period
(not `0.00`).

### Center summary row

Each center block on the Snapshot, Monthly, and Weekly tabs opens with a **pooled
center-wide row** (label `▸ TEANECK — all instructors` / `▸ ENGLEWOOD — all
instructors`), computed across every session at that center in the period (not an
average of instructor averages):

- Pages/Session — `sum(all productive pages) / count(all productive sessions)`
- Sessions — sum of productive sessions (a total)
- ≤3 pg — sum of low sessions (a total)
- ≤3 pg % — `sum(low) / sum(productive)`

Multi-instructor sessions are counted **once** in the center pooled figures (dedupe on
the underlying session), even though they credit each named instructor in that
instructor's own row.

> **Dropped in the Phase 1.1 revision:** the "Student Avg" (average of student averages)
> metric — removed from every tab. **Not adopted:** a "% of page goal" metric — page
> goals are set by the instructors themselves from each student's recent pace, so
> attainment against them is circular. A future revision may instead compare instructors
> on *the same students* (students who worked with more than one instructor).

### Periods

All periods are **calendar-aligned**. Weeks run **Sunday → Saturday**.

**History buckets (Monthly and Weekly tabs):**
- Months: every calendar month from April through the current month (current month partial).
- Weeks: every Sun–Sat week that overlaps `SEASON_START` → today. First bucket is the
  week containing April 1 (Sun 2026-03-29 → Sat 2026-04-04); last bucket is the week
  containing today. Weeks are labeled `Wk of M/D` using the Sunday date.

**Snapshot windows (as of the run date):**

| Column | Definition (example: run on 2026-08-30) |
|---|---|
| This Month So Far | 1st of current month → run date (Aug 1 – Aug 30) |
| Last Week | most recent **complete** Sun–Sat week (not the current partial week) |
| Last Month | most recent **complete** calendar month (July) |
| Last 2 Months | the 2 most recent complete calendar months combined (June + July) |
| Last 3 Months | the 3 most recent complete calendar months combined (May – July) |

---

## Output Format

**Filename:** `Instructor Pages YYYY-MM-DD.xlsx` (run date)

**One workbook. Centers are never comingled** — every analysis tab is split into a
Teaneck block on top and an Englewood block below, separated by a full-width divider row
(`━━━ TEANECK ━━━` / `━━━ ENGLEWOOD ━━━`). Within each block, instructors are listed
alphabetically, with `(Unassigned)` last. An instructor who worked at both centers in a
period appears in both blocks, counting only that center's sessions.

### Tab 1 — Snapshot

Center summary row, then one row per instructor. Five period column-groups, each with
four sub-columns:

| Sub-column | Meaning |
|---|---|
| Pages/Sess | productive-session average |
| Sessions | productive session count |
| ≤3 pg | low session count |
| ≤3 pg % | low sessions ÷ productive sessions |

**Group headers are the actual date ranges** (`Aug 23 – Aug 29`, `Jul 1 – Jul 31`,
`Jun 1 – Jul 31`, …) with the plain-English label (`Last Week`, `Last Month`, …) as a
smaller subtitle line beneath.

Footnote row under each center block: total zero/blank sessions excluded, per period.

### Tab 2 — Monthly

Four side-by-side sections, each a full grid of **center summary row + one row per
instructor** × **one column per calendar month, April → current**:

1. **AVG PAGES** — Pages/Session, blank if no sessions
2. **# SESSIONS** — productive session count
3. **# ≤3 PAGES** — low session count
4. **% ≤3 PAGES** — low ÷ productive, percent

Sections are separated by a blank spacer column and a section-title band. No more
`avg (n)` parentheses.

### Tab 3 — Weekly

Unchanged from Phase 1 except for the new center summary row: center row + one row per
instructor × one column per Sun–Sat week since the first week of April. Cell: `avg (n)`
where `avg` = Pages/Session and `n` = productive session count. Blank if no sessions.
Instructor column frozen.

### Tab 4 — Detail (long format)

The analytical source — one row per (center, instructor, period). Both period types
stacked. Columns:

`Center | Instructor | Period Type | Period Label | Period Start | Period End |
Pages/Session | Productive Sessions | Zero/Blank Sessions | Low Sessions (≤3) |
Low Session % | Distinct Students`

Matt can pivot / filter this himself.

### Tab 5 — Data

The raw DWP download, all columns, unmodified — for spot-checking any number.

### Formatting

- Header rows: blue background (`#4472C4`), white bold, centered.
- Center divider rows: dark fill, white bold, merged across the used width.
- Snapshot period groups: alternating light fills to separate the five groups visually.
- Freeze the header row and the instructor-name column on every tab.
- Column widths: instructor 26, numeric 10–12.

---

## Delivery (Phase 2 — not built until the Phase 1 workbook is approved)

**Email (success and failure):** `matt.diamond@mathnasium.com` only.
This is a management metric; center directors are **not** on it in this version.

**Schedule:** GitHub Actions cron every **Monday at 13:00 UTC** (9:00 AM ET) — captures the
just-completed Sun–Sat week. `workflow_dispatch` enabled for manual GitHub runs.

**Dashboard (Phase 2):** register on the existing `/reports` page (built in the page-goals project).
Add a hidden `instructor-pages` script entry and a `REPORTS` row pointing at this repo's
`run_log.json`, with a "Run Now" button that streams live output. Same mechanism as
page-goals — no new dashboard code, just config.

---

## Run Log

`run_log.json` in the repo root, same shape as page-goals:

```json
[
  {
    "timestamp": "2026-08-31T13:04:11Z",
    "trigger": "auto",
    "window": "2026-04-01..2026-08-30",
    "status": "success",
    "output_file": "Instructor Pages 2026-08-30.xlsx",
    "error": null
  }
]
```

`trigger` is `"auto"` (cron) or `"manual"` (dashboard / CLI). On failure `status` is
`"error"` and `error` holds the exception message. GitHub Actions commits the updated
file back after every run.

---

## GitHub Actions Secrets

Added to the `mathnasium-instructor-pages` repo (copy values from existing repos):

| Secret | Source |
|---|---|
| `RADIUS_USERNAME` | existing repos |
| `RADIUS_PASSWORD` | existing repos |
| `SMTP_USER` | `radius-morning-briefing` |
| `SMTP_PASSWORD` | `radius-morning-briefing` |
| `GH_TOKEN` | committing `run_log.json` back |

---

## Testing

`tests/test_process.py` — pure-function coverage:

- Week bucketing is Sunday-start (a Saturday and the next Sunday fall in different weeks).
- Month bucketing; partial current month handled.
- Snapshot windows: "Last Week" excludes the current partial week; "Last Month" is the
  last *complete* month; the 2- and 3-month windows are contiguous complete months.
- Zero, blank, and non-numeric `Pages Completed` excluded from averages but counted.
- Low-session (≤3) count includes 3, excludes 0 and 4.
- Pages/Session vs Student Avg diverge correctly when one student has many sessions.
- Center split; "Englewood, Teaneck Virtual" → Englewood.
- Blank instructor → `(Unassigned)` row.
- Multi-instructor field → session counted under each named instructor.
- Instructor with 0 productive sessions in a period → blank cells, not `0.00`.

`tests/test_run_log.py` — copied from page-goals.

---

## Error Handling

- Radius download failure → log, email Matt the error, exit non-zero, record in run log.
- A row that can't be parsed (bad date, etc.) → warn and skip that row, continue.
- Overall failure → failure email to Matt only.
- All output printed to stdout (visible in dashboard output panel and Actions logs).

---

## Open Items (resolve during implementation)

1. **Exact instructor column header** in the DWP export — confirm by opening a real
   download before writing `process.py`.
2. **Co-taught sessions** — confirm whether the instructor field ever holds more than one
   name and what separator it uses; adjust the split logic or delete it.
3. **DWP row granularity** — confirm one row per student per session (not per subject or
   per page range). If finer, decide how to collapse to a session.

---

## Out of Scope

- The daily center summary email "Instructors" block (Phase 2).
- Google Drive archiving of the workbook.
- Center directors as recipients.
- Alerting / thresholds (e.g. flag an instructor trending down).
- Comparisons against student page *goals* from the page-goals automation.
- Any instructor metric other than pages (attendance, punctuality, retention).
