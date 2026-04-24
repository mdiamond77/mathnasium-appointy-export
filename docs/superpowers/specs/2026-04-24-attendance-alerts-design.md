# Attendance Alerts — Design Spec

**Date:** 2026-04-24
**Status:** Approved

---

## Goal

Every Monday, email each center director a list of students who attended fewer than 75% of their allowed sessions over the past 2 weeks and/or 4 weeks. Helps directors identify students who are slipping and need outreach before they churn.

---

## Repository

New repo: `mathnasium-attendance-alerts`
Pattern mirrors `mathnasium-hold-reminders`.

---

## Schedule

GitHub Actions cron `0 15 * * 1` — every Monday at 11am ET (3pm UTC).
Also supports `workflow_dispatch` for manual testing at any time.

---

## Data Sources

### 1. Enrolled Report
**URL:** `https://radius.mathnasium.com/Enrollment/EnrollmentReport`

**Columns used:**
| Column | Purpose |
|--------|---------|
| `Account Id` | Join key to Workout Plan |
| `Student First Name` + `Student Last Name` | Display name |
| `Center` | Email routing (always clean "Englewood" / "Teaneck") |
| `Monthly Amount` | Determines allowed sessions/week |
| `Status` | Filter to `Enrolled` only |

### 2. Digital Workout Plan Report
**URL:** `https://radius.mathnasium.com/DigitalWorkoutPlan/Report`

Date range set via Playwright before export: past 28 days from today.

**Columns used:**
| Column | Purpose |
|--------|---------|
| `Account Id` | Join key to Enrolled Report |
| `Date` | Assign session to 2-week or 4-week window |
| `Session Start` | Duration calculation |
| `Session End` | Duration calculation |

---

## Processing Logic

### Session counting
- Each row in the Workout Plan = 1 session.
- If session duration ≥ 120 minutes → count as 2 sessions.
- Negative or implausible durations (data entry errors) → treat as 1 session.

### Allowed sessions per student
| Monthly Amount | Sessions/week | Allowed 2-week | Allowed 4-week |
|---------------|--------------|----------------|----------------|
| ≥ $350 | 2 | 4 | 8 |
| < $350 | 1 | 2 | 4 |

Students with missing `Monthly Amount` → treated as 1 session/week.

### Time windows
Both windows are rolling from the Monday the script runs:
- **2-week:** sessions in the past 14 days
- **4-week:** sessions in the past 28 days

### Attendance threshold
Flag students where `actual sessions / allowed sessions < 0.75`.

### Student exclusions
Skip students whose names match (case-insensitive):
- Contains "test"
- Contains "appoin"
- Exact match "x x"
- Contains "diamond" (owner's family)

### Categorization
After flagging:
- **Both windows:** flagged in 2-week AND 4-week → highest priority
- **2-week only:** flagged in 2-week but not 4-week → recent dip
- **4-week only:** flagged in 4-week but not 2-week → sustained issue, recently improved

Students with 0 sessions in 4 weeks are included — they may be informally inactive or forgotten to be placed on hold, which is exactly what this report is meant to surface.

---

## Architecture

```
mathnasium-attendance-alerts/
├── main.py           # Orchestrator: download → process → deliver → log
├── download.py       # Playwright: login, download Enrolled Report + Workout Plan (28-day window)
├── process.py        # pandas: count sessions, compute attendance %, flag students
├── deliver.py        # HTML email builder + Gmail SMTP (SSL port 465)
├── config.py         # Recipients, URLs, thresholds, center config
├── run_log.py        # Append run result to run_log.json
├── requirements.txt
└── .github/
    └── workflows/
        └── attendance_alerts.yml
```

### Module responsibilities

**`download.py`**
- Playwright (headless Chromium) logs into `https://radius.mathnasium.com`
- Downloads Enrolled Report as Excel → `input/enrolled.xlsx`
- Navigates to Digital Workout Plan Report, sets date range to past 28 days via UI, exports Excel → `input/workout_plan.xlsx`

**`process.py`**
- Loads both Excel files with pandas
- Filters enrolled report to `Status == Enrolled`, applies name exclusions
- Counts sessions per student per window (with 2-hour double-count rule)
- Computes attendance % for each window, flags at < 75%
- Returns three categorized lists per center: both-flagged, 2-week-only, 4-week-only

**`deliver.py`**
- Builds one HTML email per center
- Three sections in order: Both Windows (highest priority), 2-Week Only, 4-Week Only
- Sends via Gmail SMTP (SSL, port 465)
- Skips centers with no flagged students in any section

**`main.py`**
- Orchestrates download → process → deliver
- Accepts `--trigger manual` flag (bypasses any future date gates)
- Writes run record to `run_log.json` on completion (success or failure)

**`config.py`**
- Recipients, Radius URLs, center names, attendance threshold (0.75), session length threshold (120 min), monthly amount cutoff ($350)

---

## Email Format

**Recipients:** Each center director receives their own email (their center's students only). Matt is CC'd on both.

**Subject:** `Attendance Alerts — [Center Name] — Week of [Month D, YYYY]`
Example: `Attendance Alerts — Teaneck — Week of April 28, 2026`

**Body:**

> Happy Monday!
>
> The following students attended fewer than 75% of their allowed sessions. Please reach out as needed.
>
> **⚠ Flagged in Both Windows** *(highest priority)*
> | Student | Sessions/Week | 2-Week (actual/allowed) | 4-Week (actual/allowed) |
> |---------|--------------|------------------------|------------------------|
> | Adina Luks | 2 | 2/4 (50%) | 5/8 (63%) |
>
> **Last 2 Weeks Only**
> | Student | Sessions/Week | 2-Week (actual/allowed) |
> |---------|--------------|------------------------|
>
> **Last 4 Weeks Only**
> | Student | Sessions/Week | 4-Week (actual/allowed) |
> |---------|--------------|------------------------|
>
> *This email sends automatically every Monday.*

**Empty state:** If a center has no flagged students across all three sections, no email is sent for that center.

---

## GitHub Secrets Required

| Secret | Purpose |
|--------|---------|
| `RADIUS_USERNAME` | Radius login |
| `RADIUS_PASSWORD` | Radius login |
| `SMTP_USER` | Gmail sender |
| `SMTP_PASSWORD` | Gmail app password |

All secrets already exist in the GitHub org and will be inherited by this repo.
