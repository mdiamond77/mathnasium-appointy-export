# New Enrollments Report — Design Spec
**Date:** 2026-06-28
**Project:** `~/mathnasium-new-enrollments`

## Overview

A CLI tool that produces a historical new-enrollment report covering January 2023 through the current month. It identifies students who were genuinely new (or returned after a 3+ month gap), pulls the amount charged from Payment Reconciliation, and outputs two Excel files — one per center — each with a detail tab and a summary tab.

---

## Data Sources

### 1. Enrollment Report
- **URL:** `https://radius.mathnasium.com/Enrollment/EnrollmentReport`
- **Single pull:** date range 9/1/2022 → today (covers all students needed, including 3-month lookback window for Jan 2023)
- **Fields needed from Kendo grid:** student name, center, enrollment start date, program type (e.g., 2x/week, 3x/week)
- **Discovery gate:** before building, run a selector discovery pass to confirm these fields are exposed in the grid

### 2. Payment Reconciliation
- **URL:** `https://radius.mathnasium.com/Payment`
- **Pulled by year:** 2022 (Sept–Dec), 2023, 2024, 2025, 2026 (Jan–present) — 5 exports total
- **Fields needed:** account/student name, center, amount charged
- **Caching:** each year saved to `input/PaymentRecon_YYYY.xlsx`; re-runs skip already-cached files

---

## New Enrollment Detection Logic

Using the single enrollment export (all data from 9/1/2022):

1. For each student + center, sort all enrollment records by start date
2. A record counts as a **new enrollment** for month M if:
   - The enrollment start date falls within month M (Jan 2023 or later), AND
   - The student has no enrollment record that ended within the 3 months immediately preceding that start date
3. Students who had a prior enrollment ending more than 3 months before are treated as new (returner but outside the window)
4. The first 3 months of the report window (Jan–Mar 2023) have full lookback data since enrollment data starts 9/1/2022

---

## Join

Payment Recon amounts are joined to enrollment records by **student name + center**. Name normalization (strip/lowercase) applied before join to reduce mismatches. Unmatched enrollments (no payment record found) are included with amount shown as blank/unknown.

---

## Output

Two Excel files written to `output/`:
- `NewEnrollments_Teaneck.xlsx`
- `NewEnrollments_Englewood.xlsx`

### Tab 1: Enrollments (detail)
| Enrollment Month | Center | Student Name | Program Type | Amount Charged |
|---|---|---|---|---|
| January 2023 | Teaneck | John Smith | 2x/week | $299 |

### Tab 2: Summary
| Enrollment Month | Center | New Enrollments | Total Amount Charged |
|---|---|---|---|
| January 2023 | Teaneck | 4 | $1,196 |

Enrollment Month formatted as `Month YYYY` (e.g., `January 2023`).

---

## CLI Interface

```
python main.py
```

Defaults: enrollment data from 9/1/2022, report output from 2023-01 through current month.

Optional overrides:
- `--from YYYY-MM` — start of report window (default: 2023-01)
- `--to YYYY-MM` — end of report window (default: current month)

---

## Project Structure

```
~/mathnasium-new-enrollments/
  main.py              # orchestrator: parse args, call download/transform/output
  download.py          # Playwright: login, pull enrollment report + payment recon by year
  transform.py         # pandas: detect new enrollments, join with payment data
  output.py            # write two-tab Excel files per center
  discover_selectors.py # one-time script to inspect enrollment grid fields
  config.py            # URLs, center config, column name constants
  requirements.txt
  input/               # cached raw downloads (gitignored)
  output/              # generated Excel files (gitignored)
```

---

## Key Risks & Mitigations

| Risk | Mitigation |
|---|---|
| Enrollment Report grid doesn't expose start date or program type | Run discover_selectors.py first; adjust approach if fields missing |
| Payment Recon year pull times out or errors | Cache each year independently; retry individual years |
| Name mismatches between Enrollment and Payment Recon | Normalize names before join; flag unmatched rows visibly in output |
| Large enrollment dataset slow to process | All in pandas in-memory; 3.5 years of data should be manageable |
