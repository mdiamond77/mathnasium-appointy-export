# Summer Package Status Report — Design Spec

**Date:** 2026-09-07
**Project:** `~/mathnasium-summer-status`

## Overview

An on-demand CLI tool that produces a single self-contained HTML report showing every
"Summer 2026 (Sessions Package)" enrollment at Teaneck and Englewood and what has become
of each student now that summer is over — in particular, who has converted to a
school-year membership.

The tool is run by hand. It is repullable for a quick update but has no scheduled
GitHub Actions workflow.

---

## Data Source

### Enrollment Report (only source — Approach 1)

- **URL:** `https://radius.mathnasium.com/Enrollment/EnrollmentReport`
- **One pull per center**, date window **4/1/2026 → today**, **all enrollment statuses**
  (not just Enrolled — we need completed, expired, cancelled, and hold rows).
- **Centers:** Teaneck (`2871`), Englewood (`2428`) — pulled and reported separately.
- **Fields needed from the Kendo grid (one row per enrollment):**
  - Student name
  - Center
  - Membership / package type
  - Enrollment start date
  - Enrollment status
  - Sessions per week (or equivalent program-frequency field)
- **Summer students** are the rows where membership type == `Summer 2026 (Sessions Package)`.
- **Conversion detection** uses the *other* enrollment rows for the same student in the
  same pull — no second report, no cross-report name matching.

### Discovery Gate (before any real build)

Run `discover_selectors.py` (headed browser) to confirm the Enrollment Report grid
exposes: membership type, one row per enrollment (not one row per student), enrollment
start date, enrollment status, and a sessions/week field. Also confirm whether membership
type is available as a server-side filter or must be filtered client-side after export.

**If the report does not expose membership type or is one-row-per-student, stop and
revisit** — the fallback is Approach 3 (Enrollment Report + Student Report).

---

## Status Classification

For each student who has at least one `Summer 2026 (Sessions Package)` row, sort all of
that student's enrollment rows by start date and assign exactly one bucket:

| Status | Rule |
|---|---|
| **Converted – active** | Has a non-summer enrollment with status Enrolled that started on or after the summer enrollment's start date |
| **Converted – on hold** | Same as above, but that non-summer enrollment's status is Hold |
| **Summer still active** | The summer package row itself is still status Enrolled and there is no school-year enrollment yet |
| **Completed – not returned** | Summer package row is finished / expired and there is no later non-summer enrollment |
| **Cancelled** | Summer package row status is Cancelled / Dropped |

### Edge cases

- **Multiple non-summer enrollments:** use the most recent one (latest start date) to
  determine the "Converted to" detail.
- **Converted and then went on hold:** classified as **Converted – on hold**
  (the hold status of the school-year enrollment wins over the summer state).
- **Unrecognized Radius status strings:** passed through verbatim into the "Current
  status" column and counted in an **Other / needs review** bucket rather than being
  forced into one of the five above. The report must never silently miscategorize.

### Conversion detail

For any student in a "Converted" bucket, capture from their most-recent non-summer
enrollment: membership type, start date, sessions per week.

---

## Output

A single self-contained HTML file: `output/summer_2026_status.html`. The tool opens it
in the default browser when the run finishes. No email, no attachments.

### Structure

1. **Header** — "Summer 2026 (Sessions Package) — Enrollment Status", generated timestamp.
2. **Overall summary** — total summer students and count + % per status bucket, with
   Teaneck and Englewood shown as side-by-side columns, plus an overall column.
   Includes **conversion rate** = (Converted – active + Converted – on hold) ÷ total
   summer students, per center and overall.
3. **Teaneck section:**
   - Status breakdown table (counts + conversion rate)
   - Student detail table, sorted by status then student name:

     | Student | Summer start | Summer status | Current status | Converted to | Conv. start | Sessions/wk |
     |---|---|---|---|---|---|---|

   Blank cells where a field does not apply (non-converters have no "Converted to").
4. **Englewood section** — identical structure.

### Styling

Plain, printable, light background, consistent with the other Mathnasium reports.
Inline CSS so the file is fully portable.

---

## Project Structure

**Location:** `~/mathnasium-summer-status` (new repo, sibling of the other automations).
**GitHub:** repo under the `mdiamond77` org. Org-level secrets already provide Radius
credentials for any future CI use; none needed for local runs.
**No `.github/workflows/`** — run on demand only.

| File | Responsibility |
|---|---|
| `config.py` | URLs, center config (Teaneck 2871, Englewood 2428), the summer membership string constant, column-name constants (some `TBD_` until discovery) |
| `discover_selectors.py` | One-time headed script to confirm Enrollment Report grid fields and filter behavior |
| `download.py` | Playwright headless: login via `~/.mathnasium_env`, pull Enrollment Report per center (4/1/2026 → today, all statuses), cache raw export to `input/` |
| `transform.py` | pandas: pure functions — identify summer students, classify each per the table above, build per-center and overall summaries |
| `report.py` | Render the HTML file from the transformed data; open it in the browser |
| `main.py` | CLI orchestrator: parse args, download → transform → report |
| `tests/test_transform.py` | Unit tests for every classification bucket and edge case |
| `requirements.txt` | playwright, pandas, openpyxl (if needed for reading the export), pytest |
| `.gitignore` | `input/`, `output/`, `__pycache__/`, `*.pyc`, `.env` |

---

## CLI Interface

```
python main.py                 # log in, pull both centers, transform, render, open report
python main.py --no-download    # re-render from cached input/ (fast iteration / quick refresh)
```

No date arguments — the summer membership string is fixed and the pull window is a
constant (4/1/2026 → today).

---

## Error Handling

Fail loud and early — never emit a misleading empty or partial report:

- Login failure → stop with a clear message.
- A center returns zero enrollment rows → stop (almost certainly a scrape problem).
- The summer membership string matches zero students across both centers → stop and
  report it (the label may have changed in Radius).
- Any student lands in **Other / needs review** → the report still renders, but the run
  prints a summary line noting how many need manual review.

---

## Testing

- `transform.py` is pure functions over pandas DataFrames. `tests/test_transform.py`
  builds small in-memory fixtures covering: each of the five buckets, the multiple-
  non-summer-enrollments case, converted-then-hold, and an unrecognized status string.
- `download.py` and `report.py` are verified manually against a real pull during the
  build.

---

## Build Order

1. `discover_selectors.py` — confirm the grid exposes what Approach 1 needs (gate).
2. `download.py` — get a real Enrollment Report export saved to `input/` for both centers.
3. Inspect the real export; fill in the `TBD_` column-name constants in `config.py`.
4. `transform.py` + tests.
5. `report.py`.
6. `main.py` wiring + full manual run.
