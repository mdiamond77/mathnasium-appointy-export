# Summer Package Status Report — Design Spec

**Date:** 2026-09-07
**Project:** `~/mathnasium-summer-status`
**Status:** revised 2026-09-07 after live Radius discovery (see "Discovery findings" below)

## Overview

An on-demand CLI tool that produces a single self-contained HTML report showing every
2026 summer-package enrollment at Teaneck and Englewood and what has become of each
student now that summer is over — in particular, who has converted to an ongoing
school-year membership.

Run by hand. Repullable any time for a fresh snapshot. No scheduled GitHub Actions
workflow.

---

## Data Source

### Enrollment Report (only source — Approach 1)

- **URL:** `https://radius.mathnasium.com/Enrollment/EnrollmentReport`
- **One pull per center**, date window **4/1/2026 → 12/31/2026**. Leave the enrollment
  status filter at its default so **all** statuses come back.
- **Centers:** Teaneck (`2871`), Englewood (`2428`) — pulled and reported separately.
- The report returns **one row per enrollment** (confirmed in discovery). A student with
  three enrollments overlapping the window appears in three rows.
- The date filter matches enrollments that **overlap** the window, not just those that
  start in it — so a summer enrollment that began in June still appears.
- **Conversion detection** uses the *other* enrollment rows for the same student in the
  same pull. No second report, no cross-report matching.

### Discovery findings (2026-09-07)

Ran headless against live Radius. Gate **passed**. Details:

- **Export columns (28):** `Lead Id, Account Id, Student First Name, Student Last Name,
  Grade, Grade Range, Account Name, Center, Status, Membership Type,
  Primary Enrollment Start, Primary Enrollment End, Recurring, Enrollment Contract Length,
  Enrollment Length of Stay, Student Length of Stay, Total Sessions, Remaining,
  Session Length, Hold Count, Total Hold Length, Delivery, Monthly Amount,
  Expected Monthly Amount, Virtual Center, Guardians, Guardian Emails,
  Guardian Phone Numbers`.
  - Despite the name, **`Primary Enrollment Start` / `Primary Enrollment End` are
    per-enrollment** (a student's rows carry different values).
- **No student ID in the export.** `Account Id` is per family (siblings share it).
  Student identity key = **`Student First Name` + `Student Last Name` + `Center`**,
  normalized (trim, collapse whitespace, casefold). Collision risk (two unrelated
  same-named students at one center) is low; such cases are surfaced in a warning.
- **Enrollment status has only four possible values:** `Pre-Enrolled`, `Enrolled`,
  `On Hold`, `Inactive`. There is **no** "Completed" / "Expired" / "Cancelled" — a
  finished package and an early cancellation both read as `Inactive`.
- **`Status` == `Enrolled` does not mean currently attending.** Several summer
  enrollments still read `Enrolled` weeks after their end date. "Still active" is
  decided by **`Primary Enrollment End` >= today**, not by the status label.
- **Summer membership type differs by center** (value as it appears in the export's
  `Membership Type` column, after stripping an optional leading `"* "`):
  - Teaneck: `Summer 2026 (Sessions Package)` (24 enrollments on 2026-09-07)
  - Englewood: `2026 Summer Sessions Package (Sessions Package)` (17 enrollments)
- **`Membership Type` sometimes has a leading `"* "`** on otherwise-identical values.
  Strip it before matching.
- **No "sessions per week" column.** Frequency is embedded in the membership-type name
  for recurring plans (e.g. `2x/week - 2024 Flexible (Flexible)`). The report shows the
  membership-type name and `Monthly Amount` as the conversion detail instead.
- **School Partnership** enrollments appear as membership types containing the substring
  `School Partnership` (e.g. `* Ridgefield Park (School Partnership) (Sessions Package)`).

### Selectors confirmed

| Purpose | Selector | Notes |
|---|---|---|
| Login | `#UserName`, `#Password`, `#login` | |
| Center multiselect | `#AllCenterListMultiSelect` | Kendo MultiSelect; set via `$(...).data('kendoMultiSelect').value(['2871']); .trigger('change')` |
| From date | `#StartDate` | Kendo DatePicker; set via `.data('kendoDatePicker').value(new Date(2026,3,1))` |
| To date | `#EndDate` | Kendo DatePicker; `new Date(2026,11,31)` |
| Enrollment status filter | `#EnrollmentStatusDropDown` | leave untouched → all statuses |
| Search | `#btnsearch` | |
| Export | `#btnExport` | fires a file download |
| Grid | `#gridEnrollmentReport` | |

---

## Status Classification

For each student who has at least one summer-package row (using that center's summer
membership string), split their rows into the **summer rows** and the **other rows**,
then assign exactly one bucket. Order matters — first match wins:

| # | Bucket | Rule |
|---|---|---|
| 1 | **Converted – active** | Has an "other" enrollment that is **not** a School Partnership, status `Enrolled`, whose start date is on/after the earliest summer start date |
| 2 | **Converted – on hold** | Same as #1 but the "other" enrollment's status is `On Hold` |
| 3 | **Continued via School Partnership** | Has an "other" enrollment whose membership type contains `School Partnership`, start on/after the earliest summer start (and #1/#2 did not match) |
| 4 | **Summer still active** | No conversion; the latest summer row's `Primary Enrollment End` is >= today |
| 5 | **Did not return** | No conversion; all summer rows ended before today |
| 6 | **Needs review** | None of the above (e.g. only a `Pre-Enrolled` summer row) |

### Edge cases

- **Multiple qualifying "other" enrollments:** the "Converted to" detail comes from the
  one with the **latest start date**.
- **Converted and later put on hold:** if the student has both an `Enrolled` and an
  `On Hold` qualifying non-SP enrollment, **on-hold wins** (bucket 2), and the detail
  comes from the latest-start qualifying row.
- **A non-summer enrollment that started *before* the summer enrollment** (e.g. a spring
  membership that lapsed) does not count as a conversion.
- **Unrecognized status strings** (anything outside the four known values) are passed
  through verbatim into the detail table and land the student in **Needs review**.

### Conversion detail

For buckets 1–3, capture from the latest-start qualifying "other" enrollment:
membership type (verbatim), start date, `Monthly Amount`.

---

## Output

A single self-contained HTML file: `output/summer_2026_status.html`, opened in the
default browser when the run finishes. No email, no attachments. Inline CSS.

### Structure

1. **Header** — "2026 Summer Package — Enrollment Status", generated timestamp,
   "as of" date used for the still-active cutoff.
2. **Overall summary** — a table with one row per bucket and columns **Teaneck |
   Englewood | Overall**, showing counts. Two footer rows: **Total summer students** and
   **Conversion rate** = (Converted – active + Converted – on hold + Continued via
   School Partnership) ÷ total summer students, per column.
   - Rationale for including School Partnership in the rate: the student is still
     attending. It is still called out as its own bucket above.
3. **Teaneck section:**
   - One-line conversion-rate summary.
   - Student detail table, sorted by bucket (in the order above) then student name:

     | Student | Grade | Summer start | Summer end | Summer status | Current status | Continued as | New start | Monthly $ |
     |---|---|---|---|---|---|---|---|---|

     Blank cells where a field does not apply.
4. **Englewood section** — identical structure.
5. If any student is in **Needs review**, a highlighted callout above the sections names
   the count.

---

## Project Structure

**Location:** `~/mathnasium-summer-status` (new repo, sibling of the other automations).
**GitHub:** repo under the `mdiamond77` org. Org-level secrets cover Radius creds for any
future CI; local runs read `~/.mathnasium_env`.
**No `.github/workflows/`** — run on demand only.

| File | Responsibility |
|---|---|
| `config.py` | URLs, selectors, center config, per-center summer membership strings, pull-window dates, export column-name constants |
| `discover_selectors.py` | The headless discovery script already run on 2026-09-07; kept for re-running if Radius changes |
| `download.py` | Playwright headless: login via `~/.mathnasium_env`, pull Enrollment Report per center (4/1/2026 → 12/31/2026, all statuses), save raw export to `input/EnrollmentReport_<center>.xlsx` |
| `transform.py` | pandas pure functions: load + normalize an export, identify summer students, classify each per the table above, build per-center and overall summaries |
| `report.py` | Render `output/summer_2026_status.html` from transformed data; open it |
| `main.py` | CLI orchestrator: `--no-download` flag; download → transform → report |
| `tests/test_transform.py` | Unit tests for every classification bucket and edge case |
| `requirements.txt` | playwright, pandas, openpyxl, pytest |
| `.gitignore` | `input/`, `output/`, `__pycache__/`, `*.pyc`, `.env` |

---

## CLI Interface

```
python main.py                 # log in, pull both centers, transform, render, open report
python main.py --no-download    # re-render from cached input/ (fast iteration / quick refresh)
```

No date arguments — the summer membership strings and the pull window are constants in
`config.py`. The "as of" date for the still-active cutoff is the system date at run time.

---

## Error Handling

Fail loud and early — never emit a misleading empty or partial report:

- Login failure (still on the login page after submit) → stop with a clear message.
- A center's export has zero rows → stop (almost certainly a scrape problem).
- A center's summer membership string matches zero students → stop and name the string
  (the label may have changed in Radius — update `config.py`).
- Students in **Needs review** or a same-name collision → the report still renders, and
  the run prints a summary line with the counts.

---

## Testing

- `transform.py` is pure functions over pandas DataFrames. `tests/test_transform.py`
  builds small in-memory fixtures covering: each of the six buckets, multiple qualifying
  "other" enrollments (latest-start wins), converted-then-hold (on-hold wins), a
  pre-summer non-summer enrollment (not a conversion), and an unrecognized status string.
- `download.py` and `report.py` are verified manually against the real pull during the
  build. The two real exports pulled during discovery are cached in `input/` and double
  as manual-verification fixtures.

---

## Build Order

1. ~~`discover_selectors.py` — confirm the grid (gate).~~ **Done 2026-09-07.**
2. `config.py` — fill in real selectors, column names, membership strings (from findings).
3. `download.py` — reproduce the discovery pull as clean, reusable code for both centers.
4. `transform.py` + tests (TDD).
5. `report.py`.
6. `main.py` wiring + full manual run + spot-check against Radius.
7. README + create GitHub repo under `mdiamond77`.
