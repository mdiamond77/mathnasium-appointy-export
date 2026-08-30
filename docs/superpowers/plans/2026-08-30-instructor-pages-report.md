# Instructor Pages Report — Phase 1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local script that downloads the Radius Digital Workout Plan (DWP) report for April 1 → today and produces a 5-tab Excel workbook showing, per instructor, how many pages the average student completes — by calendar week and calendar month.

**Architecture:** New standalone repo `mathnasium-instructor-pages`, structured like `mathnasium-page-goals`. Pure date math lives in `periods.py`; DWP-DataFrame-to-metrics lives in `process.py`; the workbook writer lives in `excel_out.py`; `download.py` is a near-copy of the page-goals Radius scraper with a parameterized date window; `main.py` wires them together for a local run. No email, no cron, no dashboard in Phase 1 — those are Phase 2, after Matt reviews the first workbook.

**Tech Stack:** Python 3.12+, Playwright (Radius scraping), pandas (data), openpyxl (Excel), python-dateutil, pytest.

**Reference implementation:** `~/mathnasium-page-goals/` — copy its `download.py` login flow, `.gitignore`, and repo conventions.

**Status (2026-08-30): Phase 1 COMPLETE.** All 7 tasks done, 62 tests passing, repo pushed to `mdmathnasiums/mathnasium-instructor-pages`. First workbook (Apr 1–Aug 30) generated from live Radius data and delivered to Matt. Live check found the instructor column is `Instructors` (plural) — `config.py` updated. Phase 2 (email + cron + dashboard) awaits Matt's sign-off on the numbers.

---

## File Map

**New repo: `~/mathnasium-instructor-pages/`**

| File | Responsibility |
|---|---|
| `config.py` | Constants: Radius URLs, center IDs, instructor column name, season start, dir names |
| `periods.py` | Pure functions: Sunday-week bucketing, calendar-month bucketing, snapshot-window computation. No pandas. |
| `process.py` | DWP DataFrame → normalized session rows → per-instructor/per-period metrics dict |
| `excel_out.py` | Metrics dict + raw DataFrame → formatted `.xlsx` (Snapshot, Monthly, Weekly, Detail, Data tabs) |
| `download.py` | Playwright: log into Radius, pull the DWP report for a `[start, end]` window |
| `main.py` | CLI orchestration: `--start` / `--end` / `--skip-download`; download → process → write |
| `requirements.txt` | pandas, openpyxl, playwright, python-dateutil, pytest |
| `.gitignore` | copied from page-goals |
| `tests/__init__.py` | empty |
| `tests/test_periods.py` | date-math coverage |
| `tests/test_process.py` | classification + metrics + report-assembly coverage |
| `tests/test_excel_out.py` | workbook structure + cell-value coverage |

**The metrics dict** returned by `process.build_report(df, season_start, as_of)` — every later task depends on this exact shape:

```python
{
    "as_of": date,
    "season_start": date,
    "snapshot_labels": ["This Month So Far", "Last Week", "Last Month", "Last 2 Months", "Last 3 Months"],
    "month_labels": ["Apr 2026", "May 2026", ...],
    "week_labels": ["Wk of 3/29", "Wk of 4/5", ...],
    "periods": {
        "<label>": {"type": "snapshot" | "month" | "week", "start": date, "end": date},
        ...
    },
    "centers": {
        "Teaneck": {
            "instructors": ["Alice", "Bob", "(Unassigned)"],   # sorted, "(Unassigned)" last
            "cells": {
                ("<instructor>", "<period label>"): {
                    "pages_per_session": float | None,     # None when productive_sessions == 0
                    "student_avg": float | None,
                    "productive_sessions": int,
                    "zero_blank_sessions": int,
                    "low_sessions": int,
                    "distinct_students": int,
                },
                ...
            },
        },
        "Englewood": { ... },
    },
}
```

---

## Task 1: Repo Scaffold

**Files:**
- Create: `~/mathnasium-instructor-pages/` (new git repo)
- Create: `.gitignore`, `requirements.txt`, `config.py`
- Create: `input/`, `output/` directories

- [ ] **Step 1: Create and init the repo**

```bash
cd ~
mkdir mathnasium-instructor-pages
cd mathnasium-instructor-pages
git init
```

- [ ] **Step 2: Create `.gitignore`**

```
.venv/
__pycache__/
*.pyc
input/
output/
*.env
```

- [ ] **Step 3: Create `requirements.txt`**

```
pandas
openpyxl
playwright
python-dateutil
pytest
```

- [ ] **Step 4: Install dependencies**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

- [ ] **Step 5: Create `config.py`**

```python
from datetime import date

RADIUS_LOGIN_URL = "https://radius.mathnasium.com"
RADIUS_DWP_URL = "https://radius.mathnasium.com/DigitalWorkoutPlan/Report"

# Radius multiselect center IDs (same as mathnasium-page-goals)
CENTER_VALUES = ["2428", "2871"]  # 2428 = Englewood, 2871 = Teaneck

# Header of the instructor column in the DWP Excel export.
# Provisional — confirmed/corrected in Task 6 against a real download.
INSTRUCTOR_COL = "Instructor"

# Pages/session at or below this count flags a "low" session (still counted as productive).
LOW_SESSION_THRESHOLD = 3

# Start of the data window for the historical report.
SEASON_START = date(2026, 4, 1)

INPUT_DIR = "input"
OUTPUT_DIR = "output"
```

- [ ] **Step 6: Create directories**

```bash
mkdir -p input output tests
touch tests/__init__.py
```

- [ ] **Step 7: Commit**

```bash
git add .gitignore requirements.txt config.py tests/__init__.py
git commit -m "chore: scaffold mathnasium-instructor-pages"
```

- [ ] **Step 8: Create the private GitHub repo and push**

```bash
gh repo create mdmathnasiums/mathnasium-instructor-pages --private --source=. --push
```

---

## Task 2: `periods.py` — Date Bucketing

**Files:**
- Create: `~/mathnasium-instructor-pages/periods.py`
- Create: `~/mathnasium-instructor-pages/tests/test_periods.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_periods.py`:

```python
from datetime import date

from periods import (
    week_start,
    week_label,
    month_label,
    month_buckets,
    week_buckets,
    snapshot_windows,
)


# ── week_start: weeks run Sunday → Saturday ───────────────────────────────────

def test_week_start_of_a_sunday_is_itself():
    assert week_start(date(2026, 8, 30)) == date(2026, 8, 30)  # 2026-08-30 is a Sunday


def test_week_start_of_a_saturday_is_prior_sunday():
    assert week_start(date(2026, 8, 29)) == date(2026, 8, 23)


def test_week_start_of_a_wednesday():
    # 2026-04-01 is a Wednesday; its week began Sunday 2026-03-29
    assert week_start(date(2026, 4, 1)) == date(2026, 3, 29)


# ── labels ───────────────────────────────────────────────────────────────────

def test_week_label():
    assert week_label(date(2026, 3, 29)) == "Wk of 3/29"


def test_month_label():
    assert month_label(date(2026, 4, 1)) == "Apr 2026"


# ── month_buckets ────────────────────────────────────────────────────────────

def test_month_buckets_span_april_to_august():
    buckets = month_buckets(date(2026, 4, 1), date(2026, 8, 30))
    labels = [b[0] for b in buckets]
    assert labels == ["Apr 2026", "May 2026", "Jun 2026", "Jul 2026", "Aug 2026"]


def test_month_buckets_first_bucket_bounds():
    buckets = month_buckets(date(2026, 4, 1), date(2026, 8, 30))
    assert buckets[0][1] == date(2026, 4, 1)
    assert buckets[0][2] == date(2026, 4, 30)


def test_month_buckets_last_bucket_is_clipped_to_as_of():
    buckets = month_buckets(date(2026, 4, 1), date(2026, 8, 30))
    assert buckets[-1][1] == date(2026, 8, 1)
    assert buckets[-1][2] == date(2026, 8, 30)


# ── week_buckets ─────────────────────────────────────────────────────────────

def test_week_buckets_first_bucket_starts_on_sunday_before_season_start():
    buckets = week_buckets(date(2026, 4, 1), date(2026, 8, 30))
    assert buckets[0][0] == "Wk of 3/29"
    assert buckets[0][1] == date(2026, 3, 29)
    assert buckets[0][2] == date(2026, 4, 4)


def test_week_buckets_last_bucket_clipped_to_as_of():
    buckets = week_buckets(date(2026, 4, 1), date(2026, 8, 30))
    assert buckets[-1][1] == date(2026, 8, 30)
    assert buckets[-1][2] == date(2026, 8, 30)


def test_week_buckets_are_contiguous_seven_day_spans():
    buckets = week_buckets(date(2026, 4, 1), date(2026, 8, 30))
    for label, start, end in buckets[:-1]:
        assert (end - start).days == 6


# ── snapshot_windows (as of 2026-08-30, a Sunday) ────────────────────────────

def test_snapshot_this_month_so_far():
    windows = dict((w[0], (w[1], w[2])) for w in snapshot_windows(date(2026, 8, 30)))
    assert windows["This Month So Far"] == (date(2026, 8, 1), date(2026, 8, 30))


def test_snapshot_last_week_is_the_prior_complete_sun_sat_week():
    windows = dict((w[0], (w[1], w[2])) for w in snapshot_windows(date(2026, 8, 30)))
    assert windows["Last Week"] == (date(2026, 8, 23), date(2026, 8, 29))


def test_snapshot_last_month_is_the_prior_complete_calendar_month():
    windows = dict((w[0], (w[1], w[2])) for w in snapshot_windows(date(2026, 8, 30)))
    assert windows["Last Month"] == (date(2026, 7, 1), date(2026, 7, 31))


def test_snapshot_last_2_months_are_the_two_prior_complete_months():
    windows = dict((w[0], (w[1], w[2])) for w in snapshot_windows(date(2026, 8, 30)))
    assert windows["Last 2 Months"] == (date(2026, 6, 1), date(2026, 7, 31))


def test_snapshot_last_3_months_are_the_three_prior_complete_months():
    windows = dict((w[0], (w[1], w[2])) for w in snapshot_windows(date(2026, 8, 30)))
    assert windows["Last 3 Months"] == (date(2026, 5, 1), date(2026, 7, 31))


def test_snapshot_returns_five_windows_in_order():
    labels = [w[0] for w in snapshot_windows(date(2026, 8, 30))]
    assert labels == [
        "This Month So Far", "Last Week", "Last Month", "Last 2 Months", "Last 3 Months",
    ]
```

- [ ] **Step 2: Run the tests to confirm they fail**

```bash
cd ~/mathnasium-instructor-pages && source .venv/bin/activate
pytest tests/test_periods.py -v
```

Expected: `ModuleNotFoundError: No module named 'periods'`

- [ ] **Step 3: Implement `periods.py`**

```python
"""
periods.py
Pure date-bucketing for the instructor pages report.
Weeks run Sunday → Saturday. Months are calendar months.
No pandas — plain datetime.date only.
"""

from datetime import date, timedelta

from dateutil.relativedelta import relativedelta


def week_start(d: date) -> date:
    """Return the Sunday that begins the week containing d."""
    # Python: Monday=0 .. Sunday=6. For a Sunday-start week, Sunday must map to 0.
    return d - timedelta(days=(d.weekday() + 1) % 7)


def week_label(sunday: date) -> str:
    return f"Wk of {sunday.month}/{sunday.day}"


def month_label(d: date) -> str:
    return d.strftime("%b %Y")


def month_buckets(season_start: date, as_of: date) -> list[tuple[str, date, date]]:
    """(label, start, end) for every calendar month from season_start's month
    through as_of's month. The final bucket's end is clipped to as_of."""
    out: list[tuple[str, date, date]] = []
    cur = season_start.replace(day=1)
    last = as_of.replace(day=1)
    while cur <= last:
        nxt = cur + relativedelta(months=1)
        end = min(nxt - timedelta(days=1), as_of)
        out.append((month_label(cur), cur, end))
        cur = nxt
    return out


def week_buckets(season_start: date, as_of: date) -> list[tuple[str, date, date]]:
    """(label, start, end) for every Sun–Sat week overlapping
    [season_start, as_of]. First bucket starts on the Sunday on/before
    season_start; the final bucket's end is clipped to as_of."""
    out: list[tuple[str, date, date]] = []
    cur = week_start(season_start)
    while cur <= as_of:
        end = min(cur + timedelta(days=6), as_of)
        out.append((week_label(cur), cur, end))
        cur = cur + timedelta(days=7)
    return out


def snapshot_windows(as_of: date) -> list[tuple[str, date, date]]:
    """The five calendar-aligned windows shown on the Snapshot tab."""
    first_of_month = as_of.replace(day=1)

    current_week = week_start(as_of)
    last_week_start = current_week - timedelta(days=7)
    last_week_end = current_week - timedelta(days=1)

    last_month_start = first_of_month - relativedelta(months=1)
    last_month_end = first_of_month - timedelta(days=1)
    two_months_start = first_of_month - relativedelta(months=2)
    three_months_start = first_of_month - relativedelta(months=3)

    return [
        ("This Month So Far", first_of_month, as_of),
        ("Last Week", last_week_start, last_week_end),
        ("Last Month", last_month_start, last_month_end),
        ("Last 2 Months", two_months_start, last_month_end),
        ("Last 3 Months", three_months_start, last_month_end),
    ]
```

- [ ] **Step 4: Run the tests to confirm they pass**

```bash
pytest tests/test_periods.py -v
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add periods.py tests/test_periods.py
git commit -m "feat: add periods.py date bucketing with tests"
```

---

## Task 3: `process.py` — Session Classification & Metrics

**Files:**
- Create: `~/mathnasium-instructor-pages/process.py`
- Create: `~/mathnasium-instructor-pages/tests/test_process.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_process.py`:

```python
import pandas as pd
import pytest

from process import (
    to_pages,
    assign_center,
    split_instructors,
    normalize_sessions,
    session_metrics,
)


# ── to_pages ─────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("value, expected", [
    ("5", 5.0),
    (7, 7.0),
    (3.5, 3.5),
    (0, 0.0),
    ("", None),
    (None, None),
    ("abc", None),
])
def test_to_pages(value, expected):
    assert to_pages(value) == expected


# ── assign_center ────────────────────────────────────────────────────────────

@pytest.mark.parametrize("value, expected", [
    ("Englewood", "Englewood"),
    ("Teaneck", "Teaneck"),
    ("Teaneck, Teaneck Virtual", "Teaneck"),
    ("Englewood, Teaneck Virtual", "Englewood"),   # Englewood checked first
    ("Paramus", None),
    ("", None),
])
def test_assign_center(value, expected):
    assert assign_center(value) == expected


# ── split_instructors ────────────────────────────────────────────────────────

@pytest.mark.parametrize("value, expected", [
    ("Alice", ["Alice"]),
    ("  Alice  ", ["Alice"]),
    ("Alice, Bob", ["Alice", "Bob"]),
    ("Alice / Bob", ["Alice", "Bob"]),
    ("Alice; Bob", ["Alice", "Bob"]),
    ("Alice and Bob", ["Alice", "Bob"]),
    ("", ["(Unassigned)"]),
    (None, ["(Unassigned)"]),
])
def test_split_instructors(value, expected):
    assert split_instructors(value) == expected


# ── normalize_sessions ───────────────────────────────────────────────────────

def _raw(rows):
    """rows: list of (date, student, pages, center, instructor)"""
    return pd.DataFrame(rows, columns=["Date", "Student Name", "Pages Completed", "Center", "Instructor"])


def test_normalize_drops_rows_with_no_recognized_center():
    df = _raw([
        ("2026-04-06", "Sam", 5, "Paramus", "Alice"),
        ("2026-04-06", "Kim", 5, "Teaneck", "Alice"),
    ])
    out = normalize_sessions(df)
    assert list(out["student"]) == ["Kim"]


def test_normalize_explodes_multi_instructor_rows():
    df = _raw([("2026-04-06", "Sam", 6, "Teaneck", "Alice, Bob")])
    out = normalize_sessions(df)
    assert sorted(out["instructor"]) == ["Alice", "Bob"]
    assert list(out["pages"]) == [6.0, 6.0]


def test_normalize_blank_instructor_becomes_unassigned():
    df = _raw([("2026-04-06", "Sam", 6, "Teaneck", None)])
    out = normalize_sessions(df)
    assert list(out["instructor"]) == ["(Unassigned)"]


# ── session_metrics ──────────────────────────────────────────────────────────

def _sessions(rows):
    """rows: list of (student, pages) — all one center+instructor+window"""
    return pd.DataFrame(
        [{"student": s, "pages": (None if p is None else float(p))} for s, p in rows]
    )


def test_metrics_average_excludes_zero_and_blank():
    m = session_metrics(_sessions([("A", 4), ("A", 8), ("B", 0), ("B", None)]))
    assert m["pages_per_session"] == 6.0
    assert m["productive_sessions"] == 2
    assert m["zero_blank_sessions"] == 2


def test_metrics_low_session_count_includes_3_excludes_0_and_4():
    m = session_metrics(_sessions([("A", 3), ("A", 4), ("A", 0), ("A", 1)]))
    # productive: 3, 4, 1 → low (<=3): 3 and 1
    assert m["low_sessions"] == 2


def test_metrics_student_avg_weights_students_equally():
    rows = [("A", 10), ("A", 10), ("A", 10), ("B", 2)]
    m = session_metrics(_sessions(rows))
    assert m["pages_per_session"] == 8.0        # (10+10+10+2) / 4
    assert m["student_avg"] == 6.0              # mean(mean(A)=10, mean(B)=2)


def test_metrics_distinct_students_counts_only_productive():
    m = session_metrics(_sessions([("A", 5), ("B", 0), ("C", None)]))
    assert m["distinct_students"] == 1


def test_metrics_empty_slice_is_all_zero_with_none_averages():
    m = session_metrics(_sessions([]))
    assert m == {
        "pages_per_session": None,
        "student_avg": None,
        "productive_sessions": 0,
        "zero_blank_sessions": 0,
        "low_sessions": 0,
        "distinct_students": 0,
    }
```

- [ ] **Step 2: Run the tests to confirm they fail**

```bash
pytest tests/test_process.py -v
```

Expected: `ModuleNotFoundError: No module named 'process'`

- [ ] **Step 3: Implement the first half of `process.py`**

```python
"""
process.py
DWP Excel DataFrame → normalized session rows → per-instructor / per-period metrics.
"""

import re
from datetime import date

import pandas as pd

from config import INSTRUCTOR_COL, LOW_SESSION_THRESHOLD
from periods import month_buckets, week_buckets, snapshot_windows

DATE_COL = "Date"
STUDENT_COL = "Student Name"
PAGES_COL = "Pages Completed"
CENTER_COL = "Center"

UNASSIGNED = "(Unassigned)"
_SPLIT_RE = re.compile(r"\s*(?:,|/|;|\band\b)\s*")


def to_pages(value):
    """Float pages when the value is a real number, else None."""
    n = pd.to_numeric(value, errors="coerce")
    if pd.isna(n):
        return None
    return float(n)


def assign_center(value):
    """Map a raw Center string to 'Englewood' / 'Teaneck' / None.
    Englewood is checked first so 'Englewood, Teaneck Virtual' routes to Englewood."""
    s = str(value)
    if "Englewood" in s:
        return "Englewood"
    if "Teaneck" in s:
        return "Teaneck"
    return None


def split_instructors(value):
    """Raw instructor cell → list of names. Blank/None → ['(Unassigned)']."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return [UNASSIGNED]
    text = str(value).strip()
    if not text:
        return [UNASSIGNED]
    names = [p.strip() for p in _SPLIT_RE.split(text) if p.strip()]
    return names or [UNASSIGNED]


def normalize_sessions(df: pd.DataFrame) -> pd.DataFrame:
    """One row per (session, instructor).
    Columns: date (datetime64), student (str), center (str), instructor (str),
    pages (float or NaN). Rows with an unrecognized center or unparseable date
    are dropped."""
    if INSTRUCTOR_COL not in df.columns:
        raise KeyError(
            f"Instructor column {INSTRUCTOR_COL!r} not in the DWP export. "
            f"Columns present: {list(df.columns)}. Fix INSTRUCTOR_COL in config.py."
        )

    out = pd.DataFrame({
        "date": pd.to_datetime(df[DATE_COL], errors="coerce"),
        "student": df[STUDENT_COL].astype(str).str.strip(),
        "center": df[CENTER_COL].apply(assign_center),
        "pages": df[PAGES_COL].apply(to_pages),
        "instructor": df[INSTRUCTOR_COL].apply(split_instructors),
    })
    out = out[out["center"].notna() & out["date"].notna()]
    out = out.explode("instructor").reset_index(drop=True)
    return out


def session_metrics(rows: pd.DataFrame) -> dict:
    """rows: normalized session rows already sliced to one center + instructor + window.
    Must contain 'student' and 'pages' columns."""
    pages = rows["pages"] if len(rows) else pd.Series(dtype=float)
    productive = rows[pages.notna() & (pages > 0)] if len(rows) else rows
    n_prod = len(productive)

    result = {
        "pages_per_session": None,
        "student_avg": None,
        "productive_sessions": n_prod,
        "zero_blank_sessions": int(len(rows) - n_prod),
        "low_sessions": int((productive["pages"] <= LOW_SESSION_THRESHOLD).sum()) if n_prod else 0,
        "distinct_students": int(productive["student"].nunique()) if n_prod else 0,
    }
    if n_prod:
        result["pages_per_session"] = round(float(productive["pages"].mean()), 2)
        per_student = productive.groupby("student")["pages"].mean()
        result["student_avg"] = round(float(per_student.mean()), 2)
    return result
```

- [ ] **Step 4: Run the tests to confirm they pass**

```bash
pytest tests/test_process.py -v
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add process.py tests/test_process.py
git commit -m "feat: add process.py session classification and metrics with tests"
```

---

## Task 4: `process.py` — Report Assembly

**Files:**
- Modify: `~/mathnasium-instructor-pages/process.py` (append `build_report`)
- Modify: `~/mathnasium-instructor-pages/tests/test_process.py` (append report tests)

- [ ] **Step 1: Append the failing tests**

Append to `tests/test_process.py`:

```python
from datetime import date
from process import build_report


def _dwp(rows):
    """rows: list of (date_str, student, pages, center, instructor)"""
    return pd.DataFrame(rows, columns=["Date", "Student Name", "Pages Completed", "Center", "Instructor"])


def test_build_report_has_both_center_keys():
    df = _dwp([("2026-04-06", "Sam", 5, "Teaneck", "Alice")])
    report = build_report(df, date(2026, 4, 1), date(2026, 8, 30))
    assert set(report["centers"]) == {"Teaneck", "Englewood"}


def test_build_report_instructors_sorted_with_unassigned_last():
    df = _dwp([
        ("2026-04-06", "S1", 5, "Teaneck", "Zoe"),
        ("2026-04-06", "S2", 5, "Teaneck", "amy"),
        ("2026-04-06", "S3", 5, "Teaneck", None),
    ])
    report = build_report(df, date(2026, 4, 1), date(2026, 8, 30))
    assert report["centers"]["Teaneck"]["instructors"] == ["amy", "Zoe", "(Unassigned)"]


def test_build_report_cell_for_a_month_bucket():
    df = _dwp([
        ("2026-04-06", "Sam", 4, "Teaneck", "Alice"),
        ("2026-04-13", "Sam", 8, "Teaneck", "Alice"),
        ("2026-05-04", "Sam", 2, "Teaneck", "Alice"),   # different month
    ])
    report = build_report(df, date(2026, 4, 1), date(2026, 8, 30))
    cell = report["centers"]["Teaneck"]["cells"][("Alice", "Apr 2026")]
    assert cell["productive_sessions"] == 2
    assert cell["pages_per_session"] == 6.0


def test_build_report_period_with_no_sessions_is_zeroed():
    df = _dwp([("2026-04-06", "Sam", 4, "Teaneck", "Alice")])
    report = build_report(df, date(2026, 4, 1), date(2026, 8, 30))
    cell = report["centers"]["Teaneck"]["cells"][("Alice", "Jul 2026")]
    assert cell["productive_sessions"] == 0
    assert cell["pages_per_session"] is None


def test_build_report_exposes_period_metadata():
    df = _dwp([("2026-04-06", "Sam", 4, "Teaneck", "Alice")])
    report = build_report(df, date(2026, 4, 1), date(2026, 8, 30))
    assert report["periods"]["Apr 2026"]["type"] == "month"
    assert report["periods"]["Wk of 3/29"]["type"] == "week"
    assert report["periods"]["Last Week"]["type"] == "snapshot"
    assert report["snapshot_labels"][0] == "This Month So Far"
```

- [ ] **Step 2: Run to confirm the new tests fail**

```bash
pytest tests/test_process.py -k build_report -v
```

Expected: `ImportError: cannot import name 'build_report'`

- [ ] **Step 3: Append `build_report` to `process.py`**

```python
def build_report(df: pd.DataFrame, season_start: date, as_of: date) -> dict:
    """DWP DataFrame → the metrics dict consumed by excel_out.write_workbook."""
    sessions = normalize_sessions(df)

    snaps = snapshot_windows(as_of)
    months = month_buckets(season_start, as_of)
    weeks = week_buckets(season_start, as_of)

    periods: dict[str, dict] = {}
    for label, start, end in snaps:
        periods[label] = {"type": "snapshot", "start": start, "end": end}
    for label, start, end in months:
        periods[label] = {"type": "month", "start": start, "end": end}
    for label, start, end in weeks:
        periods[label] = {"type": "week", "start": start, "end": end}

    report = {
        "as_of": as_of,
        "season_start": season_start,
        "snapshot_labels": [p[0] for p in snaps],
        "month_labels": [p[0] for p in months],
        "week_labels": [p[0] for p in weeks],
        "periods": periods,
        "centers": {},
    }

    for center in ("Teaneck", "Englewood"):
        cs = sessions[sessions["center"] == center]
        instructors = sorted(
            cs["instructor"].unique().tolist(),
            key=lambda n: (n == UNASSIGNED, n.lower()),
        )
        cells: dict[tuple[str, str], dict] = {}
        for inst in instructors:
            inst_rows = cs[cs["instructor"] == inst]
            for label, meta in periods.items():
                lo = pd.Timestamp(meta["start"])
                hi = pd.Timestamp(meta["end"]) + pd.Timedelta(days=1)
                window = inst_rows[(inst_rows["date"] >= lo) & (inst_rows["date"] < hi)]
                cells[(inst, label)] = session_metrics(window)
        report["centers"][center] = {"instructors": instructors, "cells": cells}

    return report
```

- [ ] **Step 4: Run the full process test file**

```bash
pytest tests/test_process.py -v
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add process.py tests/test_process.py
git commit -m "feat: add build_report assembly with tests"
```

---

## Task 5: `excel_out.py` — Workbook Writer

**Files:**
- Create: `~/mathnasium-instructor-pages/excel_out.py`
- Create: `~/mathnasium-instructor-pages/tests/test_excel_out.py`

- [ ] **Step 1: Write the failing tests**

Create `tests/test_excel_out.py`:

```python
from datetime import date

import openpyxl
import pandas as pd
import pytest

from process import build_report
from excel_out import grid_cell_value, write_workbook


def test_grid_cell_value_formats_avg_and_count():
    assert grid_cell_value({"pages_per_session": 4.25, "productive_sessions": 8}) == "4.25 (8)"


def test_grid_cell_value_blank_when_no_sessions():
    assert grid_cell_value({"pages_per_session": None, "productive_sessions": 0}) is None


@pytest.fixture
def workbook_path(tmp_path):
    df = pd.DataFrame(
        [
            ("2026-04-06", "Sam", 4, "Teaneck", "Alice"),
            ("2026-04-13", "Sam", 8, "Teaneck", "Alice"),
            ("2026-04-13", "Dana", 6, "Englewood", "Bob"),
        ],
        columns=["Date", "Student Name", "Pages Completed", "Center", "Instructor"],
    )
    report = build_report(df, date(2026, 4, 1), date(2026, 8, 30))
    out = tmp_path / "out.xlsx"
    write_workbook(report, df, str(out))
    return str(out)


def test_workbook_has_the_five_tabs(workbook_path):
    wb = openpyxl.load_workbook(workbook_path)
    assert wb.sheetnames == ["Snapshot", "Monthly", "Weekly", "Detail", "Data"]


def test_snapshot_tab_has_center_divider_rows(workbook_path):
    wb = openpyxl.load_workbook(workbook_path)
    text = "\n".join(
        str(c.value) for row in wb["Snapshot"].iter_rows() for c in row if c.value
    )
    assert "TEANECK" in text
    assert "ENGLEWOOD" in text


def test_detail_tab_has_one_row_per_instructor_period(workbook_path):
    wb = openpyxl.load_workbook(workbook_path)
    ws = wb["Detail"]
    header = [c.value for c in ws[1]]
    assert header[:4] == ["Center", "Instructor", "Period Type", "Period"]
    # Alice (Teaneck) + Bob (Englewood), each across every period, plus (Unassigned) rows
    assert ws.max_row > 1


def test_data_tab_preserves_raw_rows(workbook_path):
    wb = openpyxl.load_workbook(workbook_path)
    ws = wb["Data"]
    assert [c.value for c in ws[1]] == ["Date", "Student Name", "Pages Completed", "Center", "Instructor"]
    assert ws.max_row == 4  # header + 3 rows
```

- [ ] **Step 2: Run to confirm they fail**

```bash
pytest tests/test_excel_out.py -v
```

Expected: `ModuleNotFoundError: No module named 'excel_out'`

- [ ] **Step 3: Implement `excel_out.py`**

```python
"""
excel_out.py
Metrics dict (from process.build_report) + the raw DWP DataFrame → formatted .xlsx.

Tabs: Snapshot, Monthly, Weekly, Detail, Data.
Every analysis tab is split into a TEANECK block on top and an ENGLEWOOD block
below, separated by a full-width divider row. Centers are never comingled.
"""

import os

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

CENTERS = ("Teaneck", "Englewood")

_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
_DIVIDER_FILL = PatternFill("solid", fgColor="1F3864")
_GROUP_FILL_A = PatternFill("solid", fgColor="E7E6E6")
_GROUP_FILL_B = PatternFill("solid", fgColor="FFF2CC")
_WHITE_BOLD = Font(bold=True, color="FFFFFF")
_CENTER = Alignment(horizontal="center")
_LEFT = Alignment(horizontal="left")

# Snapshot sub-columns: (heading, metrics key)
_SNAP_SUBCOLS = [
    ("Pages/Sess", "pages_per_session"),
    ("Stu Avg", "student_avg"),
    ("Sessions", "productive_sessions"),
    ("≤3 pg", "low_sessions"),
]


def grid_cell_value(metrics: dict):
    """'avg (n)' for the Monthly/Weekly grids, or None when there were no sessions."""
    if metrics["productive_sessions"] == 0:
        return None
    return f"{metrics['pages_per_session']:.2f} ({metrics['productive_sessions']})"


def _style_header(cell):
    cell.fill = _HEADER_FILL
    cell.font = _WHITE_BOLD
    cell.alignment = _CENTER


def _write_divider(ws, row_idx, center_name, span):
    cell = ws.cell(row=row_idx, column=1, value=f"━━━ {center_name.upper()} ━━━")
    cell.fill = _DIVIDER_FILL
    cell.font = _WHITE_BOLD
    cell.alignment = _LEFT
    if span > 1:
        ws.merge_cells(
            start_row=row_idx, start_column=1, end_row=row_idx, end_column=span
        )


# ── Snapshot ─────────────────────────────────────────────────────────────────

def _write_snapshot(wb, report):
    ws = wb.create_sheet("Snapshot")
    windows = report["snapshot_labels"]
    total_cols = 1 + len(windows) * len(_SNAP_SUBCOLS)

    ws.column_dimensions["A"].width = 26
    for i in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 11

    row = 1
    for center in CENTERS:
        _write_divider(ws, row, center, total_cols)
        row += 1

        # window group header
        ws.cell(row=row, column=1, value="Instructor")
        _style_header(ws.cell(row=row, column=1))
        col = 2
        for w_idx, window in enumerate(windows):
            fill = _GROUP_FILL_A if w_idx % 2 == 0 else _GROUP_FILL_B
            start_col = col
            for sub_heading, _ in _SNAP_SUBCOLS:
                c = ws.cell(row=row + 1, column=col, value=sub_heading)
                c.fill = fill
                c.font = Font(bold=True)
                c.alignment = _CENTER
                col += 1
            gc = ws.cell(row=row, column=start_col, value=window)
            _style_header(gc)
            ws.merge_cells(
                start_row=row, start_column=start_col,
                end_row=row, end_column=col - 1,
            )
        row += 2

        cells = report["centers"][center]["cells"]
        for inst in report["centers"][center]["instructors"]:
            ws.cell(row=row, column=1, value=inst).alignment = _LEFT
            col = 2
            for window in windows:
                m = cells[(inst, window)]
                for _, key in _SNAP_SUBCOLS:
                    val = m[key]
                    ws.cell(row=row, column=col, value=val).alignment = _CENTER
                    col += 1
            row += 1

        # footnote: zero/blank sessions excluded, per window
        ws.cell(row=row, column=1, value="Zero/blank sessions excluded:").font = Font(italic=True)
        col = 2
        for window in windows:
            total_zb = sum(
                cells[(inst, window)]["zero_blank_sessions"]
                for inst in report["centers"][center]["instructors"]
            )
            fc = ws.cell(row=row, column=col, value=total_zb)
            fc.font = Font(italic=True)
            fc.alignment = _CENTER
            col += len(_SNAP_SUBCOLS)
        row += 2

    ws.freeze_panes = "B2"


# ── Monthly / Weekly grids ───────────────────────────────────────────────────

def _write_grid(wb, report, sheet_name, labels):
    ws = wb.create_sheet(sheet_name)
    total_cols = 1 + len(labels)

    ws.column_dimensions["A"].width = 26
    for i in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13

    row = 1
    for center in CENTERS:
        _write_divider(ws, row, center, total_cols)
        row += 1

        ws.cell(row=row, column=1, value="Instructor")
        _style_header(ws.cell(row=row, column=1))
        for i, label in enumerate(labels):
            _style_header(ws.cell(row=row, column=2 + i, value=label))
        row += 1

        cells = report["centers"][center]["cells"]
        for inst in report["centers"][center]["instructors"]:
            ws.cell(row=row, column=1, value=inst).alignment = _LEFT
            for i, label in enumerate(labels):
                ws.cell(
                    row=row, column=2 + i, value=grid_cell_value(cells[(inst, label)])
                ).alignment = _CENTER
            row += 1
        row += 1  # blank spacer before next center

    ws.freeze_panes = "B2"


# ── Detail (long format) ─────────────────────────────────────────────────────

_DETAIL_HEADERS = [
    "Center", "Instructor", "Period Type", "Period", "Period Start", "Period End",
    "Pages/Session", "Student Avg", "Productive Sessions", "Zero/Blank Sessions",
    "Low Sessions (<=3)", "Distinct Students",
]


def _write_detail(wb, report):
    ws = wb.create_sheet("Detail")
    for i, h in enumerate(_DETAIL_HEADERS, start=1):
        _style_header(ws.cell(row=1, column=i, value=h))
        ws.column_dimensions[get_column_letter(i)].width = 16

    row = 2
    period_order = (
        report["snapshot_labels"] + report["month_labels"] + report["week_labels"]
    )
    for center in CENTERS:
        cells = report["centers"][center]["cells"]
        for inst in report["centers"][center]["instructors"]:
            for label in period_order:
                meta = report["periods"][label]
                m = cells[(inst, label)]
                values = [
                    center, inst, meta["type"], label,
                    meta["start"].isoformat(), meta["end"].isoformat(),
                    m["pages_per_session"], m["student_avg"],
                    m["productive_sessions"], m["zero_blank_sessions"],
                    m["low_sessions"], m["distinct_students"],
                ]
                for i, v in enumerate(values, start=1):
                    ws.cell(row=row, column=i, value=v)
                row += 1
    ws.freeze_panes = "A2"


# ── Data (raw passthrough) ───────────────────────────────────────────────────

def _write_data(wb, raw_df: pd.DataFrame):
    ws = wb.create_sheet("Data")
    for i, col_name in enumerate(raw_df.columns, start=1):
        _style_header(ws.cell(row=1, column=i, value=str(col_name)))
        ws.column_dimensions[get_column_letter(i)].width = 18
    for r_idx, rec in enumerate(raw_df.itertuples(index=False), start=2):
        for c_idx, value in enumerate(rec, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    ws.freeze_panes = "A2"


# ── Entry point ──────────────────────────────────────────────────────────────

def write_workbook(report: dict, raw_df: pd.DataFrame, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _write_snapshot(wb, report)
    _write_grid(wb, report, "Monthly", report["month_labels"])
    _write_grid(wb, report, "Weekly", report["week_labels"])
    _write_detail(wb, report)
    _write_data(wb, raw_df)

    wb.save(output_path)
```

- [ ] **Step 4: Run the tests to confirm they pass**

```bash
pytest tests/test_excel_out.py -v
```

Expected: all tests pass.

- [ ] **Step 5: Run the full test suite**

```bash
pytest -v
```

Expected: every test passes.

- [ ] **Step 6: Commit**

```bash
git add excel_out.py tests/test_excel_out.py
git commit -m "feat: add excel_out.py workbook writer with tests"
```

---

## Task 6: `download.py` — Radius DWP Scraper

**Files:**
- Create: `~/mathnasium-instructor-pages/download.py`

This is a near-copy of `~/mathnasium-page-goals/download.py`. The only real change:
the date window is passed in as `(start, end)` instead of derived from a 90-day trailing rule.

- [ ] **Step 1: Implement `download.py`**

```python
"""
download.py
Logs into Radius and downloads the Digital Workout Plan report for a
[start, end] date window covering both centers. Returns the local file path.

Adapted from mathnasium-page-goals/download.py — same login flow and selectors.
"""

import json
import os
from datetime import date

from playwright.sync_api import sync_playwright

from config import RADIUS_LOGIN_URL, RADIUS_DWP_URL, CENTER_VALUES, INPUT_DIR


def download_dwp_report(start: date, end: date) -> str:
    """Download the DWP report for [start, end]. Returns the saved file path."""
    username = os.environ.get("RADIUS_USERNAME")
    password = os.environ.get("RADIUS_PASSWORD")
    if not username or not password:
        raise EnvironmentError("RADIUS_USERNAME and RADIUS_PASSWORD must be set.")

    os.makedirs(INPUT_DIR, exist_ok=True)
    output_path = os.path.join(
        INPUT_DIR,
        f"Digital_Workout_Plan_{start:%Y%m%d}_{end:%Y%m%d}.xlsx",
    )
    center_values_js = json.dumps(CENTER_VALUES)

    with sync_playwright() as p:
        with p.chromium.launch(headless=True) as browser:
            with browser.new_context(accept_downloads=True) as context:
                page = context.new_page()
                try:
                    print("[download] Logging into Radius...")
                    page.goto(RADIUS_LOGIN_URL)
                    page.wait_for_load_state("networkidle")
                    page.fill("#UserName", username)
                    page.fill("#Password", password)
                    page.click("#login")
                    page.wait_for_load_state("networkidle")
                    print("[download] Logged in.")

                    print("[download] Opening Digital Workout Plan report...")
                    page.goto(RADIUS_DWP_URL)
                    page.wait_for_load_state("networkidle")
                    page.wait_for_timeout(2000)

                    page.evaluate(f"""
                        var w = jQuery('#AllCenterListMultiSelect').data('kendoMultiSelect');
                        w.value({center_values_js});
                        w.trigger('change');
                    """)
                    page.wait_for_timeout(500)

                    page.evaluate(f"""
                        var s = jQuery('#dwpFromDate').data('kendoDatePicker');
                        s.value('{start:%m/%d/%Y}'); s.trigger('change');
                        var e = jQuery('#dwpToDate').data('kendoDatePicker');
                        e.value('{end:%m/%d/%Y}'); e.trigger('change');
                    """)
                    page.wait_for_timeout(500)

                    print("[download] Running search...")
                    page.click("#btnsearch")
                    page.wait_for_load_state("networkidle")
                    page.wait_for_selector("#dwpExcelBtn", state="visible", timeout=30_000)

                    print("[download] Downloading Excel...")
                    with page.expect_download(timeout=90_000) as dl:
                        page.click("#dwpExcelBtn")
                    dl.value.save_as(output_path)
                    print(f"[download] Saved to {output_path}")
                except Exception as e:
                    try:
                        page.screenshot(path="radius_error.png")
                        print("[download] Screenshot saved to radius_error.png")
                    except Exception:
                        pass
                    raise RuntimeError(f"[download] Radius automation failed: {e}") from e

    return output_path
```

- [ ] **Step 2: Commit**

```bash
git add download.py
git commit -m "feat: add Radius DWP downloader with parameterized date window"
```

---

## Task 7: `main.py` + Live Radius Verification + First Real Run

> **⚠️ This task needs Matt and a live Radius login.** The DWP export's exact column
> names are unconfirmed until we open a real download.

**Files:**
- Create: `~/mathnasium-instructor-pages/main.py`
- Possibly modify: `config.py` (`INSTRUCTOR_COL`), `process.py` (column constants)

- [ ] **Step 1: Implement `main.py`**

```python
"""
main.py
Local run: download the DWP report for [--start, --end] and write the
Instructor Pages workbook to output/.

Usage:
    python main.py                          # April 1 → today, fresh download
    python main.py --skip-download          # reuse newest file in input/
    python main.py --start 2026-04-01 --end 2026-08-30
"""

import argparse
import glob
import os
import sys
from datetime import date, datetime

import pandas as pd

from config import SEASON_START, INPUT_DIR, OUTPUT_DIR
from download import download_dwp_report
from process import build_report
from excel_out import write_workbook


def _parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", type=_parse_date, default=SEASON_START)
    parser.add_argument("--end", type=_parse_date, default=date.today())
    parser.add_argument(
        "--skip-download", action="store_true",
        help="Reuse the most recent Digital_Workout_Plan_*.xlsx already in input/",
    )
    args = parser.parse_args()

    if args.skip_download:
        found = sorted(glob.glob(os.path.join(INPUT_DIR, "Digital_Workout_Plan_*.xlsx")))
        if not found:
            sys.exit("[main] No file in input/ to reuse. Run without --skip-download.")
        input_path = found[-1]
        print(f"[main] Reusing {input_path}")
    else:
        print(f"[main] Downloading DWP report {args.start} → {args.end}")
        input_path = download_dwp_report(args.start, args.end)

    raw_df = pd.read_excel(input_path)
    print(f"[main] Loaded {len(raw_df)} rows. Columns: {list(raw_df.columns)}")

    report = build_report(raw_df, args.start, args.end)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUT_DIR, f"Instructor Pages {args.end.isoformat()}.xlsx")
    write_workbook(report, raw_df, output_path)
    print(f"[main] Wrote {output_path}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Commit `main.py`**

```bash
git add main.py
git commit -m "feat: add main.py local orchestration"
```

- [ ] **Step 3: Do one real download (with Matt)**

```bash
cd ~/mathnasium-instructor-pages && source .venv/bin/activate
RADIUS_USERNAME='<from ~/.mathnasium_env or page-goals secrets>' \
RADIUS_PASSWORD='<...>' \
python3 -c "
from datetime import date
from download import download_dwp_report
print(download_dwp_report(date(2026, 4, 1), date.today()))
"
```

Expected: a file at `input/Digital_Workout_Plan_20260401_<today>.xlsx`.
If it fails, inspect `radius_error.png` and compare selectors against
`~/mathnasium-page-goals/download.py` (that repo's last successful run confirms they work).

- [ ] **Step 4: Confirm the column names**

```bash
python3 -c "
import pandas as pd
df = pd.read_excel('input/Digital_Workout_Plan_20260401_<today>.xlsx')
print(list(df.columns))
print(df.head(5).to_string())
"
```

Verify these four columns exist with these exact names, used by `process.py`:
`Date`, `Student Name`, `Pages Completed`, `Center`.
Find the instructor column (candidates: `Instructor`, `Coach`, `Staff`, `Tutor`,
`Instructor Name`).

- **If the instructor column name differs from `"Instructor"`:** update
  `INSTRUCTOR_COL` in `config.py`.
- **If any of the four other headers differ:** update the corresponding constant
  (`DATE_COL`, `STUDENT_COL`, `PAGES_COL`, `CENTER_COL`) at the top of `process.py`.
- **If a single session spans multiple rows** (e.g. one row per subject): note it,
  stop, and raise with Matt — the "one row = one session" assumption needs revisiting
  before the numbers can be trusted.
- Check whether any instructor cell holds two names and what separates them; confirm
  `_SPLIT_RE` in `process.py` covers that separator.

- [ ] **Step 5: Commit any config/column corrections**

```bash
git add config.py process.py
git commit -m "fix: match DWP export column names from live download"
```

- [ ] **Step 6: Generate the workbook from the real data**

```bash
python3 main.py --skip-download
```

Expected: `output/Instructor Pages <today>.xlsx` written with 5 tabs.

- [ ] **Step 7: Eyeball the workbook**

Open `output/Instructor Pages <today>.xlsx` and sanity-check:
- Snapshot tab: TEANECK block on top, ENGLEWOOD below, five window groups each.
- Monthly tab: columns Apr 2026 → current month; cells read like `5.20 (14)`.
- Weekly tab: one column per Sun–Sat week since `Wk of 3/29`.
- Detail tab: one row per instructor per period, both centers.
- Data tab: matches the raw download row count.
- Spot-check one instructor's "Last Month" Pages/Sess against the Detail tab's
  matching calendar-month row — they should agree when "Last Month" == that month.

- [ ] **Step 8: Hand the workbook to Matt**

Send `output/Instructor Pages <today>.xlsx` to Matt for review. Phase 1 ends here.
Phase 2 (weekly email + Monday cron + dashboard row) begins only after he confirms
the numbers look right.

---

## Self-Review

**Spec coverage:**
- ✅ Same DWP report, wider window, both centers — Task 6 (`download.py`), Task 1 (`CENTER_VALUES`)
- ✅ April 1 → today window — Task 1 (`SEASON_START`), Task 7 (`main.py` defaults)
- ✅ Session classification (productive / zero-blank / low ≤3) — Task 3 (`session_metrics`)
- ✅ Averages exclude zero/blank; counts still reported — Task 3, Task 5 (footnote row)
- ✅ Pages/Session and Avg-of-student-averages side by side — Task 3, Task 5 (`_SNAP_SUBCOLS`)
- ✅ Sunday–Saturday weeks — Task 2 (`week_start`)
- ✅ Calendar months April→current — Task 2 (`month_buckets`)
- ✅ Snapshot windows incl. "This Month So Far" + last week/month/2mo/3mo — Task 2 (`snapshot_windows`)
- ✅ History back to April, weekly + monthly — Task 4 (`build_report`), Task 5 (`_write_grid`)
- ✅ Centers never comingled (top/bottom blocks + divider) — Task 5 (`_write_divider`)
- ✅ Instructor at both centers appears in each block, own-center sessions only — Task 4 (center filter before instructor loop)
- ✅ Blank instructor → `(Unassigned)`, sorted last — Task 3 (`split_instructors`), Task 4 (sort key)
- ✅ Multi-instructor session counted under each — Task 3 (`normalize_sessions` explode)
- ✅ Instructor with 0 sessions in a period → blank cells — Task 4 + Task 5 (`grid_cell_value`, snapshot writes `None`)
- ✅ 5 tabs: Snapshot, Monthly, Weekly, Detail, Data — Task 5 (`write_workbook`)
- ✅ Detail long format for pivoting — Task 5 (`_write_detail`)
- ✅ Raw Data tab — Task 5 (`_write_data`)
- ✅ Filename `Instructor Pages YYYY-MM-DD.xlsx` — Task 7 (`main.py`)
- ✅ Local run only, no email/cron/dashboard — Phase 1 scope; none of those tasks present
- ✅ `--skip-download` for formatting iteration — Task 7
- ✅ Clear error when instructor column missing — Task 3 (`normalize_sessions` KeyError)
- ✅ Open items (column names, co-taught, row granularity) — Task 7 Step 4

**Placeholder scan:** No "TBD"/"handle appropriately"/"similar to Task N" left. The
`<from ~/.mathnasium_env ...>` and `<today>` tokens in Task 7 are runtime values a human
substitutes during the live session, not unfilled plan content.

**Type consistency:** `build_report` returns the dict documented in the File Map;
`write_workbook`, `_write_snapshot`, `_write_grid`, `_write_detail` all read
`report["centers"][c]["cells"][(inst, label)]` with the six metric keys from
`session_metrics`. `grid_cell_value` reads `pages_per_session` + `productive_sessions`
— both present in every metrics dict. Period labels used as keys (`"Apr 2026"`,
`"Wk of 3/29"`, `"Last Week"`) are produced by `month_label` / `week_label` / the
`snapshot_windows` literals and reused verbatim in `build_report` and the writers.

---

## Out of Scope (Phase 2+)

- `deliver.py` — weekly email to Matt via Gmail SMTP
- `run_log.py` + `run_log.json`
- `.github/workflows/weekly_instructor_pages.yml` — Monday 13:00 UTC cron
- `/reports` dashboard row + hidden script entry + "Run Now" button
- The daily center summary email "Instructors" block (Phase 3, separate spec)
