# Attendance Alerts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Weekly Monday email to each center director listing students who attended fewer than 75% of their allowed sessions over the past 2 weeks and/or 4 weeks.

**Architecture:** New repo `mathnasium-attendance-alerts`. Playwright downloads the Enrolled Report and Digital Workout Plan (28-day window) from Radius by reading Kendo grid data directly (same pattern as `mathnasium-hold-reminders`). pandas counts sessions (with ≥120 min = 2 sessions rule), computes attendance %, and categorizes flagged students into three buckets per center. Gmail SMTP sends one HTML email per center with three sections ordered by priority: Flagged in Both (highest), 2-Week Only, 4-Week Only.

**Tech Stack:** Python 3.11, Playwright (headless Chromium), pandas, openpyxl, smtplib SSL, GitHub Actions

---

### Task 1: Repo scaffold

**Files:**
- Create: `~/mathnasium-attendance-alerts/config.py`
- Create: `~/mathnasium-attendance-alerts/run_log.py`
- Create: `~/mathnasium-attendance-alerts/requirements.txt`
- Create: `~/mathnasium-attendance-alerts/.gitignore`
- Create: `~/mathnasium-attendance-alerts/tests/__init__.py`

- [ ] **Step 1: Create repo directory and initialize git**

```bash
mkdir ~/mathnasium-attendance-alerts
cd ~/mathnasium-attendance-alerts
git init
mkdir -p input tests
touch tests/__init__.py
```

- [ ] **Step 2: Create `.gitignore`**

```
input/
__pycache__/
*.pyc
.env
*.xlsx
```

- [ ] **Step 3: Create `requirements.txt`**

```
playwright==1.44.0
pandas==2.2.2
openpyxl==3.1.2
pytest==8.2.0
```

- [ ] **Step 4: Create `config.py`**

```python
import os
from pathlib import Path

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"

RADIUS_LOGIN_URL = "https://radius.mathnasium.com"
ENROLLED_REPORT_URL = "https://radius.mathnasium.com/Enrollment/EnrollmentReport"
WORKOUT_PLAN_URL = "https://radius.mathnasium.com/DigitalWorkoutPlan/Report"

CENTERS = {
    "Englewood": {"radius_id": "2428", "recipient": "englewood@mathnasium.com"},
    "Teaneck":   {"radius_id": "2871", "recipient": "teaneck@mathnasium.com"},
}
CC_RECIPIENT = "matt.diamond@mathnasium.com"

ATTENDANCE_THRESHOLD = 0.75
SESSION_DURATION_THRESHOLD_MIN = 120
MONTHLY_AMOUNT_CUTOFF = 350

EXCLUDE_NAME_PATTERNS = ["test", "appoin", "diamond"]
EXCLUDE_EXACT_NAMES = {"x x"}

COL_ACCOUNT_ID    = "Account Id"
COL_FIRST_NAME    = "Student First Name"
COL_LAST_NAME     = "Student Last Name"
COL_CENTER        = "Center"
COL_MONTHLY_AMT   = "Monthly Amount"
COL_STATUS        = "Status"
COL_DATE          = "Date"
COL_SESSION_START = "Session Start"
COL_SESSION_END   = "Session End"
```

- [ ] **Step 5: Create `run_log.py`** (identical to `mathnasium-hold-reminders/run_log.py`)

```python
import json
from datetime import datetime, timezone
from pathlib import Path

LOG_PATH = Path(__file__).parent / "run_log.json"


def read_log() -> list:
    if not LOG_PATH.exists():
        return []
    try:
        with open(LOG_PATH) as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def write_log(run_date: str, success: bool, trigger: str = "auto", error: str = "") -> None:
    log = read_log()
    log.append({
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "trigger": trigger,
        "run_date": run_date,
        "success": success,
        "error": error,
    })
    with open(LOG_PATH, "w") as f:
        json.dump(log, f, indent=2)
```

- [ ] **Step 6: Initial commit**

```bash
cd ~/mathnasium-attendance-alerts
git add .
git commit -m "chore: initial repo scaffold"
```

---

### Task 2: `process.py` — session counting helpers

**Files:**
- Create: `~/mathnasium-attendance-alerts/process.py`
- Create: `~/mathnasium-attendance-alerts/tests/test_process.py`

- [ ] **Step 1: Write failing tests for `_to_duration_minutes`, `_session_count`, `_sessions_per_week`, `_is_excluded`**

Create `tests/test_process.py`:

```python
import pandas as pd
import pytest
from datetime import date
from process import _to_duration_minutes, _session_count, _sessions_per_week, _is_excluded


# ── _to_duration_minutes ──────────────────────────────────────────────────────

def test_duration_normal_session():
    assert _to_duration_minutes("4:00 PM", "5:00 PM") == 60.0

def test_duration_two_hour_session():
    assert _to_duration_minutes("3:00 PM", "5:00 PM") == 120.0

def test_duration_negative_returns_zero():
    # Midnight crossover or bad data
    assert _to_duration_minutes("11:30 PM", "12:30 AM") == 0.0

def test_duration_bad_value_returns_zero():
    assert _to_duration_minutes(None, None) == 0.0


# ── _session_count ────────────────────────────────────────────────────────────

def _make_wp(account_ids, dates, starts, ends):
    return pd.DataFrame({
        "Account Id": account_ids,
        "Date": pd.to_datetime(dates),
        "Session Start": starts,
        "Session End": ends,
    })


def test_session_count_basic():
    df = _make_wp(
        ["A", "A", "A", "A"],
        ["2026-04-15", "2026-04-16", "2026-04-17", "2026-04-01"],
        ["4:00 PM", "4:00 PM", "4:00 PM", "4:00 PM"],
        ["5:00 PM", "5:00 PM", "5:00 PM", "5:00 PM"],
    )
    result = _session_count(df, date(2026, 4, 10))
    assert result["A"] == 3  # Apr 15/16/17 in window; Apr 1 excluded


def test_session_count_excludes_on_cutoff_date():
    df = _make_wp(["A"], ["2026-04-10"], ["4:00 PM"], ["5:00 PM"])
    result = _session_count(df, date(2026, 4, 10))
    assert result.get("A", 0) == 0  # cutoff is exclusive


def test_session_count_two_hour_counts_as_two():
    df = _make_wp(["A"], ["2026-04-15"], ["3:00 PM"], ["5:00 PM"])
    result = _session_count(df, date(2026, 4, 10))
    assert result["A"] == 2


def test_session_count_negative_duration_counts_as_one():
    df = _make_wp(["A"], ["2026-04-15"], ["11:30 PM"], ["12:30 AM"])
    result = _session_count(df, date(2026, 4, 10))
    assert result["A"] == 1


def test_session_count_missing_student_returns_zero():
    df = _make_wp(["A"], ["2026-04-15"], ["4:00 PM"], ["5:00 PM"])
    result = _session_count(df, date(2026, 4, 10))
    assert result.get("NOTEXIST", 0) == 0


# ── _sessions_per_week ────────────────────────────────────────────────────────

def test_sessions_per_week_above_cutoff():
    assert _sessions_per_week(429.0) == 2

def test_sessions_per_week_at_cutoff():
    assert _sessions_per_week(350.0) == 2

def test_sessions_per_week_below_cutoff():
    assert _sessions_per_week(250.0) == 1

def test_sessions_per_week_none_defaults_to_one():
    assert _sessions_per_week(None) == 1


# ── _is_excluded ──────────────────────────────────────────────────────────────

def test_is_excluded_test_name():
    assert _is_excluded("AppointTest Student1") is True

def test_is_excluded_diamond():
    assert _is_excluded("Bari Diamond") is True

def test_is_excluded_exact_x_x():
    assert _is_excluded("x x") is True

def test_is_excluded_normal_name():
    assert _is_excluded("Alice Smith") is False

def test_is_excluded_case_insensitive():
    assert _is_excluded("DIAMOND Test") is True
```

- [ ] **Step 2: Run tests — confirm they all fail**

```bash
cd ~/mathnasium-attendance-alerts
pytest tests/test_process.py -v 2>&1 | head -40
```

Expected: all fail with `ImportError` or `ModuleNotFoundError`.

- [ ] **Step 3: Create `process.py` with helper implementations**

```python
import re
from datetime import date, timedelta

import pandas as pd

from config import (
    ATTENDANCE_THRESHOLD,
    SESSION_DURATION_THRESHOLD_MIN,
    MONTHLY_AMOUNT_CUTOFF,
    EXCLUDE_NAME_PATTERNS,
    EXCLUDE_EXACT_NAMES,
    CENTERS,
    COL_ACCOUNT_ID,
    COL_FIRST_NAME,
    COL_LAST_NAME,
    COL_CENTER,
    COL_MONTHLY_AMT,
    COL_STATUS,
    COL_DATE,
    COL_SESSION_START,
    COL_SESSION_END,
)

_EXCLUDE_RE = [re.compile(p, re.IGNORECASE) for p in EXCLUDE_NAME_PATTERNS]


def _to_duration_minutes(start, end) -> float:
    for fmt in ('%I:%M %p', '%H:%M:%S', '%H:%M'):
        try:
            s = pd.to_datetime(str(start).strip(), format=fmt)
            e = pd.to_datetime(str(end).strip(), format=fmt)
            diff = (e - s).total_seconds() / 60
            return diff if diff > 0 else 0.0
        except (ValueError, TypeError):
            continue
    return 0.0


def _session_count(wp_df: pd.DataFrame, cutoff: date) -> dict:
    df = wp_df.copy()
    df['_date'] = pd.to_datetime(df[COL_DATE]).dt.date
    df = df[df['_date'] > cutoff]
    df['_duration'] = df.apply(
        lambda r: _to_duration_minutes(r[COL_SESSION_START], r[COL_SESSION_END]), axis=1
    )
    df['_count'] = df['_duration'].apply(
        lambda d: 2 if d >= SESSION_DURATION_THRESHOLD_MIN else 1
    )
    return df.groupby(COL_ACCOUNT_ID)['_count'].sum().to_dict()


def _sessions_per_week(monthly_amount) -> int:
    try:
        return 2 if float(monthly_amount) >= MONTHLY_AMOUNT_CUTOFF else 1
    except (TypeError, ValueError):
        return 1


def _is_excluded(name: str) -> bool:
    lower = name.lower().strip()
    if lower in EXCLUDE_EXACT_NAMES:
        return True
    return any(p.search(lower) for p in _EXCLUDE_RE)
```

- [ ] **Step 4: Run helper tests — confirm they pass**

```bash
cd ~/mathnasium-attendance-alerts
pytest tests/test_process.py -v -k "duration or session_count or sessions_per_week or is_excluded"
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
cd ~/mathnasium-attendance-alerts
git add process.py tests/test_process.py
git commit -m "feat: add process.py session counting helpers"
```

---

### Task 3: `process.py` — student flagging and main entry point

**Files:**
- Modify: `~/mathnasium-attendance-alerts/process.py`
- Modify: `~/mathnasium-attendance-alerts/tests/test_process.py`

- [ ] **Step 1: Write failing tests for `flag_students` and `process`**

Append to `tests/test_process.py`:

```python
from process import flag_students, process
from pathlib import Path
import tempfile


# ── Shared fixtures ───────────────────────────────────────────────────────────

TODAY = date(2026, 4, 28)  # a Monday

def _enrolled(account_ids, first_names, last_names, centers, amounts, statuses=None):
    n = len(account_ids)
    return pd.DataFrame({
        "Account Id":         account_ids,
        "Student First Name": first_names,
        "Student Last Name":  last_names,
        "Center":             centers,
        "Monthly Amount":     amounts,
        "Status":             statuses or ["Enrolled"] * n,
    })


def _wp(account_ids, dates, starts=None, ends=None):
    n = len(account_ids)
    return pd.DataFrame({
        "Account Id":    account_ids,
        "Date":          pd.to_datetime(dates),
        "Session Start": starts or ["4:00 PM"] * n,
        "Session End":   ends   or ["5:00 PM"] * n,
    })


# ── flag_students ─────────────────────────────────────────────────────────────

def test_flag_students_both_flagged():
    # Alice (2x/week = $429) had 1 session in past 28 days and 0 in past 14 days → both flagged
    enrolled = _enrolled(["A1"], ["Alice"], ["Smith"], ["Englewood"], [429])
    # 1 session 20 days ago — inside 4wk window, outside 2wk window
    wp = _wp(["A1"], ["2026-04-08"])
    result = flag_students(enrolled, wp, TODAY)
    assert len(result["Englewood"]["both"]) == 1
    assert result["Englewood"]["both"][0]["name"] == "Alice Smith"


def test_flag_students_only_2wk():
    # Bob (1x/week = $250): 0 sessions in past 14 days (0/2 = 0% → flagged),
    # but 3 sessions in past 28 days (3/4 = 75% → NOT flagged, at threshold)
    # → appears in only_2wk
    enrolled = _enrolled(["B1"], ["Bob"], ["Jones"], ["Teaneck"], [250])
    # 3 sessions between 15-28 days ago (in 4wk but not 2wk)
    wp = _wp(["B1", "B1", "B1"],
             ["2026-04-06", "2026-04-07", "2026-04-08"])
    result = flag_students(enrolled, wp, TODAY)
    assert len(result["Teaneck"]["only_2wk"]) == 1
    assert result["Teaneck"]["only_2wk"][0]["name"] == "Bob Jones"


def test_flag_students_only_4wk():
    # Carol (2x/week = $429): 3 sessions in past 14 days (3/4 = 75% → NOT flagged),
    # but only 4 sessions in past 28 days (4/8 = 50% → flagged)
    # → appears in only_4wk
    enrolled = _enrolled(["C1"], ["Carol"], ["White"], ["Englewood"], [429])
    # 1 session in 4wk-only window (Apr 8, after Mar 31 but not after Apr 14)
    # 3 sessions in 2wk window (Apr 15, 20, 22, after Apr 14)
    wp = _wp(["C1"] * 4,
             ["2026-04-08", "2026-04-15", "2026-04-20", "2026-04-22"])
    result = flag_students(enrolled, wp, TODAY)
    assert len(result["Englewood"]["only_4wk"]) == 1


def test_flag_students_at_threshold_not_flagged():
    # Exactly 75% is NOT flagged (threshold is strict <)
    # 2x/week student: 3/4 in 2wk = 75%, 6/8 in 4wk = 75% → neither flagged
    enrolled = _enrolled(["D1"], ["Dave"], ["Brown"], ["Teaneck"], [429])
    # 3 sessions after Apr 14 (in 2wk), 3 sessions Apr 7-9 (in 4wk only) = 6 total in 4wk
    wp = _wp(["D1"] * 6,
             ["2026-04-07", "2026-04-08", "2026-04-09",
              "2026-04-15", "2026-04-20", "2026-04-22"])
    result = flag_students(enrolled, wp, TODAY)
    assert len(result["Teaneck"]["both"]) == 0
    assert len(result["Teaneck"]["only_2wk"]) == 0
    assert len(result["Teaneck"]["only_4wk"]) == 0


def test_flag_students_excluded_not_included():
    enrolled = _enrolled(["E1"], ["Test"], ["Student"], ["Englewood"], [250])
    wp = _wp(["E1"], ["2026-04-15"])
    result = flag_students(enrolled, wp, TODAY)
    all_students = (
        result["Englewood"]["both"] +
        result["Englewood"]["only_2wk"] +
        result["Englewood"]["only_4wk"]
    )
    assert all(s["name"] != "Test Student" for s in all_students)


def test_flag_students_unenrolled_not_included():
    enrolled = _enrolled(["F1"], ["Frank"], ["Lee"], ["Teaneck"], [250], ["Inactive"])
    wp = _wp([], [])
    result = flag_students(enrolled, wp, TODAY)
    all_students = (
        result["Teaneck"]["both"] +
        result["Teaneck"]["only_2wk"] +
        result["Teaneck"]["only_4wk"]
    )
    assert len(all_students) == 0


def test_flag_students_all_centers_present():
    enrolled = _enrolled([], [], [], [], [])
    wp = _wp([], [])
    result = flag_students(enrolled, wp, TODAY)
    assert "Englewood" in result
    assert "Teaneck" in result


def test_flag_students_uses_enrolled_center():
    # Student enrolled at Englewood goes to Englewood bucket
    enrolled = _enrolled(["G1"], ["Grace"], ["Kim"], ["Englewood"], [429])
    wp = _wp([], [])  # 0 sessions → flagged both
    result = flag_students(enrolled, wp, TODAY)
    assert len(result["Englewood"]["both"]) == 1
    assert len(result["Teaneck"]["both"]) == 0


def test_flag_students_student_dict_keys():
    enrolled = _enrolled(["H1"], ["Hana"], ["Park"], ["Teaneck"], [250])
    wp = _wp([], [])
    result = flag_students(enrolled, wp, TODAY)
    student = result["Teaneck"]["both"][0]
    assert "name" in student
    assert "sessions_per_week" in student
    assert "sessions_2wk" in student
    assert "allowed_2wk" in student
    assert "pct_2wk" in student
    assert "sessions_4wk" in student
    assert "allowed_4wk" in student
    assert "pct_4wk" in student


# ── process (integration) ─────────────────────────────────────────────────────

def test_process_reads_excel_files():
    enrolled = _enrolled(["Z1"], ["Zero"], ["Sessions"], ["Teaneck"], [429])
    wp = _wp([], [])
    with tempfile.TemporaryDirectory() as d:
        enrolled_path = Path(d) / "enrolled.xlsx"
        wp_path = Path(d) / "wp.xlsx"
        enrolled.to_excel(enrolled_path, index=False)
        wp.to_excel(wp_path, index=False)
        result = process(enrolled_path, wp_path, TODAY)
    assert "Teaneck" in result
    assert len(result["Teaneck"]["both"]) == 1
```

- [ ] **Step 2: Run new tests — confirm they fail**

```bash
cd ~/mathnasium-attendance-alerts
pytest tests/test_process.py -v -k "flag_students or test_process_reads" 2>&1 | head -30
```

Expected: fail with `ImportError`.

- [ ] **Step 3: Add `flag_students` and `process` to `process.py`**

Append to `process.py`:

```python
def flag_students(enrolled_df: pd.DataFrame, wp_df: pd.DataFrame, today: date = None) -> dict:
    if today is None:
        today = date.today()

    cutoff_2wk = today - timedelta(weeks=2)
    cutoff_4wk = today - timedelta(weeks=4)

    sessions_2wk = _session_count(wp_df, cutoff_2wk)
    sessions_4wk = _session_count(wp_df, cutoff_4wk)

    enrolled = enrolled_df[enrolled_df[COL_STATUS] == "Enrolled"].copy()
    enrolled[COL_MONTHLY_AMT] = pd.to_numeric(enrolled[COL_MONTHLY_AMT], errors="coerce")

    result = {center: {"both": [], "only_2wk": [], "only_4wk": []} for center in CENTERS}

    for _, row in enrolled.iterrows():
        name = "{} {}".format(row[COL_FIRST_NAME], row[COL_LAST_NAME])
        if _is_excluded(name):
            continue
        center = row[COL_CENTER]
        if center not in CENTERS:
            continue

        account_id = row[COL_ACCOUNT_ID]
        spw = _sessions_per_week(row[COL_MONTHLY_AMT])
        allowed_2wk = spw * 2
        allowed_4wk = spw * 4
        s2 = sessions_2wk.get(account_id, 0)
        s4 = sessions_4wk.get(account_id, 0)
        pct_2wk = s2 / allowed_2wk
        pct_4wk = s4 / allowed_4wk

        flag_2wk = pct_2wk < ATTENDANCE_THRESHOLD
        flag_4wk = pct_4wk < ATTENDANCE_THRESHOLD

        if not flag_2wk and not flag_4wk:
            continue

        student = {
            "name": name,
            "sessions_per_week": spw,
            "sessions_2wk": int(s2),
            "allowed_2wk": allowed_2wk,
            "pct_2wk": pct_2wk,
            "sessions_4wk": int(s4),
            "allowed_4wk": allowed_4wk,
            "pct_4wk": pct_4wk,
        }

        if flag_2wk and flag_4wk:
            result[center]["both"].append(student)
        elif flag_2wk:
            result[center]["only_2wk"].append(student)
        else:
            result[center]["only_4wk"].append(student)

    for center in result:
        for section in result[center]:
            result[center][section].sort(key=lambda s: s["name"])

    return result


def process(enrolled_path, workout_plan_path, today: date = None) -> dict:
    enrolled_df = pd.read_excel(enrolled_path)
    wp_df = pd.read_excel(workout_plan_path)
    return flag_students(enrolled_df, wp_df, today)
```

- [ ] **Step 4: Run all process tests — confirm they pass**

```bash
cd ~/mathnasium-attendance-alerts
pip install -r requirements.txt -q
pytest tests/test_process.py -v
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
cd ~/mathnasium-attendance-alerts
git add process.py tests/test_process.py
git commit -m "feat: add flag_students and process entry point"
```

---

### Task 4: `deliver.py`

**Files:**
- Create: `~/mathnasium-attendance-alerts/deliver.py`
- Create: `~/mathnasium-attendance-alerts/tests/test_deliver.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_deliver.py`:

```python
from deliver import build_html, _fmt


BOTH_STUDENT = {
    "name": "Adina Luks",
    "sessions_per_week": 2,
    "sessions_2wk": 2,
    "allowed_2wk": 4,
    "pct_2wk": 0.5,
    "sessions_4wk": 5,
    "allowed_4wk": 8,
    "pct_4wk": 0.625,
}
WK2_STUDENT = {
    "name": "Eden Kim",
    "sessions_per_week": 1,
    "sessions_2wk": 0,
    "allowed_2wk": 2,
    "pct_2wk": 0.0,
    "sessions_4wk": 3,
    "allowed_4wk": 4,
    "pct_4wk": 0.75,
}
WK4_STUDENT = {
    "name": "Moshe Wigod",
    "sessions_per_week": 2,
    "sessions_2wk": 3,
    "allowed_2wk": 4,
    "pct_2wk": 0.75,
    "sessions_4wk": 3,
    "allowed_4wk": 8,
    "pct_4wk": 0.375,
}

CENTER_DATA = {
    "both":     [BOTH_STUDENT],
    "only_2wk": [WK2_STUDENT],
    "only_4wk": [WK4_STUDENT],
}

WEEK_LABEL = "April 28, 2026"


# ── _fmt ──────────────────────────────────────────────────────────────────────

def test_fmt_basic():
    assert _fmt(2, 4) == "2/4 (50%)"

def test_fmt_zero():
    assert _fmt(0, 4) == "0/4 (0%)"

def test_fmt_full():
    assert _fmt(8, 8) == "8/8 (100%)"


# ── build_html ────────────────────────────────────────────────────────────────

def test_build_html_is_valid_html():
    html = build_html("Teaneck", WEEK_LABEL, CENTER_DATA)
    assert html.strip().startswith("<html")
    assert "</html>" in html

def test_build_html_happy_monday_greeting():
    html = build_html("Teaneck", WEEK_LABEL, CENTER_DATA)
    assert "Happy Monday" in html

def test_build_html_contains_center_name():
    html = build_html("Englewood", WEEK_LABEL, CENTER_DATA)
    assert "Englewood" in html

def test_build_html_contains_week_label():
    html = build_html("Teaneck", WEEK_LABEL, CENTER_DATA)
    assert WEEK_LABEL in html

def test_build_html_contains_both_student():
    html = build_html("Teaneck", WEEK_LABEL, CENTER_DATA)
    assert "Adina Luks" in html

def test_build_html_contains_2wk_student():
    html = build_html("Teaneck", WEEK_LABEL, CENTER_DATA)
    assert "Eden Kim" in html

def test_build_html_contains_4wk_student():
    html = build_html("Teaneck", WEEK_LABEL, CENTER_DATA)
    assert "Moshe Wigod" in html

def test_build_html_both_section_shows_both_windows():
    html = build_html("Teaneck", WEEK_LABEL, CENTER_DATA)
    # Both-flagged students should show 2-week AND 4-week columns
    assert "2/4 (50%)" in html
    assert "5/8 (62%)" in html

def test_build_html_2wk_only_section():
    html = build_html("Teaneck", WEEK_LABEL, {"both": [], "only_2wk": [WK2_STUDENT], "only_4wk": []})
    assert "Eden Kim" in html

def test_build_html_empty_sections_not_shown():
    html = build_html("Teaneck", WEEK_LABEL, {"both": [BOTH_STUDENT], "only_2wk": [], "only_4wk": []})
    assert "Last 2 Weeks Only" not in html
    assert "Last 4 Weeks Only" not in html

def test_build_html_auto_send_note():
    html = build_html("Teaneck", WEEK_LABEL, CENTER_DATA)
    assert "automatically" in html
```

- [ ] **Step 2: Run tests — confirm they fail**

```bash
cd ~/mathnasium-attendance-alerts
pytest tests/test_deliver.py -v 2>&1 | head -20
```

Expected: all fail with `ImportError`.

- [ ] **Step 3: Create `deliver.py`**

```python
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from config import CENTERS, CC_RECIPIENT


def _fmt(sessions: int, allowed: int) -> str:
    pct = int(round(sessions / allowed * 100))
    return "{}/{} ({}%)".format(sessions, allowed, pct)


def _th(text: str) -> str:
    return (
        "<th style='padding:6px 12px;border:1px solid #ddd;"
        "background:#f5f5f5;text-align:left;'>{}</th>"
    ).format(text)


def _td(text) -> str:
    return "<td style='padding:6px 12px;border:1px solid #ddd;'>{}</td>".format(text)


def _table(headers: list, rows: list) -> str:
    header_html = "<tr>{}</tr>".format("".join(_th(h) for h in headers))
    rows_html = "".join("<tr>{}</tr>".format("".join(_td(c) for c in row)) for row in rows)
    return (
        "<table style='border-collapse:collapse;font-family:Arial,sans-serif;"
        "font-size:14px;margin-bottom:16px;'>"
        "{}{}</table>"
    ).format(header_html, rows_html)


def _both_table(students: list) -> str:
    rows = [
        [
            s["name"],
            str(s["sessions_per_week"]),
            _fmt(s["sessions_2wk"], s["allowed_2wk"]),
            _fmt(s["sessions_4wk"], s["allowed_4wk"]),
        ]
        for s in students
    ]
    return _table(["Student", "Sessions/Week", "Last 2 Weeks", "Last 4 Weeks"], rows)


def _window_table(students: list, window_label: str, sessions_key: str, allowed_key: str) -> str:
    rows = [
        [
            s["name"],
            str(s["sessions_per_week"]),
            _fmt(s[sessions_key], s[allowed_key]),
        ]
        for s in students
    ]
    return _table(["Student", "Sessions/Week", window_label], rows)


def _section(title: str, subtitle: str, table_html: str) -> str:
    return (
        "<h3 style='font-family:Arial,sans-serif;margin-bottom:4px;'>{}</h3>"
        "<p style='font-family:Arial,sans-serif;font-size:13px;color:#666;"
        "margin-top:0;margin-bottom:8px;'>{}</p>"
        "{}"
    ).format(title, subtitle, table_html)


def build_html(center_name: str, week_label: str, center_data: dict) -> str:
    divider = "<hr style='border:none;border-top:2px solid #ccc;margin:20px 0;'>"
    sections_html = ""

    if center_data["both"]:
        sections_html += divider + _section(
            "\u26a0\ufe0f Flagged in Both Windows",
            "Highest priority — attendance below 75% in both the last 2 weeks and last 4 weeks.",
            _both_table(center_data["both"]),
        )

    if center_data["only_2wk"]:
        sections_html += divider + _section(
            "Last 2 Weeks Only",
            "Recent dip — below 75% in the past 2 weeks but on track over 4 weeks.",
            _window_table(
                center_data["only_2wk"], "Last 2 Weeks", "sessions_2wk", "allowed_2wk"
            ),
        )

    if center_data["only_4wk"]:
        sections_html += divider + _section(
            "Last 4 Weeks Only",
            "Sustained issue — below 75% over 4 weeks but recently improved.",
            _window_table(
                center_data["only_4wk"], "Last 4 Weeks", "sessions_4wk", "allowed_4wk"
            ),
        )

    return (
        "<html><body style='font-family:Arial,sans-serif;font-size:14px;"
        "max-width:750px;margin:0 auto;padding:20px;'>"
        "<p>Happy Monday!</p>"
        "<p>The following <strong>{center}</strong> students attended fewer than 75% of their "
        "allowed sessions. Please reach out as needed.</p>"
        "{sections}"
        "{divider}"
        "<p style='color:#999;font-size:12px;'><em>This email sends automatically every Monday "
        "for the week of {week}. Questions? Contact matt.diamond@mathnasium.com.</em></p>"
        "</body></html>"
    ).format(center=center_name, sections=sections_html, divider=divider, week=week_label)


def send_email(center_name: str, week_label: str, html: str) -> None:
    smtp_user = os.environ.get("SMTP_USER")
    smtp_password = os.environ.get("SMTP_PASSWORD")
    if not smtp_user or not smtp_password:
        raise EnvironmentError("SMTP_USER and SMTP_PASSWORD environment variables must be set.")
    recipient = CENTERS[center_name]["recipient"]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Attendance Alerts \u2014 {} \u2014 Week of {}".format(center_name, week_label)
    msg["From"] = smtp_user
    msg["To"] = recipient
    msg["Cc"] = CC_RECIPIENT

    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, [recipient, CC_RECIPIENT], msg.as_string())


def deliver(all_center_data: dict, week_label: str) -> None:
    for center_name, center_data in all_center_data.items():
        total = sum(len(center_data[s]) for s in ["both", "only_2wk", "only_4wk"])
        if total == 0:
            print("No flagged students for {}, skipping.".format(center_name))
            continue
        html = build_html(center_name, week_label, center_data)
        send_email(center_name, week_label, html)
        print("Sent email for {}".format(center_name))
```

- [ ] **Step 4: Run tests — confirm they pass**

```bash
cd ~/mathnasium-attendance-alerts
pytest tests/test_deliver.py -v
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
cd ~/mathnasium-attendance-alerts
git add deliver.py tests/test_deliver.py
git commit -m "feat: add deliver.py HTML builder and SMTP sender"
```

---

### Task 5: `main.py`

**Files:**
- Create: `~/mathnasium-attendance-alerts/main.py`
- Create: `~/mathnasium-attendance-alerts/tests/test_main.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_main.py`:

```python
from datetime import date
from main import get_week_label


def test_get_week_label_format():
    assert get_week_label(date(2026, 4, 28)) == "April 28, 2026"

def test_get_week_label_single_digit_day():
    assert get_week_label(date(2026, 5, 4)) == "May 4, 2026"

def test_get_week_label_december():
    assert get_week_label(date(2026, 12, 7)) == "December 7, 2026"
```

- [ ] **Step 2: Run tests — confirm they fail**

```bash
cd ~/mathnasium-attendance-alerts
pytest tests/test_main.py -v 2>&1 | head -20
```

Expected: fail with `ImportError`.

- [ ] **Step 3: Create `main.py`**

```python
import argparse
import traceback
from datetime import date

from download import download_reports
from process import process
from deliver import deliver
from run_log import write_log


def get_week_label(today: date = None) -> str:
    if today is None:
        today = date.today()
    return today.strftime("%B %-d, %Y")


def main():
    parser = argparse.ArgumentParser(description="Run Attendance Alerts automation.")
    parser.add_argument(
        "--trigger",
        default="auto",
        choices=["auto", "manual"],
        help="How this run was triggered (default: auto)",
    )
    args = parser.parse_args()

    today = date.today()
    week_label = get_week_label(today)
    run_date = today.strftime("%Y-%m-%d")

    print("Running Attendance Alerts for week of {}".format(week_label))

    try:
        print("Downloading reports...")
        paths = download_reports()

        print("Processing data...")
        center_data = process(paths["enrolled"], paths["workout_plan"])

        print("Sending emails...")
        deliver(center_data, week_label)

        write_log(run_date, success=True, trigger=args.trigger)
        print("Done.")

    except Exception as e:
        error_msg = traceback.format_exc()
        print("ERROR: {}".format(e))
        print(error_msg)
        write_log(run_date, success=False, trigger=args.trigger, error=error_msg)
        raise


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Run tests — confirm they pass**

```bash
cd ~/mathnasium-attendance-alerts
pytest tests/test_main.py -v
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
cd ~/mathnasium-attendance-alerts
git add main.py tests/test_main.py
git commit -m "feat: add main.py orchestrator"
```

---

### Task 6: `download.py` — Playwright (requires page discovery)

**Files:**
- Create: `~/mathnasium-attendance-alerts/download.py`

This task requires discovering the Radius UI selectors before writing the final implementation. Follow these steps carefully.

- [ ] **Step 1: Install dependencies**

```bash
cd ~/mathnasium-attendance-alerts
pip install -r requirements.txt -q
playwright install chromium
```

- [ ] **Step 2: Discover Enrolled Report selectors**

Run a headed Playwright session to inspect the Enrolled Report page:

```python
# run_once: discover_enrolled.py (delete after use)
import os
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_context().new_page()
    page.goto("https://radius.mathnasium.com")
    page.fill("#UserName", os.environ["RADIUS_USERNAME"])
    page.fill("#Password", os.environ["RADIUS_PASSWORD"])
    page.click("#login")
    page.wait_for_load_state("networkidle")
    page.goto("https://radius.mathnasium.com/Enrollment/EnrollmentReport")
    page.wait_for_load_state("networkidle")
    input("Inspect the page — note the grid ID and any Export button. Press Enter to close.")
    browser.close()
```

Run it:

```bash
cd ~/mathnasium-attendance-alerts
RADIUS_USERNAME=$(grep RADIUS_USERNAME ~/.mathnasium_env | cut -d= -f2) \
RADIUS_PASSWORD=$(grep RADIUS_PASSWORD ~/.mathnasium_env | cut -d= -f2) \
python discover_enrolled.py
```

Look for:
- The Kendo grid element ID (e.g., `#gridEnrollmentReport`)
- The JavaScript property names in the grid's data source (hover over rows, inspect network requests, or use browser console: `$('[id*=grid]').data('kendoGrid').dataSource.data()[0].toJSON()`)

Note the grid ID and the JS field names for: `Account Id`, `Student First Name`, `Student Last Name`, `Center`, `Monthly Amount`, `Status`.

- [ ] **Step 3: Discover Workout Plan selectors**

Create and run a similar discovery script for the Workout Plan page:

```python
# run_once: discover_wp.py (delete after use)
import os
from datetime import date, timedelta
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    page = browser.new_context().new_page()
    page.goto("https://radius.mathnasium.com")
    page.fill("#UserName", os.environ["RADIUS_USERNAME"])
    page.fill("#Password", os.environ["RADIUS_PASSWORD"])
    page.click("#login")
    page.wait_for_load_state("networkidle")
    page.goto("https://radius.mathnasium.com/DigitalWorkoutPlan/Report")
    page.wait_for_load_state("networkidle")
    input("Inspect date filter inputs and grid ID. Press Enter to close.")
    browser.close()
```

```bash
RADIUS_USERNAME=$(grep RADIUS_USERNAME ~/.mathnasium_env | cut -d= -f2) \
RADIUS_PASSWORD=$(grep RADIUS_PASSWORD ~/.mathnasium_env | cut -d= -f2) \
python discover_wp.py
```

Look for:
- Date range input IDs (likely `#ReportStart` / `#ReportEnd` like the Holds Report — confirm)
- Whether they are Kendo DatePickers (check via browser console: `$('#ReportStart').data('kendoDatePicker')`)
- The search/apply button ID (likely `#btnsearch`)
- The grid ID (e.g., `#gridDigitalWorkoutPlanReport`)
- JS field names for: `AccountId`, `Date`, `SessionStart`/`StartTime`, `SessionEnd`/`EndTime`

- [ ] **Step 4: Create `download.py` using discovered selectors**

Use the pattern from `mathnasium-hold-reminders/download.py`. Replace `ENROLLED_GRID_ID`, `WP_GRID_ID`, and JS field names with what you discovered in Steps 2-3.

```python
import os
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

from config import (
    INPUT_DIR,
    RADIUS_LOGIN_URL,
    ENROLLED_REPORT_URL,
    WORKOUT_PLAN_URL,
)


def _login(page) -> None:
    username = os.environ.get("RADIUS_USERNAME")
    password = os.environ.get("RADIUS_PASSWORD")
    if not username or not password:
        raise EnvironmentError("RADIUS_USERNAME and RADIUS_PASSWORD must be set.")
    page.goto(RADIUS_LOGIN_URL)
    page.fill("#UserName", username)
    page.fill("#Password", password)
    page.click("#login")
    page.wait_for_load_state("networkidle")


def download_enrolled_report(page) -> Path:
    out_path = INPUT_DIR / "enrolled.xlsx"
    page.goto(ENROLLED_REPORT_URL)
    page.wait_for_load_state("networkidle")

    # Click search to load all enrolled students
    page.click("#btnsearch")
    page.wait_for_load_state("networkidle")

    # Wait for Kendo grid to finish binding
    # REPLACE "gridEnrollmentReport" with the actual grid ID discovered in Step 2
    page.wait_for_function(
        "() => { var g = $('#gridEnrollmentReport').data('kendoGrid'); "
        "return g && g.dataSource && !g.dataSource._requestInProgress; }",
        timeout=30000,
    )

    # Read data from Kendo grid datasource
    # REPLACE JS field names (AccountId, FirstName, etc.) with actual names from Step 2
    records = page.evaluate("""
        () => {
            var grid = $('#gridEnrollmentReport').data('kendoGrid');
            return grid.dataSource.data().map(function(item) {
                var d = item.toJSON();
                return {
                    'Account Id':          d.AccountId,
                    'Student First Name':  d.FirstName,
                    'Student Last Name':   d.LastName,
                    'Center':              d.CenterName,
                    'Monthly Amount':      d.MonthlyAmount,
                    'Status':              d.Status,
                };
            });
        }
    """)

    pd.DataFrame(records).to_excel(out_path, index=False)
    return out_path


def download_workout_plan(page) -> Path:
    out_path = INPUT_DIR / "workout_plan.xlsx"
    today = date.today()
    start = today - timedelta(days=28)

    # Format as M/D/YYYY (Radius convention)
    date_from = "{}/{}/{}".format(start.month, start.day, start.year)
    date_to   = "{}/{}/{}".format(today.month, today.day, today.year)

    page.goto(WORKOUT_PLAN_URL)
    page.wait_for_load_state("networkidle")

    # Set date range — use Kendo DatePicker API if available (same as HoldsReport)
    # VERIFY these IDs match what you found in Step 3; update if different
    page.evaluate(
        """([from_, to_]) => {
            var dpStart = $('#ReportStart').data('kendoDatePicker');
            var dpEnd   = $('#ReportEnd').data('kendoDatePicker');
            if (dpStart) { dpStart.value(from_); dpStart.trigger('change'); }
            if (dpEnd)   { dpEnd.value(to_);     dpEnd.trigger('change'); }
        }""",
        [date_from, date_to],
    )

    page.click("#btnsearch")
    page.wait_for_load_state("networkidle")

    # Wait for Kendo grid
    # REPLACE "gridDigitalWorkoutPlanReport" with actual grid ID from Step 3
    page.wait_for_function(
        "() => { var g = $('#gridDigitalWorkoutPlanReport').data('kendoGrid'); "
        "return g && g.dataSource && !g.dataSource._requestInProgress; }",
        timeout=60000,
    )

    # Read data from grid
    # REPLACE JS field names with actual names discovered in Step 3
    records = page.evaluate("""
        () => {
            var grid = $('#gridDigitalWorkoutPlanReport').data('kendoGrid');
            return grid.dataSource.data().map(function(item) {
                var d = item.toJSON();
                return {
                    'Account Id':    d.AccountId,
                    'Date':          d.Date,
                    'Session Start': d.SessionStart,
                    'Session End':   d.SessionEnd,
                };
            });
        }
    """)

    pd.DataFrame(records).to_excel(out_path, index=False)
    return out_path


def download_reports() -> dict:
    INPUT_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        with browser:
            context = browser.new_context()
            with context:
                page = context.new_page()
                _login(page)
                enrolled_path = download_enrolled_report(page)
                wp_path = download_workout_plan(page)

    return {"enrolled": enrolled_path, "workout_plan": wp_path}
```

- [ ] **Step 5: Test download manually**

```bash
cd ~/mathnasium-attendance-alerts
RADIUS_USERNAME=$(grep RADIUS_USERNAME ~/.mathnasium_env | cut -d= -f2) \
RADIUS_PASSWORD=$(grep RADIUS_PASSWORD ~/.mathnasium_env | cut -d= -f2) \
python -c "
from download import download_reports
import pandas as pd
paths = download_reports()
print('Enrolled:', paths['enrolled'])
print(pd.read_excel(paths['enrolled']).head(3)[['Account Id','Student First Name','Center','Monthly Amount','Status']])
print()
print('Workout Plan:', paths['workout_plan'])
print(pd.read_excel(paths['workout_plan']).head(3)[['Account Id','Date','Session Start','Session End']])
"
```

Expected: both files download successfully and show correct columns.

If the grid IDs or JS field names are wrong, inspect the actual page data and update `download.py` accordingly.

- [ ] **Step 6: Clean up discovery scripts**

```bash
cd ~/mathnasium-attendance-alerts
rm -f discover_enrolled.py discover_wp.py
```

- [ ] **Step 7: Commit**

```bash
cd ~/mathnasium-attendance-alerts
git add download.py
git commit -m "feat: add download.py Playwright report downloader"
```

---

### Task 7: GitHub Actions workflow

**Files:**
- Create: `~/mathnasium-attendance-alerts/.github/workflows/attendance_alerts.yml`

- [ ] **Step 1: Create workflow file**

```bash
mkdir -p ~/mathnasium-attendance-alerts/.github/workflows
```

Create `.github/workflows/attendance_alerts.yml`:

```yaml
name: Attendance Alerts

on:
  schedule:
    - cron: "0 15 * * 1"   # Every Monday at 11am ET (3pm UTC)
  workflow_dispatch:         # Manual trigger for testing

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
        run: |
          if [ "${{ github.event_name }}" = "workflow_dispatch" ]; then
            python main.py --trigger manual
          else
            python main.py --trigger auto
          fi

      - name: Commit run log
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add run_log.json || true
          git diff --cached --quiet || git commit -m "chore: update run log [skip ci]"
          git push
```

- [ ] **Step 2: Create GitHub repo and push**

```bash
cd ~/mathnasium-attendance-alerts
gh repo create mdiamond77/mathnasium-attendance-alerts --private --source=. --remote=origin --push
```

- [ ] **Step 3: Confirm secrets are inherited**

```bash
gh secret list --org mdiamond77 2>/dev/null || \
gh secret list -R mdiamond77/mathnasium-attendance-alerts
```

Expected: `RADIUS_USERNAME`, `RADIUS_PASSWORD`, `SMTP_USER`, `SMTP_PASSWORD` all listed. If not, add them:

```bash
gh secret set RADIUS_USERNAME -R mdiamond77/mathnasium-attendance-alerts
gh secret set RADIUS_PASSWORD -R mdiamond77/mathnasium-attendance-alerts
gh secret set SMTP_USER -R mdiamond77/mathnasium-attendance-alerts
gh secret set SMTP_PASSWORD -R mdiamond77/mathnasium-attendance-alerts
```

- [ ] **Step 4: Commit workflow and push**

```bash
cd ~/mathnasium-attendance-alerts
git add .github/workflows/attendance_alerts.yml
git commit -m "chore: add GitHub Actions workflow"
git push
```

---

### Task 8: End-to-end manual test

- [ ] **Step 1: Run all unit tests**

```bash
cd ~/mathnasium-attendance-alerts
pytest tests/ -v
```

Expected: all pass.

- [ ] **Step 2: Run the full pipeline locally against real Radius data**

```bash
cd ~/mathnasium-attendance-alerts
source ~/.mathnasium_env
SMTP_USER=$(grep SMTP_USER ~/.mathnasium_env | cut -d= -f2) \
SMTP_PASSWORD=$(grep SMTP_PASSWORD ~/.mathnasium_env | cut -d= -f2) \
python main.py --trigger manual
```

Expected output:
```
Running Attendance Alerts for week of [today's date]
Downloading reports...
Processing data...
Sending emails...
Sent email for Englewood
Sent email for Teaneck
Done.
```

Check that Matt receives the CC'd emails for both centers and that the three sections render correctly.

- [ ] **Step 3: Trigger manual GitHub Actions run**

```bash
gh workflow run attendance_alerts.yml -R mdiamond77/mathnasium-attendance-alerts
```

Wait ~3 minutes, then check:

```bash
gh run list -R mdiamond77/mathnasium-attendance-alerts --limit 1
```

Expected: `completed` / `success`.

- [ ] **Step 4: Add to automation dashboard**

Open the automation dashboard repo (same location as `mathnasium-hold-reminders` dashboard entry). Find where the other automations are registered (e.g., `config.py` or equivalent list of automations) and add an entry for this project following the exact same pattern:

```python
{
    "name": "Attendance Alerts",
    "repo": "mdiamond77/mathnasium-attendance-alerts",
    "run_log": "run_log.json",
    "description": "Weekly Monday email: students attending < 75% of allowed sessions",
}
```

Pull the `run_log.json` locally after the first Actions run to confirm it appears in the dashboard at `localhost:8080`.
