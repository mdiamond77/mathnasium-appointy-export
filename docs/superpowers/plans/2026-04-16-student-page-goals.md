# Student Page Goals Automation — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a monthly automation that downloads the Radius Digital Workout Plan Report, calculates personalized student page goals, delivers a formatted Excel to email and Google Drive, and shows run history on a new `/reports` dashboard page.

**Architecture:** New repo `mathnasium-page-goals` with four modules (download, process, deliver, main) following the pattern of `radius-cc-lists` and `radius-morning-briefing`. The existing `automation-dashboard` repo gains a `/reports` page that reads `run_log.json` from the page goals repo.

**Tech Stack:** Python 3.14, Playwright (Radius scraping), pandas + openpyxl (Excel), smtplib (Gmail SMTP), google-api-python-client (Drive upload), Flask (dashboard), GitHub Actions (scheduling)

---

## File Map

**New repo: `~/mathnasium-page-goals/`**
- `config.py` — constants (recipients, folder ID, paths)
- `run_log.py` — read/write `run_log.json`
- `process.py` — calculation logic + Excel output
- `download.py` — Playwright: login to Radius, download DWP report
- `deliver.py` — email + Google Drive upload
- `main.py` — orchestration, `--month` flag, `--trigger` flag
- `run_log.json` — run history (committed back by GitHub Actions)
- `requirements.txt`
- `.github/workflows/monthly_page_goals.yml`
- `tests/test_process.py`
- `tests/test_run_log.py`

**Modified: `~/automation-dashboard/`**
- `server.py` — add `REPORTS` config, add `page-goals` to `SCRIPTS` (hidden), add `/reports` route
- `templates/reports.html` — new reports dashboard page
- `templates/index.html` — add Reports link to header, skip hidden scripts

---

## Task 1: Repo Scaffold

**Files:**
- Create: `~/mathnasium-page-goals/` (new git repo)
- Create: `requirements.txt`
- Create: `config.py`
- Create: `run_log.json`
- Create: `.gitignore`

- [ ] **Step 1: Create and init the repo**

```bash
cd ~
mkdir mathnasium-page-goals
cd mathnasium-page-goals
git init
gh repo create mdmathnasiums/mathnasium-page-goals --private --source=. --push
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
google-api-python-client
google-auth
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
import os

RADIUS_LOGIN_URL = "https://radius.mathnasium.com"
# RADIUS_DWP_URL filled in during Task 3 (Radius investigation)
RADIUS_DWP_URL = ""

DRIVE_FOLDER_ID = "1NcVaeoFtyJkJfy6-GtLrlMxoN5cT_gyY"

SUCCESS_RECIPIENTS = [
    "matt.diamond@mathnasium.com",
    "teaneck@mathnasium.com",
    "englewood@mathnasium.com",
]
ERROR_RECIPIENTS = ["matt.diamond@mathnasium.com"]

INPUT_DIR = "input"
OUTPUT_DIR = "output"
RUN_LOG_PATH = "run_log.json"
```

- [ ] **Step 6: Create empty `run_log.json`**

```json
[]
```

- [ ] **Step 7: Create `input/` and `output/` directories**

```bash
mkdir -p input output
```

- [ ] **Step 8: Commit**

```bash
git add .
git commit -m "feat: initial scaffold"
git push -u origin main
```

---

## Task 2: `run_log.py` + Tests

**Files:**
- Create: `run_log.py`
- Create: `tests/test_run_log.py`

- [ ] **Step 1: Write failing tests**

Create `tests/__init__.py` (empty), then `tests/test_run_log.py`:

```python
import json
import os
import pytest
import tempfile
import run_log


@pytest.fixture(autouse=True)
def tmp_log(monkeypatch, tmp_path):
    log_path = str(tmp_path / "run_log.json")
    monkeypatch.setattr(run_log, "RUN_LOG_PATH", log_path)
    return log_path


def test_append_creates_file(tmp_log):
    run_log.append_run("manual", "2026-04", "success", "April Page Goals.xlsx", "https://drive.google.com/x", None)
    assert os.path.exists(tmp_log)


def test_append_stores_fields(tmp_log):
    run_log.append_run("auto", "2026-04", "success", "April Page Goals.xlsx", "https://drive.google.com/x", None)
    with open(tmp_log) as f:
        log = json.load(f)
    assert len(log) == 1
    entry = log[0]
    assert entry["trigger"] == "auto"
    assert entry["month"] == "2026-04"
    assert entry["status"] == "success"
    assert entry["output_file"] == "April Page Goals.xlsx"
    assert entry["drive_link"] == "https://drive.google.com/x"
    assert entry["error"] is None
    assert "timestamp" in entry


def test_append_multiple_runs(tmp_log):
    run_log.append_run("auto", "2026-03", "success", "March Page Goals.xlsx", None, None)
    run_log.append_run("manual", "2026-04", "error", None, None, "Download failed")
    with open(tmp_log) as f:
        log = json.load(f)
    assert len(log) == 2
    assert log[1]["status"] == "error"
    assert log[1]["error"] == "Download failed"


def test_read_log_empty_when_no_file(tmp_log):
    result = run_log.read_log()
    assert result == []


def test_get_last_run_by_trigger(tmp_log):
    run_log.append_run("auto", "2026-03", "success", "March Page Goals.xlsx", None, None)
    run_log.append_run("manual", "2026-04", "success", "April Page Goals.xlsx", None, None)
    last_auto = run_log.get_last_run("auto")
    assert last_auto["month"] == "2026-03"
    last_manual = run_log.get_last_run("manual")
    assert last_manual["month"] == "2026-04"


def test_get_last_run_returns_none_when_empty(tmp_log):
    assert run_log.get_last_run("auto") is None
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
pytest tests/test_run_log.py -v
```

Expected: `ModuleNotFoundError: No module named 'run_log'`

- [ ] **Step 3: Implement `run_log.py`**

```python
import json
import os
from datetime import datetime, timezone

RUN_LOG_PATH = "run_log.json"


def read_log() -> list[dict]:
    if not os.path.exists(RUN_LOG_PATH):
        return []
    with open(RUN_LOG_PATH) as f:
        return json.load(f)


def append_run(
    trigger: str,
    month: str,
    status: str,
    output_file: str = None,
    drive_link: str = None,
    error: str = None,
) -> None:
    log = read_log()
    log.append({
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "trigger": trigger,
        "month": month,
        "status": status,
        "output_file": output_file,
        "drive_link": drive_link,
        "error": error,
    })
    with open(RUN_LOG_PATH, "w") as f:
        json.dump(log, f, indent=2)


def get_last_run(trigger: str = None) -> dict | None:
    log = read_log()
    if trigger:
        log = [r for r in log if r.get("trigger") == trigger]
    return log[-1] if log else None
```

- [ ] **Step 4: Run tests to confirm they pass**

```bash
pytest tests/test_run_log.py -v
```

Expected: 6 passed

- [ ] **Step 5: Commit**

```bash
git add run_log.py tests/
git commit -m "feat: add run_log read/write with tests"
git push
```

---

## Task 3: `process.py` + Tests

**Files:**
- Create: `process.py`
- Create: `tests/test_process.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_process.py`:

```python
import pandas as pd
import pytest
from process import calculate_student_goals


def make_df(rows):
    """rows: list of (student_name, pages, date_str, center)"""
    return pd.DataFrame(rows, columns=["Student Name", "Pages Completed", "Date", "Center"])


# ── Standard student (10+ sessions) ──────────────────────────────────────────

def test_standard_student_uses_last_10():
    rows = [("Alice", i, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 16)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    assert student["num_sessions"] == 10
    assert len([s for s in student["sessions"] if s is not None]) == 10


def test_standard_student_uses_most_recent_10():
    rows = [("Alice", i, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 16)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    # Most recent 10 sessions: pages 6–15, oldest first
    non_null = [s for s in student["sessions"] if s is not None]
    assert non_null[0] == 6   # oldest of the 10
    assert non_null[-1] == 15  # most recent


def test_standard_student_average():
    # 10 sessions with pages 1-10
    rows = [("Bob", i, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 11)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    assert student["average"] == round(sum(range(1, 11)) / 10, 2)  # 5.5


def test_standard_student_goal_is_120_pct_of_average():
    rows = [("Bob", i, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 11)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    expected_goal = round(5.5 * 1.20, 2)  # 6.6
    assert student["goal"] == expected_goal


# ── High performer: goal capped below max ─────────────────────────────────────

def test_goal_capped_below_max():
    # 10 sessions all with value 10 → avg=10, 120%=12, max=10, goal=9.99
    rows = [("Carol", 10, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 11)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    assert student["goal"] == 9.99


def test_goal_capped_when_120pct_exceeds_max():
    # sessions: 5,5,5,5,5,5,5,5,5,20 → avg=7.5, 120%=9.0, max=20, goal=9.0
    pages = [5] * 9 + [20]
    rows = [("Dave", p, f"2026-03-{i:02d}", "Teaneck") for i, p in enumerate(pages, 1)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    assert student["goal"] == 9.0


# ── Zeros excluded ────────────────────────────────────────────────────────────

def test_zeros_excluded():
    rows = [
        ("Eve", 0, "2026-03-01", "Teaneck"),
        ("Eve", 0, "2026-03-02", "Teaneck"),
        ("Eve", 8, "2026-03-03", "Teaneck"),
        ("Eve", 9, "2026-03-04", "Teaneck"),
    ]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    assert student["num_sessions"] == 2
    non_null = [s for s in student["sessions"] if s is not None]
    assert 0 not in non_null


def test_student_with_all_zeros_excluded():
    rows = [("Frank", 0, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 6)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    assert len(results["Teaneck"]) == 0


# ── Null pages excluded ────────────────────────────────────────────────────────

def test_null_pages_excluded():
    rows = [
        ("Grace", None, "2026-03-01", "Teaneck"),
        ("Grace", 7,    "2026-03-02", "Teaneck"),
        ("Grace", 8,    "2026-03-03", "Teaneck"),
    ]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    assert student["num_sessions"] == 2


# ── Partial data (fewer than 10 sessions) ─────────────────────────────────────

def test_partial_student_uses_all_sessions():
    rows = [("Hal", i, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 4)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    assert student["num_sessions"] == 3
    non_null = [s for s in student["sessions"] if s is not None]
    assert len(non_null) == 3


def test_partial_student_sessions_padded_with_none():
    rows = [("Hal", i, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 4)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    student = results["Teaneck"][0]
    assert len(student["sessions"]) == 10
    assert student["sessions"].count(None) == 7


# ── Center assignment ─────────────────────────────────────────────────────────

def test_englewood_center_assignment():
    rows = [("Ivy", 8, f"2026-03-{i:02d}", "Englewood") for i in range(1, 5)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    assert len(results["Englewood"]) == 1
    assert len(results["Teaneck"]) == 0


def test_teaneck_virtual_assigned_to_teaneck():
    rows = [("Jay", 8, f"2026-03-{i:02d}", "Teaneck, Teaneck Virtual") for i in range(1, 5)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    assert len(results["Teaneck"]) == 1


def test_englewood_virtual_assigned_to_englewood():
    rows = [("Kim", 8, f"2026-03-{i:02d}", "Englewood, Teaneck Virtual") for i in range(1, 5)]
    df = make_df(rows)
    results = calculate_student_goals(df)
    assert len(results["Englewood"]) == 1


# ── Alphabetical sort ─────────────────────────────────────────────────────────

def test_students_sorted_alphabetically():
    rows = (
        [("Zara", 8, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 5)] +
        [("Aaron", 8, f"2026-03-{i:02d}", "Teaneck") for i in range(1, 5)]
    )
    df = make_df(rows)
    results = calculate_student_goals(df)
    names = [s["name"] for s in results["Teaneck"]]
    assert names == sorted(names)
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
pytest tests/test_process.py -v
```

Expected: `ModuleNotFoundError: No module named 'process'`

- [ ] **Step 3: Implement `process.py`**

```python
import os
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


def calculate_student_goals(df: pd.DataFrame) -> dict[str, list[dict]]:
    """
    Given the full DWP DataFrame, return a dict:
      {"Englewood": [student_record, ...], "Teaneck": [student_record, ...]}

    Each student_record:
      {
        "name": str,
        "num_sessions": int,
        "sessions": list[float | None],  # always length 10, oldest first
        "average": float,
        "goal": float,
      }
    """
    df = df[["Student Name", "Pages Completed", "Date", "Center"]].copy()
    df["Date"] = pd.to_datetime(df["Date"])

    # Filter: Pages Completed must be numeric and > 0
    df = df[pd.to_numeric(df["Pages Completed"], errors="coerce").notna()]
    df["Pages Completed"] = df["Pages Completed"].astype(float)
    df = df[df["Pages Completed"] > 0]

    # Assign center
    def assign_center(center_str: str) -> str | None:
        if "Englewood" in str(center_str):
            return "Englewood"
        if "Teaneck" in str(center_str):
            return "Teaneck"
        return None

    df["AssignedCenter"] = df["Center"].apply(assign_center)
    df = df[df["AssignedCenter"].notna()]

    results = {"Englewood": [], "Teaneck": []}

    for (name, center), group in df.groupby(["Student Name", "AssignedCenter"]):
        # Sort by date descending, take most recent 10
        group = group.sort_values("Date", ascending=False).head(10)
        pages = group["Pages Completed"].tolist()
        num_sessions = len(pages)

        # Reverse so oldest is first (Session 1 = oldest)
        pages_oldest_first = list(reversed(pages))

        average = round(sum(pages_oldest_first) / num_sessions, 2)
        max_pages = max(pages_oldest_first)
        goal = round(min(average * 1.20, max_pages - 0.01), 2)

        # Pad to 10 with None (left-pad: Nones go in early session slots)
        padded = [None] * (10 - num_sessions) + pages_oldest_first

        results[center].append({
            "name": name,
            "num_sessions": num_sessions,
            "sessions": padded,
            "average": average,
            "goal": goal,
        })

    # Sort alphabetically
    for center in results:
        results[center].sort(key=lambda s: s["name"])

    return results


def write_excel(results: dict[str, list[dict]], output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    header_fill   = PatternFill("solid", fgColor="4472C4")
    session_fill  = PatternFill("solid", fgColor="E7E6E6")
    summary_fill  = PatternFill("solid", fgColor="FFF2CC")
    header_font   = Font(bold=True, color="FFFFFF")
    center_align  = Alignment(horizontal="center")
    left_align    = Alignment(horizontal="left")

    headers = (
        ["Student Name", "# Sessions"] +
        [f"Session {i}" for i in range(1, 11)] +
        ["Average Pages", "Page Goal"]
    )
    col_widths = [30, 18] + [10] * 10 + [14, 12]

    for sheet_name in ["Englewood", "Teaneck"]:
        ws = wb.create_sheet(sheet_name)
        students = results[sheet_name]

        # Header row
        for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # Data rows
        for row_idx, student in enumerate(students, start=2):
            ws.cell(row=row_idx, column=1, value=student["name"]).alignment = left_align
            ws.cell(row=row_idx, column=2, value=student["num_sessions"]).alignment = center_align

            for s_idx, pages in enumerate(student["sessions"]):
                cell = ws.cell(row=row_idx, column=3 + s_idx, value=pages)
                cell.fill = session_fill
                cell.alignment = center_align

            avg_cell = ws.cell(row=row_idx, column=13, value=student["average"])
            avg_cell.fill = summary_fill
            avg_cell.alignment = center_align

            goal_cell = ws.cell(row=row_idx, column=14, value=student["goal"])
            goal_cell.fill = summary_fill
            goal_cell.alignment = center_align

        ws.freeze_panes = "B2"

    wb.save(output_path)


def process_report(input_path: str, output_path: str) -> None:
    df = pd.read_excel(input_path)
    results = calculate_student_goals(df)
    write_excel(results, output_path)
```

- [ ] **Step 4: Run tests to confirm they pass**

```bash
pytest tests/test_process.py -v
```

Expected: all tests pass

- [ ] **Step 5: Commit**

```bash
git add process.py tests/test_process.py
git commit -m "feat: add process.py with calculation logic and Excel output"
git push
```

---

## Task 4: Investigate Radius Digital Workout Plan Selectors

> **⚠️ This task requires a live browser session with the user. Do not attempt to implement `download.py` without completing this task first.**

**Files:**
- Modify: `config.py` (fill in `RADIUS_DWP_URL`)

- [ ] **Step 1: Navigate to Radius in the browser**

Open `https://radius.mathnasium.com` and log in with your normal credentials.

- [ ] **Step 2: Find the Digital Workout Plan report**

Navigate to the reports section and open the Digital Workout Plan report. Note the full URL from the browser address bar — update `config.py`:

```python
RADIUS_DWP_URL = "https://radius.mathnasium.com/..."  # fill in actual URL
```

- [ ] **Step 3: Identify filter elements with DevTools**

Open DevTools (Cmd+Option+I → Elements tab). Inspect:
- **Center selector** — note the element ID or selector (likely a Kendo MultiSelect like `#AllCenterListMultiSelect`)
- **Date range inputs** — note the IDs for start date and end date fields
- **Export/Download button** — note its ID or selector

Document findings here:
```
Center selector:    #____________
Start date input:   #____________
End date input:     #____________
Export button:      #____________
```

- [ ] **Step 4: Confirm Kendo widget pattern**

Check if the date inputs use Kendo widgets (like the Guardian report). In the DevTools Console, run:
```javascript
jQuery('#<date_field_id>').data('kendoDatePicker')
```
If it returns an object, it's Kendo — note this for the download script.

- [ ] **Step 5: Commit config update**

```bash
git add config.py
git commit -m "config: add Radius Digital Workout Plan URL"
git push
```

---

## Task 5: `download.py`

> **Prerequisite: Task 4 must be complete.** Fill in all selectors before implementing.

**Files:**
- Create: `download.py`

- [ ] **Step 1: Implement `download.py`**

Replace `#center-selector`, `#start-date`, `#end-date`, `#export-btn` with the actual IDs found in Task 4.

```python
"""
download.py
Logs into Radius and downloads the Digital Workout Plan Report
for the specified month. Returns the local file path.
"""

import os
from datetime import date
from dateutil.relativedelta import relativedelta
from playwright.sync_api import sync_playwright

from config import RADIUS_LOGIN_URL, RADIUS_DWP_URL, INPUT_DIR

RADIUS_USERNAME = os.environ["RADIUS_USERNAME"]
RADIUS_PASSWORD = os.environ["RADIUS_PASSWORD"]

# Center IDs: Teaneck + Englewood
CENTER_VALUES = ["2871", "2428"]


def download_dwp_report(data_month: date) -> str:
    """
    Download the Digital Workout Plan report for the given month.
    data_month: first day of the month to download.
    Returns the local path to the downloaded file.
    """
    os.makedirs(INPUT_DIR, exist_ok=True)
    output_path = os.path.join(
        INPUT_DIR,
        f"Digital_Workout_Plan_{data_month.strftime('%Y_%m')}.xlsx"
    )

    start = data_month.replace(day=1)
    # Last day of month
    end = (data_month + relativedelta(months=1)) - relativedelta(days=1)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        print(f"[download] Logging into Radius...")
        page.goto(RADIUS_LOGIN_URL)
        page.wait_for_load_state("networkidle")
        page.fill("#UserName", RADIUS_USERNAME)
        page.fill("#Password", RADIUS_PASSWORD)
        page.click("#login")
        page.wait_for_load_state("networkidle")
        print("[download] Logged in.")

        print(f"[download] Navigating to Digital Workout Plan report...")
        page.goto(RADIUS_DWP_URL)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)

        # Select both centers (adjust selector based on Task 4 findings)
        page.evaluate(f"""
            var w = jQuery('#AllCenterListMultiSelect').data('kendoMultiSelect');
            w.value({CENTER_VALUES});
            w.trigger('change');
        """)
        page.wait_for_timeout(500)

        # Set date range (adjust selectors based on Task 4 findings)
        # If Kendo DatePicker:
        page.evaluate(f"""
            var startPicker = jQuery('#StartDate').data('kendoDatePicker');
            startPicker.value('{start.strftime("%m/%d/%Y")}');
            startPicker.trigger('change');
            var endPicker = jQuery('#EndDate').data('kendoDatePicker');
            endPicker.value('{end.strftime("%m/%d/%Y")}');
            endPicker.trigger('change');
        """)
        page.wait_for_timeout(500)

        # Search
        page.click("#btnsearch")
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)

        # Download
        print("[download] Downloading report...")
        with page.expect_download() as dl:
            page.click("#btnExport")
        dl.value.save_as(output_path)
        print(f"[download] Saved to {output_path}")

        browser.close()

    return output_path
```

- [ ] **Step 2: Smoke test manually**

```bash
source .venv/bin/activate
RADIUS_USERNAME=your_username RADIUS_PASSWORD=your_password python3 -c "
from datetime import date
from download import download_dwp_report
path = download_dwp_report(date(2026, 3, 1))
print('Downloaded to:', path)
"
```

Expected: file appears in `input/Digital_Workout_Plan_2026_03.xlsx`

- [ ] **Step 3: Open the file and verify columns**

```bash
python3 -c "
import pandas as pd
df = pd.read_excel('input/Digital_Workout_Plan_2026_03.xlsx')
print(df.columns.tolist())
print(df.head(3))
"
```

Confirm the columns `Date`, `Student Name`, `Pages Completed`, `Center` are present. If column names differ, update `process.py` to match.

- [ ] **Step 4: Commit**

```bash
git add download.py
git commit -m "feat: add Radius Digital Workout Plan downloader"
git push
```

---

## Task 6: `deliver.py`

**Files:**
- Create: `deliver.py`

- [ ] **Step 1: Implement `deliver.py`**

```python
"""
deliver.py
Email the output file and upload it to Google Drive.
"""

import json
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

from config import DRIVE_FOLDER_ID

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465


def send_email(
    filepath: str,
    month_name: str,
    recipients: list[str],
    smtp_user: str,
    smtp_password: str,
    error_message: str = None,
) -> None:
    """
    Send the output file as an email attachment.
    If error_message is provided, send an error notification instead.
    """
    filename = os.path.basename(filepath) if filepath else None

    if error_message:
        subject = f"⚠️ Page Goals Automation Failed — {month_name}"
        body = f"The {month_name} page goals automation failed with the following error:\n\n{error_message}"
        msg = MIMEText(body, "plain")
        msg["Subject"] = subject
        msg["From"] = smtp_user
        msg["To"] = ", ".join(recipients)
    else:
        subject = f"{month_name} Page Goals"
        body = f"Please find attached the {month_name} page goals report."
        msg = MIMEMultipart()
        msg["Subject"] = subject
        msg["From"] = smtp_user
        msg["To"] = ", ".join(recipients)
        msg.attach(MIMEText(body, "plain"))

        with open(filepath, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

    print(f"[deliver] Sending email to: {', '.join(recipients)}")
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, recipients, msg.as_string())
    print("[deliver] Email sent.")


def upload_to_drive(filepath: str, credentials_json: str) -> str:
    """
    Upload filepath to the configured Google Drive folder.
    credentials_json: the service account JSON as a string.
    Returns the webViewLink of the uploaded file.
    """
    creds_dict = json.loads(credentials_json)
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive.file"],
    )
    service = build("drive", "v3", credentials=creds)

    filename = os.path.basename(filepath)
    file_metadata = {"name": filename, "parents": [DRIVE_FOLDER_ID]}
    media = MediaFileUpload(
        filepath,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    print(f"[deliver] Uploading {filename} to Google Drive...")
    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id, webViewLink",
    ).execute()

    link = uploaded.get("webViewLink", "")
    print(f"[deliver] Uploaded. Link: {link}")
    return link
```

- [ ] **Step 2: Commit**

```bash
git add deliver.py
git commit -m "feat: add email and Google Drive delivery"
git push
```

---

## Task 7: `main.py` + End-to-End Test

**Files:**
- Create: `main.py`

- [ ] **Step 1: Implement `main.py`**

```python
"""
main.py
Orchestrates: download → process → deliver → log.
Usage:
    python main.py [--month YYYY-MM] [--trigger auto|manual]
"""

import argparse
import os
import sys
from datetime import datetime
from dateutil.relativedelta import relativedelta

from config import SUCCESS_RECIPIENTS, ERROR_RECIPIENTS, OUTPUT_DIR
from download import download_dwp_report
from process import process_report
from deliver import send_email, upload_to_drive
import run_log

SMTP_USER = os.environ["SMTP_USER"]
SMTP_PASSWORD = os.environ["SMTP_PASSWORD"]
GOOGLE_DRIVE_CREDENTIALS = os.environ.get("GOOGLE_DRIVE_CREDENTIALS", "")


def parse_month(month_str: str | None):
    if month_str:
        return datetime.strptime(month_str, "%Y-%m")
    # Default: previous month
    return datetime.now() - relativedelta(months=1)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", help="Month to process (YYYY-MM), defaults to previous month")
    parser.add_argument("--trigger", default="manual", choices=["auto", "manual"])
    args = parser.parse_args()

    data_month = parse_month(args.month)
    # Goals apply to the current month (month after data_month)
    goal_month = data_month + relativedelta(months=1)
    month_name = goal_month.strftime("%B")  # e.g. "May"
    month_key = data_month.strftime("%Y-%m")  # e.g. "2026-04"

    output_filename = f"{month_name} Page Goals.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    print(f"[main] Processing {month_name} page goals (data from {data_month.strftime('%B %Y')})")

    try:
        # 1. Download
        input_path = download_dwp_report(data_month)

        # 2. Process
        process_report(input_path, output_path)
        print(f"[main] Report written to {output_path}")

        # 3. Upload to Drive
        drive_link = ""
        if GOOGLE_DRIVE_CREDENTIALS:
            try:
                drive_link = upload_to_drive(output_path, GOOGLE_DRIVE_CREDENTIALS)
            except Exception as e:
                print(f"[main] Drive upload failed (continuing): {e}")

        # 4. Email success
        send_email(output_path, month_name, SUCCESS_RECIPIENTS, SMTP_USER, SMTP_PASSWORD)

        # 5. Log success
        run_log.append_run(
            trigger=args.trigger,
            month=month_key,
            status="success",
            output_file=output_filename,
            drive_link=drive_link or None,
        )
        print(f"[main] ✅ Done. {month_name} page goals complete.")

    except Exception as e:
        error_msg = str(e)
        print(f"[main] ❌ Error: {error_msg}", file=sys.stderr)

        # Email error to Matt only
        try:
            send_email(None, month_name, ERROR_RECIPIENTS, SMTP_USER, SMTP_PASSWORD, error_message=error_msg)
        except Exception as email_err:
            print(f"[main] Failed to send error email: {email_err}", file=sys.stderr)

        # Log failure
        run_log.append_run(
            trigger=args.trigger,
            month=month_key,
            status="error",
            error=error_msg,
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Test dry run with a real downloaded file**

Assuming Task 5 smoke test left a file in `input/`:

```bash
source .venv/bin/activate
SMTP_USER=x SMTP_PASSWORD=x python3 -c "
from process import process_report
process_report('input/Digital_Workout_Plan_2026_03.xlsx', 'output/April Page Goals.xlsx')
print('Done — check output/April Page Goals.xlsx')
"
```

Open the output file and verify: two sheets (Englewood, Teaneck), correct columns, blue headers, gray session columns, yellow average/goal columns.

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: add main.py orchestration"
git push
```

---

## Task 8: GitHub Actions Workflow

**Files:**
- Create: `.github/workflows/monthly_page_goals.yml`

- [ ] **Step 1: Add GitHub Actions secrets**

In `https://github.com/mdmathnasiums/mathnasium-page-goals/settings/secrets/actions`, add:

| Secret | Value |
|---|---|
| `RADIUS_USERNAME` | Copy from radius-morning-briefing repo |
| `RADIUS_PASSWORD` | Copy from radius-morning-briefing repo |
| `SMTP_USER` | Copy from radius-morning-briefing repo |
| `SMTP_PASSWORD` | Copy from radius-morning-briefing repo |
| `GOOGLE_DRIVE_CREDENTIALS` | Paste the full service account JSON (see setup steps below) |

**Google service account setup (one-time):**
1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a project (or use existing), enable the **Google Drive API**
3. Create a **Service Account**, download the JSON key
4. Copy the `client_email` from the JSON and share the Google Drive folder with that email address
5. Paste the full JSON content as the `GOOGLE_DRIVE_CREDENTIALS` secret

- [ ] **Step 2: Create `.github/workflows/monthly_page_goals.yml`**

```yaml
name: Monthly Page Goals

on:
  schedule:
    - cron: "0 15 1 * *"   # 1st of every month at 15:00 UTC (11:00 AM ET)
  workflow_dispatch:         # allow manual trigger from GitHub UI

permissions:
  contents: write            # needed to commit run_log.json back

jobs:
  generate-page-goals:
    runs-on: ubuntu-latest
    timeout-minutes: 20

    steps:
      - name: Check out repository
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Install Playwright browser
        run: playwright install chromium --with-deps

      - name: Run page goals
        env:
          RADIUS_USERNAME: ${{ secrets.RADIUS_USERNAME }}
          RADIUS_PASSWORD: ${{ secrets.RADIUS_PASSWORD }}
          SMTP_USER: ${{ secrets.SMTP_USER }}
          SMTP_PASSWORD: ${{ secrets.SMTP_PASSWORD }}
          GOOGLE_DRIVE_CREDENTIALS: ${{ secrets.GOOGLE_DRIVE_CREDENTIALS }}
        run: python main.py --trigger auto

      - name: Commit run log
        if: always()
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add run_log.json
          git diff --staged --quiet || git commit -m "chore: update run log [skip ci]"
          git push
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

- [ ] **Step 3: Push and verify workflow appears**

```bash
git add .github/
git commit -m "feat: add GitHub Actions monthly cron workflow"
git push
```

Go to `https://github.com/mdmathnasiums/mathnasium-page-goals/actions` and confirm the workflow appears.

- [ ] **Step 4: Trigger a manual test run**

Click "Run workflow" → "Run workflow" in the GitHub Actions UI. Watch the logs. Confirm:
- Report downloads
- Excel is generated
- Email arrives at matt.diamond@mathnasium.com
- File appears in Google Drive folder
- `run_log.json` is committed back with a success entry

---

## Task 9: Reports Dashboard — `/reports` Page

**Files:**
- Modify: `~/automation-dashboard/server.py`
- Create: `~/automation-dashboard/templates/reports.html`
- Modify: `~/automation-dashboard/templates/index.html`

- [ ] **Step 1: Add REPORTS config and page-goals script to `server.py`**

In `server.py`, add after the existing `SCRIPTS` dict:

```python
# ── Hidden scripts (not shown on main page, used by /reports) ─────────────────
SCRIPTS["page-goals"] = {
    "name": "Student Page Goals",
    "description": "Calculates monthly page goals for each student from Radius data.",
    "command": [PYTHON, "main.py", "--trigger", "manual"],
    "cwd": "/Users/mattdiamond/mathnasium-page-goals",
    "icon": "📊",
    "category": "Mathnasium",
    "hidden": True,
}

# ── Reports registry (shown on /reports page) ─────────────────────────────────
REPORTS = [
    {
        "id": "student-page-goals",
        "name": "Student Page Goals",
        "schedule": "1st of month",
        "script_id": "page-goals",
        "run_log_path": "/Users/mattdiamond/mathnasium-page-goals/run_log.json",
    },
]
```

- [ ] **Step 2: Add `/reports` route to `server.py`**

Add after the existing routes:

```python
import json as _json
from datetime import datetime as _dt

def _load_run_log(path: str) -> list[dict]:
    try:
        with open(path) as f:
            return _json.load(f)
    except (FileNotFoundError, _json.JSONDecodeError):
        return []


def _fmt_run(entry: dict | None) -> dict | None:
    if entry is None:
        return None
    ts = entry.get("timestamp", "")
    try:
        dt = _dt.strptime(ts, "%Y-%m-%dT%H:%M:%SZ")
        friendly = dt.strftime("%-m/%-d/%Y %-I:%M %p") + " UTC"
    except ValueError:
        friendly = ts
    return {**entry, "friendly_time": friendly}


@app.route("/reports")
def reports_page():
    report_rows = []
    for report in REPORTS:
        log = _load_run_log(report["run_log_path"])
        auto_runs   = [r for r in log if r.get("trigger") == "auto"]
        manual_runs = [r for r in log if r.get("trigger") == "manual"]
        report_rows.append({
            **report,
            "last_auto":   _fmt_run(auto_runs[-1]   if auto_runs   else None),
            "last_manual": _fmt_run(manual_runs[-1] if manual_runs else None),
        })
    return render_template("reports.html", reports=report_rows)
```

- [ ] **Step 3: Create `templates/reports.html`**

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Reports — Mathnasium Automations</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background: #f0f2f5; color: #1a1a1a; min-height: 100vh; }

    header { background: #1e3a5f; color: white; padding: 24px 40px; display: flex; align-items: center; gap: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.2); }
    header h1 { font-size: 1.6rem; font-weight: 700; }
    header p  { font-size: 0.9rem; opacity: 0.75; margin-top: 2px; }
    .logo { font-size: 2rem; }
    .nav-link { margin-left: auto; color: rgba(255,255,255,0.8); text-decoration: none; font-size: 0.9rem; }
    .nav-link:hover { color: white; }

    main { max-width: 1100px; margin: 0 auto; padding: 36px 24px; }

    h2 { font-size: 0.75rem; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: #6b7280; margin: 0 0 16px; }

    table { width: 100%; border-collapse: collapse; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
    th { background: #1e3a5f; color: white; font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; padding: 12px 16px; text-align: left; }
    td { padding: 14px 16px; border-bottom: 1px solid #f0f0f0; font-size: 0.875rem; vertical-align: middle; }
    tr:last-child td { border-bottom: none; }

    .status-ok  { color: #16a34a; font-weight: 600; }
    .status-err { color: #dc2626; font-weight: 600; }
    .status-none { color: #9ca3af; font-style: italic; }

    .btn { display: inline-block; padding: 7px 16px; border-radius: 8px; font-size: 0.8rem; font-weight: 600; cursor: pointer; border: none; text-decoration: none; transition: opacity 0.15s; }
    .btn:hover { opacity: 0.85; }
    .btn-primary { background: #1e3a5f; color: white; }

    .delivery-link { color: #1e3a5f; text-decoration: none; font-weight: 500; margin-right: 8px; }
    .delivery-link:hover { text-decoration: underline; }

    .running-indicator { display: none; align-items: center; font-size: 0.8rem; color: #60a5fa; margin-left: 8px; }
    .running-indicator.active { display: inline-flex; }
    .spinner { display: inline-block; width: 10px; height: 10px; border: 2px solid #64748b; border-top-color: #60a5fa; border-radius: 50%; animation: spin 0.8s linear infinite; margin-right: 6px; }
    @keyframes spin { to { transform: rotate(360deg); } }

    #output-panel { display: none; position: fixed; bottom: 0; left: 0; right: 0; height: 320px; background: #0f172a; color: #e2e8f0; font-family: "SF Mono", "Fira Code", monospace; font-size: 0.8rem; z-index: 100; flex-direction: column; box-shadow: 0 -4px 24px rgba(0,0,0,0.4); }
    #output-panel.active { display: flex; }
    #output-header { display: flex; align-items: center; justify-content: space-between; padding: 10px 20px; background: #1e293b; border-bottom: 1px solid #334155; flex-shrink: 0; }
    #output-title  { font-weight: 700; color: #94a3b8; font-size: 0.8rem; letter-spacing: 0.05em; }
    #output-status { font-size: 0.75rem; }
    #output-close  { background: none; border: none; color: #64748b; font-size: 1.2rem; cursor: pointer; padding: 2px 6px; border-radius: 4px; }
    #output-close:hover { background: #334155; color: white; }
    #output-body   { flex: 1; overflow-y: auto; padding: 16px 20px; line-height: 1.6; }
    .output-line { white-space: pre-wrap; word-break: break-all; }
    .output-line.done  { color: #60a5fa; font-weight: bold; margin-top: 8px; }
    .output-line.error { color: #f87171; }
  </style>
</head>
<body>

<header>
  <div class="logo">📊</div>
  <div>
    <h1>Reports</h1>
    <p>Scheduled automations — run history and manual triggers</p>
  </div>
  <a class="nav-link" href="/">← Back to Dashboard</a>
</header>

<main>
  <h2>Scheduled Reports</h2>
  <table>
    <thead>
      <tr>
        <th>Report</th>
        <th>Schedule</th>
        <th>Last Auto Run</th>
        <th>Last Manual Run</th>
        <th>Delivery</th>
        <th></th>
      </tr>
    </thead>
    <tbody>
      {% for r in reports %}
      <tr>
        <td><strong>{{ r.name }}</strong></td>
        <td>{{ r.schedule }}</td>

        <td>
          {% if r.last_auto %}
            <span class="{{ 'status-ok' if r.last_auto.status == 'success' else 'status-err' }}">
              {{ '✓' if r.last_auto.status == 'success' else '✗' }}
            </span>
            {{ r.last_auto.friendly_time }}
          {% else %}
            <span class="status-none">Never</span>
          {% endif %}
        </td>

        <td>
          {% if r.last_manual %}
            <span class="{{ 'status-ok' if r.last_manual.status == 'success' else 'status-err' }}">
              {{ '✓' if r.last_manual.status == 'success' else '✗' }}
            </span>
            {{ r.last_manual.friendly_time }}
          {% else %}
            <span class="status-none">Never</span>
          {% endif %}
        </td>

        <td>
          {% set last = r.last_auto or r.last_manual %}
          {% if last and last.drive_link %}
            <a class="delivery-link" href="{{ last.drive_link }}" target="_blank">📁 Drive</a>
          {% endif %}
          {% if last and last.status == 'success' %}
            <span style="color:#6b7280;font-size:0.8rem;">📧 Emailed</span>
          {% endif %}
          {% if not last %}
            <span class="status-none">—</span>
          {% endif %}
        </td>

        <td>
          <button class="btn btn-primary" id="run-btn-{{ r.script_id }}"
                  onclick="runScript('{{ r.script_id }}', '{{ r.name }}')">▶ Run Now</button>
          <span class="running-indicator" id="running-{{ r.script_id }}">
            <span class="spinner"></span> Running…
          </span>
        </td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
</main>

<div id="output-panel">
  <div id="output-header">
    <span id="output-title">OUTPUT</span>
    <span id="output-status"></span>
    <button id="output-close" onclick="closeOutput()">✕</button>
  </div>
  <div id="output-body"></div>
</div>

<script>
  let currentSource = null;
  let currentScriptId = null;

  function runScript(id, name) {
    if (currentSource) currentSource.close();
    currentScriptId = id;
    const panel    = document.getElementById("output-panel");
    const body     = document.getElementById("output-body");
    const title    = document.getElementById("output-title");
    const status   = document.getElementById("output-status");
    const runBtn   = document.getElementById(`run-btn-${id}`);
    const indicator = document.getElementById(`running-${id}`);

    body.innerHTML = "";
    title.textContent = name.toUpperCase();
    status.innerHTML = '<span style="color:#60a5fa">● Running…</span>';
    panel.classList.add("active");
    if (runBtn) runBtn.disabled = true;
    if (indicator) indicator.classList.add("active");

    currentSource = new EventSource(`/run/${id}`);

    currentSource.onmessage = function(e) {
      const text = JSON.parse(e.data);
      if (text === "__DONE__") {
        status.innerHTML = '<span style="color:#4ade80">✓ Done</span>';
        addLine("✓ Completed successfully.", "done");
        finish();
        setTimeout(() => location.reload(), 1500);
        return;
      }
      if (text.startsWith("__ERROR__")) {
        status.innerHTML = '<span style="color:#f87171">✗ Failed</span>';
        addLine("✗ " + text, "error");
        finish();
        return;
      }
      addLine(text, "");
    };

    currentSource.onerror = function() {
      status.innerHTML = '<span style="color:#f87171">✗ Connection lost</span>';
      finish();
    };

    function addLine(text, cls) {
      const line = document.createElement("div");
      line.className = "output-line " + cls;
      line.textContent = text;
      body.appendChild(line);
      body.scrollTop = body.scrollHeight;
    }

    function finish() {
      currentSource.close();
      currentSource = null;
      if (runBtn) runBtn.disabled = false;
      if (indicator) indicator.classList.remove("active");
    }
  }

  function closeOutput() {
    if (currentSource) { currentSource.close(); currentSource = null; }
    document.getElementById("output-panel").classList.remove("active");
    if (currentScriptId) {
      const indicator = document.getElementById(`running-${currentScriptId}`);
      if (indicator) indicator.classList.remove("active");
      currentScriptId = null;
    }
  }
</script>

</body>
</html>
```

- [ ] **Step 4: Add Reports link to `index.html` header and skip hidden scripts**

In `templates/index.html`, update the `<header>`:

```html
<header>
  <div class="logo">🤖</div>
  <div>
    <h1>Mathnasium Automations</h1>
    <p>Click a button to run a script or open a web app</p>
  </div>
  <a href="/reports" style="margin-left:auto;color:rgba(255,255,255,0.8);text-decoration:none;font-size:0.9rem;">📊 Reports</a>
</header>
```

Update both card loops to skip hidden scripts:

```html
{% for id, s in scripts.items() if s.category == "Mathnasium" and not s.hidden %}
```

```html
{% for id, s in scripts.items() if s.category == "Fantasy Baseball" and not s.hidden %}
```

- [ ] **Step 5: Restart the dashboard and verify**

```bash
launchctl stop com.mathnasium.dashboard
launchctl start com.mathnasium.dashboard
```

Open `http://localhost:8080/` — confirm no change to main page cards.  
Open `http://localhost:8080/reports` — confirm the Student Page Goals row appears with "Never" for both run columns and a "▶ Run Now" button.

- [ ] **Step 6: Commit dashboard changes**

```bash
cd ~/automation-dashboard
git add server.py templates/reports.html templates/index.html
git commit -m "feat: add /reports dashboard page with run history and Run Now button"
git push
```

---

## Self-Review

**Spec coverage check:**
- ✅ Radius download via Playwright — Task 5
- ✅ Calculation logic (filter, last 10, average, 20% goal, cap at max-0.01) — Task 3 (process.py)
- ✅ Two worksheets (Englewood, Teaneck), 14 columns, formatting — Task 3 (write_excel)
- ✅ Filename `[Month] Page Goals.xlsx` — Task 7 (main.py)
- ✅ Email success to all three, error to Matt only — Task 7 (main.py) + Task 6 (deliver.py)
- ✅ Google Drive upload — Task 6 (deliver.py)
- ✅ run_log.json with trigger, status, drive_link — Tasks 2 + 7
- ✅ GitHub Actions cron 1st of month at 15:00 UTC — Task 8
- ✅ `--trigger auto|manual` flag — Task 7
- ✅ `--month` flag for reprocessing past months — Task 7
- ✅ Reports dashboard with run history — Task 9
- ✅ Run Now button streams live output — Task 9 (reports.html JS)
- ✅ Page auto-reloads after successful run to show updated history — Task 9 (reports.html JS)
- ✅ Drive upload failure doesn't block email — Task 7 (main.py try/except)
