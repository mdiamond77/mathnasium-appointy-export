# New Enrollments Report Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a CLI tool that produces a historical new-enrollment Excel report (Jan 2023–present) for Teaneck and Englewood, showing student name, program type, and amount charged per enrollment month.

**Architecture:** Single pull of the Enrollment Report (9/1/2022–today) identifies new enrollments by detecting students whose enrollment start date falls in the report window with no prior enrollment ending within 90 days. Payment Reconciliation is pulled year-by-year and joined on student name + center to provide amount charged. Output is two Excel files (one per center), each with a detail tab and summary tab.

**Tech Stack:** Python 3.11+, Playwright (headless Chromium), pandas, openpyxl, pytest

---

## File Map

| File | Responsibility |
|---|---|
| `config.py` | All constants: URLs, selectors, center config, column name constants |
| `discover_selectors.py` | One-time headed browser script to confirm Enrollment Report grid fields |
| `download.py` | Playwright: login, pull enrollment report (single grid read), pull payment recon by year (Excel export) |
| `transform.py` | pandas: detect new enrollments, join payment amounts, build summary |
| `output.py` | Write two-tab Excel per center using openpyxl |
| `main.py` | CLI orchestrator: parse args, call download → transform → output |
| `tests/test_transform.py` | Unit tests for all transform logic |
| `requirements.txt` | Dependencies |
| `.gitignore` | Ignore `input/`, `output/`, `__pycache__` |

---

## Task 1: Project scaffold

**Files:**
- Create: `~/mathnasium-new-enrollments/config.py`
- Create: `~/mathnasium-new-enrollments/requirements.txt`
- Create: `~/mathnasium-new-enrollments/.gitignore`
- Create: `~/mathnasium-new-enrollments/tests/__init__.py`

- [ ] **Step 1: Create project directory and git repo**

```bash
mkdir -p ~/mathnasium-new-enrollments/tests
mkdir -p ~/mathnasium-new-enrollments/input
mkdir -p ~/mathnasium-new-enrollments/output
cd ~/mathnasium-new-enrollments
git init
```

- [ ] **Step 2: Create requirements.txt**

```
playwright==1.44.0
pandas==2.2.2
openpyxl==3.1.2
pytest==8.2.0
```

- [ ] **Step 3: Create .gitignore**

```
input/
output/
__pycache__/
*.pyc
.env
```

- [ ] **Step 4: Create config.py with placeholder column name constants**

Column name constants marked `TBD_` will be filled in after running discover_selectors.py (Task 2) and inspecting a payment recon year export (Task 6).

```python
from pathlib import Path
from datetime import date

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

RADIUS_LOGIN_URL = "https://radius.mathnasium.com"
ENROLLMENT_REPORT_URL = "https://radius.mathnasium.com/Enrollment/EnrollmentReport"
PAYMENT_RECON_URL = "https://radius.mathnasium.com/Payment"

CENTERS = {
    "Teaneck":   {"radius_id": "2871"},
    "Englewood": {"radius_id": "2428"},
}

# Enrollment report date range — covers all students plus 3-month lookback for Jan 2023
ENROLL_PULL_START = "9/1/2022"

# Report output window
REPORT_START = date(2023, 1, 1)

# Payment recon years to pull
PAYMENT_YEARS = [2022, 2023, 2024, 2025, 2026]

# ── Enrollment Report column names (fill in after discover_selectors.py) ────
# Run discover_selectors.py, download a sample export, open the Excel,
# and replace the strings below with the exact column headers.
COL_ENROLL_FIRST_NAME  = "TBD_FirstName"
COL_ENROLL_LAST_NAME   = "TBD_LastName"
COL_ENROLL_CENTER      = "TBD_Center"
COL_ENROLL_START_DATE  = "TBD_StartDate"
COL_ENROLL_END_DATE    = "TBD_EndDate"   # set to None if not in export
COL_ENROLL_PROGRAM     = "TBD_Program"

# ── Payment Recon column names (fill in after inspecting year export) ────────
COL_PAY_STUDENT        = "TBD_StudentName"   # student name field in payment recon
COL_PAY_CENTER         = "TBD_PayCenter"
COL_PAY_AMOUNT         = "TBD_Amount"        # e.g. "Amount Paid"
COL_PAY_DATE           = "TBD_Date"          # date field for filtering by month; None if not present

# Re-enrollment gap threshold in days
REENROLL_GAP_DAYS = 90
```

- [ ] **Step 5: Create tests/__init__.py**

```python
```

- [ ] **Step 6: Initial commit**

```bash
cd ~/mathnasium-new-enrollments
git add config.py requirements.txt .gitignore tests/__init__.py
git commit -m "chore: project scaffold"
```

---

## Task 2: Discover Enrollment Report selectors and column names

**Files:**
- Create: `~/mathnasium-new-enrollments/discover_selectors.py`
- Modify: `~/mathnasium-new-enrollments/config.py` (fill in column name constants)

- [ ] **Step 1: Create discover_selectors.py**

```python
"""
Headed browser to inspect the Radius Enrollment Report.

Run with:
    RADIUS_USERNAME=you@email.com RADIUS_PASSWORD=pass python discover_selectors.py

What to record:
  - Date range picker selectors (#StartDate, #EndDate, or similar)
  - Center multi-select selector
  - Search button selector (#btnsearch or similar)
  - Export button selector (#btnExport or similar)
  - Exact column headers from the downloaded Excel:
      * Student first name
      * Student last name (or combined name)
      * Center
      * Enrollment start date
      * Enrollment end date (if present)
      * Program type (2x/week, 3x/week, etc.)
"""
import os
from playwright.sync_api import sync_playwright

USERNAME = os.environ.get("RADIUS_USERNAME")
PASSWORD = os.environ.get("RADIUS_PASSWORD")

if not USERNAME or not PASSWORD:
    raise EnvironmentError("Set RADIUS_USERNAME and RADIUS_PASSWORD first.")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, slow_mo=200)
    with browser:
        context = browser.new_context(accept_downloads=True)
        with context:
            page = context.new_page()

            print("\n[1/3] Logging in...")
            page.goto("https://radius.mathnasium.com")
            page.fill("#UserName", USERNAME)
            page.fill("#Password", PASSWORD)
            page.click("#login")
            page.wait_for_load_state("networkidle")
            print("      Logged in.")

            print("\n[2/3] Opening Enrollment Report...")
            page.goto("https://radius.mathnasium.com/Enrollment/EnrollmentReport")
            page.wait_for_load_state("networkidle")
            print("      Page loaded.")

            print("\n[3/3] WHAT TO DO:")
            print("  - Open DevTools (Cmd+Option+I → Elements tab)")
            print("  - Find and record:")
            print("      * Start date picker ID (e.g. #StartDate)")
            print("      * End date picker ID (e.g. #EndDate)")
            print("      * Center filter ID")
            print("      * Search button ID")
            print("      * Export/Excel button ID")
            print("  - Set date range to 9/1/2022 → today, select both centers")
            print("  - Click Search, then Export to Excel")
            print("  - Open the downloaded file and note EXACT column headers")
            print("\nPress Enter when done.")
            input()

print("\nDone. Update config.py with your findings.")
```

- [ ] **Step 2: Install dependencies and run discover_selectors.py**

```bash
cd ~/mathnasium-new-enrollments
pip install -r requirements.txt
playwright install chromium
source ~/.mathnasium_env   # loads RADIUS_USERNAME and RADIUS_PASSWORD
python discover_selectors.py
```

- [ ] **Step 3: Update config.py with confirmed column names**

After inspecting the downloaded Excel, replace all `TBD_*` strings in `config.py` with the exact column headers seen in the file. Also update the Enrollment Report URL and selector constants based on DevTools findings. Example (actual values will differ):

```python
# Enrollment Report selectors
ENROLL_START_DATE_SELECTOR = "#StartDate"    # confirmed from DevTools
ENROLL_END_DATE_SELECTOR   = "#EndDate"
ENROLL_CENTER_SELECTOR     = "#AllCenterListMultiSelect"
ENROLL_SEARCH_BTN          = "#btnsearch"
ENROLL_EXPORT_BTN          = "#btnExport"

# Enrollment Report column names — confirmed from downloaded Excel
COL_ENROLL_FIRST_NAME  = "Student First Name"   # replace with actual
COL_ENROLL_LAST_NAME   = "Student Last Name"
COL_ENROLL_CENTER      = "Center Name"
COL_ENROLL_START_DATE  = "Enrollment Start Date"
COL_ENROLL_END_DATE    = "Enrollment End Date"  # set to None if not in export
COL_ENROLL_PROGRAM     = "Program Type"
```

- [ ] **Step 4: Commit**

```bash
cd ~/mathnasium-new-enrollments
git add discover_selectors.py config.py
git commit -m "feat: add discover_selectors and confirmed enrollment report config"
```

---

## Task 3: Transform — detect new enrollments (TDD)

**Files:**
- Create: `~/mathnasium-new-enrollments/transform.py`
- Modify: `~/mathnasium-new-enrollments/tests/test_transform.py`

- [ ] **Step 1: Write failing tests for detect_new_enrollments()**

```python
# tests/test_transform.py
import pandas as pd
import pytest
from datetime import date
from transform import detect_new_enrollments

def _make_enrollment(first, last, center, start, end=None, program="2x/week"):
    return {
        "Student First Name": first,
        "Student Last Name": last,
        "Center Name": center,
        "Enrollment Start Date": pd.Timestamp(start),
        "Enrollment End Date": pd.Timestamp(end) if end else pd.NaT,
        "Program Type": program,
    }

def test_brand_new_student_flagged():
    """Student with no prior enrollment is a new enrollment."""
    df = pd.DataFrame([
        _make_enrollment("Alice", "Smith", "Teaneck", "2023-03-01"),
    ])
    result = detect_new_enrollments(df, report_start=date(2023, 1, 1))
    assert len(result) == 1
    assert result.iloc[0]["enrollment_month"] == pd.Period("2023-03", "M")

def test_returner_within_90_days_excluded():
    """Student who left and returned within 90 days is NOT a new enrollment."""
    df = pd.DataFrame([
        _make_enrollment("Bob", "Jones", "Teaneck", "2022-10-01", end="2023-01-15"),
        _make_enrollment("Bob", "Jones", "Teaneck", "2023-03-01"),  # 45 days after end
    ])
    result = detect_new_enrollments(df, report_start=date(2023, 1, 1))
    assert len(result) == 0

def test_returner_after_90_days_included():
    """Student who left and returned after 90+ days IS a new enrollment."""
    df = pd.DataFrame([
        _make_enrollment("Carol", "Lee", "Englewood", "2022-09-01", end="2022-11-30"),
        _make_enrollment("Carol", "Lee", "Englewood", "2023-05-01"),  # 152 days after end
    ])
    result = detect_new_enrollments(df, report_start=date(2023, 1, 1))
    assert len(result) == 1
    assert result.iloc[0]["enrollment_month"] == pd.Period("2023-05", "M")

def test_enrollment_before_report_start_excluded():
    """Enrollment starting before Jan 2023 is not in the output."""
    df = pd.DataFrame([
        _make_enrollment("Dave", "Brown", "Teaneck", "2022-11-01"),
    ])
    result = detect_new_enrollments(df, report_start=date(2023, 1, 1))
    assert len(result) == 0

def test_student_name_combined():
    """Output has a combined Student Name column."""
    df = pd.DataFrame([
        _make_enrollment("Eve", "White", "Teaneck", "2023-06-01"),
    ])
    result = detect_new_enrollments(df, report_start=date(2023, 1, 1))
    assert result.iloc[0]["Student Name"] == "Eve White"

def test_multiple_centers_independent():
    """Same student name at different centers are treated independently."""
    df = pd.DataFrame([
        _make_enrollment("Frank", "Gray", "Teaneck",   "2023-02-01"),
        _make_enrollment("Frank", "Gray", "Englewood", "2023-02-01"),
    ])
    result = detect_new_enrollments(df, report_start=date(2023, 1, 1))
    assert len(result) == 2

def test_no_end_date_treated_as_ongoing():
    """Prior enrollment with no end date is treated as ongoing (returner excluded)."""
    df = pd.DataFrame([
        _make_enrollment("Grace", "Hall", "Teaneck", "2022-09-01", end=None),
        _make_enrollment("Grace", "Hall", "Teaneck", "2023-04-01"),
    ])
    result = detect_new_enrollments(df, report_start=date(2023, 1, 1))
    # Prior enrollment has no end date, treated as still active — not a gap
    assert len(result) == 0
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd ~/mathnasium-new-enrollments
pytest tests/test_transform.py -v
```

Expected: FAIL — `ModuleNotFoundError: No module named 'transform'`

- [ ] **Step 3: Implement detect_new_enrollments() in transform.py**

```python
# transform.py
import pandas as pd
from datetime import date
from config import (
    COL_ENROLL_FIRST_NAME, COL_ENROLL_LAST_NAME, COL_ENROLL_CENTER,
    COL_ENROLL_START_DATE, COL_ENROLL_END_DATE, COL_ENROLL_PROGRAM,
    REENROLL_GAP_DAYS,
)


def detect_new_enrollments(df: pd.DataFrame, report_start: date) -> pd.DataFrame:
    """Return enrollment rows that represent true new enrollments.

    A row is a new enrollment if:
    - Its start date >= report_start, AND
    - The student+center has no prior enrollment that ended within REENROLL_GAP_DAYS
      before this start date (NaT end date = ongoing, treated as within gap).

    Returns DataFrame with enrollment_month (pd.Period M) and Student Name columns added.
    """
    df = df.copy()
    df[COL_ENROLL_START_DATE] = pd.to_datetime(df[COL_ENROLL_START_DATE])
    if COL_ENROLL_END_DATE and COL_ENROLL_END_DATE in df.columns:
        df[COL_ENROLL_END_DATE] = pd.to_datetime(df[COL_ENROLL_END_DATE])
    else:
        df["_end"] = pd.NaT

    end_col = COL_ENROLL_END_DATE if (COL_ENROLL_END_DATE and COL_ENROLL_END_DATE in df.columns) else "_end"
    report_start_ts = pd.Timestamp(report_start)

    new_rows = []
    for (first, last, center), group in df.groupby(
        [COL_ENROLL_FIRST_NAME, COL_ENROLL_LAST_NAME, COL_ENROLL_CENTER]
    ):
        group = group.sort_values(COL_ENROLL_START_DATE).reset_index(drop=True)
        for i, row in group.iterrows():
            start = row[COL_ENROLL_START_DATE]
            if pd.isna(start) or start < report_start_ts:
                continue
            prior = group[group[COL_ENROLL_START_DATE] < start]
            if prior.empty:
                new_rows.append(row)
                continue
            most_recent_end = prior[end_col].max()
            if pd.isna(most_recent_end):
                continue  # prior enrollment still ongoing — not a new enrollment
            gap = (start - most_recent_end).days
            if gap > REENROLL_GAP_DAYS:
                new_rows.append(row)

    if not new_rows:
        result = df.iloc[0:0].copy()
        result["Student Name"] = pd.Series(dtype=str)
        result["enrollment_month"] = pd.Series(dtype="period[M]")
        return result

    result = pd.DataFrame(new_rows).reset_index(drop=True)
    result["Student Name"] = result[COL_ENROLL_FIRST_NAME].str.strip() + " " + result[COL_ENROLL_LAST_NAME].str.strip()
    result["enrollment_month"] = result[COL_ENROLL_START_DATE].dt.to_period("M")
    return result
```

- [ ] **Step 4: Run tests to confirm they pass**

```bash
cd ~/mathnasium-new-enrollments
pytest tests/test_transform.py -v
```

Expected: All 7 tests PASS

- [ ] **Step 5: Commit**

```bash
cd ~/mathnasium-new-enrollments
git add transform.py tests/test_transform.py
git commit -m "feat: add detect_new_enrollments with tests"
```

---

## Task 4: Transform — join payment amounts (TDD)

**Files:**
- Modify: `~/mathnasium-new-enrollments/transform.py`
- Modify: `~/mathnasium-new-enrollments/tests/test_transform.py`

- [ ] **Step 1: Add failing tests for join_payment_amounts()**

Add to `tests/test_transform.py`:

```python
from transform import join_payment_amounts

def _make_new_enrollment_row(student_name, center, month_str, program="2x/week"):
    return {
        "Student Name": student_name,
        "Center Name": center,
        "enrollment_month": pd.Period(month_str, "M"),
        "Program Type": program,
    }

def _make_payment_row(student_name, center, month_str, amount):
    return {
        "Student Name": student_name,
        "Center Name": center,
        "payment_month": pd.Period(month_str, "M"),
        "Amount": amount,
    }

def test_payment_joined_by_name_center_month():
    """Amount is matched by student name + center + month."""
    enrollments = pd.DataFrame([
        _make_new_enrollment_row("Alice Smith", "Teaneck", "2023-03"),
    ])
    payments = pd.DataFrame([
        _make_payment_row("Alice Smith", "Teaneck", "2023-03", 299.0),
    ])
    result = join_payment_amounts(enrollments, payments)
    assert result.iloc[0]["Amount Charged"] == 299.0

def test_payment_name_case_insensitive():
    """Name matching is case-insensitive."""
    enrollments = pd.DataFrame([
        _make_new_enrollment_row("bob jones", "Englewood", "2023-05"),
    ])
    payments = pd.DataFrame([
        _make_payment_row("BOB JONES", "Englewood", "2023-05", 399.0),
    ])
    result = join_payment_amounts(enrollments, payments)
    assert result.iloc[0]["Amount Charged"] == 399.0

def test_no_payment_match_is_blank():
    """Enrollment with no payment match gets NaN amount."""
    enrollments = pd.DataFrame([
        _make_new_enrollment_row("Carol Lee", "Teaneck", "2023-06"),
    ])
    payments = pd.DataFrame([])
    result = join_payment_amounts(enrollments, payments)
    assert pd.isna(result.iloc[0]["Amount Charged"])

def test_wrong_month_not_matched():
    """Payment from different month is not matched."""
    enrollments = pd.DataFrame([
        _make_new_enrollment_row("Dave Brown", "Teaneck", "2023-03"),
    ])
    payments = pd.DataFrame([
        _make_payment_row("Dave Brown", "Teaneck", "2023-04", 299.0),
    ])
    result = join_payment_amounts(enrollments, payments)
    assert pd.isna(result.iloc[0]["Amount Charged"])
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd ~/mathnasium-new-enrollments
pytest tests/test_transform.py::test_payment_joined_by_name_center_month -v
```

Expected: FAIL — `ImportError: cannot import name 'join_payment_amounts'`

- [ ] **Step 3: Implement join_payment_amounts() in transform.py**

Add to `transform.py`:

```python
def _normalize(s: str) -> str:
    return str(s).strip().lower()


def join_payment_amounts(
    df_enrollments: pd.DataFrame,
    df_payments: pd.DataFrame,
) -> pd.DataFrame:
    """Join payment amounts onto new enrollment rows by student name + center + month.

    df_payments must have columns: Student Name, Center Name, payment_month (Period M), Amount.
    Returns df_enrollments with an 'Amount Charged' column added (NaN if no match).
    """
    result = df_enrollments.copy()
    result["_name_key"]   = result["Student Name"].apply(_normalize)
    result["_center_key"] = result[COL_ENROLL_CENTER].apply(_normalize)

    if df_payments.empty:
        result["Amount Charged"] = float("nan")
        return result.drop(columns=["_name_key", "_center_key"])

    pay = df_payments.copy()
    pay["_name_key"]   = pay["Student Name"].apply(_normalize)
    pay["_center_key"] = pay["Center Name"].apply(_normalize)

    pay_agg = (
        pay.groupby(["_name_key", "_center_key", "payment_month"])["Amount"]
        .sum()
        .reset_index()
        .rename(columns={"Amount": "Amount Charged"})
    )

    merged = result.merge(
        pay_agg,
        left_on=["_name_key", "_center_key", "enrollment_month"],
        right_on=["_name_key", "_center_key", "payment_month"],
        how="left",
    )
    merged = merged.drop(columns=["_name_key", "_center_key", "payment_month"], errors="ignore")
    return merged
```

- [ ] **Step 4: Run all tests**

```bash
cd ~/mathnasium-new-enrollments
pytest tests/test_transform.py -v
```

Expected: All 11 tests PASS

- [ ] **Step 5: Commit**

```bash
cd ~/mathnasium-new-enrollments
git add transform.py tests/test_transform.py
git commit -m "feat: add join_payment_amounts with tests"
```

---

## Task 5: Transform — build output DataFrames (TDD)

**Files:**
- Modify: `~/mathnasium-new-enrollments/transform.py`
- Modify: `~/mathnasium-new-enrollments/tests/test_transform.py`

- [ ] **Step 1: Add failing tests for build_detail() and build_summary()**

Add to `tests/test_transform.py`:

```python
from transform import build_detail, build_summary

def _full_row(student, center, month_str, program, amount):
    return {
        "Student Name": student,
        "Center Name": center,
        "enrollment_month": pd.Period(month_str, "M"),
        "Program Type": program,
        "Amount Charged": amount,
    }

def test_build_detail_columns_and_sort():
    """Detail DataFrame has correct columns, sorted by month then name."""
    rows = [
        _full_row("Bob Jones",  "Teaneck", "2023-02", "3x/week", 399.0),
        _full_row("Alice Smith","Teaneck", "2023-01", "2x/week", 299.0),
    ]
    df = pd.DataFrame(rows)
    result = build_detail(df, center="Teaneck")
    assert list(result.columns) == [
        "Enrollment Month", "Center", "Student Name", "Program Type", "Amount Charged"
    ]
    assert result.iloc[0]["Enrollment Month"] == "January 2023"
    assert result.iloc[1]["Enrollment Month"] == "February 2023"

def test_build_detail_filters_center():
    """build_detail returns only rows for the specified center."""
    rows = [
        _full_row("Alice Smith", "Teaneck",   "2023-01", "2x/week", 299.0),
        _full_row("Bob Jones",   "Englewood", "2023-01", "3x/week", 399.0),
    ]
    df = pd.DataFrame(rows)
    result = build_detail(df, center="Teaneck")
    assert len(result) == 1
    assert result.iloc[0]["Student Name"] == "Alice Smith"

def test_build_summary_aggregation():
    """Summary aggregates count and total amount per month."""
    rows = [
        _full_row("Alice Smith", "Teaneck", "2023-01", "2x/week", 299.0),
        _full_row("Bob Jones",   "Teaneck", "2023-01", "3x/week", 399.0),
        _full_row("Carol Lee",   "Teaneck", "2023-02", "2x/week", 299.0),
    ]
    df = pd.DataFrame(rows)
    result = build_summary(df, center="Teaneck")
    assert list(result.columns) == [
        "Enrollment Month", "Center", "New Enrollments", "Total Amount Charged"
    ]
    jan = result[result["Enrollment Month"] == "January 2023"].iloc[0]
    assert jan["New Enrollments"] == 2
    assert jan["Total Amount Charged"] == 698.0
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd ~/mathnasium-new-enrollments
pytest tests/test_transform.py::test_build_detail_columns_and_sort -v
```

Expected: FAIL — `ImportError: cannot import name 'build_detail'`

- [ ] **Step 3: Implement build_detail() and build_summary() in transform.py**

Add to `transform.py`:

```python
def build_detail(df: pd.DataFrame, center: str) -> pd.DataFrame:
    """Return the Enrollments detail tab DataFrame for one center."""
    filtered = df[df[COL_ENROLL_CENTER].str.strip() == center].copy()
    filtered = filtered.sort_values(["enrollment_month", "Student Name"])
    filtered["Enrollment Month"] = filtered["enrollment_month"].apply(
        lambda p: p.to_timestamp().strftime("%B %Y")
    )
    return filtered[["Enrollment Month", COL_ENROLL_CENTER, "Student Name", COL_ENROLL_PROGRAM, "Amount Charged"]].rename(
        columns={COL_ENROLL_CENTER: "Center", COL_ENROLL_PROGRAM: "Program Type"}
    ).reset_index(drop=True)


def build_summary(df: pd.DataFrame, center: str) -> pd.DataFrame:
    """Return the Summary tab DataFrame for one center."""
    filtered = df[df[COL_ENROLL_CENTER].str.strip() == center].copy()
    agg = (
        filtered.groupby("enrollment_month")
        .agg(
            count=("Student Name", "count"),
            total=("Amount Charged", "sum"),
        )
        .reset_index()
        .sort_values("enrollment_month")
    )
    agg["Enrollment Month"] = agg["enrollment_month"].apply(
        lambda p: p.to_timestamp().strftime("%B %Y")
    )
    agg["Center"] = center
    return agg[["Enrollment Month", "Center"]].assign(
        **{"New Enrollments": agg["count"], "Total Amount Charged": agg["total"]}
    ).reset_index(drop=True)
```

- [ ] **Step 4: Run all tests**

```bash
cd ~/mathnasium-new-enrollments
pytest tests/test_transform.py -v
```

Expected: All 14 tests PASS

- [ ] **Step 5: Commit**

```bash
cd ~/mathnasium-new-enrollments
git add transform.py tests/test_transform.py
git commit -m "feat: add build_detail and build_summary with tests"
```

---

## Task 6: Download — enrollment report

**Files:**
- Create: `~/mathnasium-new-enrollments/download.py`

- [ ] **Step 1: Create download.py with login + enrollment report pull**

The enrollment report is exported to Excel (not read from the Kendo grid directly) to get all fields reliably. Selectors confirmed in Task 2.

```python
# download.py
import os
import calendar
from pathlib import Path
from datetime import date

from playwright.sync_api import sync_playwright

from config import (
    INPUT_DIR, RADIUS_LOGIN_URL, ENROLLMENT_REPORT_URL, PAYMENT_RECON_URL,
    ENROLL_PULL_START, PAYMENT_YEARS, CENTERS,
    ENROLL_START_DATE_SELECTOR, ENROLL_END_DATE_SELECTOR,
    ENROLL_CENTER_SELECTOR, ENROLL_SEARCH_BTN, ENROLL_EXPORT_BTN,
)

CENTER_IDS = [c["radius_id"] for c in CENTERS.values()]


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


def _set_enrollment_date_range(page, start: str, end: str) -> None:
    """Set enrollment report start/end date pickers."""
    page.evaluate(
        """([start, end]) => {
            var dpStart = $(arguments[0]).data('kendoDatePicker') ||
                          $(arguments[0]).data('kendoDateTimePicker');
            var dpEnd   = $(arguments[1]).data('kendoDatePicker') ||
                          $(arguments[1]).data('kendoDateTimePicker');
            if (dpStart) { dpStart.value(start); dpStart.trigger('change'); }
            if (dpEnd)   { dpEnd.value(end);     dpEnd.trigger('change'); }
        }""".replace("arguments[0]", f'"{ENROLL_START_DATE_SELECTOR}"')
           .replace("arguments[1]", f'"{ENROLL_END_DATE_SELECTOR}"'),
        [start, end],
    )


def download_enrollment_report() -> Path:
    """Pull the full enrollment report (9/1/2022 → today) as a single Excel export."""
    out_path = INPUT_DIR / "Enrollment_All.xlsx"
    if out_path.exists():
        print(f"  [skip] {out_path.name} already cached")
        return out_path

    today = date.today().strftime("%-m/%-d/%Y")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        with browser:
            context = browser.new_context(accept_downloads=True)
            with context:
                page = context.new_page()
                _login(page)

                page.goto(ENROLLMENT_REPORT_URL)
                page.wait_for_load_state("networkidle")

                _set_enrollment_date_range(page, ENROLL_PULL_START, today)

                # Select all centers
                page.evaluate(
                    """(ids) => {
                        var ms = $('#AllCenterListMultiSelect').data('kendoMultiSelect');
                        if (ms) { ms.value(ids); ms.trigger('change'); }
                    }""",
                    CENTER_IDS,
                )

                page.click(ENROLL_SEARCH_BTN)
                page.wait_for_load_state("networkidle")

                with page.expect_download() as dl:
                    page.click(ENROLL_EXPORT_BTN)
                dl.value.save_as(out_path)
                print(f"  [ok] Downloaded {out_path.name}")

    return out_path


def download_payment_recon_year(page, year: int) -> Path:
    """Download Payment Reconciliation for a full year. Returns cached path if exists."""
    out_path = INPUT_DIR / f"PaymentRecon_{year}.xlsx"
    if out_path.exists():
        print(f"  [skip] {out_path.name} already cached")
        return out_path

    if year == 2022:
        start = "9/1/2022"
        end = "12/31/2022"
    elif year == date.today().year:
        end = date.today().strftime("%-m/%-d/%Y")
        start = f"1/1/{year}"
    else:
        start = f"1/1/{year}"
        end = f"12/31/{year}"

    page.goto(PAYMENT_RECON_URL)
    page.wait_for_load_state("networkidle")

    page.evaluate(
        """([start, end]) => {
            var dpS = $('#startDateSelect').data('kendoDatePicker');
            var dpE = $('#endDateSelect').data('kendoDatePicker');
            dpS.value(start); dpS.trigger('change');
            dpE.value(end);   dpE.trigger('change');
        }""",
        [start, end],
    )

    page.evaluate(
        """(ids) => {
            var ms = $('#AllCenterListMultiSelect').data('kendoMultiSelect');
            ms.value(ids); ms.trigger('change');
        }""",
        CENTER_IDS,
    )

    page.click("#btnsearch")
    page.wait_for_load_state("networkidle")

    with page.expect_download(timeout=120_000) as dl:
        page.click("#btnExport")
    dl.value.save_as(out_path)
    print(f"  [ok] Downloaded {out_path.name}")
    return out_path


def download_payment_recon_all_years() -> dict[int, Path]:
    """Download payment recon for all configured years. Returns {year: Path}."""
    paths = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        with browser:
            context = browser.new_context(accept_downloads=True)
            with context:
                page = context.new_page()
                _login(page)
                for year in PAYMENT_YEARS:
                    paths[year] = download_payment_recon_year(page, year)
    return paths


def download_all() -> dict:
    """Download all data needed for the report. Caches to input/."""
    INPUT_DIR.mkdir(exist_ok=True)
    print("\n--- Downloading Enrollment Report ---")
    enrollment_path = download_enrollment_report()
    print("\n--- Downloading Payment Recon by Year ---")
    payment_paths = download_payment_recon_all_years()
    return {"enrollment": enrollment_path, "payment": payment_paths}
```

- [ ] **Step 2: Commit**

```bash
cd ~/mathnasium-new-enrollments
git add download.py
git commit -m "feat: add download.py for enrollment report and payment recon"
```

---

## Task 7: Inspect payment recon year export and update config

**Files:**
- Modify: `~/mathnasium-new-enrollments/config.py`

After downloading the first payment recon year export in Task 6, open it and record the exact column names.

- [ ] **Step 1: Run a one-year payment recon pull manually**

```bash
cd ~/mathnasium-new-enrollments
source ~/.mathnasium_env
python -c "
from download import download_payment_recon_all_years
from config import PAYMENT_YEARS
import config; config.PAYMENT_YEARS = [2023]
download_payment_recon_all_years()
"
```

Open `input/PaymentRecon_2023.xlsx` and record:
- Is there a row-per-month, or one aggregated row per account for the full year?
- What is the exact column name for student name (if present)?
- What is the exact column name for center?
- What is the exact column name for amount paid?
- Is there a date column? If yes, what is it called?

- [ ] **Step 2: Update config.py with confirmed payment recon column names**

Example (replace with actual confirmed values):

```python
# Payment Recon column names — confirmed from 2023 year export
COL_PAY_STUDENT = "Student Names"   # replace with actual column header
COL_PAY_CENTER  = "Center"
COL_PAY_AMOUNT  = "Amount Paid"
COL_PAY_DATE    = "Payment Date"    # set to None if no date column present
```

If the export has NO date column (i.e., amounts are annual totals), note this and proceed to Task 7b. Otherwise continue to Task 8.

- [ ] **Step 3: Commit config update**

```bash
cd ~/mathnasium-new-enrollments
git add config.py
git commit -m "chore: confirm payment recon column names from live export"
```

---

## Task 8: Transform — load and prepare payment data

**Files:**
- Modify: `~/mathnasium-new-enrollments/transform.py`
- Modify: `~/mathnasium-new-enrollments/tests/test_transform.py`

- [ ] **Step 1: Add failing test for load_payment_data()**

Add to `tests/test_transform.py`:

```python
import os, tempfile
from transform import load_payment_data

def test_load_payment_data_concatenates_years():
    """load_payment_data reads multiple Excel files and concatenates them."""
    row = {"Student Names": "Alice Smith", "Center": "Teaneck",
           "Amount Paid": 299.0, "Payment Date": "2023-03-01"}
    with tempfile.TemporaryDirectory() as tmp:
        import pathlib
        p = pathlib.Path(tmp)
        pd.DataFrame([row]).to_excel(p / "PaymentRecon_2023.xlsx", index=False)
        result = load_payment_data(p)
    assert len(result) == 1
    assert "Student Name" in result.columns
    assert "payment_month" in result.columns
```

- [ ] **Step 2: Run test to confirm it fails**

```bash
cd ~/mathnasium-new-enrollments
pytest tests/test_transform.py::test_load_payment_data_concatenates_years -v
```

Expected: FAIL

- [ ] **Step 3: Implement load_payment_data() in transform.py**

Add to `transform.py`:

```python
import glob
from config import COL_PAY_STUDENT, COL_PAY_CENTER, COL_PAY_AMOUNT, COL_PAY_DATE


def load_payment_data(input_dir: Path) -> pd.DataFrame:
    """Read all PaymentRecon_YYYY.xlsx files, concatenate, and normalize.

    Returns DataFrame with columns: Student Name, Center Name, Amount, payment_month (Period M).
    If COL_PAY_DATE is None, payment_month is set to NaT and the join will match on name+center only.
    """
    files = sorted(input_dir.glob("PaymentRecon_*.xlsx"))
    if not files:
        return pd.DataFrame(columns=["Student Name", "Center Name", "Amount", "payment_month"])

    frames = [pd.read_excel(f) for f in files]
    df = pd.concat(frames, ignore_index=True)

    df = df.rename(columns={
        COL_PAY_STUDENT: "Student Name",
        COL_PAY_CENTER:  "Center Name",
        COL_PAY_AMOUNT:  "Amount",
    })

    if COL_PAY_DATE and COL_PAY_DATE in df.columns:
        df["payment_month"] = pd.to_datetime(df[COL_PAY_DATE], errors="coerce").dt.to_period("M")
    else:
        df["payment_month"] = pd.NaT

    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    return df[["Student Name", "Center Name", "Amount", "payment_month"]]
```

- [ ] **Step 4: Run all tests**

```bash
cd ~/mathnasium-new-enrollments
pytest tests/test_transform.py -v
```

Expected: All 15 tests PASS

- [ ] **Step 5: Commit**

```bash
cd ~/mathnasium-new-enrollments
git add transform.py tests/test_transform.py
git commit -m "feat: add load_payment_data with tests"
```

---

## Task 9: Output — write Excel files

**Files:**
- Create: `~/mathnasium-new-enrollments/output.py`

- [ ] **Step 1: Create output.py**

```python
# output.py
from pathlib import Path
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def _autofit_columns(ws) -> None:
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 50)


def _style_header(ws) -> None:
    fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center")


def write_center_excel(
    center: str,
    df_detail: pd.DataFrame,
    df_summary: pd.DataFrame,
    output_dir: Path,
) -> Path:
    """Write a two-tab Excel file for one center."""
    output_dir.mkdir(exist_ok=True)
    out_path = output_dir / f"NewEnrollments_{center}.xlsx"

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_detail.to_excel(writer, sheet_name="Enrollments", index=False)
        df_summary.to_excel(writer, sheet_name="Summary", index=False)

        for sheet_name in ["Enrollments", "Summary"]:
            ws = writer.sheets[sheet_name]
            _style_header(ws)
            _autofit_columns(ws)

    print(f"  [ok] Written: {out_path}")
    return out_path
```

- [ ] **Step 2: Commit**

```bash
cd ~/mathnasium-new-enrollments
git add output.py
git commit -m "feat: add output.py Excel writer"
```

---

## Task 10: Main orchestrator

**Files:**
- Create: `~/mathnasium-new-enrollments/main.py`

- [ ] **Step 1: Create main.py**

```python
# main.py
import argparse
from datetime import date
import pandas as pd

from config import INPUT_DIR, OUTPUT_DIR, CENTERS, REPORT_START, COL_ENROLL_CENTER
from download import download_all
from transform import detect_new_enrollments, load_payment_data, join_payment_amounts, build_detail, build_summary
from output import write_center_excel


def parse_args():
    parser = argparse.ArgumentParser(description="Generate new enrollments report.")
    parser.add_argument("--from", dest="from_date", default=None,
                        help="Report start YYYY-MM (default: 2023-01)")
    parser.add_argument("--to", dest="to_date", default=None,
                        help="Report end YYYY-MM (default: current month)")
    return parser.parse_args()


def main():
    args = parse_args()

    report_start = (
        date.fromisoformat(args.from_date + "-01") if args.from_date
        else REPORT_START
    )
    if args.to_date:
        to_period = pd.Period(args.to_date, "M")
    else:
        today = date.today()
        to_period = pd.Period(f"{today.year}-{today.month:02d}", "M")

    print("=== New Enrollments Report ===")
    print(f"Report window: {report_start.strftime('%B %Y')} → {to_period}")

    # Download (uses cache when available)
    paths = download_all()

    # Load enrollment data
    print("\n--- Processing ---")
    df_enroll = pd.read_excel(paths["enrollment"])
    df_new = detect_new_enrollments(df_enroll, report_start=report_start)
    print(f"  Found {len(df_new)} new enrollment records")

    # Filter to report window
    df_new = df_new[df_new["enrollment_month"] <= to_period]

    # Load and join payment data
    df_pay = load_payment_data(INPUT_DIR)
    df_joined = join_payment_amounts(df_new, df_pay)

    # Write one Excel per center
    print("\n--- Writing output ---")
    for center in CENTERS:
        df_detail  = build_detail(df_joined, center=center)
        df_summary = build_summary(df_joined, center=center)
        write_center_excel(center, df_detail, df_summary, OUTPUT_DIR)

    print("\nDone.")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Commit**

```bash
cd ~/mathnasium-new-enrollments
git add main.py
git commit -m "feat: add main.py orchestrator"
```

---

## Task 11: End-to-end run and validation

- [ ] **Step 1: Run the full pipeline**

```bash
cd ~/mathnasium-new-enrollments
source ~/.mathnasium_env
python main.py
```

Expected output:
```
=== New Enrollments Report ===
Report window: January 2023 → June 2026

--- Downloading Enrollment Report ---
  [ok] Downloaded Enrollment_All.xlsx

--- Downloading Payment Recon by Year ---
  [ok] Downloaded PaymentRecon_2022.xlsx
  ...

--- Processing ---
  Found NNN new enrollment records

--- Writing output ---
  [ok] Written: output/NewEnrollments_Teaneck.xlsx
  [ok] Written: output/NewEnrollments_Englewood.xlsx

Done.
```

- [ ] **Step 2: Open and verify output files**

Open `output/NewEnrollments_Teaneck.xlsx` and check:
- Enrollments tab has correct columns: Enrollment Month, Center, Student Name, Program Type, Amount Charged
- Months are formatted as "January 2023", "February 2023", etc.
- Summary tab shows count and total per month
- No obvious errors (e.g., all Amount Charged blank → column name mismatch in config.py)

- [ ] **Step 3: Spot-check a few known students**

Pick 2–3 students you know enrolled in a specific month and verify they appear in the correct month with a reasonable amount.

- [ ] **Step 4: Re-run to verify caching**

```bash
python main.py
```

All download steps should print `[skip]` and the run should complete in seconds.

- [ ] **Step 5: Final commit**

```bash
cd ~/mathnasium-new-enrollments
git add -A
git commit -m "feat: new enrollments report complete"
```
