# Birthdays & Level Ups Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a monthly automation that downloads two Radius reports, calculates birthdays and level-ups per center, and emails a formatted HTML report to each center director on the 1st of every month.

**Architecture:** Four-module Python script (`download.py` → `process.py` → `deliver.py` → `main.py`) in a new GitHub repo `mathnasium-birthdays-levelups`, triggered by GitHub Actions cron `0 15 1 * *`. Sends one HTML email per center with no attachment.

**Tech Stack:** Python 3.11+, Playwright (headless Chromium), pandas, smtplib (Gmail SMTP, SSL port 465), GitHub Actions

---

## File Map

| File | Responsibility |
|---|---|
| `config.py` | All constants: URLs, center IDs, recipients, column names, paths |
| `download.py` | Playwright: log into Radius, download enrollment + birthdays reports |
| `process.py` | pandas: parse reports, compute level-ups and birthdays per center |
| `deliver.py` | Build HTML email body, send via Gmail SMTP |
| `run_log.py` | Read/write `run_log.json` |
| `main.py` | Orchestrate download → process → deliver → log; `--month` flag |
| `tests/test_process.py` | Unit tests for level-up logic and birthday parsing |
| `.github/workflows/monthly_birthdays_levelups.yml` | GitHub Actions cron + manual trigger |

---

### Task 1: Scaffold repo

**Files:**
- Create: `~/mathnasium-birthdays-levelups/`
- Create: `requirements.txt`
- Create: `.gitignore`
- Create: `config.py`
- Create: `input/.gitkeep`, `output/.gitkeep`
- Create: `run_log.json`

- [ ] **Step 1: Create the local directory and initialize git**

```bash
mkdir ~/mathnasium-birthdays-levelups
cd ~/mathnasium-birthdays-levelups
git init -b main
```

- [ ] **Step 2: Create `requirements.txt`**

```
playwright==1.44.0
pandas==2.2.2
openpyxl==3.1.2
pytest==8.2.0
```

- [ ] **Step 3: Create `.gitignore`**

```
input/
__pycache__/
*.pyc
.env
```

- [ ] **Step 4: Create `config.py`**

```python
import os
from pathlib import Path

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"

# ── Radius ────────────────────────────────────────────────────────────────────
# Verify this URL against ~/mathnasium-page-goals/download.py
RADIUS_LOGIN_URL = "https://go.mathnasium.com/"

CENTERS = {
    "Englewood": {
        "radius_id": "2428",
        "recipient": "englewood@mathnasium.com",
    },
    "Teaneck": {
        "radius_id": "2871",
        "recipient": "teaneck@mathnasium.com",
    },
}
CC_RECIPIENT = "matt.diamond@mathnasium.com"

# ── Column names ──────────────────────────────────────────────────────────────
# These must be verified against real Radius export files in Task 2 and 3.
# Update here if the actual column names differ.
ENROLLMENT_COL_STUDENT = "Student Name"
ENROLLMENT_COL_MONTHS  = "Length of Stay"
ENROLLMENT_COL_CENTER  = "Center"

BIRTHDAY_COL_STUDENT   = "Student Name"
BIRTHDAY_COL_BIRTHDAY  = "Birthday"
BIRTHDAY_COL_CENTER    = "Center"
```

- [ ] **Step 5: Create input/output dirs and empty run log**

```bash
mkdir input output
touch input/.gitkeep output/.gitkeep
echo "[]" > run_log.json
```

- [ ] **Step 6: Create the GitHub repo**

Go to github.com → New repository → name: `mathnasium-birthdays-levelups` → Private → Create. Then:

```bash
cd ~/mathnasium-birthdays-levelups
git add .
git commit -m "chore: initial scaffold"
git remote add origin https://github.com/mdmathnasiums/mathnasium-birthdays-levelups.git
git push -u origin main
```

- [ ] **Step 7: Add GitHub secrets to the new repo**

Go to the repo on GitHub → Settings → Secrets and variables → Actions → New repository secret. Add:
- `RADIUS_USERNAME` — your Radius login email
- `RADIUS_PASSWORD` — your Radius password
- `SMTP_USER` — the Gmail address used in page goals
- `SMTP_PASSWORD` — the Gmail App Password used in page goals

---

### Task 2: Discover Radius selectors and download Enrollment Report

> **Why this comes first:** The biggest risk in any automation is the data acquisition step. Get a real file out before writing any processing code. Verify actual column names.

**Files:**
- Create: `download.py` (partial — login + enrollment report only)

- [ ] **Step 1: Install Playwright browsers**

```bash
cd ~/mathnasium-birthdays-levelups
pip install -r requirements.txt
playwright install chromium
```

- [ ] **Step 2: Write a headed discovery script**

Create `discover_selectors.py` in the repo root (will be deleted after Task 3):

```python
"""
Temporary headed browser script to navigate Radius and find selectors.
Run with: python discover_selectors.py
"""
import os
from playwright.sync_api import sync_playwright

USERNAME = os.environ.get("RADIUS_USERNAME")
PASSWORD = os.environ.get("RADIUS_PASSWORD")
LOGIN_URL = "https://go.mathnasium.com/"  # update if wrong

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, slow_mo=500)
    with browser:
        page = browser.new_page()
        page.goto(LOGIN_URL)
        page.fill("#UserName", USERNAME)
        page.fill("#Password", PASSWORD)
        page.click("#login")
        page.wait_for_load_state("networkidle")
        print("Logged in. Navigate to the Enrollment Report manually.")
        print("Use DevTools (F12) to inspect elements and note selectors.")
        print("Press Enter in this terminal when done.")
        input()
```

- [ ] **Step 3: Run the discovery script with real credentials**

```bash
RADIUS_USERNAME="your@email.com" RADIUS_PASSWORD="yourpass" python discover_selectors.py
```

Navigate to the Enrollment Report in the headed browser. Record these selectors (you'll need them in Step 5):
- Navigation menu item to reach the Enrollment Report
- Center MultiSelect selector (likely `#AllCenterListMultiSelect` — same as DWP)
- Any date or filter controls on the Enrollment Report
- The Excel/export button selector
- The column header row in the downloaded file — note exact names for `Student Name`, `Length of Stay`, `Center`

- [ ] **Step 4: Download a real Enrollment Report file and inspect it**

After discovering navigation, manually trigger the export in the headed browser. Open the downloaded file and note:
- Exact column name for student name
- Exact column name for length of stay (is it a decimal like `12.1`, whole number, or formatted string?)
- Exact column name for center

Update `config.py` if any column names differ from the defaults set in Task 1.

- [ ] **Step 5: Write `download.py` with login + enrollment download**

Fill in `[ENROLLMENT_NAV_SELECTOR]` and `[ENROLLMENT_EXPORT_BTN]` with what you discovered:

```python
import os
import json
from pathlib import Path
from playwright.sync_api import sync_playwright
from config import RADIUS_LOGIN_URL, INPUT_DIR, CENTERS

CENTER_IDS = [c["radius_id"] for c in CENTERS.values()]


def _login(page):
    """Log into Radius."""
    page.goto(RADIUS_LOGIN_URL)
    page.fill("#UserName", os.environ.get("RADIUS_USERNAME"))
    page.fill("#Password", os.environ.get("RADIUS_PASSWORD"))
    page.click("#login")
    page.wait_for_load_state("networkidle")


def _select_all_centers(page):
    """Select both Englewood and Teaneck in the center MultiSelect."""
    page.evaluate(
        """(ids) => {
            const ms = $('#AllCenterListMultiSelect').data('kendoMultiSelect');
            ms.value(ids);
            ms.trigger('change');
        }""",
        json.dumps(CENTER_IDS),
    )


def download_enrollment_report(page, month_label: str) -> Path:
    """
    Navigate to the Enrollment Report, select both centers, export to Excel.
    Returns path to downloaded file.
    """
    out_path = INPUT_DIR / f"Enrollment_{month_label}.xlsx"

    # Navigate to Enrollment Report
    # TODO: replace with actual selector discovered in Step 3
    # e.g., page.click("text=Enrollment Report") or page.goto(URL)
    page.click("[ENROLLMENT_NAV_SELECTOR]")
    page.wait_for_load_state("networkidle")

    _select_all_centers(page)

    # Click search/filter to load results
    page.click("#btnsearch")  # verify this is the correct search button
    page.wait_for_load_state("networkidle")

    # Export to Excel
    with page.expect_download() as dl:
        page.click("[ENROLLMENT_EXPORT_BTN]")  # replace with real selector
    download = dl.value
    download.save_as(out_path)

    return out_path


def download_reports(month_label: str) -> dict:
    """Download both reports. Returns dict with paths."""
    INPUT_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        with browser:
            context = browser.new_context(accept_downloads=True)
            with context:
                page = context.new_page()
                _login(page)
                enrollment_path = download_enrollment_report(page, month_label)

    return {"enrollment": enrollment_path}
```

- [ ] **Step 6: Replace placeholder selectors with discovered values**

Edit `download.py` and replace `[ENROLLMENT_NAV_SELECTOR]` and `[ENROLLMENT_EXPORT_BTN]` with the actual selectors found in Step 3.

- [ ] **Step 7: Test the download**

```bash
RADIUS_USERNAME="your@email.com" RADIUS_PASSWORD="yourpass" python -c "
from download import download_reports
paths = download_reports('2026-04')
print('Downloaded:', paths)
"
```

Expected: a file appears at `input/Enrollment_2026-04.xlsx`. Open it and verify it has data for both centers.

- [ ] **Step 8: Commit**

```bash
git add download.py config.py
git commit -m "feat: add enrollment report download"
```

---

### Task 3: Add Birthdays Report download

**Files:**
- Modify: `download.py`

- [ ] **Step 1: Navigate to Birthdays Report in headed browser**

Reuse `discover_selectors.py` to navigate to the Birthdays Report. Record:
- Navigation selector to reach the Birthdays Report
- The "Enrolled" filter selector (the spec says you need to change the enrollment filter to "Enrolled")
- The export button selector
- Exact column names: student name, birthday date, center

Update `config.py` if birthday column names differ from defaults.

- [ ] **Step 2: Add `download_birthdays_report()` to `download.py`**

Add this function to `download.py`, filling in discovered selectors:

```python
def download_birthdays_report(page, month_label: str) -> Path:
    """
    Navigate to the Birthdays Report, filter to enrolled students,
    select both centers, export to Excel.
    Returns path to downloaded file.
    """
    out_path = INPUT_DIR / f"Birthdays_{month_label}.xlsx"

    # Navigate to Birthdays Report
    # TODO: replace with actual selector
    page.click("[BIRTHDAYS_NAV_SELECTOR]")
    page.wait_for_load_state("networkidle")

    # Set enrollment filter to "Enrolled"
    # TODO: replace with actual selector (e.g., a dropdown)
    page.select_option("[ENROLLMENT_FILTER_SELECTOR]", "Enrolled")

    _select_all_centers(page)

    page.click("#btnsearch")  # verify correct search button
    page.wait_for_load_state("networkidle")

    with page.expect_download() as dl:
        page.click("[BIRTHDAYS_EXPORT_BTN]")  # replace with real selector
    download = dl.value
    download.save_as(out_path)

    return out_path
```

- [ ] **Step 3: Replace placeholder selectors with discovered values**

Edit `download.py` and replace all `[BIRTHDAYS_*]` placeholders with real selectors.

- [ ] **Step 4: Update `download_reports()` to include birthdays**

Replace the existing `download_reports()` function:

```python
def download_reports(month_label: str) -> dict:
    """Download both reports. Returns dict with paths to each file."""
    INPUT_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        with browser:
            context = browser.new_context(accept_downloads=True)
            with context:
                page = context.new_page()
                _login(page)
                enrollment_path = download_enrollment_report(page, month_label)
                birthdays_path = download_birthdays_report(page, month_label)

    return {
        "enrollment": enrollment_path,
        "birthdays": birthdays_path,
    }
```

- [ ] **Step 5: Test both downloads**

```bash
RADIUS_USERNAME="your@email.com" RADIUS_PASSWORD="yourpass" python -c "
from download import download_reports
paths = download_reports('2026-04')
print('Enrollment:', paths['enrollment'])
print('Birthdays:', paths['birthdays'])
"
```

Expected: two files in `input/`. Open both and verify data.

- [ ] **Step 6: Delete `discover_selectors.py` and commit**

```bash
rm discover_selectors.py
git add download.py config.py
git rm discover_selectors.py
git commit -m "feat: add birthdays report download, complete download.py"
```

---

### Task 4: Verify SMTP credentials

> **Why this comes before process.py:** Confirm email delivery works before building the full pipeline. A wrong App Password wastes time later.

**Files:**
- Create: `test_smtp.py` (temporary, deleted after this task)

- [ ] **Step 1: Create `test_smtp.py`**

```python
"""
Temporary script to verify Gmail SMTP credentials.
Run with: SMTP_USER=... SMTP_PASSWORD=... python test_smtp.py
Delete after confirming email arrives.
"""
import os
import smtplib
from email.mime.text import MIMEText

smtp_user = os.environ.get("SMTP_USER")
smtp_password = os.environ.get("SMTP_PASSWORD")

msg = MIMEText("SMTP test from mathnasium-birthdays-levelups. If you see this, it works!")
msg["Subject"] = "SMTP Test — Birthdays & Level Ups"
msg["From"] = smtp_user
msg["To"] = smtp_user  # send to yourself

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(smtp_user, smtp_password)
    server.sendmail(smtp_user, [smtp_user], msg.as_string())

print("Test email sent. Check your inbox.")
```

- [ ] **Step 2: Run it**

```bash
SMTP_USER="your@gmail.com" SMTP_PASSWORD="your-app-password" python test_smtp.py
```

Expected output: `Test email sent. Check your inbox.`  
Expected result: email arrives in inbox within 30 seconds.

If you get error 534: the password is a regular Gmail password, not an App Password. Go to passwords.google.com, create an App Password for "Mail", use that 16-character string instead.

- [ ] **Step 3: Delete `test_smtp.py` and commit**

```bash
git rm test_smtp.py
git commit -m "chore: verify SMTP (test script removed)"
```

---

### Task 5: TDD for `process.py`

**Files:**
- Create: `tests/test_process.py`
- Create: `process.py`

- [ ] **Step 1: Create `tests/` directory and write failing tests**

```bash
mkdir tests
touch tests/__init__.py
```

Create `tests/test_process.py`:

```python
import pandas as pd
import pytest
from process import calculate_level, get_levelups, get_birthdays, split_by_center

# ── calculate_level ──────────────────────────────────────────────────────────

def test_calculate_level_zero():
    assert calculate_level(0.0) == 1

def test_calculate_level_level1_boundary():
    assert calculate_level(11.9) == 1

def test_calculate_level_level2_start():
    assert calculate_level(12.0) == 2

def test_calculate_level_level2_mid():
    assert calculate_level(18.5) == 2

def test_calculate_level_level3_start():
    assert calculate_level(24.0) == 3

def test_calculate_level_level4_start():
    assert calculate_level(36.0) == 4

def test_calculate_level_level5_start():
    assert calculate_level(48.0) == 5

def test_calculate_level_capped_at_5():
    assert calculate_level(72.0) == 5

# ── get_levelups ─────────────────────────────────────────────────────────────

@pytest.fixture
def enrollment_df():
    return pd.DataFrame({
        "Student Name": ["Alice", "Bob", "Carol", "Dave", "Eve"],
        "Length of Stay": [12.1, 24.3, 13.0, 5.5, 36.0],
        "Center": ["Englewood", "Englewood", "Teaneck", "Teaneck", "Englewood"],
    })

def test_get_levelups_current_month(enrollment_df):
    result = get_levelups(enrollment_df, current=True)
    names = [r["name"] for r in result]
    assert "Alice" in names   # 12.1 → level up this month
    assert "Carol" not in names  # 13.0 → last month
    assert "Bob" in names     # 24.3 → level up this month
    assert "Dave" not in names  # 5.5 → not a level-up month
    assert "Eve" in names     # 36.0 → level up this month

def test_get_levelups_last_month(enrollment_df):
    result = get_levelups(enrollment_df, current=False)
    names = [r["name"] for r in result]
    assert "Carol" in names   # 13.0 → leveled up last month
    assert "Alice" not in names  # 12.1 → this month

def test_get_levelups_includes_old_and_new_level(enrollment_df):
    result = get_levelups(enrollment_df, current=True)
    alice = next(r for r in result if r["name"] == "Alice")
    assert alice["old_level"] == 1
    assert alice["new_level"] == 2

def test_get_levelups_level5_never_appears(enrollment_df):
    """Students already at level 5 (60+ months) should not appear as level-ups."""
    df = pd.DataFrame({
        "Student Name": ["Zara"],
        "Length of Stay": [60.5],  # floor=60, not in threshold sets
        "Center": ["Englewood"],
    })
    assert get_levelups(df, current=True) == []
    assert get_levelups(df, current=False) == []

def test_get_levelups_sorted_by_name(enrollment_df):
    result = get_levelups(enrollment_df, current=True)
    names = [r["name"] for r in result]
    assert names == sorted(names)

# ── get_birthdays ─────────────────────────────────────────────────────────────

@pytest.fixture
def birthday_df():
    return pd.DataFrame({
        "Student Name": ["Alice", "Bob", "Carol"],
        "Birthday": ["1915-04-03", "1913-11-17", "1918-04-22"],
        "Center": ["Englewood", "Englewood", "Teaneck"],
    })

def test_get_birthdays_returns_only_this_month(birthday_df):
    result = get_birthdays(birthday_df, month=4)
    names = [r["name"] for r in result]
    assert "Alice" in names
    assert "Carol" in names
    assert "Bob" not in names  # November birthday

def test_get_birthdays_formats_date(birthday_df):
    result = get_birthdays(birthday_df, month=4)
    alice = next(r for r in result if r["name"] == "Alice")
    assert alice["birthday"] == "Apr 3"

def test_get_birthdays_calculates_age(birthday_df):
    result = get_birthdays(birthday_df, month=4)
    alice = next(r for r in result if r["name"] == "Alice")
    # Alice born 1915, turning 111 in 2026 — age = current_year - birth_year
    from datetime import date
    expected_age = date.today().year - 1915
    assert alice["age"] == expected_age

def test_get_birthdays_sorted_by_day(birthday_df):
    result = get_birthdays(birthday_df, month=4)
    days = [r["_day"] for r in result]
    assert days == sorted(days)

# ── split_by_center ───────────────────────────────────────────────────────────

def test_split_by_center(enrollment_df):
    result = split_by_center(enrollment_df, center_col="Center")
    assert set(result.keys()) == {"Englewood", "Teaneck"}
    assert len(result["Englewood"]) == 3
    assert len(result["Teaneck"]) == 2
```

- [ ] **Step 2: Run tests to confirm they fail**

```bash
cd ~/mathnasium-birthdays-levelups
pytest tests/test_process.py -v
```

Expected: multiple `ModuleNotFoundError` or `ImportError` — `process.py` doesn't exist yet.

- [ ] **Step 3: Create `process.py`**

```python
import math
from datetime import date

import pandas as pd

from config import (
    ENROLLMENT_COL_STUDENT,
    ENROLLMENT_COL_MONTHS,
    ENROLLMENT_COL_CENTER,
    BIRTHDAY_COL_STUDENT,
    BIRTHDAY_COL_BIRTHDAY,
    BIRTHDAY_COL_CENTER,
)

CURRENT_MONTH_THRESHOLDS = {12, 24, 36, 48}
LAST_MONTH_THRESHOLDS = {13, 25, 37, 49}


def calculate_level(months: float) -> int:
    """Return the Mathnasium level for a given length of stay in months.
    Levels: 1 (0-11mo), 2 (12-23mo), 3 (24-35mo), 4 (36-47mo), 5 (48+mo, capped).
    """
    return min(int(months // 12) + 1, 5)


def get_levelups(df: pd.DataFrame, current: bool) -> list[dict]:
    """Return students who leveled up this month (current=True) or last month (current=False).

    Args:
        df: Enrollment report DataFrame for one center.
        current: If True, use current-month thresholds; if False, use last-month thresholds.

    Returns:
        List of dicts with keys: name, old_level, new_level. Sorted by name.
    """
    thresholds = CURRENT_MONTH_THRESHOLDS if current else LAST_MONTH_THRESHOLDS
    results = []
    for _, row in df.iterrows():
        months = float(row[ENROLLMENT_COL_MONTHS])
        if int(months) in thresholds:
            new_level = calculate_level(months)
            results.append({
                "name": row[ENROLLMENT_COL_STUDENT],
                "old_level": new_level - 1,
                "new_level": new_level,
            })
    return sorted(results, key=lambda x: x["name"])


def get_birthdays(df: pd.DataFrame, month: int) -> list[dict]:
    """Return students with birthdays in the given month number (1=Jan, ..., 12=Dec).

    Returns:
        List of dicts with keys: name, birthday (formatted "Apr 3"), age, _day. Sorted by day.
    """
    results = []
    today = date.today()
    for _, row in df.iterrows():
        bday = pd.to_datetime(row[BIRTHDAY_COL_BIRTHDAY])
        if bday.month == month:
            age = today.year - bday.year
            results.append({
                "name": row[BIRTHDAY_COL_STUDENT],
                "birthday": bday.strftime("%b %-d"),
                "age": age,
                "_day": bday.day,
            })
    return sorted(results, key=lambda x: x["_day"])


def split_by_center(df: pd.DataFrame, center_col: str) -> dict:
    """Split a DataFrame into a dict of {center_name: sub-DataFrame}."""
    return {center: group.reset_index(drop=True) for center, group in df.groupby(center_col)}


def process(enrollment_path, birthdays_path, month: int) -> dict:
    """
    Load both reports, split by center, compute level-ups and birthdays.

    Returns:
        {
          "Englewood": {"current_levelups": [...], "last_levelups": [...], "birthdays": [...]},
          "Teaneck":   {"current_levelups": [...], "last_levelups": [...], "birthdays": [...]},
        }
    """
    enrollment_df = pd.read_excel(enrollment_path)
    birthdays_df = pd.read_excel(birthdays_path)

    enrollment_by_center = split_by_center(enrollment_df, ENROLLMENT_COL_CENTER)
    birthdays_by_center = split_by_center(birthdays_df, BIRTHDAY_COL_CENTER)

    result = {}
    for center_name in enrollment_by_center:
        enroll = enrollment_by_center[center_name]
        bdays = birthdays_by_center.get(center_name, pd.DataFrame(columns=birthdays_df.columns))
        result[center_name] = {
            "current_levelups": get_levelups(enroll, current=True),
            "last_levelups": get_levelups(enroll, current=False),
            "birthdays": get_birthdays(bdays, month=month),
        }
    return result
```

- [ ] **Step 4: Run tests to confirm they pass**

```bash
pytest tests/test_process.py -v
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add process.py tests/
git commit -m "feat: add process.py with level-up and birthday logic (TDD)"
```

---

### Task 6: Build `deliver.py`

**Files:**
- Create: `deliver.py`

- [ ] **Step 1: Create `deliver.py`**

```python
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from config import CENTERS, CC_RECIPIENT


def _table_rows(rows: list[dict], cols: list[str]) -> str:
    if not rows:
        return "<tr><td colspan='{}' style='color:#888;font-style:italic;'>None this month.</td></tr>".format(len(cols))
    return "".join(
        "<tr>" + "".join(f"<td style='padding:6px 12px;border:1px solid #ddd;'>{row.get(c, '')}</td>" for c in cols) + "</tr>"
        for row in rows
    )


def _table(headers: list[str], rows: list[dict], cols: list[str]) -> str:
    header_html = "".join(
        f"<th style='padding:6px 12px;border:1px solid #ddd;background:#f5f5f5;text-align:left;'>{h}</th>"
        for h in headers
    )
    return (
        "<table style='border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;margin-bottom:8px;'>"
        f"<tr>{header_html}</tr>"
        f"{_table_rows(rows, cols)}"
        "</table>"
    )


def build_html(center_name: str, month_label: str,
               birthdays: list[dict],
               current_levelups: list[dict],
               last_levelups: list[dict]) -> str:
    """Build the full HTML email body."""

    bday_rows = [{"Name": r["name"], "Birthday": r["birthday"], "Age": r["age"]} for r in birthdays]
    bday_table = _table(["Name", "Birthday", "Age"], bday_rows, ["Name", "Birthday", "Age"])

    def levelup_table(rows):
        display = [{"Name": r["name"], "Old Level": f"Level {r['old_level']}", "New Level": f"Level {r['new_level']}"} for r in rows]
        return _table(["Name", "Old Level", "New Level"], display, ["Name", "Old Level", "New Level"])

    divider = "<hr style='border:none;border-top:2px solid #ccc;margin:24px 0;'>"

    return f"""
<html><body style="font-family:Arial,sans-serif;font-size:14px;max-width:700px;margin:0 auto;padding:20px;">

<p>Hi {center_name} Center Directors,</p>

<p>Please find this month's Birthdays &amp; Level Ups below. As a reminder:</p>
<ul>
  <li>Please have an instructor <strong>update the student binders</strong> for any level ups.</li>
  <li>Please <strong>add this month's birthdays to the whiteboard</strong> so we can celebrate with our students!</li>
</ul>

{divider}

<h2 style="color:#333;">🎂 Birthdays &mdash; {month_label}</h2>
{bday_table}

{divider}

<h2 style="color:#333;">⭐ Level Ups This Month</h2>
{levelup_table(current_levelups)}

{divider}

<h2 style="color:#333;">✅ Last Month's Level Ups &mdash; Please Confirm Binders Were Updated</h2>
<p style="color:#555;">The following students leveled up last month. Please confirm their binders have been updated.</p>
{levelup_table(last_levelups)}

{divider}

<p style="color:#999;font-size:12px;"><em>This email was generated automatically. Questions? Contact matt.diamond@mathnasium.com.</em></p>

</body></html>
"""


def send_email(center_name: str, month_label: str, html: str) -> None:
    """Send the HTML email for one center via Gmail SMTP."""
    smtp_user = os.environ.get("SMTP_USER")
    smtp_password = os.environ.get("SMTP_PASSWORD")

    recipient = CENTERS[center_name]["recipient"]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Mathnasium {center_name} \u2014 Birthdays & Level Ups: {month_label}"
    msg["From"] = smtp_user
    msg["To"] = recipient
    msg["Cc"] = CC_RECIPIENT

    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, [recipient, CC_RECIPIENT], msg.as_string())


def deliver(center_data: dict, month_label: str) -> None:
    """Send one email per center.

    Args:
        center_data: output of process.process() — {center_name: {current_levelups, last_levelups, birthdays}}
        month_label: human-readable label e.g. "April 2026"
    """
    for center_name, data in center_data.items():
        html = build_html(
            center_name=center_name,
            month_label=month_label,
            birthdays=data["birthdays"],
            current_levelups=data["current_levelups"],
            last_levelups=data["last_levelups"],
        )
        send_email(center_name, month_label, html)
        print(f"Sent email for {center_name}")
```

- [ ] **Step 2: Do a quick local HTML preview**

```bash
python -c "
from deliver import build_html
html = build_html(
    'Englewood', 'April 2026',
    birthdays=[{'name': 'Alex Johnson', 'birthday': 'Apr 3', 'age': 9}],
    current_levelups=[{'name': 'Jordan Kim', 'old_level': 1, 'new_level': 2}],
    last_levelups=[{'name': 'Casey Lee', 'old_level': 2, 'new_level': 3}],
)
with open('/tmp/preview.html', 'w') as f:
    f.write(html)
print('Open /tmp/preview.html in a browser to review.')
"
open /tmp/preview.html
```

Review the email in the browser. Confirm layout, wording, and "None this month." fallback looks right.

- [ ] **Step 3: Commit**

```bash
git add deliver.py
git commit -m "feat: add deliver.py with HTML builder and SMTP sender"
```

---

### Task 7: Build `run_log.py` and `main.py`

**Files:**
- Create: `run_log.py`
- Create: `main.py`

- [ ] **Step 1: Create `run_log.py`**

```python
import json
from datetime import datetime, timezone
from pathlib import Path

LOG_PATH = Path(__file__).parent / "run_log.json"


def read_log() -> list:
    if not LOG_PATH.exists():
        return []
    with open(LOG_PATH) as f:
        return json.load(f)


def write_log(month: str, success: bool, error: str = "") -> None:
    log = read_log()
    log.append({
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "month": month,
        "success": success,
        "error": error,
    })
    with open(LOG_PATH, "w") as f:
        json.dump(log, f, indent=2)
```

- [ ] **Step 2: Create `main.py`**

```python
import argparse
import traceback
from datetime import date

from config import INPUT_DIR
from download import download_reports
from process import process
from deliver import deliver
from run_log import write_log


def get_month_label(month_str: str) -> str:
    """Convert 'YYYY-MM' to 'Month YYYY' e.g. '2026-04' → 'April 2026'."""
    year, month = month_str.split("-")
    month_name = date(int(year), int(month), 1).strftime("%B")
    return f"{month_name} {year}"


def main():
    parser = argparse.ArgumentParser(description="Run Birthdays & Level Ups automation.")
    parser.add_argument(
        "--month",
        default=date.today().strftime("%Y-%m"),
        help="Month to process in YYYY-MM format (default: current month)",
    )
    args = parser.parse_args()
    month_str = args.month
    month_num = int(month_str.split("-")[1])
    month_label = get_month_label(month_str)

    print(f"Running Birthdays & Level Ups for {month_label}")

    try:
        print("Downloading reports...")
        paths = download_reports(month_str)

        print("Processing data...")
        center_data = process(paths["enrollment"], paths["birthdays"], month=month_num)

        print("Sending emails...")
        deliver(center_data, month_label)

        write_log(month_str, success=True)
        print("Done.")

    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"ERROR: {e}")
        print(error_msg)
        write_log(month_str, success=False, error=error_msg)
        raise


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Commit**

```bash
git add run_log.py main.py
git commit -m "feat: add main.py orchestrator and run_log.py"
```

---

### Task 8: GitHub Actions workflow

**Files:**
- Create: `.github/workflows/monthly_birthdays_levelups.yml`

- [ ] **Step 1: Create the workflow file**

```bash
mkdir -p .github/workflows
```

Create `.github/workflows/monthly_birthdays_levelups.yml`:

```yaml
name: Monthly Birthdays & Level Ups

on:
  schedule:
    - cron: "0 15 1 * *"   # 1st of month, 11am ET (3pm UTC)
  workflow_dispatch:         # manual trigger for testing

jobs:
  run:
    runs-on: ubuntu-latest

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
        run: python main.py

      - name: Commit run log
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add run_log.json
          git diff --cached --quiet || git commit -m "chore: update run log [skip ci]"
          git push
```

- [ ] **Step 2: Commit and push**

```bash
git add .github/
git commit -m "feat: add GitHub Actions workflow"
git push
```

- [ ] **Step 3: Trigger a manual test run**

Go to the GitHub repo → Actions tab → "Monthly Birthdays & Level Ups" → Run workflow → Run workflow.

Watch the run complete. If it fails, click the failing step for logs.

- [ ] **Step 4: Confirm emails arrived**

Check matt.diamond@mathnasium.com, englewood@mathnasium.com, and teaneck@mathnasium.com inboxes. Verify:
- Both center emails arrived
- Subject line is correct
- Birthday table is populated
- Level-up tables are populated (or show "None this month." if no level-ups)
- Formatting is clean and printable

- [ ] **Step 5: Confirm `run_log.json` was updated**

Pull latest and check:

```bash
git pull
cat run_log.json
```

Expected: one entry with `"success": true` and the current month.

---

## Appendix: Column name verification checklist

After downloading real Radius files in Tasks 2 and 3, verify these column names and update `config.py` if they differ:

| Constant in `config.py` | Default value | Actual value from file |
|---|---|---|
| `ENROLLMENT_COL_STUDENT` | `"Student Name"` | _(fill in)_ |
| `ENROLLMENT_COL_MONTHS` | `"Length of Stay"` | _(fill in)_ |
| `ENROLLMENT_COL_CENTER` | `"Center"` | _(fill in)_ |
| `BIRTHDAY_COL_STUDENT` | `"Student Name"` | _(fill in)_ |
| `BIRTHDAY_COL_BIRTHDAY` | `"Birthday"` | _(fill in)_ |
| `BIRTHDAY_COL_CENTER` | `"Center"` | _(fill in)_ |
