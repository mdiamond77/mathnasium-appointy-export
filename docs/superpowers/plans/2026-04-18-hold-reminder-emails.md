# Hold Reminder Emails Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Email each center director a per-center list of students coming off hold this month, automatically on the last Monday, Tuesday, and Thursday of each month.

**Architecture:** New repo `mathnasium-hold-reminders` with the same 5-module structure as `mathnasium-birthdays-levelups` (download → process → deliver → log). A cron fires Mon/Tue/Thu; `main.py` checks whether today is in the last 7 days of the month and exits early if not. Per-center HTML emails go to each CD with guardian contact info and a reminder to add new holds.

**Tech Stack:** Python 3.11, Playwright (Radius scraping), pandas + openpyxl (Excel), smtplib SSL (Gmail), GitHub Actions (cron + secrets), Flask dashboard (run_log integration)

---

## File Map

| File | Create/Modify | Purpose |
|------|--------------|---------|
| `config.py` | Create | Constants: URLs, centers, recipients, column names |
| `process.py` | Create | Filter holds by end-date month/year, split by center |
| `deliver.py` | Create | HTML email builder + Gmail SMTP sender |
| `download.py` | Create | Playwright: login → HoldsReport → export Excel |
| `main.py` | Create | Orchestrator with last-7-days date check |
| `run_log.py` | Create | Append run result to run_log.json |
| `requirements.txt` | Create | Python dependencies |
| `tests/__init__.py` | Create | Empty — makes tests/ a package |
| `tests/test_process.py` | Create | Unit tests for filter_holds, split_by_center |
| `tests/test_deliver.py` | Create | Unit tests for build_html |
| `tests/test_main.py` | Create | Unit tests for is_last_week_of_month |
| `.github/workflows/hold_reminders.yml` | Create | Mon/Tue/Thu cron + workflow_dispatch |
| `automation-dashboard/server.py` | Modify | Add SCRIPTS + REPORTS entries |
| `automation-dashboard/future_projects.json` | Modify | Set hold-reminder-emails status to "Live" |

---

## Task 1: Project scaffolding

**Files:**
- Create: `~/mathnasium-hold-reminders/config.py`
- Create: `~/mathnasium-hold-reminders/requirements.txt`
- Create: `~/mathnasium-hold-reminders/tests/__init__.py`

- [ ] **Step 1: Create GitHub repo and clone locally**

```bash
gh repo create mdmathnasiums/mathnasium-hold-reminders --private --clone
cd ~/mathnasium-hold-reminders
```

- [ ] **Step 2: Create `config.py`**

```python
import os
from pathlib import Path

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"

RADIUS_LOGIN_URL = "https://radius.mathnasium.com"
HOLDS_REPORT_URL = "https://radius.mathnasium.com/Holds/HoldsReport"

CENTERS = {
    "Englewood": {"radius_id": "2428", "recipient": "englewood@mathnasium.com"},
    "Teaneck":   {"radius_id": "2871", "recipient": "teaneck@mathnasium.com"},
}
CC_RECIPIENT = "matt.diamond@mathnasium.com"

# Column names verified from real Radius Holds Report export (4/18/2026)
COL_STUDENT_NAME   = "Student Name"
COL_HOLD_END_DATE  = "Hold End Date"
COL_GUARDIAN_NAME  = "Guardian Name"
COL_GUARDIAN_PHONE = "Guardian Phone"
COL_GUARDIAN_EMAIL = "Guardian Email"
COL_CENTER         = "Center Name"
```

- [ ] **Step 3: Create `requirements.txt`**

```
playwright==1.44.0
pandas==2.2.2
openpyxl==3.1.2
pytest==8.2.0
```

- [ ] **Step 4: Create `tests/__init__.py`**

Empty file — its presence allows pytest to find imports from the project root.

```bash
mkdir tests && touch tests/__init__.py
```

- [ ] **Step 5: Install dependencies**

```bash
pip install -r requirements.txt
playwright install chromium
```

- [ ] **Step 6: Initial commit**

```bash
git add config.py requirements.txt tests/__init__.py
git commit -m "feat: initial scaffolding"
git push -u origin main
```

---

## Task 2: process.py

**Files:**
- Create: `~/mathnasium-hold-reminders/process.py`
- Create: `~/mathnasium-hold-reminders/tests/test_process.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_process.py`:

```python
import pandas as pd
import pytest
from process import filter_holds, split_by_center


@pytest.fixture
def holds_df():
    return pd.DataFrame({
        "Student Name":   ["Alice Smith",    "Bob Jones",     "Carol White",   "Dave Brown"],
        "Hold End Date":  ["4/30/2026",      "4/15/2026",     "5/31/2026",     "4/30/2026"],
        "Guardian Name":  ["Amy Smith",      "Bill Jones",    "Clara White",   "Dan Brown"],
        "Guardian Phone": ["(201) 555-0001", "(201) 555-0002","(201) 555-0003","(201) 555-0004"],
        "Guardian Email": ["amy@ex.com",     "bill@ex.com",   "clara@ex.com",  "dan@ex.com"],
        "Center Name":    ["Englewood",      "Teaneck",       "Englewood",     "Teaneck"],
    })


# ── filter_holds ──────────────────────────────────────────────────────────────

def test_filter_holds_returns_current_month(holds_df):
    result = filter_holds(holds_df, month=4, year=2026)
    names = [r["name"] for r in result]
    assert "Alice Smith" in names
    assert "Bob Jones" in names
    assert "Dave Brown" in names
    assert "Carol White" not in names  # May end date


def test_filter_holds_excludes_wrong_year(holds_df):
    result = filter_holds(holds_df, month=4, year=2025)
    assert result == []


def test_filter_holds_formats_end_date(holds_df):
    result = filter_holds(holds_df, month=4, year=2026)
    alice = next(r for r in result if r["name"] == "Alice Smith")
    assert alice["hold_end_date"] == "4/30/2026"


def test_filter_holds_includes_contact_info(holds_df):
    result = filter_holds(holds_df, month=4, year=2026)
    alice = next(r for r in result if r["name"] == "Alice Smith")
    assert alice["guardian_name"] == "Amy Smith"
    assert alice["guardian_phone"] == "(201) 555-0001"
    assert alice["guardian_email"] == "amy@ex.com"


def test_filter_holds_sorted_by_center_then_name(holds_df):
    result = filter_holds(holds_df, month=4, year=2026)
    keys = [(r["center"], r["name"]) for r in result]
    assert keys == sorted(keys)


# ── split_by_center ───────────────────────────────────────────────────────────

def test_split_by_center_all_centers_present_when_empty():
    result = split_by_center([])
    assert "Englewood" in result
    assert "Teaneck" in result


def test_split_by_center_splits_correctly(holds_df):
    holds = filter_holds(holds_df, month=4, year=2026)
    result = split_by_center(holds)
    assert len(result["Englewood"]) == 1   # Alice
    assert len(result["Teaneck"]) == 2     # Bob, Dave


def test_split_by_center_empty_center_is_empty_list(holds_df):
    # Filter for a month with only Englewood holds
    df = holds_df[holds_df["Center Name"] == "Englewood"].copy()
    df["Hold End Date"] = "4/30/2026"
    holds = filter_holds(df, month=4, year=2026)
    result = split_by_center(holds)
    assert result["Teaneck"] == []
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
cd ~/mathnasium-hold-reminders
pytest tests/test_process.py -v
```

Expected: `ImportError: No module named 'process'`

- [ ] **Step 3: Write `process.py`**

```python
import pandas as pd

from config import (
    COL_STUDENT_NAME,
    COL_HOLD_END_DATE,
    COL_GUARDIAN_NAME,
    COL_GUARDIAN_PHONE,
    COL_GUARDIAN_EMAIL,
    COL_CENTER,
    CENTERS,
)


def filter_holds(df: pd.DataFrame, month: int, year: int) -> list[dict]:
    """Return holds where Hold End Date is in the given month and year.

    Returns list of dicts: {name, hold_end_date, guardian_name, guardian_phone,
    guardian_email, center}. Sorted by center then name.
    """
    results = []
    for _, row in df.iterrows():
        end_date = pd.to_datetime(row[COL_HOLD_END_DATE])
        if end_date.month == month and end_date.year == year:
            results.append({
                "name": row[COL_STUDENT_NAME],
                "hold_end_date": "{}/{}/{}".format(
                    end_date.month, end_date.day, end_date.year
                ),
                "guardian_name": row[COL_GUARDIAN_NAME],
                "guardian_phone": row[COL_GUARDIAN_PHONE],
                "guardian_email": row[COL_GUARDIAN_EMAIL],
                "center": row[COL_CENTER],
            })
    return sorted(results, key=lambda x: (x["center"], x["name"]))


def split_by_center(holds: list[dict]) -> dict:
    """Split hold list into {center_name: [hold dicts]}.

    All CENTERS keys are always present (empty list if no holds for that center).
    """
    result = {center: [] for center in CENTERS}
    for hold in holds:
        center = hold["center"]
        if center in result:
            result[center].append(hold)
    return result


def process(holds_path, month: int, year: int) -> dict:
    """Load Holds Report Excel, filter by month/year, split by center.

    Returns {center_name: [hold dicts]} for all centers in CENTERS.
    Raises ValueError if the report is empty.
    """
    df = pd.read_excel(holds_path)
    if df.empty:
        raise ValueError("Holds report is empty — check Radius export and re-run.")
    holds = filter_holds(df, month, year)
    return split_by_center(holds)
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pytest tests/test_process.py -v
```

Expected: all 9 tests PASS

- [ ] **Step 5: Commit**

```bash
git add process.py tests/test_process.py
git commit -m "feat: add process module with hold filtering and center split"
```

---

## Task 3: deliver.py

**Files:**
- Create: `~/mathnasium-hold-reminders/deliver.py`
- Create: `~/mathnasium-hold-reminders/tests/test_deliver.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_deliver.py`:

```python
from deliver import build_html


SAMPLE_HOLDS = [
    {
        "name": "Alice Smith",
        "hold_end_date": "4/30/2026",
        "guardian_name": "Amy Smith",
        "guardian_phone": "(201) 555-0001",
        "guardian_email": "amy@example.com",
    }
]


def test_build_html_contains_student_name():
    html = build_html("Teaneck", "April 2026", SAMPLE_HOLDS)
    assert "Alice Smith" in html


def test_build_html_contains_hold_end_date():
    html = build_html("Teaneck", "April 2026", SAMPLE_HOLDS)
    assert "4/30/2026" in html


def test_build_html_contains_guardian_info():
    html = build_html("Teaneck", "April 2026", SAMPLE_HOLDS)
    assert "Amy Smith" in html
    assert "(201) 555-0001" in html
    assert "amy@example.com" in html


def test_build_html_contains_center_name():
    html = build_html("Englewood", "April 2026", [])
    assert "Englewood" in html


def test_build_html_contains_hold_reminder():
    html = build_html("Teaneck", "April 2026", [])
    assert "going on hold" in html


def test_build_html_is_valid_html():
    html = build_html("Teaneck", "April 2026", SAMPLE_HOLDS)
    assert html.strip().startswith("<html")
    assert "</html>" in html
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_deliver.py -v
```

Expected: `ImportError: No module named 'deliver'`

- [ ] **Step 3: Write `deliver.py`**

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
        "<tr>" + "".join(
            "<td style='padding:6px 12px;border:1px solid #ddd;'>{}</td>".format(row.get(c, ""))
            for c in cols
        ) + "</tr>"
        for row in rows
    )


def _table(headers: list[str], rows: list[dict], cols: list[str]) -> str:
    header_html = "".join(
        "<th style='padding:6px 12px;border:1px solid #ddd;background:#f5f5f5;text-align:left;'>{}</th>".format(h)
        for h in headers
    )
    return (
        "<table style='border-collapse:collapse;font-family:Arial,sans-serif;"
        "font-size:14px;margin-bottom:8px;'>"
        "<tr>{}</tr>"
        "{}"
        "</table>"
    ).format(header_html, _table_rows(rows, cols))


def build_html(center_name: str, month_label: str, holds: list[dict]) -> str:
    """Build the full HTML email body for one center."""
    display_rows = [
        {
            "Student": h["name"],
            "Hold End Date": h["hold_end_date"],
            "Guardian": h["guardian_name"],
            "Phone": h["guardian_phone"],
            "Email": h["guardian_email"],
        }
        for h in holds
    ]
    table_html = _table(
        ["Student", "Hold End Date", "Guardian", "Phone", "Email"],
        display_rows,
        ["Student", "Hold End Date", "Guardian", "Phone", "Email"],
    )
    divider = "<hr style='border:none;border-top:2px solid #ccc;margin:24px 0;'>"

    return (
        "<html><body style='font-family:Arial,sans-serif;font-size:14px;"
        "max-width:700px;margin:0 auto;padding:20px;'>"
        "<p>Hi {center} Team,</p>"
        "<p>The following students are scheduled to come off hold this month and return "
        "to billing for the next cycle. Please confirm each one is still correct before "
        "the end of the month.</p>"
        "{divider}"
        "{table}"
        "{divider}"
        "<p><strong>Reminder:</strong> Please also add any students to next month's holds "
        "if they have notified you that they are going on hold.</p>"
        "<p style='color:#999;font-size:12px;'><em>This email sends automatically on the "
        "last Monday, Tuesday, and Thursday of each month. "
        "Questions? Contact matt.diamond@mathnasium.com.</em></p>"
        "</body></html>"
    ).format(center=center_name, divider=divider, table=table_html)


def send_email(center_name: str, month_label: str, html: str) -> None:
    """Send the HTML email for one center via Gmail SMTP."""
    smtp_user = os.environ.get("SMTP_USER")
    smtp_password = os.environ.get("SMTP_PASSWORD")
    recipient = CENTERS[center_name]["recipient"]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Hold Reminders \u2014 {} \u2014 {}".format(center_name, month_label)
    msg["From"] = smtp_user
    msg["To"] = recipient
    msg["Cc"] = CC_RECIPIENT

    msg.attach(MIMEText(html, "html"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, [recipient, CC_RECIPIENT], msg.as_string())


def deliver(center_data: dict, month_label: str) -> None:
    """Send one email per center. Centers with no holds are skipped silently."""
    for center_name, holds in center_data.items():
        if not holds:
            print("No holds for {} this month, skipping.".format(center_name))
            continue
        html = build_html(center_name, month_label, holds)
        send_email(center_name, month_label, html)
        print("Sent email for {}".format(center_name))
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pytest tests/test_deliver.py -v
```

Expected: all 6 tests PASS

- [ ] **Step 5: Commit**

```bash
git add deliver.py tests/test_deliver.py
git commit -m "feat: add deliver module with HTML email builder and SMTP sender"
```

---

## Task 4: download.py

**Files:**
- Create: `~/mathnasium-hold-reminders/download.py`

No unit tests — this module wraps Playwright and requires a live Radius session. It is tested end-to-end via GitHub Actions.

- [ ] **Step 1: Write `download.py`**

```python
import json
import os
from pathlib import Path

from playwright.sync_api import sync_playwright

from config import INPUT_DIR, RADIUS_LOGIN_URL, HOLDS_REPORT_URL, CENTERS

CENTER_IDS = [c["radius_id"] for c in CENTERS.values()]


def _login(page) -> None:
    """Log into Radius."""
    page.goto(RADIUS_LOGIN_URL)
    page.fill("#UserName", os.environ.get("RADIUS_USERNAME"))
    page.fill("#Password", os.environ.get("RADIUS_PASSWORD"))
    page.click("#login")
    page.wait_for_load_state("networkidle")


def _select_all_centers(page) -> None:
    """Select both Englewood and Teaneck in the center Kendo MultiSelect."""
    page.evaluate(
        """(ids) => {
            const ms = $('#AllCenterListMultiSelect').data('kendoMultiSelect');
            ms.value(ids);
            ms.trigger('change');
        }""",
        json.dumps(CENTER_IDS),
    )


def download_holds_report(page, month_label: str) -> Path:
    """Navigate to HoldsReport, select both centers, export to Excel."""
    out_path = INPUT_DIR / "Holds_{}.xlsx".format(month_label)

    page.goto(HOLDS_REPORT_URL)
    page.wait_for_load_state("networkidle")

    _select_all_centers(page)

    page.click("#btnsearch")
    page.wait_for_load_state("networkidle")

    with page.expect_download() as dl:
        page.click("#btnExport")
    dl.value.save_as(out_path)

    return out_path


def download_reports(month_label: str) -> dict:
    """Download the Holds Report. Returns {"holds": Path}."""
    INPUT_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        with browser:
            context = browser.new_context(accept_downloads=True)
            with context:
                page = context.new_page()
                _login(page)
                holds_path = download_holds_report(page, month_label)

    return {"holds": holds_path}
```

**Note on selectors:** `#AllCenterListMultiSelect`, `#btnsearch`, and `#btnExport` are verified on other Radius reports and are expected to be consistent. If the Holds Report uses different IDs, the first GitHub Actions run will error with a timeout — check the Actions log, use `discover_selectors.py` from `mathnasium-birthdays-levelups` as a reference to find the correct IDs, and update accordingly.

- [ ] **Step 2: Commit**

```bash
git add download.py
git commit -m "feat: add download module for Radius Holds Report"
```

---

## Task 5: main.py and run_log.py

**Files:**
- Create: `~/mathnasium-hold-reminders/main.py`
- Create: `~/mathnasium-hold-reminders/run_log.py`
- Create: `~/mathnasium-hold-reminders/tests/test_main.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_main.py`:

```python
from datetime import date
from main import is_last_week_of_month


def test_last_day_of_april():
    assert is_last_week_of_month(date(2026, 4, 30)) is True


def test_day_24_of_april_is_in_last_week():
    # April has 30 days. 30 - 24 = 6 < 7.
    assert is_last_week_of_month(date(2026, 4, 24)) is True


def test_day_23_of_april_is_not_in_last_week():
    # 30 - 23 = 7, not < 7.
    assert is_last_week_of_month(date(2026, 4, 23)) is False


def test_first_day_of_month():
    assert is_last_week_of_month(date(2026, 4, 1)) is False


def test_last_day_of_march():
    # March has 31 days.
    assert is_last_week_of_month(date(2026, 3, 31)) is True


def test_day_25_of_march_is_in_last_week():
    # 31 - 25 = 6 < 7.
    assert is_last_week_of_month(date(2026, 3, 25)) is True


def test_day_24_of_march_is_not_in_last_week():
    # 31 - 24 = 7, not < 7.
    assert is_last_week_of_month(date(2026, 3, 24)) is False
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_main.py -v
```

Expected: `ImportError: No module named 'main'`

- [ ] **Step 3: Write `run_log.py`**

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


def write_log(month: str, success: bool, trigger: str = "auto", error: str = "") -> None:
    log = read_log()
    log.append({
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "trigger": trigger,
        "month": month,
        "success": success,
        "error": error,
    })
    with open(LOG_PATH, "w") as f:
        json.dump(log, f, indent=2)
```

- [ ] **Step 4: Write `main.py`**

```python
import argparse
import traceback
from calendar import monthrange
from datetime import date

from download import download_reports
from process import process
from deliver import deliver
from run_log import write_log


def is_last_week_of_month(today: date = None) -> bool:
    """Return True if today is within the last 7 days of the month."""
    if today is None:
        today = date.today()
    last_day = monthrange(today.year, today.month)[1]
    return (last_day - today.day) < 7


def get_month_label(month_str: str) -> str:
    """Convert 'YYYY-MM' to 'Month YYYY', e.g. '2026-04' -> 'April 2026'."""
    year, month = month_str.split("-")
    month_name = date(int(year), int(month), 1).strftime("%B")
    return "{} {}".format(month_name, year)


def main():
    parser = argparse.ArgumentParser(description="Run Hold Reminder Emails automation.")
    parser.add_argument(
        "--trigger",
        default="auto",
        choices=["auto", "manual"],
        help="How this run was triggered (default: auto)",
    )
    parser.add_argument(
        "--month",
        default=date.today().strftime("%Y-%m"),
        help="Month to process in YYYY-MM format (default: current month)",
    )
    args = parser.parse_args()

    month_str = args.month
    month_label = get_month_label(month_str)
    year = int(month_str.split("-")[0])
    month_num = int(month_str.split("-")[1])

    if args.trigger == "auto" and not is_last_week_of_month():
        print("Not in last 7 days of month — skipping.")
        return

    print("Running Hold Reminders for {}".format(month_label))

    try:
        print("Downloading holds report...")
        paths = download_reports(month_label)

        print("Processing data...")
        center_data = process(paths["holds"], month=month_num, year=year)

        print("Sending emails...")
        deliver(center_data, month_label)

        write_log(month_str, success=True, trigger=args.trigger)
        print("Done.")

    except Exception as e:
        error_msg = traceback.format_exc()
        print("ERROR: {}".format(e))
        print(error_msg)
        write_log(month_str, success=False, trigger=args.trigger, error=error_msg)
        raise


if __name__ == "__main__":
    main()
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
pytest tests/test_main.py -v
```

Expected: all 7 tests PASS

- [ ] **Step 6: Run full test suite**

```bash
pytest -v
```

Expected: all tests PASS (process + deliver + main)

- [ ] **Step 7: Commit**

```bash
git add main.py run_log.py tests/test_main.py
git commit -m "feat: add main orchestrator with last-week-of-month check and run logging"
```

---

## Task 6: GitHub Actions workflow

**Files:**
- Create: `~/mathnasium-hold-reminders/.github/workflows/hold_reminders.yml`

- [ ] **Step 1: Create workflow directory and file**

```bash
mkdir -p .github/workflows
```

Create `.github/workflows/hold_reminders.yml`:

```yaml
name: Hold Reminder Emails

on:
  schedule:
    - cron: "0 15 * * 1,2,4"   # Mon/Tue/Thu at 11am ET (3pm UTC)
  workflow_dispatch:             # manual trigger for testing

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
        run: python main.py --trigger auto

      - name: Commit run log
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add run_log.json
          git diff --cached --quiet || git commit -m "chore: update run log [skip ci]"
          git push
```

- [ ] **Step 2: Add GitHub secrets**

The repo needs these secrets (same values as `mathnasium-birthdays-levelups`):

Go to: `https://github.com/mdmathnasiums/mathnasium-hold-reminders/settings/secrets/actions`

Add:
- `RADIUS_USERNAME`
- `RADIUS_PASSWORD`
- `SMTP_USER`
- `SMTP_PASSWORD`

- [ ] **Step 3: Commit and push**

```bash
git add .github/workflows/hold_reminders.yml
git commit -m "feat: add GitHub Actions workflow for monthly hold reminders"
git push
```

- [ ] **Step 4: Trigger a manual test run**

Go to: `https://github.com/mdmathnasiums/mathnasium-hold-reminders/actions`

Click `Hold Reminder Emails` → `Run workflow` → `Run workflow`.

Expected: workflow completes successfully, emails arrive in each center director's inbox (or Matt's if no holds for the current month match).

**If it fails:** Check the Actions log. Common issues:
- Selector not found (`#AllCenterListMultiSelect`): the Holds Report uses a different center selector. Add a debug step to print `document.querySelectorAll('[id*=Center]').map(e=>e.id)` and update `download.py` with the correct ID.
- Column name mismatch: add `print(df.columns.tolist())` to `process.py` after `pd.read_excel()`, re-run, update `config.py` with the real column names.

---

## Task 7: Dashboard integration

**Files:**
- Modify: `~/automation-dashboard/server.py`
- Modify: `~/automation-dashboard/future_projects.json`

- [ ] **Step 1: Add entry to SCRIPTS in `server.py`**

Open `/Users/mattdiamond/automation-dashboard/server.py`.

After the `SCRIPTS["birthdays-levelups"]` block (around line 64), add:

```python
SCRIPTS["hold-reminders"] = {
    "name": "Hold Reminder Emails",
    "description": "Emails each center director a list of students coming off hold this month.",
    "command": [PYTHON, "main.py", "--trigger", "manual"],
    "cwd": "/Users/mattdiamond/mathnasium-hold-reminders",
    "icon": "⏸️",
    "category": "Mathnasium",
    "hidden": True,
}
```

- [ ] **Step 2: Add entry to REPORTS in `server.py`**

In the `REPORTS` list (around line 67), add after the `birthdays-levelups` entry:

```python
{
    "id": "hold-reminders",
    "name": "Hold Reminder Emails",
    "schedule": "Last Mon/Tue/Thu of month",
    "script_id": "hold-reminders",
    "run_log_path": "/Users/mattdiamond/mathnasium-hold-reminders/run_log.json",
},
```

- [ ] **Step 3: Update `future_projects.json`**

Open `/Users/mattdiamond/automation-dashboard/future_projects.json`.

Find the entry with `"id": "hold-reminder-emails"` and change `"status": "Planned"` to `"status": "Live"`.

- [ ] **Step 4: Restart dashboard server**

```bash
launchctl stop com.mathnasium.dashboard && sleep 2 && launchctl start com.mathnasium.dashboard
```

- [ ] **Step 5: Verify dashboard shows Hold Reminder Emails in Live Automations**

Open `http://localhost:8080` and confirm the Hold Reminder Emails row appears in the Live Automations table.

- [ ] **Step 6: Commit dashboard changes**

```bash
cd /Users/mattdiamond/automation-dashboard
git add server.py future_projects.json
git commit -m "feat: add hold-reminders to dashboard live automations"
git push
```
