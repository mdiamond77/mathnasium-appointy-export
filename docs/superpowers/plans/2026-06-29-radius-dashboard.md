# Radius Dashboard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a GitHub Actions automation that scrapes two Radius reports, transforms the data, and writes a Google Sheet with 6 tabs covering a lead funnel and enrollment/MRR report for two centers.

**Architecture:** Playwright exports CSVs from Radius for each center and report type. Pandas transforms the data (filtering event leads, backfilling assessed status, aggregating monthly metrics). A gspread client rewrites all 6 tabs of a shared Google Sheet on each run.

**Tech Stack:** Python 3.11, Playwright, pandas, gspread, google-auth, python-dateutil, GitHub Actions

---

## File Map

```
radius-dashboard/
├── main.py                    # Orchestrator; CLI entry point
├── config.py                  # All constants, column names, selectors
├── radius_export.py           # Playwright: login + export both reports per center
├── discover_selectors.py      # One-time tool: find Radius report URLs + selectors
├── transform.py               # pandas: filter, join, aggregate, format for sheets
├── sheets.py                  # Google Sheets API: authenticate + write all 6 tabs
├── requirements.txt
├── .github/workflows/run.yml
└── tests/
    └── test_transform.py
```

---

## Task 1: Project scaffolding

**Files:**
- Create: `requirements.txt`
- Create: `config.py`
- Create: `tests/__init__.py`
- Create: `tests/test_transform.py` (empty stub)

- [ ] **Step 1: Create the project directory and initialize git**

```bash
mkdir ~/radius-dashboard && cd ~/radius-dashboard
git init
echo ".env\n__pycache__/\n*.pyc\n.playwright/\n*.csv" > .gitignore
```

- [ ] **Step 2: Create requirements.txt**

```
playwright==1.44.0
pandas==2.2.2
gspread==6.1.2
google-auth==2.30.0
python-dateutil==2.9.0
pytest==8.2.2
```

- [ ] **Step 3: Create config.py**

```python
import os
from datetime import date

CENTERS = ['Teaneck', 'Englewood']

RADIUS_URL = 'https://radius.mathnasium.com'
RADIUS_USERNAME = os.environ['RADIUS_USERNAME']
RADIUS_PASSWORD = os.environ['RADIUS_PASSWORD']

# Filled in after running discover_selectors.py
LEADS_REPORT_URL = ''
ENROLLMENT_REPORT_URL = ''

# Selectors — filled in after running discover_selectors.py
LEADS_CENTER_SELECTOR = ''
LEADS_DATE_START_SELECTOR = ''
LEADS_DATE_END_SELECTOR = ''
LEADS_SEARCH_SELECTOR = ''
LEADS_EXPORT_SELECTOR = ''

ENROLL_CENTER_SELECTOR = ''
ENROLL_DATE_START_SELECTOR = ''
ENROLL_DATE_END_SELECTOR = ''
ENROLL_SEARCH_SELECTOR = ''
ENROLL_EXPORT_SELECTOR = ''

# Date ranges
ENROLLMENT_HISTORY_START = date(2018, 1, 1)
DISPLAY_COMPLETED_YEARS = 2
DETAIL_MONTHS = 4  # current month + 3 completed

# Confirmed during selector discovery — update these after reviewing the actual CSVs
LEADS_COL_NAME = 'Student Name'
LEADS_COL_DATE = 'Lead Date'
LEADS_COL_SOURCE = 'Lead Source'
LEADS_COL_CENTER = 'Center'
LEADS_COL_ASSESSED = 'Assessed'
LEADS_COL_ENROLLED = 'Enrolled'

ENROLL_COL_NAME = 'Student Name'
ENROLL_COL_DATE = 'Enrollment Date'
ENROLL_COL_CENTER = 'Center'
ENROLL_COL_TYPE = 'Enrollment Type'
ENROLL_COL_MRR = 'Monthly Amount'

# Lead filtering
EVENT_LEAD_SOURCES = ['Event']  # confirmed during selector discovery

# Enrollment type filtering
PRIVATE_TUTORING_TYPE = 'Private Tutoring'  # confirmed during selector discovery

# Google Sheets
SPREADSHEET_ID = os.environ['SPREADSHEET_ID']
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ['GOOGLE_SERVICE_ACCOUNT_JSON']

TAB_NAMES = [
    'Funnel - Combined',
    'Funnel - Teaneck',
    'Funnel - Englewood',
    'Enrollments - Combined',
    'Enrollments - Teaneck',
    'Enrollments - Englewood',
]
```

- [ ] **Step 4: Create empty test file and tests package**

```bash
mkdir tests && touch tests/__init__.py tests/test_transform.py
```

- [ ] **Step 5: Install dependencies**

```bash
pip install -r requirements.txt
playwright install chromium
```

- [ ] **Step 6: Commit**

```bash
git add .
git commit -m "feat: initial project scaffolding"
```

---

## Task 2: discover_selectors.py

One-time diagnostic script. Run it against live Radius to find report URLs, filter selectors, and CSV column names. Update config.py after running.

**Files:**
- Create: `discover_selectors.py`

- [ ] **Step 1: Create discover_selectors.py**

```python
"""
Run once to discover Radius report URLs and selectors.
Usage: python discover_selectors.py
Opens a headed browser so you can observe what's happening.
"""
import asyncio
from playwright.async_api import async_playwright
import config

REPORTS_TO_CHECK = [
    '/LeadTracking',
    '/Leads',
    '/LeadReport',
    '/Enrollment',
    '/EnrollmentReport',
    '/Enrollments',
    '/Student/Enrollment',
]

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=500)
        page = await browser.new_page()

        print('Logging in...')
        await page.goto(f'{config.RADIUS_URL}/Login')
        await page.fill('input[name="UserName"], #UserName', config.RADIUS_USERNAME)
        await page.fill('input[name="Password"], #Password', config.RADIUS_PASSWORD)
        await page.click('input[type="submit"], button[type="submit"]')
        await page.wait_for_load_state('networkidle')
        print(f'Logged in. Current URL: {page.url}')

        print('\n--- Checking report URLs ---')
        for path in REPORTS_TO_CHECK:
            url = f'{config.RADIUS_URL}{path}'
            await page.goto(url)
            await page.wait_for_load_state('networkidle')
            title = await page.title()
            status = '404' if '404' in title or 'not found' in title.lower() else 'OK'
            print(f'{status}  {url}  [{title}]')

        print('\n--- Now navigate manually to the Leads Tracking report ---')
        print('When you find it, type the URL here and press Enter.')
        leads_url = input('Leads Tracking URL: ').strip()

        await page.goto(leads_url)
        await page.wait_for_load_state('networkidle')

        print('\nInteractive inputs on this page:')
        inputs = await page.query_selector_all('input, select, [data-role]')
        for el in inputs:
            id_ = await el.get_attribute('id') or ''
            name = await el.get_attribute('name') or ''
            role = await el.get_attribute('data-role') or ''
            tag = await el.evaluate('el => el.tagName')
            print(f'  <{tag.lower()}> id="{id_}" name="{name}" data-role="{role}"')

        print('\nButtons on this page:')
        buttons = await page.query_selector_all('button, input[type="button"], input[type="submit"], a[id*="btn"], a[id*="export"]')
        for btn in buttons:
            id_ = await btn.get_attribute('id') or ''
            text = (await btn.inner_text()).strip()[:60]
            print(f'  id="{id_}" text="{text}"')

        print('\nExport the CSV from the Leads Tracking report manually.')
        print('Save it to: /tmp/leads_sample.csv')
        input('Press Enter when done...')

        import pandas as pd
        try:
            df = pd.read_csv('/tmp/leads_sample.csv')
            print('\nLeads CSV columns:', list(df.columns))
            print(df.head(3).to_string())
        except Exception as e:
            print(f'Could not read CSV: {e}')

        print('\n--- Now navigate to the Enrollment report ---')
        enroll_url = input('Enrollment Report URL: ').strip()
        await page.goto(enroll_url)
        await page.wait_for_load_state('networkidle')

        print('\nInteractive inputs on this page:')
        inputs = await page.query_selector_all('input, select, [data-role]')
        for el in inputs:
            id_ = await el.get_attribute('id') or ''
            name = await el.get_attribute('name') or ''
            role = await el.get_attribute('data-role') or ''
            tag = await el.evaluate('el => el.tagName')
            print(f'  <{tag.lower()}> id="{id_}" name="{name}" data-role="{role}"')

        print('\nExport the CSV from the Enrollment report manually.')
        print('Save it to: /tmp/enrollments_sample.csv')
        input('Press Enter when done...')

        try:
            df = pd.read_csv('/tmp/enrollments_sample.csv')
            print('\nEnrollments CSV columns:', list(df.columns))
            print(df.head(3).to_string())
            print('\nUnique enrollment types:', df.iloc[:, -1].unique() if len(df.columns) > 0 else 'unknown')
        except Exception as e:
            print(f'Could not read CSV: {e}')

        print('\nDone. Update config.py with the URLs, selectors, and column names above.')
        await browser.close()

asyncio.run(main())
```

- [ ] **Step 2: Run the script and record findings**

```bash
python discover_selectors.py
```

Follow the prompts. Note:
- Both report URLs
- Selector IDs for center filter, date pickers, search button, export button
- Exact column names from both CSVs
- All values in the enrollment type column (to confirm `PRIVATE_TUTORING_TYPE`)
- All values in the lead source column (to confirm `EVENT_LEAD_SOURCES`)

- [ ] **Step 3: Update config.py with discovered values**

Fill in all the blank selector and column name constants based on the output above.

- [ ] **Step 4: Commit**

```bash
git add discover_selectors.py config.py
git commit -m "feat: add discover_selectors.py; update config with live selectors"
```

---

## Task 3: radius_export.py

**Files:**
- Create: `radius_export.py`

- [ ] **Step 1: Create radius_export.py**

```python
"""
Exports Leads Tracking and Enrollment CSVs from Radius per center.
Returns DataFrames; does not write to disk.
"""
import asyncio
import io
from datetime import date
from playwright.async_api import async_playwright, Page
import pandas as pd
import config


async def _login(page: Page) -> None:
    await page.goto(f'{config.RADIUS_URL}/Login')
    await page.fill('input[name="UserName"], #UserName', config.RADIUS_USERNAME)
    await page.fill('input[name="Password"], #Password', config.RADIUS_PASSWORD)
    await page.click('input[type="submit"], button[type="submit"]')
    await page.wait_for_load_state('networkidle')


async def _set_date_range(page: Page, start_sel: str, end_sel: str,
                          start: date, end: date) -> None:
    await page.fill(start_sel, start.strftime('%m/%d/%Y'))
    await page.fill(end_sel, end.strftime('%m/%d/%Y'))


async def _select_center(page: Page, center_sel: str, center: str) -> None:
    """Select a single center in a Kendo MultiSelect or standard select."""
    el = await page.query_selector(center_sel)
    tag = await el.evaluate('el => el.tagName') if el else ''
    if tag == 'SELECT':
        await page.select_option(center_sel, label=center)
    else:
        # Kendo MultiSelect: clear, type, click matching item
        await page.click(center_sel)
        await page.keyboard.type(center)
        await page.click(f'li.k-item:has-text("{center}")')


async def _export_csv(page: Page, search_sel: str, export_sel: str) -> pd.DataFrame:
    async with page.expect_download() as dl:
        await page.click(search_sel)
        await page.wait_for_load_state('networkidle')
        await page.click(export_sel)
    download = await dl.value
    content = await download.read()
    return pd.read_csv(io.BytesIO(content))


async def export_leads(center: str, start: date, end: date) -> pd.DataFrame:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await _login(page)
        await page.goto(config.LEADS_REPORT_URL)
        await page.wait_for_load_state('networkidle')
        await _select_center(page, config.LEADS_CENTER_SELECTOR, center)
        await _set_date_range(page, config.LEADS_DATE_START_SELECTOR,
                              config.LEADS_DATE_END_SELECTOR, start, end)
        df = await _export_csv(page, config.LEADS_SEARCH_SELECTOR,
                               config.LEADS_EXPORT_SELECTOR)
        await browser.close()
    df['_center'] = center
    return df


async def export_enrollments(center: str, start: date, end: date) -> pd.DataFrame:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await _login(page)
        await page.goto(config.ENROLLMENT_REPORT_URL)
        await page.wait_for_load_state('networkidle')
        await _select_center(page, config.ENROLL_CENTER_SELECTOR, center)
        await _set_date_range(page, config.ENROLL_DATE_START_SELECTOR,
                              config.ENROLL_DATE_END_SELECTOR, start, end)
        df = await _export_csv(page, config.ENROLL_SEARCH_SELECTOR,
                               config.ENROLL_EXPORT_SELECTOR)
        await browser.close()
    df['_center'] = center
    return df


def export_all(leads_start: date, leads_end: date,
               enroll_start: date, enroll_end: date) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Export both reports for all centers. Returns (leads_df, enrollments_df)."""
    leads_frames = []
    enroll_frames = []
    for center in config.CENTERS:
        leads_frames.append(asyncio.run(export_leads(center, leads_start, leads_end)))
        enroll_frames.append(asyncio.run(export_enrollments(center, enroll_start, enroll_end)))
    return pd.concat(leads_frames, ignore_index=True), pd.concat(enroll_frames, ignore_index=True)
```

- [ ] **Step 2: Smoke test against live Radius (manual, not automated)**

```bash
python - <<'EOF'
import asyncio
from datetime import date
from radius_export import export_leads, export_enrollments
import config

df = asyncio.run(export_leads('Teaneck', date(2026, 1, 1), date(2026, 6, 30)))
print('Leads columns:', list(df.columns))
print('Rows:', len(df))
print(df.head(3))
EOF
```

Expected: DataFrame with columns matching `config.LEADS_COL_*` values, at least some rows.

- [ ] **Step 3: Commit**

```bash
git add radius_export.py
git commit -m "feat: add radius_export.py for leads and enrollment CSVs"
```

---

## Task 4: transform.py — lead filtering and assessed backfill helpers

**Files:**
- Create: `transform.py`
- Modify: `tests/test_transform.py`

- [ ] **Step 1: Write failing tests for lead filtering and assessed backfill**

```python
# tests/test_transform.py
import pandas as pd
import pytest
from transform import filter_event_leads, apply_assessed_backfill, get_enrolled_keys


LEADS = pd.DataFrame({
    'Student Name': ['Alice', 'Bob', 'Carol', 'Dave', 'Eve'],
    'Lead Date':    ['2026-01-15', '2026-01-20', '2026-02-01', '2026-02-10', '2026-03-05'],
    'Lead Source':  ['Walk-in', 'Event', 'Referral', 'Event', 'Web'],
    'Center':       ['Teaneck', 'Teaneck', 'Englewood', 'Englewood', 'Teaneck'],
    'Assessed':     [True, False, True, False, False],
    'Enrolled':     [True, False, True, False, False],
})

# Enrollments: Alice enrolled standard, Carol enrolled standard, Dave enrolled PT (excluded)
ENROLLMENTS = pd.DataFrame({
    'Student Name':    ['Alice', 'Carol', 'Dave'],
    'Enrollment Date': ['2026-01-20', '2026-02-15', '2026-02-20'],
    'Center':          ['Teaneck', 'Englewood', 'Englewood'],
    'Enrollment Type': ['Standard 2x', 'Standard 3x', 'Private Tutoring'],
    'Monthly Amount':  [400, 550, 200],
})


def test_get_enrolled_keys_excludes_private_tutoring():
    keys = get_enrolled_keys(ENROLLMENTS)
    assert ('Alice', 'Teaneck') in keys
    assert ('Carol', 'Englewood') in keys
    assert ('Dave', 'Englewood') not in keys  # PT excluded


def test_filter_event_leads_removes_non_converting_events():
    keys = get_enrolled_keys(ENROLLMENTS)
    result = filter_event_leads(LEADS, keys)
    assert 'Bob' not in result['Student Name'].values   # event lead, didn't enroll
    assert 'Dave' not in result['Student Name'].values  # event lead, only PT enroll


def test_filter_event_leads_keeps_event_leads_that_enrolled():
    # Add an event lead who enrolled non-PT
    extra_lead = pd.DataFrame({
        'Student Name': ['Frank'],
        'Lead Date': ['2026-01-10'],
        'Lead Source': ['Event'],
        'Center': ['Teaneck'],
        'Assessed': [False],
        'Enrolled': [False],
    })
    extra_enroll = pd.DataFrame({
        'Student Name': ['Frank'],
        'Enrollment Date': ['2026-01-25'],
        'Center': ['Teaneck'],
        'Enrollment Type': ['Standard 2x'],
        'Monthly Amount': [400],
    })
    leads = pd.concat([LEADS, extra_lead], ignore_index=True)
    enrollments = pd.concat([ENROLLMENTS, extra_enroll], ignore_index=True)
    keys = get_enrolled_keys(enrollments)
    result = filter_event_leads(leads, keys)
    assert 'Frank' in result['Student Name'].values


def test_filter_event_leads_keeps_non_event_leads_always():
    keys = get_enrolled_keys(ENROLLMENTS)
    result = filter_event_leads(LEADS, keys)
    for name in ['Alice', 'Carol', 'Eve']:
        assert name in result['Student Name'].values


def test_apply_assessed_backfill_marks_enrolled_students_as_assessed():
    keys = get_enrolled_keys(ENROLLMENTS)
    leads = LEADS.copy()
    result = apply_assessed_backfill(leads, keys)
    # Alice already assessed; Carol already assessed; Eve not enrolled, stays False
    alice_row = result[result['Student Name'] == 'Alice']
    assert alice_row['Assessed'].values[0] is True
    carol_row = result[result['Student Name'] == 'Carol']
    assert carol_row['Assessed'].values[0] is True
    eve_row = result[result['Student Name'] == 'Eve']
    assert eve_row['Assessed'].values[0] is False


def test_apply_assessed_backfill_does_not_use_private_tutoring():
    # Dave has only a PT enrollment — should NOT be marked assessed
    keys = get_enrolled_keys(ENROLLMENTS)
    leads = LEADS.copy()
    result = apply_assessed_backfill(leads, keys)
    dave_row = result[result['Student Name'] == 'Dave']
    assert dave_row['Assessed'].values[0] is False
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_transform.py -v
```

Expected: `ModuleNotFoundError: No module named 'transform'`

- [ ] **Step 3: Implement the helpers in transform.py**

```python
# transform.py
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta
import config


def get_enrolled_keys(enrollments_df: pd.DataFrame) -> set[tuple[str, str]]:
    """Return (student_name, center) pairs for non-PT enrollments."""
    non_pt = enrollments_df[
        enrollments_df[config.ENROLL_COL_TYPE] != config.PRIVATE_TUTORING_TYPE
    ]
    return set(zip(non_pt[config.ENROLL_COL_NAME], non_pt[config.ENROLL_COL_CENTER]))


def filter_event_leads(leads_df: pd.DataFrame, enrolled_keys: set[tuple[str, str]]) -> pd.DataFrame:
    """Remove event leads unless they have a non-PT enrollment."""
    is_event = leads_df[config.LEADS_COL_SOURCE].isin(config.EVENT_LEAD_SOURCES)
    key_col = list(zip(leads_df[config.LEADS_COL_NAME], leads_df[config.LEADS_COL_CENTER]))
    has_enrollment = pd.Series([k in enrolled_keys for k in key_col], index=leads_df.index)
    keep = ~is_event | has_enrollment
    return leads_df[keep].copy()


def apply_assessed_backfill(leads_df: pd.DataFrame, enrolled_keys: set[tuple[str, str]]) -> pd.DataFrame:
    """Mark students as assessed if they appear in non-PT enrollments."""
    df = leads_df.copy()
    key_col = list(zip(df[config.LEADS_COL_NAME], df[config.LEADS_COL_CENTER]))
    in_enrollments = pd.Series([k in enrolled_keys for k in key_col], index=df.index)
    df[config.LEADS_COL_ASSESSED] = df[config.LEADS_COL_ASSESSED] | in_enrollments
    return df
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pytest tests/test_transform.py -v
```

Expected: all 6 tests pass.

- [ ] **Step 5: Commit**

```bash
git add transform.py tests/test_transform.py
git commit -m "feat: add lead filter and assessed backfill helpers with tests"
```

---

## Task 5: transform.py — monthly funnel aggregation

**Files:**
- Modify: `transform.py`
- Modify: `tests/test_transform.py`

- [ ] **Step 1: Write failing tests for monthly funnel aggregation**

Add to `tests/test_transform.py`:

```python
from transform import build_funnel


def test_build_funnel_counts_by_month():
    """Funnel aggregates leads/assessed/enrolled by month."""
    leads = pd.DataFrame({
        'Student Name': ['Alice', 'Bob', 'Carol'],
        'Lead Date':    ['2026-01-05', '2026-01-20', '2026-02-10'],
        'Lead Source':  ['Walk-in', 'Referral', 'Walk-in'],
        'Center':       ['Teaneck', 'Teaneck', 'Teaneck'],
        'Assessed':     [True, False, True],
        'Enrolled':     [True, False, False],
    })
    result = build_funnel(leads)
    jan = result[result['month_date'] == pd.Timestamp('2026-01-01')]
    assert jan['leads'].values[0] == 2
    assert jan['assessed'].values[0] == 1
    assert jan['enrolled'].values[0] == 1


def test_build_funnel_calculates_rates():
    leads = pd.DataFrame({
        'Student Name': ['Alice', 'Bob', 'Carol', 'Dave'],
        'Lead Date':    ['2026-01-05', '2026-01-10', '2026-01-15', '2026-01-20'],
        'Lead Source':  ['Walk-in'] * 4,
        'Center':       ['Teaneck'] * 4,
        'Assessed':     [True, True, False, True],
        'Enrolled':     [True, False, False, True],
    })
    result = build_funnel(leads)
    jan = result[result['month_date'] == pd.Timestamp('2026-01-01')]
    assert jan['pct_leads_assessed'].values[0] == pytest.approx(75.0)
    assert jan['pct_assessments_converted'].values[0] == pytest.approx(100 * 2 / 3)
    assert jan['pct_leads_converted'].values[0] == pytest.approx(50.0)


def test_build_funnel_filters_by_center():
    leads = pd.DataFrame({
        'Student Name': ['Alice', 'Bob'],
        'Lead Date':    ['2026-01-05', '2026-01-10'],
        'Lead Source':  ['Walk-in', 'Walk-in'],
        'Center':       ['Teaneck', 'Englewood'],
        'Assessed':     [True, True],
        'Enrolled':     [True, False],
    })
    result = build_funnel(leads, center='Teaneck')
    assert len(result) == 1
    assert result['leads'].values[0] == 1


def test_build_funnel_sorted_newest_first():
    leads = pd.DataFrame({
        'Student Name': ['Alice', 'Bob'],
        'Lead Date':    ['2026-01-05', '2026-03-10'],
        'Lead Source':  ['Walk-in', 'Walk-in'],
        'Center':       ['Teaneck', 'Teaneck'],
        'Assessed':     [True, True],
        'Enrolled':     [True, False],
    })
    result = build_funnel(leads)
    assert result['month_date'].iloc[0] > result['month_date'].iloc[1]
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_transform.py -v -k "funnel"
```

Expected: `AttributeError` or `ImportError` on `build_funnel`.

- [ ] **Step 3: Implement build_funnel in transform.py**

Add to `transform.py`:

```python
def build_funnel(leads_df: pd.DataFrame, center: str | None = None) -> pd.DataFrame:
    """Aggregate leads/assessed/enrolled by month. Returns DataFrame sorted newest first."""
    df = leads_df.copy()
    if center:
        df = df[df[config.LEADS_COL_CENTER] == center]

    df[config.LEADS_COL_DATE] = pd.to_datetime(df[config.LEADS_COL_DATE])
    df['month_date'] = df[config.LEADS_COL_DATE].dt.to_period('M').dt.to_timestamp()

    # Normalize boolean columns (Radius may export as 'Yes'/'No' strings)
    for col in [config.LEADS_COL_ASSESSED, config.LEADS_COL_ENROLLED]:
        if df[col].dtype == object:
            df[col] = df[col].str.strip().str.lower().isin(['yes', 'true', '1'])
        else:
            df[col] = df[col].astype(bool)

    agg = df.groupby('month_date').agg(
        leads=(config.LEADS_COL_NAME, 'count'),
        assessed=(config.LEADS_COL_ASSESSED, 'sum'),
        enrolled=(config.LEADS_COL_ENROLLED, 'sum'),
    ).reset_index()

    agg['pct_leads_assessed'] = (agg['assessed'] / agg['leads'] * 100).round(1)
    agg['pct_assessments_converted'] = (agg['enrolled'] / agg['assessed'].replace(0, pd.NA) * 100).round(1)
    agg['pct_leads_converted'] = (agg['enrolled'] / agg['leads'] * 100).round(1)

    return agg.sort_values('month_date', ascending=False).reset_index(drop=True)
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pytest tests/test_transform.py -v
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add transform.py tests/test_transform.py
git commit -m "feat: add build_funnel with monthly aggregation and rate calculations"
```

---

## Task 6: transform.py — enrollment summary and detail

**Files:**
- Modify: `transform.py`
- Modify: `tests/test_transform.py`

- [ ] **Step 1: Write failing tests**

Add to `tests/test_transform.py`:

```python
from transform import build_enrollment_summary, build_enrollment_detail, get_display_date_range, get_detail_month_range
from unittest.mock import patch
from datetime import date


ENROLL_DF = pd.DataFrame({
    'Student Name':    ['Alice', 'Bob', 'Carol', 'Dave', 'Eve'],
    'Enrollment Date': ['2026-01-10', '2026-01-20', '2026-02-05', '2026-03-01', '2026-03-15'],
    'Center':          ['Teaneck', 'Teaneck', 'Englewood', 'Teaneck', 'Teaneck'],
    'Enrollment Type': ['Standard 2x', 'Standard 3x', 'Standard 2x', 'Private Tutoring', 'Standard 2x'],
    'Monthly Amount':  [400, 550, 400, 200, 400],
})


def test_build_enrollment_summary_excludes_pt_from_main_counts():
    result = build_enrollment_summary(ENROLL_DF)
    jan = result[result['month_date'] == pd.Timestamp('2026-01-01')]
    assert jan['total_enrollments'].values[0] == 2  # Alice + Bob, not PT
    assert jan['pt_count'].values[0] == 0


def test_build_enrollment_summary_tracks_pt_separately():
    result = build_enrollment_summary(ENROLL_DF)
    mar = result[result['month_date'] == pd.Timestamp('2026-03-01')]
    assert mar['total_enrollments'].values[0] == 1  # Eve only
    assert mar['pt_count'].values[0] == 1           # Dave


def test_build_enrollment_summary_mrr_excludes_pt():
    result = build_enrollment_summary(ENROLL_DF)
    jan = result[result['month_date'] == pd.Timestamp('2026-01-01')]
    assert jan['mrr'].values[0] == 950  # 400 + 550


def test_build_enrollment_summary_has_type_columns():
    result = build_enrollment_summary(ENROLL_DF)
    assert 'Standard 2x' in result.columns
    assert 'Standard 3x' in result.columns
    assert 'Private Tutoring' not in result.columns


def test_build_enrollment_summary_sorted_newest_first():
    result = build_enrollment_summary(ENROLL_DF)
    assert result['month_date'].iloc[0] > result['month_date'].iloc[1]


def test_build_enrollment_summary_filters_by_center():
    result = build_enrollment_summary(ENROLL_DF, center='Teaneck')
    feb = result[result['month_date'] == pd.Timestamp('2026-02-01')]
    assert len(feb) == 0  # Carol is Englewood


def test_build_enrollment_detail_excludes_pt():
    with patch('transform.date') as mock_date:
        mock_date.today.return_value = date(2026, 3, 20)
        mock_date.side_effect = lambda *a, **k: date(*a, **k)
        result = build_enrollment_detail(ENROLL_DF)
    assert 'Dave' not in result['Student Name'].values  # Dave is PT


def test_build_enrollment_detail_covers_4_months():
    with patch('transform.date') as mock_date:
        mock_date.today.return_value = date(2026, 3, 20)
        mock_date.side_effect = lambda *a, **k: date(*a, **k)
        result = build_enrollment_detail(ENROLL_DF)
    months = result['month_date'].dt.to_period('M').unique()
    # Expecting Dec 2025, Jan 2026, Feb 2026, Mar 2026 — but ENROLL_DF only has Jan-Mar
    assert pd.Period('2026-01') in months
    assert pd.Period('2026-02') in months
    assert pd.Period('2026-03') in months
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
pytest tests/test_transform.py -v -k "enrollment"
```

Expected: `ImportError` on new functions.

- [ ] **Step 3: Implement enrollment functions in transform.py**

Add to `transform.py`:

```python
def get_display_date_range() -> tuple[date, date]:
    """Start of (current_year - DISPLAY_COMPLETED_YEARS), through today."""
    today = date.today()
    start = date(today.year - config.DISPLAY_COMPLETED_YEARS, 1, 1)
    return start, today


def get_detail_month_range() -> tuple[date, date]:
    """Start of month (today - 3 months), through today. Covers 4 months."""
    today = date.today()
    start = (today - relativedelta(months=config.DETAIL_MONTHS - 1)).replace(day=1)
    return start, today


def build_enrollment_summary(enrollments_df: pd.DataFrame, center: str | None = None) -> pd.DataFrame:
    """Monthly enrollment summary. PT tracked separately in pt_count / pt_mrr columns."""
    df = enrollments_df.copy()
    if center:
        df = df[df[config.ENROLL_COL_CENTER] == center]

    df[config.ENROLL_COL_DATE] = pd.to_datetime(df[config.ENROLL_COL_DATE])
    df['month_date'] = df[config.ENROLL_COL_DATE].dt.to_period('M').dt.to_timestamp()

    non_pt = df[df[config.ENROLL_COL_TYPE] != config.PRIVATE_TUTORING_TYPE]
    pt = df[df[config.ENROLL_COL_TYPE] == config.PRIVATE_TUTORING_TYPE]

    # Main aggregation (non-PT)
    base = non_pt.groupby('month_date').agg(
        total_enrollments=(config.ENROLL_COL_NAME, 'count'),
        mrr=(config.ENROLL_COL_MRR, 'sum'),
    ).reset_index()

    # Per-type pivot (non-PT)
    type_pivot = non_pt.pivot_table(
        index='month_date',
        columns=config.ENROLL_COL_TYPE,
        values=config.ENROLL_COL_NAME,
        aggfunc='count',
        fill_value=0,
    ).reset_index()
    type_pivot.columns.name = None

    # PT aggregation
    pt_agg = pt.groupby('month_date').agg(
        pt_count=(config.ENROLL_COL_NAME, 'count'),
        pt_mrr=(config.ENROLL_COL_MRR, 'sum'),
    ).reset_index()

    result = base.merge(type_pivot, on='month_date', how='left')
    result = result.merge(pt_agg, on='month_date', how='left')
    result[['pt_count', 'pt_mrr']] = result[['pt_count', 'pt_mrr']].fillna(0).astype(int)

    return result.sort_values('month_date', ascending=False).reset_index(drop=True)


def build_enrollment_detail(enrollments_df: pd.DataFrame, center: str | None = None) -> pd.DataFrame:
    """Individual enrollment rows for the past DETAIL_MONTHS months. Excludes PT."""
    df = enrollments_df.copy()
    if center:
        df = df[df[config.ENROLL_COL_CENTER] == center]

    df[config.ENROLL_COL_DATE] = pd.to_datetime(df[config.ENROLL_COL_DATE])
    df['month_date'] = df[config.ENROLL_COL_DATE].dt.to_period('M').dt.to_timestamp()

    start, _ = get_detail_month_range()
    cutoff = pd.Timestamp(start)

    df = df[
        (df['month_date'] >= cutoff) &
        (df[config.ENROLL_COL_TYPE] != config.PRIVATE_TUTORING_TYPE)
    ]

    return df.sort_values([config.ENROLL_COL_DATE], ascending=False).reset_index(drop=True)
```

- [ ] **Step 4: Run all tests**

```bash
pytest tests/test_transform.py -v
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add transform.py tests/test_transform.py
git commit -m "feat: add enrollment summary and detail transforms with tests"
```

---

## Task 7: sheets.py

**Files:**
- Create: `sheets.py`

No automated tests for this task — the Google Sheets API is an external integration. Manual smoke test only.

- [ ] **Step 1: Create a Google Cloud service account and share the sheet**

1. Go to Google Cloud Console → create a project → enable Google Sheets API
2. Create a service account → create JSON key → download it
3. Create a new Google Sheet → share it with the service account email (Editor)
4. Copy the spreadsheet ID from the sheet URL (the long string between `/d/` and `/edit`)
5. Create all 6 tabs with the exact names from `config.TAB_NAMES`
6. Set env vars locally:
   ```bash
   export SPREADSHEET_ID='your-sheet-id'
   export GOOGLE_SERVICE_ACCOUNT_JSON='{"type":"service_account",...}'  # full JSON as string
   ```

- [ ] **Step 2: Create sheets.py**

```python
"""
Writes all 6 tabs to the Google Sheet. Fully rewrites each tab on every run.
"""
import json
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import config

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

FUNNEL_HEADERS = [
    'Month', 'Leads', 'Assessed', 'Enrolled',
    '% leads assessed', '% assessments converted', '% leads converted',
]

ENROLL_SUMMARY_STATIC_HEADERS = ['Month', 'Total enrollments', 'MRR']
ENROLL_DETAIL_HEADERS = ['Month', 'Student name', 'Center', 'Enrollment type', 'MRR', 'Enrollment date']


def _get_client() -> gspread.Client:
    creds = Credentials.from_service_account_info(
        json.loads(config.GOOGLE_SERVICE_ACCOUNT_JSON),
        scopes=SCOPES,
    )
    return gspread.authorize(creds)


def _write_tab(worksheet: gspread.Worksheet, rows: list[list]) -> None:
    worksheet.clear()
    if rows:
        worksheet.update(rows, value_input_option='USER_ENTERED')


def _fmt_month(ts: pd.Timestamp) -> str:
    return ts.strftime('%b %Y')


def _fmt_pct(val) -> str:
    if pd.isna(val):
        return 'N/A'
    return f'{val:.1f}%'


def _funnel_rows(funnel_df: pd.DataFrame) -> list[list]:
    """Convert funnel DataFrame to sheet rows with year group headers."""
    rows = [FUNNEL_HEADERS]
    current_year = None
    for _, row in funnel_df.iterrows():
        year = row['month_date'].year
        if year != current_year:
            label = str(year) if year < pd.Timestamp.now().year else f'{year} (YTD)'
            rows.append([label] + [''] * (len(FUNNEL_HEADERS) - 1))
            current_year = year
        rows.append([
            _fmt_month(row['month_date']),
            int(row['leads']),
            int(row['assessed']),
            int(row['enrolled']),
            _fmt_pct(row['pct_leads_assessed']),
            _fmt_pct(row['pct_assessments_converted']),
            _fmt_pct(row['pct_leads_converted']),
        ])
    return rows


def _enrollment_rows(summary_df: pd.DataFrame, detail_df: pd.DataFrame,
                     type_cols: list[str]) -> list[list]:
    """Build all rows for an enrollment tab: summary section then detail section."""
    summary_headers = ENROLL_SUMMARY_STATIC_HEADERS + type_cols + ['Private Tutoring', 'PT MRR']
    rows = [summary_headers]

    current_year = None
    for _, row in summary_df.iterrows():
        year = row['month_date'].year
        if year != current_year:
            label = str(year) if year < pd.Timestamp.now().year else f'{year} (YTD)'
            rows.append([label] + [''] * (len(summary_headers) - 1))
            current_year = year
        type_counts = [int(row.get(t, 0)) for t in type_cols]
        rows.append([
            _fmt_month(row['month_date']),
            int(row['total_enrollments']),
            f"${int(row['mrr']):,}",
            *type_counts,
            int(row['pt_count']),
            f"${int(row['pt_mrr']):,}",
        ])

    # Blank separator + detail header
    rows.append([])
    rows.append(['--- Individual enrollment detail (recent 4 months) ---'] + [''] * (len(summary_headers) - 1))
    rows.append(ENROLL_DETAIL_HEADERS + [''] * (len(summary_headers) - len(ENROLL_DETAIL_HEADERS)))

    current_month = None
    for _, row in detail_df.iterrows():
        month_label = _fmt_month(row['month_date'])
        if month_label != current_month:
            rows.append([month_label] + [''] * (len(summary_headers) - 1))
            current_month = month_label
        rows.append([
            '',
            row[config.ENROLL_COL_NAME],
            row[config.ENROLL_COL_CENTER],
            row[config.ENROLL_COL_TYPE],
            f"${int(row[config.ENROLL_COL_MRR]):,}",
            pd.to_datetime(row[config.ENROLL_COL_DATE]).strftime('%-m/%-d/%Y'),
            *([''] * (len(summary_headers) - 6)),
        ])
    return rows


def write_all_tabs(
    funnel_combined: pd.DataFrame,
    funnel_teaneck: pd.DataFrame,
    funnel_englewood: pd.DataFrame,
    enroll_summary_combined: pd.DataFrame,
    enroll_summary_teaneck: pd.DataFrame,
    enroll_summary_englewood: pd.DataFrame,
    enroll_detail_combined: pd.DataFrame,
    enroll_detail_teaneck: pd.DataFrame,
    enroll_detail_englewood: pd.DataFrame,
    type_cols: list[str],
) -> None:
    client = _get_client()
    sheet = client.open_by_key(config.SPREADSHEET_ID)

    tab_data = {
        'Funnel - Combined':      _funnel_rows(funnel_combined),
        'Funnel - Teaneck':       _funnel_rows(funnel_teaneck),
        'Funnel - Englewood':     _funnel_rows(funnel_englewood),
        'Enrollments - Combined': _enrollment_rows(enroll_summary_combined, enroll_detail_combined, type_cols),
        'Enrollments - Teaneck':  _enrollment_rows(enroll_summary_teaneck, enroll_detail_teaneck, type_cols),
        'Enrollments - Englewood':_enrollment_rows(enroll_summary_englewood, enroll_detail_englewood, type_cols),
    }

    for tab_name, rows in tab_data.items():
        ws = sheet.worksheet(tab_name)
        _write_tab(ws, rows)
        print(f'  wrote {len(rows)} rows to "{tab_name}"')
```

- [ ] **Step 3: Smoke test against the real sheet**

```bash
python - <<'EOF'
import pandas as pd
from sheets import write_all_tabs

# Minimal fake data
funnel = pd.DataFrame({
    'month_date': [pd.Timestamp('2026-01-01')],
    'leads': [10], 'assessed': [8], 'enrolled': [5],
    'pct_leads_assessed': [80.0], 'pct_assessments_converted': [62.5], 'pct_leads_converted': [50.0],
})
enroll_summary = pd.DataFrame({
    'month_date': [pd.Timestamp('2026-01-01')],
    'total_enrollments': [5], 'mrr': [2000],
    'Standard 2x': [3], 'Standard 3x': [2],
    'pt_count': [1], 'pt_mrr': [200],
})
enroll_detail = pd.DataFrame({
    'month_date': [pd.Timestamp('2026-01-10')],
    'Student Name': ['Test Student'], 'Center': ['Teaneck'],
    'Enrollment Type': ['Standard 2x'], 'Monthly Amount': [400],
    'Enrollment Date': ['2026-01-10'],
})
type_cols = ['Standard 2x', 'Standard 3x']

write_all_tabs(funnel, funnel, funnel, enroll_summary, enroll_summary, enroll_summary,
               enroll_detail, enroll_detail, enroll_detail, type_cols)
print('Done — check the Google Sheet.')
EOF
```

Expected: all 6 tabs in the sheet are updated with the fake data.

- [ ] **Step 4: Commit**

```bash
git add sheets.py
git commit -m "feat: add sheets.py to write all 6 Google Sheet tabs"
```

---

## Task 8: main.py — orchestrator

**Files:**
- Create: `main.py`

- [ ] **Step 1: Create main.py**

```python
"""
Entry point. Exports from Radius, transforms, writes to Google Sheet.
Usage:
  python main.py
  python main.py --date 2026-06-15   # override today's date for testing
"""
import argparse
from datetime import date
import pandas as pd
from dateutil.relativedelta import relativedelta

import config
from radius_export import export_all
from transform import (
    get_enrolled_keys,
    filter_event_leads,
    apply_assessed_backfill,
    build_funnel,
    build_enrollment_summary,
    build_enrollment_detail,
    get_display_date_range,
)
from sheets import write_all_tabs


def run(today: date) -> None:
    display_start, display_end = get_display_date_range()
    enroll_start = config.ENROLLMENT_HISTORY_START

    print('Exporting from Radius...')
    leads_raw, enrollments_raw = export_all(
        leads_start=display_start,
        leads_end=today,
        enroll_start=enroll_start,
        enroll_end=today,
    )

    print('Transforming data...')
    enrolled_keys = get_enrolled_keys(enrollments_raw)
    leads_filtered = filter_event_leads(leads_raw, enrolled_keys)
    leads_final = apply_assessed_backfill(leads_filtered, enrolled_keys)

    # Filter leads to display range (enrollments already cover full history)
    leads_display = leads_final[
        pd.to_datetime(leads_final[config.LEADS_COL_DATE]) >= pd.Timestamp(display_start)
    ]

    # Filter enrollments to display range for summary
    enrollments_display = enrollments_raw[
        pd.to_datetime(enrollments_raw[config.ENROLL_COL_DATE]) >= pd.Timestamp(display_start)
    ]

    # Discover enrollment type columns (non-PT, sorted)
    non_pt = enrollments_display[
        enrollments_display[config.ENROLL_COL_TYPE] != config.PRIVATE_TUTORING_TYPE
    ]
    type_cols = sorted(non_pt[config.ENROLL_COL_TYPE].dropna().unique().tolist())

    print('Building report tables...')
    funnel_combined  = build_funnel(leads_display)
    funnel_teaneck   = build_funnel(leads_display, center='Teaneck')
    funnel_englewood = build_funnel(leads_display, center='Englewood')

    es_combined  = build_enrollment_summary(enrollments_display)
    es_teaneck   = build_enrollment_summary(enrollments_display, center='Teaneck')
    es_englewood = build_enrollment_summary(enrollments_display, center='Englewood')

    ed_combined  = build_enrollment_detail(enrollments_raw)
    ed_teaneck   = build_enrollment_detail(enrollments_raw, center='Teaneck')
    ed_englewood = build_enrollment_detail(enrollments_raw, center='Englewood')

    print('Writing to Google Sheets...')
    write_all_tabs(
        funnel_combined, funnel_teaneck, funnel_englewood,
        es_combined, es_teaneck, es_englewood,
        ed_combined, ed_teaneck, ed_englewood,
        type_cols,
    )
    print('Done.')


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument('--date', help='Override today (YYYY-MM-DD)')
    args = parser.parse_args()
    today = date.fromisoformat(args.date) if args.date else date.today()
    run(today)


if __name__ == '__main__':
    main()
```

- [ ] **Step 2: Run end-to-end against live Radius and the real sheet**

```bash
python main.py
```

Expected:
- Radius exports complete without errors
- Google Sheet updated with real data across all 6 tabs
- No tracebacks

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: add main.py orchestrator"
```

---

## Task 9: GitHub Actions workflow

**Files:**
- Create: `.github/workflows/run.yml`

- [ ] **Step 1: Create the workflow**

```yaml
# .github/workflows/run.yml
name: Radius Dashboard

on:
  schedule:
    - cron: '0 14 1 * *'   # 9am ET on the 1st
    - cron: '0 14 15 * *'  # 9am ET on the 15th
  workflow_dispatch:        # manual trigger

jobs:
  run:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: |
          pip install -r requirements.txt
          playwright install chromium --with-deps

      - name: Run dashboard
        env:
          RADIUS_USERNAME: ${{ secrets.RADIUS_USERNAME }}
          RADIUS_PASSWORD: ${{ secrets.RADIUS_PASSWORD }}
          SPREADSHEET_ID: ${{ secrets.SPREADSHEET_ID }}
          GOOGLE_SERVICE_ACCOUNT_JSON: ${{ secrets.GOOGLE_SERVICE_ACCOUNT_JSON }}
        run: python main.py
```

- [ ] **Step 2: Add GitHub secrets**

In the GitHub repo settings → Secrets and variables → Actions, add:
- `RADIUS_USERNAME`
- `RADIUS_PASSWORD`
- `SPREADSHEET_ID`
- `GOOGLE_SERVICE_ACCOUNT_JSON` (the full JSON key as a string)

These can be added to the existing `mdiamond77` org-level secrets if you want to reuse them across repos.

- [ ] **Step 3: Push and verify the workflow triggers manually**

```bash
git add .github/
git commit -m "feat: add GitHub Actions workflow with schedule and manual trigger"
git remote add origin https://github.com/mdiamond77/radius-dashboard.git
git push -u origin main
```

Then go to the Actions tab in GitHub → select "Radius Dashboard" → Run workflow → Run workflow.

Expected: workflow runs green, Google Sheet updated.

---

## Self-Review

**Spec coverage check:**

| Requirement | Covered by |
|---|---|
| Lead funnel by month (leads/assessed/enrolled + rates) | Tasks 4, 5 |
| Enrollment types + MRR by month | Task 6 |
| Exclude event leads unless enrolled | Task 4 |
| Any enrolled (non-PT) student counts as assessed | Task 4 |
| PT excluded from main counts, shown separately | Task 6 |
| Individual detail for last 3 completed + current month | Task 6 |
| Per-center + combined tabs | Tasks 7, 8 |
| Newest months first with year headers | Tasks 7 |
| 2 completed years + YTD display | Tasks 5, 6 |
| Enrollment report from Jan 1, 2018 | Task 8 |
| Google Sheets delivery | Task 7 |
| GitHub Actions cron (1st and 15th) + manual trigger | Task 9 |
| Selector discovery before implementation | Task 2 |

All requirements covered. No placeholders remain.
