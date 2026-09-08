# Summer Package Status Report Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an on-demand CLI tool that pulls the Radius Enrollment Report for Teaneck and Englewood, finds every "Summer 2026 (Sessions Package)" enrollment, classifies each student by current status and school-year conversion, and renders a single self-contained HTML report.

**Architecture:** One Playwright headless pull of the Enrollment Report per center (4/1/2026 → today, all statuses). A pure-function transform layer (pandas) identifies summer students and buckets each one using only that student's other enrollment rows. A render layer emits one portable HTML file with an overall summary plus a per-center section. Run by hand; no GitHub Actions.

**Tech Stack:** Python 3.11+, Playwright (headless Chromium), pandas, openpyxl (read the export if it is `.xlsx`), pytest. Radius credentials from `~/.mathnasium_env`.

---

## Reference: spec

Design spec: `~/Mathnasium_automation/docs/superpowers/specs/2026-09-07-summer-package-status-report-design.md`

## Reference: existing patterns to copy

These live under `~/` as sibling repos. Read before starting — they show the exact Radius/Playwright idioms this project follows:

- `~/mathnasium-new-enrollments/` — closest cousin. `download.py` (login + Enrollment Report), `config.py` (center IDs, `TBD_` column constants), `discover_selectors.py`.
- `~/mathnasium-binder-audit/discover_selectors.py` — reusable selector-dump template.

**Radius selector constants (known, stable across reports):**

```python
LOGIN_URL = "https://radius.mathnasium.com"
# login form
SEL_USERNAME = "#UserName"
SEL_PASSWORD = "#Password"
SEL_LOGIN_BTN = "#login"
# center multiselect (Kendo)
SEL_CENTER_MULTISELECT = "#AllCenterListMultiSelect"
CENTER_ID_TEANECK = "2871"
CENTER_ID_ENGLEWOOD = "2428"
# search / export
SEL_SEARCH_BTN = "#btnsearch"
SEL_EXPORT_BTN = "#btnExport"
# enrollment report
ENROLLMENT_REPORT_URL = "https://radius.mathnasium.com/Enrollment/EnrollmentReport"
SEL_ENROLLMENT_STATUS_DD = "#EnrollmentStatusDropDown"  # value "3" = Enrolled; we want ALL, so leave default/unset
```

Build JS arrays passed to `page.evaluate()` with `json.dumps(...)`, never `str(list)`.

---

## File Structure

| File | Responsibility |
|---|---|
| `~/mathnasium-summer-status/config.py` | Constants: URLs, selectors, center config, the summer membership string, pull window, `TBD_` column-name constants filled after discovery |
| `~/mathnasium-summer-status/discover_selectors.py` | One-time headed script: dump Enrollment Report grid columns + confirm row granularity and membership-type filter behavior |
| `~/mathnasium-summer-status/download.py` | Playwright headless: login, pull Enrollment Report per center, save raw export to `input/EnrollmentReport_<center>.xlsx` |
| `~/mathnasium-summer-status/transform.py` | Pure functions over DataFrames: normalize, identify summer students, classify each, build summaries |
| `~/mathnasium-summer-status/report.py` | Render `output/summer_2026_status.html` from transformed data; open it in the browser |
| `~/mathnasium-summer-status/main.py` | CLI orchestrator: `--no-download` flag, download → transform → report |
| `~/mathnasium-summer-status/tests/test_transform.py` | Unit tests for every classification bucket + edge cases |
| `~/mathnasium-summer-status/requirements.txt` | Dependencies |
| `~/mathnasium-summer-status/.gitignore` | `input/`, `output/`, `__pycache__/`, `*.pyc`, `.env` |

---

## Task 1: Project scaffold

**Files:**
- Create: `~/mathnasium-summer-status/requirements.txt`
- Create: `~/mathnasium-summer-status/.gitignore`
- Create: `~/mathnasium-summer-status/config.py`
- Create: `~/mathnasium-summer-status/tests/__init__.py`

- [ ] **Step 1: Create the project directory and git repo**

```bash
mkdir -p ~/mathnasium-summer-status/tests ~/mathnasium-summer-status/input ~/mathnasium-summer-status/output
cd ~/mathnasium-summer-status
git init
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
output/
__pycache__/
*.pyc
.env
```

- [ ] **Step 4: Create `config.py`**

```python
from pathlib import Path
from datetime import date

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

# ── Radius ────────────────────────────────────────────────────────────────
LOGIN_URL = "https://radius.mathnasium.com"
ENROLLMENT_REPORT_URL = "https://radius.mathnasium.com/Enrollment/EnrollmentReport"

SEL_USERNAME = "#UserName"
SEL_PASSWORD = "#Password"
SEL_LOGIN_BTN = "#login"
SEL_CENTER_MULTISELECT = "#AllCenterListMultiSelect"
SEL_SEARCH_BTN = "#btnsearch"
SEL_EXPORT_BTN = "#btnExport"

CENTERS = {
    "Teaneck":   {"radius_id": "2871"},
    "Englewood": {"radius_id": "2428"},
}

# ── Report parameters ─────────────────────────────────────────────────────
SUMMER_MEMBERSHIP = "Summer 2026 (Sessions Package)"
PULL_START = "4/1/2026"   # passed to the report's From date field
# PULL_END is "today" — computed at runtime

# ── Enrollment Report column names — FILL IN AFTER discover_selectors.py ───
# Run discover_selectors.py, open input/EnrollmentReport_Teaneck.xlsx, and
# replace each string below with the exact column header from the export.
COL_STUDENT      = "TBD_StudentName"     # full student name; if split, see COL_FIRST/COL_LAST
COL_FIRST        = None                  # set if name is split across two columns, else None
COL_LAST         = None
COL_CENTER       = "TBD_Center"
COL_MEMBERSHIP   = "TBD_MembershipType"  # must contain the SUMMER_MEMBERSHIP string
COL_START_DATE   = "TBD_EnrollmentStart"
COL_STATUS       = "TBD_EnrollmentStatus"
COL_SESSIONS_WK  = "TBD_SessionsPerWeek" # set to None if the export has no such field

# ── Status string mapping — FILL IN / CONFIRM AFTER inspecting real export ─
# Map the raw Radius status strings (lowercased) to canonical tokens.
STATUS_ENROLLED  = {"enrolled", "active"}
STATUS_HOLD      = {"hold", "on hold"}
STATUS_CANCELLED = {"cancelled", "canceled", "dropped"}
STATUS_DONE      = {"completed", "expired", "finished", "ended"}
```

- [ ] **Step 5: Create `tests/__init__.py`** (empty file)

```python
```

- [ ] **Step 6: Commit**

```bash
cd ~/mathnasium-summer-status
git add -A
git commit -m "chore: project scaffold for summer package status report"
```

---

## Task 2: Selector / column discovery (GATE)

**Files:**
- Create: `~/mathnasium-summer-status/discover_selectors.py`

This task produces information, not shippable code. Its output decides whether Approach 1 is viable.

- [ ] **Step 1: Write `discover_selectors.py`**

```python
"""One-time headed discovery for the Enrollment Report.

Confirms: (a) the grid/export exposes membership type, enrollment start date,
enrollment status, and a sessions/week field; (b) the export has one row per
enrollment, not one per student; (c) whether membership type can be filtered
server-side or must be filtered after export.

Run:  export $(cat ~/.mathnasium_env | xargs) && python discover_selectors.py
"""
import os
from playwright.sync_api import sync_playwright
import config as c

def main():
    user = os.environ["RADIUS_USERNAME"]
    pw = os.environ["RADIUS_PASSWORD"]
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=300)
        with browser:
            context = browser.new_context(accept_downloads=True)
            with context:
                page = context.new_page()
                page.goto(c.LOGIN_URL)
                page.fill(c.SEL_USERNAME, user)
                page.fill(c.SEL_PASSWORD, pw)
                page.click(c.SEL_LOGIN_BTN)
                page.wait_for_load_state("networkidle")

                page.goto(c.ENROLLMENT_REPORT_URL)
                page.wait_for_load_state("networkidle")

                # Dump every element that has an id — find date fields,
                # membership-type filter, status dropdown.
                ids = page.evaluate(
                    "Array.from(document.querySelectorAll('[id]')).map(e => "
                    "({id: e.id, tag: e.tagName, type: e.type, name: e.name}))"
                )
                print("=== ELEMENTS WITH IDS ===")
                for el in ids:
                    print(el)

                input("Manually run the report for Teaneck, all statuses, "
                      "4/1/2026 -> today. Then press Enter to dump grid headers...")

                headers = page.evaluate(
                    "Array.from(document.querySelectorAll('.k-grid-header th')).map(e => e.innerText.trim())"
                )
                print("=== GRID COLUMN HEADERS ===")
                for h in headers:
                    print(repr(h))

                input("Now export the grid to Excel manually and save it to "
                      "input/ as EnrollmentReport_Teaneck.xlsx. Press Enter when done.")

if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Run it**

```bash
cd ~/mathnasium-summer-status
export $(cat ~/.mathnasium_env | xargs)
python discover_selectors.py
```

- [ ] **Step 3: Record findings in a scratch note**

Create `~/mathnasium-summer-status/DISCOVERY_NOTES.md` with:
- The exact From/To date field selectors (ids)
- Whether a membership-type filter control exists on the report (id, or "none — filter after export")
- The exact grid column headers, verbatim
- Confirmation: does the export have multiple rows for a student with multiple enrollments? (open the xlsx, look for a known multi-enrollment student, or sort by name)
- The distinct values seen in the status column

```bash
git add DISCOVERY_NOTES.md
git commit -m "docs: record Enrollment Report discovery findings"
```

- [ ] **Step 4: GATE CHECK**

If the export is **one row per student** (not per enrollment) OR has **no membership-type column**, STOP. Report to the user: "Approach 1 isn't viable — the Enrollment Report is [one-row-per-student / missing membership type]. The spec's fallback is Approach 3 (add the Student Report). Want me to revise the spec?"

Otherwise continue.

- [ ] **Step 5: Fill in `config.py`**

Replace every `TBD_` constant with the verbatim column header from the export. Set `COL_SESSIONS_WK = None` if there is no such column. Set `COL_FIRST`/`COL_LAST` if the name is split. Adjust the `STATUS_*` sets to include every distinct raw status value seen (lowercased). Commit:

```bash
git add config.py
git commit -m "chore: fill in Enrollment Report column names from discovery"
```

---

## Task 3: `download.py` — pull the Enrollment Report

**Files:**
- Create: `~/mathnasium-summer-status/download.py`

Selectors for the date fields and export flow come from Task 2. The code below uses
placeholder names `SEL_FROM_DATE` / `SEL_TO_DATE` — **replace with the real ids
recorded in DISCOVERY_NOTES.md** and add them to `config.py` first.

- [ ] **Step 1: Add the date-field selectors to `config.py`**

```python
# Enrollment Report date fields (from discovery — replace with real ids)
SEL_FROM_DATE = "#ReportStart"   # REPLACE with actual id from DISCOVERY_NOTES.md
SEL_TO_DATE = "#ReportEnd"       # REPLACE with actual id from DISCOVERY_NOTES.md
```

```bash
git add config.py && git commit -m "chore: add enrollment report date-field selectors"
```

- [ ] **Step 2: Write `download.py`**

```python
"""Pull the Radius Enrollment Report for each center and cache the raw export."""
import os
from datetime import date
from playwright.sync_api import sync_playwright
import config as c


def _login(page):
    page.goto(c.LOGIN_URL)
    page.fill(c.SEL_USERNAME, os.environ["RADIUS_USERNAME"])
    page.fill(c.SEL_PASSWORD, os.environ["RADIUS_PASSWORD"])
    page.click(c.SEL_LOGIN_BTN)
    page.wait_for_load_state("networkidle")


def _set_center(page, radius_id: str):
    # Kendo MultiSelect: clear any existing tags, then set the single value via JS.
    page.evaluate(
        """(id) => {
            const ms = $('#AllCenterListMultiSelect').data('kendoMultiSelect');
            ms.value([id]);
            ms.trigger('change');
        }""",
        radius_id,
    )


def _pull_center(page, center_name: str, radius_id: str) -> str:
    page.goto(c.ENROLLMENT_REPORT_URL)
    page.wait_for_load_state("networkidle")
    _set_center(page, radius_id)
    page.fill(c.SEL_FROM_DATE, c.PULL_START)
    page.fill(c.SEL_TO_DATE, date.today().strftime("%-m/%-d/%Y"))
    # Leave the status dropdown at its default so ALL statuses are returned.
    page.click(c.SEL_SEARCH_BTN)
    page.wait_for_load_state("networkidle")

    out_path = c.INPUT_DIR / f"EnrollmentReport_{center_name}.xlsx"
    with page.expect_download() as dl:
        page.click(c.SEL_EXPORT_BTN)
    dl.value.save_as(str(out_path))
    return str(out_path)


def download_all() -> dict[str, str]:
    """Returns {center_name: path_to_xlsx}. Raises on login failure."""
    c.INPUT_DIR.mkdir(exist_ok=True)
    paths = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        with browser:
            context = browser.new_context(accept_downloads=True)
            with context:
                page = context.new_page()
                _login(page)
                if page.query_selector(c.SEL_USERNAME):
                    raise RuntimeError("Radius login failed — still on the login page")
                for name, cfg in c.CENTERS.items():
                    paths[name] = _pull_center(page, name, cfg["radius_id"])
    return paths


if __name__ == "__main__":
    for name, path in download_all().items():
        print(f"{name}: {path}")
```

- [ ] **Step 3: Run it against real Radius**

```bash
cd ~/mathnasium-summer-status
export $(cat ~/.mathnasium_env | xargs)
python download.py
```

Expected: `input/EnrollmentReport_Teaneck.xlsx` and `input/EnrollmentReport_Englewood.xlsx` exist and open in Excel with the columns seen during discovery.

- [ ] **Step 4: Fix selectors until it works.** The Kendo `_set_center` JS and the date-field fill are the likely failure points — adjust against the live page. Re-run until both files download with the right center's data (spot-check a few student names).

- [ ] **Step 5: Commit**

```bash
git add download.py config.py
git commit -m "feat: pull Enrollment Report per center via Playwright"
```

---

## Task 4: `transform.py` — load + normalize

**Files:**
- Create: `~/mathnasium-summer-status/transform.py`
- Test: `~/mathnasium-summer-status/tests/test_transform.py`

- [ ] **Step 1: Write the failing test**

```python
import pandas as pd
from transform import normalize


def _raw_row(**kw):
    base = {
        "TBD_StudentName": "Jane Doe",
        "TBD_Center": "Teaneck",
        "TBD_MembershipType": "Summer 2026 (Sessions Package)",
        "TBD_EnrollmentStart": "6/15/2026",
        "TBD_EnrollmentStatus": "Completed",
        "TBD_SessionsPerWeek": "",
    }
    base.update(kw)
    return base


def test_normalize_renames_and_types():
    raw = pd.DataFrame([_raw_row()])
    df = normalize(raw)
    assert list(df.columns) == [
        "student", "center", "membership", "start_date", "status_raw",
        "status_token", "sessions_wk",
    ]
    assert df.loc[0, "student"] == "Jane Doe"
    assert df.loc[0, "start_date"] == pd.Timestamp("2026-06-15")
    assert df.loc[0, "status_token"] == "done"


def test_normalize_maps_status_tokens():
    rows = pd.DataFrame([
        _raw_row(TBD_EnrollmentStatus="Enrolled"),
        _raw_row(TBD_EnrollmentStatus="On Hold"),
        _raw_row(TBD_EnrollmentStatus="Cancelled"),
        _raw_row(TBD_EnrollmentStatus="Weird New Status"),
    ])
    df = normalize(rows)
    assert list(df["status_token"]) == ["enrolled", "hold", "cancelled", "other"]
```

- [ ] **Step 2: Run to verify it fails**

Run: `cd ~/mathnasium-summer-status && python -m pytest tests/test_transform.py -v`
Expected: FAIL — `ImportError: cannot import name 'normalize'`

- [ ] **Step 3: Write `normalize`**

```python
"""Pure transforms over the Enrollment Report exports."""
import pandas as pd
import config as c

_CANON_COLS = ["student", "center", "membership", "start_date",
               "status_raw", "status_token", "sessions_wk"]


def _status_token(raw: str) -> str:
    s = str(raw).strip().lower()
    if s in c.STATUS_ENROLLED:
        return "enrolled"
    if s in c.STATUS_HOLD:
        return "hold"
    if s in c.STATUS_CANCELLED:
        return "cancelled"
    if s in c.STATUS_DONE:
        return "done"
    return "other"


def _student_name(row) -> str:
    if c.COL_FIRST and c.COL_LAST:
        return f"{row[c.COL_FIRST]} {row[c.COL_LAST]}".strip()
    return str(row[c.COL_STUDENT]).strip()


def normalize(raw: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame()
    out["student"] = raw.apply(_student_name, axis=1)
    out["center"] = raw[c.COL_CENTER].astype(str).str.strip()
    out["membership"] = raw[c.COL_MEMBERSHIP].astype(str).str.strip()
    out["start_date"] = pd.to_datetime(raw[c.COL_START_DATE], errors="coerce")
    out["status_raw"] = raw[c.COL_STATUS].astype(str).str.strip()
    out["status_token"] = out["status_raw"].map(_status_token)
    if c.COL_SESSIONS_WK:
        out["sessions_wk"] = raw[c.COL_SESSIONS_WK].astype(str).str.strip()
    else:
        out["sessions_wk"] = ""
    return out[_CANON_COLS]


def load_center(path: str) -> pd.DataFrame:
    return normalize(pd.read_excel(path))
```

- [ ] **Step 4: Run to verify it passes**

Run: `python -m pytest tests/test_transform.py -v`
Expected: PASS (both tests)

- [ ] **Step 5: Commit**

```bash
git add transform.py tests/test_transform.py
git commit -m "feat: normalize Enrollment Report exports"
```

---

## Task 5: `transform.py` — classify each summer student

**Files:**
- Modify: `~/mathnasium-summer-status/transform.py`
- Test: `~/mathnasium-summer-status/tests/test_transform.py`

Classification rules (from the spec):

| Bucket key | Label | Rule |
|---|---|---|
| `converted_active` | Converted – active | non-summer row, token `enrolled`, start >= summer start |
| `converted_hold` | Converted – on hold | non-summer row, token `hold`, start >= summer start |
| `summer_active` | Summer still active | summer row token `enrolled`; no non-summer row qualifying above |
| `not_returned` | Completed – not returned | summer row token `done`; no later non-summer row |
| `cancelled` | Cancelled | summer row token `cancelled`; no later non-summer row |
| `needs_review` | Other / needs review | none of the above (e.g. summer row token `other`) |

If a student has a qualifying converted row, the converted bucket wins regardless of the summer row's status. If both a `converted_active` and `converted_hold` row exist, `converted_hold` wins (matches spec: "converted and then went on hold"). Among multiple non-summer rows, "Converted to" detail comes from the one with the latest start date.

- [ ] **Step 1: Write the failing tests**

```python
from transform import classify_students
import pandas as pd


def _df(rows):
    return pd.DataFrame(rows, columns=[
        "student", "center", "membership", "start_date",
        "status_raw", "status_token", "sessions_wk",
    ])


SUMMER = "Summer 2026 (Sessions Package)"


def test_converted_active():
    df = _df([
        ["A", "Teaneck", SUMMER, pd.Timestamp("2026-06-01"), "Completed", "done", ""],
        ["A", "Teaneck", "School Year 2x", pd.Timestamp("2026-09-01"), "Enrolled", "enrolled", "2"],
    ])
    res = classify_students(df)
    row = res.loc[res["student"] == "A"].iloc[0]
    assert row["bucket"] == "converted_active"
    assert row["converted_to"] == "School Year 2x"
    assert row["converted_start"] == pd.Timestamp("2026-09-01")
    assert row["converted_sessions_wk"] == "2"


def test_converted_hold_wins_over_active():
    df = _df([
        ["B", "Teaneck", SUMMER, pd.Timestamp("2026-06-01"), "Completed", "done", ""],
        ["B", "Teaneck", "School Year 2x", pd.Timestamp("2026-09-01"), "Enrolled", "enrolled", "2"],
        ["B", "Teaneck", "School Year 2x", pd.Timestamp("2026-10-01"), "On Hold", "hold", "2"],
    ])
    row = classify_students(df).iloc[0]
    assert row["bucket"] == "converted_hold"
    # detail from latest-start non-summer row
    assert row["converted_start"] == pd.Timestamp("2026-10-01")


def test_summer_still_active():
    df = _df([
        ["C", "Englewood", SUMMER, pd.Timestamp("2026-07-01"), "Enrolled", "enrolled", ""],
    ])
    assert classify_students(df).iloc[0]["bucket"] == "summer_active"


def test_completed_not_returned():
    df = _df([
        ["D", "Englewood", SUMMER, pd.Timestamp("2026-06-01"), "Completed", "done", ""],
    ])
    assert classify_students(df).iloc[0]["bucket"] == "not_returned"


def test_cancelled():
    df = _df([
        ["E", "Teaneck", SUMMER, pd.Timestamp("2026-06-01"), "Cancelled", "cancelled", ""],
    ])
    assert classify_students(df).iloc[0]["bucket"] == "cancelled"


def test_needs_review_on_unknown_status():
    df = _df([
        ["F", "Teaneck", SUMMER, pd.Timestamp("2026-06-01"), "Frozen", "other", ""],
    ])
    assert classify_students(df).iloc[0]["bucket"] == "needs_review"


def test_non_summer_row_before_summer_start_does_not_count():
    df = _df([
        ["G", "Teaneck", "School Year 2x", pd.Timestamp("2026-01-01"), "Cancelled", "cancelled", "2"],
        ["G", "Teaneck", SUMMER, pd.Timestamp("2026-06-01"), "Completed", "done", ""],
    ])
    assert classify_students(df).iloc[0]["bucket"] == "not_returned"


def test_one_row_per_student_in_output():
    df = _df([
        ["H", "Teaneck", SUMMER, pd.Timestamp("2026-06-01"), "Completed", "done", ""],
        ["H", "Teaneck", SUMMER, pd.Timestamp("2026-07-15"), "Completed", "done", ""],
    ])
    res = classify_students(df)
    assert len(res) == 1
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_transform.py -v -k classify or converted or summer_still or completed or cancelled or needs_review or non_summer or one_row`
Expected: FAIL — `cannot import name 'classify_students'`

- [ ] **Step 3: Implement `classify_students`**

```python
BUCKET_LABELS = {
    "converted_active": "Converted – active",
    "converted_hold": "Converted – on hold",
    "summer_active": "Summer still active",
    "not_returned": "Completed – not returned",
    "cancelled": "Cancelled",
    "needs_review": "Other / needs review",
}

_OUT_COLS = ["student", "center", "summer_start", "summer_status_raw", "bucket",
             "bucket_label", "converted_to", "converted_start",
             "converted_sessions_wk"]


def _classify_one(g: pd.DataFrame) -> dict:
    is_summer = g["membership"].str.strip() == c.SUMMER_MEMBERSHIP
    summer = g[is_summer].sort_values("start_date")
    non_summer = g[~is_summer]
    summer_start = summer["start_date"].min()
    summer_status_raw = summer.iloc[-1]["status_raw"] if len(summer) else ""

    qualifying = non_summer[non_summer["start_date"] >= summer_start]
    conv_active = qualifying[qualifying["status_token"] == "enrolled"]
    conv_hold = qualifying[qualifying["status_token"] == "hold"]

    result = {
        "student": g.iloc[0]["student"],
        "center": g.iloc[0]["center"],
        "summer_start": summer_start,
        "summer_status_raw": summer_status_raw,
        "converted_to": "",
        "converted_start": pd.NaT,
        "converted_sessions_wk": "",
    }

    if len(conv_hold) or len(conv_active):
        bucket = "converted_hold" if len(conv_hold) else "converted_active"
        detail = qualifying.sort_values("start_date").iloc[-1]
        result.update({
            "bucket": bucket,
            "converted_to": detail["membership"],
            "converted_start": detail["start_date"],
            "converted_sessions_wk": detail["sessions_wk"],
        })
    else:
        summer_token = summer.iloc[-1]["status_token"] if len(summer) else "other"
        bucket = {
            "enrolled": "summer_active",
            "done": "not_returned",
            "cancelled": "cancelled",
        }.get(summer_token, "needs_review")
        result["bucket"] = bucket

    result["bucket_label"] = BUCKET_LABELS[result["bucket"]]
    return result


def classify_students(df: pd.DataFrame) -> pd.DataFrame:
    """One row per student who has any Summer membership row."""
    summer_students = set(
        df.loc[df["membership"].str.strip() == c.SUMMER_MEMBERSHIP, "student"]
    )
    sub = df[df["student"].isin(summer_students)]
    records = [
        _classify_one(g.reset_index(drop=True))
        for _, g in sub.groupby("student", sort=True)
    ]
    return pd.DataFrame(records, columns=_OUT_COLS)
```

- [ ] **Step 4: Run to verify it passes**

Run: `python -m pytest tests/test_transform.py -v`
Expected: PASS (all tests)

- [ ] **Step 5: Commit**

```bash
git add transform.py tests/test_transform.py
git commit -m "feat: classify summer students by status and conversion"
```

---

## Task 6: `transform.py` — summaries

**Files:**
- Modify: `~/mathnasium-summer-status/transform.py`
- Test: `~/mathnasium-summer-status/tests/test_transform.py`

- [ ] **Step 1: Write the failing test**

```python
from transform import build_summary


def test_build_summary_counts_and_conversion_rate():
    classified = pd.DataFrame([
        {"student": "A", "center": "Teaneck", "bucket": "converted_active"},
        {"student": "B", "center": "Teaneck", "bucket": "converted_hold"},
        {"student": "C", "center": "Teaneck", "bucket": "not_returned"},
        {"student": "D", "center": "Teaneck", "bucket": "cancelled"},
    ])
    s = build_summary(classified)
    assert s["total"] == 4
    assert s["counts"]["converted_active"] == 1
    assert s["counts"]["needs_review"] == 0
    assert round(s["conversion_rate"], 2) == 0.50
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_transform.py -v -k build_summary`
Expected: FAIL — `cannot import name 'build_summary'`

- [ ] **Step 3: Implement `build_summary`**

```python
def build_summary(classified: pd.DataFrame) -> dict:
    total = len(classified)
    counts = {k: int((classified["bucket"] == k).sum()) for k in BUCKET_LABELS}
    converted = counts["converted_active"] + counts["converted_hold"]
    return {
        "total": total,
        "counts": counts,
        "converted": converted,
        "conversion_rate": (converted / total) if total else 0.0,
        "needs_review": counts["needs_review"],
    }
```

- [ ] **Step 4: Run to verify it passes**

Run: `python -m pytest tests/test_transform.py -v`
Expected: PASS (all tests)

- [ ] **Step 5: Commit**

```bash
git add transform.py tests/test_transform.py
git commit -m "feat: per-center summary with conversion rate"
```

---

## Task 7: `report.py` — render the HTML

**Files:**
- Create: `~/mathnasium-summer-status/report.py`
- Test: `~/mathnasium-summer-status/tests/test_report.py`

- [ ] **Step 1: Write the failing test**

```python
import pandas as pd
from report import render_html


def _classified():
    return pd.DataFrame([
        {"student": "Anna A", "center": "Teaneck", "summer_start": pd.Timestamp("2026-06-01"),
         "summer_status_raw": "Completed", "bucket": "converted_active",
         "bucket_label": "Converted – active", "converted_to": "School Year 2x",
         "converted_start": pd.Timestamp("2026-09-02"), "converted_sessions_wk": "2"},
        {"student": "Ben B", "center": "Englewood", "summer_start": pd.Timestamp("2026-07-01"),
         "summer_status_raw": "Completed", "bucket": "not_returned",
         "bucket_label": "Completed – not returned", "converted_to": "",
         "converted_start": pd.NaT, "converted_sessions_wk": ""},
    ])


def test_render_html_has_both_center_sections_and_names():
    html = render_html(_classified(), generated_at="2026-09-07 10:00")
    assert "<h2>Teaneck</h2>" in html
    assert "<h2>Englewood</h2>" in html
    assert "Anna A" in html and "Ben B" in html
    assert "School Year 2x" in html
    assert "2026-09-07 10:00" in html
    # self-contained: no external stylesheet links
    assert "<link" not in html


def test_render_html_flags_needs_review_count():
    df = _classified()
    df.loc[0, "bucket"] = "needs_review"
    df.loc[0, "bucket_label"] = "Other / needs review"
    html = render_html(df, generated_at="x")
    assert "needs review" in html.lower()
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_report.py -v`
Expected: FAIL — `No module named 'report'`

- [ ] **Step 3: Implement `report.py`**

```python
"""Render the self-contained HTML status report."""
import webbrowser
from datetime import datetime
import pandas as pd
import config as c
from transform import BUCKET_LABELS, build_summary

_CSS = """
body{font-family:-apple-system,Segoe UI,Arial,sans-serif;margin:2rem;color:#222;background:#fff}
h1{font-size:1.4rem} h2{font-size:1.15rem;margin-top:2rem;border-bottom:2px solid #ddd}
table{border-collapse:collapse;margin:0.75rem 0;font-size:0.9rem}
th,td{border:1px solid #ccc;padding:4px 8px;text-align:left}
th{background:#f2f2f2}
.rate{font-weight:bold}
.flag{background:#fff4e5;border:1px solid #f0c36d;padding:6px 10px;margin:0.5rem 0}
"""

_ORDER = ["converted_active", "converted_hold", "summer_active",
          "not_returned", "cancelled", "needs_review"]


def _fmt_date(v):
    return "" if pd.isna(v) else pd.Timestamp(v).strftime("%Y-%m-%d")


def _summary_table(per_center: dict, overall: dict) -> str:
    centers = list(per_center)
    head = "".join(f"<th>{name}</th>" for name in centers) + "<th>Overall</th>"
    rows = ""
    for key in _ORDER:
        cells = "".join(
            f"<td>{per_center[name]['counts'][key]}</td>" for name in centers
        )
        rows += f"<tr><td>{BUCKET_LABELS[key]}</td>{cells}<td>{overall['counts'][key]}</td></tr>"
    tot = "".join(f"<td>{per_center[name]['total']}</td>" for name in centers)
    rate = "".join(
        f"<td class='rate'>{per_center[name]['conversion_rate']:.0%}</td>" for name in centers
    )
    rows += f"<tr><td><b>Total summer students</b></td>{tot}<td>{overall['total']}</td></tr>"
    rows += (f"<tr><td class='rate'>Conversion rate</td>{rate}"
             f"<td class='rate'>{overall['conversion_rate']:.0%}</td></tr>")
    return f"<table><tr><th></th>{head}</tr>{rows}</table>"


def _detail_table(df: pd.DataFrame) -> str:
    order = {k: i for i, k in enumerate(_ORDER)}
    df = df.sort_values(by=["bucket", "student"],
                        key=lambda s: s.map(order) if s.name == "bucket" else s)
    head = ("<tr><th>Student</th><th>Summer start</th><th>Summer status</th>"
            "<th>Current status</th><th>Converted to</th><th>Conv. start</th>"
            "<th>Sessions/wk</th></tr>")
    body = ""
    for _, r in df.iterrows():
        body += (
            f"<tr><td>{r['student']}</td><td>{_fmt_date(r['summer_start'])}</td>"
            f"<td>{r['summer_status_raw']}</td><td>{r['bucket_label']}</td>"
            f"<td>{r['converted_to']}</td><td>{_fmt_date(r['converted_start'])}</td>"
            f"<td>{r['converted_sessions_wk']}</td></tr>"
        )
    return f"<table>{head}{body}</table>"


def render_html(classified: pd.DataFrame, generated_at: str) -> str:
    centers = list(c.CENTERS)
    per_center = {name: build_summary(classified[classified["center"] == name])
                  for name in centers}
    overall = build_summary(classified)

    parts = [
        f"<!doctype html><html><head><meta charset='utf-8'>",
        f"<title>Summer 2026 Enrollment Status</title><style>{_CSS}</style></head><body>",
        f"<h1>Summer 2026 (Sessions Package) &mdash; Enrollment Status</h1>",
        f"<p>Generated {generated_at}</p>",
    ]
    if overall["needs_review"]:
        parts.append(f"<div class='flag'>{overall['needs_review']} student(s) "
                     f"could not be auto-classified &mdash; see “Other / needs review” rows.</div>")
    parts.append("<h2>Overall summary</h2>")
    parts.append(_summary_table(per_center, overall))

    for name in centers:
        sub = classified[classified["center"] == name]
        parts.append(f"<h2>{name}</h2>")
        s = per_center[name]
        parts.append(f"<p class='rate'>Conversion rate: {s['conversion_rate']:.0%} "
                     f"({s['converted']} of {s['total']})</p>")
        parts.append(_detail_table(sub))

    parts.append("</body></html>")
    return "".join(parts)


def write_and_open(classified: pd.DataFrame) -> str:
    c.OUTPUT_DIR.mkdir(exist_ok=True)
    html = render_html(classified, datetime.now().strftime("%Y-%m-%d %H:%M"))
    out = c.OUTPUT_DIR / "summer_2026_status.html"
    out.write_text(html, encoding="utf-8")
    webbrowser.open(out.as_uri())
    return str(out)
```

- [ ] **Step 4: Run to verify it passes**

Run: `python -m pytest tests/test_report.py -v`
Expected: PASS (both tests)

- [ ] **Step 5: Commit**

```bash
git add report.py tests/test_report.py
git commit -m "feat: render self-contained HTML status report"
```

---

## Task 8: `main.py` — orchestrator

**Files:**
- Create: `~/mathnasium-summer-status/main.py`

- [ ] **Step 1: Write `main.py`**

```python
"""Summer package status report — CLI entry point.

Usage:
    export $(cat ~/.mathnasium_env | xargs)
    python main.py                # pull fresh + render
    python main.py --no-download  # re-render from cached input/
"""
import argparse
import sys
import pandas as pd
import config as c
from transform import load_center, classify_students
from report import write_and_open


def run(no_download: bool) -> None:
    if not no_download:
        from download import download_all
        download_all()

    frames = []
    for name in c.CENTERS:
        path = c.INPUT_DIR / f"EnrollmentReport_{name}.xlsx"
        if not path.exists():
            sys.exit(f"Missing {path}. Run without --no-download first.")
        df = load_center(str(path))
        if df.empty:
            sys.exit(f"{name}: Enrollment Report returned zero rows — aborting.")
        frames.append(df)

    all_rows = pd.concat(frames, ignore_index=True)
    classified = classify_students(all_rows)
    if classified.empty:
        sys.exit(f"No rows matched membership '{c.SUMMER_MEMBERSHIP}' in either "
                 f"center. The Radius label may have changed — check config.SUMMER_MEMBERSHIP.")

    out = write_and_open(classified)
    n_review = int((classified["bucket"] == "needs_review").sum())
    print(f"Wrote {out} — {len(classified)} summer students"
          + (f", {n_review} need manual review." if n_review else "."))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-download", action="store_true",
                    help="re-render from cached input/ without hitting Radius")
    args = ap.parse_args()
    run(no_download=args.no_download)
```

- [ ] **Step 2: Full run against real Radius**

```bash
cd ~/mathnasium-summer-status
export $(cat ~/.mathnasium_env | xargs)
python main.py
```

Expected: both reports download, the browser opens `output/summer_2026_status.html`, both center sections are populated, counts look plausible (cross-check the total summer count against a manual Radius filter on the membership type).

- [ ] **Step 3: Test `--no-download`**

```bash
python main.py --no-download
```

Expected: no Radius login, report regenerates from cached files in seconds.

- [ ] **Step 4: Sanity-check the classification**

Pick 2-3 students you know personally (one who converted, one who didn't) and confirm the report puts them in the right bucket. Eyeball the "Other / needs review" list — if it's large, the `STATUS_*` sets in `config.py` need more entries; add them and re-run `--no-download`.

- [ ] **Step 5: Commit**

```bash
git add main.py
git commit -m "feat: CLI orchestrator for summer package status report"
```

---

## Task 9: README + final commit

**Files:**
- Create: `~/mathnasium-summer-status/README.md`

- [ ] **Step 1: Write `README.md`**

```markdown
# Summer Package Status Report

On-demand HTML report of every "Summer 2026 (Sessions Package)" enrollment at
Teaneck and Englewood, showing each student's current status and whether they
converted to a school-year membership.

## Run

```bash
cd ~/mathnasium-summer-status
export $(cat ~/.mathnasium_env | xargs)
pip install -r requirements.txt        # first time
playwright install chromium            # first time
python main.py                         # pull fresh + open report
python main.py --no-download           # re-render from cached input/
```

Output: `output/summer_2026_status.html` (opens automatically).

## When the summer label changes next year

Edit `SUMMER_MEMBERSHIP` in `config.py` to match the new Radius membership type
string exactly.

## If Radius changes the Enrollment Report

Re-run `python discover_selectors.py` and update the `COL_*` / `SEL_*` constants
in `config.py`. See `DISCOVERY_NOTES.md` for the last known-good field list.
```

- [ ] **Step 2: Run the full test suite**

Run: `cd ~/mathnasium-summer-status && python -m pytest -v`
Expected: PASS (all tests in `tests/`)

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: usage README"
```

- [ ] **Step 4: Create the GitHub repo under the mdiamond77 org**

```bash
cd ~/mathnasium-summer-status
gh repo create mdiamond77/mathnasium-summer-status --private --source=. --push
```

---

## Self-Review Notes

**Spec coverage:**
- Data source / one pull per center / all statuses → Task 3
- Discovery gate → Task 2 (with explicit STOP condition)
- Five status buckets + Other/needs-review → Task 5 (tests cover each)
- Conversion detail (type, start, sessions/wk) → Task 5
- Edge cases (multiple non-summer, converted-then-hold, unknown status) → Task 5 tests
- HTML layout: header, overall summary side-by-side, per-center sections, conversion rate → Task 7
- Fail-loud error handling (login, zero rows, zero membership matches, needs-review count) → Tasks 3 & 8
- CLI `--no-download` → Task 8
- Pure-function tests with fixtures → Tasks 4-6
- New repo under `mdiamond77`, no workflow → Tasks 1 & 9

**Known deferred decisions (resolved during Task 2, not plan failures):**
- Exact date-field selector ids on the Enrollment Report
- Exact column header strings (the `TBD_` constants)
- Whether "sessions per week" exists as a column at all (spec already flags this may be blank for the summer row)
- The full set of raw status strings Radius emits

**Type consistency:** `classify_students` output columns (`bucket`, `bucket_label`, `converted_to`, `converted_start`, `converted_sessions_wk`, `summer_start`, `summer_status_raw`) are used identically in `report.py`. `build_summary` keys (`total`, `counts`, `converted`, `conversion_rate`, `needs_review`) match between Task 6 and Task 7.
