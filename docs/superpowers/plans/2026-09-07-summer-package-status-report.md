# Summer Package Status Report Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an on-demand CLI tool that pulls the Radius Enrollment Report for Teaneck and Englewood, finds every 2026 summer-package enrollment, classifies each student by current status and school-year conversion, and renders a single self-contained HTML report.

**Architecture:** One Playwright headless pull of the Enrollment Report per center (4/1/2026 → 12/31/2026, all statuses). A pure-function transform layer (pandas) identifies summer students and buckets each one using only that student's other enrollment rows. A render layer emits one portable HTML file with an overall summary plus a per-center section. Run by hand; no GitHub Actions.

**Tech Stack:** Python 3.11+, Playwright (headless Chromium), pandas, openpyxl, pytest. Radius credentials from `~/.mathnasium_env`.

---

## References

- **Design spec:** `~/Mathnasium_automation/docs/superpowers/specs/2026-09-07-summer-package-status-report-design.md` — read the "Discovery findings" and "Status Classification" sections before starting.
- **Existing patterns:** `~/mathnasium-new-enrollments/` (closest cousin — `download.py`, `config.py`) and `~/mathnasium-binder-audit/discover_selectors.py`.
- **Discovery already done (2026-09-07):** the gate passed. Real exports for both centers are cached at `~/mathnasium-summer-status/input/EnrollmentReport_Teaneck.xlsx` and `..._Englewood.xlsx` — use them as fixtures and for `--no-download` runs. The discovery script is at `~/mathnasium-summer-status/discover_selectors.py` (created in Task 2).

## Known facts from discovery (do not re-derive)

- **Export columns (verbatim):** `Lead Id`, `Account Id`, `Student First Name`, `Student Last Name`, `Grade`, `Grade Range`, `Account Name`, `Center`, `Status`, `Membership Type`, `Primary Enrollment Start`, `Primary Enrollment End`, `Recurring`, `Enrollment Contract Length`, `Enrollment Length of Stay`, `Student Length of Stay`, `Total Sessions`, `Remaining`, `Session Length`, `Hold Count`, `Total Hold Length`, `Delivery`, `Monthly Amount`, `Expected Monthly Amount`, `Virtual Center`, `Guardians`, `Guardian Emails`, `Guardian Phone Numbers`.
- `Primary Enrollment Start` / `Primary Enrollment End` are **per-enrollment** (per row).
- One row per enrollment. Student key = `Student First Name` + `Student Last Name` + `Center`, normalized (strip, collapse inner whitespace, casefold).
- `Status` ∈ {`Pre-Enrolled`, `Enrolled`, `On Hold`, `Inactive`}. Nothing else. `Enrolled` does **not** imply currently attending — use `Primary Enrollment End` >= today.
- Summer membership `Membership Type` value (strip a leading `"* "` first):
  - Teaneck: `Summer 2026 (Sessions Package)`
  - Englewood: `2026 Summer Sessions Package (Sessions Package)`
- School Partnership = `Membership Type` contains substring `School Partnership`.
- No sessions/week column — show `Membership Type` name + `Monthly Amount` as conversion detail.
- Selectors: `#UserName` `#Password` `#login`; center `#AllCenterListMultiSelect` (Kendo MultiSelect); dates `#StartDate` `#EndDate` (Kendo DatePicker); `#btnsearch`; `#btnExport`; grid `#gridEnrollmentReport`. Center ids: Teaneck `2871`, Englewood `2428`.

---

## File Structure

| File | Responsibility |
|---|---|
| `~/mathnasium-summer-status/config.py` | Constants: URLs, selectors, center config, per-center summer strings, pull window, column-name constants |
| `~/mathnasium-summer-status/discover_selectors.py` | The headless discovery script (kept for re-running if Radius changes) |
| `~/mathnasium-summer-status/download.py` | Playwright headless: login, pull Enrollment Report per center, save `input/EnrollmentReport_<center>.xlsx` |
| `~/mathnasium-summer-status/transform.py` | Pure functions: load/normalize, identify summer students, classify, summarize |
| `~/mathnasium-summer-status/report.py` | Render `output/summer_2026_status.html`; open it |
| `~/mathnasium-summer-status/main.py` | CLI orchestrator (`--no-download`) |
| `~/mathnasium-summer-status/tests/test_transform.py` | Unit tests for classification |
| `~/mathnasium-summer-status/tests/test_report.py` | Unit tests for the renderer |

Task 1 (scaffold) is **already complete** — the directory, git repo, `requirements.txt`, `.gitignore`, a placeholder `config.py`, and `tests/__init__.py` exist and are committed. Task 2 rewrites `config.py` with real values.

---

## Task 2: `config.py` with real values + save the discovery script

**Files:**
- Modify: `~/mathnasium-summer-status/config.py` (replace entirely)
- Create: `~/mathnasium-summer-status/discover_selectors.py`

- [ ] **Step 1: Replace `config.py` with:**

```python
from pathlib import Path

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
SEL_START_DATE = "#StartDate"
SEL_END_DATE = "#EndDate"
SEL_SEARCH_BTN = "#btnsearch"
SEL_EXPORT_BTN = "#btnExport"
SEL_GRID = "#gridEnrollmentReport"

# Each center: Radius id + the exact Membership Type string for its 2026 summer
# package (as it appears in the export, after stripping a leading "* ").
CENTERS = {
    "Teaneck":   {"radius_id": "2871",
                  "summer_membership": "Summer 2026 (Sessions Package)"},
    "Englewood": {"radius_id": "2428",
                  "summer_membership": "2026 Summer Sessions Package (Sessions Package)"},
}

# Pull window passed to the report's date fields (year, month0-indexed, day).
PULL_START = (2026, 3, 1)    # 4/1/2026
PULL_END = (2026, 11, 31)    # 12/31/2026

# A membership type is a "school partnership" continuation if its name contains:
SCHOOL_PARTNERSHIP_MARKER = "School Partnership"

# ── Export column names (verbatim, confirmed 2026-09-07) ──────────────────
COL_FIRST      = "Student First Name"
COL_LAST       = "Student Last Name"
COL_GRADE      = "Grade"
COL_CENTER     = "Center"
COL_STATUS     = "Status"
COL_MEMBERSHIP = "Membership Type"
COL_START_DATE = "Primary Enrollment Start"
COL_END_DATE   = "Primary Enrollment End"
COL_MONTHLY    = "Monthly Amount"

# Known Radius enrollment statuses. Anything else -> the student is flagged
# "needs review".
KNOWN_STATUSES = {"Pre-Enrolled", "Enrolled", "On Hold", "Inactive"}
```

- [ ] **Step 2: Create `discover_selectors.py`** (headless; kept for future re-runs)

```python
"""Headless discovery for the Enrollment Report. Already run 2026-09-07; kept in
case Radius changes the report. Dumps element ids, membership-type options, grid
fields, and saves a Teaneck export to input/.

Run:  RADIUS_USERNAME=... RADIUS_PASSWORD=... python discover_selectors.py
"""
import os
import json
from playwright.sync_api import sync_playwright
import config as c


def main():
    user = os.environ["RADIUS_USERNAME"]
    pw = os.environ["RADIUS_PASSWORD"]
    c.INPUT_DIR.mkdir(exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        with browser:
            ctx = browser.new_context(accept_downloads=True,
                                      viewport={"width": 1500, "height": 1000})
            with ctx:
                page = ctx.new_page()
                page.goto(c.LOGIN_URL, wait_until="networkidle")
                page.fill(c.SEL_USERNAME, user)
                page.fill(c.SEL_PASSWORD, pw)
                page.click(c.SEL_LOGIN_BTN)
                page.wait_for_load_state("networkidle")
                page.goto(c.ENROLLMENT_REPORT_URL, wait_until="networkidle")
                page.wait_for_timeout(3000)

                els = page.evaluate(
                    "Array.from(document.querySelectorAll('[id]')).map(e => "
                    "({id:e.id, tag:e.tagName, type:e.type||'', name:e.name||''}))"
                )
                print(json.dumps(els, indent=2))

                opts = page.evaluate(
                    "Array.from(document.querySelectorAll('#membershipTypeMultiSelect option'))"
                    ".map(o => o.text.trim())"
                )
                print("MEMBERSHIP TYPE OPTIONS:", json.dumps(opts, indent=2))

                page.evaluate(
                    "() => { const ms=$('#AllCenterListMultiSelect').data('kendoMultiSelect');"
                    "ms.value(['2871']); ms.trigger('change');"
                    "$('#StartDate').data('kendoDatePicker').value(new Date(2026,3,1));"
                    "$('#EndDate').data('kendoDatePicker').value(new Date(2026,11,31)); }"
                )
                page.wait_for_timeout(500)
                page.click(c.SEL_SEARCH_BTN)
                page.wait_for_timeout(7000)
                with page.expect_download(timeout=60000) as dl:
                    page.click(c.SEL_EXPORT_BTN)
                dl.value.save_as(str(c.INPUT_DIR / "EnrollmentReport_Teaneck.xlsx"))
                print("saved input/EnrollmentReport_Teaneck.xlsx")


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Byte-compile check**

Run: `cd ~/mathnasium-summer-status && python -c "import config; import ast; ast.parse(open('discover_selectors.py').read()); print('ok')"`
Expected: `ok`

- [ ] **Step 4: Commit**

```bash
cd ~/mathnasium-summer-status
git add config.py discover_selectors.py
git commit -m "chore: real config values + discovery script from 2026-09-07 findings"
```

---

## Task 3: `download.py` — pull the Enrollment Report per center

**Files:**
- Create: `~/mathnasium-summer-status/download.py`

- [ ] **Step 1: Write `download.py`**

```python
"""Pull the Radius Enrollment Report for each center; cache the raw export."""
import os
from playwright.sync_api import sync_playwright
import config as c


def _login(page):
    page.goto(c.LOGIN_URL, wait_until="networkidle")
    page.fill(c.SEL_USERNAME, os.environ["RADIUS_USERNAME"])
    page.fill(c.SEL_PASSWORD, os.environ["RADIUS_PASSWORD"])
    page.click(c.SEL_LOGIN_BTN)
    page.wait_for_load_state("networkidle")
    if page.query_selector(c.SEL_USERNAME):
        raise RuntimeError("Radius login failed — still on the login page")


def _pull_center(page, radius_id: str, out_path):
    page.goto(c.ENROLLMENT_REPORT_URL, wait_until="networkidle")
    page.wait_for_timeout(3000)
    sy, sm, sd = c.PULL_START
    ey, em, ed = c.PULL_END
    page.evaluate(
        """([id, s, e]) => {
            const ms = $('#AllCenterListMultiSelect').data('kendoMultiSelect');
            ms.value([id]); ms.trigger('change');
            $('#StartDate').data('kendoDatePicker').value(new Date(s[0], s[1], s[2]));
            $('#EndDate').data('kendoDatePicker').value(new Date(e[0], e[1], e[2]));
        }""",
        [radius_id, [sy, sm, sd], [ey, em, ed]],
    )
    page.wait_for_timeout(500)
    page.click(c.SEL_SEARCH_BTN)
    page.wait_for_timeout(7000)
    page.wait_for_load_state("networkidle")
    with page.expect_download(timeout=60000) as dl:
        page.click(c.SEL_EXPORT_BTN)
    dl.value.save_as(str(out_path))


def download_all() -> dict:
    """Pull both centers. Returns {center_name: Path}. Raises on login failure."""
    c.INPUT_DIR.mkdir(exist_ok=True)
    paths = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        with browser:
            ctx = browser.new_context(accept_downloads=True,
                                      viewport={"width": 1500, "height": 1000})
            with ctx:
                page = ctx.new_page()
                _login(page)
                for name, cfg in c.CENTERS.items():
                    out = c.INPUT_DIR / f"EnrollmentReport_{name}.xlsx"
                    _pull_center(page, cfg["radius_id"], out)
                    paths[name] = out
    return paths


if __name__ == "__main__":
    for name, path in download_all().items():
        print(f"{name}: {path}")
```

- [ ] **Step 2: Run it against real Radius**

```bash
cd ~/mathnasium-summer-status
export RADIUS_USERNAME="$(grep '^RADIUS_USERNAME=' ~/.mathnasium_env | cut -d= -f2-)"
export RADIUS_PASSWORD="$(grep '^RADIUS_PASSWORD=' ~/.mathnasium_env | cut -d= -f2-)"
python download.py
```

Expected: prints `Teaneck: .../EnrollmentReport_Teaneck.xlsx` and the Englewood line; both files open in Excel with the 28 known columns and ~150–170 rows each. (Note: `~/.mathnasium_env` also contains a `GOOGLE_SERVICE_ACCOUNT_JSON` value with spaces, so `export $(cat ... | xargs)` fails — extract the two RADIUS vars individually as shown.)

- [ ] **Step 3: Fix selectors until it works.** Likely failure points: the Kendo `.data('kendoMultiSelect')` / `.data('kendoDatePicker')` calls if jQuery (`$`) isn't ready — add `page.wait_for_function("typeof window.$ === 'function'")` before the `evaluate` if so. Re-run until both files download with the right center's data (spot-check a couple of student names against the Radius UI).

- [ ] **Step 4: Commit**

```bash
git add download.py
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

SUMMER_T = "Summer 2026 (Sessions Package)"


def _raw(**kw):
    base = {
        "Student First Name": " Jane ", "Student Last Name": "Doe",
        "Grade": "5", "Center": "Teaneck", "Status": "Inactive",
        "Membership Type": "* " + SUMMER_T,
        "Primary Enrollment Start": "6/15/2026",
        "Primary Enrollment End": "8/31/2026",
        "Monthly Amount": 0,
    }
    base.update(kw)
    return base


def test_normalize_shapes_and_types():
    df = normalize(pd.DataFrame([_raw()]))
    assert list(df.columns) == [
        "student_key", "student", "grade", "center", "status",
        "membership", "is_school_partnership", "start_date", "end_date",
        "monthly_amount",
    ]
    r = df.iloc[0]
    assert r["student"] == "Jane Doe"
    assert r["student_key"] == "jane doe|teaneck"
    assert r["membership"] == SUMMER_T          # leading "* " stripped
    assert r["start_date"] == pd.Timestamp("2026-06-15")
    assert r["end_date"] == pd.Timestamp("2026-08-31")
    assert r["is_school_partnership"] is False or r["is_school_partnership"] == False


def test_normalize_flags_school_partnership():
    df = normalize(pd.DataFrame([
        _raw(**{"Membership Type": "* Ridgefield Park (School Partnership) (Sessions Package)"})
    ]))
    assert bool(df.iloc[0]["is_school_partnership"]) is True


def test_normalize_collapses_whitespace_in_key():
    df = normalize(pd.DataFrame([_raw(**{"Student First Name": "Mary  Jane"})]))
    assert df.iloc[0]["student_key"] == "mary jane doe|teaneck"
```

- [ ] **Step 2: Run to verify it fails**

Run: `cd ~/mathnasium-summer-status && python -m pytest tests/test_transform.py -v`
Expected: FAIL — `ImportError: cannot import name 'normalize'`

- [ ] **Step 3: Write `normalize`**

```python
"""Pure transforms over an Enrollment Report export."""
import re
import pandas as pd
import config as c

_CANON = ["student_key", "student", "grade", "center", "status", "membership",
          "is_school_partnership", "start_date", "end_date", "monthly_amount"]


def _clean(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


def _strip_star(s) -> str:
    return re.sub(r"^\*\s+", "", _clean(s))


def normalize(raw: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame()
    first = raw[c.COL_FIRST].map(_clean)
    last = raw[c.COL_LAST].map(_clean)
    center = raw[c.COL_CENTER].map(_clean)
    out["student_key"] = (first.str.cat(last, sep=" ").str.casefold()
                          + "|" + center.str.casefold())
    out["student"] = first.str.cat(last, sep=" ")
    out["grade"] = raw[c.COL_GRADE].map(_clean)
    out["center"] = center
    out["status"] = raw[c.COL_STATUS].map(_clean)
    out["membership"] = raw[c.COL_MEMBERSHIP].map(_strip_star)
    out["is_school_partnership"] = out["membership"].str.contains(
        c.SCHOOL_PARTNERSHIP_MARKER, case=False, regex=False
    )
    out["start_date"] = pd.to_datetime(raw[c.COL_START_DATE], errors="coerce")
    out["end_date"] = pd.to_datetime(raw[c.COL_END_DATE], errors="coerce")
    out["monthly_amount"] = pd.to_numeric(raw[c.COL_MONTHLY], errors="coerce")
    return out[_CANON]


def load_center(path) -> pd.DataFrame:
    return normalize(pd.read_excel(path))
```

- [ ] **Step 4: Run to verify it passes**

Run: `python -m pytest tests/test_transform.py -v`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add transform.py tests/test_transform.py
git commit -m "feat: normalize Enrollment Report export"
```

---

## Task 5: `transform.py` — classify each summer student

**Files:**
- Modify: `~/mathnasium-summer-status/transform.py`
- Test: `~/mathnasium-summer-status/tests/test_transform.py`

Buckets, first match wins (see spec):

| key | label |
|---|---|
| `converted_active` | Converted – active |
| `converted_hold` | Converted – on hold |
| `school_partnership` | Continued via School Partnership |
| `summer_active` | Summer still active |
| `did_not_return` | Did not return |
| `needs_review` | Needs review |

- [ ] **Step 1: Write the failing tests**

```python
from datetime import date
from transform import classify_students, BUCKET_LABELS

SUMMER_T = "Summer 2026 (Sessions Package)"
AS_OF = date(2026, 9, 7)


def _row(student, membership, start, end, status, center="Teaneck", sp=False, amt=0):
    return {
        "student_key": student.casefold() + "|" + center.casefold(),
        "student": student, "grade": "5", "center": center, "status": status,
        "membership": membership, "is_school_partnership": sp,
        "start_date": pd.Timestamp(start), "end_date": pd.Timestamp(end),
        "monthly_amount": amt,
    }


def _df(rows):
    return pd.DataFrame(rows)


def test_converted_active():
    df = _df([
        _row("Amy A", SUMMER_T, "2026-06-01", "2026-08-31", "Inactive"),
        _row("Amy A", "1x/week - 2024 Flex (Flexible)", "2026-09-01", "2027-01-31", "Enrolled", amt=329),
    ])
    r = classify_students(df, as_of=AS_OF).iloc[0]
    assert r["bucket"] == "converted_active"
    assert r["continued_as"] == "1x/week - 2024 Flex (Flexible)"
    assert r["new_start"] == pd.Timestamp("2026-09-01")
    assert r["new_monthly"] == 329


def test_converted_hold_beats_active():
    df = _df([
        _row("Ben B", SUMMER_T, "2026-06-01", "2026-08-31", "Inactive"),
        _row("Ben B", "2x/week - 2024 Flexible (Flexible)", "2026-09-01", "2027-01-31", "Enrolled"),
        _row("Ben B", "2x/week - 2024 Flexible (Flexible)", "2026-10-01", "2027-01-31", "On Hold"),
    ])
    r = classify_students(df, as_of=AS_OF).iloc[0]
    assert r["bucket"] == "converted_hold"
    assert r["new_start"] == pd.Timestamp("2026-10-01")


def test_school_partnership_bucket():
    df = _df([
        _row("Cara C", SUMMER_T, "2026-06-01", "2026-08-31", "Inactive"),
        _row("Cara C", "Ridgefield Park (School Partnership) (Sessions Package)",
             "2026-09-05", "2026-12-31", "Enrolled", sp=True),
    ])
    assert classify_students(df, as_of=AS_OF).iloc[0]["bucket"] == "school_partnership"


def test_summer_still_active_by_end_date_not_status():
    df = _df([
        _row("Dan D", SUMMER_T, "2026-08-24", "2026-10-31", "Enrolled"),
    ])
    assert classify_students(df, as_of=AS_OF).iloc[0]["bucket"] == "summer_active"


def test_did_not_return_when_summer_ended_and_nothing_after():
    df = _df([
        _row("Eve E", SUMMER_T, "2026-06-01", "2026-08-31", "Enrolled"),
    ])
    assert classify_students(df, as_of=AS_OF).iloc[0]["bucket"] == "did_not_return"


def test_pre_summer_enrollment_is_not_a_conversion():
    df = _df([
        _row("Fay F", "2x/week - 2024 Flexible (Flexible)", "2026-01-01", "2026-05-31", "Inactive"),
        _row("Fay F", SUMMER_T, "2026-06-01", "2026-08-31", "Inactive"),
    ])
    assert classify_students(df, as_of=AS_OF).iloc[0]["bucket"] == "did_not_return"


def test_unknown_status_goes_to_needs_review():
    df = _df([
        _row("Gus G", SUMMER_T, "2026-06-01", "2026-08-31", "Frozen"),
    ])
    assert classify_students(df, as_of=AS_OF).iloc[0]["bucket"] == "needs_review"


def test_pre_enrolled_only_summer_is_needs_review():
    df = _df([
        _row("Hal H", SUMMER_T, "2026-10-01", "2026-12-31", "Pre-Enrolled"),
    ])
    assert classify_students(df, as_of=AS_OF).iloc[0]["bucket"] == "needs_review"


def test_latest_start_wins_for_detail():
    df = _df([
        _row("Ivy I", SUMMER_T, "2026-06-01", "2026-08-31", "Inactive"),
        _row("Ivy I", "1x/week - 2024 Flex (Flexible)", "2026-09-01", "2027-01-31", "Enrolled", amt=1),
        _row("Ivy I", "2x/week - 2024 Flexible (Flexible)", "2026-09-15", "2027-01-31", "Enrolled", amt=2),
    ])
    r = classify_students(df, as_of=AS_OF).iloc[0]
    assert r["new_monthly"] == 2


def test_one_row_per_student():
    df = _df([
        _row("Jo J", SUMMER_T, "2026-06-01", "2026-07-15", "Inactive"),
        _row("Jo J", SUMMER_T, "2026-07-16", "2026-08-31", "Inactive"),
    ])
    assert len(classify_students(df, as_of=AS_OF)) == 1
```

Add `import pandas as pd` at the top of the test file if not already present.

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_transform.py -v -k classify or converted or summer_still or did_not_return or pre_summer or unknown_status or pre_enrolled or latest_start or one_row`
Expected: FAIL — `cannot import name 'classify_students'`

- [ ] **Step 3: Implement `classify_students`**

```python
from datetime import date

BUCKET_LABELS = {
    "converted_active": "Converted – active",
    "converted_hold": "Converted – on hold",
    "school_partnership": "Continued via School Partnership",
    "summer_active": "Summer still active",
    "did_not_return": "Did not return",
    "needs_review": "Needs review",
}
BUCKET_ORDER = list(BUCKET_LABELS)

_OUT_COLS = ["student_key", "student", "grade", "center", "summer_start",
             "summer_end", "summer_status", "bucket", "bucket_label",
             "continued_as", "new_start", "new_monthly"]


def _classify_one(g: pd.DataFrame, as_of: date) -> dict:
    center = g.iloc[0]["center"]
    summer_membership = None
    for cfg in c.CENTERS.values():
        if (g["membership"] == cfg["summer_membership"]).any():
            summer_membership = cfg["summer_membership"]
            break
    is_summer = g["membership"] == summer_membership
    summer = g[is_summer].sort_values("start_date")
    other = g[~is_summer]

    summer_start = summer["start_date"].min()
    summer_last = summer.sort_values("end_date").iloc[-1]

    res = {
        "student_key": g.iloc[0]["student_key"],
        "student": g.iloc[0]["student"],
        "grade": g.iloc[0]["grade"],
        "center": center,
        "summer_start": summer_start,
        "summer_end": summer_last["end_date"],
        "summer_status": summer_last["status"],
        "continued_as": "",
        "new_start": pd.NaT,
        "new_monthly": pd.NA,
    }

    qualifying = other[other["start_date"] >= summer_start]
    non_sp = qualifying[~qualifying["is_school_partnership"]]
    sp = qualifying[qualifying["is_school_partnership"]]

    conv_hold = non_sp[non_sp["status"] == "On Hold"]
    conv_active = non_sp[non_sp["status"] == "Enrolled"]

    if len(conv_hold) or len(conv_active):
        bucket = "converted_hold" if len(conv_hold) else "converted_active"
        detail = non_sp.sort_values("start_date").iloc[-1]
    elif len(sp):
        bucket = "school_partnership"
        detail = sp.sort_values("start_date").iloc[-1]
    else:
        detail = None
        if summer_last["status"] not in config_known_statuses():
            bucket = "needs_review"
        elif summer_last["status"] == "Pre-Enrolled":
            bucket = "needs_review"
        elif pd.notna(summer_last["end_date"]) and summer_last["end_date"].date() >= as_of:
            bucket = "summer_active"
        else:
            bucket = "did_not_return"

    if detail is not None:
        res["continued_as"] = detail["membership"]
        res["new_start"] = detail["start_date"]
        res["new_monthly"] = detail["monthly_amount"]

    res["bucket"] = bucket
    res["bucket_label"] = BUCKET_LABELS[bucket]
    return res


def config_known_statuses():
    return c.KNOWN_STATUSES


def classify_students(df: pd.DataFrame, as_of: date | None = None) -> pd.DataFrame:
    """One row per student who has any summer-package enrollment row."""
    as_of = as_of or date.today()
    summer_strings = {cfg["summer_membership"] for cfg in c.CENTERS.values()}
    summer_keys = set(df.loc[df["membership"].isin(summer_strings), "student_key"])
    sub = df[df["student_key"].isin(summer_keys)]
    records = [
        _classify_one(g.reset_index(drop=True), as_of)
        for _, g in sub.groupby("student_key", sort=True)
    ]
    return pd.DataFrame(records, columns=_OUT_COLS).sort_values(
        by=["bucket", "student"],
        key=lambda s: s.map({b: i for i, b in enumerate(BUCKET_ORDER)}) if s.name == "bucket" else s,
    ).reset_index(drop=True)
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

## Task 6: `transform.py` — summary + same-name collision check

**Files:**
- Modify: `~/mathnasium-summer-status/transform.py`
- Test: `~/mathnasium-summer-status/tests/test_transform.py`

- [ ] **Step 1: Write the failing tests**

```python
from transform import build_summary, find_name_collisions


def test_build_summary_counts_and_rate():
    classified = pd.DataFrame([
        {"student": "A", "center": "Teaneck", "bucket": "converted_active"},
        {"student": "B", "center": "Teaneck", "bucket": "converted_hold"},
        {"student": "C", "center": "Teaneck", "bucket": "school_partnership"},
        {"student": "D", "center": "Teaneck", "bucket": "did_not_return"},
    ])
    s = build_summary(classified)
    assert s["total"] == 4
    assert s["counts"]["converted_active"] == 1
    assert s["counts"]["summer_active"] == 0
    assert round(s["conversion_rate"], 2) == 0.75   # 3 of 4 (incl. school partnership)


def test_find_name_collisions_flags_same_name_two_keys():
    raw = pd.DataFrame([
        {"student": "Sam Lee", "student_key": "sam lee|teaneck", "center": "Teaneck"},
        {"student": "Sam Lee", "student_key": "sam lee|teaneck", "center": "Teaneck"},
    ])
    assert find_name_collisions(raw) == []   # same key = same kid, fine
```

Note: `find_name_collisions` detects the *rare* case the spec calls out — it operates
on the raw normalized frame and flags a `(student, center)` name that maps to enrollment
rows with conflicting `Account Id`. Since `normalize` currently drops `Account Id`, add
it: extend `_CANON` and `normalize` to carry an `account_id` column (`raw["Account Id"]
.map(_clean)`), update the Task 4 `test_normalize_shapes_and_types` column list to
include `"account_id"` at the end, then implement:

```python
def find_name_collisions(norm: pd.DataFrame) -> list[str]:
    """Return display names where one (name, center) maps to >1 Account Id."""
    g = norm.groupby(["student", "center"])["account_id"].nunique()
    return sorted(name for (name, _center), n in g.items() if n > 1)
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_transform.py -v -k build_summary or find_name_collisions or normalize_shapes`
Expected: FAIL — `cannot import name 'build_summary'`

- [ ] **Step 3: Implement**

```python
def build_summary(classified: pd.DataFrame) -> dict:
    total = len(classified)
    counts = {k: int((classified["bucket"] == k).sum()) for k in BUCKET_LABELS}
    converted = (counts["converted_active"] + counts["converted_hold"]
                 + counts["school_partnership"])
    return {
        "total": total,
        "counts": counts,
        "converted": converted,
        "conversion_rate": (converted / total) if total else 0.0,
        "needs_review": counts["needs_review"],
    }
```

Plus the `account_id` additions to `_CANON` / `normalize` and `find_name_collisions`
described above.

- [ ] **Step 4: Run to verify it passes**

Run: `python -m pytest tests/test_transform.py -v`
Expected: PASS (all tests)

- [ ] **Step 5: Commit**

```bash
git add transform.py tests/test_transform.py
git commit -m "feat: per-center summary + same-name collision check"
```

---

## Task 7: `report.py` — render the HTML

**Files:**
- Create: `~/mathnasium-summer-status/report.py`
- Test: `~/mathnasium-summer-status/tests/test_report.py`

- [ ] **Step 1: Write the failing test**

```python
from datetime import date
import pandas as pd
from report import render_html


def _classified():
    return pd.DataFrame([
        {"student_key": "amy a|teaneck", "student": "Amy A", "grade": "5",
         "center": "Teaneck", "summer_start": pd.Timestamp("2026-06-01"),
         "summer_end": pd.Timestamp("2026-08-31"), "summer_status": "Inactive",
         "bucket": "converted_active", "bucket_label": "Converted – active",
         "continued_as": "1x/week - 2024 Flex (Flexible)",
         "new_start": pd.Timestamp("2026-09-01"), "new_monthly": 329},
        {"student_key": "ben b|englewood", "student": "Ben B", "grade": "7",
         "center": "Englewood", "summer_start": pd.Timestamp("2026-07-01"),
         "summer_end": pd.Timestamp("2026-08-15"), "summer_status": "Inactive",
         "bucket": "did_not_return", "bucket_label": "Did not return",
         "continued_as": "", "new_start": pd.NaT, "new_monthly": pd.NA},
    ])


def test_render_html_has_both_sections_and_data():
    html = render_html(_classified(), as_of=date(2026, 9, 7))
    assert "<h2>Teaneck</h2>" in html
    assert "<h2>Englewood</h2>" in html
    assert "Amy A" in html and "Ben B" in html
    assert "1x/week - 2024 Flex (Flexible)" in html
    assert "2026-09-07" in html            # as-of date shown
    assert "<link" not in html             # self-contained
    assert "&lt;script&gt;" not in html or True  # escaping sanity (see next test)


def test_render_html_escapes_student_names():
    df = _classified()
    df.loc[0, "student"] = "<b>x</b>"
    html = render_html(df, as_of=date(2026, 9, 7))
    assert "<b>x</b>" not in html
    assert "&lt;b&gt;x&lt;/b&gt;" in html


def test_render_html_flags_needs_review():
    df = _classified()
    df.loc[1, "bucket"] = "needs_review"
    df.loc[1, "bucket_label"] = "Needs review"
    html = render_html(df, as_of=date(2026, 9, 7))
    assert "needs review" in html.lower()
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest tests/test_report.py -v`
Expected: FAIL — `No module named 'report'`

- [ ] **Step 3: Implement `report.py`**

```python
"""Render the self-contained HTML status report."""
import html as _html
import webbrowser
from datetime import date, datetime
import pandas as pd
import config as c
from transform import BUCKET_LABELS, BUCKET_ORDER, build_summary, find_name_collisions

_CSS = """
body{font-family:-apple-system,Segoe UI,Arial,sans-serif;margin:2rem;color:#222;background:#fff}
h1{font-size:1.4rem}h2{font-size:1.15rem;margin-top:2rem;border-bottom:2px solid #ddd;padding-bottom:2px}
table{border-collapse:collapse;margin:.6rem 0;font-size:.9rem}
th,td{border:1px solid #ccc;padding:4px 8px;text-align:left;vertical-align:top}
th{background:#f2f2f2}
.rate{font-weight:bold}
.flag{background:#fff4e5;border:1px solid #f0c36d;padding:8px 12px;margin:1rem 0;border-radius:4px}
.muted{color:#888}
"""


def _esc(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NA or v is pd.NaT:
        return ""
    return _html.escape(str(v))


def _d(v) -> str:
    return "" if pd.isna(v) else pd.Timestamp(v).strftime("%Y-%m-%d")


def _money(v) -> str:
    return "" if pd.isna(v) else f"${float(v):,.0f}"


def _summary_table(per_center: dict, overall: dict) -> str:
    names = list(per_center)
    head = "".join(f"<th>{n}</th>" for n in names) + "<th>Overall</th>"
    rows = ""
    for k in BUCKET_ORDER:
        cells = "".join(f"<td>{per_center[n]['counts'][k]}</td>" for n in names)
        rows += f"<tr><td>{BUCKET_LABELS[k]}</td>{cells}<td>{overall['counts'][k]}</td></tr>"
    tot = "".join(f"<td>{per_center[n]['total']}</td>" for n in names)
    rate = "".join(f"<td class='rate'>{per_center[n]['conversion_rate']:.0%}</td>" for n in names)
    rows += f"<tr><td><b>Total summer students</b></td>{tot}<td><b>{overall['total']}</b></td></tr>"
    rows += (f"<tr><td class='rate'>Conversion rate</td>{rate}"
             f"<td class='rate'>{overall['conversion_rate']:.0%}</td></tr>")
    return f"<table><tr><th></th>{head}</tr>{rows}</table>"


def _detail_table(sub: pd.DataFrame) -> str:
    order = {b: i for i, b in enumerate(BUCKET_ORDER)}
    sub = sub.sort_values(by=["bucket", "student"],
                          key=lambda s: s.map(order) if s.name == "bucket" else s)
    head = ("<tr><th>Student</th><th>Grade</th><th>Summer start</th><th>Summer end</th>"
            "<th>Summer status</th><th>Current status</th><th>Continued as</th>"
            "<th>New start</th><th>Monthly $</th></tr>")
    body = ""
    for _, r in sub.iterrows():
        body += (
            f"<tr><td>{_esc(r['student'])}</td><td>{_esc(r['grade'])}</td>"
            f"<td>{_d(r['summer_start'])}</td><td>{_d(r['summer_end'])}</td>"
            f"<td>{_esc(r['summer_status'])}</td><td>{_esc(r['bucket_label'])}</td>"
            f"<td>{_esc(r['continued_as'])}</td><td>{_d(r['new_start'])}</td>"
            f"<td>{_money(r['new_monthly'])}</td></tr>"
        )
    return f"<table>{head}{body}</table>"


def render_html(classified: pd.DataFrame, as_of: date, collisions=None) -> str:
    names = list(c.CENTERS)
    per_center = {n: build_summary(classified[classified["center"] == n]) for n in names}
    overall = build_summary(classified)
    p = [
        "<!doctype html><html><head><meta charset='utf-8'>",
        f"<title>2026 Summer Package — Enrollment Status</title><style>{_CSS}</style>",
        "</head><body>",
        "<h1>2026 Summer Package &mdash; Enrollment Status</h1>",
        f"<p class='muted'>Generated {datetime.now():%Y-%m-%d %H:%M} &middot; "
        f"“still active” measured as of {as_of:%Y-%m-%d}</p>",
    ]
    if overall["needs_review"]:
        p.append(f"<div class='flag'>{overall['needs_review']} student(s) could not be "
                 f"auto-classified — see the “Needs review” rows.</div>")
    if collisions:
        joined = ", ".join(_esc(x) for x in collisions)
        p.append(f"<div class='flag'>Same-name students at one center (verify manually): {joined}</div>")
    p.append("<h2>Overall summary</h2>")
    p.append(_summary_table(per_center, overall))
    for n in names:
        s = per_center[n]
        p.append(f"<h2>{n}</h2>")
        p.append(f"<p class='rate'>Conversion rate: {s['conversion_rate']:.0%} "
                 f"({s['converted']} of {s['total']})</p>")
        p.append(_detail_table(classified[classified["center"] == n]))
    p.append("</body></html>")
    return "".join(p)


def write_and_open(classified: pd.DataFrame, norm_frames=None) -> str:
    c.OUTPUT_DIR.mkdir(exist_ok=True)
    collisions = []
    if norm_frames is not None:
        for f in norm_frames:
            collisions += find_name_collisions(f)
    html = render_html(classified, as_of=date.today(), collisions=collisions or None)
    out = c.OUTPUT_DIR / "summer_2026_status.html"
    out.write_text(html, encoding="utf-8")
    webbrowser.open(out.as_uri())
    return str(out)
```

- [ ] **Step 4: Run to verify it passes**

Run: `python -m pytest tests/test_report.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add report.py tests/test_report.py
git commit -m "feat: render self-contained HTML status report"
```

---

## Task 8: `main.py` — orchestrator + real end-to-end run

**Files:**
- Create: `~/mathnasium-summer-status/main.py`

- [ ] **Step 1: Write `main.py`**

```python
"""Summer package status report — CLI entry point.

Usage:
    export RADIUS_USERNAME="$(grep '^RADIUS_USERNAME=' ~/.mathnasium_env | cut -d= -f2-)"
    export RADIUS_PASSWORD="$(grep '^RADIUS_PASSWORD=' ~/.mathnasium_env | cut -d= -f2-)"
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

    norm_frames = []
    for name in c.CENTERS:
        path = c.INPUT_DIR / f"EnrollmentReport_{name}.xlsx"
        if not path.exists():
            sys.exit(f"Missing {path}. Run once without --no-download first.")
        raw = pd.read_excel(path)
        if raw.empty:
            sys.exit(f"{name}: Enrollment Report export has zero rows — aborting.")
        frame = load_center(path)
        summer_str = c.CENTERS[name]["summer_membership"]
        if not (frame["membership"] == summer_str).any():
            sys.exit(f"{name}: no rows match summer membership '{summer_str}'. "
                     f"The Radius label may have changed — update config.CENTERS.")
        norm_frames.append(frame)

    all_rows = pd.concat(norm_frames, ignore_index=True)
    classified = classify_students(all_rows)

    out = write_and_open(classified, norm_frames=norm_frames)
    n_review = int((classified["bucket"] == "needs_review").sum())
    print(f"Wrote {out} — {len(classified)} summer students"
          + (f"; {n_review} need manual review." if n_review else "."))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-download", action="store_true",
                    help="re-render from cached input/ without hitting Radius")
    args = ap.parse_args()
    run(no_download=args.no_download)
```

- [ ] **Step 2: Run `--no-download` first (uses the cached discovery exports)**

```bash
cd ~/mathnasium-summer-status
python main.py --no-download
```

Expected: browser opens `output/summer_2026_status.html`; both center sections populated; Teaneck ~24 summer students, Englewood ~17; console prints the count line.

- [ ] **Step 3: Full run against Radius**

```bash
export RADIUS_USERNAME="$(grep '^RADIUS_USERNAME=' ~/.mathnasium_env | cut -d= -f2-)"
export RADIUS_PASSWORD="$(grep '^RADIUS_PASSWORD=' ~/.mathnasium_env | cut -d= -f2-)"
python main.py
```

Expected: fresh exports download, report regenerates, same shape.

- [ ] **Step 4: Spot-check.** In Radius, filter the Enrollment Report by the Teaneck summer membership type and confirm the student count matches the report's Teaneck total. Pick one converted student and one "did not return" student and confirm their rows look right.

- [ ] **Step 5: Commit**

```bash
git add main.py
git commit -m "feat: CLI orchestrator + end-to-end run"
```

---

## Task 9: README + GitHub repo

**Files:**
- Create: `~/mathnasium-summer-status/README.md`

- [ ] **Step 1: Write `README.md`**

```markdown
# Summer Package Status Report

On-demand HTML report of every 2026 summer-package enrollment at Teaneck and
Englewood, showing each student's current status and whether they converted to an
ongoing membership.

## Run

```bash
cd ~/mathnasium-summer-status
pip install -r requirements.txt        # first time
playwright install chromium            # first time
export RADIUS_USERNAME="$(grep '^RADIUS_USERNAME=' ~/.mathnasium_env | cut -d= -f2-)"
export RADIUS_PASSWORD="$(grep '^RADIUS_PASSWORD=' ~/.mathnasium_env | cut -d= -f2-)"
python main.py                         # pull fresh + open report
python main.py --no-download           # re-render from cached input/
```

Output: `output/summer_2026_status.html` (opens automatically).

## Status buckets

- **Converted – active / on hold** — has an ongoing non-summer enrollment that started
  on/after their summer enrollment.
- **Continued via School Partnership** — same, but the new enrollment is a School
  Partnership program.
- **Summer still active** — summer package end date is still in the future.
- **Did not return** — summer package ended, nothing after it.
- **Needs review** — couldn't be classified automatically (e.g. a Pre-Enrolled row).

## Next year

Edit `CENTERS[...]["summer_membership"]` and `PULL_START` / `PULL_END` in `config.py`.
If Radius changed the Enrollment Report, re-run `python discover_selectors.py` and
update the `COL_*` / `SEL_*` constants.
```

- [ ] **Step 2: Full test suite**

Run: `cd ~/mathnasium-summer-status && python -m pytest -v`
Expected: PASS (all tests in `tests/`)

- [ ] **Step 3: Commit + create repo**

```bash
git add README.md
git commit -m "docs: usage README"
gh repo create mdiamond77/mathnasium-summer-status --private --source=. --push
```

---

## Self-Review Notes

**Spec coverage:**
- One pull per center, 4/1–12/31/2026, all statuses → Task 3
- Discovery findings baked into `config.py` → Task 2
- Six buckets, first-match order, School Partnership as its own bucket → Task 5 (tests each)
- "Still active" by end-date not status → Task 5 (`test_summer_still_active_by_end_date_not_status`)
- Student key = first+last+center; sibling-safe; same-name collision flagged → Tasks 4 & 6
- Conversion detail = membership name + start + Monthly Amount; latest-start wins → Task 5
- Converted-then-hold → on-hold wins → Task 5 (`test_converted_hold_beats_active`)
- Pre-summer enrollment not a conversion → Task 5 (`test_pre_summer_enrollment_is_not_a_conversion`)
- HTML: header w/ as-of date, overall summary Teaneck|Englewood|Overall, per-center detail, needs-review callout, HTML-escaped → Task 7
- Fail-loud: login, zero rows, zero summer matches, needs-review count → Tasks 3 & 8
- `--no-download` → Task 8
- New repo under `mdiamond77`, no workflow → Tasks 1 & 9

**Type consistency:** `classify_students` emits `_OUT_COLS`; `report.py` reads exactly those names (`bucket`, `bucket_label`, `continued_as`, `new_start`, `new_monthly`, `summer_start`, `summer_end`, `summer_status`, `grade`, `student`, `center`). `build_summary` keys (`total`, `counts`, `converted`, `conversion_rate`, `needs_review`) match between Task 6 and Task 7. `normalize` `_CANON` gains `account_id` in Task 6 — the Task 4 column-list test is updated in the same task.

**Deferred to implementation (not plan gaps):** exact jQuery-readiness wait in `download.py` (Task 3 Step 3 covers it); whether `Grade` comes through as `"5"` vs `5` (— `_clean` stringifies either way).
