# Appointy Nightly Export Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A Python script runs nightly on GitHub Actions, logs into Appointy via HTTP, and triggers the Appointment Detailed Report email export for Sunday of the current week through 10 weeks out.

**Architecture:** Direct HTTP requests (no browser) bypass Cloudflare Turnstile entirely. A GitHub Actions cron workflow runs the script nightly, injecting credentials from repository secrets. The existing Gmail → Google Sheet pipeline receives the exported email unchanged.

**Tech Stack:** Python 3.11, `requests`, `pytest`, GitHub Actions

---

## File Structure

```
mathnasium-appointy-export/
├── .github/
│   └── workflows/
│       └── nightly_export.yml        # GitHub Actions cron workflow
├── scripts/
│   └── appointy_export.py            # Main export script
├── tests/
│   └── test_appointy_export.py       # Unit tests (date logic + request building)
├── requirements.txt                  # Python dependencies
└── docs/
    └── superpowers/
        ├── specs/
        │   └── 2026-03-30-appointy-nightly-export-design.md
        └── plans/
            └── 2026-03-30-appointy-nightly-export.md
```

---

## Task 1: Connect Local Repo to GitHub

**Files:** none (git setup only)

### Step 1a: Create a GitHub Personal Access Token

A Personal Access Token (PAT) lets your Mac push code to GitHub without a password.

1. Go to **github.com** → click your profile photo (top right) → **Settings**
2. Scroll to the bottom of the left sidebar → click **Developer settings**
3. Click **Personal access tokens** → **Tokens (classic)**
4. Click **Generate new token** → **Generate new token (classic)**
5. Give it a name like `mac-push`
6. Under **Expiration**, choose **No expiration** (or 1 year)
7. Check the box next to **repo** (this gives full repo access)
8. Scroll down and click **Generate token**
9. **Copy the token immediately** — GitHub only shows it once. It starts with `ghp_`

### Step 1b: Push the repo

Open Terminal and run these commands one at a time. When prompted for a username, type your GitHub username (`mdmathnasiums`). When prompted for a password, paste the token you just copied (it won't show as you type — that's normal).

```bash
cd /Users/mattdiamond/Mathnasium_automation
git remote add origin https://github.com/mdmathnasiums/mathnasium-appointy-export.git
git push -u origin main
```

Expected output ends with: `Branch 'main' set up to track remote branch 'main' from 'origin'.`

- [ ] Confirm the push succeeded by visiting https://github.com/mdmathnasiums/mathnasium-appointy-export — you should see the `docs/` folder there.

---

## Task 2: DevTools Investigation (capture Appointy's HTTP requests)

**Goal:** Identify the exact login POST and export POST requests so the script can replicate them.

**Files:** none (research only — findings feed into Task 3)

- [ ] **Step 2.1: Open Chrome DevTools**

  1. Go to `https://mathnasium-admin.appointy.com` in Chrome
  2. Press **F12** (or right-click anywhere → **Inspect**) to open DevTools
  3. Click the **Network** tab at the top of the DevTools panel
  4. Check the **Preserve log** checkbox (so requests don't disappear after login)
  5. Make sure the filter dropdown shows **All** (not just XHR or Fetch)

- [ ] **Step 2.2: Capture the login request**

  1. With the Network tab open and recording, fill in your username and password on the Appointy login page and click **Sign In**
  2. After logging in, look at the Network tab — you'll see a list of requests
  3. Find the one that looks like a login — it will be a **POST** request, often named `login`, `signin`, `auth`, or similar
  4. Click on it
  5. In the panel that opens on the right, click the **Headers** sub-tab
  6. Note down (screenshot or copy):
     - The **Request URL** (e.g., `https://mathnasium-admin.appointy.com/api/v1/login`)
     - Under **Request Headers**: copy all headers listed
  7. Click the **Payload** sub-tab (sometimes called **Form Data** or **Request Body**)
  8. Note down all the fields and values shown (e.g., `email=...`, `password=...`, any hidden fields like `_token`)

- [ ] **Step 2.3: Capture the export request**

  1. Keep DevTools open and the Network tab recording
  2. Navigate to the **Appointment Detailed Report** page in Appointy
  3. Set the date range fields to any values (exact dates don't matter for capture)
  4. Click the **Email Export** (or similar) button
  5. In the Network tab, find the new POST request that appeared
  6. Click on it and note down:
     - The **Request URL**
     - All **Request Headers**
     - All **Payload** fields (especially which field holds the date range — likely `start_date`, `end_date`, or similar)

- [ ] **Step 2.4: Share findings**

  Paste the captured URLs, headers, and payload fields into the chat. This determines the exact code for Task 3.

---

## Task 3: Set Up Project Structure

**Files:**
- Create: `requirements.txt`
- Create: `scripts/appointy_export.py` (skeleton)
- Create: `tests/test_appointy_export.py` (skeleton)

- [ ] **Step 3.1: Create requirements.txt**

```
requests==2.31.0
pytest==8.1.1
```

- [ ] **Step 3.2: Create script skeleton**

Create `scripts/appointy_export.py`:

```python
import os
import sys
from datetime import date, timedelta
import requests


def get_week_start(today: date) -> date:
    """Return the most recent Sunday on or before today."""
    days_since_sunday = (today.weekday() + 1) % 7
    return today - timedelta(days=days_since_sunday)


def get_date_range(today: date) -> tuple[date, date]:
    """Return (start, end) where start is Sunday of current week, end is start + 70 days."""
    start = get_week_start(today)
    end = start + timedelta(days=70)
    return start, end


def login(session: requests.Session, username: str, password: str) -> None:
    """Log into Appointy. Raises on failure. Populated in Task 4."""
    raise NotImplementedError


def trigger_export(session: requests.Session, start: date, end: date) -> None:
    """Trigger the email export. Raises on failure. Populated in Task 4."""
    raise NotImplementedError


def main() -> None:
    username = os.environ["APPOINTY_USERNAME"]
    password = os.environ["APPOINTY_PASSWORD"]

    today = date.today()
    start, end = get_date_range(today)

    session = requests.Session()
    login(session, username, password)
    trigger_export(session, start, end)
    print(f"Export triggered: {start} to {end}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 3.3: Create test skeleton**

Create `tests/test_appointy_export.py`:

```python
from datetime import date
from scripts.appointy_export import get_week_start, get_date_range
```

- [ ] **Step 3.4: Install dependencies**

```bash
cd /Users/mattdiamond/Mathnasium_automation
pip install -r requirements.txt
```

Expected: packages install without errors.

- [ ] **Step 3.5: Commit**

```bash
git add requirements.txt scripts/appointy_export.py tests/test_appointy_export.py
git commit -m "feat: add project skeleton and date logic"
git push
```

---

## Task 4: Date Calculation Logic (TDD)

**Files:**
- Modify: `tests/test_appointy_export.py`
- Modify: `scripts/appointy_export.py` (already implemented — tests verify it)

- [ ] **Step 4.1: Write failing tests**

Replace the contents of `tests/test_appointy_export.py`:

```python
from datetime import date
from scripts.appointy_export import get_week_start, get_date_range


def test_get_week_start_on_sunday():
    assert get_week_start(date(2026, 3, 29)) == date(2026, 3, 29)  # Sunday → itself


def test_get_week_start_on_wednesday():
    assert get_week_start(date(2026, 4, 1)) == date(2026, 3, 29)  # Wed → prior Sunday


def test_get_week_start_on_saturday():
    assert get_week_start(date(2026, 4, 4)) == date(2026, 3, 29)  # Sat → prior Sunday


def test_get_date_range_start_is_sunday():
    start, end = get_date_range(date(2026, 4, 1))  # Wednesday
    assert start == date(2026, 3, 29)


def test_get_date_range_end_is_70_days_out():
    start, end = get_date_range(date(2026, 4, 1))
    assert end == date(2026, 6, 7)
```

- [ ] **Step 4.2: Run tests**

```bash
cd /Users/mattdiamond/Mathnasium_automation
pytest tests/test_appointy_export.py -v
```

Expected: all 5 tests **PASS** (the logic is already written in Task 3).

- [ ] **Step 4.3: Commit**

```bash
git add tests/test_appointy_export.py
git commit -m "test: verify date range calculation"
git push
```

---

## Task 5: Login and Export HTTP Functions

> **Prerequisite:** Task 2 (DevTools investigation) must be complete. Replace the placeholder values below with the actual URLs, headers, and field names captured in Task 2.

**Files:**
- Modify: `scripts/appointy_export.py`

- [ ] **Step 5.1: Implement login()**

Replace the `login` function with the actual captured request. Example structure (exact values from Task 2):

```python
def login(session: requests.Session, username: str, password: str) -> None:
    """Log into Appointy and store session cookie."""
    url = "REPLACE_WITH_LOGIN_URL_FROM_TASK_2"
    payload = {
        "REPLACE_WITH_EMAIL_FIELD": username,
        "REPLACE_WITH_PASSWORD_FIELD": password,
        # Add any other fields captured from DevTools (e.g. _token, remember_me)
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Content-Type": "application/x-www-form-urlencoded",
        # Add any other required headers captured from DevTools
    }
    response = session.post(url, data=payload, headers=headers)
    response.raise_for_status()
    # Verify login succeeded — check for redirect to dashboard or absence of login form
    if "login" in response.url.lower() or "sign" in response.url.lower():
        raise RuntimeError(f"Login failed — still on login page: {response.url}")
```

- [ ] **Step 5.2: Implement trigger_export()**

Replace the `trigger_export` function:

```python
def trigger_export(session: requests.Session, start: date, end: date) -> None:
    """Trigger the Appointment Detailed Report email export."""
    url = "REPLACE_WITH_EXPORT_URL_FROM_TASK_2"
    payload = {
        "REPLACE_WITH_START_DATE_FIELD": start.strftime("%Y-%m-%d"),
        "REPLACE_WITH_END_DATE_FIELD": end.strftime("%Y-%m-%d"),
        # Add any other required fields captured from DevTools
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        # Add any other required headers captured from DevTools
    }
    response = session.post(url, data=payload, headers=headers)
    response.raise_for_status()
```

- [ ] **Step 5.3: Test login manually**

```bash
cd /Users/mattdiamond/Mathnasium_automation
APPOINTY_USERNAME="your@email.com" APPOINTY_PASSWORD="yourpassword" python scripts/appointy_export.py
```

Expected output: `Export triggered: 2026-03-29 to 2026-06-07` (with current week's dates)

Check your email — the report email from Appointy should arrive within a few minutes.

- [ ] **Step 5.4: Commit**

```bash
git add scripts/appointy_export.py
git commit -m "feat: implement login and export trigger"
git push
```

---

## Task 6: GitHub Actions Workflow

**Files:**
- Create: `.github/workflows/nightly_export.yml`

- [ ] **Step 6.1: Add GitHub Secrets**

These store your Appointy credentials securely in GitHub (never in code):

1. Go to https://github.com/mdmathnasiums/mathnasium-appointy-export
2. Click **Settings** (top tab bar)
3. In the left sidebar: **Secrets and variables** → **Actions**
4. Click **New repository secret**
5. Name: `APPOINTY_USERNAME` / Value: your Appointy login email → **Add secret**
6. Click **New repository secret** again
7. Name: `APPOINTY_PASSWORD` / Value: your Appointy password → **Add secret**

- [ ] **Step 6.2: Create the workflow file**

Create `.github/workflows/nightly_export.yml`:

```yaml
name: Nightly Appointy Export

on:
  schedule:
    - cron: '0 7 * * *'   # 7:00 AM UTC = midnight PT (adjust for daylight saving)
  workflow_dispatch:        # allows manual trigger from GitHub UI

jobs:
  export:
    runs-on: ubuntu-latest
    steps:
      - name: Check out repo
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Run export
        env:
          APPOINTY_USERNAME: ${{ secrets.APPOINTY_USERNAME }}
          APPOINTY_PASSWORD: ${{ secrets.APPOINTY_PASSWORD }}
        run: python scripts/appointy_export.py
```

- [ ] **Step 6.3: Commit and push**

```bash
git add .github/workflows/nightly_export.yml
git commit -m "feat: add nightly GitHub Actions workflow"
git push
```

- [ ] **Step 6.4: Trigger a manual test run**

1. Go to https://github.com/mdmathnasiums/mathnasium-appointy-export
2. Click the **Actions** tab
3. Click **Nightly Appointy Export** in the left sidebar
4. Click **Run workflow** → **Run workflow**
5. Wait ~1 minute, then click the run to see its output
6. Expected: green checkmark and `Export triggered: ...` in the logs
7. Check email — report should arrive within a few minutes

---

## Task 7: Verify End-to-End

- [ ] Confirm the GitHub Actions run shows green
- [ ] Confirm the Appointy report email arrives in the expected inbox
- [ ] Confirm the existing Google Sheet picks up the new data as expected
- [ ] If any step fails, check the Actions run logs — errors will be printed there

---

## Fallback Plan (if HTTP login is blocked by Cloudflare)

If Task 5.3 fails because Appointy returns a Cloudflare challenge page instead of logging in, the script will need to use a real browser. Skip to this section only if that happens.

The fallback uses Playwright with a visible (non-headless) browser, which typically bypasses Turnstile without needing to solve it. This runs on your Mac via `launchd` instead of GitHub Actions. Details to be designed if needed.
