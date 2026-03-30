# Appointy Nightly Export — Design Spec
**Date:** 2026-03-30

## Overview

A Python script runs on a nightly GitHub Actions cron schedule. It authenticates to Appointy (`mathnasium-admin.appointy.com`) via direct HTTP requests (no browser), navigates to the Appointment Detailed Report, and triggers the email export for the current week's date range. The existing Gmail → Google Sheet pipeline receives the email and processes it downstream — no changes needed there.

## Problem Context

Appointy's web UI is protected by Cloudflare Turnstile (the "I am human" checkbox). Previous attempts using headless Playwright were blocked at this challenge. By making direct HTTP requests instead of driving a browser, we bypass Cloudflare entirely — the challenge only appears in browser contexts, not raw HTTP calls.

## Architecture

```
GitHub Actions (nightly cron)
        │
        ▼
   Python script
   ├── POST /login  (username + password → session cookie)
   ├── GET  /export-page  (capture CSRF token if required)
   └── POST /export  (date range: week-start Sunday → Sunday + 70 days)
        │
        ▼
  Appointy sends report email
        │
        ▼
  Existing Gmail → Google Sheet pipeline (unchanged)
```

## Components

### 1. GitHub Actions Workflow (`.github/workflows/appointy_export.yml`)
- Trigger: nightly cron schedule (time TBD based on when downstream sheet needs the data)
- Checks out the repo, sets up Python, runs the export script
- Injects `APPOINTY_USERNAME` and `APPOINTY_PASSWORD` from repository secrets

### 2. Export Script (`scripts/appointy_export.py`)
- **Login:** POST credentials to Appointy's login endpoint; capture session cookie
- **CSRF:** If required, fetch the export page and extract a CSRF/nonce token
- **Date range:** Compute start = most recent Sunday ≤ today; end = start + 70 days
- **Export trigger:** POST to Appointy's export endpoint with date range; email destination is whatever Appointy already has configured (confirmed during DevTools capture)
- **Error handling:** Exit with a non-zero code on failure so GitHub Actions marks the run as failed and sends a notification email

### 3. GitHub Secrets
- `APPOINTY_USERNAME` — Appointy login email
- `APPOINTY_PASSWORD` — Appointy login password

## Date Range Logic

Computed at runtime each night:
- **Start:** The most recent Sunday on or before today (stable within a week — Mon–Sat all produce the same Sunday)
- **End:** Start + 70 days (exactly 10 weeks)

Example: if the script runs on Wednesday 2026-04-01, start = 2026-03-29 (Sunday), end = 2026-06-07.

## Credentials & Security

- Credentials stored exclusively as GitHub Actions secrets — never in source code or committed files
- Script reads credentials from environment variables at runtime
- Repository can be public or private; secrets are safe either way

## Fallback Plan

If Appointy's login endpoint itself triggers a Cloudflare challenge (possible but unlikely for direct HTTP), the fallback is:
- Run the same script on a Mac using `launchd` on a nightly schedule
- Replace the HTTP-based login with headed Playwright (real Chrome, non-headless), which typically bypasses Turnstile without needing to solve it

This fallback will only be evaluated after the DevTools investigation confirms whether direct HTTP login is viable.

## DevTools Investigation (Pre-Implementation Step)

Before writing the script, we need to capture the exact HTTP requests Appointy makes when:
1. Logging in
2. Triggering the email export

This is done by:
1. Opening Appointy in Chrome with DevTools → Network tab open
2. Logging in normally and identifying the login POST request
3. Navigating to the Appointment Detailed Report, setting the date range, and clicking export
4. Capturing the export POST request (URL, headers, body/payload)

The script will replicate these requests exactly.

## Success Criteria

- GitHub Actions workflow runs nightly without manual intervention
- Appointy sends the report email to the configured address each night
- Existing Gmail → Google Sheet pipeline continues to work with no changes
- Failed runs surface as GitHub Actions failures (email notification to repo owner)
- No credentials stored in the repository

## Out of Scope

- Parsing or processing the CSV contents (handled by existing downstream pipeline)
- Changes to the Gmail → Google Sheet integration
- Any UI or dashboard for monitoring
