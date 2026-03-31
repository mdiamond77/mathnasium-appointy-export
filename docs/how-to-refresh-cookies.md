# How to Refresh Appointy Cookies

The nightly export uses browser cookies to authenticate with Appointy. These expire
periodically (roughly every 30 days). When they do, the GitHub Actions run will fail
and you'll receive a notification email.

**This process takes about 5 minutes.**

---

## Step 1: Get fresh cookies from Appointy

1. Open **Chrome** and press **Cmd+Shift+N** to open an Incognito window
2. Go to `https://mathnasium-admin.appointy.com`
3. Press **F12** to open DevTools
4. Click the **Network** tab — check the **Preserve log** checkbox
5. Log in with your Appointy username and password
6. Once on the dashboard, navigate to **Reports → Appointment Detailed Report**
7. Set any date range and click the **Email export** button → select **Send to email** → click **Submit**
8. In the Network tab, find the `graphql?companyId=...` POST request that just appeared
9. Right-click it → **Copy → Copy as cURL**

---

## Step 2: Extract the cookie string

In Terminal, run:

```bash
pbpaste | python3 -c "
import sys, re
curl = sys.stdin.read()
match = re.search(r\"-b '([^']+)'\", curl)
if match:
    print(match.group(1))
else:
    print('ERROR: cookie not found — make sure you copied the cURL command first')
" | pbcopy
```

This reads the cURL command from your clipboard, extracts just the cookie string,
and copies it back to your clipboard — ready to paste.

---

## Step 3: Update the GitHub secret

1. Go to https://github.com/mdmathnasiums/mathnasium-appointy-export/settings/secrets/actions
2. Click the **pencil icon** next to `APPOINTY_COOKIES`
3. Clear the existing value and paste the new cookie string (Cmd+V)
4. Click **Save changes**

---

## Step 4: Verify it works

1. Go to https://github.com/mdmathnasiums/mathnasium-appointy-export/actions/workflows/nightly_export.yml
2. Click **Run workflow** → **Run workflow**
3. Wait ~30 seconds — you should see a green checkmark ✅
4. Check your email for the Appointy report

---

## Notes

- You do **not** need to change any code — just update the secret
- The cookies from an Incognito window are fresh and reliable
- If Step 2's Terminal command doesn't work, contact whoever set this up for help
