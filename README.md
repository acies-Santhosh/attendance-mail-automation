# Attendance Compliance Mailer

A Google Apps Script web app that:
- Accepts 4 weekly attendance file uploads from HR
- Cross-checks biometric, Ullen Ayyah, and Zoho leave data
- Identifies employees who came to office < 3 days
- Shows a preview table with per-day breakdown
- Sends compliance reminder emails from Gmail in one click
- Logs every sent email to a Google Sheet

---

## Files in this repo

| File | Purpose |
|------|---------|
| `Code.gs` | Server-side Apps Script — parses files, computes compliance, sends emails, writes log |
| `Index.html` | Client-side UI — upload page, preview table, send confirmation |
| `SETUP_INSTRUCTIONS.md` | Step-by-step deployment guide |

---

## One-time setup (do this once, ~10 minutes)

### Step 1 — Create a Google Sheet for logging

1. Go to [sheets.google.com](https://sheets.google.com) and create a blank spreadsheet
2. Name it: **Attendance Compliance Log**
3. Copy the ID from the URL:
   `https://docs.google.com/spreadsheets/d/` **← this part →** `/edit`

### Step 2 — Create the Apps Script project

1. In your Google Sheet → **Extensions → Apps Script**
2. You'll see a default `Code.gs` file — delete everything in it
3. Paste the entire contents of `Code.gs` from this repo
4. The Spreadsheet ID is already filled in at the top of the file:
   ```js
   SPREADSHEET_ID: "1pE1DEIpytWLRzsfUxSn9MQ-PFUE8Rrg5oodKEEj4noo",
   ```
   If you created a **new** sheet, replace this value with your sheet's ID.

5. Click **+ (New file) → HTML** → name it exactly `Index` (no `.html` suffix)
6. Delete everything in it → paste the entire contents of `Index.html` from this repo
7. Press **Ctrl+S** to save

### Step 3 — Enable the Drive API

1. In the Apps Script sidebar, click **Services (+)**
2. Find **Drive API** → click **Add**

### Step 4 — Run setup once

1. In the top toolbar, select the function: `setupSheets`
2. Click **▶ Run**
3. A permissions dialog will appear — click **Review permissions → Allow**
   (It needs Gmail to send emails, Drive to read uploaded files, Sheets to log results)
4. You should see `✅ Sheets created successfully.` in the execution log

### Step 5 — Deploy as a Web App

1. Click **Deploy → New deployment**
2. Click the gear ⚙ next to "Select type" → choose **Web app**
3. Set:
   - **Execute as:** Me *(your Gmail account sends the emails)*
   - **Who has access:** Anyone *(so all HR can open the link without logging in)*
4. Click **Deploy** → copy the `/exec` URL

That URL is what HR bookmarks. Share it with everyone who needs to run the weekly check.

---

## Every week — how HR uses it (3 minutes)

1. Open the web app URL in Chrome
2. Upload the 4 files (drag-and-drop or click each box):

   | Box | File | Format |
   |-----|------|--------|
   | Chennai Attendance | Biometric export from IndiQube Chennai | `.xlsx` |
   | Bangalore Attendance | Biometric/manual entry export | `.csv` or `.xlsx` |
   | Zoho Leave Report | Export from Zoho People (Leave module) | `.xls` or `.xlsx` |
   | Ullen Ayyah Report | Google Sheets export from the Ullen Ayyah bot | `.csv` or `.xlsx` |

3. Confirm the week start (Monday) and end (Friday) dates
4. Click **Process & Preview →**
   - Wait ~20–30 seconds while files are parsed
   - A table appears showing every employee, their score, and per-day breakdown
5. Review the **Non-compliant** table
   - Uncheck any employees who should be excluded (e.g., already spoken to)
   - Click **Preview** on any row to read the draft email before sending
6. Click **Send Emails →** → confirm → done
   - Emails send from your Gmail
   - Results are logged to the Google Sheet automatically

---

## How compliance is calculated

- Biometric (Chennai + Bangalore) is the **primary source** — if the system recorded a clock-in, the day counts
- Zoho leave codes (EL, SCL, MSL, ML, PLM, H) count as office days
- Ullen Ayyah (WFB / WFC / WFCO / L) is the **fallback** for employees not in biometric
- If a day has both biometric presence AND a leave code, it counts once
- Employees need **≥ 3 effective days** per week (Mon–Fri) to be compliant

## Attendance code reference

| Code | Meaning | Counts as office? |
|------|---------|-------------------|
| WFO | Work From Office (Bangalore biometric) | ✅ Yes |
| WFB | Work From Bangalore (Ullen Ayyah) | ✅ Yes |
| WFC | Work From Chennai (Ullen Ayyah) | ✅ Yes |
| WFCO | Work From Client Office (Ullen Ayyah) | ✅ Yes |
| EL | Earned (Planned) Leave | ✅ Yes |
| SCL | Unplanned / Sick Leave | ✅ Yes |
| MSL | Menstrual Leave | ✅ Yes |
| ML | Maternity Leave | ✅ Yes |
| H | Public / Government Holiday | ✅ Yes |
| WR | Work Remote / WFH (Ullen Ayyah) | ❌ No |
| L | Leave (Ullen Ayyah) | ✅ Yes |
| — (blank) | Absent or no record | ❌ No |

---

## Re-deploying after a code change

If you update `Code.gs` or `Index.html`:
1. Apps Script → **Deploy → Manage deployments**
2. Click the pencil ✏ icon → set Version to **New version** → **Deploy**
3. The same `/exec` URL keeps working — no need to reshare it

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| "Drive API not enabled" | Apps Script → Services → Add Drive API |
| "Permission denied" | Re-run `setupSheets()` and re-approve all permissions |
| Emails not sending | Check Gmail daily limit: 500/day (free), 1500/day (Workspace) |
| "No employees found" | Make sure Ullen Ayyah CSV has `Employee ID`, `Employee Name`, `Email` columns |
| File parse error | Check correct file is uploaded in each slot; Bangalore can be CSV or XLSX |
