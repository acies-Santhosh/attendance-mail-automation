# Attendance Compliance Web App — Setup Guide

## What you'll have after setup
A web app link HR can open in any browser. They upload 4 files,
review the compliance table, and click Send. Emails go from your
Gmail automatically. Everything logs to a Google Sheet.

---

## Step 1 — Create the Google Sheet (1 min)

1. Go to [sheets.google.com](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it: **Attendance Compliance Log**
4. Copy the Spreadsheet ID from the URL:
   `https://docs.google.com/spreadsheets/d/ **COPY THIS PART** /edit`
https://docs.google.com/spreadsheets/d/1pE1DEIpytWLRzsfUxSn9MQ-PFUE8Rrg5oodKEEj4noo/edit?usp=sharing
---

## Step 2 — Create the Apps Script project (3 min)

1. In your Google Sheet → **Extensions → Apps Script**
2. Delete everything in the default `Code.gs` file
3. Paste the entire contents of **`Code.gs`** (the file you downloaded)
4. Find this line at the top and paste your Spreadsheet ID:
   ```
   SPREADSHEET_ID: "YOUR_SPREADSHEET_ID_HERE",
   ```
5. Click **+ (New file) → HTML** → name it exactly **`Index`** (no .html)
6. Delete everything in it → paste the entire contents of **`Index.html`**
7. Save both files (Ctrl+S)

---

## Step 3 — Enable required APIs (2 min)

In Apps Script:
1. Click **Services (+)** on the left sidebar
2. Add **Drive API** (needed to parse uploaded Excel files)
3. Click Add

---

## Step 4 — Run setup once (1 min)

1. In Apps Script, select function: **`setupSheets`**
2. Click **Run**
3. Approve the permissions when prompted (uses your Gmail + Drive + Sheets)
4. You'll see a success message — two tabs are created in your Sheet

---

## Step 5 — Deploy as Web App (2 min)

1. Click **Deploy → New deployment**
2. Click the gear ⚙ next to "Select type" → choose **Web app**
3. Fill in:
   - Description: `Attendance Compliance v1`
   - Execute as: **Me** (your Gmail sends the emails)
   - Who has access: **Anyone** (so all HR can open the link)
4. Click **Deploy**
5. Copy the **/exec URL** — this is the link you share with HR

---

## Step 6 — Share the link

Send the `/exec` URL to everyone in HR. They open it in Chrome, 
upload files, and send emails. No login required, no install.

---

## Every week — HR workflow (3 clicks)

1. Open the web app link
2. Upload 4 files (Chennai biometric, Bangalore biometric, Zoho, Ullen Ayyah)
3. Confirm the week dates → **Process & Preview**
4. Review the table — click **Preview** on any row to see the email
5. Click **Send Emails →** → confirm → done

Emails send from your Gmail. Log writes to the Google Sheet automatically.

---

## Re-deploying after changes

If you update Code.gs or Index.html:
- Deploy → **Manage deployments** → Edit → **New version** → Deploy
- The same URL continues to work

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| "Drive API not enabled" | Apps Script → Services → Add Drive API |
| "Permission denied" | Re-run setupSheets() and approve all permissions |
| Emails not sending | Check Gmail daily send limit (500/day for free, 1500 for Workspace) |
| File parse error | Check that you're uploading the correct file for each slot |
