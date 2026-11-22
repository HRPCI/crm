# HRPCI PR CRM – setup (about 20 minutes, one time)

Your Google Sheet stays the database. The mobile app (PWA) reads and writes your PR 1 to PR 5 tabs.
Each PR executive logs in with a PIN and sees only their own leads. You (Admin) see everyone.

## Step 1 – Make your tracker a real Google Sheet
Open HRPCI_Leads_Tracker_PR_Sheets.xlsx in Drive, then File > Save as Google Sheets.
(Apps Script does not work on .xlsx files.) Keep tab names PR 1 to PR 5, Setup, as they are.

## Step 2 – Add the backend
1. In the new Google Sheet: Extensions > Apps Script.
2. Delete the sample code, paste all of Code.gs, click Save.
3. Choose function `setup` at the top, click Run, allow permissions.
   A new tab AppUsers appears with a PIN for each executive and for ADMIN. You can change PINs there.
4. Deploy > New deployment > Web app.
   Execute as: Me. Who has access: Anyone. Click Deploy and copy the Web app URL.

## Step 3 – Connect and publish the app
1. Open index.html, find the line `const API_URL = '...'` and paste your Web app URL.
2. Put all files in a GitHub repo and connect it to Cloudflare Pages (same as your other sites).
   Use a subdomain, for example crm.hrpci.org. HTTPS is automatic.
3. Send the link and the PIN to each PR executive.

## Step 4 – Install on phone
Open the link in Chrome (Android) and tap Install app / Add to Home Screen.
On iPhone: Safari > Share > Add to Home Screen.

## Good to know
- If you change Code.gs later: Deploy > Manage deployments > edit > New version.
- Rows 8 to 407 of each PR tab are used. The app never touches formula columns.
- To stop an executive: set Active to N in AppUsers.
- PINs are checked on the server. Wrong PIN 6 times locks that user for 15 minutes.
- Do not share the Sheet itself with executives. They only need the app.
