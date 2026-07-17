# ITF152X Online Attendance Register — Setup Guide

Two files:
- **`Code.gs`** — the backend (Google Apps Script), stores everything in a Google Sheet.
- **`index.html`** — the standalone web page, used by both the lecturer and students (student view is triggered by a URL parameter, no separate hosting needed for that part).

The 57 students from `ITF152X_Class_list.xlsx` are already built into `Code.gs` — no manual data entry needed.

## 1. Create the backend (Google Apps Script)

1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet. Name it e.g. **"ITF152X Attendance Register"**.
2. In the sheet, go to **Extensions → Apps Script**.
3. Delete the default `myFunction() {}` code and paste in the entire contents of **`Code.gs`**.
4. Save (Ctrl/Cmd+S). Give the project a name if prompted.
5. In the function dropdown at the top of the editor, select **`setupSheets`** and click **Run**.
   - The first time, Google will ask you to authorize the script — click through **Review permissions → (your account) → Advanced → Go to project (unsafe) → Allow**. This is your own script, so this is expected.
   - Check the Execution log — it should say `Setup complete: 57 students loaded.`
   - Go back to the spreadsheet — you'll now see three tabs: **Students**, **Attendance**, **Sessions**.

## 2. Deploy it as a web app

1. In the Apps Script editor, click **Deploy → New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Configure:
   - **Execute as:** Me (your account)
   - **Who has access:** Anyone
4. Click **Deploy**, authorize again if prompted, and **copy the Web app URL** (ends in `/exec`).

> Any time you edit `Code.gs` afterwards, go to **Deploy → Manage deployments → ✏️ Edit → New version → Deploy** to publish the changes — editing the script alone does not update the live URL.

## 3. Connect the front end

1. Open `index.html` in a text editor.
2. Find this line near the top of the `<script>` section:
   ```js
   const APPS_SCRIPT_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
   ```
3. Replace the placeholder with the Web app URL you copied, e.g.:
   ```js
   const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycb.../exec";
   ```
4. Save the file.

## 4. Host the page

`index.html` is a single self-contained file — host it anywhere that serves static files:
- **Google Sites**, **GitHub Pages**, **Netlify/Vercel**, a **shared drive with a static host**, or your institution's LMS as an embedded page.
- You can also just open it locally in a browser for lecturer-only use, but for students to scan a QR code and reach it on their own phones, it needs to be at a real public URL.

Once hosted, that URL is what the lecturer uses day to day, and it's also what the QR codes will point to automatically.

## 5. Using it

**Lecturer view** (the URL with no parameters):
- Pick the class **date** (calendar picker, or type it directly).
- **Search** by student number, surname or first name to jump to a student.
- Every student defaults to **Absent**; click a student's status pill to toggle **Present/Absent**.
- **Mark all present / Mark all absent** for quick bulk actions.
- Click **💾 Save register** to write the day's marks to the Attendance sheet. Attendance % updates live above the table.
- **Generate QR code for selected date**: the browser will ask for your location — allow it. This "pins" the classroom location for that session. Display the QR code (e.g. on a projector); students scan it with their own phones.
- **Attendance summary**: pick a date range and click **Run summary** to see each student's sessions-present count and attendance %, then **⬇ Download CSV** for a file you can keep or import elsewhere.

**Student self sign-in** (what students see after scanning the QR code, URL like `...index.html?checkin=ABC123`):
- Their phone asks for location access — this is compared against the lecturer's location when the QR code was generated (within ~60 metres by default) to confirm they're in the same room. Adjust `MAX_CHECKIN_DISTANCE_METERS` in `Code.gs` if your venue needs a different radius.
- The student enters their **student number, surname and initials**. These are checked against the Students sheet — a mismatch is rejected.
- The backend also stores a random ID for that phone (in its browser storage). If the same phone tries to sign in a second time on the same day, it's rejected, so one phone can't mark several students.
- A student already marked present that day can't be marked again either.
- Successful sign-ins appear immediately in the lecturer's register with method "QR" — reload the register or revisit the date to see them.

## Notes & limitations

- **Location accuracy**: GPS accuracy indoors can vary (10–50m+). The default 60m radius is a reasonable starting point for a single classroom/lab, but widen or narrow it in `Code.gs` (`MAX_CHECKIN_DISTANCE_METERS`) to suit your venue.
- **QR sessions expire** after 3 hours (`SESSION_VALID_MINUTES` in `Code.gs`) and generating a new QR code for the same date automatically deactivates the previous one.
- **Data lives in the Google Sheet** behind the script — you can open it any time to view/edit raw data, or build your own charts/pivot tables from the `Attendance` tab.
- If you add, remove, or correct students later, edit the `Students` sheet directly, or update the `STUDENT_SEED` array in `Code.gs` and re-run `setupSheets` (note: re-running `setupSheets` clears and reloads the Students tab only — Attendance and Sessions history is also cleared, so only re-run it at the very start of a new module, not mid-term).
