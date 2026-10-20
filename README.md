# Exam invigilators

A small website where invigilators type their name and see where they are invigilating each week, with a PDF download. The exams office manages the timetable from an admin page by uploading the Excel spreadsheet and making changes in the browser.

- **Invigilators** use the main address, for example `https://your-site.netlify.app/`
- **The exams office** uses `/admin.html`, for example `https://your-site.netlify.app/admin.html`

It follows the LwLAT Brand Guidelines (v1.0, October 2026): teal as the lead colour, Arial as the everyday font and the master logo artwork from the Marketing SharePoint.

## Schools

The main page opens with three tiles: Lancaster Academy, Babington Academy and South Wigston High School. Each school has its own timetable. A school with no published timetable shows "Coming soon".

The admin page uses one password for all three, with a tab for each school. Only Lancaster Academy can be uploaded to or changed at the moment. The other two show "Not set up yet", and the server refuses any change to them.

**To switch on another school later:** in `netlify/functions/timetable.mjs`, change `editable: false` to `editable: true` for that school, and in `admin.html` change the "Upload and edit" label line so it covers that school too. Then upload the files to GitHub as usual.

## How it works

The site is plain HTML pages hosted on Netlify. There is one small server function (`netlify/functions/timetable.mjs`) that checks passwords and stores the timetable in Netlify's built-in storage (Netlify Blobs). There is no separate database, and nothing is connected to SharePoint, Microsoft 365 or Claude. Excel files are read inside the exams officer's browser and only the timetable itself is sent to Netlify.

Each time the exams office saves, the previous version is kept. The last 30 versions can be restored from the admin page.

## Setting it up on Netlify (one-off, about 15 minutes)

1. Sign in to [Netlify](https://app.netlify.com) (a free account is enough).
2. Choose **Add new site**, then **Import an existing project**, then **GitHub**, and pick the `laninvigilators` repository.
3. Leave the build settings as they are and press **Deploy**.
4. In the site, go to **Site configuration**, then **Environment variables**, and add:
   - `ADMIN_PASSWORD`: the password for the exams office. Make it long.
   - `VIEWER_PASSWORD`: the password you give to invigilators. Leave this out if you want anyone with the link to see the timetable (not recommended).
5. Go to **Deploys** and choose **Trigger deploy**, then **Deploy site**, so the passwords take effect.
6. Optional: under **Domain management** you can rename the site (for example `lancaster-invigilation.netlify.app`) or add a school web address, which needs IT to add a DNS record.

To change a password later, edit the environment variable and redeploy (step 5).

## Using the admin page

1. Open `/admin.html` and sign in with the admin password.
2. **Upload spreadsheet**: drop in the `.xlsx` file. If the file is `.xlsb`, open it in Excel and use *File, Save As, Excel Workbook (.xlsx)* first. Check the summary, then press **Use this spreadsheet**.
3. **Timetable**: click any box and type to change it. Type `N/A` for not available. Leave it empty for no duty. Each row is one exam, with its own start time and duration. Use **+** to add another exam row to a session, the arrows to move a row within its day (between AM and PM too), and **✕** to remove one.
4. **Exam details**: each shaded row is one exam, with its name, start time and duration, which apply to every room under it. Rooms can have their own name or times. The **+ ↑ ↓ ✕** buttons add, move and remove exam rows here too.
5. **Names, dates and settings**: the heading, a notice for all invigilators, and adding, renaming or removing people and sessions.
6. Press **Save and publish**. Nothing changes for invigilators until you do.
7. **Download as Excel** gives you the current timetable in the same layout, with an extra "Exam details" sheet. You can upload that file again later.

### What the upload understands

- A row of first names across the top, with columns of days, dates (such as "9th Nov") and sessions (AM or PM) down the left.
- Each cell holds the room or duty, such as Hall, Gym, Ma 5 or Roving.
- Optional columns headed **Start time** and **Duration** give each session's exam start time and its duration, not including extra time. They can sit anywhere in the row. Times can be Excel times or text such as 9:00. Durations can be Excel times such as 1:30, a number of minutes, or text such as 1h 30m or 90 mins. A bare number of 8 or less is read as hours.
- If one room in a session has a different start time or duration, change it for that room on the Exam details tab.
- A session can have more than one exam row. Each extra row can be labelled AM or PM again, or left unlabelled under the row above. Either way it is kept as a separate exam.
- **Grey-shaded empty cells mean "not available"**, as do cells containing `N/A`.
- Cells reading `CLASH` are shown to the invigilator as "Clash: please speak to the Exams Officer".
- Dates without a year are given the year that matches the day name.

## Data protection

The site holds staff first names and where they will be during exam sessions. It holds no learner data. Please confirm the arrangements with the Trust's Data Protection Officer before going live, and record what was approved:

| Item | Detail |
| --- | --- |
| Data held | Invigilator first names, rooms and duties, exam names and times |
| Where it is stored | Netlify Blobs, as part of the Netlify site |
| Who can see it | Anyone with the invigilator password (or anyone with the link, if no password is set) |
| Who can change it | Anyone with the admin password |
| Kept for | Until deleted. The last 30 saved versions are kept |
| DPO approval | *Name, date and what was approved* |

Spreadsheets are deliberately excluded from this repository (see `.gitignore`), so staff names are never stored in GitHub.

## Files

| File | What it is |
| --- | --- |
| `index.html` | The invigilator page: name search, weekly timetable, PDF download |
| `admin.html` | The exams office page |
| `assets/common.js` | Shared helpers (dates, data, talking to the server) |
| `assets/excel.js` | Reads and writes the Excel spreadsheet |
| `assets/style.css` | Brand colours and layout |
| `assets/vendor/` | ExcelJS 4.4.0, jsPDF 2.5.2 and jsPDF-AutoTable 3.8.4, bundled so school web filters cannot block them |
| `netlify/functions/timetable.mjs` | Password checks and storage |
