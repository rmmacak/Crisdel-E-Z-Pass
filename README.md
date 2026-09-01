# Crisdel Toll Reconciliation Dashboard

A local dashboard for cross-referencing E-ZPass and SunPass toll statements
against the Crisdel fleet, flagging high-value transactions, and producing
one combined Excel report.

## What it does

1. You drop in the month's **E-ZPass PDF** and **SunPass PDF**.
2. You drop in the **Crisdel equipment/transponder mapping spreadsheet**
   (the one with columns for Truck #, License Plate, EZPASS Transponder ID,
   Sunpass, and Assigned Driver). The app remembers the last one you gave
   it, so you only need to re-upload this when the fleet list changes.
3. Click **Generate Report**.
4. Download the finished Excel workbook: all transactions matched to a
   Truck #/Driver, a summary tab per vehicle, a flagged-transactions tab
   (over your chosen dollar threshold), and an unmatched-transponders tab.

## Easiest way to run it: double-click launcher

No terminal, no typing commands. Just double-click the file for your OS:

- **Mac:** `start_dashboard_mac.command`
  - First time only, macOS may block it with a security warning. Right-click
    the file -> **Open** -> confirm **Open** in the dialog. After that,
    plain double-clicking works.
- **Windows:** `start_dashboard_windows.bat`

Either one will, the first time you run it: check that Python is installed,
set up a private environment for the app, and install the required
packages. Every time after that, it starts in a few seconds. Your browser
opens automatically to the dashboard. To stop it, just close the black
terminal/command window that opened alongside it.

(If you don't have Python at all yet, both scripts will tell you and point
you to python.org -- install it once, then double-click the launcher again.)

## Running it in GitHub Codespaces (no install needed)

This repo includes a `.devcontainer/devcontainer.json`, so Codespaces sets
itself up automatically:

1. Push this folder to a new GitHub repo (see below if you haven't done
   this before).
2. On the repo page, click the green **Code** button -> **Codespaces** tab
   -> **Create codespace on main**.
3. Wait for it to build (~1 minute) -- it runs `pip install -r
   requirements.txt` for you.
4. In the terminal that opens, run:
   ```bash
   streamlit run app.py
   ```
5. A popup will offer to open the forwarded port in your browser --
   click it (or open the **Ports** tab and click the globe icon next to
   port 8501). That's your dashboard, running entirely inside the
   Codespace.

**Pushing this folder to GitHub for the first time**, from this folder:
```bash
git init
git add .
git commit -m "Crisdel toll dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```
(Create the empty repo first at github.com/new -- don't initialize it
with a README, so the push above doesn't conflict.)

**A note on data:** your toll PDFs and equipment list are only ever
processed inside the Codespace's temporary container -- they aren't
committed to the repo unless you explicitly `git add` them. Since a
Codespace resets when deleted, don't rely on it for storing your monthly
reports; download each finished Excel file before you close the session.

## Deploying to Azure (public URL, always-on)

Use **Azure App Service** (not Azure Static Web Apps -- that product only
hosts static HTML/JS/CSS, and can't run a Python process like this app).
App Service runs the dashboard continuously and gives you a real URL like
`https://crisdel-toll-dashboard.azurewebsites.net` that anyone on your team
can open.

This repo already includes what Azure needs (`startup.sh` and
`.streamlit/config.toml`).

1. In the [Azure Portal](https://portal.azure.com), search **App Services**
   -> **Create** -> **Web App**.
2. Fill in:
   - **Publish:** Code
   - **Runtime stack:** Python 3.11
   - **Operating System:** Linux
   - **Region:** whichever is closest to you
   - **Pricing plan:** Free (F1) or Basic (B1) is plenty for this
3. Click **Review + create** -> **Create**. Wait for deployment to finish,
   then click **Go to resource**.
4. In the left menu, click **Deployment Center**.
   - **Source:** GitHub
   - Sign in and pick your repo and branch (`main`)
   - Save -- Azure will build and deploy automatically, and redeploy every
     time you push to GitHub from now on.
5. In the left menu, click **Configuration** -> **General settings**:
   - **Startup Command:** `bash startup.sh`
   - Save.
6. Still in **Configuration**, click **Application settings** -> **+ New
   application setting**:
   - Name: `WEBSITES_PORT`  Value: `8000`
   - Save (this tells Azure which port the app is listening on).
7. Give it a few minutes to redeploy after saving, then open the URL shown
   at the top of the Overview page.

**Worth knowing before you open this up to your team:**
- This URL will be reachable by anyone who has it, with no login screen,
  unless you add one. For toll/vehicle data, that's worth locking down:
  in the App Service resource, **Authentication** -> **Add identity
  provider** -> **Microsoft** lets you require an Azure AD (company
  Microsoft 365) sign-in before anyone can reach the app -- a few clicks,
  no code changes needed.
- The saved equipment list (`.cache/last_equipment_list.xlsx`) persists on
  App Service's disk between visits, but a fresh deployment (new git push)
  can reset it depending on your plan -- if the sidebar ever shows "no
  equipment list saved" unexpectedly, just re-upload it once.

## Running it locally via terminal (alternative to the launcher)

You'll need **Python 3.9+** installed. Then, from this folder:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Streamlit will print a local URL (usually `http://localhost:8501`) --
open that in your browser. The app runs entirely on your own machine;
nothing is uploaded anywhere else.

## Editing the generated Excel

The output is a completely normal `.xlsx` file -- open it in Excel and
edit, re-sort, or add to it like any spreadsheet. The Summary by Truck tab
uses live formulas (SUMIFS/COUNTIFS), so if you correct a Truck # or
Assigned Driver on the "All Transactions" tab, the summary numbers will
recalculate automatically when you save.

If you want to change what gets flagged, either:
- Use the **threshold box in the sidebar** before generating (no code
  changes needed), or
- Edit `crisdel_toll_report.py` directly -- the matching logic, column
  layout, and report formatting are all in one file, with comments
  marking each stage (mapping load -> PDF parsing -> matching -> Excel
  build) so you can find what to change.

## If the PDF layout ever changes

E-ZPass and SunPass occasionally tweak their statement format. If a new
month's PDF stops parsing correctly (e.g., the transaction count comes
back as 0), the likely fix is in `parse_ezpass()` or `parse_sunpass()` in
`crisdel_toll_report.py` -- specifically the column order assumed in the
`row[:12]` unpacking, and the header text used to detect the transaction
table (`'TAG # / PLATE'` / `'TRANSPONDER'`).

## Files in this folder

| File | Purpose |
|---|---|
| `app.py` | The dashboard UI (Streamlit) |
| `crisdel_toll_report.py` | All the actual logic: PDF parsing, vehicle matching, Excel report building |
| `requirements.txt` | Python packages needed |
| `start_dashboard_mac.command` | Double-click launcher (Mac) |
| `start_dashboard_windows.bat` | Double-click launcher (Windows) |
| `.cache/last_equipment_list.xlsx` | Auto-saved copy of your most recent equipment list (created after first run) |
