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

## Setup (one-time)

You'll need **Python 3.9+** installed. Then, from this folder:

```bash
pip install -r requirements.txt
```

## Running it

```bash
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
| `.cache/last_equipment_list.xlsx` | Auto-saved copy of your most recent equipment list (created after first run) |
