# HR Hiring Cost &amp; Vacancy Management System

A browser-based application that turns the **New Joiners Cost** hiring
spreadsheet into a professional workforce-cost planning tool: vacancy
management, monthly cost forecasting, and full 24-month employment-cost
analysis by company and department.

Runs entirely in the browser — no server, no database — and deploys on
**GitHub Pages**. Data lives in a single `data.json`; edits are saved to the
browser and exported back to JSON when you want to make them permanent.

---

## Features

- **Overview dashboard** — 9 KPI cards, six charts (monthly cost, cost by
  company, cost by department, vacancy status, cost breakdown), an auto-generated
  management-insights panel, and an upcoming-hires widget.
- **Global filters** — Year, Month, Company, Department, Status. Every KPI,
  chart and table updates live. One-click **Reset filters**.
- **Vacancy management** — add, edit, duplicate and delete vacancies through a
  guided form (no JSON editing). Search, filter, sort and paginate the table.
- **Monthly Cost matrix** — Jan–Dec grid per vacancy with automatic
  frequency-aware calculation, column/row totals, and year selection.
- **Forecast horizons** — current month, next 3 / 6 / 12 / 24 months.
- **Cost Analysis** — one-time vs recurring split, and a full 24-month
  breakdown by every cost line from the sheet.
- **Company dashboards** — a comparison table plus a dedicated dashboard per
  company (KPIs, department split, 12-month forecast, vacancy list).
- **Department analysis** — cost and headcount per department, drill-down to
  positions.
- **Cost audit** — click any vacancy for a cost timeline (24 monthly bars),
  a month-by-month table, and an itemised breakdown showing exactly what lands
  in each month.
- **Import / Export** — JSON export, CSV export (vacancies + monthly report),
  full backup, import, and reset to the original dataset.

---

## File structure

```
index.html    The complete application (HTML + CSS + JavaScript in one file)
data.json     The dataset (generated from the Excel sheet by convert.py)
convert.py    Excel -> data.json converter
README.md     This file
```

`index.html` loads Chart.js and two Google Fonts from CDNs; everything else is
self-contained.

---

## Requirements

- **To run the app:** any modern browser (Chrome, Edge, Firefox, Safari) and a
  way to serve the folder over HTTP (see below). No install needed.
- **To regenerate `data.json` from Excel:** Python 3.8+ with `pandas` and
  `openpyxl`:
  ```bash
  pip install pandas openpyxl
  ```

---

## Excel → JSON conversion

The Excel sheet is the source of truth for the field structure. Cost columns
are bucketed by their **spreadsheet column letter**:

| Columns | Meaning              | Charged                              |
|---------|----------------------|--------------------------------------|
| M–O     | One-time             | once, in the hiring month            |
| P–R     | Monthly recurring    | every month from the hiring month    |
| S       | Once every 2 years   | hiring month, then every 24 months   |
| T–X     | Annual               | hiring month, then every 12 months   |
| Y       | Monthly              | every month from the hiring month    |
| Z       | Two-year (Gratuity)  | hiring month, then every 24 months   |

Convert:

```bash
python convert.py hiring.xlsx
```

Produces `data.json`. Optional custom output name:

```bash
python convert.py hiring.xlsx custom-data.json
```

The converter validates required columns, converts dates to ISO format
(`2026-08-15`), turns blanks into `0`, generates unique vacancy IDs
(`VAC-2026-001`), and prints a data-quality report (missing company/department/
date, possible duplicates).

> **Note on the sheet:** the workbook has no *status* column and no employee
> names, so every row is imported as a **Planned** vacancy. Set status per
> vacancy inside the app. The sheet's `Total` column is **not** trusted — the
> app always recalculates the 2-year cost from the rules above.

---

## Running locally

Because the app fetches `data.json`, opening `index.html` directly from disk
(`file://`) is blocked by the browser. Serve the folder instead:

```bash
cd hr-hiring-app
python -m http.server 8000
```

Then open:

```
http://localhost:8000
```

---

## GitHub Pages deployment

1. Create a repository and add `index.html`, `data.json`, `convert.py`,
   `README.md`.
2. Push to GitHub.
3. Repository **Settings → Pages → Build and deployment**, set **Source** to
   *Deploy from a branch*, pick `main` and `/root`, save.
4. Your app is live at `https://<user>.github.io/<repo>/`.

`data.json` loads automatically over HTTPS on Pages.

---

## Adding a vacancy

Go to **Vacancies → + Add vacancy**. Fill in company, department, position,
status and hiring date, then the cost fields (each section is labelled with its
frequency). The **Auto-calculated** panel updates the 2-year cost live. Click
**Save vacancy** — a unique ID is assigned automatically.

## Editing a vacancy

Row **Actions (⋮) → Edit**, or open a vacancy and click **Edit**. Change any
field and save. All KPIs, charts, company/department totals and the 2-year cost
recalculate immediately.

## Deleting a vacancy

**Actions (⋮) → Delete**, then confirm in the dialog. Deletion is local; you can
always **Reset** to the original `data.json` from **Data &amp; Settings**.

## Duplicating a vacancy

**Actions (⋮) → Duplicate** creates a copy with a new unique ID — useful when
hiring several people for the same role.

---

## Importing data

**Data &amp; Settings → Import → Choose JSON file.** Load a file exported from
this app or produced by `convert.py`. This replaces the working dataset in your
browser.

## Exporting data

From **Data &amp; Settings**:

- **Export JSON** — the current dataset as `data.json` (drop-in replacement).
- **Export CSV (all)** — all vacancies with computed costs.
- **Monthly cost CSV** — the monthly matrix for the selected year.
- **Backup (.json)** — a dated full backup.

---

## Cost calculation rules

For a vacancy hired on date *H* (year `hy`, month `hm`), and any calendar month
*(Y, M)*, let `k = (Y − hy)·12 + (M − hm)`:

- If `k < 0` → cost is **0** (not yet hired).
- **Recurring** (P–R and Y) is added **every** month `k ≥ 0`.
- **One-time** (M–O) is added only when `k = 0`.
- **Annual** (T–X) is added when `k` is a multiple of **12**.
- **Two-year** (S and Z) is added when `k` is a multiple of **24**.

Costs are multiplied by the vacancy **quantity**. Cancelled vacancies are
excluded from all cost totals. The **2-year employment cost** is the sum of the
monthly cost over the first 24 months (`k = 0…23`), so it includes one-time ×1,
recurring ×24, annual ×2 and two-year ×1.

Worked example (MGE-UAE Mechanical Helper, hired Aug 2026):
one-time 1,918 + recurring 2,450×24 + annual 2,413×2 + two-year 6,000 =
**AED 71,544**.

---

## Data persistence

- On load, the app reads `data.json`, then merges any changes saved in this
  browser's `localStorage`.
- Add / edit / delete actions are written to `localStorage` immediately. A
  marker in the sidebar shows when you have unsaved-to-file changes.
- The architecture keeps the data layer separate from the UI, so a real
  backend/API can be added later without rewriting the frontend.

## GitHub Pages limitations

Changes made through Add / Edit / Delete are saved to **the current browser
only**. They do **not** update the GitHub repository or the original
`data.json`. To publish changes for everyone:

1. **Export JSON** from **Data &amp; Settings**.
2. Replace `data.json` in the repository with the exported file.
3. Commit and push.

## Updating the source Excel file

When the master spreadsheet changes, regenerate the dataset and redeploy:

```bash
python convert.py hiring.xlsx
```

Commit the new `data.json`. Anyone with local browser edits should export a
backup first (their local edits are not merged into the new file automatically).

---

## Troubleshooting

- **Blank page / "Could not load data.json"** — you opened `index.html` from
  disk. Run `python -m http.server 8000` and use `http://localhost:8000`.
- **Charts not showing** — check the network tab; the Chart.js CDN must be
  reachable. On an offline network, download `chart.umd.min.js` and reference it
  locally.
- **My edits disappeared** — edits live in one browser's `localStorage`; they
  don't move between browsers/devices and are cleared by **Reset**. Export JSON
  to keep them.
- **`convert.py` errors** — ensure `pip install pandas openpyxl` succeeded and
  that you passed the correct `.xlsx` path.

---

*Currency is displayed as AED. Cost frequencies follow the column rules above;
the 2-year employment cost is always recalculated, never taken from the sheet.*
