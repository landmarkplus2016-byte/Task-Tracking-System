# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working
with code in this repository.

## What this project is

A fully static, no-build-step web app (HTML + CSS + vanilla JS) for the
Telecom Department. It runs directly from the filesystem or any static
host (e.g. GitHub Pages). There is no package.json, no bundler, no
server, and no test framework. "Running" the app means opening
`index.html` in a browser.

## Architecture

### Module pattern

Every JS file exposes a single `const` IIFE that returns a public API
object. Load order in `index.html` matters — each module depends only
on those declared before it:
```
fileHandler.js → comparison.js → excelExport.js → siteIdJc.js
→ allowanceChecker.js → app.js
```

### Tab structure

The UI has four tabs. Each tab is fully independent — separate DOM IDs,
separate state, separate logic:

| Tab | Panel ID | Logic owner |
|---|---|---|
| RF-TX Tracking Update | `#panelTracking` | `app.js` |
| POC Tracking Update | `#panelPocTracking` | `pocTracking.js` |
| Site ID-JC File | `#panelSiteId` | `siteIdJc.js` |
| Allowance Checker | `#panelAllowanceChecker` | `allowanceChecker.js` |

Tab switching is handled by `initTabs()` in `app.js` using
`aria-controls` as the link between button and panel.

---

## Tab Descriptions

### 1. RF-TX Tracking Update (`#panelTracking`)
Compares coordinator Excel files against a master tracking file to
detect new and changed entries.

- User uploads one or more coordinator files and one master file
- The app auto-detects the header row and the correct sheet tab
- Coordinator sheets are merged into one dataset keyed by `ID#`
- Each row is classified as: New, Changed, or Unchanged vs the master
- **New entry filter**: an ID is only classified as "New" if it is
  absent from **both** the `"Invoicing Track"` sheet **and** the
  `"Old Tasks"` sheet in the master file. If the `"Old Tasks"` sheet
  does not exist in the master file, this filter is silently skipped.
- Post-comparison, duplicate Job Codes across different Site IDs are flagged
- Output is a downloadable Excel file with tabs: New Entries, Collective Tasks
- A "↺ New Analysis" button resets all state for a fresh run

### 2. POC Tracking Update (`#panelPocTracking`)
Same structure as RF-TX Tracking Update but uses different identifiers.

- Keyed on `"Job Code"` column instead of `ID#`
- Looks for the sheet tab named `"POC3 Tracking"` in the master file
- Otherwise identical flow to RF-TX: upload → compare → export

### 3. Site ID-JC File (`#panelSiteId`)
Validates and processes Site ID to Job Code mapping files.

- User uploads one or more tracking files
- Sheet detection: any sheet whose name **contains "Tracking"**
  (case-insensitive). If multiple matching sheets exist in one file,
  an error is shown listing the ambiguous names — the file is skipped.
- Column detection uses fuzzy matching for both PC and POC variants:
  - Site ID: `"Physical Site ID"` (PC) | `"Site ID"` (POC)
  - Job Code: `"Job Code"` (both)
  - Task Date: `"Task Date"` (PC) | `"Installation Date"` (POC)
  - Contractor: `"Contractor"` (PC) | `"Installation Team"` (POC)
- Dates are parsed from any recognised format and **always output as
  `dd-mmm-yyyy`** (e.g. `21-Mar-2026`) in the Excel file
- Old/New cutoff is **2026-01-01 local time** — dates on or after
  2026-01-01 are "New"; dates before are "Old". The cutoff is
  constructed with `new Date(2026, 0, 1)` (not from a string) to
  avoid UTC-offset misclassification in non-UTC timezones.
- Output is a single-sheet Excel file: Site ID-JC | Task Date | Old/New | Contractor
- Fully self-contained, no dependency on other tabs

### 4. Allowance Checker (`#panelAllowanceChecker`)
Calculates team allowances for a selected month/half-month period
by combining data from multiple Google Sheets and validating it
against a master tracking file.

**Data Sources:**
- `list.xlsx` (in project root) — loaded on startup, contains:
  - `"Google Sheets URLs"` tab: Sheet Name + URL for each coordinator's
    Google Sheet
  - `"Salaries"` tab: Name, Account Number, Salary, Salary/Day for
    each team member
- Master Tracking Excel file — uploaded fresh at the start of
  each analysis run

**Flow:**
1. On startup, reads `list.xlsx` to load Google Sheet URLs and
   salary data into app state
2. User selects Month and Month Half (First / Second) from dropdowns
3. User uploads the Master Tracking Excel file
4. App fetches and combines all Google Sheets from the URLs list
   using `{ cache: 'no-store' }` so every run hits the network fresh
5. Combined data is filtered by selected Month + Month Half
6. Analysis runs on the filtered dataset:
   - **Allowance Calculation**: counts non-empty team member fields
     (Engineer, Tech-1, Tech-2, Tech-3, Driver), multiplies by the
     Allowance value, then adds vacation allowance per person based
     on their Salary/Day
   - **Repetition Check**: flags any team member appearing more than
     once on the same day
   - **Site-JC Validation**: splits Site and JC fields by `/`, pairs
     them positionally, checks each pair against the Master Tracking
     file, and flags missing or mismatched combos
7. If errors exist, they are shown in the Errors & Warnings panel
   before output generation
8. Output is a downloadable Excel file named
   `Allowance_Report_[Month]_[Half].xlsx` with tabs:
   - **Total Tracking**: all combined rows from Google Sheets
   - **Allowance Amount**: two tables — Team (Name, Total Amount,
     Account Number) and Driver (Name, Amount)

**Results UI:**
- Stat cards row: Sheets fetched · Total rows · Filtered count ·
  Grand Total · **Avg Team Utilization** (team members only, `Math.ceil`)
- **Team Utilization table**: Name, Days Worked (unique days via Set),
  Utilization % (`Math.ceil(days/13×100)`). Team members only — drivers
  excluded. Sorted by days worked descending, then alphabetically.
- Data Sources table: rows loaded vs matched per coordinator sheet

**Key rules:**
- Empty team member fields = not counted (truly absent, not zero)
- Vacation allowance is per person, based on individual Salary/Day
- Site-JC pairing is positional (index-based)
- All errors must appear before output is generated
- Days worked = unique calendar days (Set-based), not row count

---

## Fixed Configuration Constants (top of `app.js`)
```js
const ID_COLUMN        = 'ID#';
const MASTER_SHEET     = 'Invoicing Track';
const CASE_SENSITIVE   = false;
const INCLUDE_UNCHANGED = false;
```

These are hardcoded — there is no settings UI.

## Key Files

- **`js/app.js`** — RF-TX tab wiring, tab switching,
  `findSheetWithId()`, `checkJobCodeDuplicates()`, Old Tasks filter,
  reset logic
- **`js/pocTracking.js`** — POC Tracking tab, same structure as
  app.js but keyed on Job Code and POC3 Tracking sheet
- **`js/siteIdJc.js`** — Site ID-JC tab, fully self-contained
- **`js/allowanceChecker.js`** — Allowance Checker tab, reads
  list.xlsx on startup, fetches Google Sheets, runs all analysis
- **`js/comparison.js`** — pure data logic, no DOM
- **`js/fileHandler.js`** — file I/O and drag-drop
- **`js/excelExport.js`** — output workbook builder
- **`list.xlsx`** — reference data file (Google Sheet URLs +
  salary data), stored in project root, updated manually and
  pushed to GitHub when changes are needed

## PWA

- `manifest.json` requires `icons/icon-192.png` and
  `icons/icon-512.png`
- `sw.js` caches all static assets for offline use
- **Always bump the cache version string in `sw.js` before
  pushing any update**
- Current cache version: `task-tracker-v2.168`
- Version format: always two digits after the dot (e.g. `v2.10`,
  `v2.11`) — never single digit minor (not `v2.9`)

## Deployment Checklist

1. Bump version in `sw.js` (e.g. `v2.160` → `v2.161`)
2. Commit with a descriptive message
3. Push to GitHub
4. Reopen the installed app to load the update

## Known Decisions & Gotchas

### Site ID-JC: sheet detection
Previously hard-coded to `"Invoicing Track"` and `"POC3 Tracking"`.
Changed to match **any sheet name containing "Tracking"** so that
variant names like `"Tracking"` or `"Gendy Tracking"` are accepted.
Ambiguity (multiple Tracking sheets in one file) is an error, not
a silent pick.

### RF-TX: Old Tasks filter
The master file may contain an optional `"Old Tasks"` sheet alongside
`"Invoicing Track"`. After `Comparison.compare()` produces its
`newEntries` list, `app.js` loads this sheet (using the same
`parseMasterData()` path) and filters out any entry whose `ID#` appears
in it. The filter is case-insensitive (both sides are lowercased).
If the sheet is absent, no error is raised and no entries are filtered.
This logic lives entirely in `runProcess()` in `app.js` — `comparison.js`
and `excelExport.js` are unchanged.

### Excel output file size
`excelExport.js` previously applied border styles to every data cell,
and `XLSX.writeFile` used no ZIP compression (store mode). Combined,
these inflated a ~4 MB input to ~11 MB output. Both fixed:
- `applyStyles()` now only styles the header row — data rows carry no
  style objects, which eliminates the per-cell XML bloat.
- `XLSX.writeFile(wb, filename, { compression: true })` — enables deflate
  compression on the output ZIP, bringing file size back in line with
  the input.

### Allowance Checker: Google Sheets fetch always hits network
Two layers of caching were silently serving stale CSV data:
1. **Browser HTTP cache** — fixed with `fetch(url, { cache: 'no-store' })`
   in `fetchGoogleSheets()`.
2. **Service Worker Cache Storage** — the SW's cache-first handler was
   intercepting the fetch and returning a cached CSV before it ever
   reached the network. `{ cache: 'no-store' }` does NOT bypass SW cache.
   Fixed in `sw.js` by only intercepting same-origin requests and
   explicitly whitelisted external assets (CDN). Any other external URL
   (Google Sheets export, etc.) is not intercepted at all — the browser
   handles it directly.

### CSS specificity: numeric column alignment
`.allowance-table th` (specificity 0,1,1) overrides `.allowance-th-num`
(0,1,0), causing header cells to stay left-aligned even when the class
sets `text-align: center`. Fixed by scoping the rule to
`.allowance-table .allowance-th-num` (0,2,0).

### Site ID-JC: date handling
- All source dates are normalised to `dd-mmm-yyyy` on output regardless
  of source format (ISO, slash-delimited, SheetJS serial, etc.)
- The Old/New cutoff uses `new Date(2026, 0, 1)` — local-time
  constructor — not `new Date('2026-01-01')` which is UTC and would
  misclassify Jan 1 as "Old" in UTC+ timezones (e.g. Egypt UTC+2/+3).
