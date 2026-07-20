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
fileHandler.js → appData.js → comparison.js → excelExport.js
→ siteIdJc.js → pocTracking.js → allowanceChecker.js
→ adminSettings.js → app.js
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
| Settings (admin) | `#panelSettings` | `adminSettings.js` |

Tab switching is handled by `initTabs()` in `app.js` using
`aria-controls` as the link between button and panel.

The **Settings** tab button is hidden by default. It is revealed by
clicking the sidebar logo (`#brandLogo`) **7 times within 2 s** — the
gesture is wired in `app.js` and calls `AdminSettings.reveal()`.

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
- **AppList Google Sheet** (via `appData.js` → Apps Script web app) —
  loaded on startup from a remote JSON endpoint (see "App Data / Settings"
  below). Replaces the old `list.xlsx`. Provides:
  - Google Sheet URLs: Sheet Name + URL for each coordinator's Google Sheet
  - Team & Driver salaries: Name, **full monthly Salary**, Account Number.
    `appData.js` derives the daily rate as `round(monthly / 26)` and exposes
    it as `dailySalary` — every downstream consumer is unchanged.
- Master Tracking Excel file (the Site ID-JC file) — uploaded fresh
  at the start of each analysis run. Must contain a sheet whose name
  includes "Tracking" with columns `SiteID-JC` and `Old/New`.

**Flow:**
1. On startup, `appData.js` loads Google Sheet URLs and salary data
   from the AppList Apps Script endpoint into app state
2. User selects Month and Month Half (First / Second) from dropdowns
3. User uploads the Master Tracking Excel file (Site ID-JC file)
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
8. Three download actions are available:
   - **Download Output** (green) — the original complete report named
     `Allowance_Report_[Month]_[Half].xlsx` with tabs:
     - **[Month] - [Half]**: filtered tracking rows
     - **Allowance Amount**: two tables — Team (Name, Total Amount,
       Account Number) and Driver (Name, Amount)
     - **Per Person**: one row per team member per source row,
       sectioned by role (see Per Person tab below)
   - **New/Old Files** (red) — generates two separate workbooks with the
     same structure, splitting rows by Old/New classification:
     - `Allowance_Old_[Month]_[Half].xlsx`
     - `Allowance_New_[Month]_[Half].xlsx`
     - Each workbook also contains a **Per Person** tab scoped to
       its Old or New rows only
   - **Cost per Site** (blue) — generates `SiteCost_[Month]_[Half].xlsx`
     with one row per site per day (see Cost per Site section below)

**Results UI:**
- Stat cards row: Sheets fetched · Total rows · Filtered count ·
  Grand Total · **Avg Team Utilization** (team members only, `Math.ceil`)
- **Team Utilization table**: Name, Days Worked (unique days via Set),
  Utilization % (`Math.ceil(days/13×100)`). Team members only — drivers
  excluded. Sorted by days worked descending, then alphabetically.
- Data Sources table: rows loaded vs matched per coordinator sheet

**Per Person tabs** (present in all Allowance Checker workbooks):
- Each employee gets their own dedicated worksheet tab named after them
- Tab order: Engineers (alphabetical) → Tech-1 → Tech-2 → Tech-3 →
  Drivers (alphabetical within each role)
- Each tab contains: orange section-label row (Engineer / Technicians /
  Driver) → blue column-header row → data rows sorted by day → green
  total row (`"Name — Total"` with summed Allowance and Vacation Allowance)
- Columns: Month, Day, Month Half, Coordinator, Site, Area, Project,
  Sub Project, Name, Allowance, Vacation Allowance, Work Details, JC
  (Start Time, End Time, and Role are intentionally excluded)
- Vacation Allowance column shows the person's actual daily salary
  amount (looked up from the `AppData` salary map) — blank when no vacation
- Empty roles produce no tabs. Duplicate tab names (same employee name
  in multiple roles) get a `_2` suffix to avoid collisions.
- Built by `buildPerPersonSheets()` in `allowanceChecker.js`

**Key rules:**
- Empty team member fields = not counted (truly absent, not zero)
- Vacation allowance is per person, based on individual Salary/Day
- Site-JC pairing is positional (index-based)
- All errors must appear before output is generated
- Days worked = unique calendar days (Set-based), not row count

**Cost per Site output** (`SiteCost_[Month]_[Half].xlsx`):
- One sheet, one row per site per day
- Columns: Date, Site ID, Job Code, Cost/Site, Coordinator, Engineer,
  Tech-1, Tech-2, Tech-3, Driver
- **Date** is formatted as `{day}-{monthAbbr}` (e.g. `15-Jan`) from row data
- **Cost/Site calculation**:
  1. For every row compute its total cost:
     `rowCost = allowancePerPerson × memberCount + Σ dailySalary (vacation rows only)`
  2. Group by day → sum all row costs and all site counts for that day
  3. `costPerSite = totalDayCost / totalSitesForDay`
  4. Multi-site rows (e.g. `K3960 / K5402`) are expanded — one output row
     per site, same cost/site value, team-member fields repeated on each row
- Rows with no site data count as 1 site and produce one output row with
  a blank Site ID
- Output rows are sorted by day (ascending), then Site ID
- Built by `buildSiteCostSheet()` and `generateSiteCostFile()` in
  `allowanceChecker.js`; button ID is `allowanceSiteCostBtn`

**New/Old split rules (applied per site-JC pair, not per row):**
- JC contains "CCTV" (case-insensitive) → **always Old**, highest priority
- Combo found in master's `Old/New` column with value "Old" → Old
- Everything else (not found, blank, "New") → New
- Mixed rows (e.g. `K3960 / K5402` where one is Old and one is New)
  are **split into separate rows** per file — Site and JC columns are
  rebuilt to contain only the pairs belonging to that file. The allowance
  is divided proportionally (`origAllow / totalPairs × pairsInThisFile`).
  A row never bleeds the wrong site into the wrong file.
- The `masterOldNewMap` is populated by `parseMasterTracking()` at the
  same time as `masterJcSet` — no second file upload needed.

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
  reset logic, 7-click logo gesture for the Settings tab
- **`js/pocTracking.js`** — POC Tracking tab, same structure as
  app.js but keyed on Job Code and POC3 Tracking sheet
- **`js/siteIdJc.js`** — Site ID-JC tab, fully self-contained
- **`js/allowanceChecker.js`** — Allowance Checker tab, reads
  data from `AppData`, fetches Google Sheets, runs all analysis
- **`js/appData.js`** — loads the AppList (Google Sheet URLs +
  salaries) from the Apps Script endpoint on startup, derives the
  daily rate (`monthly / 26`), exposes getters + `getRawData()` /
  `setData()` / `reload()` for the Settings tab
- **`js/adminSettings.js`** — admin Settings tab: editable URL/Team/
  Driver tables, saves back to the Google Sheet via `doPost`
- **`js/comparison.js`** — pure data logic, no DOM
- **`js/fileHandler.js`** — file I/O and drag-drop
- **`js/excelExport.js`** — output workbook builder
- **`apps-script/Code.gs`** — Google Apps Script web app backing the
  AppList: `doGet` returns JSON, `doPost` (password-guarded) writes
  the three Sheet tabs. Deployed manually; its `/exec` URL is the
  `APPLIST_ENDPOINT` constant in `appData.js`

### App Data / Settings (AppList)
The reference list that used to live in `list.xlsx` now lives in a
Google Sheet with three tabs — `Google Sheets URLs`, `Team Salaries`,
`Driver Salaries` (Salary column = **full monthly** amount). An Apps
Script web app (`apps-script/Code.gs`) serves it as JSON and accepts
password-guarded writes. The password is validated **server-side only**
— it is never stored in the client.

**Access flow:** the 7-click logo gesture reveals the Settings tab,
which opens on a **password lock screen**. Entering the password fires
`doPost` with `action:'login'`, which validates it server-side and, on
success, returns the current data — used to populate the tables (so the
tab never depends on the slow startup fetch having finished). The
verified password is kept in memory for the session and reused when
saving. Reopening the app hides the tab again and re-requires the
password (session-only reveal). Both login and save POST with
`Content-Type: text/plain` to avoid a CORS preflight (Apps Script can't
answer OPTIONS). `list.xlsx` is retired (no fallback).

**Redeploy gotcha:** editing `apps-script/Code.gs` has no effect until
the Apps Script web app is redeployed. To keep the same `/exec` URL,
edit the **existing** deployment (Deploy → Manage deployments → pencil →
New version). A brand-new deployment mints a **different** URL, which
must then be pasted into `APPLIST_ENDPOINT` in `appData.js`.

## PWA

- `manifest.json` requires `icons/icon-192.png` and
  `icons/icon-512.png`
- `sw.js` caches all static assets for offline use
- **Always bump the cache version string in `sw.js` before
  pushing any update**
- Current cache version: `task-tracker-v2.190`
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

### Allowance Checker: New/Old split — pair-level not row-level
The initial implementation copied the entire row (with all sites) into
both Old and New files for mixed rows. This caused a "New" site like
K3960 to appear in the Old file when paired with an Old site. Fixed in
`buildSplitTrackingRows()`: each site-JC pair is classified individually,
and the `site` and `jc` fields of the output row are rebuilt to contain
only the pairs that belong to that file. The allowance is divided as
`origAllow / totalPairs × countInThisFile`.

### Allowance Checker: CCTV → always Old
Any JC value containing "CCTV" (case-insensitive) is unconditionally
classified as Old in the New/Old split, regardless of what the master
file's Old/New column says. This rule takes priority over the lookup.
Implemented inside `classifyRowPairs()` as the first check per pair.

### Allowance Checker: Old/New map comes from the master file
The `Old/New` column is read from the same Tracking tab of the Site
ID-JC master file that is already uploaded for JC validation.
`parseMasterTracking()` now returns both `jcSet` and `oldNewMap` in one
pass. No second file upload is required — the "New/Old Files" button
works as soon as `runAnalysis()` has completed.

### Allowance Checker: error boxes show coordinator source context
Missing Names and Missing Job Codes boxes follow the same format as
Repeated Names — each entry shows which coordinator sheet(s) triggered it.

**Missing Names** (`computeAllowances()`): replaced the `seenNames` Set
(which suppressed after first occurrence) with a `missingMap` that
accumulates every source sheet where the name appears. Errors are built
as pre-HTML strings at the end and rendered without escaping (same as
`repeatedErrors`). Format: `"Name" was not found in X Salaries list —
found in N sheets: [Sheet A] [Sheet B]`.

**Missing Job Codes** (`runAnalysis()`): `missingCombos` now stores
`{ display, sources: Set }` instead of just the display string. The
warning is pre-HTML, showing each combo with its source sheets.

**Site with no JC** (`runAnalysis()`): rows where `site` is non-empty
but `jc` is empty were previously silently skipped (`!siteRaw || !jcRaw`).
Now collected into `missingJcRows` (Map: source → `[{ site, day }]`) and
rendered as a separate item in the Missing Job Codes box. Format:
`"N rows with a site but no JC: [Sheet A] — K rows: [Site / Day] …"`.

### Site ID-JC: date handling
- All source dates are normalised to `dd-mmm-yyyy` on output regardless
  of source format (ISO, slash-delimited, SheetJS serial, etc.)
- The Old/New cutoff uses `new Date(2026, 0, 1)` — local-time
  constructor — not `new Date('2026-01-01')` which is UTC and would
  misclassify Jan 1 as "Old" in UTC+ timezones (e.g. Egypt UTC+2/+3).

### Allowance Checker: download button color convention
Three buttons sit in the `download-bar` of the results section:
- **Download Output** — `btn-success` (filled green) — the main full report
- **New/Old Files** — `btn-outline-red` (filled red, white text) — split output
- **Cost per Site** — `btn-primary` (filled blue, white text) — site cost report

`btn-outline-red` is a custom class in `css/styles.css` (background `#e02424`,
white text). It is not a Bootstrap utility — don't rename it to `btn-danger`.

### Allowance Checker: Per Person — one tab per employee, not one shared tab
Each employee gets their own worksheet tab rather than a shared "Per Person"
sheet. Tab order mirrors the role ordering: engineers first, then Tech-1/2/3
(as separate role passes), then drivers — all sorted alphabetically within
each role. A person who appears as both Tech-1 and Tech-2 in different rows
will get two tabs (one per role); their tab name gets a `_2` suffix on the
second occurrence to avoid Excel collisions.

Each tab has the same layout: orange section-label row (Engineer /
Technicians / Driver), blue column headers, data rows sorted by day, and a
green total row. Tab names are sanitised (forbidden Excel chars replaced with
`_`, truncated to 31 chars).

The Vacation Allowance column resolves to the actual EGP amount (each
person's `dailySalary` from `AppData`) rather than the raw flag text.
The total row sums Allowance and Vacation Allowance independently.
Start Time, End Time, and Role columns are excluded by design.
