# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

A fully static, no-build-step web app (HTML + CSS + vanilla JS) for the Telecom Department. It runs directly from the filesystem or any static host (e.g. GitHub Pages). There is no package.json, no bundler, no server, and no test framework. "Running" the app means opening `index.html` in a browser.

## Architecture

### Module pattern
Every JS file exposes a single `const` IIFE that returns a public API object. Load order in `index.html` matters — each module depends only on those declared before it:

```
fileHandler.js  →  comparison.js  →  excelExport.js  →  siteIdJc.js  →  app.js
```

### Tab structure
The UI has two tabs. Each tab is fully independent — separate DOM IDs, separate state, separate logic:

| Tab | Panel ID | Logic owner |
|---|---|---|
| Tracking Update | `#panelTracking` | `app.js` |
| Site ID-JC File | `#panelSiteId` | `siteIdJc.js` |

Tab switching is handled by `initTabs()` in `app.js` using `aria-controls` as the link between button and panel.

### Data flow — Tracking Update tab

1. **FileHandler.readFile()** — reads any file SheetJS can parse; auto-detects the header row using a 3-strategy algorithm anchored by the `ID#` column hint
2. **findSheetWithId()** in `app.js` — selects the right sheet tab: coordinator files use ID-column auto-detection; master file looks for the tab named `"Invoicing Track"` by exact then partial name match
3. **Comparison.combineCoordinatorSheets()** — merges N coordinator sheets into one Map keyed by `id.toLowerCase()`
4. **Comparison.parseMasterData()** — builds a Map from the master sheet, also keyed by `id.toLowerCase()`
5. **Comparison.compare()** — iterates coordinator rows, classifies each as new/changed/unchanged
6. **checkJobCodeDuplicates()** in `app.js` — post-compare validation; flags any JC value assigned to more than one Site ID across all coordinator rows
7. **ExcelExport.generate()** — builds output workbook with sheets: Summary, New Entries, Changed Entries, Change Details, Combined Coordinators

### Fixed configuration constants (top of `app.js`)
```js
const ID_COLUMN       = 'ID#';
const MASTER_SHEET    = 'Invoicing Track';
const CASE_SENSITIVE  = false;
const INCLUDE_UNCHANGED = false;
```
These are hardcoded — there is no settings UI. Change them here if the source files change column/sheet names.

### Header row auto-detection (`fileHandler.js → detectHeaderRow`)
Three strategies in priority order:
1. **ID-hint scan** — finds the first row containing a cell that matches the supplied `idHint` string (e.g. `"ID#"`). Most reliable.
2. **Density scan** — finds the first row with ≥ 70 % of max column count AND ≥ 70 % text-label cells.
3. **Absolute fallback** — first non-empty row.

Always pass `ID#` as the `idColumnHint` when calling `FileHandler.readFile` so strategy 1 fires correctly for the master sheet.

### Column fuzzy-matching pattern
Used in both `siteIdJc.js` (`detectColumns`) and `app.js` (`findNormKey`): two passes — exact match against a list of terms, then contains match. Terms are ordered most-specific first to prevent shorter terms (e.g. `"jc"`) from shadowing longer ones (e.g. `"jc#"`).

### Date handling (`siteIdJc.js`)
SheetJS returns dates from this particular master file as `"dd-Mon-yy"` strings (e.g. `"29-Apr-25"`), not ISO format. The `parseDate()` function handles multiple formats; do **not** use plain string comparison against `"2026-01-01"` for Old/New classification — it will produce wrong results for this format.

## Key files

- **`js/app.js`** — Tracking Update tab wiring, `findSheetWithId()`, `checkJobCodeDuplicates()`, auto-trigger debounce (600 ms), tab switching
- **`js/siteIdJc.js`** — Site ID-JC tab, fully self-contained, exposes only `SiteIdJc.init()`
- **`js/comparison.js`** — pure data logic, no DOM; all Map keys are `.toLowerCase()` for case-insensitive ID matching
- **`js/fileHandler.js`** — file I/O and drag-drop; `setupDropZone` accepts all file types (SheetJS handles format errors at parse time)
- **`js/excelExport.js`** — output workbook builder; `publicHeaders()` strips the internal `__source__` column from all output sheets

## PWA
- `manifest.json` requires `icons/icon-192.png` and `icons/icon-512.png` to trigger the browser install prompt. These are not in the repo — open `icons/generate-icons.html` in a browser to download them.
- `sw.js` caches all static assets for offline use under cache key `task-tracker-v1`. Bump the version string there when deploying significant updates.
