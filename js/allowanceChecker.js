/**
 * allowanceChecker.js
 * ──────────────────────────────────────────────────────────────
 * Allowance Checker tab logic.
 *
 * UI flow:
 *   1. User picks a Month and Month Half (First / Second)
 *   2. User uploads the Master Tracking file (.xlsx)
 *   3. Run Analysis → reads master file, then fetches every Google
 *      Sheet URL stored in AppData (loaded from list.xlsx), combines
 *      them into a single unified dataset, then runs analysis.
 *   4. Results + any errors/warnings are shown.
 *
 * Load order: must come after fileHandler.js and appData.js
 */

const AllowanceChecker = (() => {
    'use strict';

    const MONTHS = [
        'January','February','March','April','May','June',
        'July','August','September','October','November','December',
    ];

    /* Unified column schema — keys used throughout this module */
    const COLUMNS = [
        { key: 'month',             terms: ['month'] },
        { key: 'day',               terms: ['day'] },
        { key: 'monthHalf',         terms: ['month half', 'monthhalf', 'half'] },
        { key: 'coordinator',       terms: ['coordinator', 'coord'] },
        { key: 'site',              terms: ['site id', 'site name', 'site'] },
        { key: 'area',              terms: ['area', 'region', 'zone'] },
        { key: 'startTime',         terms: ['start time', 'time start', 'time in', 'start'] },
        { key: 'endTime',           terms: ['end time', 'time end', 'time out', 'end'] },
        { key: 'project',           terms: ['project', 'proj'] },
        { key: 'subProject',        terms: ['sub project', 'sub-project', 'subproject', 'sub proj'] },
        { key: 'engineer',          terms: ['engineer', 'eng'] },
        { key: 'tech1',             terms: ['tech-1', 'tech 1', 'technician 1', 'tech1'] },
        { key: 'tech2',             terms: ['tech-2', 'tech 2', 'technician 2', 'tech2'] },
        { key: 'tech3',             terms: ['tech-3', 'tech 3', 'technician 3', 'tech3'] },
        { key: 'driver',            terms: ['driver', 'drv'] },
        { key: 'allowance',         terms: ['allowance', 'allow'] },
        { key: 'vacationAllowance', terms: ['vacation allowance', 'vac allowance', 'vacation'] },
        { key: 'workDetails',       terms: ['work details', 'work description', 'details', 'description'] },
        { key: 'jc',                terms: ['jc#', 'jc', 'job code', 'job#', 'jobcode'] },
    ];

    const state = {
        masterFile:    null,
        sheetRows:     [],   // unified rows from all Google Sheets
        filteredRows:  [],   // rows matching selected month + half
        results:       null,
    };

    const $ = id => document.getElementById(id);

    /* ── HTML escape ─────────────────────────────────────────── */
    function esc(str) {
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    /* ── Column fuzzy-match (exact first, then contains) ─────── */
    function findColIdx(headers, terms) {
        const lower = headers.map(h => h.toLowerCase().trim());
        for (const t of terms) {
            const i = lower.indexOf(t);
            if (i !== -1) return i;
        }
        for (const t of terms) {
            const i = lower.findIndex(h => h.includes(t));
            if (i !== -1) return i;
        }
        return -1;
    }

    /* ── Google Sheets URL helpers ───────────────────────────── */

    /**
     * Extract the spreadsheet ID from any Google Sheets URL.
     * Handles share links, edit links, and published links.
     */
    function extractSheetId(url) {
        const m = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        return m ? m[1] : null;
    }

    /**
     * Extract the gid (sheet tab ID) from a URL.
     * Falls back to '0' (first sheet) when absent.
     */
    function extractGid(url) {
        const m = url.match(/[#&?]gid=(\d+)/);
        return m ? m[1] : '0';
    }

    /**
     * Build a CSV export URL from any Google Sheets share/edit link.
     * Returns null when the URL is not a recognisable Sheets URL.
     */
    function buildExportUrl(url) {
        const id  = extractSheetId(url);
        const gid = extractGid(url);
        if (!id) return null;
        return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`;
    }

    /* ── Parse raw 2-D array → unified row objects ───────────── */
    function parseSheetRows(raw, sourceName) {
        if (!raw.length) return [];

        const headerIdx = FileHandler.detectHeaderRow(raw, 40);
        const headers   = (raw[headerIdx] || []).map(h =>
            h != null ? h.toString().trim() : ''
        );
        const dataRows  = raw
            .slice(headerIdx + 1)
            .filter(r => r.some(c => c !== '' && c != null));

        // Build { key → columnIndex } map using fuzzy matching
        const colMap = {};
        for (const { key, terms } of COLUMNS) {
            colMap[key] = findColIdx(headers, terms);
        }

        function get(row, key) {
            const idx = colMap[key];
            if (idx === -1 || idx == null) return '';
            return (row[idx] ?? '').toString().trim();
        }

        return dataRows.map(row => {
            const obj = { __source__: sourceName };
            for (const { key } of COLUMNS) obj[key] = get(row, key);
            return obj;
        });
    }

    /* ── Fetch all Google Sheets and combine ─────────────────── */
    /**
     * @param {Function} onProgress  (stepIndex, totalSteps, sheetName) => void
     * @returns {{ rows: object[], warnings: string[], sourceCounts: Map }}
     */
    async function fetchGoogleSheets(onProgress) {
        const urlList = AppData.getSheetUrls();   // [{ name, url }]

        if (!urlList.length) {
            return {
                rows:         [],
                warnings:     ['No Google Sheet URLs found in list.xlsx — nothing to fetch.'],
                sourceCounts: new Map(),
            };
        }

        const allRows     = [];
        const warnings    = [];
        const sourceCounts = new Map();   // name → row count

        for (let i = 0; i < urlList.length; i++) {
            const { name, url } = urlList[i];
            onProgress(i, urlList.length, name);

            const exportUrl = buildExportUrl(url);
            if (!exportUrl) {
                warnings.push(`"${name}": could not parse as a Google Sheets URL — skipped.`);
                continue;
            }

            try {
                const res = await fetch(exportUrl);
                if (!res.ok) throw new Error(`HTTP ${res.status}`);

                const csvText = await res.text();

                // Parse CSV with SheetJS
                const wb  = XLSX.read(csvText, { type: 'string', raw: false });
                const ws  = wb.Sheets[wb.SheetNames[0]];
                const raw = XLSX.utils.sheet_to_json(ws, {
                    header: 1, defval: '', raw: false,
                });

                const rows = parseSheetRows(raw, name);
                allRows.push(...rows);
                sourceCounts.set(name, rows.length);

            } catch (err) {
                warnings.push(`"${name}": fetch failed — ${err.message}.`);
                sourceCounts.set(name, 0);
            }
        }

        return { rows: allRows, warnings, sourceCounts };
    }

    /* ── Month / Half filtering ──────────────────────────────── */

    /**
     * Returns true when a cell value represents the given month.
     * Handles: numeric ("3", "03"), full name ("March"), 3-letter
     * abbreviation ("Mar"), and "March 2025"-style strings.
     */
    function matchesMonth(cellValue, monthVal, monthName) {
        const v = cellValue.toString().trim().toLowerCase();
        if (!v) return false;

        // Numeric: "3", "03", "3.0"
        const num = parseFloat(v);
        if (!isNaN(num) && Math.round(num) === monthVal) return true;

        const name = monthName.toLowerCase();

        // Exact full name: "march"
        if (v === name) return true;

        // 3-letter abbreviation: "mar"
        if (v === name.slice(0, 3)) return true;

        // Cell starts with month name: "march 2025", "march-25"
        if (v.startsWith(name)) return true;

        return false;
    }

    /**
     * Returns true when a cell value represents the given half.
     * Handles: "First" / "Second", "First Half" / "Second Half",
     * "1" / "2", "1st" / "2nd", "1st Half" / "2nd Half".
     */
    function matchesHalf(cellValue, half) {
        const v = cellValue.toString().trim().toLowerCase();
        if (!v) return false;

        if (half === 'first') {
            return v === 'first' || v === 'first half' ||
                   v === '1'     || v === '1st'        ||
                   v === '1st half' || v.startsWith('first');
        }
        return v === 'second' || v === 'second half' ||
               v === '2'      || v === '2nd'         ||
               v === '2nd half' || v.startsWith('second');
    }

    /**
     * Filter the unified dataset to only rows that match both the
     * selected month and the selected half.
     */
    function filterRows(rows, monthVal, monthName, half) {
        return rows.filter(row =>
            matchesMonth(row.month,     monthVal, monthName) &&
            matchesHalf( row.monthHalf, half)
        );
    }

    /* ── Month dropdown ──────────────────────────────────────── */
    function populateMonthDropdown() {
        const sel = $('allowanceMonth');
        const now = new Date();
        MONTHS.forEach((name, i) => {
            const opt = document.createElement('option');
            opt.value = i + 1;
            opt.textContent = name;
            if (i === now.getMonth()) opt.selected = true;
            sel.appendChild(opt);
        });
    }

    /* ── Master file rendering ───────────────────────────────── */
    function renderMasterFile() {
        const listEl   = $('allowanceMasterFileList');
        const dropZone = $('allowanceMasterDropZone');
        const runBtn   = $('allowanceRunBtn');

        if (!state.masterFile) {
            listEl.innerHTML = '<p class="no-files">No file uploaded yet</p>';
            dropZone.classList.remove('has-files');
            runBtn.disabled = true;
            return;
        }

        dropZone.classList.add('has-files');
        listEl.innerHTML = `
            <div class="file-item">
                <div class="file-item-name">
                    <span>📊</span>
                    <span class="fname" title="${esc(state.masterFile.name)}">${esc(state.masterFile.name)}</span>
                    <span class="file-status">✓</span>
                </div>
                <button class="file-remove" id="allowanceRemoveMasterBtn" title="Remove this file">✕</button>
            </div>
        `;

        $('allowanceRemoveMasterBtn').addEventListener('click', () => {
            state.masterFile = null;
            renderMasterFile();
        });

        runBtn.disabled = false;
    }

    /* ── Progress ────────────────────────────────────────────── */
    function setProgress(pct, text) {
        $('allowanceProgressBar').style.width = pct + '%';
        $('allowanceProgressText').textContent = text;
    }

    function showProgress() {
        $('allowanceProgressSection').hidden = false;
        setProgress(0, 'Starting…');
    }

    function hideProgress() {
        $('allowanceProgressSection').hidden = true;
    }

    /* ── Issues panel ────────────────────────────────────────── */
    function showIssues(errors, warnings) {
        const panel       = $('allowanceIssuesPanel');
        const errSection  = $('allowanceErrorsSection');
        const warnSection = $('allowanceWarningsSection');

        if (errors.length) {
            $('allowanceErrorsList').innerHTML = errors.map(e => `<li>${esc(e)}</li>`).join('');
            errSection.hidden = false;
        } else {
            errSection.hidden = true;
        }

        if (warnings.length) {
            $('allowanceWarningsList').innerHTML = warnings.map(w => `<li>${esc(w)}</li>`).join('');
            warnSection.hidden = false;
        } else {
            warnSection.hidden = true;
        }

        panel.hidden = !(errors.length || warnings.length);
    }

    function clearIssues() {
        $('allowanceIssuesPanel').hidden     = true;
        $('allowanceErrorsSection').hidden   = true;
        $('allowanceWarningsSection').hidden = true;
    }

    /* ── Results display ─────────────────────────────────────── */
    function showResults(monthName, halfLabel, sourceCounts, totalRows, filteredCount) {
        $('allowanceResultsSummary').textContent = `${monthName} — ${halfLabel}`;

        // Fetch summary table
        const tableRows = Array.from(sourceCounts.entries())
            .map(([name, count]) => `
                <tr>
                    <td>${esc(name)}</td>
                    <td class="allowance-count ${count === 0 ? 'allowance-count--zero' : ''}">${count}</td>
                </tr>
            `).join('');

        $('allowanceResultsBody').innerHTML = `
            <div class="allowance-fetch-summary">
                <div class="allowance-fetch-stat">
                    <span class="allowance-fetch-num">${sourceCounts.size}</span>
                    <span class="allowance-fetch-label">Sheets fetched</span>
                </div>
                <div class="allowance-fetch-stat">
                    <span class="allowance-fetch-num">${totalRows}</span>
                    <span class="allowance-fetch-label">Total rows loaded</span>
                </div>
                <div class="allowance-fetch-stat allowance-fetch-stat--filtered">
                    <span class="allowance-fetch-num">${filteredCount}</span>
                    <span class="allowance-fetch-label">${esc(monthName)} — ${esc(halfLabel)}</span>
                </div>
            </div>
            <div class="allowance-table-wrap">
                <table class="allowance-table">
                    <thead>
                        <tr><th>Coordinator / Sheet</th><th>Rows loaded</th></tr>
                    </thead>
                    <tbody>${tableRows || '<tr><td colspan="2" style="text-align:center;color:var(--gray-400);">No data</td></tr>'}</tbody>
                </table>
            </div>
        `;

        $('allowanceResultsSection').hidden = false;
        $('allowanceResultsSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    /* ── Run Analysis ────────────────────────────────────────── */
    async function runAnalysis() {
        if (!state.masterFile) return;

        const monthVal  = parseInt($('allowanceMonth').value, 10);
        const monthName = MONTHS[monthVal - 1];
        const half      = $('allowanceHalf').value;
        const halfLabel = half === 'first' ? 'First Half' : 'Second Half';

        $('allowanceResultsSection').hidden = true;
        clearIssues();
        showProgress();

        const warnings = [];
        const errors   = [];

        try {
            /* ── Step 1: Read master file ────────────────────── */
            setProgress(5, 'Reading master tracking file…');
            const masterSheets = await FileHandler.readFile(state.masterFile, undefined, 'ID#');

            /* ── Step 2: Fetch coordinator Google Sheets ─────── */
            setProgress(10, 'Fetching coordinator sheets…');

            const urlList  = AppData.getSheetUrls();
            const total    = urlList.length;

            const { rows, warnings: fetchWarnings, sourceCounts } =
                await fetchGoogleSheets((stepIdx, stepTotal, sheetName) => {
                    const pct = total > 0
                        ? 10 + Math.round((stepIdx / stepTotal) * 75)
                        : 10;
                    setProgress(pct, `Fetching ${stepIdx + 1} / ${stepTotal}: ${sheetName}…`);
                });

            warnings.push(...fetchWarnings);
            state.sheetRows = rows;

            /* ── Step 3: Filter by month & half ─────────────── */
            setProgress(88, `Filtering for ${monthName} — ${halfLabel}…`);

            const filteredRows = filterRows(rows, monthVal, monthName, half);
            state.filteredRows = filteredRows;

            if (rows.length > 0 && filteredRows.length === 0) {
                warnings.push(
                    `No rows matched "${monthName} — ${halfLabel}". ` +
                    `Verify that the Month and Month Half columns in your sheets ` +
                    `use a recognised format (e.g. "March" / "First Half").`
                );
            }

            /* ── Step 4: Analysis ────────────────────────────── */
            setProgress(94, 'Analysing…');

            // ── Analysis logic goes here ──────────────────────
            // Available:
            //   state.filteredRows — rows matching selected month + half
            //   state.sheetRows    — all combined rows (pre-filter)
            //   masterSheets       — [{ name, headers, rows, detectedHeaderRow }]
            //   monthVal           — 1-based month number (e.g. 3 = March)
            //   half               — 'first' | 'second'
            //   monthName, halfLabel — display strings
            //
            // Each row has these keys:
            //   __source__, month, day, monthHalf, coordinator, site, area,
            //   startTime, endTime, project, subProject, engineer,
            //   tech1, tech2, tech3, driver, allowance, vacationAllowance,
            //   workDetails, jc
            // ─────────────────────────────────────────────────

            setProgress(100, 'Done!');
            hideProgress();

            state.results = { monthVal, monthName, half, halfLabel, masterSheets };
            showResults(monthName, halfLabel, sourceCounts, rows.length, filteredRows.length);
            showIssues(errors, warnings);

        } catch (err) {
            hideProgress();
            errors.push(err.message);
            showIssues(errors, warnings);
        }
    }

    /* ── Reset ───────────────────────────────────────────────── */
    function reset() {
        state.masterFile   = null;
        state.sheetRows    = [];
        state.filteredRows = [];
        state.results      = null;

        renderMasterFile();
        $('allowanceMasterInput').value      = '';
        $('allowanceResultsSection').hidden  = true;
        $('allowanceProgressSection').hidden = true;
        clearIssues();
    }

    /* ── Init ────────────────────────────────────────────────── */
    function init() {
        populateMonthDropdown();

        FileHandler.setupDropZone(
            $('allowanceMasterDropZone'),
            $('allowanceMasterInput'),
            (files) => { state.masterFile = files[0]; renderMasterFile(); },
            false
        );

        $('allowanceRunBtn').addEventListener('click', runAnalysis);
        $('allowanceResetBtn').addEventListener('click', reset);
    }

    return { init };

})();
