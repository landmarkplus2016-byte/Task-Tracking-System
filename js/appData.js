/**
 * appData.js
 * ──────────────────────────────────────────────────────────────
 * Loads ./list.xlsx on startup and exposes its contents to
 * other modules.
 *
 * Reads two sheets:
 *   'Google Sheets URLs'  → [{ name, url }]
 *   'Salaries'            → two tables detected automatically:
 *                           team    → [{ name, dailySalary, bankAccount }]
 *                           drivers → [{ name, dailySalary, bankAccount }]
 *
 * If the file is missing or a tab is not found / unparseable,
 * an error banner (#appDataError) is shown in the UI.
 *
 * Load order: must come after fileHandler.js (uses FileHandler.detectHeaderRow)
 */

const AppData = (() => {
    'use strict';

    const DATA_FILE    = './list.xlsx';
    const URLS_TAB     = 'Google Sheets URLs';
    const SALARIES_TAB = 'Salaries';

    const state = {
        sheetUrls: [],              // [{ name, url }]
        salaries:  { team: [], drivers: [] },
    };

    /* ── Column fuzzy-match ──────────────────────────────────── */
    // Two passes: exact match against terms list, then contains match.
    // Terms are ordered most-specific first so shorter terms (e.g. "name")
    // don't shadow longer ones (e.g. "sheet name").
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

    /* ── Raw sheet → { headers, rows } ──────────────────────── */
    function parseRawSheet(ws) {
        const raw = XLSX.utils.sheet_to_json(ws, {
            header: 1, defval: '', raw: false, dateNF: 'yyyy-mm-dd',
        });
        if (!raw.length) return { headers: [], rows: [] };

        const headerIdx = FileHandler.detectHeaderRow(raw, 40);
        const headers   = (raw[headerIdx] || []).map(h =>
            h != null ? h.toString().trim() : ''
        );
        const rows = raw
            .slice(headerIdx + 1)
            .filter(r => r.some(c => c !== '' && c != null));

        return { headers, rows };
    }

    /* ── Tab parsers ─────────────────────────────────────────── */
    function parseUrlsTab(ws) {
        const { headers, rows } = parseRawSheet(ws);

        const nameIdx = findColIdx(headers, ['sheet name', 'name', 'title']);
        const urlIdx  = findColIdx(headers, ['url', 'sheet url', 'link', 'google sheet']);

        const missing = [];
        if (nameIdx === -1) missing.push('Name');
        if (urlIdx  === -1) missing.push('URL');
        if (missing.length) {
            throw new Error(
                `'${URLS_TAB}' tab: could not locate column(s): ${missing.join(', ')}.`
            );
        }

        return rows
            .filter(r => r[nameIdx] || r[urlIdx])
            .map(r => ({
                name: (r[nameIdx] || '').toString().trim(),
                url:  (r[urlIdx]  || '').toString().trim(),
            }))
            .filter(r => r.name || r.url);
    }

    /**
     * Detect whether a raw row looks like a table header.
     * Uses EXACT cell value matches only — data rows (person names, numbers,
     * bank account codes) will never exactly equal these header keywords.
     */
    const HEADER_NAME_TERMS = new Set([
        'name', 'member name', 'member', 'employee', 'employee name',
        'driver name', 'driver',
    ]);
    const HEADER_SALARY_TERMS = new Set([
        'daily salary', 'salary', 'daily rate', 'rate',
        'bank account', 'account number', 'account no', 'account', 'bank',
    ]);
    function isHeaderCandidate(row) {
        const cells = row.map(c => (c || '').toString().toLowerCase().trim());
        return cells.some(c => HEADER_NAME_TERMS.has(c)) &&
               cells.some(c => HEADER_SALARY_TERMS.has(c));
    }

    /**
     * Parse one contiguous table block from `raw` (a 2-D array).
     * headerIdx = index of the header row in `raw`.
     * endIdx    = index where this table ends (exclusive), or undefined for EOF.
     * bankAccount column is optional (bankIdx may be -1 → stored as '').
     */
    function parseSalaryTable(raw, headerIdx, endIdx) {
        const headers   = (raw[headerIdx] || []).map(h => (h || '').toString().trim());
        const nameIdx   = findColIdx(headers, ['name', 'member name', 'member', 'employee']);
        // 'salary/day' and 'salary per day' must come before 'salary' so the
        // daily-rate column is preferred over the monthly-salary column.
        const salaryIdx = findColIdx(headers, ['daily salary', 'salary/day', 'salary per day', 'daily rate', 'salary', 'rate']);
        const bankIdx   = findColIdx(headers, ['bank account', 'account number', 'account no', 'account', 'bank']);

        if (nameIdx === -1 || salaryIdx === -1) return [];

        return raw
            .slice(headerIdx + 1, endIdx)
            .filter(r => r.some(c => c !== '' && c != null))
            .map(r => ({
                name:        (r[nameIdx]   || '').toString().trim(),
                dailySalary: (r[salaryIdx] || '').toString().trim(),
                bankAccount: bankIdx !== -1 ? (r[bankIdx] || '').toString().trim() : '',
            }))
            .filter(r => r.name);
    }

    const DRIVER_COL_TERMS = new Set(['driver name', 'driver']);

    function parseSalariesTab(ws) {
        const raw = XLSX.utils.sheet_to_json(ws, {
            header: 1, defval: '', raw: false, dateNF: 'yyyy-mm-dd',
        });
        if (!raw.length) return { team: [], drivers: [] };

        // Find the main header row (must have both a name col and a salary/account col)
        const teamHeaderIdx = raw.findIndex(isHeaderCandidate);
        if (teamHeaderIdx === -1) {
            throw new Error(
                `'${SALARIES_TAB}' tab: could not detect header row. ` +
                `The team table must have a "Name" column and a "Salary"/"Account"/"Rate" column.`
            );
        }

        const team = parseSalaryTable(raw, teamHeaderIdx, undefined);

        // Detect driver column within the SAME header row (side-by-side layout).
        // Column G only has "Driver Name" with no salary column, so isHeaderCandidate
        // won't fire for it — instead we scan the header row cells directly.
        const headers = (raw[teamHeaderIdx] || []).map(h => (h || '').toString().toLowerCase().trim());
        const driverColIdx = headers.findIndex(h => DRIVER_COL_TERMS.has(h));

        let drivers = [];
        if (driverColIdx !== -1) {
            drivers = raw
                .slice(teamHeaderIdx + 1)
                .filter(r => r.some(c => c !== '' && c != null))
                .map(r => (r[driverColIdx] || '').toString().trim())
                .filter(name => name)
                .map(name => ({ name, dailySalary: '', bankAccount: '' }));
        } else {
            // Fallback: look for a second vertical header block below team table
            const secondHeaderIdx = raw.findIndex((row, i) => i > teamHeaderIdx && isHeaderCandidate(row));
            if (secondHeaderIdx !== -1) {
                drivers = parseSalaryTable(raw, secondHeaderIdx, undefined);
            }
        }

        return { team, drivers };
    }

    /* ── Error banner ────────────────────────────────────────── */
    function showError(messages) {
        const el = document.getElementById('appDataError');
        if (!el) return;
        el.innerHTML = messages
            .map(m => `<p>${m.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</p>`)
            .join('');
        el.hidden = false;
    }

    /* ── Public init ─────────────────────────────────────────── */
    async function init() {
        // fetch() is blocked by the browser when the app is opened directly
        // from the filesystem (file:// protocol). Skip silently — the file
        // will load correctly when served via GitHub Pages or a local server.
        if (window.location.protocol === 'file:') return;

        let workbook;

        try {
            const res = await fetch(DATA_FILE);
            if (!res.ok) throw new Error(`HTTP ${res.status} — file not found`);
            const buf = await res.arrayBuffer();
            workbook = XLSX.read(new Uint8Array(buf), {
                type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd',
            });
        } catch (err) {
            showError([`Could not load list.xlsx: ${err.message}`]);
            return;
        }

        const errors = [];

        // 'Google Sheets URLs' tab
        const urlsWs = workbook.Sheets[URLS_TAB];
        if (!urlsWs) {
            errors.push(`Tab '${URLS_TAB}' not found in list.xlsx.`);
        } else {
            try {
                state.sheetUrls = parseUrlsTab(urlsWs);
            } catch (err) {
                errors.push(err.message);
            }
        }

        // 'Salaries' tab
        const salariesWs = workbook.Sheets[SALARIES_TAB];
        if (!salariesWs) {
            errors.push(`Tab '${SALARIES_TAB}' not found in list.xlsx.`);
        } else {
            try {
                const parsed = parseSalariesTab(salariesWs);
                state.salaries = parsed;
                console.log(
                    `Salaries loaded — team: ${parsed.team.length}, drivers: ${parsed.drivers.length}`
                );
            } catch (err) {
                errors.push(err.message);
            }
        }

        if (errors.length) showError(errors);
    }

    /* ── Public API ──────────────────────────────────────────── */
    return {
        init,
        getSheetUrls:      () => state.sheetUrls,
        getSalaries:       () => state.salaries.team,     // [{ name, dailySalary, bankAccount }]
        getDriverSalaries: () => state.salaries.drivers,  // [{ name, dailySalary, bankAccount }]
    };

})();
