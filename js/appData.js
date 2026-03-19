/**
 * appData.js
 * ──────────────────────────────────────────────────────────────
 * Loads ./data/list.xlsx on startup and exposes its contents to
 * other modules.
 *
 * Reads two sheets:
 *   'Google Sheets URLs'  → [{ name, url }]
 *   'Salaries'            → [{ name, dailySalary, bankAccount }]
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
        sheetUrls: [],   // [{ name, url }]
        salaries:  [],   // [{ name, dailySalary, bankAccount }]
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

    function parseSalariesTab(ws) {
        const { headers, rows } = parseRawSheet(ws);

        const nameIdx   = findColIdx(headers, ['name', 'member name', 'member', 'employee']);
        const salaryIdx = findColIdx(headers, ['daily salary', 'daily rate', 'salary', 'rate']);
        const bankIdx   = findColIdx(headers, ['bank account', 'account number', 'account no', 'account', 'bank']);

        const missing = [];
        if (nameIdx   === -1) missing.push('Name');
        if (salaryIdx === -1) missing.push('Daily Salary');
        if (bankIdx   === -1) missing.push('Bank Account');
        if (missing.length) {
            throw new Error(
                `'${SALARIES_TAB}' tab: could not locate column(s): ${missing.join(', ')}.`
            );
        }

        return rows
            .filter(r => r[nameIdx])
            .map(r => ({
                name:        (r[nameIdx]   || '').toString().trim(),
                dailySalary: (r[salaryIdx] || '').toString().trim(),
                bankAccount: (r[bankIdx]   || '').toString().trim(),
            }))
            .filter(r => r.name);
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
                state.salaries = parseSalariesTab(salariesWs);
            } catch (err) {
                errors.push(err.message);
            }
        }

        if (errors.length) showError(errors);
    }

    /* ── Public API ──────────────────────────────────────────── */
    return {
        init,
        getSheetUrls: () => state.sheetUrls,
        getSalaries:  () => state.salaries,
    };

})();
