/**
 * siteIdJc.js
 * ──────────────────────────────────────────────────────────────
 * Self-contained module for the "Site ID-JC File" tab.
 *
 * Accepts multiple tracking files (PC Tracking and/or POC Tracking).
 * For each file it locates the relevant sheet tab and extracts rows,
 * then combines everything into a single 4-column output:
 *
 *   A  Site ID-JC   – Site ID + "-" + Job Code  (e.g. "D1234-MK001")
 *   B  Task Date    – date value from the sheet
 *   C  Old/New      – date < 2026-01-01 → "Old", else → "New"
 *   D  Contractor   – contractor / installation team
 *
 * Supported sheet names (searched in order):
 *   "Invoicing Track"  – PC Tracking files
 *   "POC3 Tracking"    – POC Tracking files
 *
 * Column name variants handled automatically:
 *   Site ID      : "Physical Site ID" (PC) | "Site ID" (POC)
 *   Job Code     : "Job Code" (both)
 *   Task Date    : "Task Date" (PC) | "Installation Date" (POC)
 *   Contractor   : "Contractor" (PC) | "Installation Team" (POC)
 *
 * Dependencies: FileHandler (fileHandler.js), XLSX (SheetJS CDN)
 */

const SiteIdJc = (() => {
    'use strict';

    /* ── Configuration ────────────────────────────────────────── */
    // Sheet names to search for, in priority order.
    const SHEET_NAMES = ['Invoicing Track', 'POC3 Tracking'];
    const OLD_CUTOFF  = '2026-01-01';

    /* ── Column detection rules ───────────────────────────────── */
    // Listed most-specific first so exact match on the longer term wins.
    const COL_RULES = [
        // PC: "Physical Site ID"  |  POC: "Site ID"
        { key: 'siteId',     terms: ['physical site id', 'physical site_id', 'site id', 'site_id', 'siteid'] },
        // Both: "Job Code"
        { key: 'jobCode',    terms: ['job code', 'job_code', 'jobcode', 'jc#', 'jc'] },
        // PC: "Task Date"  |  POC: "Installation Date"
        { key: 'taskDate',   terms: ['installation date', 'task date', 'task_date', 'install date'] },
        // PC: "Contractor"  |  POC: "Installation Team"
        { key: 'contractor', terms: ['installation team', 'contractor', 'sub-contractor', 'subcontractor'] },
    ];

    /* ── Module-private state ─────────────────────────────────── */
    let _files    = [];     // File[] — currently loaded files
    let _workbook = null;   // held for the download button

    /* ── Helpers ──────────────────────────────────────────────── */
    const $       = id => document.getElementById(id);
    const escHtml = s  => String(s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    /* ── Fuzzy column finder ──────────────────────────────────── */
    function detectColumns(headers) {
        const result = {};
        const lower  = headers.map(h => (h || '').toString().toLowerCase().trim());

        for (const { key, terms } of COL_RULES) {
            // Pass 1: exact match (header equals one of the terms exactly)
            let idx = lower.findIndex(h => terms.includes(h));
            // Pass 2: contains match
            if (idx === -1) {
                idx = lower.findIndex(h => terms.some(t => h.includes(t)));
            }
            result[key] = idx;
        }
        return result;   // { siteId: N, jobCode: N, taskDate: N, contractor: N }
    }

    /* ── Sheet finder ─────────────────────────────────────────── */
    /**
     * Search the parsed sheets array for the first recognised sheet name.
     * Returns { sheet, sheetType } or { sheet: null, sheetType: null }.
     */
    function findTargetSheet(sheets) {
        for (const name of SHEET_NAMES) {
            const target = name.toLowerCase().trim();
            const found  =
                sheets.find(s => s.name.toLowerCase().trim() === target) ||
                sheets.find(s => s.name.toLowerCase().includes(target));
            if (found) return { sheet: found, sheetType: name };
        }
        return { sheet: null, sheetType: null };
    }

    /* ── Date parser ──────────────────────────────────────────── */
    const MONTH_MAP = {
        jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
        jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
    };

    function parseDate(dateStr) {
        const s = (dateStr || '').toString().trim();
        if (!s) return null;

        // ISO: yyyy-mm-dd
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s);

        // dd-Mon-yy or dd-Mon-yyyy  (e.g. 29-Apr-25 / 29-Apr-2025)
        const dmy = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
        if (dmy) {
            const day   = parseInt(dmy[1], 10);
            const month = MONTH_MAP[dmy[2].toLowerCase()];
            let   year  = parseInt(dmy[3], 10);
            if (year < 100) year += 2000;
            if (month === undefined) return null;
            return new Date(year, month, day);
        }

        // dd/mm/yyyy or mm/dd/yyyy
        const slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (slash) {
            const a = parseInt(slash[1], 10);
            const b = parseInt(slash[2], 10);
            const y = parseInt(slash[3], 10);
            if (a > 12) return new Date(y, b - 1, a);
            return new Date(y, a - 1, b);
        }

        const d = new Date(s);
        return isNaN(d.getTime()) ? null : d;
    }

    /* ── Date classifier ──────────────────────────────────────── */
    const CUTOFF = new Date(OLD_CUTOFF);

    function classifyDate(dateStr) {
        const d = parseDate(dateStr);
        if (!d) return '';
        return d < CUTOFF ? 'Old' : 'New';
    }

    /* ── Cell styles ──────────────────────────────────────────── */
    const HEADER_STYLE = {
        fill: { fgColor: { rgb: '0070C0' } },
        font: { color: { rgb: 'FFFFFF' }, bold: true },
        border: {
            top:    { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left:   { style: 'thin', color: { rgb: '000000' } },
            right:  { style: 'thin', color: { rgb: '000000' } },
        },
    };

    const DATA_STYLE = {
        border: {
            top:    { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left:   { style: 'thin', color: { rgb: '000000' } },
            right:  { style: 'thin', color: { rgb: '000000' } },
        },
    };

    /* ── Excel workbook builder ───────────────────────────────── */
    function buildWorkbook(dataRows) {
        const headers = ['Site ID-JC', 'Task Date', 'Old/New', 'Contractor'];
        const ws      = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);

        const widths = headers.map(h => h.length);
        dataRows.forEach(row => {
            row.forEach((cell, i) => {
                if (cell != null) {
                    widths[i] = Math.min(Math.max(widths[i], String(cell).length), 60);
                }
            });
        });
        ws['!cols'] = widths.map(w => ({ wch: w + 2 }));

        // Apply header colour and border styles
        for (let c = 0; c < headers.length; c++) {
            const ref = XLSX.utils.encode_cell({ r: 0, c });
            if (ws[ref]) ws[ref].s = HEADER_STYLE;
        }
        for (let r = 1; r <= dataRows.length; r++) {
            for (let c = 0; c < headers.length; c++) {
                const ref = XLSX.utils.encode_cell({ r, c });
                if (ws[ref]) ws[ref].s = DATA_STYLE;
            }
        }

        ws['!freeze'] = { xSplit: 0, ySplit: 1 };

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Site ID-JC');
        return wb;
    }

    /* ── Progress helpers ─────────────────────────────────────── */
    function setProgress(pct, text) {
        $('siteIdProgressBar').style.width  = pct + '%';
        $('siteIdProgressText').textContent = text;
    }
    function showProgress() { $('siteIdProgressSection').hidden = false; setProgress(0, 'Starting…'); }
    function hideProgress() { $('siteIdProgressSection').hidden = true; }

    function flashCardBar() {
        const prog = $('siteIdProgress');
        const bar  = $('siteIdBar');
        prog.hidden     = false;
        bar.style.width = '0%';
        requestAnimationFrame(() => requestAnimationFrame(() => {
            bar.style.width = '100%';
        }));
        setTimeout(() => { bar.style.width = '0%'; prog.hidden = true; }, 800);
    }

    /* ── File list renderer ───────────────────────────────────── */
    function renderFiles() {
        const fileList = $('siteIdFileList');
        const dropZone = $('siteIdDropZone');

        if (_files.length === 0) {
            fileList.innerHTML = '<p class="no-files">No files uploaded yet</p>';
            dropZone.classList.remove('has-files');
            return;
        }

        dropZone.classList.add('has-files');
        fileList.innerHTML = _files.map((f, i) => `
            <div class="file-item">
                <div class="file-item-name">
                    <span>📄</span>
                    <span class="fname" title="${escHtml(f.name)}">${escHtml(f.name)}</span>
                    <span class="file-status">✓</span>
                </div>
                <button class="file-remove" data-index="${i}" title="Remove this file">✕</button>
            </div>
        `).join('');

        fileList.querySelectorAll('.file-remove').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(e.currentTarget.dataset.index, 10);
                _files.splice(idx, 1);
                renderFiles();
                if (_files.length === 0) {
                    $('siteIdResultsSection').hidden = true;
                    hideProgress();
                    _workbook = null;
                } else {
                    processAll();
                }
            });
        });
    }

    /* ── Main processing function ─────────────────────────────── */
    async function processAll() {
        if (_files.length === 0) return;

        $('siteIdResultsSection').hidden = true;
        _workbook = null;
        showProgress();

        const allRows    = [];
        const skipped    = [];

        try {
            for (let i = 0; i < _files.length; i++) {
                const file = _files[i];
                const pct  = Math.round(10 + 80 * i / _files.length);
                setProgress(pct, `Reading ${file.name}… (${i + 1} / ${_files.length})`);

                try {
                    // Use "job code" as the hint — present in both file types
                    const sheets = await FileHandler.readFile(file, undefined, 'job code');
                    const { sheet, sheetType } = findTargetSheet(sheets);

                    if (!sheet) {
                        const names = sheets.map(s => `"${s.name}"`).join(', ');
                        skipped.push(`"${file.name}" — no recognised sheet found (available: ${names})`);
                        continue;
                    }

                    if (!sheet.headers || sheet.headers.length === 0) {
                        skipped.push(`"${file.name}" — sheet "${sheetType}" is empty`);
                        continue;
                    }

                    const cols = detectColumns(sheet.headers);

                    // Validate required columns
                    const missing = Object.entries(cols)
                        .filter(([, idx]) => idx === -1)
                        .map(([key]) => key);

                    if (missing.length > 0) {
                        skipped.push(
                            `"${file.name}" (${sheetType}) — could not find columns: ${missing.join(', ')}. ` +
                            `Headers: ${sheet.headers.filter(Boolean).join(' | ')}`
                        );
                        continue;
                    }

                    const rows = sheet.rows
                        .map(row => {
                            const siteId     = String(row[cols.siteId]     || '').trim();
                            const jobCode    = String(row[cols.jobCode]    || '').trim();
                            const taskDate   = String(row[cols.taskDate]   || '').trim();
                            const contractor = String(row[cols.contractor] || '').trim();

                            const siteIdJc = (siteId && jobCode)
                                ? `${siteId}-${jobCode}`
                                : (siteId || jobCode);

                            return [siteIdJc, taskDate, classifyDate(taskDate), contractor];
                        })
                        .filter(row => row[0] || row[1]);

                    allRows.push(...rows);
                    console.log(`✓ "${file.name}" [${sheetType}] — ${rows.length} rows`);

                } catch (err) {
                    skipped.push(`"${file.name}" — ${err.message}`);
                }
            }

            if (allRows.length === 0 && skipped.length > 0) {
                throw new Error('No data could be extracted from any file:\n\n' + skipped.join('\n'));
            }

            setProgress(95, 'Generating Excel workbook…');
            _workbook = buildWorkbook(allRows);

            setProgress(100, 'Done!');
            $('siteIdRowCount').textContent  = allRows.length;
            $('siteIdResultsSection').hidden = false;
            $('siteIdResultsSection').scrollIntoView({ behavior: 'smooth', block: 'start' });

            if (skipped.length > 0) {
                console.warn('Skipped files:\n' + skipped.join('\n'));
                alert(`⚠ ${skipped.length} file(s) could not be processed:\n\n${skipped.join('\n')}`);
            }

        } catch (err) {
            console.error(err);
            hideProgress();
            alert(`Site ID-JC processing failed:\n\n${err.message}`);
        }
    }

    /* ── Public init — called once by app.js ──────────────────── */
    function init() {
        FileHandler.setupDropZone(
            $('siteIdDropZone'),
            $('siteIdInput'),
            (files) => {
                files.forEach(f => {
                    if (!_files.find(e => e.name === f.name && e.size === f.size)) {
                        _files.push(f);
                    }
                });
                renderFiles();
                flashCardBar();
                processAll();
            },
            true   // multiple files
        );

        $('siteIdDownloadBtn').addEventListener('click', () => {
            if (!_workbook) return;
            const now      = new Date();
            const datePart = [
                now.getFullYear(),
                String(now.getMonth() + 1).padStart(2, '0'),
                String(now.getDate()).padStart(2, '0'),
            ].join('');
            XLSX.writeFile(_workbook, `SiteID_JC_${datePart}.xlsx`);
        });
    }

    /* ── Reset ────────────────────────────────────────────────── */
    function reset() {
        _files    = [];
        _workbook = null;
        renderFiles();
        $('siteIdInput').value            = '';
        $('siteIdResultsSection').hidden  = true;
        $('siteIdProgressSection').hidden = true;
    }

    /* ── Public API ───────────────────────────────────────────── */
    return { init, reset };

})();
