/**
 * siteIdJc.js
 * ──────────────────────────────────────────────────────────────
 * Self-contained module for the "Site ID-JC File" tab.
 *
 * Reads the "Invoicing Track" sheet from an uploaded master file,
 * builds a 4-column output:
 *   A  Site ID-JC   – Site ID + "-" + Job Code  (e.g. "D1234-MK001")
 *   B  Task Date    – as-is from the sheet
 *   C  Old/New      – date < 2026-01-01 → "Old", else → "New"
 *   D  Contractor   – as-is from the sheet
 *
 * Dependencies: FileHandler (fileHandler.js), XLSX (SheetJS CDN)
 */

const SiteIdJc = (() => {
    'use strict';

    /* ── Configuration ────────────────────────────────────────── */
    const MASTER_SHEET = 'Invoicing Track';
    const OLD_CUTOFF   = '2026-01-01';   // ISO string — safe to compare as string

    /* ── Column detection rules ───────────────────────────────── */
    // Listed from most-specific to least-specific so the first match wins.
    const COL_RULES = [
        { key: 'siteId',     terms: ['site id', 'site_id', 'siteid'] },
        { key: 'jobCode',    terms: ['job code', 'job_code', 'jobcode', 'jc#', 'jc'] },
        { key: 'taskDate',   terms: ['task date', 'task_date'] },
        { key: 'contractor', terms: ['contractor', 'sub-contractor', 'subcontractor'] },
    ];

    /* ── Module-private state ─────────────────────────────────── */
    let _workbook = null;   // held for the download button

    /* ── Helpers ──────────────────────────────────────────────── */
    const $        = id => document.getElementById(id);
    const escHtml  = s  => String(s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    /* ── Fuzzy column finder ──────────────────────────────────── */
    function detectColumns(headers) {
        const result = {};
        const lower  = headers.map(h => (h || '').toString().toLowerCase().trim());

        for (const { key, terms } of COL_RULES) {
            // Pass 1: exact match (any term equals the full header)
            let idx = lower.findIndex(h => terms.includes(h));
            // Pass 2: contains match
            if (idx === -1) {
                idx = lower.findIndex(h => terms.some(t => h.includes(t)));
            }
            result[key] = idx;
        }
        return result;   // { siteId: N, jobCode: N, taskDate: N, contractor: N }
    }

    /* ── Date parser ──────────────────────────────────────────── */
    const MONTH_MAP = {
        jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
        jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
    };

    /**
     * Parse a date string in any of these common formats into a Date:
     *   yyyy-mm-dd   (ISO — SheetJS default)
     *   dd-Mon-yy    e.g. 29-Apr-25
     *   dd-Mon-yyyy  e.g. 29-Apr-2025
     *   dd/mm/yyyy
     *   mm/dd/yyyy
     * Returns null if unparseable.
     */
    function parseDate(dateStr) {
        const s = (dateStr || '').toString().trim();
        if (!s) return null;

        // ISO: yyyy-mm-dd
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
            return new Date(s);
        }

        // dd-Mon-yy or dd-Mon-yyyy  (e.g. 29-Apr-25 / 29-Apr-2025)
        const dmy = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
        if (dmy) {
            const day   = parseInt(dmy[1], 10);
            const month = MONTH_MAP[dmy[2].toLowerCase()];
            let   year  = parseInt(dmy[3], 10);
            if (year < 100) year += 2000;   // 25 → 2025
            if (month === undefined) return null;
            return new Date(year, month, day);
        }

        // dd/mm/yyyy or mm/dd/yyyy — try both; pick whichever is a valid date
        const slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (slash) {
            const a = parseInt(slash[1], 10);
            const b = parseInt(slash[2], 10);
            const y = parseInt(slash[3], 10);
            // If first part > 12 it must be dd/mm/yyyy
            if (a > 12) return new Date(y, b - 1, a);
            return new Date(y, a - 1, b);   // assume mm/dd/yyyy
        }

        // Last resort: let the JS engine try
        const d = new Date(s);
        return isNaN(d.getTime()) ? null : d;
    }

    /* ── Date classifier ──────────────────────────────────────── */
    const CUTOFF = new Date(OLD_CUTOFF);   // 2026-01-01

    function classifyDate(dateStr) {
        const d = parseDate(dateStr);
        if (!d) return '';
        return d < CUTOFF ? 'Old' : 'New';
    }

    /* ── Sheet finder ─────────────────────────────────────────── */
    function findInvoicingSheet(sheets) {
        const target = MASTER_SHEET.toLowerCase().trim();
        return (
            sheets.find(s => s.name.toLowerCase().trim() === target) ||
            sheets.find(s => s.name.toLowerCase().includes(target))  ||
            null
        );
    }

    /* ── Excel workbook builder ───────────────────────────────── */
    function buildWorkbook(dataRows) {
        const headers = ['Site ID-JC', 'Task Date', 'Old/New', 'Contractor'];
        const ws      = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);

        // Auto-fit column widths
        const widths = headers.map(h => h.length);
        dataRows.forEach(row => {
            row.forEach((cell, i) => {
                if (cell != null) {
                    widths[i] = Math.min(Math.max(widths[i], String(cell).length), 60);
                }
            });
        });
        ws['!cols'] = widths.map(w => ({ wch: w + 2 }));

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
        prog.hidden      = false;
        bar.style.width  = '0%';
        requestAnimationFrame(() => requestAnimationFrame(() => {
            bar.style.width = '100%';
        }));
        setTimeout(() => { bar.style.width = '0%'; prog.hidden = true; }, 800);
    }

    /* ── File list renderer ───────────────────────────────────── */
    function renderFile(file) {
        const fileList = $('siteIdFileList');
        const dropZone = $('siteIdDropZone');

        if (!file) {
            fileList.innerHTML = '<p class="no-files">No file uploaded yet</p>';
            dropZone.classList.remove('has-files');
            return;
        }

        dropZone.classList.add('has-files');
        fileList.innerHTML = `
            <div class="file-item">
                <div class="file-item-name">
                    <span>📊</span>
                    <span class="fname" title="${escHtml(file.name)}">${escHtml(file.name)}</span>
                    <span class="file-status">✓</span>
                </div>
                <button class="file-remove" id="siteIdRemoveBtn" title="Remove this file">✕</button>
            </div>
        `;

        $('siteIdRemoveBtn').addEventListener('click', () => {
            renderFile(null);
            $('siteIdResultsSection').hidden = true;
            hideProgress();
            _workbook = null;
        });
    }

    /* ── Main processing function ─────────────────────────────── */
    async function process(file) {
        $('siteIdResultsSection').hidden = true;
        _workbook = null;
        showProgress();

        try {
            setProgress(20, 'Reading file…');
            const sheets = await FileHandler.readFile(file, undefined, 'ID#');

            setProgress(45, 'Locating "Invoicing Track" sheet…');
            const sheet = findInvoicingSheet(sheets);

            if (!sheet) {
                throw new Error(
                    `Sheet "${MASTER_SHEET}" not found in this file.\n` +
                    `Available sheets: ${sheets.map(s => `"${s.name}"`).join(', ')}`
                );
            }
            if (!sheet.headers || sheet.headers.length === 0) {
                throw new Error(`Sheet "${MASTER_SHEET}" appears to be empty.`);
            }

            setProgress(60, 'Detecting columns…');
            const cols = detectColumns(sheet.headers);

            const missing = Object.entries(cols)
                .filter(([, idx]) => idx === -1)
                .map(([key]) => key);

            if (missing.length > 0) {
                throw new Error(
                    `Could not locate required columns: ${missing.join(', ')}.\n` +
                    `Headers detected: ${sheet.headers.filter(Boolean).join(' | ')}`
                );
            }

            setProgress(75, 'Building output rows…');
            const outputRows = sheet.rows
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
                .filter(row => row[0] || row[1]);   // drop fully blank rows

            setProgress(92, 'Generating Excel workbook…');
            _workbook = buildWorkbook(outputRows);

            setProgress(100, 'Done!');
            $('siteIdRowCount').textContent   = outputRows.length;
            $('siteIdResultsSection').hidden  = false;
            $('siteIdResultsSection').scrollIntoView({ behavior: 'smooth', block: 'start' });

        } catch (err) {
            console.error(err);
            hideProgress();
            alert(`Site ID-JC processing failed:\n\n${err.message}`);
        }
    }

    /* ── Public init — called once by app.js ──────────────────── */
    function init() {
        // Wire drop zone
        FileHandler.setupDropZone(
            $('siteIdDropZone'),
            $('siteIdInput'),
            (files) => {
                const file = files[0];
                renderFile(file);
                flashCardBar();
                process(file);
            },
            false   // single file only
        );

        // Download button
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

    /* ── Public API ───────────────────────────────────────────── */
    return { init };

})();
