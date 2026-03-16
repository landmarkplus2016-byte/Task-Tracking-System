/**
 * fileHandler.js
 * ──────────────────────────────────────────────────────────────
 * Handles file upload, drag-and-drop interaction, and spreadsheet
 * parsing via the SheetJS (XLSX) library.
 *
 * Accepts ANY file type — SheetJS will handle what it can and
 * throw a meaningful error for anything it cannot parse.
 *
 * Header row auto-detection: scans up to the first 40 rows to
 * find the row that looks most like a table header (multiple
 * text-label cells), regardless of how many blank or title rows
 * appear above it.
 */

const FileHandler = (() => {

    /* ── Header Row Auto-detection ────────────────────────────── */

    /**
     * Returns true when a cell value looks like a column label
     * (contains at least one letter, is not a plain number, is not
     * a recognisable date string).
     */
    function isTextLabel(val) {
        if (val === null || val === undefined) return false;
        const s = val.toString().trim();
        if (!s) return false;
        if (!isNaN(Number(s))) return false;
        // Common date patterns  MM/DD/YYYY  DD-MM-YYYY  YYYY-MM-DD
        if (/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}$/.test(s)) return false;
        if (/^\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}$/.test(s)) return false;
        return /[a-zA-Z]/.test(s);
    }

    /**
     * Scan rawData and return the 0-based index of the header row.
     *
     * Detection strategy (in priority order):
     *
     *  1. ID-hint scan — if idHint is supplied, look for the first row
     *     that contains a cell whose value exactly matches idHint
     *     (case-insensitive).  This is the most reliable signal because
     *     real data rows never contain the column name as a cell value.
     *
     *  2. Density scan — find the first row that reaches ≥ 70 % of the
     *     maximum non-empty-cell count seen in the first N rows AND has
     *     ≥ 70 % of those cells as text labels.  The header row almost
     *     always spans the full width of the table, while title / metadata
     *     rows above it use only a few cells.
     *
     *  3. Absolute fallback — first row with any non-empty content.
     *
     * @param  {any[][]} rawData
     * @param  {number}  [maxScan=40]
     * @param  {string}  [idHint]     User-supplied ID column name
     * @returns {number} 0-based row index
     */
    function detectHeaderRow(rawData, maxScan = 40, idHint = '') {
        const limit = Math.min(rawData.length, maxScan);
        if (limit === 0) return 0;

        // ── Strategy 1: look for the row that contains the ID column name ──
        // Data rows will never have the column name as a cell value, so this
        // is unambiguous even when the sheet has many leading metadata rows.
        if (idHint) {
            const target = idHint.toLowerCase().trim();
            for (let i = 0; i < limit; i++) {
                const row = rawData[i] || [];
                if (row.some(c =>
                    c !== null && c !== undefined &&
                    c.toString().trim().toLowerCase() === target
                )) return i;
            }
            // Partial fallback: cell contains the hint (handles "Task ID#" etc.)
            for (let i = 0; i < limit; i++) {
                const row = rawData[i] || [];
                if (row.some(c => {
                    if (c === null || c === undefined) return false;
                    const v = c.toString().trim().toLowerCase();
                    return v.includes(target) || target.includes(v);
                })) return i;
            }
        }

        // ── Strategy 2: density-based detection ───────────────────────────
        // Count non-empty cells per row.
        const counts = [];
        for (let i = 0; i < limit; i++) {
            const row = rawData[i] || [];
            counts.push(
                row.filter(c =>
                    c !== null && c !== undefined && c.toString().trim() !== ''
                ).length
            );
        }

        const maxCount = Math.max(...counts, 0);
        if (maxCount < 1) return 0;

        // The header row is the first row that:
        //   a) reaches ≥ 70 % of the max column count  (dense row)
        //   b) has ≥ 70 % of its non-empty cells as text labels
        const densityThreshold = Math.max(2, Math.ceil(maxCount * 0.7));

        for (let i = 0; i < limit; i++) {
            if (counts[i] < densityThreshold) continue;

            const row = rawData[i] || [];
            let nonEmpty = 0, textLabels = 0;
            for (const cell of row) {
                const v = (cell !== null && cell !== undefined)
                    ? cell.toString().trim() : '';
                if (!v) continue;
                nonEmpty++;
                if (isTextLabel(v)) textLabels++;
            }

            if (nonEmpty > 0 && textLabels / nonEmpty >= 0.7) return i;
        }

        // ── Strategy 3: absolute fallback ─────────────────────────────────
        const firstNonEmpty = counts.findIndex(c => c > 0);
        return firstNonEmpty !== -1 ? firstNonEmpty : 0;
    }

    /* ── Spreadsheet Reader ───────────────────────────────────── */

    /**
     * Read any File object that SheetJS can parse and return an
     * array of sheet objects.  The header row is located automatically
     * using detectHeaderRow() unless forceHeaderRow is supplied.
     *
     * @param  {File}   file
     * @param  {number} [forceHeaderRow]  0-based row index override
     * @param  {string} [idColumnHint]    ID column name to anchor detection
     * @returns {Promise<Array<{name:string, headers:string[], rows:any[][], detectedHeaderRow:number}>>}
     */
    function readFile(file, forceHeaderRow, idColumnHint = '') {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data     = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {
                        type:      'array',
                        cellDates: true,
                        dateNF:    'yyyy-mm-dd',
                    });

                    const sheets = workbook.SheetNames.map((sheetName) => {
                        const ws  = workbook.Sheets[sheetName];
                        const raw = XLSX.utils.sheet_to_json(ws, {
                            header: 1,
                            defval: '',
                            raw:    false,
                            dateNF: 'yyyy-mm-dd',
                        });

                        if (raw.length === 0) {
                            return { name: sheetName, headers: [], rows: [], detectedHeaderRow: 0 };
                        }

                        const headerIdx = (forceHeaderRow !== undefined && forceHeaderRow >= 0)
                            ? forceHeaderRow
                            : detectHeaderRow(raw, 40, idColumnHint);

                        const headerRow = raw[headerIdx] || raw[0];
                        const headers   = headerRow.map(h =>
                            h !== null && h !== undefined ? h.toString().trim() : ''
                        );

                        const dataRows = raw
                            .slice(headerIdx + 1)
                            .filter(row =>
                                row.some(cell =>
                                    cell !== '' && cell !== null && cell !== undefined
                                )
                            );

                        return {
                            name:              sheetName,
                            headers,
                            rows:              dataRows,
                            detectedHeaderRow: headerIdx,
                        };
                    });

                    resolve(sheets);
                } catch (err) {
                    reject(new Error(`Could not parse "${file.name}": ${err.message}`));
                }
            };

            reader.onerror = () =>
                reject(new Error(`Could not read file "${file.name}"`));

            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Return a specific sheet by 0-based index, falling back to the first.
     *
     * @param  {Array}  sheets
     * @param  {number} index  0-based
     * @returns {Object|null}
     */
    function getSheet(sheets, index) {
        if (!sheets || sheets.length === 0) return null;
        if (index < 0 || index >= sheets.length) return sheets[0];
        return sheets[index];
    }

    /* ── Drag-and-Drop Setup ──────────────────────────────────── */

    /**
     * Wire up a drop zone with drag-and-drop and click-to-browse.
     * No file-type filtering is applied — any file is accepted and
     * passed to the callback; SheetJS will report an error at
     * parse-time for unsupported formats.
     *
     * @param {HTMLElement}      dropZoneEl
     * @param {HTMLInputElement} inputEl       hidden <input type="file">
     * @param {Function}         onFilesAdded  callback(files: File[])
     * @param {boolean}          [multiple=true]
     */
    function setupDropZone(dropZoneEl, inputEl, onFilesAdded, multiple = true) {

        dropZoneEl.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZoneEl.classList.add('drag-over');
        });

        dropZoneEl.addEventListener('dragleave', (e) => {
            if (!dropZoneEl.contains(e.relatedTarget)) {
                dropZoneEl.classList.remove('drag-over');
            }
        });

        dropZoneEl.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZoneEl.classList.remove('drag-over');
            const files = Array.from(e.dataTransfer.files);
            if (files.length > 0) {
                onFilesAdded(multiple ? files : [files[0]]);
            }
        });

        dropZoneEl.addEventListener('click', (e) => {
            if (!e.target.closest('label') && e.target !== inputEl) {
                inputEl.click();
            }
        });

        inputEl.addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            if (files.length > 0) {
                onFilesAdded(files);
            }
            e.target.value = '';
        });
    }

    /* ── Public API ───────────────────────────────────────────── */
    return { readFile, getSheet, setupDropZone, detectHeaderRow };

})();
