/**
 * excelExport.js
 * ──────────────────────────────────────────────────────────────
 * Builds and downloads the output Excel workbook that contains:
 *
 *   Sheet 1 – Summary           Overview stats & metadata
 *   Sheet 2 – New Entries       Full rows of entries not in master
 *   Sheet 3 – Changed Entries   Full current rows with "Changed Columns" note
 *   Sheet 4 – Change Details    One row per changed cell (old vs new)
 *   Sheet 5 – Combined          Full merged coordinator dataset
 *   Sheet 6 – Unchanged         (optional) rows with no changes
 */

const ExcelExport = (() => {

    /* ── Helpers ──────────────────────────────────────────────── */

    /**
     * Build a worksheet from a 2-D array (header row + data rows).
     * @param  {string[]} headers
     * @param  {any[][]}  dataRows
     * @returns {Object} SheetJS worksheet
     */
    function buildSheet(headers, dataRows) {
        return XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
    }

    /**
     * Set column widths automatically based on cell content.
     * @param  {Object}   ws
     * @param  {string[]} headers
     * @param  {any[][]}  dataRows
     * @returns {Object} ws (mutated)
     */
    function autoWidth(ws, headers, dataRows) {
        const widths = headers.map(h => Math.max(String(h || '').length, 8));

        dataRows.forEach(row => {
            row.forEach((cell, i) => {
                const len = cell !== null && cell !== undefined
                    ? String(cell).length : 0;
                if (widths[i] !== undefined) {
                    widths[i] = Math.min(Math.max(widths[i], len), 60);
                }
            });
        });

        ws['!cols'] = widths.map(w => ({ wch: w + 2 }));
        return ws;
    }

    /**
     * Filter out internal "__source__" header/column from display.
     */
    function publicHeaders(coordinatorData) {
        return coordinatorData.normHeaders
            .map((normH, i) => ({ normH, original: coordinatorData.headers[i] }))
            .filter(({ normH }) => normH !== '__source__');
    }

    /* ── Sheet builders ───────────────────────────────────────── */

    function buildSummarySheet(coordinatorData, masterData, results, options) {
        const { newEntries, changedEntries, unchangedEntries, changeDetails } = results;
        const now = new Date();

        const rows = [
            ['Task Tracking System – Comparison Report'],
            [],
            ['Generated On',             now.toLocaleString()],
            ['ID Column Used',           coordinatorData.idColumnOriginal || options.idColumnName],
            ['Case-sensitive Comparison', options.caseSensitive ? 'Yes' : 'No'],
            [],
            ['Coordinator Files Processed', options.coordinatorFileCount || ''],
            ['Total Coordinator Entries',   coordinatorData.rows.size],
            ['Master Tracking Entries',     masterData.rows.size],
            [],
            ['─── RESULTS ───',  ''],
            ['New Entries (not in Master)',       newEntries.length],
            ['Changed Entries',                  changedEntries.length],
            ['Unchanged Entries',               unchangedEntries.length],
            ['Total Cell-level Changes',         changeDetails.length],
        ];

        const dupes = options.duplicates || [];
        if (dupes.length > 0) {
            rows.push([], ['─── WARNINGS – Duplicate IDs ───', dupes.length]);
            dupes.forEach(d => rows.push(['', d]));
        }

        const errors = options.errors || [];
        if (errors.length > 0) {
            rows.push([], ['─── ERRORS ───', errors.length]);
            errors.forEach(e => rows.push(['', e]));
        }

        const ws = XLSX.utils.aoa_to_sheet(rows);
        ws['!cols'] = [{ wch: 38 }, { wch: 60 }];
        return ws;
    }

    function buildNewEntriesSheet(coordinatorData, newEntries) {
        const cols    = publicHeaders(coordinatorData);
        const headers = [...cols.map(c => c.original), 'Source File'];

        if (newEntries.length === 0) {
            return XLSX.utils.aoa_to_sheet([['No new entries found.']]);
        }

        const dataRows = newEntries.map(entry => {
            const cells = cols.map(({ normH }) => entry.row[normH] || '');
            cells.push(entry.source || '');
            return cells;
        });

        return autoWidth(buildSheet(headers, dataRows), headers, dataRows);
    }

    function buildChangedEntriesSheet(coordinatorData, changedEntries) {
        const cols    = publicHeaders(coordinatorData);
        const headers = [...cols.map(c => c.original), 'Changed Columns', 'Source File'];

        if (changedEntries.length === 0) {
            return XLSX.utils.aoa_to_sheet([['No changed entries found.']]);
        }

        const dataRows = changedEntries.map(entry => {
            const cells = cols.map(({ normH }) => entry.row[normH] || '');
            cells.push(entry.changedColumns || '');
            cells.push(entry.source || '');
            return cells;
        });

        return autoWidth(buildSheet(headers, dataRows), headers, dataRows);
    }

    function buildChangeDetailsSheet(changeDetails) {
        if (changeDetails.length === 0) {
            return XLSX.utils.aoa_to_sheet([['No cell-level changes found.']]);
        }

        const headers  = ['ID#', 'Column Name', 'Old Value (Master)', 'New Value (Coordinator)'];
        const dataRows = changeDetails.map(d => [d.id, d.column, d.oldValue, d.newValue]);
        return autoWidth(buildSheet(headers, dataRows), headers, dataRows);
    }

    function buildCombinedSheet(coordinatorData) {
        const cols    = publicHeaders(coordinatorData);
        const headers = [...cols.map(c => c.original), 'Source File'];

        const dataRows = Array.from(coordinatorData.rows.values()).map(row => {
            const cells = cols.map(({ normH }) => row[normH] || '');
            cells.push(row['__source__'] || '');
            return cells;
        });

        return autoWidth(buildSheet(headers, dataRows), headers, dataRows);
    }

    function buildUnchangedSheet(coordinatorData, unchangedEntries) {
        const cols    = publicHeaders(coordinatorData);
        const headers = cols.map(c => c.original);

        if (unchangedEntries.length === 0) {
            return XLSX.utils.aoa_to_sheet([['No unchanged entries.']]);
        }

        const dataRows = unchangedEntries.map(entry =>
            cols.map(({ normH }) => entry.row[normH] || '')
        );

        return autoWidth(buildSheet(headers, dataRows), headers, dataRows);
    }

    /* ── Main generate function ───────────────────────────────── */

    /**
     * Build the output workbook.
     *
     * @param  {Object} coordinatorData  From Comparison.combineCoordinatorSheets()
     * @param  {Object} masterData       From Comparison.parseMasterData()
     * @param  {Object} results          From Comparison.compare()
     * @param  {Object} options
     *   @param {string}   options.idColumnName
     *   @param {number}   options.coordinatorFileCount
     *   @param {string[]} options.duplicates
     *   @param {string[]} options.errors
     *   @param {boolean}  options.includeUnchanged
     *   @param {boolean}  options.caseSensitive
     * @returns {Object} SheetJS workbook
     */
    function generate(coordinatorData, masterData, results, options = {}) {
        const wb = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(wb,
            buildSummarySheet(coordinatorData, masterData, results, options),
            'Summary'
        );

        XLSX.utils.book_append_sheet(wb,
            buildNewEntriesSheet(coordinatorData, results.newEntries),
            'New Entries'
        );

        XLSX.utils.book_append_sheet(wb,
            buildChangedEntriesSheet(coordinatorData, results.changedEntries),
            'Changed Entries'
        );

        XLSX.utils.book_append_sheet(wb,
            buildChangeDetailsSheet(results.changeDetails),
            'Change Details'
        );

        XLSX.utils.book_append_sheet(wb,
            buildCombinedSheet(coordinatorData),
            'Combined Coordinators'
        );

        if (options.includeUnchanged) {
            XLSX.utils.book_append_sheet(wb,
                buildUnchangedSheet(coordinatorData, results.unchangedEntries),
                'Unchanged'
            );
        }

        return wb;
    }

    /**
     * Trigger a browser download of the workbook.
     * @param {Object} wb       SheetJS workbook
     * @param {string} filename Target filename (.xlsx)
     */
    function download(wb, filename) {
        XLSX.writeFile(wb, filename);
    }

    /* ── Public API ───────────────────────────────────────────── */
    return { generate, download };

})();
