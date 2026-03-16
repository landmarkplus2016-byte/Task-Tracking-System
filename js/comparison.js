/**
 * comparison.js
 * ──────────────────────────────────────────────────────────────
 * Combines multiple coordinator sheets into one dataset, then
 * compares that dataset against the master tracking sheet using
 * the ID# column as the unique key.
 *
 * Returns:
 *   newEntries     – rows present in coordinators but not in master
 *   changedEntries – rows present in both, with at least one value diff
 *   unchangedEntries – rows present in both with no differences
 *   changeDetails  – flat list of per-cell changes (ID, column, old, new)
 */

const Comparison = (() => {

    /* ── Helpers ──────────────────────────────────────────────── */

    /** Lowercase + trim a header for normalised comparison. */
    function normHeader(h) {
        return h ? h.toString().toLowerCase().trim().replace(/\s+/g, ' ') : '';
    }

    /** Normalise a cell value for comparison. */
    function normValue(v, caseSensitive = false) {
        if (v === null || v === undefined) return '';
        const s = v.toString().trim();
        return caseSensitive ? s : s.toLowerCase();
    }

    /**
     * Locate the ID column inside a normalised-headers array.
     * Tries several common patterns before falling back to any
     * header that contains the word "id".
     *
     * @param  {string[]} normHeaders  Already-normalised headers
     * @param  {string}   idName       User-supplied ID column name
     * @returns {number}  Index, or -1 if not found
     */
    function findIdIndex(normHeaders, idName) {
        const target = normHeader(idName);

        // 1. Exact match
        let idx = normHeaders.indexOf(target);
        if (idx !== -1) return idx;

        // 2. One contains the other
        idx = normHeaders.findIndex(
            h => h.includes(target) || target.includes(h)
        );
        if (idx !== -1) return idx;

        // 3. Fallback – anything with "id"
        idx = normHeaders.findIndex(h => /\bid\b/.test(h) || h.startsWith('id'));
        return idx;
    }

    /* ── Combine coordinator sheets ───────────────────────────── */

    /**
     * Merge multiple coordinator sheet objects into a single dataset.
     *
     * @param  {Array<{fileName:string, sheet:{name,headers,rows}}>} sheetsData
     * @param  {string}  idColumnName  User-supplied ID column name
     * @param  {boolean} caseSensitive
     * @returns {{
     *   headers: string[],        original header names (union of all sheets)
     *   normHeaders: string[],    normalised versions of headers
     *   idColumnNorm: string,     normalised ID column name
     *   idColumnOriginal: string, original ID column name
     *   rows: Map<string,Object>, id → rowObject (normHeader keys)
     *   duplicates: string[],     warning messages for duplicate IDs
     *   errors: string[]          error messages for skipped sheets
     * }}
     */
    function combineCoordinatorSheets(sheetsData, idColumnName, caseSensitive = false) {
        const headerOriginalByNorm = new Map(); // normH -> first-seen original
        const duplicates = [];
        const errors     = [];

        // ── Pass 1: Collect all unique headers (union) ──────────
        sheetsData.forEach(({ fileName, sheet }) => {
            if (!sheet || !sheet.headers || sheet.headers.length === 0) {
                errors.push(`"${fileName}": sheet appears to be empty or has no headers.`);
                return;
            }
            sheet.headers.forEach(h => {
                const n = normHeader(h);
                if (n && !headerOriginalByNorm.has(n)) {
                    headerOriginalByNorm.set(n, h);
                }
            });
        });

        const normHeadersList     = Array.from(headerOriginalByNorm.keys());
        const originalHeadersList = normHeadersList.map(n => headerOriginalByNorm.get(n));

        // ── Find ID column in the combined header list ──────────
        const idIdx = findIdIndex(normHeadersList, idColumnName);
        if (idIdx === -1) {
            errors.push(
                `ID column "${idColumnName}" was not found in any coordinator sheet. ` +
                `Please check the "ID Column Name" option.`
            );
            return {
                headers: originalHeadersList,
                normHeaders: normHeadersList,
                idColumnNorm: null,
                idColumnOriginal: null,
                rows: new Map(),
                duplicates,
                errors,
            };
        }

        const idColNorm     = normHeadersList[idIdx];
        const idColOriginal = originalHeadersList[idIdx];

        // ── Pass 2: Build row map keyed by ID ───────────────────
        const rowMap = new Map();

        sheetsData.forEach(({ fileName, sheet }) => {
            if (!sheet || !sheet.headers || sheet.headers.length === 0) return;

            const sheetNorm = sheet.headers.map(h => normHeader(h));
            const localIdIdx = findIdIndex(sheetNorm, idColumnName);

            if (localIdIdx === -1) {
                errors.push(
                    `"${fileName}": ID column "${idColumnName}" not found — sheet skipped.`
                );
                return;
            }

            sheet.rows.forEach((row, rowIdx) => {
                const rawId  = row[localIdIdx];
                const idRaw  = rawId !== undefined ? rawId.toString().trim() : '';
                if (!idRaw) return; // skip blank IDs

                // Use a lowercase key so lookups are always case-insensitive.
                // The original value is preserved inside rowObj.
                const idKey = idRaw.toLowerCase();

                if (rowMap.has(idKey)) {
                    duplicates.push(
                        `Duplicate ID "${idRaw}" in "${fileName}" (data row ${rowIdx + 1}) — ` +
                        `previous entry overwritten.`
                    );
                }

                // Build a normalised-keyed row object from all known columns
                const rowObj = {};
                normHeadersList.forEach((normH) => {
                    const srcIdx = sheetNorm.indexOf(normH);
                    rowObj[normH] =
                        srcIdx !== -1 && row[srcIdx] !== undefined
                            ? row[srcIdx].toString().trim()
                            : '';
                });
                rowObj['__source__'] = fileName;

                rowMap.set(idKey, rowObj);
            });
        });

        return {
            headers: originalHeadersList,
            normHeaders: normHeadersList,
            idColumnNorm: idColNorm,
            idColumnOriginal: idColOriginal,
            rows: rowMap,
            duplicates,
            errors,
        };
    }

    /* ── Parse Master Sheet ───────────────────────────────────── */

    /**
     * Convert the master tracking sheet into a searchable Map.
     *
     * @param  {{name,headers,rows}} sheet
     * @param  {string} idColumnName
     * @returns {{
     *   normHeaders: string[],
     *   originalHeaders: string[],
     *   idColumnNorm: string,
     *   rows: Map<string,Object>,
     *   error: string|null
     * }}
     */
    function parseMasterData(sheet, idColumnName) {
        if (!sheet || !sheet.headers || sheet.headers.length === 0) {
            return {
                normHeaders: [],
                originalHeaders: [],
                idColumnNorm: null,
                rows: new Map(),
                error: 'Master Tracking sheet is empty or has no headers.',
            };
        }

        const normHeaders = sheet.headers.map(h => normHeader(h));
        const idIdx       = findIdIndex(normHeaders, idColumnName);

        if (idIdx === -1) {
            return {
                normHeaders,
                originalHeaders: sheet.headers,
                idColumnNorm: null,
                rows: new Map(),
                error:
                    `ID column "${idColumnName}" was not found in the Master Tracking sheet. ` +
                    `Please check the "ID Column Name" option.`,
            };
        }

        const rowMap = new Map();

        sheet.rows.forEach((row) => {
            const rawId = row[idIdx];
            const idRaw = rawId !== undefined ? rawId.toString().trim() : '';
            if (!idRaw) return;

            // Lowercase key so it always matches coordinator keys
            const idKey = idRaw.toLowerCase();

            const rowObj = {};
            normHeaders.forEach((normH, i) => {
                rowObj[normH] = row[i] !== undefined ? row[i].toString().trim() : '';
            });

            rowMap.set(idKey, rowObj);
        });

        return {
            normHeaders,
            originalHeaders: sheet.headers,
            idColumnNorm: normHeaders[idIdx],
            rows: rowMap,
            error: null,
        };
    }

    /* ── Compare ──────────────────────────────────────────────── */

    /**
     * Compare combined coordinator data against master data.
     *
     * @param  {Object}  coordinatorData  Result of combineCoordinatorSheets()
     * @param  {Object}  masterData       Result of parseMasterData()
     * @param  {boolean} caseSensitive
     * @returns {{
     *   newEntries:       Array,
     *   changedEntries:   Array,
     *   unchangedEntries: Array,
     *   changeDetails:    Array   flat list of {id, column, oldValue, newValue}
     * }}
     */
    function compare(coordinatorData, masterData, caseSensitive = false) {
        const newEntries       = [];
        const changedEntries   = [];
        const unchangedEntries = [];
        const changeDetails    = [];

        coordinatorData.rows.forEach((coordRow, id) => {

            if (!masterData.rows.has(id)) {
                // ── New entry ──────────────────────────────────
                newEntries.push({
                    id,
                    row:    coordRow,
                    source: coordRow['__source__'] || '',
                });
                return;
            }

            // ── Compare field by field ─────────────────────────
            const masterRow = masterData.rows.get(id);
            const changes   = [];

            coordinatorData.normHeaders.forEach((normH, colIdx) => {
                if (normH === '__source__') return;

                const coordVal  = normValue(coordRow[normH],  caseSensitive);
                const masterVal = normValue(masterRow[normH], caseSensitive);

                if (coordVal !== masterVal) {
                    const colLabel = coordinatorData.headers[colIdx] || normH;
                    changes.push({
                        column:   colLabel,
                        oldValue: masterRow[normH] !== undefined ? masterRow[normH] : '',
                        newValue: coordRow[normH]  !== undefined ? coordRow[normH]  : '',
                    });
                    changeDetails.push({
                        id,
                        column:   colLabel,
                        oldValue: masterRow[normH] !== undefined ? masterRow[normH] : '',
                        newValue: coordRow[normH]  !== undefined ? coordRow[normH]  : '',
                    });
                }
            });

            if (changes.length > 0) {
                changedEntries.push({
                    id,
                    row:            coordRow,
                    masterRow,
                    changes,
                    changedColumns: changes.map(c => c.column).join(', '),
                    source:         coordRow['__source__'] || '',
                });
            } else {
                unchangedEntries.push({ id, row: coordRow });
            }
        });

        return { newEntries, changedEntries, unchangedEntries, changeDetails };
    }

    /* ── Public API ───────────────────────────────────────────── */
    return { combineCoordinatorSheets, parseMasterData, compare };

})();
