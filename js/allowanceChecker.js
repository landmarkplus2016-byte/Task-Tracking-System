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
        sheetRows:     [],      // unified rows from all Google Sheets
        filteredRows:  [],      // rows matching selected month + half
        masterJcSet:   new Set(), // SiteID-JC values from master tracking tab
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

    /* ── Number formatter ────────────────────────────────────── */
    function fmt(n) {
        return Number(n).toLocaleString(undefined, {
            minimumFractionDigits: 2, maximumFractionDigits: 2,
        });
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
    // Matches "Jan", "Feb", "Mar" … (case-insensitive 3-letter abbreviation)
    function matchesMonth(cellValue, _monthVal, monthName) {
        const v    = cellValue.toString().trim().toLowerCase();
        const abbr = monthName.slice(0, 3).toLowerCase();   // e.g. "jan", "feb"
        return v === abbr;
    }

    // Matches "First" or "Second" only (case-insensitive)
    function matchesHalf(cellValue, half) {
        const v = cellValue.toString().trim().toLowerCase();
        return half === 'first' ? v === 'first' : v === 'second';
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

    /* ── Allowance calculation ───────────────────────────────── */

    const MEMBER_FIELDS = ['engineer', 'tech1', 'tech2', 'tech3', 'driver'];

    /** Normalize a name for salary map lookups: lowercase, trim, collapse spaces. */
    function normName(s) {
        return (s || '').toString().toLowerCase().trim().replace(/\s+/g, ' ');
    }

    /**
     * Build a salary lookup Map keyed by normalized name.
     * Each entry: { dailySalary: number, bankAccount: string, name: string (canonical) }
     */
    function buildSalaryMap(salaryList) {
        const map = new Map();
        for (const s of salaryList) {
            map.set(normName(s.name), {
                dailySalary: parseFloat(s.dailySalary) || 0,
                bankAccount: s.bankAccount || '',
                name:        s.name,   // canonical form from list.xlsx
            });
        }
        return map;
    }

    /**
     * Look up a name (from a Google Sheet) in a salary map.
     *
     * Strategy:
     *   1. Exact normalized match: "Mostafa Ahmed Mohamed" → "mostafa ahmed mohamed"
     *   2. Word-prefix match: sheet "mostafa ahmed" is a word-aligned prefix of
     *      canonical "mostafa ahmed mohamed" → returns the canonical entry.
     *      This handles sheets that only store the first two name parts.
     *
     * Returns the salary entry or null if no match.
     */
    function lookupSalary(salaryMap, rawName) {
        const key = normName(rawName);
        if (!key) return null;

        // Pass 1: exact match
        if (salaryMap.has(key)) return salaryMap.get(key);

        // Pass 2: word-prefix match
        // "mostafa ahmed" matches "mostafa ahmed mohamed" (canonical starts with key + space)
        // Also handles the reverse: canonical "ali" matches sheet "ali hassan" (less common)
        for (const [canonKey, entry] of salaryMap) {
            if (canonKey.startsWith(key + ' ') || key.startsWith(canonKey + ' ')) {
                return entry;
            }
        }

        return null;
    }

    /**
     * Compute per-person allowance totals from the filtered rows.
     *
     * Names from Google Sheets are normalized and matched against the
     * team / driver salary maps from list.xlsx. When a match is found,
     * the canonical name from list.xlsx is used in the output.
     *
     * Per row:
     *   memberCount       = non-empty fields among engineer/tech1-3/driver
     *   baseAllowance     = row.allowance × memberCount
     *   vacationAllowance = Σ dailySalary for each member (only when
     *                       row.vacationAllowance is non-empty)
     *   rowTotal          = baseAllowance + vacationAllowance
     *
     * Per person:
     *   allowanceTotal    = Σ row.allowance for every row they appear in
     *   vacationTotal     = Σ their dailySalary for rows with vacation flag
     *   grandTotal        = allowanceTotal + vacationTotal
     *
     * @returns {{ people: object[], grandTotal: number, calcWarnings: string[] }}
     */
    function computeAllowances(filteredRows) {
        const teamSalaryMap   = buildSalaryMap(AppData.getSalaries());
        const driverSalaryMap = buildSalaryMap(AppData.getDriverSalaries());

        // normalized name → accumulator object
        const personMap    = new Map();
        const calcWarnings = [];
        const warnedNames  = new Set();   // suppress duplicate warnings
        let   grandTotal   = 0;

        const TEAM_FIELDS_LOCAL = ['engineer', 'tech1', 'tech2', 'tech3'];

        for (const row of filteredRows) {
            const allowancePerPerson = parseFloat(row.allowance) || 0;
            const hasVacation        = (row.vacationAllowance || '').trim() !== '';
            let   rowVacationTotal   = 0;
            let   rowMemberCount     = 0;

            /* ── Team fields (engineer / tech1-3) ───────────── */
            for (const field of TEAM_FIELDS_LOCAL) {
                const rawName = (row[field] || '').trim();
                if (!rawName) continue;

                rowMemberCount++;
                const normKey     = normName(rawName);
                const sal         = lookupSalary(teamSalaryMap, rawName);
                const displayName = sal ? sal.name : rawName;

                if (!personMap.has(normKey)) {
                    personMap.set(normKey, {
                        name:           displayName,
                        rows:           0,
                        allowanceTotal: 0,
                        vacationTotal:  0,
                        bankAccount:    sal ? sal.bankAccount : '',
                        isTeam:         true,
                    });
                } else {
                    // If this person was first seen as a driver, upgrade to team
                    personMap.get(normKey).isTeam = true;
                }

                const person = personMap.get(normKey);
                person.rows++;
                person.allowanceTotal += allowancePerPerson;

                if (hasVacation) {
                    if (sal && sal.dailySalary > 0) {
                        person.vacationTotal += sal.dailySalary;
                        rowVacationTotal     += sal.dailySalary;
                    } else if (!warnedNames.has(normKey)) {
                        warnedNames.add(normKey);
                        calcWarnings.push(
                            `"${rawName}" not found in Team Salaries — vacation allowance skipped.`
                        );
                    }
                }
            }

            /* ── Driver field ────────────────────────────────── */
            const drvRaw = (row.driver || '').trim();
            if (drvRaw) {
                rowMemberCount++;
                const normKey = normName(drvRaw);
                // Fall back to teamSalaryMap if driver is not in driver list
                const sal         = lookupSalary(driverSalaryMap, drvRaw) || lookupSalary(teamSalaryMap, drvRaw);
                const displayName = sal ? sal.name : drvRaw;

                if (!personMap.has(normKey)) {
                    personMap.set(normKey, {
                        name:           displayName,
                        rows:           0,
                        allowanceTotal: 0,
                        vacationTotal:  0,
                        bankAccount:    sal ? sal.bankAccount : '',
                        isTeam:         false,   // driver unless also in a team field
                    });
                }

                const person = personMap.get(normKey);
                person.rows++;
                person.allowanceTotal += allowancePerPerson;

                if (hasVacation) {
                    if (sal && sal.dailySalary > 0) {
                        person.vacationTotal += sal.dailySalary;
                        rowVacationTotal     += sal.dailySalary;
                    } else if (!sal && !warnedNames.has(normKey)) {
                        warnedNames.add(normKey);
                        calcWarnings.push(
                            `"${drvRaw}" not found in Driver Salaries — vacation allowance skipped.`
                        );
                    }
                }
            }

            grandTotal += (allowancePerPerson * rowMemberCount) + rowVacationTotal;
        }

        const people = Array.from(personMap.values())
            .map(p => ({ ...p, grandTotal: p.allowanceTotal + p.vacationTotal }))
            .sort((a, b) => a.name.localeCompare(b.name));

        return { people, grandTotal, calcWarnings };
    }

    /* ── Repetition check ────────────────────────────────────── */

    /**
     * For each unique day in the filtered dataset, scan all team-member
     * fields. If the same person appears in more than one row on the same
     * day, generate an error entry (pre-formatted HTML, values escaped).
     *
     * @returns {string[]}  Array of HTML error strings, one per duplicate.
     */
    function checkRepetitions(filteredRows) {
        // day → Map( name.toLowerCase() → { name, labels[] } )
        const dayMap = new Map();

        for (const row of filteredRows) {
            const day = (row.day || '').trim();
            if (!day) continue;

            const members = MEMBER_FIELDS
                .map(f => (row[f] || '').trim())
                .filter(Boolean);
            if (!members.length) continue;

            // Readable row label for the error message
            const parts = [];
            if (row.__source__) parts.push(esc(row.__source__));
            if (row.site)       parts.push(`Site: ${esc(row.site)}`);
            if (row.jc)         parts.push(`JC: ${esc(row.jc)}`);
            const label = parts.length ? parts.join(' / ') : '(unknown row)';

            if (!dayMap.has(day)) dayMap.set(day, new Map());
            const personMap = dayMap.get(day);

            for (const memberName of members) {
                const key = memberName.toLowerCase();
                if (!personMap.has(key)) personMap.set(key, { name: memberName, labels: [] });
                personMap.get(key).labels.push(label);
            }
        }

        const errors = [];

        // Sort days numerically where possible
        const sortedDays = Array.from(dayMap.keys())
            .sort((a, b) => (Number(a) - Number(b)) || a.localeCompare(b));

        for (const day of sortedDays) {
            const people = Array.from(dayMap.get(day).values())
                .sort((a, b) => a.name.localeCompare(b.name));

            for (const { name, labels } of people) {
                if (labels.length < 2) continue;

                const entries = labels
                    .map(l => `<span class="rep-entry">${l}</span>`)
                    .join('');

                errors.push(
                    `<span class="rep-header">Day&nbsp;<strong>${esc(day)}</strong> — ` +
                    `<strong>${esc(name)}</strong> appears in ${labels.length} rows:</span> ` +
                    entries
                );
            }
        }

        return errors;
    }

    /* ── Excel export ────────────────────────────────────────── */

    const TEAM_FIELDS = ['engineer', 'tech1', 'tech2', 'tech3'];

    /** Split the people array into team members vs drivers using the isTeam flag set during computeAllowances(). */
    function categorizeMembers(_filteredRows, people) {
        return {
            team:    people.filter(p => p.isTeam),
            drivers: people.filter(p => !p.isTeam),
        };
    }

    /** Apply bold + blue header style to a range of cells in a worksheet. */
    function styleHeaderRow(ws, rowIdx, colCount) {
        const hStyle = {
            font: { bold: true, color: { rgb: 'FFFFFF' } },
            fill: { fgColor: { rgb: '1A56DB' } },
            alignment: { horizontal: 'center' },
        };
        for (let c = 0; c < colCount; c++) {
            const addr = XLSX.utils.encode_cell({ r: rowIdx, c });
            if (ws[addr]) ws[addr].s = hStyle;
        }
    }

    /** Apply bold section-label style to a single cell. */
    function styleSectionLabel(ws, rowIdx) {
        const addr = XLSX.utils.encode_cell({ r: rowIdx, c: 0 });
        if (ws[addr]) ws[addr].s = {
            font: { bold: true, sz: 12, color: { rgb: '1A56DB' } },
        };
    }

    function generateExcel() {
        if (!state.results) return;

        const { people, monthName, half, masterSheets } = state.results;
        const allRows      = state.sheetRows;
        const filteredRows = state.filteredRows;
        const monthAbbr    = monthName.slice(0, 3);
        const halfStr      = half === 'first' ? 'First' : 'Second';

        const wb = XLSX.utils.book_new();

        /* ── Sheet 1: Filtered Tracking (selected month + half) ─ */
        const trackingTabName = `${monthAbbr} - ${halfStr}`;
        const trackingHeaders = [
            'Month', 'Day', 'Month Half', 'Coordinator', 'Site', 'Area',
            'Start Time', 'End Time', 'Project', 'Sub Project',
            'Engineer', 'Tech-1', 'Tech-2', 'Tech-3', 'Driver',
            'Allowance', 'Vacation Allowance', 'Work Details', 'JC',
        ];
        const trackingData = [
            trackingHeaders,
            ...filteredRows.map(r => [
                r.month, r.day, r.monthHalf, r.coordinator, r.site, r.area,
                r.startTime, r.endTime, r.project, r.subProject,
                r.engineer, r.tech1, r.tech2, r.tech3, r.driver,
                r.allowance, r.vacationAllowance, r.workDetails, r.jc,
            ]),
        ];
        const trackingSheet = XLSX.utils.aoa_to_sheet(trackingData);
        styleHeaderRow(trackingSheet, 0, trackingHeaders.length);
        trackingSheet['!cols'] = [
            { wch: 6 }, { wch: 5 }, { wch: 12 }, { wch: 22 }, { wch: 20 },
            { wch: 14 }, { wch: 11 }, { wch: 11 }, { wch: 18 }, { wch: 18 },
            { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 },
            { wch: 11 }, { wch: 18 }, { wch: 30 }, { wch: 12 },
        ];
        XLSX.utils.book_append_sheet(wb, trackingSheet, trackingTabName);

        /* ── Sheet 2: Allowance Amount ───────────────────────── */
        const { team, drivers } = categorizeMembers(filteredRows, people);

        const teamRows = team.map(p => [p.name, p.grandTotal, p.bankAccount]);
        const drvRows  = drivers.map(p => [p.name, p.grandTotal]);

        // Build as array-of-arrays so we control the layout precisely
        const aoa = [];
        const markers = {};   // { rowIdx → 'teamHeader' | 'drvHeader' | 'teamCols' | 'drvCols' }

        // Team section
        markers[aoa.length] = 'section';
        aoa.push(['Team']);
        markers[aoa.length] = 'cols3';
        aoa.push(['Name', 'Total Amount', 'Bank Account #']);
        teamRows.forEach(r => aoa.push(r));

        // Gap
        aoa.push([]);

        // Driver section
        markers[aoa.length] = 'section';
        aoa.push(['Driver']);
        markers[aoa.length] = 'cols2';
        aoa.push(['Name', 'Amount']);
        drvRows.forEach(r => aoa.push(r));

        const allowanceSheet = XLSX.utils.aoa_to_sheet(aoa);

        // Apply styles
        Object.entries(markers).forEach(([ri, type]) => {
            const rowIdx = parseInt(ri, 10);
            if (type === 'section')  styleSectionLabel(allowanceSheet, rowIdx);
            if (type === 'cols3')    styleHeaderRow(allowanceSheet, rowIdx, 3);
            if (type === 'cols2')    styleHeaderRow(allowanceSheet, rowIdx, 2);
        });

        allowanceSheet['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 22 }];
        XLSX.utils.book_append_sheet(wb, allowanceSheet, 'Allowance Amount');

        /* ── Download ────────────────────────────────────────── */
        XLSX.writeFile(wb, `Allowance_Report_${monthAbbr}_${halfStr}.xlsx`);
    }

    /* ── Master tracking parser ──────────────────────────────── */

    /**
     * Locate the "Tracking" tab in the master file and extract all
     * values from the "SiteID-JC" column into a Set for fast lookup.
     *
     * Tab matching:   exact name "tracking" → partial contains match
     * Column matching: normalised (spaces/dashes stripped) exact "siteidjc"
     *                  → any header that contains both "siteid" and "jc"
     *
     * @param  {Array}  masterSheets  Output of FileHandler.readFile()
     * @returns {{ jcSet: Set, tabName: string|null, colName: string|null, error: string|null }}
     */
    function parseMasterTracking(masterSheets) {
        const lc = s => s.toLowerCase().trim();

        // Find the 'Tracking' tab
        let sheet = masterSheets.find(s => lc(s.name) === 'tracking');
        if (!sheet) sheet = masterSheets.find(s => lc(s.name).includes('tracking'));

        if (!sheet) {
            return {
                jcSet: new Set(), tabName: null, colName: null,
                error: `Tab "Tracking" not found in master file. ` +
                       `Tabs available: ${masterSheets.map(s => `"${s.name}"`).join(', ')}`,
            };
        }

        // Find the 'SiteID-JC' column — normalise by removing spaces and dashes
        const norm = h => h.toLowerCase().replace(/[\s\-_]/g, '');
        const normHeaders = sheet.headers.map(norm);

        let colIdx = normHeaders.indexOf('siteidjc');
        if (colIdx === -1) colIdx = normHeaders.findIndex(h => h.includes('siteid') && h.includes('jc'));

        if (colIdx === -1) {
            return {
                jcSet: new Set(), tabName: sheet.name, colName: null,
                error: `Column "SiteID-JC" not found in "${sheet.name}" tab. ` +
                       `Headers found: ${sheet.headers.filter(Boolean).slice(0, 10).join(', ')}`,
            };
        }

        // Build a Set of every "SiteID-JC" combo in the master, lowercased.
        // Values in the master column look like "CABH185-MK01", "R5061-MK10", etc.
        const jcSet = new Set();
        for (const row of sheet.rows) {
            const val = (row[colIdx] ?? '').toString().trim();
            if (val) jcSet.add(val.toLowerCase());
        }

        return { jcSet, tabName: sheet.name, colName: sheet.headers[colIdx], error: null };
    }

    /* ── Month dropdown ──────────────────────────────────────── */
    function populateMonthDropdown() {
        const sel = $('allowanceMonth');
        const now = new Date();
        MONTHS.forEach((name, i) => {
            const opt = document.createElement('option');
            opt.value = i + 1;
            opt.textContent = name.slice(0, 3);
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
            // Error items are pre-formatted HTML — individual user-data values
            // are already escaped at the point of construction.
            $('allowanceErrorsList').innerHTML = errors.map(e => `<li>${e}</li>`).join('');
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
    function showResults(monthName, halfLabel, sourceCounts, filteredSourceCounts, totalRows, filteredCount, people, grandTotal) {
        $('allowanceResultsSummary').textContent = `${monthName} — ${halfLabel}`;

        /* ── Stat cards ─────────────────────────────────────── */
        const statCards = `
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
                <div class="allowance-fetch-stat allowance-fetch-stat--total">
                    <span class="allowance-fetch-num allowance-fetch-num--total">${fmt(grandTotal)}</span>
                    <span class="allowance-fetch-label">Grand Total</span>
                </div>
            </div>
        `;

        /* ── Per-person breakdown table ─────────────────────── */
        const totAllowance = people.reduce((s, p) => s + p.allowanceTotal, 0);
        const totVacation  = people.reduce((s, p) => s + p.vacationTotal,  0);
        const totGrand     = people.reduce((s, p) => s + p.grandTotal,     0);

        const personRows = people.map(p => `
            <tr>
                <td>${esc(p.name)}</td>
                <td class="allowance-td-num">${p.rows}</td>
                <td class="allowance-td-num">${fmt(p.allowanceTotal)}</td>
                <td class="allowance-td-num">${p.vacationTotal > 0 ? fmt(p.vacationTotal) : '<span class="allowance-nil">—</span>'}</td>
                <td class="allowance-td-num allowance-td-total">${fmt(p.grandTotal)}</td>
                <td class="allowance-td-bank">${esc(p.bankAccount) || '<span class="allowance-nil">—</span>'}</td>
            </tr>
        `).join('');

        const personTable = `
            <h3 class="allowance-section-title">Per-Person Breakdown</h3>
            <div class="allowance-table-wrap">
                <table class="allowance-table">
                    <thead>
                        <tr>
                            <th>Name</th>
                            <th class="allowance-th-num">Rows</th>
                            <th class="allowance-th-num">Allowance</th>
                            <th class="allowance-th-num">Vacation</th>
                            <th class="allowance-th-num">Total</th>
                            <th>Bank Account</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${personRows || '<tr><td colspan="6" class="allowance-empty">No data</td></tr>'}
                    </tbody>
                    <tfoot>
                        <tr class="allowance-table-tfoot">
                            <td>Totals</td>
                            <td class="allowance-td-num">—</td>
                            <td class="allowance-td-num">${fmt(totAllowance)}</td>
                            <td class="allowance-td-num">${totVacation > 0 ? fmt(totVacation) : '<span class="allowance-nil">—</span>'}</td>
                            <td class="allowance-td-num allowance-td-total">${fmt(totGrand)}</td>
                            <td></td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        `;

        /* ── Source breakdown table ──────────────────────────── */
        const sourceRows = Array.from(sourceCounts.entries()).map(([name, loaded]) => {
            const matched = filteredSourceCounts.get(name) || 0;
            return `
                <tr>
                    <td>${esc(name)}</td>
                    <td class="allowance-td-num">${loaded}</td>
                    <td class="allowance-td-num ${matched === 0 ? 'allowance-count--zero' : 'allowance-count'}">${matched}</td>
                </tr>
            `;
        }).join('');

        const sourceTable = `
            <h3 class="allowance-section-title">Data Sources</h3>
            <div class="allowance-table-wrap">
                <table class="allowance-table">
                    <thead>
                        <tr>
                            <th>Coordinator / Sheet</th>
                            <th class="allowance-th-num">Rows loaded</th>
                            <th class="allowance-th-num">Matched</th>
                        </tr>
                    </thead>
                    <tbody>${sourceRows || '<tr><td colspan="3" class="allowance-empty">No data</td></tr>'}</tbody>
                </table>
            </div>
        `;

        $('allowanceResultsBody').innerHTML = statCards + personTable + sourceTable;

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

            /* ── Step 1b: Parse master Tracking tab ──────────── */
            setProgress(8, 'Parsing master tracking tab…');
            const { jcSet, tabName, colName, error: masterError } =
                parseMasterTracking(masterSheets);

            if (masterError) {
                warnings.push(masterError);
            } else {
                state.masterJcSet = jcSet;
                console.log(`Master tracking: tab "${tabName}", column "${colName}", ${jcSet.size} combos`);
            }

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

            // Per-source filtered counts (for the Data Sources table)
            const filteredSourceCounts = new Map();
            for (const row of filteredRows) {
                const src = row.__source__ || '(unknown)';
                filteredSourceCounts.set(src, (filteredSourceCounts.get(src) || 0) + 1);
            }

            if (rows.length > 0 && filteredRows.length === 0) {
                warnings.push(
                    `No rows matched "${monthName} — ${halfLabel}". ` +
                    `Verify that the Month column uses a 3-letter abbreviation (e.g. "Jan", "Feb") ` +
                    `and Month Half uses "First" or "Second".`
                );
            }

            /* ── Step 3b: Compare SiteID-JC combos against master ── */
            if (jcSet.size > 0 && filteredRows.length > 0) {
                setProgress(90, 'Comparing SiteID-JC combos against master tracking…');

                // Each coordinator row can have multiple sites AND multiple job codes
                // separated by "/", e.g.:
                //   Site:  "CABH317/CABH338/CABH337"
                //   JC:    "MK22/MK23/MK24"
                //
                // Pairs are positional: CABH317-MK22, CABH338-MK23, CABH337-MK24.
                // The master "SiteID-JC" column stores these exact combinations.
                //
                // Rules:
                //   • If site count ≠ JC count → error (data problem in the sheet)
                //   • For each pair, check the combo against the master jcSet
                //   • Missing combos → warning

                const missingCombos = new Map();  // lowercase combo → display string

                for (const row of filteredRows) {
                    const siteRaw = (row.site || '').trim();
                    const jcRaw   = (row.jc   || '').trim();
                    if (!siteRaw || !jcRaw) continue;

                    const sites = siteRaw.split('/').map(s => s.trim()).filter(Boolean);
                    const jcs   = jcRaw.split('/').map(s => s.trim()).filter(Boolean);

                    // Count mismatch = error (shown in red)
                    if (sites.length !== jcs.length) {
                        const label = [row.__source__, row.day && `Day ${row.day}`]
                            .filter(Boolean).join(', ');
                        errors.push(
                            `Site/JC count mismatch${label ? ` (${esc(label)})` : ''}: ` +
                            `${sites.length} site(s) [${esc(siteRaw)}] but ` +
                            `${jcs.length} job code(s) [${esc(jcRaw)}]`
                        );
                        continue;
                    }

                    // Check each paired combo
                    for (let i = 0; i < sites.length; i++) {
                        const combo      = `${sites[i]}-${jcs[i]}`;
                        const comboLower = combo.toLowerCase();
                        if (!jcSet.has(comboLower) && !missingCombos.has(comboLower)) {
                            missingCombos.set(comboLower, combo);
                        }
                    }
                }

                if (missingCombos.size > 0) {
                    const list = [...missingCombos.values()].sort().join(', ');
                    warnings.push(
                        `${missingCombos.size} SiteID-JC combo(s) not found in master tracking ` +
                        `(tab: "${tabName}", column: "${colName}"): ${list}`
                    );
                }
            }

            /* ── Step 4: Compute allowances ──────────────────── */
            setProgress(92, 'Computing allowances…');

            const { people, grandTotal, calcWarnings } = computeAllowances(filteredRows);
            warnings.push(...calcWarnings);

            /* ── Step 5: Repetition check ────────────────────── */
            setProgress(97, 'Checking for repetitions…');

            const repErrors = checkRepetitions(filteredRows);
            errors.push(...repErrors);

            setProgress(100, 'Done!');
            hideProgress();

            state.results = { monthVal, monthName, half, halfLabel, masterSheets, people, grandTotal };
            showResults(monthName, halfLabel, sourceCounts, filteredSourceCounts, rows.length, filteredRows.length, people, grandTotal);
            showIssues(errors, warnings);

        } catch (err) {
            hideProgress();
            errors.push(esc(err.message));
            showIssues(errors, warnings);
        }
    }

    /* ── Reset ───────────────────────────────────────────────── */
    function reset() {
        state.masterFile   = null;
        state.sheetRows    = [];
        state.filteredRows = [];
        state.masterJcSet  = new Set();
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
        $('allowanceDownloadBtn').addEventListener('click', generateExcel);
        $('allowanceResetBtn').addEventListener('click', reset);
    }

    return { init };

})();
