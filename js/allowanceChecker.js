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
        masterFile:      null,
        sheetRows:       [],      // unified rows from all Google Sheets
        filteredRows:    [],      // rows matching selected month + half
        masterJcSet:     new Set(), // SiteID-JC values from master tracking tab
        masterOldNewMap: new Map(), // SiteID-JC combo → 'Old' | 'New' | '' from master file
        results:         null,
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
                const res = await fetch(exportUrl, { cache: 'no-store' });
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
    /**
     * Parse a salary value that may be stored as a formatted string like
     * "EGP 308" or "EGP 8,000".  Strips any non-numeric characters (letters,
     * currency symbols, commas) before calling parseFloat.
     */
    function parseSalaryValue(str) {
        return parseFloat((str || '').toString().replace(/[^\d.]/g, '')) || 0;
    }

    function buildSalaryMap(salaryList) {
        const map = new Map();
        for (const s of salaryList) {
            map.set(normName(s.name), {
                dailySalary: parseSalaryValue(s.dailySalary),
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
     * @returns {{ people: object[], grandTotal: number, calcWarnings: string[], calcErrors: string[] }}
     */
    function computeAllowances(filteredRows) {
        const teamSalaryMap   = buildSalaryMap(AppData.getSalaries());
        const driverSalaryMap = buildSalaryMap(AppData.getDriverSalaries());

        const personMap    = new Map();
        const calcWarnings = [];
        // normKey → { rawName, listType: 'Team'|'Driver', sources: Set<string> }
        const missingMap   = new Map();
        let   grandTotal   = 0;

        const TEAM_FIELDS_LOCAL = ['engineer', 'tech1', 'tech2', 'tech3'];

        for (const row of filteredRows) {
            const allowancePerPerson = parseFloat(row.allowance) || 0;
            const hasVacation        = (row.vacationAllowance || '').trim() !== '';
            let   rowVacationTotal   = 0;
            let   rowMemberCount     = 0;
            const source             = row.__source__ || '(unknown)';

            /* ── Team fields (engineer / tech1-3) ───────────── */
            for (const field of TEAM_FIELDS_LOCAL) {
                const rawName = (row[field] || '').trim();
                if (!rawName) continue;

                rowMemberCount++;
                const sal         = lookupSalary(teamSalaryMap, rawName);
                const displayName = sal ? sal.name : rawName;
                const normKey     = normName(displayName);

                if (!sal) {
                    if (!missingMap.has(normKey)) {
                        missingMap.set(normKey, { rawName, listType: 'Team', sources: new Set() });
                    }
                    missingMap.get(normKey).sources.add(source);
                }

                if (!personMap.has(normKey)) {
                    personMap.set(normKey, {
                        name:           displayName,
                        rows:           0,
                        daysWorked:     new Set(),
                        allowanceTotal: 0,
                        vacationTotal:  0,
                        bankAccount:    sal ? sal.bankAccount : '',
                        isTeam:         true,
                    });
                } else {
                    personMap.get(normKey).isTeam = true;
                }

                const person = personMap.get(normKey);
                person.rows++;
                person.daysWorked.add((row.day || '').trim());
                person.allowanceTotal += allowancePerPerson;

                if (hasVacation && sal && sal.dailySalary > 0) {
                    person.vacationTotal += sal.dailySalary;
                    rowVacationTotal     += sal.dailySalary;
                }
            }

            /* ── Driver field ────────────────────────────────── */
            const drvRaw = (row.driver || '').trim();
            if (drvRaw) {
                rowMemberCount++;
                const sal         = lookupSalary(driverSalaryMap, drvRaw)
                                 || lookupSalary(teamSalaryMap, drvRaw);
                const displayName = sal ? sal.name : drvRaw;
                const normKey     = normName(displayName);

                if (!sal) {
                    if (!missingMap.has(normKey)) {
                        missingMap.set(normKey, { rawName: drvRaw, listType: 'Driver', sources: new Set() });
                    }
                    missingMap.get(normKey).sources.add(source);
                }

                if (!personMap.has(normKey)) {
                    personMap.set(normKey, {
                        name:           displayName,
                        rows:           0,
                        daysWorked:     new Set(),
                        allowanceTotal: 0,
                        vacationTotal:  0,
                        bankAccount:    sal ? sal.bankAccount : '',
                        isTeam:         false,
                    });
                }

                const person = personMap.get(normKey);
                person.rows++;
                person.daysWorked.add((row.day || '').trim());
                person.allowanceTotal += allowancePerPerson;

                if (hasVacation && sal && sal.dailySalary > 0) {
                    person.vacationTotal += sal.dailySalary;
                    rowVacationTotal     += sal.dailySalary;
                }
            }

            grandTotal += (allowancePerPerson * rowMemberCount) + rowVacationTotal;
        }

        // Build pre-HTML error strings (same style as checkRepetitions)
        const calcErrors = Array.from(missingMap.values())
            .sort((a, b) => a.rawName.localeCompare(b.rawName))
            .map(({ rawName, listType, sources }) => {
                const srcSpans = [...sources].sort()
                    .map(s => `<span class="rep-entry">${esc(s)}</span>`)
                    .join('');
                return `<span class="rep-header"><strong>${esc(rawName)}</strong> was not found in the ${listType} Salaries list — ` +
                       `found in ${sources.size} sheet${sources.size > 1 ? 's' : ''}:</span> ${srcSpans}`;
            });

        const people = Array.from(personMap.values())
            .map(p => ({ ...p, daysWorked: p.daysWorked.size, grandTotal: p.allowanceTotal + p.vacationTotal }))
            .sort((a, b) => a.name.localeCompare(b.name));

        return { people, grandTotal, calcWarnings, calcErrors };
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

    /* ── New/Old split helpers ───────────────────────────────── */

    /**
     * For a single row, classify each site-JC pair as Old or New.
     * Rules (applied in order per pair):
     *   1. JC contains "CCTV" (case-insensitive) → always Old
     *   2. Combo found in masterOldNewMap with value "Old" → Old
     *   3. Everything else (not found, blank, "New") → New
     * Returns { oldCount, newCount, total }.
     */
    function classifyRowPairs(row, siteJcMap) {
        const siteRaw = (row.site || '').trim();
        const jcRaw   = (row.jc   || '').trim();

        // No site/jc: treat whole row as one pair; CCTV → Old, else New
        if (!siteRaw || !jcRaw) {
            const isCctv = jcRaw.toUpperCase().includes('CCTV');
            return isCctv ? { oldCount: 1, newCount: 0, total: 1 }
                          : { oldCount: 0, newCount: 1, total: 1 };
        }

        const sites = siteRaw.split('/').map(s => s.trim()).filter(Boolean);
        const jcs   = jcRaw.split('/').map(s => s.trim()).filter(Boolean);
        const n     = Math.min(sites.length, jcs.length);

        if (n === 0) return { oldCount: 0, newCount: 1, total: 1 };

        let oldCount = 0, newCount = 0;
        for (let i = 0; i < n; i++) {
            const isCctv = jcs[i].toUpperCase().includes('CCTV');
            const combo  = `${sites[i]}-${jcs[i]}`.toLowerCase();
            const val    = (siteJcMap.get(combo) || '').toLowerCase();
            if (isCctv || val === 'old') oldCount++;
            else newCount++;   // 'new', blank, or not found → New
        }

        return { oldCount, newCount, total: n };
    }

    /**
     * Like computeAllowances() but splits each person's totals into Old and New
     * based on the site-JC pair classification from siteJcMap.
     */
    function computeSplitAllowances(filteredRows, siteJcMap) {
        const teamSalaryMap   = buildSalaryMap(AppData.getSalaries());
        const driverSalaryMap = buildSalaryMap(AppData.getDriverSalaries());

        // normKey → { name, oldAllowance, oldVacation, newAllowance, newVacation, bankAccount, isTeam }
        const personMap = new Map();
        let oldGrandTotal = 0;
        let newGrandTotal = 0;

        const TEAM_FIELDS_LOCAL = ['engineer', 'tech1', 'tech2', 'tech3'];

        for (const row of filteredRows) {
            const allowancePerPerson = parseFloat(row.allowance) || 0;
            const hasVacation        = (row.vacationAllowance || '').trim() !== '';

            const { oldCount, newCount, total } = classifyRowPairs(row, siteJcMap);
            const oldFrac = total > 0 ? oldCount / total : 0;
            const newFrac = total > 0 ? newCount / total : 1;

            const oldAllow = allowancePerPerson * oldFrac;
            const newAllow = allowancePerPerson * newFrac;

            let rowOldVac = 0, rowNewVac = 0, memberCount = 0;

            const processMember = (rawName, isTeam) => {
                if (!rawName) return;
                memberCount++;

                const sal         = isTeam
                    ? lookupSalary(teamSalaryMap, rawName)
                    : (lookupSalary(driverSalaryMap, rawName) || lookupSalary(teamSalaryMap, rawName));
                const displayName = sal ? sal.name : rawName;
                const normKey     = normName(displayName);

                if (!personMap.has(normKey)) {
                    personMap.set(normKey, {
                        name:         displayName,
                        oldAllowance: 0, oldVacation: 0,
                        newAllowance: 0, newVacation: 0,
                        bankAccount:  sal ? sal.bankAccount : '',
                        isTeam,
                    });
                } else if (isTeam) {
                    personMap.get(normKey).isTeam = true;
                }

                const p = personMap.get(normKey);
                p.oldAllowance += oldAllow;
                p.newAllowance += newAllow;

                if (hasVacation && sal && sal.dailySalary > 0) {
                    p.oldVacation += sal.dailySalary * oldFrac;
                    p.newVacation += sal.dailySalary * newFrac;
                    rowOldVac     += sal.dailySalary * oldFrac;
                    rowNewVac     += sal.dailySalary * newFrac;
                }
            };

            for (const field of TEAM_FIELDS_LOCAL) processMember((row[field] || '').trim(), true);
            processMember((row.driver || '').trim(), false);

            oldGrandTotal += oldAllow * memberCount + rowOldVac;
            newGrandTotal += newAllow * memberCount + rowNewVac;
        }

        const toPeople = (isOld) => Array.from(personMap.values())
            .map(p => ({
                name:        p.name,
                grandTotal:  isOld ? p.oldAllowance + p.oldVacation : p.newAllowance + p.newVacation,
                bankAccount: p.bankAccount,
                isTeam:      p.isTeam,
            }))
            .filter(p => p.grandTotal > 0)
            .sort((a, b) => a.name.localeCompare(b.name));

        return {
            oldPeople: toPeople(true),
            newPeople: toPeople(false),
            oldGrandTotal,
            newGrandTotal,
        };
    }

    /**
     * Split filtered rows into Old and New subsets.
     * Site and JC fields are rebuilt to contain ONLY the pairs that belong
     * to each category, so a row with "K3960 / K5402" where K3960 is New
     * and K5402 is Old will produce:
     *   Old row:  site="K5402",  jc=<jc2>, allowance = origAllow / 2
     *   New row:  site="K3960",  jc=<jc1>, allowance = origAllow / 2
     */
    function buildSplitTrackingRows(filteredRows, siteJcMap) {
        const oldRows = [], newRows = [];

        for (const row of filteredRows) {
            const siteRaw   = (row.site || '').trim();
            const jcRaw     = (row.jc   || '').trim();
            const origAllow = parseFloat(row.allowance) || 0;

            const sites = siteRaw ? siteRaw.split('/').map(s => s.trim()).filter(Boolean) : [];
            const jcs   = jcRaw  ? jcRaw.split('/').map(s => s.trim()).filter(Boolean)   : [];
            const n     = Math.min(sites.length, jcs.length);

            if (n === 0) {
                // No parseable pairs — classify the whole row and keep as-is
                const { oldCount, newCount, total } = classifyRowPairs(row, siteJcMap);
                const oldFrac = oldCount / (total || 1);
                const newFrac = newCount / (total || 1);
                if (oldFrac > 0) oldRows.push({ ...row, allowance: String(origAllow * oldFrac) });
                if (newFrac > 0) newRows.push({ ...row, allowance: String(origAllow * newFrac) });
                continue;
            }

            // Separate each pair into its category
            const oldSites = [], oldJcs = [];
            const newSites = [], newJcs = [];

            for (let i = 0; i < n; i++) {
                const isCctv = jcs[i].toUpperCase().includes('CCTV');
                const combo  = `${sites[i]}-${jcs[i]}`.toLowerCase();
                const val    = (siteJcMap.get(combo) || '').toLowerCase();
                if (isCctv || val === 'old') {
                    oldSites.push(sites[i]);
                    oldJcs.push(jcs[i]);
                } else {
                    newSites.push(sites[i]);
                    newJcs.push(jcs[i]);
                }
            }

            const allowPerPair = origAllow / n;

            if (oldSites.length > 0) {
                oldRows.push({
                    ...row,
                    site:      oldSites.join(' / '),
                    jc:        oldJcs.join(' / '),
                    allowance: String(allowPerPair * oldSites.length),
                });
            }
            if (newSites.length > 0) {
                newRows.push({
                    ...row,
                    site:      newSites.join(' / '),
                    jc:        newJcs.join(' / '),
                    allowance: String(allowPerPair * newSites.length),
                });
            }
        }
        return { oldRows, newRows };
    }

    /** Build and download one workbook (Old or New) with Tracking + Allowance Amount sheets. */
    function buildSplitWorkbook(trackingRows, people, grandTotal, monthAbbr, halfStr, label) {
        const wb = XLSX.utils.book_new();

        /* ── Sheet 1: Tracking ─────────────────────────────── */
        const trackingHeaders = [
            'Month', 'Day', 'Month Half', 'Coordinator', 'Site', 'Area',
            'Start Time', 'End Time', 'Project', 'Sub Project',
            'Engineer', 'Tech-1', 'Tech-2', 'Tech-3', 'Driver',
            'Allowance', 'Vacation Allowance', 'Work Details', 'JC',
        ];

        const summaryRow = new Array(trackingHeaders.length).fill('');
        summaryRow[8] = 'Total Allowance';
        summaryRow[9] = grandTotal;

        const trackingData = [
            summaryRow,
            new Array(trackingHeaders.length).fill(''),
            trackingHeaders,
            ...trackingRows.map(r => [
                r.month, r.day, r.monthHalf, r.coordinator, r.site, r.area,
                r.startTime, r.endTime, r.project, r.subProject,
                r.engineer, r.tech1, r.tech2, r.tech3, r.driver,
                r.allowance, r.vacationAllowance, r.workDetails, r.jc,
            ]),
        ];

        const trackingSheet = XLSX.utils.aoa_to_sheet(trackingData);
        const totalStyle = {
            font:      { bold: true, color: { rgb: 'FFFFFF' } },
            fill:      { fgColor: { rgb: '00B050' } },
            alignment: { horizontal: 'center', vertical: 'center' },
        };
        ['I1', 'J1'].forEach(addr => {
            if (trackingSheet[addr]) trackingSheet[addr].s = totalStyle;
        });
        styleHeaderRow(trackingSheet, 2, trackingHeaders.length);
        trackingSheet['!cols'] = [
            { wch: 6 }, { wch: 5 }, { wch: 12 }, { wch: 22 }, { wch: 20 },
            { wch: 14 }, { wch: 11 }, { wch: 11 }, { wch: 18 }, { wch: 18 },
            { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 },
            { wch: 11 }, { wch: 18 }, { wch: 30 }, { wch: 12 },
        ];
        trackingSheet['!freeze'] = { xSplit: 0, ySplit: 3 };

        const tabName = `${monthAbbr} - ${halfStr} - ${label}`.slice(0, 31);
        XLSX.utils.book_append_sheet(wb, trackingSheet, tabName);

        /* ── Sheet 2: Allowance Amount ─────────────────────── */
        const { team, drivers } = categorizeMembers([], people);
        const teamRows = team.map(p => [p.name, p.grandTotal, p.bankAccount]);
        const drvRows  = drivers.map(p => [p.name, p.grandTotal]);

        const aoa = [], markers = {};
        markers[aoa.length] = 'section'; aoa.push(['Team']);
        markers[aoa.length] = 'cols3';   aoa.push(['Name', 'Total Amount', 'Bank Account #']);
        teamRows.forEach(r => aoa.push(r));
        aoa.push([]);
        markers[aoa.length] = 'section'; aoa.push(['Driver']);
        markers[aoa.length] = 'cols2';   aoa.push(['Name', 'Amount']);
        drvRows.forEach(r => aoa.push(r));

        const allowanceSheet = XLSX.utils.aoa_to_sheet(aoa);
        Object.entries(markers).forEach(([ri, type]) => {
            const rowIdx = parseInt(ri, 10);
            if (type === 'section') styleSectionLabel(allowanceSheet, rowIdx);
            if (type === 'cols3')   styleHeaderRow(allowanceSheet, rowIdx, 3);
            if (type === 'cols2')   styleHeaderRow(allowanceSheet, rowIdx, 2);
        });
        allowanceSheet['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 22 }];
        XLSX.utils.book_append_sheet(wb, allowanceSheet, 'Allowance Amount');

        /* ── Sheet 3: Per Person breakdown ─────────────────── */
        XLSX.utils.book_append_sheet(wb, buildPerPersonSheet(trackingRows), 'Per Person');

        XLSX.writeFile(wb, `Allowance_${label}_${monthAbbr}_${halfStr}.xlsx`, { compression: true });
    }

    /** Generate the two split Excel files (Old and New). */
    function generateNewOldFiles() {
        if (!state.results) {
            alert('Please run analysis first.');
            return;
        }

        const { monthName, half } = state.results;
        const monthAbbr  = monthName.slice(0, 3);
        const halfStr    = half === 'first' ? 'First' : 'Second';
        const siteJcMap  = state.masterOldNewMap;  // populated during runAnalysis()

        const { oldPeople, newPeople, oldGrandTotal, newGrandTotal } =
            computeSplitAllowances(state.filteredRows, siteJcMap);

        const { oldRows, newRows } = buildSplitTrackingRows(state.filteredRows, siteJcMap);

        buildSplitWorkbook(oldRows, oldPeople, oldGrandTotal, monthAbbr, halfStr, 'Old');
        buildSplitWorkbook(newRows, newPeople, newGrandTotal, monthAbbr, halfStr, 'New');
    }

    /** Build a per-person breakdown sheet: sectioned by role, each person gets their own header + rows + total. */
    function buildPerPersonSheet(trackingRows) {
        const teamSalaryMap   = buildSalaryMap(AppData.getSalaries());
        const driverSalaryMap = buildSalaryMap(AppData.getDriverSalaries());

        const ROLE_LABELS = { engineer: 'Engineer', tech1: 'Tech-1', tech2: 'Tech-2', tech3: 'Tech-3', driver: 'Driver' };

        // One Map per role: normKey → { displayName, entries[] }
        const roleGroups = new Map([
            ['engineer', new Map()],
            ['tech1',    new Map()],
            ['tech2',    new Map()],
            ['tech3',    new Map()],
            ['driver',   new Map()],
        ]);

        for (const row of trackingRows) {
            const allowance   = parseFloat(row.allowance) || 0;
            const hasVacation = (row.vacationAllowance || '').trim() !== '';

            for (const field of ['engineer', 'tech1', 'tech2', 'tech3', 'driver']) {
                const rawName = (row[field] || '').trim();
                if (!rawName) continue;

                const isDriver    = field === 'driver';
                const sal         = isDriver
                    ? (lookupSalary(driverSalaryMap, rawName) || lookupSalary(teamSalaryMap, rawName))
                    : lookupSalary(teamSalaryMap, rawName);
                const displayName = sal ? sal.name : rawName;
                const key         = normName(displayName);
                const vacAmt      = (hasVacation && sal && sal.dailySalary > 0) ? sal.dailySalary : 0;

                const group = roleGroups.get(field);
                if (!group.has(key)) group.set(key, { displayName, entries: [] });
                group.get(key).entries.push({
                    month: row.month, day: row.day, monthHalf: row.monthHalf,
                    coordinator: row.coordinator, site: row.site, area: row.area,
                    startTime: row.startTime, endTime: row.endTime,
                    project: row.project, subProject: row.subProject,
                    role: ROLE_LABELS[field], allowance, vacAmt,
                    workDetails: row.workDetails, jc: row.jc,
                });
            }
        }

        const HEADERS  = [
            'Month', 'Day', 'Month Half', 'Coordinator', 'Site', 'Area',
            'Project', 'Sub Project',
            'Name', 'Allowance', 'Vacation Allowance', 'Work Details', 'JC',
        ];
        const N        = HEADERS.length;
        const COL_NAME  = 8, COL_ALLOW = 9, COL_VAC = 10;

        const aoa              = [];
        const sectionIndices   = [];  // orange section-label rows
        const colHeaderIndices = [];  // blue column-header rows (one per employee)
        const totalIndices     = [];  // green total rows

        function hasAny(roleKey) {
            return [...roleGroups.get(roleKey).values()].some(g => g.entries.length > 0);
        }

        function addSectionLabel(label) {
            const row = new Array(N).fill('');
            row[0] = label;
            sectionIndices.push(aoa.length);
            aoa.push(row);
        }

        function addRoleGroup(roleKey) {
            const sorted = [...roleGroups.get(roleKey).values()]
                .filter(g => g.entries.length > 0)
                .sort((a, b) => a.displayName.localeCompare(b.displayName));

            for (const { displayName, entries } of sorted) {
                entries.sort((a, b) =>
                    (Number(a.day) - Number(b.day)) || String(a.day).localeCompare(String(b.day))
                );

                // Column headers for this employee
                colHeaderIndices.push(aoa.length);
                aoa.push([...HEADERS]);

                // Data rows
                for (const e of entries) {
                    aoa.push([
                        e.month, e.day, e.monthHalf, e.coordinator, e.site, e.area,
                        e.project, e.subProject,
                        displayName, e.allowance,
                        e.vacAmt > 0 ? e.vacAmt : '',
                        e.workDetails, e.jc,
                    ]);
                }

                // Total row (allowance + vacation)
                const totalAllow = entries.reduce((s, e) => s + e.allowance, 0);
                const totalVac   = entries.reduce((s, e) => s + e.vacAmt,    0);
                const totalRow   = new Array(N).fill('');
                totalRow[COL_NAME]  = displayName + ' — Total';
                totalRow[COL_ALLOW] = totalAllow;
                if (totalVac > 0) totalRow[COL_VAC] = totalVac;
                totalIndices.push(aoa.length);
                aoa.push(totalRow);
                aoa.push(new Array(N).fill(''));  // spacer
            }
        }

        if (hasAny('engineer')) { addSectionLabel('Engineer');     addRoleGroup('engineer'); }
        if (['tech1','tech2','tech3'].some(k => hasAny(k))) {
            addSectionLabel('Technicians');
            addRoleGroup('tech1'); addRoleGroup('tech2'); addRoleGroup('tech3');
        }
        if (hasAny('driver'))   { addSectionLabel('Driver');       addRoleGroup('driver');   }

        const ws = XLSX.utils.aoa_to_sheet(aoa);

        // Section labels — orange background
        const sectionStyle = {
            font:      { bold: true, sz: 12, color: { rgb: 'FFFFFF' } },
            fill:      { fgColor: { rgb: 'E67700' } },
            alignment: { horizontal: 'left' },
        };
        for (const ri of sectionIndices) {
            for (let c = 0; c < N; c++) {
                const addr = XLSX.utils.encode_cell({ r: ri, c });
                if (!ws[addr]) ws[addr] = { v: '', t: 's' };
                ws[addr].s = sectionStyle;
            }
        }

        // Column headers — blue (reuse styleHeaderRow)
        for (const ri of colHeaderIndices) styleHeaderRow(ws, ri, N);

        // Total rows — green
        const totalStyle = {
            font:      { bold: true, color: { rgb: 'FFFFFF' } },
            fill:      { fgColor: { rgb: '00B050' } },
            alignment: { horizontal: 'center' },
        };
        for (const ri of totalIndices) {
            for (let c = 0; c < N; c++) {
                const addr = XLSX.utils.encode_cell({ r: ri, c });
                if (!ws[addr]) ws[addr] = { v: '', t: 's' };
                ws[addr].s = totalStyle;
            }
        }

        ws['!cols'] = [
            { wch: 6 }, { wch: 5 }, { wch: 12 }, { wch: 22 }, { wch: 20 },
            { wch: 14 }, { wch: 18 }, { wch: 18 },
            { wch: 22 }, { wch: 11 }, { wch: 18 }, { wch: 30 }, { wch: 12 },
        ];

        return ws;
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

        const { people, monthName, half, masterSheets, grandTotal } = state.results;
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

        // Row 1: Total Allowance summary at columns I & J (indices 8 & 9)
        const summaryRow = new Array(trackingHeaders.length).fill('');
        summaryRow[8] = 'Total Allowance';
        summaryRow[9] = grandTotal;

        const trackingData = [
            summaryRow,                                    // row 1 — summary
            new Array(trackingHeaders.length).fill(''),    // row 2 — empty spacer
            trackingHeaders,                               // row 3 — headers
            ...filteredRows.map(r => [
                r.month, r.day, r.monthHalf, r.coordinator, r.site, r.area,
                r.startTime, r.endTime, r.project, r.subProject,
                r.engineer, r.tech1, r.tech2, r.tech3, r.driver,
                r.allowance, r.vacationAllowance, r.workDetails, r.jc,
            ]),
        ];
        const trackingSheet = XLSX.utils.aoa_to_sheet(trackingData);

        // Style Total Allowance cells — green (#00B050) background, white bold text
        const totalStyle = {
            font:      { bold: true, color: { rgb: 'FFFFFF' } },
            fill:      { fgColor: { rgb: '00B050' } },
            alignment: { horizontal: 'center', vertical: 'center' },
        };
        ['I1', 'J1'].forEach(addr => {
            if (trackingSheet[addr]) trackingSheet[addr].s = totalStyle;
        });

        // Headers are now at row index 2 (0-based)
        styleHeaderRow(trackingSheet, 2, trackingHeaders.length);

        trackingSheet['!cols'] = [
            { wch: 6 }, { wch: 5 }, { wch: 12 }, { wch: 22 }, { wch: 20 },
            { wch: 14 }, { wch: 11 }, { wch: 11 }, { wch: 18 }, { wch: 18 },
            { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 },
            { wch: 11 }, { wch: 18 }, { wch: 30 }, { wch: 12 },
        ];
        // Freeze pane below the header row (row 3)
        trackingSheet['!freeze'] = { xSplit: 0, ySplit: 3 };
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

        /* ── Sheet 3: Per Person breakdown ───────────────────── */
        XLSX.utils.book_append_sheet(wb, buildPerPersonSheet(filteredRows), 'Per Person');

        /* ── Download ────────────────────────────────────────── */
        XLSX.writeFile(wb, `Allowance_Report_${monthAbbr}_${halfStr}.xlsx`, { compression: true });
    }

    /* ── Master tracking parser ──────────────────────────────── */

    /**
     * Locate the "Tracking" tab in the master file and extract:
     *   - jcSet:     Set of all SiteID-JC combos (lowercased) for JC validation
     *   - oldNewMap: Map of combo (lowercase) → 'Old' | 'New' | '' for the split feature
     *
     * Tab matching:   exact name "tracking" → partial contains match
     * Column matching: normalised (spaces/dashes stripped) exact "siteidjc"
     *                  → any header that contains both "siteid" and "jc"
     *
     * @param  {Array}  masterSheets  Output of FileHandler.readFile()
     * @returns {{ jcSet: Set, oldNewMap: Map, tabName: string|null, colName: string|null, error: string|null }}
     */
    function parseMasterTracking(masterSheets) {
        const lc = s => s.toLowerCase().trim();

        // Find the 'Tracking' tab
        let sheet = masterSheets.find(s => lc(s.name) === 'tracking');
        if (!sheet) sheet = masterSheets.find(s => lc(s.name).includes('tracking'));

        if (!sheet) {
            return {
                jcSet: new Set(), oldNewMap: new Map(), tabName: null, colName: null,
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
                jcSet: new Set(), oldNewMap: new Map(), tabName: sheet.name, colName: null,
                error: `Column "SiteID-JC" not found in "${sheet.name}" tab. ` +
                       `Headers found: ${sheet.headers.filter(Boolean).slice(0, 10).join(', ')}`,
            };
        }

        // Find the 'Old/New' column for the split feature
        let oldNewIdx = normHeaders.indexOf('oldnew');
        if (oldNewIdx === -1) oldNewIdx = normHeaders.findIndex(h => h.includes('old') && h.includes('new'));

        // Build Set (for JC validation) and Map (for Old/New split) together
        const jcSet    = new Set();
        const oldNewMap = new Map();
        for (const row of sheet.rows) {
            const val = (row[colIdx] ?? '').toString().trim();
            if (!val) continue;
            const key = val.toLowerCase();
            jcSet.add(key);
            if (oldNewIdx !== -1) {
                oldNewMap.set(key, (row[oldNewIdx] ?? '').toString().trim());
            }
        }

        return { jcSet, oldNewMap, tabName: sheet.name, colName: sheet.headers[colIdx], error: null };
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

        if (!state.masterFile) {
            listEl.innerHTML = '<p class="no-files">No file uploaded yet</p>';
            dropZone.classList.remove('has-files');
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
            $('allowanceResultsSection').hidden = true;
            hideProgress();
            clearIssues();
        });
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
    /**
     * @param {string[]} repeatedErrors  Pre-HTML strings from checkRepetitions()
     * @param {string[]} missingNames    Plain-text strings from computeAllowances() calcErrors
     * @param {string[]} jcWarnings      Plain-text: missing SiteID-JC combo messages
     * @param {string[]} generalIssues   Plain-text: fetch failures, count mismatches, etc.
     */
    function showIssues(repeatedErrors, missingNames, jcWarnings, generalIssues) {
        const panel = $('allowanceIssuesPanel');

        if (repeatedErrors.length) {
            $('allowanceRepeatedList').innerHTML = repeatedErrors.map(e => `<li>${e}</li>`).join('');
            $('allowanceRepeatedSection').hidden = false;
        } else {
            $('allowanceRepeatedSection').hidden = true;
        }

        if (missingNames.length) {
            $('allowanceMissingNamesList').innerHTML = missingNames.map(e => `<li>${e}</li>`).join('');
            $('allowanceMissingNamesSection').hidden = false;
        } else {
            $('allowanceMissingNamesSection').hidden = true;
        }

        if (jcWarnings.length) {
            $('allowanceWarningsList').innerHTML = jcWarnings.map(w => `<li>${w}</li>`).join('');
            $('allowanceWarningsSection').hidden = false;
        } else {
            $('allowanceWarningsSection').hidden = true;
        }

        if (generalIssues.length) {
            $('allowanceErrorsList').innerHTML = generalIssues.map(e => `<li>${esc(e)}</li>`).join('');
            $('allowanceErrorsSection').hidden = false;
        } else {
            $('allowanceErrorsSection').hidden = true;
        }

        panel.hidden = !(repeatedErrors.length || missingNames.length || jcWarnings.length || generalIssues.length);
    }

    function clearIssues() {
        $('allowanceIssuesPanel').hidden         = true;
        $('allowanceRepeatedSection').hidden     = true;
        $('allowanceMissingNamesSection').hidden = true;
        $('allowanceWarningsSection').hidden     = true;
        $('allowanceErrorsSection').hidden       = true;
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
                <div class="allowance-fetch-stat allowance-fetch-stat--util">
                    <span class="allowance-fetch-num">${(() => {
                        const team = people.filter(p => p.isTeam);
                        if (!team.length) return '0%';
                        const avg = team.reduce((s, p) => s + p.daysWorked, 0) / team.length / 13 * 100;
                        return Math.ceil(avg) + '%';
                    })()}</span>
                    <span class="allowance-fetch-label">Avg Team Utilization</span>
                </div>
            </div>
        `;

        /* ── Utilization table (team members only, most days first) ── */
        const teamPeople = people
            .filter(p => p.isTeam)
            .sort((a, b) => b.daysWorked - a.daysWorked || a.name.localeCompare(b.name));
        const utilizationRows = teamPeople.map(p => {
            const utilPct = Math.ceil(p.daysWorked / 13 * 100);
            return `
                <tr>
                    <td>${esc(p.name)}</td>
                    <td class="allowance-td-num">${p.daysWorked}</td>
                    <td class="allowance-td-num">${utilPct}%</td>
                </tr>
            `;
        }).join('');

        const personTable = `
            <h3 class="allowance-section-title">Team Utilization</h3>
            <div class="allowance-table-wrap">
                <table class="allowance-table">
                    <thead>
                        <tr>
                            <th>Name</th>
                            <th class="allowance-th-num">Days Worked</th>
                            <th class="allowance-th-num">Utilization</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${utilizationRows || '<tr><td colspan="3" class="allowance-empty">No data</td></tr>'}
                    </tbody>
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

        // Separated issue buckets
        const generalIssues  = [];   // fetch failures, count mismatches, master errors
        const jcWarnings     = [];   // missing SiteID-JC combos
        let   missingNames   = [];   // names not in salary list
        let   repeatedErrors = [];   // same person on same day in multiple rows

        try {
            /* ── Step 1: Read master file ────────────────────── */
            setProgress(5, 'Reading master tracking file…');
            const masterSheets = await FileHandler.readFile(state.masterFile, undefined, 'ID#');

            /* ── Step 1b: Parse master Tracking tab ──────────── */
            setProgress(8, 'Parsing master tracking tab…');
            const { jcSet, oldNewMap, tabName, colName, error: masterError } =
                parseMasterTracking(masterSheets);

            if (masterError) {
                generalIssues.push(masterError);
            } else {
                state.masterJcSet     = jcSet;
                state.masterOldNewMap = oldNewMap;
                console.log(`Master tracking: tab "${tabName}", column "${colName}", ${jcSet.size} combos, ${oldNewMap.size} Old/New entries`);
            }

            /* ── Step 2: Fetch coordinator Google Sheets ─────── */
            setProgress(10, 'Fetching coordinator sheets…');

            const urlList = AppData.getSheetUrls();
            const total   = urlList.length;

            const { rows, warnings: fetchWarnings, sourceCounts } =
                await fetchGoogleSheets((stepIdx, stepTotal, sheetName) => {
                    const pct = total > 0
                        ? 10 + Math.round((stepIdx / stepTotal) * 75)
                        : 10;
                    setProgress(pct, `Fetching ${stepIdx + 1} / ${stepTotal}: ${sheetName}…`);
                });

            generalIssues.push(...fetchWarnings);
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
                generalIssues.push(
                    `No rows matched "${monthName} — ${halfLabel}". ` +
                    `Verify that the Month column uses a 3-letter abbreviation (e.g. "Jan", "Feb") ` +
                    `and Month Half uses "First" or "Second".`
                );
            }

            /* ── Step 3b: Compare SiteID-JC combos against master ── */
            if (jcSet.size > 0 && filteredRows.length > 0) {
                setProgress(90, 'Comparing SiteID-JC combos against master tracking…');

                // lowercase combo → { display: string, sources: Set<string> }
                const missingCombos = new Map();
                // source → [{ site, day }]  for rows that have a site but no JC
                const missingJcRows = new Map();

                for (const row of filteredRows) {
                    const siteRaw = (row.site || '').trim();
                    const jcRaw   = (row.jc   || '').trim();
                    if (!siteRaw) continue;

                    const source = row.__source__ || '(unknown)';

                    if (!jcRaw) {
                        if (!missingJcRows.has(source)) missingJcRows.set(source, []);
                        missingJcRows.get(source).push({ site: siteRaw, day: (row.day || '').trim() });
                        continue;
                    }

                    const sites  = siteRaw.split('/').map(s => s.trim()).filter(Boolean);
                    const jcs    = jcRaw.split('/').map(s => s.trim()).filter(Boolean);

                    // Count mismatch → general issues (data problem in the sheet)
                    if (sites.length !== jcs.length) {
                        const label = [row.__source__, row.day && `Day ${row.day}`]
                            .filter(Boolean).join(', ');
                        generalIssues.push(
                            `Site/JC count mismatch${label ? ` (${label})` : ''}: ` +
                            `${sites.length} site(s) [${siteRaw}] but ` +
                            `${jcs.length} job code(s) [${jcRaw}]`
                        );
                        continue;
                    }

                    for (let i = 0; i < sites.length; i++) {
                        const combo      = `${sites[i]}-${jcs[i]}`;
                        const comboLower = combo.toLowerCase();
                        if (!jcSet.has(comboLower)) {
                            if (!missingCombos.has(comboLower)) {
                                missingCombos.set(comboLower, { display: combo, sources: new Set() });
                            }
                            missingCombos.get(comboLower).sources.add(source);
                        }
                    }
                }

                if (missingCombos.size > 0) {
                    const sorted = [...missingCombos.values()]
                        .sort((a, b) => a.display.localeCompare(b.display));

                    const header = `<span class="rep-header">${missingCombos.size} SiteID-JC combo(s) not found in master tracking ` +
                                   `(tab: &quot;${esc(tabName)}&quot;, column: &quot;${esc(colName)}&quot;):</span>`;

                    const comboEntries = sorted.map(({ display, sources }) => {
                        const srcSpans = [...sources].sort()
                            .map(s => `<span class="rep-entry">${esc(s)}</span>`)
                            .join('');
                        return `<span class="rep-header"><strong>${esc(display)}</strong> — in ` +
                               `${sources.size} sheet${sources.size > 1 ? 's' : ''}:</span> ${srcSpans}`;
                    }).join('');

                    jcWarnings.push(header + comboEntries);
                }

                if (missingJcRows.size > 0) {
                    const total = [...missingJcRows.values()].reduce((s, a) => s + a.length, 0);
                    const header2 = `<span class="rep-header">${total} row${total > 1 ? 's' : ''} ` +
                                    `with a site but <strong>no JC</strong>:</span>`;
                    const sheetEntries = [...missingJcRows.entries()]
                        .sort(([a], [b]) => a.localeCompare(b))
                        .map(([src, rows]) => {
                            const rowSpans = rows
                                .map(r => `<span class="rep-entry">${esc(r.site)}${r.day ? ` / Day ${esc(r.day)}` : ''}</span>`)
                                .join('');
                            return `<span class="rep-header"><strong>${esc(src)}</strong> — ${rows.length} row${rows.length > 1 ? 's' : ''}:</span> ${rowSpans}`;
                        }).join('');
                    jcWarnings.push(header2 + sheetEntries);
                }
            }

            /* ── Step 4: Compute allowances ──────────────────── */
            setProgress(92, 'Computing allowances…');

            const { people, grandTotal, calcWarnings, calcErrors } = computeAllowances(filteredRows);
            generalIssues.push(...calcWarnings);
            missingNames = calcErrors;   // plain-text strings, escaped in showIssues

            /* ── Step 5: Repetition check ────────────────────── */
            setProgress(97, 'Checking for repetitions…');

            repeatedErrors = checkRepetitions(filteredRows);   // pre-HTML strings

            setProgress(100, 'Done!');
            hideProgress();

            state.results = { monthVal, monthName, half, halfLabel, masterSheets, people, grandTotal };
            showResults(monthName, halfLabel, sourceCounts, filteredSourceCounts, rows.length, filteredRows.length, people, grandTotal);
            showIssues(repeatedErrors, missingNames, jcWarnings, generalIssues);

        } catch (err) {
            hideProgress();
            generalIssues.push(err.message);
            showIssues(repeatedErrors, missingNames, jcWarnings, generalIssues);
        }
    }

    /* ── Reset ───────────────────────────────────────────────── */
    function reset() {
        state.masterFile      = null;
        state.sheetRows       = [];
        state.filteredRows    = [];
        state.masterJcSet     = new Set();
        state.masterOldNewMap = new Map();
        state.results         = null;

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
            (files) => {
                state.masterFile = files[0];
                renderMasterFile();
                runAnalysis();
            },
            false
        );

        $('allowanceDownloadBtn').addEventListener('click', generateExcel);
        $('allowanceNewOldBtn').addEventListener('click', generateNewOldFiles);
    }

    return { init, reset };

})();
