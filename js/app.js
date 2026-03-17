/**
 * app.js
 * ──────────────────────────────────────────────────────────────
 * Main application entry point.
 * Ties together FileHandler, Comparison, and ExcelExport modules.
 * Auto-triggers comparison 600 ms after files are added/removed.
 */

(() => {
    'use strict';

    /* ── Fixed Configuration ──────────────────────────────────── */
    const ID_COLUMN       = 'ID#';
    const MASTER_SHEET    = 'Invoicing Track';
    const CASE_SENSITIVE  = false;
    const INCLUDE_UNCHANGED = false;

    /* ── Application State ────────────────────────────────────── */
    const state = {
        coordinatorFiles:    [],
        masterFile:          null,
        coordinatorCombined: null,
        masterParsed:        null,
        results:             null,
    };

    /* ── Sheet Selection Helper ───────────────────────────────── */
    function findSheetWithId(sheets, preferredIdx, idColumnName, sheetName) {
        if (!sheets || sheets.length === 0) {
            return { sheet: null, sheetIdx: 0, autoSelected: false, noNameMatch: false };
        }

        const idTarget   = idColumnName.toLowerCase().trim();
        const nameFilter = sheetName ? sheetName.toLowerCase().trim() : '';

        function sheetHasId(sheet) {
            if (!sheet || !sheet.headers) return false;
            return sheet.headers.some(h =>
                h.toLowerCase().trim() === idTarget ||
                h.toLowerCase().trim().includes(idTarget)
            );
        }

        if (nameFilter) {
            for (let i = 0; i < sheets.length; i++) {
                if (sheets[i].name.toLowerCase().trim() === nameFilter) {
                    return { sheet: sheets[i], sheetIdx: i, autoSelected: false, noNameMatch: false };
                }
            }
            for (let i = 0; i < sheets.length; i++) {
                if (sheets[i].name.toLowerCase().includes(nameFilter)) {
                    return { sheet: sheets[i], sheetIdx: i, autoSelected: false, noNameMatch: false };
                }
            }
            return { sheet: null, sheetIdx: -1, autoSelected: false, noNameMatch: true };
        }

        const preferred = sheets[preferredIdx] || sheets[0];
        if (sheetHasId(preferred)) {
            return { sheet: preferred, sheetIdx: preferredIdx, autoSelected: false, noNameMatch: false };
        }

        for (let i = 0; i < sheets.length; i++) {
            if (i === preferredIdx) continue;
            if (sheetHasId(sheets[i])) {
                return { sheet: sheets[i], sheetIdx: i, autoSelected: true, noNameMatch: false };
            }
        }

        return { sheet: preferred, sheetIdx: preferredIdx, autoSelected: false, noNameMatch: false };
    }

    /* ── DOM References ───────────────────────────────────────── */
    const $ = id => document.getElementById(id);

    const coordinatorDropZone  = $('coordinatorDropZone');
    const coordinatorInput     = $('coordinatorInput');
    const coordinatorFileList  = $('coordinatorFileList');
    const coordinatorProgress  = $('coordinatorProgress');
    const coordinatorBar       = $('coordinatorBar');

    const masterDropZone       = $('masterDropZone');
    const masterInput          = $('masterInput');
    const masterFileList       = $('masterFileList');
    const masterProgress       = $('masterProgress');
    const masterBar            = $('masterBar');

    const progressSection      = $('progressSection');
    const progressBar          = $('progressBar');
    const progressText         = $('progressText');

    const resultsSection       = $('resultsSection');
    const warningsPanel        = $('warningsPanel');
    const warningsList         = $('warningsList');
    const downloadBtn          = $('downloadBtn');

    /* ── Helpers ──────────────────────────────────────────────── */
    function escHtml(str) {
        return str
            .replace(/&/g,  '&amp;')
            .replace(/</g,  '&lt;')
            .replace(/>/g,  '&gt;')
            .replace(/"/g,  '&quot;');
    }

    /* ── File List Rendering ──────────────────────────────────── */
    function renderCoordinatorFiles() {
        if (state.coordinatorFiles.length === 0) {
            coordinatorFileList.innerHTML = '<p class="no-files">No files uploaded yet</p>';
            coordinatorDropZone.classList.remove('has-files');
            return;
        }

        coordinatorDropZone.classList.add('has-files');
        coordinatorFileList.innerHTML = state.coordinatorFiles.map((f, i) => `
            <div class="file-item">
                <div class="file-item-name">
                    <span>📄</span>
                    <span class="fname" title="${escHtml(f.name)}">${escHtml(f.name)}</span>
                    <span class="file-status">✓</span>
                </div>
                <button class="file-remove" data-index="${i}" title="Remove this file">✕</button>
            </div>
        `).join('');

        coordinatorFileList.querySelectorAll('.file-remove').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = parseInt(e.currentTarget.dataset.index, 10);
                state.coordinatorFiles.splice(idx, 1);
                renderCoordinatorFiles();
                scheduleAutoProcess();
            });
        });
    }

    function renderMasterFile() {
        if (!state.masterFile) {
            masterFileList.innerHTML = '<p class="no-files">No file uploaded yet</p>';
            masterDropZone.classList.remove('has-files');
            return;
        }

        masterDropZone.classList.add('has-files');
        masterFileList.innerHTML = `
            <div class="file-item">
                <div class="file-item-name">
                    <span>📊</span>
                    <span class="fname" title="${escHtml(state.masterFile.name)}">${escHtml(state.masterFile.name)}</span>
                    <span class="file-status">✓</span>
                </div>
                <button class="file-remove" id="removeMasterBtn" title="Remove this file">✕</button>
            </div>
        `;

        $('removeMasterBtn').addEventListener('click', () => {
            state.masterFile = null;
            renderMasterFile();
            scheduleAutoProcess();
        });
    }

    /* ── Tab Switching ────────────────────────────────────────── */
    function initTabs() {
        const buttons    = document.querySelectorAll('.tab-btn');
        const panels     = document.querySelectorAll('.tab-panel');
        const topbarTitle = document.getElementById('topbarTitle');

        buttons.forEach(btn => {
            btn.addEventListener('click', () => {
                const targetId = btn.getAttribute('aria-controls');

                buttons.forEach(b => {
                    b.classList.remove('tab-btn--active');
                    b.setAttribute('aria-selected', 'false');
                });
                panels.forEach(p => {
                    p.hidden = true;
                    p.classList.remove('tab-panel--active');
                });

                btn.classList.add('tab-btn--active');
                btn.setAttribute('aria-selected', 'true');
                const panel = document.getElementById(targetId);
                if (panel) {
                    panel.hidden = false;
                    panel.classList.add('tab-panel--active');
                }

                // Update topbar title to match the active tab label
                if (topbarTitle) {
                    topbarTitle.textContent = btn.textContent.trim();
                }
            });
        });
    }

    /* ── Drop Zone Setup ──────────────────────────────────────── */
    FileHandler.setupDropZone(coordinatorDropZone, coordinatorInput, (files) => {
        files.forEach(f => {
            if (!state.coordinatorFiles.find(e => e.name === f.name && e.size === f.size)) {
                state.coordinatorFiles.push(f);
            }
        });
        renderCoordinatorFiles();
        flashCardBar(coordinatorProgress, coordinatorBar);
        scheduleAutoProcess();
    }, true);

    FileHandler.setupDropZone(masterDropZone, masterInput, (files) => {
        state.masterFile = files[0];
        renderMasterFile();
        flashCardBar(masterProgress, masterBar);
        scheduleAutoProcess();
    }, false);

    /* ── Auto-trigger ─────────────────────────────────────────── */
    let autoTimer = null;

    function scheduleAutoProcess() {
        if (autoTimer) clearTimeout(autoTimer);

        // Need both coordinator files and master file to run
        if (state.coordinatorFiles.length === 0 || !state.masterFile) return;

        autoTimer = setTimeout(() => runProcess(), 600);
    }

    /* ── Card Upload Progress Flash ───────────────────────────── */
    // Shows a quick fill animation on drop, then hides after 800 ms.
    function flashCardBar(progressEl, barEl) {
        progressEl.hidden = false;
        barEl.style.width = '0%';
        // Kick off the animation in the next frame so the transition fires
        requestAnimationFrame(() => {
            requestAnimationFrame(() => {
                barEl.style.width = '100%';
            });
        });
        setTimeout(() => {
            barEl.style.width = '0%';
            progressEl.hidden = true;
        }, 800);
    }

    /* ── Global Progress ──────────────────────────────────────── */
    function setProgress(pct, text) {
        progressBar.style.width = pct + '%';
        progressText.textContent = text;
    }

    function showProgress() {
        progressSection.hidden = false;
        setProgress(0, 'Starting…');
    }

    function hideProgress() {
        progressSection.hidden = true;
    }

    /* ── Main Process ─────────────────────────────────────────── */
    async function runProcess() {
        // Guard: both sides must be loaded
        if (state.coordinatorFiles.length === 0 || !state.masterFile) return;

        resultsSection.hidden = true;
        warningsPanel.hidden  = true;
        showProgress();

        try {

            /* ── Step 1: Read coordinator files ─────────────────── */
            setProgress(8, 'Reading coordinator files…');

            const coordinatorSheetsData = [];

            for (let i = 0; i < state.coordinatorFiles.length; i++) {
                const file = state.coordinatorFiles[i];

                try {
                    const sheets = await FileHandler.readFile(file, undefined, ID_COLUMN);
                    const { sheet, sheetIdx, autoSelected } = findSheetWithId(sheets, 0, ID_COLUMN);

                    if (!sheet || sheet.headers.length === 0) {
                        console.warn(`"${file.name}": no data found on any sheet.`);
                    } else {
                        if (autoSelected) {
                            console.warn(`"${file.name}": auto-selected sheet "${sheet.name}"`);
                        }
                        coordinatorSheetsData.push({ fileName: file.name, sheet });
                        console.log(`✓ "${file.name}" — ${sheet.rows.length} rows [sheet: "${sheet.name}"] [header row: ${sheet.detectedHeaderRow + 1}]`);
                    }
                } catch (err) {
                    console.error(`✗ "${file.name}": ${err.message}`);
                }

                setProgress(
                    8 + Math.round(32 * (i + 1) / state.coordinatorFiles.length),
                    `Reading coordinator files… (${i + 1} / ${state.coordinatorFiles.length})`
                );
            }

            if (coordinatorSheetsData.length === 0) {
                throw new Error('No valid coordinator data could be loaded. Please check your files.');
            }

            /* ── Step 2: Read master file ────────────────────────── */
            setProgress(44, 'Reading master tracking file…');

            const masterSheets = await FileHandler.readFile(state.masterFile, undefined, ID_COLUMN);
            const { sheet: masterSheet, noNameMatch } = findSheetWithId(masterSheets, 0, ID_COLUMN, MASTER_SHEET);

            if (noNameMatch) {
                throw new Error(
                    `Sheet "${MASTER_SHEET}" was not found in the Master Tracking file.\n` +
                    `Available sheets: ${masterSheets.map(s => `"${s.name}"`).join(', ')}`
                );
            }

            if (!masterSheet || masterSheet.headers.length === 0) {
                throw new Error('Master Tracking sheet is empty. Please check the file.');
            }

            console.log(`✓ "${state.masterFile.name}" — ${masterSheet.rows.length} rows [sheet: "${masterSheet.name}"] [header row: ${masterSheet.detectedHeaderRow + 1}]`);

            /* ── Step 3: Combine coordinator sheets ──────────────── */
            setProgress(58, 'Combining coordinator data…');

            const coordinatorCombined = Comparison.combineCoordinatorSheets(
                coordinatorSheetsData, ID_COLUMN, CASE_SENSITIVE
            );

            if (!coordinatorCombined.idColumnNorm) {
                throw new Error(
                    coordinatorCombined.errors.join(' ') ||
                    `ID column "${ID_COLUMN}" not found in coordinator sheets.`
                );
            }

            console.log(`Combined: ${coordinatorCombined.rows.size} unique entries`);

            /* ── Step 4: Parse master ─────────────────────────────── */
            setProgress(70, 'Parsing master tracking…');

            const masterParsed = Comparison.parseMasterData(masterSheet, ID_COLUMN);

            if (masterParsed.error) throw new Error(masterParsed.error);

            console.log(`Master: ${masterParsed.rows.size} entries`);

            /* ── Step 5: Compare ─────────────────────────────────── */
            setProgress(84, 'Comparing…');

            const results = Comparison.compare(coordinatorCombined, masterParsed, CASE_SENSITIVE);

            console.log(`New: ${results.newEntries.length}  Changed: ${results.changedEntries.length}  Unchanged: ${results.unchangedEntries.length}`);

            /* ── Step 6: Job Code duplicate check ────────────────── */
            setProgress(95, 'Checking Job Code conflicts…');
            const jcConflicts = checkJobCodeDuplicates(coordinatorCombined);
            if (jcConflicts.length > 0) {
                console.warn(`JC conflicts found: ${jcConflicts.length}`);
            }

            setProgress(100, 'Done!');

            // Persist for download
            state.coordinatorCombined = coordinatorCombined;
            state.masterParsed        = masterParsed;
            state.results             = results;

            showResults(results, coordinatorCombined, jcConflicts);

        } catch (err) {
            console.error(err);
            hideProgress();
            alert(`Processing failed:\n\n${err.message}`);
        }
    }

    /* ── Job Code Duplicate Check ─────────────────────────────── */

    /** Find a normalised header key by fuzzy terms (exact first, contains second). */
    function findNormKey(normHeaders, terms) {
        const targets = terms.map(t => t.toLowerCase().trim());
        // Pass 1: exact
        let found = normHeaders.find(h => targets.includes(h));
        if (found) return found;
        // Pass 2: contains
        found = normHeaders.find(h => targets.some(t => h.includes(t)));
        return found || null;
    }

    /**
     * Returns an array of conflict objects:
     *   { jc: string, sites: [{ siteId, source }] }
     * A conflict exists when the same JC value appears with more than one Site ID.
     */
    function checkJobCodeDuplicates(coordinatorCombined) {
        const { normHeaders, rows } = coordinatorCombined;

        const siteIdKey = findNormKey(normHeaders, ['site id', 'site_id', 'siteid', 'site']);
        const jcKey     = findNormKey(normHeaders, ['job code', 'job_code', 'jobcode', 'jc#', 'jc']);

        if (!siteIdKey || !jcKey) return [];   // columns not found — skip silently

        // jc (lowercase) → { original, sites: Map<siteId, source> }
        const jcMap = new Map();

        rows.forEach((row) => {
            const siteId = (row[siteIdKey] || '').trim();
            const jc     = (row[jcKey]     || '').trim();
            if (!siteId || !jc) return;

            const key = jc.toLowerCase();
            if (!jcMap.has(key)) jcMap.set(key, { original: jc, sites: new Map() });
            const entry = jcMap.get(key);
            if (!entry.sites.has(siteId)) {
                entry.sites.set(siteId, row['__source__'] || '');
            }
        });

        const conflicts = [];
        jcMap.forEach(({ original, sites }) => {
            if (sites.size > 1) {
                conflicts.push({
                    jc:    original,
                    sites: Array.from(sites.entries())
                               .map(([siteId, source]) => ({ siteId, source }))
                               .sort((a, b) => a.siteId.localeCompare(b.siteId)),
                });
            }
        });

        return conflicts.sort((a, b) => a.jc.localeCompare(b.jc));
    }

    /* ── Results Display ──────────────────────────────────────── */
    function showResults(results, coordinatorCombined, jcConflicts = []) {
        $('newCount').textContent       = results.newEntries.length;
        $('changedCount').textContent   = results.changedEntries.length;
        $('unchangedCount').textContent = results.unchangedEntries.length;
        $('totalCount').textContent     = coordinatorCombined.rows.size;

        // ── JC Conflicts panel ──────────────────────────────────
        const jcPanel = $('jcConflictsPanel');
        if (jcConflicts.length > 0) {
            $('jcConflictsTitle').textContent =
                `🔴 Job Code Conflicts — ${jcConflicts.length} JC${jcConflicts.length > 1 ? 's' : ''} assigned to multiple sites`;

            $('jcConflictsList').innerHTML = jcConflicts.map(({ jc, sites }) => {
                const siteTags = sites.map(({ siteId, source }) => {
                    const src = source ? ` <span style="color:var(--gray-400);font-size:.74rem;">(${escHtml(source)})</span>` : '';
                    return `<span class="jc-site-item">${escHtml(siteId)}${src}</span>`;
                }).join('');
                return `<li>
                    <span class="jc-tag">${escHtml(jc)}</span>
                    <span class="jc-sites">${siteTags}</span>
                </li>`;
            }).join('');

            jcPanel.hidden = false;
        } else {
            jcPanel.hidden = true;
        }

        // ── Warnings panel ──────────────────────────────────────
        const allWarnings = [
            ...(coordinatorCombined.duplicates || []),
            ...(coordinatorCombined.errors     || []),
        ];

        if (allWarnings.length > 0) {
            warningsList.innerHTML = allWarnings
                .map(w => `<li>${escHtml(w)}</li>`)
                .join('');
            warningsPanel.hidden = false;
        } else {
            warningsPanel.hidden = true;
        }

        resultsSection.hidden = false;
        resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    /* ── Reset ────────────────────────────────────────────────── */
    function resetTracking() {
        state.coordinatorFiles    = [];
        state.masterFile          = null;
        state.coordinatorCombined = null;
        state.masterParsed        = null;
        state.results             = null;

        renderCoordinatorFiles();
        renderMasterFile();
        coordinatorInput.value = '';
        masterInput.value      = '';

        resultsSection.hidden            = true;
        progressSection.hidden           = true;
        warningsPanel.hidden             = true;
        $('jcConflictsPanel').hidden     = true;
    }

    $('resetTrackingBtn').addEventListener('click', resetTracking);

    /* ── Initialise ───────────────────────────────────────────── */
    initTabs();
    SiteIdJc.init();
    PocTracking.init();

    /* ── Download ─────────────────────────────────────────────── */
    downloadBtn.addEventListener('click', () => {
        if (!state.results || !state.coordinatorCombined || !state.masterParsed) return;

        try {
            const wb = ExcelExport.generate(
                state.coordinatorCombined,
                state.masterParsed,
                state.results,
                {
                    idColumnName:         ID_COLUMN,
                    coordinatorFileCount: state.coordinatorFiles.length,
                    duplicates:           state.coordinatorCombined.duplicates,
                    errors:               state.coordinatorCombined.errors,
                    includeUnchanged:     INCLUDE_UNCHANGED,
                    caseSensitive:        CASE_SENSITIVE,
                }
            );

            const now      = new Date();
            const datePart = [
                now.getFullYear(),
                String(now.getMonth() + 1).padStart(2, '0'),
                String(now.getDate()).padStart(2, '0'),
            ].join('');

            ExcelExport.download(wb, `Task_Tracking_Report_${datePart}.xlsx`);

        } catch (err) {
            console.error(err);
            alert(`Download failed:\n\n${err.message}`);
        }
    });

})();
