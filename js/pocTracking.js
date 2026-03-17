/**
 * pocTracking.js
 * ──────────────────────────────────────────────────────────────
 * POC Tracking Update tab.
 * Same logic as app.js but uses "Job Code" as the unique identifier.
 * Duplicate Job Codes are highlighted prominently in the results.
 */

const PocTracking = (() => {
    'use strict';

    /* ── Fixed Configuration ──────────────────────────────────── */
    const ID_COLUMN        = 'Job Code';
    const MASTER_SHEET     = 'POC3 Tracking';
    const CASE_SENSITIVE   = false;
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
        const coordinatorFileList = $('pocCoordinatorFileList');
        const coordinatorDropZone = $('pocCoordinatorDropZone');

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
        const masterFileList  = $('pocMasterFileList');
        const masterDropZone  = $('pocMasterDropZone');

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
                <button class="file-remove" id="pocRemoveMasterBtn" title="Remove this file">✕</button>
            </div>
        `;

        $('pocRemoveMasterBtn').addEventListener('click', () => {
            state.masterFile = null;
            renderMasterFile();
            scheduleAutoProcess();
        });
    }

    /* ── Auto-trigger ─────────────────────────────────────────── */
    let autoTimer = null;

    function scheduleAutoProcess() {
        if (autoTimer) clearTimeout(autoTimer);
        if (state.coordinatorFiles.length === 0 || !state.masterFile) return;
        autoTimer = setTimeout(() => runProcess(), 600);
    }

    /* ── Card Upload Progress Flash ───────────────────────────── */
    function flashCardBar(progressEl, barEl) {
        progressEl.hidden = false;
        barEl.style.width = '0%';
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
        $('pocProgressBar').style.width = pct + '%';
        $('pocProgressText').textContent = text;
    }

    function showProgress() {
        $('pocProgressSection').hidden = false;
        setProgress(0, 'Starting…');
    }

    function hideProgress() {
        $('pocProgressSection').hidden = true;
    }

    /* ── Main Process ─────────────────────────────────────────── */
    async function runProcess() {
        if (state.coordinatorFiles.length === 0 || !state.masterFile) return;

        $('pocResultsSection').hidden = true;
        $('pocWarningsPanel').hidden  = true;
        showProgress();

        try {

            /* ── Step 1: Read coordinator files ─────────────────── */
            setProgress(8, 'Reading coordinator files…');

            const coordinatorSheetsData = [];

            for (let i = 0; i < state.coordinatorFiles.length; i++) {
                const file = state.coordinatorFiles[i];

                try {
                    const sheets = await FileHandler.readFile(file, undefined, ID_COLUMN);
                    const { sheet, autoSelected } = findSheetWithId(sheets, 0, ID_COLUMN);

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
            if (coordinatorCombined.duplicates.length > 0) {
                console.warn(`Duplicate Job Codes found: ${coordinatorCombined.duplicates.length}`);
            }

            /* ── Step 4: Parse master ─────────────────────────────── */
            setProgress(70, 'Parsing master tracking…');

            const masterParsed = Comparison.parseMasterData(masterSheet, ID_COLUMN);

            if (masterParsed.error) throw new Error(masterParsed.error);

            console.log(`Master: ${masterParsed.rows.size} entries`);

            /* ── Step 5: Compare ─────────────────────────────────── */
            setProgress(84, 'Comparing…');

            const results = Comparison.compare(coordinatorCombined, masterParsed, CASE_SENSITIVE);

            console.log(`New: ${results.newEntries.length}  Changed: ${results.changedEntries.length}  Unchanged: ${results.unchangedEntries.length}`);

            setProgress(100, 'Done!');

            state.coordinatorCombined = coordinatorCombined;
            state.masterParsed        = masterParsed;
            state.results             = results;

            showResults(results, coordinatorCombined);

        } catch (err) {
            console.error(err);
            hideProgress();
            alert(`Processing failed:\n\n${err.message}`);
        }
    }

    /* ── Results Display ──────────────────────────────────────── */
    function showResults(results, coordinatorCombined) {
        $('pocNewCount').textContent       = results.newEntries.length;
        $('pocChangedCount').textContent   = results.changedEntries.length;
        $('pocUnchangedCount').textContent = results.unchangedEntries.length;
        $('pocTotalCount').textContent     = coordinatorCombined.rows.size;

        // ── Duplicate Job Codes panel ────────────────────────────
        // Since Job Code is the identifier, any duplicate means the same
        // Job Code appeared more than once across the coordinator files.
        const duplicates = coordinatorCombined.duplicates || [];
        const dupPanel   = $('pocDuplicatesPanel');

        if (duplicates.length > 0) {
            $('pocDuplicatesTitle').textContent =
                `🔴 Duplicate Job Codes — ${duplicates.length} duplicate${duplicates.length > 1 ? 's' : ''} detected (only last entry kept)`;

            $('pocDuplicatesList').innerHTML = duplicates
                .map(msg => `<li>${escHtml(msg)}</li>`)
                .join('');

            dupPanel.hidden = false;
        } else {
            dupPanel.hidden = true;
        }

        // ── Warnings panel (errors only) ─────────────────────────
        const errors = coordinatorCombined.errors || [];

        if (errors.length > 0) {
            $('pocWarningsList').innerHTML = errors
                .map(w => `<li>${escHtml(w)}</li>`)
                .join('');
            $('pocWarningsPanel').hidden = false;
        } else {
            $('pocWarningsPanel').hidden = true;
        }

        $('pocResultsSection').hidden = false;
        $('pocResultsSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    /* ── Initialise ───────────────────────────────────────────── */
    function init() {
        const coordinatorDropZone = $('pocCoordinatorDropZone');
        const coordinatorInput    = $('pocCoordinatorInput');
        const masterDropZone      = $('pocMasterDropZone');
        const masterInput         = $('pocMasterInput');

        FileHandler.setupDropZone(coordinatorDropZone, coordinatorInput, (files) => {
            files.forEach(f => {
                if (!state.coordinatorFiles.find(e => e.name === f.name && e.size === f.size)) {
                    state.coordinatorFiles.push(f);
                }
            });
            renderCoordinatorFiles();
            flashCardBar($('pocCoordinatorProgress'), $('pocCoordinatorBar'));
            scheduleAutoProcess();
        }, true);

        FileHandler.setupDropZone(masterDropZone, masterInput, (files) => {
            state.masterFile = files[0];
            renderMasterFile();
            flashCardBar($('pocMasterProgress'), $('pocMasterBar'));
            scheduleAutoProcess();
        }, false);

        /* ── Download ───────────────────────────────────────────── */
        $('pocDownloadBtn').addEventListener('click', () => {
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

                ExcelExport.download(wb, `POC_Tracking_Report_${datePart}.xlsx`);

            } catch (err) {
                console.error(err);
                alert(`Download failed:\n\n${err.message}`);
            }
        });
    }

    return { init };

})();
