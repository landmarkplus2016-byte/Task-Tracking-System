/**
 * adminSettings.js
 * ──────────────────────────────────────────────────────────────
 * Admin-only "Settings" tab. Lets the admin edit the App List
 * (Google Sheet URLs + Team/Driver salaries) that used to live in
 * list.xlsx, and save it back to the shared Google Sheet via the
 * AppList Apps Script web app (doPost).
 *
 * - Read data comes from AppData (already loaded on startup).
 * - Salaries are edited as FULL MONTHLY amounts (the raw value);
 *   AppData derives the daily rate (÷26) for the rest of the app.
 * - Saving requires the admin password, which is validated
 *   server-side by the Apps Script — never stored in this file.
 *
 * The tab itself is revealed by a 7-click gesture on the sidebar
 * logo (wired in app.js), then AdminSettings.reveal() is called.
 *
 * Load order: after appData.js.
 */

const AdminSettings = (() => {
    'use strict';

    const $ = id => document.getElementById(id);

    let unlocked = false;   // has the 7-click gesture revealed the tab this session

    /* ── Cell factories ──────────────────────────────────────── */
    function textCell(value, cls) {
        const td = document.createElement('td');
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'settings-input' + (cls ? ' ' + cls : '');
        input.value = value == null ? '' : String(value);
        td.appendChild(input);
        return td;
    }

    function delCell() {
        const td = document.createElement('td');
        td.className = 'settings-col-del';
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'settings-del-btn';
        btn.title = 'Remove this row';
        btn.textContent = '✕';
        btn.addEventListener('click', () => td.parentElement.remove());
        td.appendChild(btn);
        return td;
    }

    function addUrlRow(tbody, row = {}) {
        const tr = document.createElement('tr');
        tr.appendChild(textCell(row.name, 'js-name'));
        tr.appendChild(textCell(row.url, 'js-url'));
        tr.appendChild(delCell());
        tbody.appendChild(tr);
    }

    function addSalaryRow(tbody, row = {}) {
        const tr = document.createElement('tr');
        tr.appendChild(textCell(row.name, 'js-name'));
        tr.appendChild(textCell(row.salary, 'js-salary settings-col-num'));
        tr.appendChild(textCell(row.bankAccount, 'js-bank'));
        tr.appendChild(delCell());
        tbody.appendChild(tr);
    }

    /* ── Render tables from AppData raw state ────────────────── */
    function render() {
        const data = AppData.getRawData();

        const urlsBody   = $('settingsUrlsTable').querySelector('tbody');
        const teamBody   = $('settingsTeamTable').querySelector('tbody');
        const driverBody = $('settingsDriverTable').querySelector('tbody');

        urlsBody.innerHTML = '';
        teamBody.innerHTML = '';
        driverBody.innerHTML = '';

        (data.sheetUrls || []).forEach(r => addUrlRow(urlsBody, r));
        (data.salaries.team || []).forEach(r => addSalaryRow(teamBody, r));
        (data.salaries.drivers || []).forEach(r => addSalaryRow(driverBody, r));

        setStatus('', '');
    }

    /* ── Collect edited tables back into a raw payload ───────── */
    function collectUrls() {
        const rows = $('settingsUrlsTable').querySelectorAll('tbody tr');
        return Array.from(rows).map(tr => ({
            name: tr.querySelector('.js-name').value.trim(),
            url:  tr.querySelector('.js-url').value.trim(),
        })).filter(r => r.name || r.url);
    }

    function collectSalaries(tableId) {
        const rows = $(tableId).querySelectorAll('tbody tr');
        return Array.from(rows).map(tr => ({
            name:        tr.querySelector('.js-name').value.trim(),
            salary:      tr.querySelector('.js-salary').value.trim(),
            bankAccount: tr.querySelector('.js-bank').value.trim(),
        })).filter(r => r.name);
    }

    function collectData() {
        return {
            sheetUrls: collectUrls(),
            salaries: {
                team:    collectSalaries('settingsTeamTable'),
                drivers: collectSalaries('settingsDriverTable'),
            },
        };
    }

    /* ── Status line ─────────────────────────────────────────── */
    function setStatus(msg, kind) {
        const el = $('settingsStatus');
        if (!el) return;
        el.textContent = msg;
        el.className = 'settings-status' + (kind ? ' settings-status--' + kind : '');
    }

    /* ── Save (POST to Apps Script) ──────────────────────────── */
    async function save() {
        const password = $('settingsPassword').value;
        if (!password) {
            setStatus('Enter the admin password to save.', 'error');
            $('settingsPassword').focus();
            return;
        }

        const saveBtn = $('settingsSaveBtn');
        saveBtn.disabled = true;
        setStatus('Saving…', 'pending');

        const payload = JSON.stringify({ password, data: collectData() });

        try {
            // text/plain avoids a CORS preflight (Apps Script can't answer OPTIONS).
            const res = await fetch(AppData.getEndpoint(), {
                method:  'POST',
                cache:   'no-store',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body:    payload,
            });
            const result = await res.json();

            if (!result || !result.ok) {
                setStatus((result && result.error) || 'Save failed.', 'error');
                return;
            }

            // Server echoes the saved state — adopt it so the rest of the
            // app (Allowance Checker) uses the new data immediately.
            AppData.setData(result.data);
            $('settingsPassword').value = '';
            render();
            setStatus('✓ Saved. All users get the update on their next load.', 'success');

        } catch (err) {
            setStatus(`Save failed: ${err.message}`, 'error');
        } finally {
            saveBtn.disabled = false;
        }
    }

    /* ── Reveal / activate the tab ───────────────────────────── */
    function reveal() {
        const btn = $('tabBtnSettings');
        if (!btn) return;
        unlocked = true;
        btn.hidden = false;
        render();
        btn.click();   // switch to the Settings tab
    }

    /* ── Init ────────────────────────────────────────────────── */
    function init() {
        $('settingsAddUrl').addEventListener('click', () =>
            addUrlRow($('settingsUrlsTable').querySelector('tbody')));
        $('settingsAddTeam').addEventListener('click', () =>
            addSalaryRow($('settingsTeamTable').querySelector('tbody')));
        $('settingsAddDriver').addEventListener('click', () =>
            addSalaryRow($('settingsDriverTable').querySelector('tbody')));

        $('settingsSaveBtn').addEventListener('click', save);
        $('settingsReloadBtn').addEventListener('click', async () => {
            setStatus('Reloading…', 'pending');
            const ok = await AppData.reload();
            render();
            setStatus(ok ? 'Reloaded from server.' : 'Reload failed.', ok ? 'success' : 'error');
        });
    }

    return { init, reveal, isUnlocked: () => unlocked };

})();
