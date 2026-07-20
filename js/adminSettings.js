/**
 * adminSettings.js
 * ──────────────────────────────────────────────────────────────
 * Admin-only "Settings" tab. Lets the admin edit the App List
 * (Google Sheet URLs + Team/Driver salaries) that used to live in
 * list.xlsx, and save it back to the shared Google Sheet via the
 * AppList Apps Script web app (doPost).
 *
 * Access flow:
 *   1. The Settings tab is always visible in the sidebar; opening it
 *      shows a password lock screen (the tables are hidden).
 *   2. The admin enters the password → it is verified SERVER-SIDE
 *      (doPost action:'login'). Only on success are the tables shown,
 *      populated from the login response (so there is no dependency on
 *      the startup fetch having finished).
 *   3. Saving reuses the verified password for that session.
 *
 * The password is never stored in this file — it is validated by the
 * Apps Script. Unlock is session-only: reopening the app returns to the
 * lock screen and re-requires the password.
 *
 * Salaries are edited as FULL MONTHLY amounts (the raw value);
 * AppData derives the daily rate (÷26) for the rest of the app.
 *
 * Load order: after appData.js.
 */

const AdminSettings = (() => {
    'use strict';

    const $ = id => document.getElementById(id);

    let sessionPassword = null;    // verified password, kept in memory only

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

    /* ── Status lines ────────────────────────────────────────── */
    function setStatus(msg, kind) { setLine('settingsStatus', msg, kind); }
    function setLockStatus(msg, kind) { setLine('settingsLockStatus', msg, kind); }
    function setLine(id, msg, kind) {
        const el = $(id);
        if (!el) return;
        el.textContent = msg;
        el.className = 'settings-status' + (kind ? ' settings-status--' + kind : '');
    }

    /* ── Endpoint POST helper ────────────────────────────────── */
    // text/plain avoids a CORS preflight (Apps Script can't answer OPTIONS).
    async function postJson(body) {
        const res = await fetch(AppData.getEndpoint(), {
            method:  'POST',
            cache:   'no-store',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body:    JSON.stringify(body),
        });
        return res.json();
    }

    /* ── Lock / unlock UI ────────────────────────────────────── */
    function showLock() {
        $('settingsLock').hidden = false;
        $('settingsContent').hidden = true;
        setLockStatus('', '');
    }
    function showContent() {
        $('settingsLock').hidden = true;
        $('settingsContent').hidden = false;
    }

    async function unlock() {
        const pw = $('settingsLockPw').value;
        if (!pw) {
            setLockStatus('Enter the admin password.', 'error');
            $('settingsLockPw').focus();
            return;
        }

        const btn = $('settingsUnlockBtn');
        btn.disabled = true;
        setLockStatus('Verifying…', 'pending');

        try {
            const result = await postJson({ action: 'login', password: pw });
            if (!result || !result.ok) {
                setLockStatus((result && result.error) || 'Wrong password.', 'error');
                return;
            }
            sessionPassword = pw;
            AppData.setData(result.data);   // fresh data straight from the server
            $('settingsLockPw').value = '';
            render();
            showContent();
        } catch (err) {
            setLockStatus(`Could not verify: ${err.message}`, 'error');
        } finally {
            btn.disabled = false;
        }
    }

    function lockNow() {
        sessionPassword = null;
        $('settingsLockPw').value = '';
        showLock();
    }

    /* ── Save (POST to Apps Script) ──────────────────────────── */
    async function save() {
        if (!sessionPassword) { showLock(); return; }

        const saveBtn = $('settingsSaveBtn');
        saveBtn.disabled = true;
        setStatus('Saving…', 'pending');

        try {
            const result = await postJson({ password: sessionPassword, data: collectData() });

            if (!result || !result.ok) {
                // Password may have changed on the server — force re-login.
                if (result && /password/i.test(result.error || '')) {
                    sessionPassword = null;
                    showLock();
                    setLockStatus('Session expired — enter the password again.', 'error');
                    return;
                }
                setStatus((result && result.error) || 'Save failed.', 'error');
                return;
            }

            AppData.setData(result.data);   // adopt saved state app-wide
            render();
            setStatus('✓ Saved. All users get the update on their next load.', 'success');

        } catch (err) {
            setStatus(`Save failed: ${err.message}`, 'error');
        } finally {
            saveBtn.disabled = false;
        }
    }

    /* ── Init ────────────────────────────────────────────────── */
    function init() {
        showLock();   // panel always starts locked

        // Focus the password box whenever the tab is opened while locked.
        const tabBtn = $('tabBtnSettings');
        if (tabBtn) {
            tabBtn.addEventListener('click', () => {
                if (!sessionPassword) setTimeout(() => $('settingsLockPw').focus(), 0);
            });
        }

        $('settingsUnlockBtn').addEventListener('click', unlock);
        $('settingsLockPw').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); unlock(); }
        });

        $('settingsAddUrl').addEventListener('click', () =>
            addUrlRow($('settingsUrlsTable').querySelector('tbody')));
        $('settingsAddTeam').addEventListener('click', () =>
            addSalaryRow($('settingsTeamTable').querySelector('tbody')));
        $('settingsAddDriver').addEventListener('click', () =>
            addSalaryRow($('settingsDriverTable').querySelector('tbody')));

        $('settingsSaveBtn').addEventListener('click', save);
        $('settingsLockBtn').addEventListener('click', lockNow);
        $('settingsReloadBtn').addEventListener('click', async () => {
            setStatus('Reloading…', 'pending');
            const ok = await AppData.reload();
            render();
            setStatus(ok ? 'Reloaded from server.' : 'Reload failed.', ok ? 'success' : 'error');
        });
    }

    return { init };

})();
