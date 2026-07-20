/**
 * appData.js
 * ──────────────────────────────────────────────────────────────
 * Loads the app's reference list (Google Sheet URLs + salaries)
 * from the AppList Google Apps Script web app and exposes it to
 * other modules. Replaces the old ./list.xlsx workflow.
 *
 * Endpoint returns JSON:
 *   {
 *     sheetUrls: [{ name, url }],
 *     salaries: {
 *       team:    [{ name, salary, bankAccount }],   // salary = FULL MONTHLY
 *       drivers: [{ name, salary, bankAccount }],
 *     }
 *   }
 *
 * The full monthly salary is converted to a daily rate here — once —
 * as round(monthly / DAYS_PER_MONTH). Every downstream consumer keeps
 * reading `dailySalary` exactly as before.
 *
 * If the endpoint is unreachable / unparseable, an error banner
 * (#appDataError) is shown in the UI.
 *
 * Load order: no hard dependency, but kept after fileHandler.js.
 */

const AppData = (() => {
    'use strict';

    // AppList Apps Script web app — read (doGet) / write (doPost).
    const APPLIST_ENDPOINT =
        'https://script.google.com/macros/s/AKfycbzAHl04tWg21Xs3B6fVpAlaBIukhqXO1GN4KX4YABnCBeUGq8E97hS7bWXhazVeCl-4NA/exec';

    // Working days per month — monthly salary ÷ this = daily rate.
    // Consistent with the utilization model (13 working days per half-month).
    const DAYS_PER_MONTH = 26;

    const EMPTY_RAW = { sheetUrls: [], salaries: { team: [], drivers: [] } };

    const state = {
        raw:       EMPTY_RAW,       // exactly as received from the server
        sheetUrls: [],              // [{ name, url }]
        salaries:  { team: [], drivers: [] },   // [{ name, dailySalary, bankAccount }]
    };

    /* ── Monthly → daily ─────────────────────────────────────── */
    function toDaily(monthly) {
        const n = parseFloat((monthly == null ? '' : monthly).toString().replace(/[^\d.]/g, '')) || 0;
        return n > 0 ? Math.round(n / DAYS_PER_MONTH) : 0;
    }

    function deriveSalaries(list) {
        return (list || [])
            .map(s => ({
                name:        (s.name || '').toString().trim(),
                dailySalary: toDaily(s.salary),
                bankAccount: (s.bankAccount || '').toString().trim(),
            }))
            .filter(s => s.name);
    }

    /* ── Apply a raw server payload to state ─────────────────── */
    function applyData(raw) {
        const safe = raw && typeof raw === 'object' ? raw : EMPTY_RAW;
        const sal  = safe.salaries || {};
        state.raw = {
            sheetUrls: Array.isArray(safe.sheetUrls) ? safe.sheetUrls : [],
            salaries: {
                team:    Array.isArray(sal.team)    ? sal.team    : [],
                drivers: Array.isArray(sal.drivers) ? sal.drivers : [],
            },
        };
        state.sheetUrls = state.raw.sheetUrls
            .map(r => ({ name: (r.name || '').toString().trim(), url: (r.url || '').toString().trim() }))
            .filter(r => r.name || r.url);
        state.salaries = {
            team:    deriveSalaries(state.raw.salaries.team),
            drivers: deriveSalaries(state.raw.salaries.drivers),
        };
    }

    /* ── Error banner ────────────────────────────────────────── */
    function showError(messages) {
        const el = document.getElementById('appDataError');
        if (!el) return;
        el.innerHTML = messages
            .map(m => `<p>${m.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</p>`)
            .join('');
        el.hidden = false;
    }

    function clearError() {
        const el = document.getElementById('appDataError');
        if (el) el.hidden = true;
    }

    /* ── Fetch from endpoint ─────────────────────────────────── */
    async function fetchData() {
        const res = await fetch(APPLIST_ENDPOINT, { cache: 'no-store' });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        if (!data || typeof data !== 'object') throw new Error('Malformed response.');
        return data;
    }

    /* ── Public init ─────────────────────────────────────────── */
    async function init() {
        try {
            applyData(await fetchData());
            clearError();
            console.log(
                `AppData loaded — sheets: ${state.sheetUrls.length}, ` +
                `team: ${state.salaries.team.length}, drivers: ${state.salaries.drivers.length}`
            );
            return true;
        } catch (err) {
            showError([`Could not load app data from the server: ${err.message}`]);
            return false;
        }
    }

    // Re-fetch on demand (e.g. after the admin saves changes).
    async function reload() {
        return init();
    }

    /* ── Public API ──────────────────────────────────────────── */
    return {
        init,
        reload,
        getEndpoint:       () => APPLIST_ENDPOINT,
        getSheetUrls:      () => state.sheetUrls,
        getSalaries:       () => state.salaries.team,     // [{ name, dailySalary, bankAccount }]
        getDriverSalaries: () => state.salaries.drivers,  // [{ name, dailySalary, bankAccount }]
        // Raw monthly-salary data for the admin Settings tab (deep copy).
        getRawData:        () => JSON.parse(JSON.stringify(state.raw)),
        // Replace state from a fresh server payload (used after a save).
        setData:           (raw) => { applyData(raw); clearError(); },
    };

})();
