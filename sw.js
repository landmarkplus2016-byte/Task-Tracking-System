/**
 * sw.js — Service Worker for Documents Control System
 * Caches static assets for offline use.
 */

const CACHE_NAME = 'task-tracker-v2.173';

const STATIC_ASSETS = [
    './',
    './index.html',
    './css/styles.css',
    './js/fileHandler.js',
    './js/appData.js',
    './js/comparison.js',
    './js/excelExport.js',
    './js/siteIdJc.js',
    './js/pocTracking.js',
    './js/allowanceChecker.js',
    './js/app.js',
    './manifest.json',
    './icons/icon.svg',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return cache.addAll(STATIC_ASSETS).catch(() => {
                // Non-fatal: some CDN resources may fail in offline install
            });
        })
    );
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((keys) =>
            Promise.all(
                keys
                    .filter((key) => key !== CACHE_NAME)
                    .map((key) => caches.delete(key))
            )
        )
    );
    self.clients.claim();
});

// Only cache same-origin assets + explicitly whitelisted external URLs (e.g. CDN).
// External dynamic URLs (Google Sheets exports, etc.) are never intercepted so
// they always go straight to the network and are never stale.
const STATIC_ASSET_URLS = new Set(STATIC_ASSETS);

self.addEventListener('fetch', (event) => {
    const url = event.request.url;
    const isSameOrigin = url.startsWith(self.location.origin);
    const isWhitelistedExternal = STATIC_ASSET_URLS.has(url);

    if (!isSameOrigin && !isWhitelistedExternal) return; // pass through unmodified

    event.respondWith(
        caches.match(event.request).then((cached) => {
            if (cached) return cached;
            return fetch(event.request).then((response) => {
                if (response && response.status === 200 && response.type !== 'opaque') {
                    const clone = response.clone();
                    caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
                }
                return response;
            }).catch(() => cached);
        })
    );
});
