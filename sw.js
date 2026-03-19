/**
 * sw.js — Service Worker for Task Tracking System
 * Caches static assets for offline use.
 */

const CACHE_NAME = 'task-tracker-v1.8';

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

self.addEventListener('fetch', (event) => {
    // Cache-first for same-origin assets; network-first for everything else
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
