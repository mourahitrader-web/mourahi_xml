const CACHE_NAME = 'bigsi-xml-v6';
const APP_SHELL = [
    './',
    './index.html',
    './css/styles.css',
    './js/compatibility.js',
    './js/app.js',
    './vendor/tailwind/tailwindcss.js',
    './vendor/xlsx/xlsx.full.min.js',
    './vendor/fontawesome/all.min.css',
    './vendor/webfonts/fa-solid-900.woff2',
    './vendor/webfonts/fa-regular-400.woff2',
    './vendor/webfonts/fa-brands-400.woff2',
    './vendor/fontawesome/webfonts/fa-solid-900.woff2',
    './vendor/fontawesome/webfonts/fa-regular-400.woff2',
    './vendor/fontawesome/webfonts/fa-brands-400.woff2'
];

self.addEventListener('install', function (event) {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(function (cache) { return cache.addAll(APP_SHELL); })
            .then(function () { return self.skipWaiting(); })
    );
});

self.addEventListener('activate', function (event) {
    event.waitUntil(
        caches.keys().then(function (cacheNames) {
            return Promise.all(cacheNames
                .filter(function (cacheName) { return cacheName !== CACHE_NAME; })
                .map(function (cacheName) { return caches.delete(cacheName); }));
        }).then(function () { return self.clients.claim(); })
    );
});

self.addEventListener('fetch', function (event) {
    if (event.request.method !== 'GET') return;
    event.respondWith(
        caches.match(event.request).then(function (cachedResponse) {
            return cachedResponse || fetch(event.request).then(function (networkResponse) {
                if (new URL(event.request.url).origin === self.location.origin) {
                    const responseCopy = networkResponse.clone();
                    caches.open(CACHE_NAME).then(function (cache) {
                        cache.put(event.request, responseCopy);
                    });
                }
                return networkResponse;
            });
        })
    );
});
