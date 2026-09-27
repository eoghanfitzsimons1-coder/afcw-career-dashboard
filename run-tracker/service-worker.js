const CACHE_NAME = 'run-tracker-v2';
const STATIC_ASSETS = [
    'index.html',
    'manifest.json',
    'https://cdn.jsdelivr.net/npm/chart.js',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
];

// Install: pre-cache static assets and activate immediately
self.addEventListener('install', event => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => cache.addAll(STATIC_ASSETS))
    );
});

// Activate: delete old caches and take control of open pages
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(
                keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
            )
        ).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', event => {
    const url = new URL(event.request.url);

    // Never cache API calls — always go straight to the network
    if (url.pathname.startsWith('/api/')) {
        return; // let the browser handle it normally
    }

    // Only handle GET requests for static assets
    if (event.request.method !== 'GET') {
        return;
    }

    // Cache-first for static assets
    event.respondWith(
        caches.match(event.request).then(response => response || fetch(event.request))
    );
});
