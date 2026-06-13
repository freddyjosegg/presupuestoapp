const CACHE_NAME = 'presupuestoapp-v10';
const ASSETS = [
    './',
    'index.html',
    'style.css',
    'main.js',
    'worker.js',
    'xlsx.full.min.js',
    'icon-512.png',
    'manifest.json'
];

// Install Event
self.addEventListener('install', (e) => {
    e.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            console.log('[Service Worker] Caching app shell');
            return cache.addAll(ASSETS);
        }).then(() => self.skipWaiting())
    );
});

// Activate Event
self.addEventListener('activate', (e) => {
    e.waitUntil(
        caches.keys().then((keys) => {
            return Promise.all(
                keys.map((key) => {
                    if (key !== CACHE_NAME) {
                        console.log('[Service Worker] Removing old cache', key);
                        return caches.delete(key);
                    }
                })
            );
        }).then(() => self.clients.claim())
    );
});

// Fetch Event
self.addEventListener('fetch', (e) => {
    const url = new URL(e.request.url);

    // Skip caching for API requests and Excel file downloads to let the app handle updates
    if (url.pathname.endsWith('.xlsm') || url.pathname.endsWith('.xlsx') || url.hostname.includes('dolarapi.com')) {
        e.respondWith(fetch(e.request));
        return;
    }

    // Use Stale-While-Revalidate strategy for internal static assets
    e.respondWith(
        caches.match(e.request).then((cachedResponse) => {
            if (cachedResponse) {
                // Fetch fresh version in background
                fetch(e.request).then((networkResponse) => {
                    if (networkResponse.status === 200) {
                        caches.open(CACHE_NAME).then((cache) => {
                            cache.put(e.request, networkResponse);
                        });
                    }
                }).catch(() => {/* Ignore network errors offline */});
                
                return cachedResponse;
            }

            return fetch(e.request).then((networkResponse) => {
                // If it's a valid static resource, cache it dynamically
                if (networkResponse.status === 200 && e.request.method === 'GET') {
                    const responseClone = networkResponse.clone();
                    caches.open(CACHE_NAME).then((cache) => {
                        cache.put(e.request, responseClone);
                    });
                }
                return networkResponse;
            });
        })
    );
});
