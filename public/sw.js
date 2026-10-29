// TAIGA TALK Progressive Web App (PWA) Service Worker
const CACHE_NAME = 'taiga-talk-v2.2';
const PRECACHE_ASSETS = [
  '/',
  '/home',
  '/roadmap',
  '/articles',
  '/test',
  '/app',
  '/contact',
  '/forumflash-theme.css?v=2.2',
  '/home.css?v=2.2',
  '/articles.css?v=2.2',
  '/app.css?v=2.2',
  '/roadmap.css?v=2.2',
  '/dictionary-widget.css?v=2.0',
  '/dictionary-widget.js',
  '/forumflash-engine.js',
  '/vocab-25-topics.js',
  '/taiga-audio.js',
  '/pwa-install.js?v=2.2',
  '/manifest.json',
  '/icon-192.png',
  '/icon-512.png',
  '/favicon.png'
];

// Install: Pre-cache static assets
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(PRECACHE_ASSETS).catch((err) => {
        console.warn('[SW] Pre-cache warning:', err);
      });
    })
  );
});

// Activate: Clean up previous caches

// Message: Listen for SKIP_WAITING from client
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((name) => {
          if (name !== CACHE_NAME) {
            return caches.delete(name);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch: Strategy dispatcher
self.addEventListener('fetch', (event) => {
  const request = event.request;
  const url = new URL(request.url);

  // 1. Never cache non-GET requests, API calls, or auth endpoints
  if (request.method !== 'GET' || url.pathname.startsWith('/api') || url.pathname.startsWith('/auth')) {
    return;
  }

  // 2. HTML navigation: Stale-While-Revalidate (0ms instant cache response + background update)
  if (request.mode === 'navigate' || request.headers.get('accept')?.includes('text/html')) {
    event.respondWith(
      caches.match(request).then((cachedResponse) => {
        const fetchPromise = fetch(request)
          .then((networkResponse) => {
            if (networkResponse && networkResponse.status === 200) {
              const responseClone = networkResponse.clone();
              caches.open(CACHE_NAME).then((cache) => cache.put(request, responseClone));
            }
            return networkResponse;
          })
          .catch(() => cachedResponse || caches.match('/home'));

        // Return immediately from local cache (0ms instant!) if present, updating in background
        return cachedResponse || fetchPromise;
      })
    );
    return;
  }

  // 3. Static assets (JS, CSS, images, fonts): Cache-first, background revalidate
  event.respondWith(
    caches.match(request).then((cachedResponse) => {
      if (cachedResponse) {
        // Asynchronously update cache in background
        fetch(request).then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            caches.open(CACHE_NAME).then((cache) => cache.put(request, networkResponse));
          }
        }).catch(() => {});
        return cachedResponse;
      }

      return fetch(request).then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200) {
          const responseClone = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(request, responseClone));
        }
        return networkResponse;
      });
    })
  );
});
