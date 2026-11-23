/**
 * FCMS Service Worker (PWA Offline Shell & Performance Acceleration)
 * Strategy:
 * - Static Assets (HTML, CSS, JS, Fonts, Icons): Cache-First / Stale-While-Revalidate
 * - Dynamic API Requests (/api/*): Network-Only (No sensitive data persisted offline without credentials)
 */

const CACHE_NAME = 'fcms-shell-v20261120-r2';
const STATIC_ASSETS = [
  '/',
  '/index.html',
  '/manifest.json',
  '/favicon.svg',
  '/icon-192.png',
  '/icon-512.png',
  '/icon-maskable-192.png',
  '/icon-maskable-512.png',
  '/assets/style.css?v=20261120r1',
  '/assets/responsive.css?v=20261120r1'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(STATIC_ASSETS).catch((err) => {
        console.warn('[SW] Some static assets could not be precached during install:', err);
      });
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

function isBypassCacheRequest(request, url) {
  if (request.method !== 'GET') return true;
  if (url.pathname.startsWith('/api')) return true;
  if (request.headers && (request.headers.has('Authorization') || request.headers.has('authorization'))) return true;
  return false;
}

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // 0. Bypass cross-origin requests (CDNs, Google Fonts, external Worker API) completely
  // so browser fetches them natively under document CSP without SW connect-src 503 blocks.
  if (url.origin !== self.location.origin) {
    return;
  }

  // 1. Strictly Network-Only for same-origin /api, auth-header, and non-GET requests
  if (isBypassCacheRequest(event.request, url)) {
    event.respondWith(
      fetch(event.request).catch(() => {
        return new Response(
          JSON.stringify({
            success: false,
            error: 'Offline: Koneksi terputus, permintaan tidak dapat diproses tanpa jaringan.'
          }),
          {
            status: 503,
            statusText: 'Service Unavailable',
            headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }
          }
        );
      })
    );
    return;
  }

  // 2. Static assets & Shell: Stale-While-Revalidate (excluding no-store / private responses)
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      const fetchPromise = fetch(event.request).then((networkResponse) => {
        const cacheControl = (networkResponse && networkResponse.headers && networkResponse.headers.get('Cache-Control')) || '';
        const allowCache =
          networkResponse &&
          networkResponse.status === 200 &&
          (networkResponse.type === 'basic' || networkResponse.type === 'default') &&
          !/\b(no-store|private)\b/i.test(cacheControl);

        if (allowCache) {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      }).catch(async () => {
        if (event.request.mode === 'navigate') {
          const fallback = await caches.match('/index.html');
          if (fallback) return fallback;
        }
        return cachedResponse || new Response('Offline', { status: 503, statusText: 'Service Unavailable' });
      });

      return cachedResponse || fetchPromise;
    })
  );
});
