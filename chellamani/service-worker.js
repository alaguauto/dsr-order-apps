// DSR Order Apps - PWA service worker (added 15 Sep 2026)
//
// Purpose: this file's ONLY real job is to make the app eligible for
// "Add to Home Screen" / install - Chrome and other browsers require a
// registered service worker with a fetch handler before they'll offer the
// install prompt. It is NOT meant to make the app work offline in a big way,
// because the whole point of this app is fresh data (stock, rates, dealer
// outstanding, live Tally sync) - serving old cached data instead of the
// network would be wrong here.
//
// Strategy: network-first, cache as a fallback only.
//   - Every GET request goes to the network first.
//   - If it succeeds, the response is also stashed in the cache (so there's
//     something to fall back to) and returned as normal.
//   - Only if the network fails outright (phone offline, no signal) does it
//     fall back to whatever was last cached, so the app shell can still open
//     instead of showing a browser error page.
//   - POST requests (saving an order to the backend) are left completely
//     alone - never intercepted, never cached. This is checked explicitly
//     below so an order save can never be accidentally served from cache.

const CACHE_NAME = 'dsr-app-shell-v1';
const CORE_ASSET = './';

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.add(CORE_ASSET)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  // Only ever handle simple GET page/asset loads. Everything else (POSTs to
  // the orders API, live-sync polling, etc.) passes straight through
  // untouched - not intercepted at all.
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy)).catch(() => {});
        return response;
      })
      .catch(() =>
        caches.match(event.request).then((cached) => cached || caches.match(CORE_ASSET))
      )
  );
});
