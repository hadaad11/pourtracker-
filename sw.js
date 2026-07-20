/* Pour Tracker service worker — makes the installed PWA open OFFLINE and load instantly.
   Strategy:
   - Cross-origin requests (the live Google Sheets data) are NEVER touched -> always straight to the network.
   - Same-origin navigations (the app shell) are NETWORK-FIRST: always fetch the latest build when online
     (so the built-in self-update keeps working), and fall back to the cached shell only when offline.
   - Static assets (icons/manifest) are cache-first for speed.
   Bump CACHE only when this file itself changes. */
const CACHE = 'pourtracker-v1';
const SHELL = [
  './index.html', './manifest.webmanifest',
  './icons/icon-192.png', './icons/icon-512.png',
  './icons/favicon-32.png', './icons/apple-touch-icon.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(ks => Promise.all(ks.map(k => k === CACHE ? null : caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;   // live data -> let it go to the network untouched

  const isNav = req.mode === 'navigate' || (req.headers.get('accept') || '').includes('text/html');
  if (isNav) {
    // network-first: newest build wins; cache a clean copy under a stable key for the offline fallback
    e.respondWith(
      fetch(req)
        .then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put('./index.html', copy)); return res; })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }
  // static assets: cache-first, refresh in the background
  e.respondWith(
    caches.match(req).then(hit => hit || fetch(req).then(res => {
      const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return res;
    }))
  );
});
