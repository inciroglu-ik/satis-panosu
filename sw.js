const CACHE = 'satis-panosu-v2';
const ASSETS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png', './logo.png'];

self.addEventListener('install', (e) => {
  // Yeni sürümde varlıkları TAZE (HTTP önbelleğini atlayarak) çek.
  e.waitUntil(caches.open(CACHE).then((c) => Promise.all(
    ASSETS.map((u) => fetch(new Request(u, { cache: 'reload' })).then((r) => c.put(u, r)).catch(() => {}))
  )));
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// HTML (sayfa) her zaman sunucudan DOĞRULANIR (no-cache) → yeni kod anında gelir;
// diğer varlıklar ağ öncelikli, çevrimdışıysa önbelleğe düşer.
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  const url = e.request.url;
  const isDoc = e.request.mode === 'navigate' || url.endsWith('/') || url.endsWith('/index.html');
  const req = isDoc ? new Request(url, { cache: 'no-cache' }) : e.request;
  e.respondWith(
    fetch(req)
      .then((res) => { const c = res.clone(); caches.open(CACHE).then((x) => x.put(e.request, c)); return res; })
      .catch(() => caches.match(e.request).then((r) => r || caches.match('./index.html')))
  );
});
