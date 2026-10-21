const CACHE = 'simulador-esp-v7';
const ASSETS = ['./', './index.html'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // data.js: sempre busca da rede ignorando cache (CDN e browser)
  if (e.request.url.includes('/data/data.js')) {
    e.respondWith(
      fetch(new Request(e.request.url, { cache: 'no-cache' }))
        .catch(() => caches.match(e.request))
    );
    return;
  }
  e.respondWith(
    fetch(e.request).then(res => {
      if (res && res.status === 200 && e.request.method === 'GET') {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
      }
      return res;
    }).catch(() => caches.match(e.request))
  );
});
