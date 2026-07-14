const CACHE_NAME = 'inventory-app-v3';
const APP_SHELL = [
  './',
  './index.html',
  './core.js?v=20260714a',
  './manifest.webmanifest',
  './icons/icon.svg'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(
      keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
    ))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Apps Script API는 항상 네트워크 우선 (데이터 최신성)
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(fetch(req));
    return;
  }

  if (req.method !== 'GET') {
    event.respondWith(fetch(req));
    return;
  }

  // 문서 요청: 최신 HTML을 우선하고 성공 응답은 다음 오프라인 실행을 위해 갱신한다.
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).then(async (res) => {
        if (res.ok) {
          const cache = await caches.open(CACHE_NAME);
          await cache.put('./index.html', res.clone());
        }
        return res;
      }).catch(() => caches.match('./index.html'))
    );
    return;
  }

  // 정적 리소스: 즉시 캐시를 사용하되 백그라운드에서 최신 파일로 갱신한다.
  const cachedResponse = caches.match(req);
  const networkUpdate = fetch(req).then(async (res) => {
    if (res.ok) {
      const cache = await caches.open(CACHE_NAME);
      await cache.put(req, res.clone());
    }
    return res;
  });
  event.waitUntil(networkUpdate.then(() => undefined, () => undefined));
  event.respondWith(cachedResponse.then((cached) => cached || networkUpdate));
});
