const CACHE_NAME = 'tenkawa-pwa-v1';

// アプリのコア資産（オフライン対応・高速化）
const PRECACHE_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './qr.png',
  './icons/favicon.svg',
  './icons/favicon.ico',
  './icons/favicon-32.png',
  './icons/favicon-16.png',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-192.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png'
];

// インストール時：コアリソースのキャッシュ
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(PRECACHE_ASSETS);
    }).then(() => {
      return self.skipWaiting();
    })
  );
});

// アクティベーション時：古いキャッシュの削除
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
    }).then(() => {
      return self.clients.claim();
    })
  );
});

// リクエスト取得時（ネットワーク & キャッシュ戦略）
self.addEventListener('fetch', (event) => {
  const req = event.request;

  // GETリクエスト以外はそのまま
  if (req.method !== 'GET') {
    return;
  }

  // ナビゲーション（HTMLページ遷移）：Network First (最新表示優先、オフライン時はキャッシュ)
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            const responseClone = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, responseClone));
          }
          return networkResponse;
        })
        .catch(async () => {
          const cachedResponse = await caches.match(req);
          if (cachedResponse) {
            return cachedResponse;
          }
          // fallback to index.html
          return caches.match('./index.html');
        })
    );
    return;
  }

  // 静的リソース・CDN：Stale-While-Revalidate
  event.respondWith(
    caches.match(req).then((cachedResponse) => {
      const fetchPromise = fetch(req).then((networkResponse) => {
        if (networkResponse && (networkResponse.status === 200 || networkResponse.type === 'opaque')) {
          const responseClone = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, responseClone));
        }
        return networkResponse;
      }).catch(() => {
        // オフライン時のフォールバック（エラー抑制）
      });

      return cachedResponse || fetchPromise;
    })
  );
});
