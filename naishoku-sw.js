/* 内職伝票 — オフラインで開けるようにするための小さな仕掛け
   ------------------------------------------------------------
   ページを一度でも開いていれば、圏外でも起動できます。
   通信できるときは必ず新しい方を取りに行くので、
   HTMLを差し替えたら次に開いたときに反映されます。 */
const CACHE = 'naishoku-v1';

self.addEventListener('install',  e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // 保管サーバへの送信には触らない
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;             // CDN などは触らない

  e.respondWith(
    fetch(req)
      .then(res => {
        if (res && res.status === 200) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
        }
        return res;
      })
      .catch(() => caches.match(req).then(m => m || caches.match('naishoku-test.html')))
  );
});
