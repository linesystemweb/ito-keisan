/* 内職伝票 — オフラインで開けるようにするための小さな仕掛け（v2）
   ------------------------------------------------------------
   ページを一度でも開いていれば、圏外でも起動できます。
   通信できるときは必ず新しい方を取りに行きます。
   v1 では端末のHTTPキャッシュが古いHTMLを返すことがあったため、
   HTML だけは cache:'reload' で必ず取り直します。 */
const CACHE = 'naishoku-v2';

self.addEventListener('install', e => self.skipWaiting());

self.addEventListener('activate', e => e.waitUntil((async () => {
  // 古いキャッシュを捨てる
  const keys = await caches.keys();
  await Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)));
  await self.clients.claim();
})()));

self.addEventListener('message', e => { if (e.data === 'skipWaiting') self.skipWaiting(); });

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                 // 保管サーバへの送信には触らない
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;       // CDN などは触らない

  // HTML（ページそのもの）は、端末のキャッシュを無視して取り直す
  const isDoc = req.mode === 'navigate' || url.pathname.endsWith('.html');
  const go = isDoc ? fetch(new Request(url.href, { cache: 'reload', credentials: 'same-origin' }))
                   : fetch(req);

  e.respondWith(
    go.then(res => {
      if (res && res.status === 200) {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy)).catch(() => {});
      }
      return res;
    })
    .catch(() => caches.match(req).then(m => m || caches.match('naishoku-test.html')))
  );
});
