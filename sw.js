const CACHE_NAME = 'nbl-encroachment-v3';
const CORE = ['index.html', 'manifest.json', 'seal-province.png', 'seal-dopa.png', 'nbl_districts.geojson'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(CORE)));
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

// network-first เสมอ: หน้าเว็บและข้อมูลจะสดใหม่ ถ้าออฟไลน์จึงใช้ cache
// ไม่ cache คำขอไปยัง Apps Script (ข้อมูลต้องสดเสมอ)
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  if (e.request.url.indexOf('script.google.com') !== -1) return; // ปล่อยให้ยิง API ตรง ไม่แคช
  e.respondWith(
    fetch(e.request)
      .then(res => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then(c => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
