// Service worker לאפליקציית הנוכחות (PWA). מטמין את מעטפת האפליקציה כדי שתיפתח
// מהר וגם ללא רשת. לעולם לא מיירט קריאות ל-Google (התחברות/גיליון) — הן תמיד מהרשת.
const CACHE = 'att-shell-v3';
const ASSETS = ['attendance.html', 'att-icon-192.png', 'att-icon-512.png', 'attendance.webmanifest'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  // גוגל (התחברות + Sheets API) — תמיד רשת, בלי מטמון
  if (u.hostname.indexOf('google') !== -1 || u.hostname.indexOf('gstatic') !== -1) return;
  if (e.request.method !== 'GET') return;
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).then(resp => {
      if (resp && resp.ok && u.origin === location.origin) {
        const cp = resp.clone(); caches.open(CACHE).then(c => c.put(e.request, cp));
      }
      return resp;
    }).catch(() => caches.match('attendance.html')))
  );
});
