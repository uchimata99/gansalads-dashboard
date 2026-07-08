/* Service Worker מינימלי לאפליקציית הרכש — נועד רק לאפשר התקנה ("הוסף למסך הבית").
   מעביר בקשות ישירות לרשת (בלי מטמון) כדי שעדכונים תמיד יגיעו טריים. */
self.addEventListener('install', (e) => { self.skipWaiting(); });
self.addEventListener('activate', (e) => { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', (e) => { /* pass-through: הדפדפן מטפל כרגיל (רשת) */ });
