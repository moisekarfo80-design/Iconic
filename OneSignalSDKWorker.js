// ICONIC · Service worker : notifications (OneSignal) + ouverture instantanée (cache de l'app)
try { importScripts('https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js'); } catch (e) { /* hors ligne : les notifications se rechargeront plus tard */ }

const CACHE = 'iconic-v3';
const BASE = self.registration.scope;
const FICHIERS = ['', 'index.html', 'manifest.json', 'Image%20PNG.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FICHIERS.map(f => BASE + f))).catch(() => {}).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(k => Promise.all(k.filter(n => n.startsWith('iconic-') && n !== CACHE).map(n => caches.delete(n)))).then(() => self.clients.claim()));
});

// L'app s'ouvre tout de suite depuis le cache, puis se met à jour en arrière-plan.
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  const local = u.href.startsWith(BASE);
  const polices = u.host === 'fonts.googleapis.com' || u.host === 'fonts.gstatic.com';
  if (!local && !polices) return;
  e.respondWith(caches.open(CACHE).then(async c => {
    const enCache = await c.match(r, { ignoreSearch: local });
    const reseau = fetch(r).then(rep => { if (rep && (rep.ok || rep.type === 'opaque')) c.put(r, rep.clone()); return rep; }).catch(() => enCache);
    return enCache || reseau;
  }));
});
