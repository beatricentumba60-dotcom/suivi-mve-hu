/* Service worker de « Qualité MVE » : fonctionnement hors connexion.
   Le fichier chiffré qualite.enc est toujours demandé au réseau (jamais mis en cache ici).
   v2 : réseau d'abord pour les fichiers de l'application (une nouvelle version est prise
   dès l'ouverture suivante) ; le cache ne sert que sans Internet. */
var VERSION = 'qualite-mve-v6'; // v4 : contrôles « confirmée et non-cas » et « classée sans date d’analyse »
var SHELL = ['../admin.js', './', 'index.html', 'app-qualite.js', 'engine-qualite.js', '../xlsx.mini.min.js', 'manifest.webmanifest',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/maskable-512.png', 'icons/apple-touch-icon.png'];
self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(VERSION).then(function (c) {
    return c.addAll(SHELL.map(function (u) { return new Request(u, { cache: 'reload' }); }));
  }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k.indexOf('qualite-mve-') === 0 && k !== VERSION; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  if (url.pathname.endsWith('.enc')) return;
  if (url.pathname.indexOf('/bureau/') === 0) return; // programme d'ordinateur : classeurs toujours relus
  e.respondWith(caches.open(VERSION).then(function (c) {
    return fetch(e.request, { cache: 'no-cache' }).then(function (res) {
      if (res && res.ok) c.put(e.request, res.clone());
      return res;
    }).catch(function () {
      return c.match(e.request, { ignoreSearch: true }).then(function (hit) { return hit || c.match('index.html'); });
    });
  }));
});
