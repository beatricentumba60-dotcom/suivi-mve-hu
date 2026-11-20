/* Service worker : fonctionnement hors connexion + réception de fichiers partagés.
   Aucune donnée n'est envoyée sur Internet : le fichier partagé est gardé
   dans le cache local du téléphone le temps que l'application le lise. */
var VERSION = 'suivi-mve-v15'; // v15 : onglet Carte (propagation, foyers) ; v14 : population 2026 (classeur) ; v13 : carte des aires et aires actives (carte.js) ; v12 : axe du graphique « Cas par semaine de notification » ; v11 : décès confirmés et non-cas alignés sur DHIS2 ; v8 : lien « Qualité Contacts » ; v7 : lien « Contacts » ; v5 : réseau d'abord (nouvelle version prise dès l'ouverture), cache seulement hors ligne
var SHELL = ['./', 'index.html', 'app.js', 'engine.js', 'carte.js', 'xlsx.mini.min.js', 'manifest.webmanifest',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/maskable-512.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(VERSION).then(function (c) {
    return c.addAll(SHELL.map(function (u) { return new Request(u, { cache: 'reload' }); }));
  }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k.indexOf('suivi-mve-') === 0 && k !== VERSION; })
      .map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  if (e.request.method === 'POST' && url.pathname.endsWith('/partage')) {
    e.respondWith((async function () {
      try {
        var form = await e.request.formData();
        var file = form.get('fichier');
        if (file) {
          var cache = await caches.open('partage');
          await cache.put('fichier-partage', new Response(file, { headers: { 'X-Nom': encodeURIComponent(file.name || 'fichier.xlsx') } }));
        }
      } catch (err) { /* l'application affichera l'absence de fichier */ }
      return Response.redirect('./?partage=1', 303);
    })());
    return;
  }
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  if (url.pathname.endsWith('.enc')) return; // données chiffrées : toujours depuis le réseau
  if (url.pathname.indexOf('/bureau/') === 0) return; // programme d'ordinateur : classeurs toujours relus
  if (url.pathname.indexOf('/qualite/') >= 0) return; // application Qualité : son propre service worker
  if (url.pathname.indexOf('/vaccination/') >= 0) return; // application Vaccination : son propre service worker
  if (url.pathname.indexOf('/contacts/') >= 0) return; // application Contacts : son propre service worker
  if (url.pathname.indexOf('/qualite-contacts/') >= 0) return; // application Qualité Contacts : son propre service worker
  if (url.pathname.indexOf('/chaines/') >= 0) return; // application Chaînes de transmission : son propre service worker
  e.respondWith(fetch(e.request, { cache: 'no-cache' }).then(function (res) {
    if (res && res.ok) { var copy = res.clone(); caches.open(VERSION).then(function (c) { c.put(e.request, copy); }); }
    return res;
  }).catch(function () {
    return caches.match(e.request, { ignoreSearch: true }).then(function (hit) { return hit || caches.match('index.html'); });
  }));
});
