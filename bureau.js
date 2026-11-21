/* Mode « programme d'ordinateur » : actif seulement quand l'application est servie par
   serveur.js sur cet ordinateur (127.0.0.1). Les données viennent alors directement du
   classeur Excel du dossier choisi, relu dès qu'il est enregistré de nouveau.
   Sur Internet (GitHub Pages), ce fichier ne fait rien. */
(function () {
  'use strict';
  var h = location.hostname;
  if (h !== '127.0.0.1' && h !== 'localhost') return;
  var moi = document.currentScript, type = moi && moi.getAttribute('data-classeur');
  if (!type) return;
  var CLE = 'mve-bureau-' + type, enCours = false, actif = false;
  function lsGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* rien */ } }
  function $(s) { return document.querySelector(s); }
  function cacher(sel) { var e = $(sel); if (e) { e.hidden = true; e.style.display = 'none'; } }
  function vide() { var e = $('#empty'); return !!e && !e.hidden; }
  function message(t) {
    var e = $('#empty'); if (!e) return;
    var p = e.querySelector('.bureau-msg');
    if (!p) { p = document.createElement('p'); p.className = 'bureau-msg'; p.style.cssText = 'font-weight:600'; e.insertBefore(p, e.querySelector('.btn')); }
    p.textContent = t;
  }
  function preparer() {
    document.documentElement.classList.add('bureau');
    ['#btn-auto-empty', '#btn-auto', '#btn-sync', '#btn-sync-menu', '#btn-auto-off', '#btn-install'].forEach(cacher);
    // Lien vers l'accueil du programme, en tête du menu
    var pile = $('#sheet-menu .stack');
    if (pile && !pile.querySelector('.bureau-accueil')) {
      var a = document.createElement('a'); a.className = 'btn bureau-accueil'; a.href = '/accueil/'; a.textContent = 'Accueil : tous les tableaux de bord';
      pile.insertBefore(a, pile.firstChild);
    }
  }
  var dernierEtat = null;
  function infoMenu() {
    var m = $('#menu-info'), c = dernierEtat && dernierEtat.classeurs[type]; if (!m) return;
    m.textContent = c ? 'Classeur « ' + c.nom + ' », enregistré le ' + new Date(c.mtime).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) +
      ', dans le dossier ' + dernierEtat.dossier + '. Il est relu automatiquement à chaque nouvel enregistrement.' : 'Aucun classeur de ce tableau de bord dans le dossier ' + (dernierEtat ? dernierEtat.dossier : '') + '.';
  }
  function toast(t) { var e = $('#toast'); if (!e) return; e.textContent = t; e.hidden = false; setTimeout(function () { e.hidden = true; }, 4000); }
  function bouton() {
    // Bouton « Recharger maintenant » dans l'en-tête, à côté du menu
    var menu = $('#btn-menu'); if (!menu || $('#btn-bureau-sync')) return;
    var b = document.createElement('button'); b.className = 'icon-btn'; b.id = 'btn-bureau-sync';
    b.setAttribute('aria-label', 'Recharger le classeur maintenant'); b.title = 'Recharger le classeur maintenant';
    b.innerHTML = '<svg width="21" height="21" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 11a8 8 0 0 0-14.3-4.9L4 8"/><path d="M4 3v5h5"/><path d="M4 13a8 8 0 0 0 14.3 4.9L20 16"/><path d="M20 21v-5h-5"/></svg>';
    b.addEventListener('click', function () { verifier(true); });
    menu.parentNode.insertBefore(b, menu);
    var pile = $('#sheet-menu .stack');
    if (pile && !pile.querySelector('.bureau-recharger')) {
      var m = document.createElement('button'); m.className = 'btn ghost bureau-recharger'; m.type = 'button'; m.textContent = 'Recharger le classeur maintenant';
      m.addEventListener('click', function () { history.back(); verifier(true); });
      var acc = pile.querySelector('.bureau-accueil'); pile.insertBefore(m, acc ? acc.nextSibling : pile.firstChild);
    }
  }
  function tourne(on) { var b = $('#btn-bureau-sync'); if (b) b.classList.toggle('spin', on); }
  function verifier(force) {
    if (enCours) return;
    enCours = true; if (force) tourne(true);
    fetch('/bureau/etat', { cache: 'no-store' }).then(function (r) { if (!r.ok) throw new Error('hors programme'); return r.json(); }).then(function (e) {
      if (!e.programme) throw new Error('hors programme');
      dernierEtat = e;
      if (!actif) { actif = true; preparer(); bouton(); var bm = $('#btn-menu'); if (bm) bm.addEventListener('click', infoMenu); }
      var c = e.classeurs[type];
      if (!c) { if (force) toast('Aucun classeur de ce tableau de bord dans le dossier choisi.'); if (vide()) message('Aucun classeur de ce tableau de bord dans le dossier « ' + e.dossier + ' ». Ajoutez-le dans ce dossier ou changez de dossier depuis l’accueil.'); return; }
      var deja = lsGet(CLE);
      if (!force && deja && deja.mtime === c.mtime && deja.nom === c.nom && !vide()) return; // déjà à jour
      if (typeof window.MVE_lireClasseur !== 'function') return;
      return fetch('/bureau/classeur/' + encodeURIComponent(type), { cache: 'no-store' }).then(function (r) {
        if (!r.ok) return r.json().then(function (j) { throw new Error(j.erreur || 'Lecture impossible'); });
        return r.arrayBuffer();
      }).then(function (b) {
        return Promise.resolve(window.MVE_lireClasseur(new Uint8Array(b), c.nom, !vide())).then(function () { lsSet(CLE, { nom: c.nom, mtime: c.mtime }); });
      });
    }).catch(function (err) {
      if (actif && err && err.message !== 'hors programme') { if (vide()) message(err.message); else if (force) toast(err.message); }
      else if (force) toast('Le programme ne répond pas : relancez « Tableaux de bord MVE » depuis le bureau.');
    }).then(function () { enCours = false; tourne(false); });
  }
  function demarrer() { setTimeout(function () { verifier(false); }, 700); }
  if (document.readyState === 'complete') demarrer(); else window.addEventListener('load', demarrer);
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible') verifier(false); });
  setInterval(function () { if (document.visibilityState === 'visible') verifier(false); else fetch('/bureau/ping').catch(function () {}); }, 60000);
})();
