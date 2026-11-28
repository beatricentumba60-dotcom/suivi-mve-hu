/* Interface « Chaînes de transmission MVE ». Calculs : engine-chaines.js.
   Données : classeurs « Qualité des données MVE » (alertes) et « Contacts MVE », lus dans le dossier
   du programme d'ordinateur (relus à chaque enregistrement) ou ouverts à la main. Rien n'est enregistré. */
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };
  var nf = new Intl.NumberFormat('fr-FR'), pf = new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 0 }), df1 = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 });
  var state = { D: null, S: null, f: { zone: '', debut: null, fin: null }, tab: 'synthese', sel: null, noeud: null, gen: '2', nCh: 40, nL3: 40, sources: null, bureau: false, etat: null };

  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function fd(s, court) { if (s == null) return '–'; var d = MVECH.serialToUTC(s); return pad(d.getUTCDate()) + '/' + pad(d.getUTCMonth() + 1) + (court ? '' : '/' + d.getUTCFullYear()); }
  function sToInput(s) { var d = MVECH.serialToUTC(s); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function lsGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* rien */ } }
  var toastT;
  function toast(m, ms) { var t = $('#toast'); t.textContent = m; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(function () { t.hidden = true; }, ms || 3500); }
  function busy(on) { $('#busy').hidden = !on; }
  function de(n) { return /^[aeiouyéèêâîôûh]/i.test(n) ? 'd’' + n : 'de ' + n; }

  /* ---------- Lecture des classeurs ---------- */
  function feuille(buf) {
    var wb = XLSX.read(buf, { type: 'array', sheets: ['Données'], cellDates: false, cellFormula: false, cellHTML: false, dense: false });
    if (!wb.Sheets['Données']) return null;
    return XLSX.utils.sheet_to_json(wb.Sheets['Données'], { header: 1, raw: true, defval: null, blankrows: false });
  }
  function genre(t) { // alertes (avec Nom) ou contacts ?
    var h = (t && t[0] || []).map(MVECH.norm);
    if (h.indexOf('n° epid cas source') >= 0) return 'contacts';
    if (h.indexOf('id alerte') >= 0 && h.indexOf('classification finale') >= 0) return h.indexOf('nom') >= 0 ? 'alertes' : 'alertes-sans-nom';
    return null;
  }
  function calculer(alertes, contacts, meta) { afficher(MVECH.charger(alertes, contacts), meta); }
  // Téléphone : liens déjà calculés par la publication (aucun nom)
  function appliquer(p, meta) { afficher(MVECH.assembler(p), meta); }
  function afficher(D, meta) {
    state.D = D; state.meta = meta;
    var ex = state.f; state.f = { zone: ex.zone, debut: ex.debut, fin: ex.fin };
    $('#empty').hidden = true; $('#tabs').hidden = false; $('#chips').hidden = false;
    rendre(); montrer(state.tab);
    toast(nf.format(D.aretes.filter(function (e) { return !e.ignore; }).length) + ' liens entre cas, ' + nf.format(D.chaines.filter(function (c) { return c.contacts || c.nbSecondaires; }).length) + ' chaînes.');
  }
  function lireFichiers(files) {
    busy(true);
    Promise.all(Array.prototype.map.call(files, function (f) { return f.arrayBuffer().then(function (b) { return { nom: f.name, t: feuille(new Uint8Array(b)) }; }); })).then(function (l) {
      var a = l.find(function (x) { return genre(x.t) === 'alertes'; }), c = l.find(function (x) { return genre(x.t) === 'contacts'; });
      if (l.some(function (x) { return genre(x.t) === 'alertes-sans-nom'; }) && !a) throw new Error('Pour les alertes, ouvrez le classeur « Qualité des données MVE » (celui de « Suivi MVE » n’a pas les noms nécessaires au rapprochement).');
      if (!a || !c) throw new Error('Sélectionnez les deux classeurs ensemble : « Qualité des données MVE » et « Contacts MVE ».');
      calculer(a.t, c.t, { fichiers: [a.nom, c.nom], le: Date.now() });
    }).catch(function (e) { toast('Lecture impossible : ' + (e && e.message ? e.message : e), 7000); }).then(function () { busy(false); });
  }
  /* Programme d'ordinateur : les deux classeurs du dossier choisi */
  var enCours = false;
  function bureau(force) {
    if (enCours) return Promise.resolve();
    enCours = true; if (force) $('#btn-sync').classList.add('spin');
    return fetch('/bureau/etat', { cache: 'no-store' }).then(function (r) { if (!r.ok) throw new Error('hors'); return r.json(); }).then(function (e) {
      if (!e.programme) throw new Error('hors');
      state.bureau = true; state.etat = e; $('#btn-sync').hidden = false; $('#btn-sync-menu').hidden = false; $('#lien-accueil').hidden = false;
      var q = e.classeurs.qualite, c = e.classeurs.contacts;
      if (!q || !c) {
        if (!state.D) $('#empty-msg').textContent = 'Il manque dans le dossier « ' + e.dossier + ' » ' + (!q && !c ? 'les classeurs « Qualité des données MVE » et « Contacts MVE »' : !q ? 'le classeur « Qualité des données MVE »' : 'le classeur « Contacts MVE »') + '. Ajoutez-les ou changez de dossier depuis l’accueil.';
        if (force) toast('Classeur manquant dans le dossier.');
        return;
      }
      var cle = q.nom + '|' + q.mtime + '|' + c.nom + '|' + c.mtime;
      if (!force && state.D && state.sources === cle) return;
      busy(true);
      var get = function (t) { return fetch('/bureau/classeur/' + t, { cache: 'no-store' }).then(function (r) { if (!r.ok) return r.json().then(function (j) { throw new Error(j.erreur); }); return r.arrayBuffer(); }); };
      return Promise.all([get('qualite'), get('contacts')]).then(function (b) {
        var ta = feuille(new Uint8Array(b[0])), tc = feuille(new Uint8Array(b[1]));
        if (genre(ta) !== 'alertes') throw new Error('Le classeur « ' + q.nom + ' » n’a pas la colonne « Nom » nécessaire au rapprochement.');
        calculer(ta, tc, { fichiers: [q.nom, c.nom], le: Date.now() }); state.sources = cle;
      });
    }).catch(function (e) {
      if (e && e.message !== 'hors') toast(e.message, 7000);
    }).then(function () { enCours = false; busy(false); $('#btn-sync').classList.remove('spin'); });
  }

  /* ---------- Stockage local ---------- */
  function idb(name) {
    return new Promise(function (res, rej) {
      var rq = indexedDB.open(name || 'chaines-mve', 1);
      rq.onupgradeneeded = function () { rq.result.createObjectStore('kv'); };
      rq.onsuccess = function () { res(rq.result); }; rq.onerror = function () { rej(rq.error); };
    });
  }
  function kv(mode, fn, name) {
    return idb(name).then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction('kv', mode), out = fn(tx.objectStore('kv'));
        tx.oncomplete = function () { res(out && out.result !== undefined ? out.result : undefined); };
        tx.onerror = function () { rej(tx.error); };
      });
    });
  }
  var store = {
    get: function (k, name) { return kv('readonly', function (s) { return s.get(k); }, name); },
    set: function (k, v) { return kv('readwrite', function (s) { s.put(v, k); }); },
    del: function (k) { return kv('readwrite', function (s) { s.delete(k); }); }
  };

  /* ---------- Mise à jour automatique (chaines.enc, même mot de passe que Suivi MVE) ---------- */
  var REMOTE = 'chaines.enc', sync = { key: null, salt: null, iter: 0, running: false };
  function b64(u8) { var s = ''; for (var i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]); return btoa(s); }
  function err(m, c) { var e = new Error(m); e.code = c; return e; }
  function parseHeader(buf) {
    var u8 = new Uint8Array(buf);
    if (u8.length < 40 || String.fromCharCode(u8[0], u8[1], u8[2], u8[3]) !== 'MVE1') throw err('Le fichier publié est illisible.');
    return { iter: new DataView(buf).getUint32(4), salt: u8.slice(8, 24), iv: u8.slice(24, 36), data: u8.slice(36) };
  }
  function deriveKey(pw, salt, iter) {
    return crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']).then(function (b) {
      return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: salt, iterations: iter, hash: 'SHA-256' }, b, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
    });
  }
  function decrypt(h, key) {
    return crypto.subtle.decrypt({ name: 'AES-GCM', iv: h.iv }, key, h.data).then(function (p) {
      return new Response(new Blob([p]).stream().pipeThrough(new DecompressionStream('gzip'))).text();
    }, function () { throw err('Mot de passe incorrect.', 'cle'); }).then(JSON.parse);
  }
  function fetchRemote() {
    return fetch(REMOTE + '?t=' + Date.now(), { cache: 'no-store' }).then(function (r) {
      if (r.status === 404) throw err('Aucune chaîne de transmission n’a encore été publiée.', 'absent');
      if (!r.ok) throw err('Le site de l’application ne répond pas (erreur ' + r.status + ').');
      return r.arrayBuffer();
    }, function () { throw err('Connexion impossible : les dernières données restent affichées.', 'reseau'); });
  }
  function applyRemote(p, force) {
    var cur = state.meta;
    if (!force && cur && cur.publie && p.publie && p.publie <= cur.publie) return Promise.resolve(false);
    var meta = { fichiers: ['mise à jour automatique'], le: Date.now(), publie: p.publie };
    appliquer(p, meta);
    return store.set('dataset', { p: p, meta: meta }).catch(function () { /* stockage plein */ }).then(function () { return true; });
  }
  function setSyncUI() {
    $('#btn-sync').hidden = !sync.key; $('#btn-sync').classList.toggle('spin', sync.running);
    $('#btn-sync-menu').hidden = !sync.key; $('#btn-auto-off').hidden = !sync.key;
    $('#btn-auto').textContent = sync.key ? 'Changer le mot de passe d’équipe' : 'Activer la mise à jour automatique';
  }
  function syncNow(manual) {
    if (!sync.key || sync.running) return Promise.resolve();
    sync.running = true; setSyncUI();
    return fetchRemote().then(function (buf) {
      var h = parseHeader(buf);
      if (b64(h.salt) !== sync.salt || h.iter !== sync.iter) throw err('Le mot de passe d’équipe a changé.', 'cle');
      return decrypt(h, sync.key);
    }).then(function (p) { return applyRemote(p, false); }).then(function (ch) {
      lsSet('mvech-sync-last', Date.now()); lsSet('mvech-sync-err', null);
      if (ch && !manual) toast('Données mises à jour automatiquement.'); else if (manual) toast(ch ? 'Données mises à jour.' : 'Les données sont déjà à jour.');
    }).catch(function (e) {
      var m = (e && e.message) || String(e);
      lsSet('mvech-sync-err', { m: m, t: Date.now() });
      if (e.code === 'cle') { sync.key = null; store.del('cle'); toast(m + ' Saisissez-le à nouveau.', 6000); openPassword(); }
      else if (manual || !state.D || e.code !== 'reseau') toast('Mise à jour impossible : ' + m, 6000);
    }).then(function () { sync.running = false; setSyncUI(); });
  }
  function openPassword() { $('#pw').value = ''; $('#pw-msg').textContent = ''; ouvrir('#sheet-pass'); setTimeout(function () { $('#pw').focus(); }, 50); }
  function adopt(key, h) {
    sync.key = key; sync.salt = b64(h.salt); sync.iter = h.iter;
    return store.set('cle', { key: key, salt: sync.salt, iter: h.iter }).catch(function () { /* mémoire */ });
  }
  function connect(pw) {
    $('#pw-msg').textContent = 'Connexion…'; $('#btn-pw').disabled = true;
    fetchRemote().then(function (buf) {
      var h = parseHeader(buf);
      return deriveKey(pw, h.salt, h.iter).then(function (key) {
        return decrypt(h, key).then(function (p) {
          return adopt(key, h).then(function () { fermer(); return applyRemote(p, true); })
            .then(function () { lsSet('mvech-sync-last', Date.now()); setSyncUI(); toast('Mise à jour automatique activée.'); });
        });
      });
    }).catch(function (e) { $('#pw-msg').textContent = e.message; }).then(function () { $('#btn-pw').disabled = false; });
  }
  // Si « Suivi MVE » est déjà activé sur ce téléphone, on réutilise sa clé (même mot de passe)
  function cleSuivi() {
    return store.get('cle', 'suivi-mve').then(function (c) {
      if (!c || !c.key) return false;
      return fetchRemote().then(function (buf) {
        var h = parseHeader(buf);
        if (b64(h.salt) !== c.salt || h.iter !== c.iter) return false;
        return decrypt(h, c.key).then(function (p) { return adopt(c.key, h).then(function () { return applyRemote(p, true); }).then(function () { return true; }); });
      });
    }).catch(function () { return false; });
  }


  /* ---------- Rendu ---------- */
  function rendre() {
    var D = state.D, f = state.f; if (!D) return;
    state.S = MVECH.synthese(D, f);
    var m = state.meta || {};
    $('#src').textContent = m.publie ? 'DHIS2 SNIS · publication du ' + new Date(m.publie).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + ' · ' + nf.format(state.D.cas.length) + ' cas, ' + nf.format(state.D.contacts.length) + ' contacts'
      : (m.fichiers ? m.fichiers.join(' + ') : '') + ' · lus le ' + new Date(m.le).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' });
    var per = f.debut == null && f.fin == null ? 'Toute l’épidémie' : fd(f.debut) + ' → ' + fd(f.fin), c = $('#chips');
    c.replaceChildren.apply(c, [['Lieu', f.zone || 'Toutes les zones'], ['Chaînes commencées', per]].map(function (it) {
      var b = document.createElement('button'); b.className = 'chip'; var bb = document.createElement('b'); bb.textContent = it[1]; b.append(it[0] + ' : ', bb); b.addEventListener('click', ouvrirFiltres); return b;
    }));
    if (!state.sel || state.S.chaines.indexOf(state.sel) < 0) state.sel = state.S.chaines[0] || null;
    renderSynthese(); renderChaines(); renderValider(); renderLacunes();
  }
  function tuile(lab, v, fmt, defn, lect, cls) {
    var b = document.createElement('button'); b.className = 'tile' + (cls ? ' ' + cls : '');
    var l = document.createElement('span'); l.className = 'lab'; l.textContent = lab;
    var x = document.createElement('span'); x.className = 'val'; x.textContent = v == null ? '–' : fmt === 'p' ? pf.format(v) : fmt === 'd1' ? df1.format(v) : nf.format(v);
    b.append(l, x);
    b.addEventListener('click', function () { $('#t-def').textContent = lab; $('#def-val').textContent = x.textContent; $('#def-def').textContent = defn; $('#def-int').textContent = lect; ouvrir('#sheet-def'); });
    return b;
  }
  function remplir(id, l) { var h = $(id); h.replaceChildren.apply(h, l); }
  function carteChaine(ch, petit) {
    var d = document.createElement('div'); d.className = 'item click';
    var top = document.createElement('div'); top.className = 'top-line';
    var id = document.createElement('span'); id.className = 'idn'; id.textContent = 'Chaîne ' + ch.num + ' · ' + ch.racine.court;
    var nb = document.createElement('span'); nb.className = 'nb'; nb.textContent = ch.generations + ' génération' + (ch.generations > 1 ? 's' : '');
    if (ch.generations < 2) { nb.style.background = 'var(--bg)'; nb.style.color = 'var(--muted)'; }
    top.append(id, nb);
    var m = document.createElement('div'); m.className = 'meta';
    m.textContent = (ch.nbSecondaires ? ch.nbSecondaires + ' cas secondaire' + (ch.nbSecondaires > 1 ? 's' : '') + ' · ' : '') + nf.format(ch.contacts) + ' contacts (' + (ch.contacts ? pf.format(ch.vus / ch.contacts) : '–') + ' vus au moins une fois)' +
      ' · ' + (ch.zones.join(', ') || 'zone ?') + ' · ' + fd(ch.debut) + (ch.fin !== ch.debut ? ' → ' + fd(ch.fin) : '');
    d.append(top, m);
    if (!petit && ch.liensProbables) { var t = document.createElement('div'); t.className = 'tags'; var s = document.createElement('span'); s.className = 'tag-prob'; s.textContent = ch.liensProbables + ' lien' + (ch.liensProbables > 1 ? 's' : '') + ' probable' + (ch.liensProbables > 1 ? 's' : '') + ' à valider'; t.appendChild(s); d.appendChild(t); }
    d.addEventListener('click', function () { state.sel = ch; state.noeud = null; montrer('chaines'); });
    return d;
  }
  function renderSynthese() {
    var k = state.S.k, f = state.f;
    remplir('#tiles-ch', [
      tuile('Chaînes (cas avec contacts)', k.chaines, 'n', 'Groupes formés par un cas source, ses contacts et, le cas échéant, les cas secondaires reliés.', 'Une chaîne peut se limiter à un cas et ses contacts (1 génération).', 'hero'),
      tuile('Chaînes de 2 générations et plus', k.multi, 'n', 'Chaînes où au moins un contact est devenu un cas retrouvé dans les alertes.', 'Transmission documentée d’un cas à un autre.'),
      tuile('Plus longue chaîne (générations)', k.maxGen, 'n', 'Nombre maximal de générations successives de cas reliés.', 'Avec les données actuelles, les générations suivantes sont souvent perdues faute de lien saisi.')
    ]);
    $('#sub-cas').textContent = 'Cas notifiés ' + (f.debut == null && f.fin == null ? 'depuis le début' : 'du ' + fd(f.debut) + ' au ' + fd(f.fin)) + (f.zone ? ', zone ' + de(f.zone) : '') + '.';
    remplir('#tiles-cas', [
      tuile('Cas confirmés', k.casConf, 'n', 'Alertes dont l’indicateur DHIS2 « Cas confirmés » = 1 (même règle que « Suivi MVE »).', 'Base de comparaison avec « Suivi MVE ».', 'hero'),
      tuile('Cas probables', k.casProb, 'n', 'Alertes classées « Cas probable » (non confirmées).', 'Additionnés aux cas confirmés pour les indicateurs ci-dessous.'),
      tuile('Cas avec contacts listés', k.pAvecContacts, 'p', 'Part des cas confirmés ou probables dont le N° Epid figure comme cas source sur au moins une fiche contact.', 'Viser 100 % : chaque cas doit avoir ses contacts listés.', k.pAvecContacts == null ? '' : k.pAvecContacts >= .9 ? 'q-good' : k.pAvecContacts >= .8 ? 'q-mid' : 'q-bad'),
      tuile('Cas sans aucun contact listé', k.sansContacts, 'n', 'Cas confirmés ou probables sans fiche contact rattachée.', 'Liste dans l’onglet « Lacunes ».', k.sansContacts ? 'q-bad' : 'q-good'),
      tuile('Contacts listés', k.contacts, 'n', 'Contacts rattachés aux cas sources des chaînes affichées.', ''),
      tuile('Contacts par cas source', k.parSource, 'd1', 'Contacts ÷ cas sources ayant au moins un contact.', 'Un chiffre faible peut indiquer un listage incomplet.'),
      tuile('Contacts devenus confirmés', k.devenus, 'n', 'Contacts dont l’issue du suivi est « Confirmé ».', 'Cas secondaires détectés par le suivi des contacts.')
    ]);
    remplir('#tiles-liens', [
      tuile('Liens forts', k.liensForts, 'n', 'Contact retrouvé dans les alertes : nom identique (ou contact noté devenu confirmé/suspect), même sexe, âge ±3 ans, notifié entre le dernier contact et 42 jours après.', 'Lien quasi certain.', 'q-good'),
      tuile('Liens probables', k.liensProbables, 'n', '2 mots du nom sur 3 en commun, même sexe, âge et dates compatibles, même zone.', 'À confirmer par l’équipe : onglet « À valider ».', k.liensProbables ? 'q-mid' : ''),
      tuile('Cas reliés à une chaîne', k.pRelies, 'p', 'Part des cas confirmés ou probables reliés à au moins un autre cas.', 'Faible tant que le N° Epid de l’alerte n’est pas saisi sur la fiche du contact devenu cas.'),
      tuile('Cas sources hors liste des alertes', k.horsListe, 'n', 'N° Epid saisis comme cas source mais absents des alertes (format erroné ou alerte d’une autre province).', 'Liste dans l’onglet « Lacunes ».', k.horsListe ? 'q-bad' : 'q-good')
    ]);
    var lieu = f.zone ? 'la zone ' + de(f.zone) : 'le Haut-Uélé';
    var parts = ['Dans ' + lieu + ', ', ['b', nf.format(k.chaines)], ' chaînes de transmission, dont ', ['b', nf.format(k.multi)], ' sur 2 générations ou plus (au maximum ', ['b', nf.format(k.maxGen)], '). ',
      ['b', pf.format(k.pAvecContacts || 0)], ' des cas ont des contacts listés ; ', ['b', nf.format(k.liensForts + k.liensProbables)], ' liens entre cas ont été retrouvés, dont ', ['b', nf.format(k.liensProbables)], ' à valider.'];
    var l = $('#lede'); l.replaceChildren.apply(l, parts.map(function (p) { if (typeof p === 'string') return document.createTextNode(p); var e = document.createElement(p[0]); e.textContent = p[1]; return e; }));
    remplir('#list-top', state.S.chaines.slice(0, 6).map(function (ch) { return carteChaine(ch, true); }));
  }

  /* ---------- Schéma d'une chaîne (arbre de gauche à droite) ---------- */
  var NS = 'http://www.w3.org/2000/svg';
  function el(tag, a, p) { var e = document.createElementNS(NS, tag); for (var k in a) e.setAttribute(k, a[k]); if (p) p.appendChild(e); return e; }
  function schema(ch) {
    var host = $('#schema'); host.replaceChildren();
    if (!ch) { host.textContent = 'Aucune chaîne pour ces filtres.'; return; }
    var W = 168, H = 58, GX = 56, GY = 14, ligne = 0, pos = {};
    (function place(n) {
      if (!n.enfants.length) { pos[n.id] = { x: n.gen * (W + GX) + 12, y: ligne * (H + GY) + 12 }; ligne++; return; }
      n.enfants.forEach(place);
      var ys = n.enfants.map(function (e) { return pos[e.id].y; });
      pos[n.id] = { x: n.gen * (W + GX) + 12, y: (Math.min.apply(null, ys) + Math.max.apply(null, ys)) / 2 };
    })(ch.racine);
    var largeur = ch.generations * (W + GX) - GX + 24, hauteur = ligne * (H + GY) - GY + 24;
    var svg = el('svg', { width: largeur, height: hauteur, viewBox: '0 0 ' + largeur + ' ' + hauteur, role: 'img', 'aria-label': 'Schéma de la chaîne ' + ch.num }, host);
    ch.membres.forEach(function (n) {
      if (!n.parent) return;
      var a = pos[n.parent.id], b = pos[n.id], x1 = a.x + W, y1 = a.y + H / 2, x2 = b.x, y2 = b.y + H / 2, mx = (x1 + x2) / 2;
      var p = el('path', { d: 'M' + x1 + ',' + y1 + ' C' + mx + ',' + y1 + ' ' + mx + ',' + y2 + ' ' + x2 + ',' + y2, fill: 'none', 'stroke-width': 2, style: 'stroke:var(--fg)' }, svg);
      if (n.lienParent.niveau === 'probable') p.setAttribute('stroke-dasharray', '6 4');
      var t = el('title', {}, p); t.textContent = 'Lien ' + n.lienParent.niveau + ' : contact ' + (n.lienParent.contact.id || '(sans ID)');
    });
    ch.membres.forEach(function (n) {
      var p = pos[n.id], g = el('g', { class: 'nd' + (state.noeud === n ? ' sel' : ''), transform: 'translate(' + p.x + ',' + p.y + ')', tabindex: 0, role: 'button', 'aria-label': n.id }, svg);
      var coul = n.cas === 'Confirmé' ? ['var(--bad-bg)', 'var(--bad)'] : n.cas === 'Probable' ? ['var(--warn-bg)', 'var(--warn)'] : ['var(--panel)', 'var(--muted)'];
      el('rect', { width: W, height: H, rx: 8, style: 'fill:' + coul[0] + ';stroke:' + coul[1], 'stroke-width': 1.5 }, g);
      var t1 = el('text', { x: 8, y: 17, class: 'b' }, g); t1.textContent = n.court + (n.decede ? '  ✝' : '');
      var t2 = el('text', { x: 8, y: 33, class: 'm' }, g); t2.textContent = (n.dn != null ? fd(n.dn) : 'date ?') + ' · ' + (n.zone || (n.alerte ? 'zone ?' : 'hors liste'));
      var t3 = el('text', { x: 8, y: 49, class: 'm' }, g); t3.textContent = n.contacts.length + ' contact' + (n.contacts.length > 1 ? 's' : '') + (n.cas ? ' · ' + n.cas.toLowerCase() : '');
      var go = function () { state.noeud = n; schema(ch); detail(n); };
      g.addEventListener('click', go); g.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } });
    });
  }
  function detail(n) {
    var d = $('#detail'); d.hidden = false; d.replaceChildren();
    var add = function (lab, val) { var p = document.createElement('div'); var b = document.createElement('b'); b.textContent = lab + ' : '; p.append(b, val); d.appendChild(p); };
    add('N° Epid', n.id); add('Classification', n.classif + (n.decede ? ' (décédé)' : ''));
    add('Notification', fd(n.dn) + ' · ' + (n.zone || '–') + (n.aire ? ' › ' + n.aire : ''));
    var vus = n.contacts.filter(function (c) { return c.vu; }).length, dev = n.contacts.filter(function (c) { return /confirm/i.test(c.issue || ''); }).length;
    add('Contacts listés', n.contacts.length + (n.contacts.length ? ' (' + vus + ' vus au moins une fois' + (dev ? ', ' + dev + ' devenu' + (dev > 1 ? 's' : '') + ' confirmé' + (dev > 1 ? 's' : '') : '') + ')' : ''));
    if (n.parent) add('Contaminé par', n.parent.id + ' — lien ' + n.lienParent.niveau + ' (fiche contact ' + (n.lienParent.contact.id || 'sans ID') + ')');
    if (n.enfants.length) add('A contaminé', n.enfants.map(function (e) { return e.court; }).join(', '));
  }
  function renderChaines() {
    if (state.tab !== 'chaines') return;
    var ch = state.sel;
    $('#schema-titre').textContent = ch ? 'Chaîne ' + ch.num + ' – cas source ' + ch.racine.id : 'Schéma';
    $('#schema-sub').textContent = ch ? ch.generations + ' génération' + (ch.generations > 1 ? 's' : '') + ', ' + ch.membres.length + ' cas, ' + nf.format(ch.contacts) + ' contacts. Touchez un cas pour le détail.' : '';
    schema(ch); if (state.noeud && ch && ch.membres.indexOf(state.noeud) >= 0) detail(state.noeud); else $('#detail').hidden = true;
    $$('#q-gen button').forEach(function (b) { b.setAttribute('aria-pressed', b.dataset.g === state.gen ? 'true' : 'false'); });
    var q = MVECH.norm($('#q-chaines').value), l = state.S.chaines.filter(function (c) { return c.generations >= +state.gen; });
    if (q) l = l.filter(function (c) { return c.membres.some(function (m) { return MVECH.norm(m.id + ' ' + m.court + ' ' + (m.zone || '')).indexOf(q) >= 0; }); });
    remplir('#list-chaines', l.slice(0, state.nCh).map(function (c) { var x = carteChaine(c); if (c === state.sel) x.style.borderColor = 'var(--ink-2)'; return x; }));
    if (!l.length) { var p = document.createElement('p'); p.className = 'sub'; p.textContent = 'Aucune chaîne ne correspond.'; $('#list-chaines').appendChild(p); }
    $('#more-chaines').hidden = l.length <= state.nCh; $('#more-chaines').textContent = 'Afficher plus (' + (l.length - state.nCh) + ' restantes)';
  }
  function item(titre, badge, cls, meta, tags, clic) {
    var d = document.createElement('div'); d.className = 'item' + (clic ? ' click' : '');
    var top = document.createElement('div'); top.className = 'top-line';
    var a = document.createElement('span'); a.className = 'idn'; a.textContent = titre; top.appendChild(a);
    if (badge) { var b = document.createElement('span'); b.className = 'nb ' + (cls || ''); b.textContent = badge; top.appendChild(b); }
    var m = document.createElement('div'); m.className = 'meta'; m.textContent = meta; d.append(top, m);
    if (tags && tags.length) { var t = document.createElement('div'); t.className = 'tags'; tags.forEach(function (x) { var s = document.createElement('span'); s.textContent = x; t.appendChild(s); }); d.appendChild(t); }
    if (clic) d.addEventListener('click', clic);
    return d;
  }
  function renderValider() {
    if (state.tab !== 'valider') return;
    var l = state.S.aValider, np = l.filter(function (e) { return e.niveau === 'probable'; }).length;
    $('#sub-valider').textContent = l.length + ' liens retrouvés par rapprochement (' + np + ' probables d’abord, puis ' + (l.length - np) + ' forts). Pour chacun, vérifiez dans DHIS2 que la fiche contact et l’alerte concernent bien la même personne, puis saisissez le N° Epid de l’alerte sur la fiche du contact.';
    remplir('#list-valider', l.map(function (e) {
      var c = e.contact, cible = state.D.noeuds[e.vers];
      return item(MVECH.court(e.de) + ' → ' + MVECH.court(e.vers), 'lien ' + e.niveau, e.niveau === 'fort' ? 'tag-fort' : 'tag-prob',
        'Fiche contact ' + (c.id || '(sans ID)') + ' (cas source ' + e.de + '), listée le ' + fd(c.dEnr) + ', dernier contact le ' + fd(c.dDern) + ' · alerte ' + e.vers + ' notifiée le ' + fd(cible && cible.dn) + ' · ' + (c.zone || '–'),
        [c.issue ? 'Issue du suivi : ' + c.issue : 'Issue du suivi non renseignée', cible && cible.cas ? 'Alerte : cas ' + cible.cas.toLowerCase() : ''].filter(Boolean),
        function () { if (cible && cible.chaine) { state.sel = cible.chaine; state.noeud = cible; montrer('chaines'); } });
    }));
    if (!l.length) $('#list-valider').textContent = 'Aucun lien à valider.';
  }
  function renderLacunes() {
    if (state.tab !== 'lacunes') return;
    var S = state.S;
    $('#sub-l1').textContent = S.devenusSansAlerte.length + ' contact(s) noté(s) « Confirmé » dont l’alerte n’a pas été retrouvée (nom absent, différent ou incomplet). Saisissez le N° Epid de leur alerte pour prolonger la chaîne.';
    remplir('#list-l1', S.devenusSansAlerte.map(function (x) { return item(x.c.id || '(sans ID contact)', 'devenu confirmé', '', 'Cas source ' + x.source.id + ' · listé le ' + fd(x.c.dEnr) + ' · ' + (x.c.zone || '–'), []); }));
    var hors = []; state.S.chaines.forEach(function (ch) { ch.membres.forEach(function (m) { if (!m.alerte && m.contacts.length) hors.push(m); }); });
    hors.sort(function (a, b) { return b.contacts.length - a.contacts.length; });
    $('#sub-l2').textContent = hors.length + ' N° Epid saisis comme cas source sur des fiches contact mais absents de la liste des alertes : souvent une faute de frappe (format attendu RDC-HUE-XXX-26-0000000-MVE).';
    remplir('#list-l2', hors.map(function (m) { var ids = m.contacts.slice(0, 3).map(function (c) { return c.id || '(sans ID)'; }); return item(m.id, m.contacts.length + ' contact' + (m.contacts.length > 1 ? 's' : ''), '', 'Exemples de fiches : ' + ids.join(', ') + (m.contacts.length > 3 ? '…' : ''), []); }));
    var q = MVECH.norm($('#q-l3').value), l3 = S.sansContacts.filter(function (a) { return !q || MVECH.norm(a.id + ' ' + a.court + ' ' + (a.zone || '')).indexOf(q) >= 0; });
    $('#sub-l3').textContent = S.sansContacts.length + ' cas confirmés ou probables sans aucune fiche contact rattachée (les plus récents d’abord). Vérifiez que leurs contacts ont été listés et que le N° Epid du cas source est bien saisi sur les fiches.';
    remplir('#list-l3', l3.slice(0, state.nL3).map(function (a) { return item(a.id, a.cas, '', 'Notifié le ' + fd(a.dn) + ' · ' + (a.zone || '–') + (a.aire ? ' › ' + a.aire : '') + (a.decede ? ' · décédé' : ''), []); }));
    $('#more-l3').hidden = l3.length <= state.nL3; $('#more-l3').textContent = 'Afficher plus (' + (l3.length - state.nL3) + ' restants)';
  }

  /* ---------- Onglets, feuilles, filtres ---------- */
  function montrer(t) {
    state.tab = t; lsSet('mvech-onglet', t);
    ['synthese', 'chaines', 'valider', 'lacunes'].forEach(function (x) { $('#tab-' + x).hidden = x !== t || !state.D; });
    $$('#tabs button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.tab === t ? 'true' : 'false'); });
    window.scrollTo(0, 0);
    if (state.S) { renderChaines(); renderValider(); renderLacunes(); }
  }
  function ouvrir(id) { fermer(true); $('#sheet-bg').hidden = false; $(id).hidden = false; history.pushState({ sheet: id }, ''); }
  function fermer(silent) {
    var o = $$('.sheet').some(function (s) { return !s.hidden; });
    $$('.sheet').forEach(function (s) { s.hidden = true; }); $('#sheet-bg').hidden = true;
    if (o && !silent && history.state && history.state.sheet) history.back();
  }
  window.addEventListener('popstate', function () { $$('.sheet').forEach(function (s) { s.hidden = true; }); $('#sheet-bg').hidden = true; });
  function ouvrirFiltres() {
    var z = {}; state.D.cas.forEach(function (a) { if (a.zone) z[a.zone] = 1; });
    var sel = $('#f-zone'); sel.replaceChildren(new Option('Toutes les zones', ''));
    Object.keys(z).sort(function (a, b) { return a.localeCompare(b, 'fr'); }).forEach(function (x) { sel.appendChild(new Option(x, x)); });
    sel.value = state.f.zone || '';
    $('#f-debut').value = state.f.debut != null ? sToInput(state.f.debut) : ''; $('#f-fin').value = state.f.fin != null ? sToInput(state.f.fin) : '';
    ouvrir('#sheet-filtres');
  }
  $('#btn-apply').addEventListener('click', function () {
    var d = $('#f-debut').value ? MVECH.toSerial($('#f-debut').value) : null, e = $('#f-fin').value ? MVECH.toSerial($('#f-fin').value) : null;
    if (d != null && e != null && d > e) { toast('La date de début est après la date de fin.'); return; }
    state.f = { zone: $('#f-zone').value, debut: d, fin: e }; lsSet('mvech-filtres', state.f); state.sel = null; state.noeud = null; fermer(); rendre();
  });
  $('#btn-reset').addEventListener('click', function () { state.f = { zone: '', debut: null, fin: null }; lsSet('mvech-filtres', state.f); fermer(); rendre(); });
  $$('#tabs button').forEach(function (b) { b.addEventListener('click', function () { montrer(b.dataset.tab); }); });
  $('#sheet-bg').addEventListener('click', function () { fermer(); });
  $$('[data-close]').forEach(function (b) { b.addEventListener('click', function () { fermer(); }); });
  $('#btn-open-empty').addEventListener('click', function () { $('#file').click(); });
  $('#btn-open').addEventListener('click', function () { fermer(); $('#file').click(); });
  $('#file').addEventListener('change', function (e) { if (e.target.files.length) lireFichiers(e.target.files); e.target.value = ''; });
  $('#btn-menu').addEventListener('click', function () {
    var m = state.meta;
    var last = lsGet('mvech-sync-last');
    $('#menu-info').textContent = (m ? (m.publie ? 'Données publiées le ' + new Date(m.publie).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + '.' : 'Classeurs lus : ' + m.fichiers.join(' et ') + ', le ' + new Date(m.le).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + '.' + (state.bureau ? ' Ils sont relus automatiquement à chaque nouvel enregistrement.' : '')) : 'Aucune donnée chargée.') +
      (state.bureau ? '' : sync.key ? ' Mise à jour automatique activée' + (last ? ', dernière vérification le ' + new Date(last).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) : '') + '.' : ' Mise à jour automatique désactivée.');
    if (!state.bureau) setSyncUI();
    ouvrir('#sheet-menu');
  });
  $('#btn-sync').addEventListener('click', function () { if (state.bureau) bureau(true); else syncNow(true); });
  $('#btn-sync-menu').addEventListener('click', function () { fermer(); if (state.bureau) bureau(true); else syncNow(true); });
  $('#btn-auto').addEventListener('click', openPassword);
  $('#btn-auto-empty').addEventListener('click', function () {
    cleSuivi().then(function (ok) { if (ok) { setSyncUI(); toast('Mise à jour automatique activée avec le mot de passe de « Suivi MVE ».'); } else openPassword(); });
  });
  $('#btn-auto-off').addEventListener('click', function () { sync.key = null; store.del('cle'); setSyncUI(); fermer(); toast('Mise à jour automatique désactivée sur ce téléphone.'); });
  $('#form-pw').addEventListener('submit', function (e) { e.preventDefault(); var pw = $('#pw').value; if (!pw) { $('#pw-msg').textContent = 'Saisissez le mot de passe d’équipe.'; return; } connect(pw); });
  $('#pw-show').addEventListener('change', function () { $('#pw').type = $('#pw-show').checked ? 'text' : 'password'; });
  var installEvt = null;
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); installEvt = e; if (!state.bureau) $('#btn-install').hidden = false; });
  $('#btn-install').addEventListener('click', function () { if (!installEvt) return; fermer(); installEvt.prompt(); installEvt.userChoice.finally(function () { installEvt = null; $('#btn-install').hidden = true; }); });
  $('#q-gen').addEventListener('click', function (e) { var g = e.target.dataset && e.target.dataset.g; if (!g) return; state.gen = g; state.nCh = 40; renderChaines(); });
  $('#q-chaines').addEventListener('input', function () { state.nCh = 40; renderChaines(); });
  $('#more-chaines').addEventListener('click', function () { state.nCh += 40; renderChaines(); });
  $('#q-l3').addEventListener('input', function () { state.nL3 = 40; renderLacunes(); });
  $('#more-l3').addEventListener('click', function () { state.nL3 += 40; renderLacunes(); });

  var sf = lsGet('mvech-filtres'); if (sf) state.f = { zone: sf.zone || '', debut: sf.debut == null ? null : sf.debut, fin: sf.fin == null ? null : sf.fin };
  montrer(lsGet('mvech-onglet') || 'synthese');
  var h = location.hostname;
  if (h === '127.0.0.1' || h === 'localhost') {
    $('#btn-auto-empty').hidden = true; $('#btn-auto').hidden = true;
    bureau(false).then(function () { if (!state.bureau) telephone(); }); // servi en local sans le programme : comme sur le téléphone
    document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible') bureau(false); });
    setInterval(function () { if (document.visibilityState === 'visible') bureau(false); else fetch('/bureau/ping').catch(function () {}); }, 60000);
  }

  /* Téléphone (site publié) : données publiées, même mot de passe que « Suivi MVE » */
  function telephone() {
    $('#btn-auto-empty').hidden = false; $('#btn-auto').hidden = false;
    store.get('cle').then(function (c) { if (c && c.key) { sync.key = c.key; sync.salt = c.salt; sync.iter = c.iter; } }).catch(function () { /* aucune clé */ }).then(function () {
      setSyncUI();
      return store.get('dataset').then(function (d) { if (d && d.p) appliquer(d.p, d.meta); }).catch(function () { /* première ouverture */ });
    }).then(function () {
      if (sync.key) return syncNow(false);
      return cleSuivi().then(function (ok) { if (ok) { setSyncUI(); toast('Mise à jour automatique activée avec le mot de passe de « Suivi MVE ».'); } });
    });
    document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible' && Date.now() - (lsGet('mvech-sync-last') || 0) > 2 * 60000) syncNow(false); });
    window.addEventListener('online', function () { syncNow(false); });
    if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(function () { /* hors ligne indisponible */ });
  }
  if (h !== '127.0.0.1' && h !== 'localhost') telephone();
})();
