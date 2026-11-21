/* Interface de l'application « Contacts MVE ». Calculs : engine-contacts.js */
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };
  var state = { data: null, filters: null, tab: 'synthese', result: null, nRel: 50, cat: '' };
  var nf = new Intl.NumberFormat('fr-FR');
  var pf = new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 1 });
  var df1 = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 });
  var OBJ = 0.9, ORANGE = 0.8;

  /* ---------- Dates ---------- */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function sToDate(s) { return MVEC.serialToUTC(s); }
  function fmtDate(s, short) { if (s == null) return '–'; var d = sToDate(s); return pad(d.getUTCDate()) + '/' + pad(d.getUTCMonth() + 1) + (short ? '' : '/' + d.getUTCFullYear()); }
  function sToInput(s) { var d = sToDate(s); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function inputToS(v) { return v ? MVEC.toSerial(v) : null; }
  function today() { return Math.floor(MVEC.nowSerial()); }

  /* ---------- Stockage local ---------- */
  function idb(name) {
    return new Promise(function (res, rej) {
      var rq = indexedDB.open(name || 'contacts-mve', 1);
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
  function lsGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* indisponible */ } }

  var toastT;
  function toast(m, ms) { var t = $('#toast'); t.textContent = m; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(function () { t.hidden = true; }, ms || 3500); }
  function busy(on) { $('#busy').hidden = !on; }

  /* ---------- Chargement ---------- */
  function setFromTables(contacts, suivis, meta) {
    var out = MVEC.normalize(contacts, suivis);
    return Object.assign({ rows: out.rows, visites: out.visites, extrait: out.extrait, charge: out.charge, nbSuivis: out.nbSuivis, chargeLe: Date.now() }, meta);
  }
  /* Seules les colonnes A à V de « Données » et A à J de « Suivis » sont gardées : les noms (colonnes W à Y) sont écartés dès la lecture. */
  function readWorkbook(buf, name, garder) {
    busy(true);
    return new Promise(function (r) { setTimeout(r, 30); }).then(function () {
      var wb = XLSX.read(buf, { type: 'array', sheets: ['Données', 'Suivis'], cellDates: false, cellFormula: false, cellHTML: false, dense: false });
      if (!wb.Sheets['Données']) throw new Error('Ce fichier ne contient pas la feuille « Données ».');
      var T = function (n) { return wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: false }) : []; };
      var contacts = MVEC.epurer(T('Données'), MVEC.COLS), suivis = MVEC.epurer(T('Suivis'), MVEC.SCOLS);
      var data = setFromTables(contacts, suivis, { fichier: name || 'fichier.xlsx' });
      return store.set('dataset', { contacts: contacts, suivis: suivis, fichier: data.fichier, publie: null, chargeLe: data.chargeLe })
        .catch(function () { /* stockage plein */ }).then(function () { return data; });
    }).then(function (d) { busy(false); setData(d, !garder); toast(nf.format(d.charge) + ' contacts chargés.'); })
      .catch(function (e) { busy(false); toast('Lecture impossible : ' + (e && e.message ? e.message : e), 6000); });
  }
  // Programme d'ordinateur (dossier de classeurs) : lecture directe d'un classeur sans passer par le sélecteur de fichier
  window.MVE_lireClasseur = function (buf, name, garder) { return readWorkbook(buf, name, garder); };
  function openFile(f) { if (f) f.arrayBuffer().then(function (b) { return readWorkbook(new Uint8Array(b), f.name); }); }

  /* ---------- Mise à jour automatique (contacts.enc, même mot de passe que Suivi MVE) ---------- */
  var REMOTE = 'contacts.enc', sync = { key: null, salt: null, iter: 0, running: false };
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
      if (r.status === 404) throw err('Aucune donnée de suivi des contacts n’a encore été publiée.', 'absent');
      if (!r.ok) throw err('Le site de l’application ne répond pas (erreur ' + r.status + ').');
      return r.arrayBuffer();
    }, function () { throw err('Connexion impossible : les dernières données restent affichées.', 'reseau'); });
  }
  function applyRemote(p, force) {
    var cur = state.data;
    if (!force && cur && cur.publie && p.publie && p.publie <= cur.publie) return Promise.resolve(false);
    var d = setFromTables(p.contacts, p.suivis, { fichier: 'mise à jour automatique', publie: p.publie });
    return store.set('dataset', { contacts: p.contacts, suivis: p.suivis, fichier: d.fichier, publie: p.publie, chargeLe: d.chargeLe })
      .catch(function () { /* stockage plein */ }).then(function () { setData(d, !cur); return true; });
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
      lsSet('mvec-sync-last', Date.now()); lsSet('mvec-sync-err', null);
      if (ch && !manual) toast('Données mises à jour automatiquement.'); else if (manual) toast(ch ? 'Données mises à jour.' : 'Les données sont déjà à jour.');
    }).catch(function (e) {
      var m = (e && e.message) || String(e);
      lsSet('mvec-sync-err', { m: m, t: Date.now() });
      if (e.code === 'cle') { sync.key = null; store.del('cle'); toast(m + ' Saisissez-le à nouveau.', 6000); openPassword(); }
      else if (manual || !state.data || e.code !== 'reseau') toast('Mise à jour impossible : ' + m, 6000);
    }).then(function () { sync.running = false; setSyncUI(); });
  }
  function openPassword() { $('#pw').value = ''; $('#pw-msg').textContent = ''; openSheet('#sheet-pass'); setTimeout(function () { $('#pw').focus(); }, 50); }
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
          return adopt(key, h).then(function () { closeSheets(); return applyRemote(p, true); })
            .then(function () { lsSet('mvec-sync-last', Date.now()); setSyncUI(); toast('Mise à jour automatique activée.'); });
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

  /* ---------- Filtres ---------- */
  function minN(rows) { var m = Infinity; rows.forEach(function (r) { if (r.dEnr != null && r.dEnr < m) m = r.dEnr; }); return m === Infinity ? today() : m; }
  function maxN(rows) { var m = 0; rows.forEach(function (r) { if (r.dEnr != null && r.dEnr > m) m = r.dEnr; }); return m; }
  function resolveFilters() {
    var s = lsGet('mvec-filtres') || {}, rows = state.data.rows, t = today();
    var f = { zone: s.zone || '', aire: s.aire || '', periode: s.periode || 'tout', debutFixe: s.debutFixe, finFixe: s.finFixe, refMode: s.refMode || '0', refFixe: s.refFixe };
    if (f.zone && MVEC.zonesOf(rows).indexOf(f.zone) < 0) { f.zone = ''; f.aire = ''; }
    f.ref = f.refMode === 'fixe' ? f.refFixe : t - (+f.refMode || 0);
    if (f.periode === 'fixe') { f.debut = f.debutFixe; f.fin = f.finFixe; }
    else if (f.periode === 'tout') { f.debut = minN(rows); f.fin = Math.max(t, maxN(rows)); }
    else { f.fin = f.ref; f.debut = f.ref - (+f.periode) + 1; }
    return f;
  }
  function saveFilters(f) { lsSet('mvec-filtres', { zone: f.zone, aire: f.aire, periode: f.periode, debutFixe: f.debutFixe, finFixe: f.finFixe, refMode: f.refMode, refFixe: f.refFixe }); }
  function fillZone(sel, v) { sel.replaceChildren(new Option('Toutes les zones', '')); MVEC.zonesOf(state.data.rows).forEach(function (z) { sel.appendChild(new Option(z, z)); }); sel.value = v || ''; }
  function fillAire(sel, z, v) {
    sel.replaceChildren(new Option(z ? 'Toutes les aires de ' + z : 'Toutes les aires', ''));
    MVEC.airesOf(state.data.rows, z).forEach(function (a) { sel.appendChild(new Option(a, a)); }); sel.value = v || ''; if (sel.value !== (v || '')) sel.value = '';
  }
  var draft = null;
  function openFilters() {
    var f = state.filters; draft = JSON.parse(JSON.stringify(f));
    fillZone($('#f-zone'), f.zone); fillAire($('#f-aire'), f.zone, f.aire);
    $('#f-debut').value = sToInput(f.debut); $('#f-fin').value = sToInput(f.fin); $('#f-ref').value = sToInput(f.ref);
    openSheet('#sheet-filtres');
  }
  function applyFilters() {
    var f = draft; f.zone = $('#f-zone').value; f.aire = $('#f-aire').value;
    var d = inputToS($('#f-debut').value), e = inputToS($('#f-fin').value), r = inputToS($('#f-ref').value);
    if (d == null || e == null || r == null) { toast('Renseignez les trois dates.'); return; }
    if (d > e) { toast('La date de début est après la date de fin.'); return; }
    if (f.periode === 'fixe') { f.debutFixe = d; f.finFixe = e; }
    if (f.refMode === 'fixe') f.refFixe = r;
    saveFilters(f); closeSheets(); state.filters = resolveFilters(); state.nRel = 50; render();
  }
  function setLieu(zone, aire, msg) {
    var s = lsGet('mvec-filtres') || {}; s.zone = zone; s.aire = aire; lsSet('mvec-filtres', s);
    state.filters = resolveFilters(); render(); toast(msg);
  }

  /* ---------- Rendu ---------- */
  function setData(d, fresh) {
    state.data = d; state.filters = resolveFilters();
    $('#empty').hidden = true; $('#tabs').hidden = false; $('#chips').hidden = false;
    showTab(fresh ? 'synthese' : (lsGet('mvec-onglet') || 'synthese'));
    render();
  }
  function cls(v) { return v == null ? '' : v >= OBJ ? 'good' : v >= ORANGE ? 'mid' : 'bad'; }
  function render() {
    var d = state.data, f = state.filters;
    state.result = MVEC.compute(d, f);
    var ex = d.extrait != null ? sToDate(d.extrait) : null;
    $('#src').textContent = 'DHIS2 SNIS · extrait le ' + (ex ? pad(ex.getUTCDate()) + '/' + pad(ex.getUTCMonth() + 1) + '/' + ex.getUTCFullYear() + ' à ' + pad(ex.getUTCHours()) + 'h' + pad(ex.getUTCMinutes()) : '(date inconnue)') +
      ' · ' + nf.format(d.charge) + ' contacts, ' + nf.format(d.nbSuivis) + ' jours de suivi';
    var age = d.extrait != null ? (Date.now() - (sToDate(d.extrait).getTime() + new Date().getTimezoneOffset() * 60000)) / 864e5 : null;
    $('#src').classList.toggle('stale', age != null && age > 2);
    renderChips(); renderSynthese(); renderGraphiques(); renderZones(); renderAnalyse(); renderRelancer();
  }
  function renderChips() {
    var f = state.filters, c = $('#chips');
    var lieu = f.aire ? f.aire + ' (' + (f.zone || 'aire') + ')' : (f.zone || 'Toutes les zones');
    var per = f.periode === 'tout' ? 'Depuis le début' : (/^\d+$/.test(f.periode) ? f.periode + ' derniers jours' : fmtDate(f.debut) + ' → ' + fmtDate(f.fin));
    c.replaceChildren.apply(c, [['Lieu', lieu], ['Listés', per], ['Réf.', fmtDate(f.ref)]].map(function (it) {
      var b = document.createElement('button'); b.className = 'chip'; var bb = document.createElement('b'); bb.textContent = it[1];
      b.append(it[0] + ' : ', bb); b.addEventListener('click', openFilters); return b;
    }));
  }

  /* Indicateurs (dictionnaire du classeur) : [clé, libellé, format, définition, lecture, filtres, coloration, hero] */
  var P_ = 'Zone + aire + période (date d’enregistrement du/au)', R_ = 'Zone + aire + date de référence (ignore la période)';
  var TILES = {
    ref: [
      ['aSuivre', 'Contacts à suivre – jour de réf.', 'n', 'Contacts « En suivi » : enregistrés au plus tard ce jour, entre J1 et J21 après le dernier contact, sans issue (suspect, confirmé, décédé, transféré, doublon) enregistrée avant ce jour.', 'Dénominateur du taux de suivi.', R_, null, true],
      ['vus', 'Contacts vus – jour de réf.', 'n', 'Contacts à suivre ayant au moins une fiche de suivi ce jour avec « Contact a été vu ? » = Oui (plusieurs fiches le même jour = 1 visite).', 'Numérateur du taux de suivi.', R_],
      ['taux', 'Taux de suivi – jour de réf.', 'p', 'Contacts vus ÷ contacts à suivre.', 'Objectif ≥ 90 %. Rouge < 80 %, orange 80–89 %, vert ≥ 90 %. Un taux bas sur les derniers jours peut refléter un retard de saisie.', R_, 'taux'],
      ['nonVus', 'Contacts non vus – jour de réf.', 'n', 'Contacts à suivre − contacts vus.', 'À relancer en priorité (onglet « À relancer »).', R_],
      ['nouveauxJour', 'Nouveaux contacts – jour de réf.', 'n', 'Contacts dont la date d’enregistrement = date de référence.', 'Activité de listage du jour.', R_],
      ['nouveaux7', 'Nouveaux contacts (7 j)', 'n', 'Date d’enregistrement > date de réf. − 7 et ≤ date de réf.', 'Tendance récente du listage.', R_]
    ],
    relance: [
      ['jamaisVus', 'Jamais vus (en suivi)', 'n', 'Contacts « En suivi » sans aucune fiche « vu = Oui » à une date ≤ date de réf.', 'En rouge si > 0 : contacts jamais atteints, risque majeur.', R_, 'zero'],
      ['nonVus2', 'Non vus depuis ≥ 2 j (en suivi)', 'n', 'Contacts « En suivi » dont la dernière visite « vu = Oui » date de 2 jours ou plus avant la date de réf.', 'En rouge si > 0 : à rechercher activement.', R_, 'zero'],
      ['nonVus24', 'Non vus 24 h (vus la veille)', 'n', 'Contacts « En suivi » non vus le jour de réf. mais vus la veille.', 'À visiter dans la journée.', R_]
    ],
    per: [
      ['listes', 'Contacts listés (période)', 'n', 'Contacts enregistrés dans DHIS2 pendant la période choisie.', 'Volume de listage ; dénominateur du % de contacts liés à un cas.', P_, null, true],
      ['casSources', 'Cas sources (période)', 'n', 'N° Epid distincts ayant généré des contacts ; chaque cas est compté à la date d’enregistrement de son premier contact.', 'Nombre de cas couverts par le listage.', P_],
      ['parCas', 'Contacts par cas source', 'd1', 'Contacts ayant un N° Epid renseigné ÷ cas sources.', 'Un chiffre faible peut indiquer un listage incomplet.', P_],
      ['pLies', '% contacts liés à un cas (N° Epid)', 'p', 'Contacts ayant un N° Epid renseigné ÷ contacts listés.', 'Indicateur de complétude : viser 100 %.', P_, 'taux'],
      ['delai', 'Délai moyen dernier contact → listage (j)', 'd1', 'Moyenne de (date d’enregistrement − date du dernier contact), valeurs négatives exclues.', 'Plus il est court, plus le suivi démarre tôt dans la période d’incubation.', P_],
      ['vusUneFois', 'Contacts vus au moins une fois', 'n', 'Contacts listés dans la période dont au moins une fiche de suivi porte « vu = Oui », quelle que soit la date.', 'À comparer aux contacts listés : l’écart = contacts jamais vus ou sans fiche saisie.', P_],
      ['suspects', 'Devenus suspects (période)', 'n', 'Contacts listés dans la période dont l’issue = « Devenu suspect ».', 'Efficacité de la surveillance des contacts.', P_],
      ['confirmes', 'Devenus cas confirmés (période)', 'n', 'Contacts listés dans la période dont l’issue = « Confirmé ».', 'Part des nouveaux cas issus de la liste des contacts.', P_],
      ['symptomes', 'Ont signalé des symptômes (période)', 'n', 'Contacts dont « Symptômes signalés » = Oui.', 'Alerte : chaque contact symptomatique doit être investigué.', P_, 'zero'],
      ['decedes', 'Contacts décédés (période)', 'n', 'Contacts dont l’issue = « Décédé ».', 'Tout décès de contact doit faire l’objet d’une investigation.', P_, 'zero'],
      ['sortis', 'Sortis – 21 j terminés (période)', 'n', 'Contacts au statut « Sorti (21 j terminés) » : date de réf. postérieure à J21 sans issue enregistrée.', 'Contacts libérés du suivi.', P_]
    ]
  };
  var T21 = [
    ['nouveaux', 'Contacts nouvellement listés', 'n', 'Contacts enregistrés pendant les 21 jours.'],
    ['casSources', 'Nouveaux cas sources', 'n', 'N° Epid dont le 1er contact a été listé pendant les 21 jours.'],
    ['parCas', 'Contacts par nouveau cas source', 'd1', 'Contacts avec N° Epid ÷ nouveaux cas sources.'],
    ['pLies', '% contacts liés à un cas', 'p', 'Contacts avec N° Epid ÷ contacts listés.', 'taux'],
    ['delai', 'Délai dernier contact → listage (j)', 'd1', 'Moyenne, valeurs négatives exclues.'],
    ['cjSuivre', 'Contacts-jours à suivre', 'n', 'Somme, sur les 21 jours, des contacts à suivre chaque jour.'],
    ['visites', 'Visites réalisées (vus)', 'n', 'Somme des contacts vus chaque jour (visites comptées entre J1 et J21, avant toute issue).'],
    ['taux', 'Taux de suivi moyen (21 j)', 'p', 'Visites ÷ contacts-jours à suivre. Objectif ≥ 90 %.', 'taux', true],
    ['tauxMin', 'Taux le plus bas', 'p', 'Jour au taux de suivi le plus faible.', 'taux'],
    ['tauxMax', 'Taux le plus haut', 'p', 'Jour au taux de suivi le plus élevé.', 'taux'],
    ['joursObjectif', 'Jours avec taux ≥ 90 %', 'n', 'Nombre de jours (sur 21) où l’objectif est atteint.'],
    ['suspects', 'Devenus suspects', 'n', 'Contacts dont la date d’issue « Devenu suspect » tombe dans les 21 jours.'],
    ['confirmes', 'Devenus confirmés', 'n', 'Contacts dont la date d’issue « Confirmé » tombe dans les 21 jours.'],
    ['symptomes', 'Listés ayant signalé des symptômes', 'n', 'Contacts listés pendant les 21 jours avec « Symptômes signalés » = Oui.', 'zero'],
    ['termineSansIssue', 'Arrivés au terme J21 sans issue', 'n', 'Contacts dont la fin du suivi (J21) tombe dans les 21 jours, sans issue.'],
    ['enSuivi', 'En suivi (date de réf.)', 'n', 'Contacts au statut « En suivi » à la date de référence.'],
    ['jamaisVus', 'dont jamais vus', 'n', 'Contacts en suivi sans aucune visite « vu = Oui ».', 'zero'],
    ['nonVus2', 'dont non vus depuis ≥ 2 j', 'n', 'Contacts en suivi dont la dernière visite date de 2 jours ou plus.', 'zero']
  ];
  function fmtVal(v, t) {
    if (v == null) return '–';
    if (t === 'p') return pf.format(v); if (t === 'd1') return df1.format(v);
    return nf.format(v);
  }
  function tile(t, v, note, def) {
    var b = document.createElement('button');
    b.className = 'tile' + (t.hero ? ' hero' : '');
    if (t.col === 'taux' && v != null) b.className += t.hero ? ' h-' + cls(v) : ' q-' + cls(v);
    else if (t.col === 'zero' && v != null) b.className += v === 0 ? ' q-good' : ' q-bad';
    var lab = document.createElement('span'); lab.className = 'lab'; lab.textContent = t.lab;
    var val = document.createElement('span'); val.className = 'val'; val.textContent = fmtVal(v, t.fmt);
    b.append(lab, val);
    if (note) { var n = document.createElement('span'); n.className = 'note'; n.textContent = note; b.appendChild(n); }
    b.addEventListener('click', function () {
      $('#t-def').textContent = t.lab; $('#def-val').textContent = fmtVal(v, t.fmt);
      $('#def-def').textContent = def[0]; $('#def-int').textContent = def[1] || ''; $('#def-act').textContent = def[2] || ''; openSheet('#sheet-def');
    });
    return b;
  }
  function renderSynthese() {
    var k = state.result.k, f = state.filters;
    Object.keys(TILES).forEach(function (g) {
      var host = $('#tiles-' + g);
      host.replaceChildren.apply(host, TILES[g].map(function (t) {
        return tile({ lab: t[1], fmt: t[2], col: t[6], hero: t[7] }, k[t[0]], null, [t[3], t[4], t[5]]);
      }));
    });
    $('#sub-ref').textContent = 'Situation du suivi au ' + fmtDate(f.ref) + ' (quel que soit le filtre de période).';
    $('#sub-per').textContent = 'Contacts enregistrés du ' + fmtDate(f.debut) + ' au ' + fmtDate(f.fin) + '.';
    var lieu = f.aire ? 'l’aire ' + de(f.aire) : f.zone ? 'la zone ' + de(f.zone) : 'le Haut-Uélé';
    var parts = ['Au ' + fmtDate(f.ref) + ', dans ' + lieu + ', ', ['b', nf.format(k.aSuivre)], ' contact' + (k.aSuivre > 1 ? 's' : '') + ' à suivre, dont ', ['b', nf.format(k.vus)], ' vu' + (k.vus > 1 ? 's' : '')];
    if (k.taux != null) parts.push(' (taux de suivi ', ['b', pf.format(k.taux)], ', objectif 90 %)');
    parts.push('. ');
    if (k.jamaisVus || k.nonVus2) parts.push(['b', nf.format(k.jamaisVus)], ' jamais vu' + (k.jamaisVus > 1 ? 's' : '') + ' et ', ['b', nf.format(k.nonVus2)], ' non vu' + (k.nonVus2 > 1 ? 's' : '') + ' depuis 2 jours ou plus. ');
    parts.push(['b', nf.format(k.listes)], ' contact' + (k.listes > 1 ? 's' : '') + ' listé' + (k.listes > 1 ? 's' : '') + ' sur la période, dont ', ['b', nf.format(k.suspects)], ' devenu' + (k.suspects > 1 ? 's' : '') + ' suspect' + (k.suspects > 1 ? 's' : '') + ' et ', ['b', nf.format(k.confirmes)], ' confirmé' + (k.confirmes > 1 ? 's' : '') + '.');
    var lede = $('#lede');
    lede.replaceChildren.apply(lede, parts.map(function (p) { if (typeof p === 'string') return document.createTextNode(p); var e = document.createElement(p[0]); e.textContent = p[1]; return e; }));
  }
  function de(n) { return /^[aeiouyéèêâîôûh]/i.test(n) ? 'd’' + n : 'de ' + n; }

  /* ---------- Graphiques SVG ---------- */
  var NS = 'http://www.w3.org/2000/svg';
  function el(tag, a, p) { var e = document.createElementNS(NS, tag); for (var k in a) e.setAttribute(k, a[k]); if (p) p.appendChild(e); return e; }
  function tw(svg, t) { var x = el('text', { x: -999, y: -999 }, svg); x.textContent = t; var w = x.getComputedTextLength ? x.getComputedTextLength() : t.length * 6.5; svg.removeChild(x); return w || t.length * 6.5; }
  function nice(v) { if (v <= 5) return 5; var p = Math.pow(10, Math.floor(Math.log10(v))), m = v / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p; }
  function tipOn(host, svg, W, ml, bw, n, lines) {
    var tip = document.createElement('div'); tip.className = 'tip'; tip.hidden = true; host.appendChild(tip);
    var at = function (ev) {
      var r = svg.getBoundingClientRect(), sx = (ev.clientX - r.left) * (W / r.width), i = Math.floor((sx - ml) / bw);
      if (i < 0 || i >= n) return; var L = lines(i);
      tip.replaceChildren(); var b = document.createElement('b'); b.textContent = L[0]; tip.appendChild(b);
      L.slice(1).forEach(function (l) { var x = document.createElement('div'); x.textContent = l; tip.appendChild(x); });
      tip.hidden = false; var w = tip.offsetWidth;
      tip.style.left = Math.max(0, Math.min(host.clientWidth - w, (ml + i * bw + bw / 2) * (r.width / W) - w / 2)) + 'px'; tip.style.top = '0px';
    };
    svg.addEventListener('pointerdown', at); svg.addEventListener('pointermove', at);
    svg.addEventListener('pointerleave', function (e) { if (e.pointerType === 'mouse') tip.hidden = true; });
  }
  // Barres « à suivre » (fond) et « vus » (devant), taux en ligne sur l'axe de droite (0–100 %)
  function suiviJours(host, data) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), H = 230, svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Contacts à suivre, vus et taux de suivi par jour' }, host);
    var maxN = nice(Math.max(1, Math.max.apply(null, data.map(function (d) { return d.aSuivre; }))));
    var ml = tw(svg, nf.format(maxN)) + 8, mr = tw(svg, '100 %') + 8, mt = 8, mb = 24, iw = W - ml - mr, ih = H - mt - mb, bw = iw / data.length;
    var yN = function (v) { return mt + ih - v / maxN * ih; }, yP = function (v) { return mt + ih - Math.min(1, v) * ih; };
    [0, .5, 1].forEach(function (p) {
      el('line', { x1: ml, x2: W - mr, y1: mt + ih - p * ih, y2: mt + ih - p * ih, style: 'stroke:var(--line)' }, svg);
      var a = el('text', { x: ml - 6, y: mt + ih - p * ih, 'text-anchor': 'end', dy: '0.32em' }, svg); a.textContent = nf.format(maxN * p);
      var b = el('text', { x: W - mr + 6, y: mt + ih - p * ih, dy: '0.32em' }, svg); b.textContent = pf.format(p);
    });
    el('line', { x1: ml, x2: W - mr, y1: yP(OBJ), y2: yP(OBJ), 'stroke-dasharray': '4 3', style: 'stroke:var(--ok)' }, svg);
    data.forEach(function (d, i) {
      var x = ml + i * bw;
      if (d.aSuivre) el('rect', { x: x + bw * .12, width: bw * .76, y: yN(d.aSuivre), height: ih - (yN(d.aSuivre) - mt), rx: 2, style: 'fill:var(--line-2)' }, svg);
      if (d.vus) el('rect', { x: x + bw * .25, width: bw * .5, y: yN(d.vus), height: ih - (yN(d.vus) - mt), rx: 2, style: 'fill:var(--ink-2)' }, svg);
      if (i % 5 === 4 || i === data.length - 1) { var t = el('text', { x: x + bw / 2, y: H - 6, 'text-anchor': 'middle' }, svg); t.textContent = fmtDate(d.date, true); }
    });
    el('polyline', { points: data.map(function (d, i) { return (ml + i * bw + bw / 2) + ',' + yP(d.taux); }).join(' '), fill: 'none', 'stroke-width': 2, style: 'stroke:var(--accent)' }, svg);
    tipOn(host, svg, W, ml, bw, data.length, function (i) { var d = data[i]; return [fmtDate(d.date), 'À suivre : ' + nf.format(d.aSuivre), 'Vus : ' + nf.format(d.vus), 'Taux : ' + pf.format(d.taux)]; });
  }
  var CIRC_COUL = { 'Foyer': 'var(--ink-2)', 'Hôpital': 'var(--men)', 'Funérailles': 'var(--dec)', 'Lieu du travail': 'var(--prob)', 'Non précisé': 'var(--line-2)' };
  function semaines(host, data) {
    host.replaceChildren();
    var keys = MVEC.CIRC.concat(['Non précisé']);
    var W = Math.max(280, host.clientWidth), H = 220, svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Contacts listés par semaine et circonstance' }, host);
    if (!data.length) { var t0 = el('text', { x: W / 2, y: H / 2, 'text-anchor': 'middle' }, svg); t0.textContent = 'Aucun contact listé sur la période.'; return; }
    var max = nice(Math.max(1, Math.max.apply(null, data.map(function (d) { return d.total; }))));
    var ml = tw(svg, nf.format(max)) + 8, mr = 6, mt = 8, mb = 24, iw = W - ml - mr, ih = H - mt - mb, bw = iw / data.length;
    [0, .5, 1].forEach(function (p) {
      el('line', { x1: ml, x2: W - mr, y1: mt + ih - p * ih, y2: mt + ih - p * ih, style: 'stroke:var(--line)' }, svg);
      var a = el('text', { x: ml - 6, y: mt + ih - p * ih, 'text-anchor': 'end', dy: '0.32em' }, svg); a.textContent = nf.format(max * p);
    });
    var pas = Math.max(1, Math.ceil(data.length / 6));
    data.forEach(function (d, i) {
      var x = ml + i * bw, y = mt + ih;
      keys.forEach(function (k) { var v = d[k]; if (!v) return; var h = v / max * ih; y -= h; el('rect', { x: x + bw * .15, width: bw * .7, y: y, height: h, style: 'fill:' + CIRC_COUL[k] }, svg); });
      if ((data.length - 1 - i) % pas === 0) { var t = el('text', { x: x + bw / 2, y: H - 6, 'text-anchor': 'middle' }, svg); t.textContent = fmtDate(d.semaine, true); }
    });
    tipOn(host, svg, W, ml, bw, data.length, function (i) {
      var d = data[i]; return ['Semaine du ' + fmtDate(d.semaine) + ' : ' + nf.format(d.total)].concat(keys.filter(function (k) { return d[k]; }).map(function (k) { return k + ' : ' + nf.format(d[k]); }));
    });
    var lg = $('#lg-sem');
    lg.replaceChildren.apply(lg, keys.map(function (k) { var s = document.createElement('span'), i = document.createElement('i'); i.style.background = CIRC_COUL[k]; s.append(i, k); return s; }));
  }
  function zonesBars(host, list) {
    host.replaceChildren();
    var items = list.filter(function (z) { return z.aSuivre || z.vus; }); if (!items.length) items = list;
    var W = Math.max(280, host.clientWidth), rowH = 30, H = items.length * rowH + 4;
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Contacts à suivre et vus par zone' }, host);
    var lw = Math.min(W * 0.4, Math.max.apply(null, items.map(function (i) { return tw(svg, i.lieu); })) + 10);
    var max = Math.max(1, Math.max.apply(null, items.map(function (i) { return i.aSuivre; })));
    var vw = tw(svg, nf.format(max) + ' / ' + nf.format(max)) + 10, iw = W - lw - vw;
    items.forEach(function (z, i) {
      var y = i * rowH + 2;
      var t = el('text', { x: lw - 8, y: y + rowH / 2, 'text-anchor': 'end', dy: '0.32em' }, svg); t.textContent = z.lieu;
      el('rect', { x: lw, y: y + 5, width: Math.max(0, z.aSuivre / max * iw), height: rowH - 10, rx: 3, style: 'fill:var(--line-2)' }, svg);
      if (z.vus) el('rect', { x: lw, y: y + 10, width: Math.max(2, z.vus / max * iw), height: rowH - 20, rx: 2, style: 'fill:var(--ink-2)' }, svg);
      var v = el('text', { x: lw + iw + 6, y: y + rowH / 2, dy: '0.32em', class: 'vlab' }, svg); v.textContent = nf.format(z.vus) + ' / ' + nf.format(z.aSuivre);
    });
  }
  function pyramide(host, data) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), rowH = 24, H = data.length * rowH + 4, svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Contacts listés par tranche d’âge et sexe' }, host);
    var max = Math.max(1, Math.max.apply(null, data.map(function (d) { return Math.max(d.h, d.f); })));
    var lw = Math.max.apply(null, data.map(function (d) { return tw(svg, d.tranche); })) + 12, vw = tw(svg, nf.format(max)) + 8;
    var half = (W - lw) / 2 - vw;
    data.forEach(function (d, i) {
      var y = i * rowH + 2, wh = d.h / max * half, wf = d.f / max * half;
      var t = el('text', { x: W / 2, y: y + rowH / 2, 'text-anchor': 'middle', dy: '0.32em' }, svg); t.textContent = d.tranche;
      if (wh) el('rect', { x: W / 2 - lw / 2 - wh, y: y + 4, width: wh, height: rowH - 8, rx: 3, style: 'fill:var(--men)' }, svg);
      if (wf) el('rect', { x: W / 2 + lw / 2, y: y + 4, width: wf, height: rowH - 8, rx: 3, style: 'fill:var(--women)' }, svg);
      var a = el('text', { x: W / 2 - lw / 2 - wh - 4, y: y + rowH / 2, 'text-anchor': 'end', dy: '0.32em', class: 'vlab' }, svg); a.textContent = d.h ? nf.format(d.h) : '';
      var b = el('text', { x: W / 2 + lw / 2 + wf + 4, y: y + rowH / 2, dy: '0.32em', class: 'vlab' }, svg); b.textContent = d.f ? nf.format(d.f) : '';
    });
  }
  function renderGraphiques() {
    if (state.tab !== 'graphiques') return;
    var r = state.result, f = state.filters;
    suiviJours($('#ch-jours'), r.jours);
    // Semaines de la période (dans la fenêtre de 30 semaines du classeur)
    semaines($('#ch-sem'), r.semaines.filter(function (s) { return s.semaine + 6 >= f.debut && s.semaine <= f.fin; }));
    zonesBars($('#ch-zones'), r.zones);
    pyramide($('#ch-age'), r.ageSexe);
  }

  function td(t, c) { var e = document.createElement('td'); e.textContent = t; if (c) e.className = c; return e; }
  function pctTd(v) { var e = document.createElement('td'); e.className = 'n'; if (v == null) { e.textContent = '–'; return e; } var s = document.createElement('span'); s.className = 'pct ' + cls(v); s.textContent = pf.format(v); e.appendChild(s); return e; }
  function rougeTd(v) { var e = td(nf.format(v), 'n'); if (v > 0) { e.style.color = 'var(--bad)'; e.style.fontWeight = '600'; } return e; }
  function lieuRows(host, list, tot, clic) {
    var row = function (z, total) {
      var tr = document.createElement('tr'); if (total) tr.className = 'total'; else if (z.aSuivre) tr.className = 'actif';
      tr.appendChild(td(z.lieu, 'first'));
      [z.listes, z.n7, z.aSuivre, z.vus].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); });
      tr.appendChild(pctTd(z.taux)); tr.appendChild(rougeTd(z.nonVus2)); tr.appendChild(rougeTd(z.jamais));
      [z.suspects, z.confirmes, z.sortis].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); });
      if (!total && clic) { tr.style.cursor = 'pointer'; tr.addEventListener('click', function () { clic(z); }); }
      return tr;
    };
    host.replaceChildren.apply(host, list.map(function (z) { return row(z); }).concat([row(tot, true)]));
  }
  function renderZones() {
    if (state.tab !== 'zones') return;
    var r = state.result, f = state.filters;
    lieuRows($('#tb-zones'), r.zones, r.totalZones, function (z) { setLieu(z.lieu, '', 'Filtre : zone ' + de(z.lieu) + '.'); });
    $('#t-aires').textContent = f.zone ? 'Situation par aire de santé – zone ' + de(f.zone) : 'Situation par aire de santé – toutes les zones';
    lieuRows($('#tb-aires'), r.aires, r.totalAires, function (a) { setLieu(f.zone, a.lieu, 'Filtre : aire ' + de(a.lieu) + '.'); });
    $('#note-zones').textContent = 'Fond bleu : contacts à suivre au jour de référence. Taux de suivi : vert ≥ 90 %, orange 80–89 %, rouge < 80 %. Non vus ≥ 2 j et jamais vus en rouge si > 0. Touchez une ligne pour filtrer.';
  }

  function renderAnalyse() {
    if (state.tab !== 'analyse') return;
    var A = state.result.analyse21, k = A.cles, f = state.filters;
    var lieu = f.aire ? 'aire ' + de(f.aire) : f.zone ? 'zone ' + de(f.zone) : 'toutes zones';
    $('#sub-21').textContent = '21 jours complets du ' + fmtDate(A.debut) + ' au ' + fmtDate(A.fin) + ' (se terminant la veille de la date de référence) – ' + lieu + '.';
    var host = $('#tiles-21');
    host.replaceChildren.apply(host, T21.map(function (t) {
      var note = t[0] === 'tauxMin' ? 'le ' + fmtDate(k.dateMin) : t[0] === 'tauxMax' ? 'le ' + fmtDate(k.dateMax) : null;
      return tile({ lab: t[1], fmt: t[2], col: t[4], hero: t[5] }, k[t[0]], note, [t[3], '', 'Lieu choisi ; 21 jours se terminant la veille de la date de référence']);
    }));
    var s = $('#tb-21sem');
    s.replaceChildren.apply(s, A.semaines.map(function (w) {
      var tr = document.createElement('tr'); tr.appendChild(td('Semaine ' + w.n, 'first'));
      tr.appendChild(td(fmtDate(w.du), 'n')); tr.appendChild(td(fmtDate(w.au), 'n'));
      tr.appendChild(td(nf.format(w.aSuivre), 'n')); tr.appendChild(td(nf.format(w.vus), 'n')); tr.appendChild(pctTd(w.aSuivre ? w.taux : null));
      tr.appendChild(td(nf.format(w.nouveaux), 'n')); tr.appendChild(td(nf.format(w.suspects), 'n')); return tr;
    }));
    var z = $('#tb-21zones'), zl = A.zones.filter(function (x) { return x.listes || x.aSuivre || x.enSuivi; });
    var zrow = function (x, total) {
      var tr = document.createElement('tr'); if (total) tr.className = 'total'; tr.appendChild(td(x.lieu, 'first'));
      [x.listes, x.aSuivre, x.vus].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); });
      tr.appendChild(pctTd(x.taux)); tr.appendChild(pctTd(x.taux7)); tr.appendChild(td(nf.format(x.enSuivi), 'n')); tr.appendChild(rougeTd(x.jamais)); return tr;
    };
    z.replaceChildren.apply(z, zl.map(function (x) { return zrow(x); }).concat([zrow(A.totalZones, true)]));
    var j = $('#tb-21jours');
    j.replaceChildren.apply(j, A.jours.slice().reverse().map(function (d) {
      var tr = document.createElement('tr'); tr.appendChild(td(fmtDate(d.date), 'first'));
      tr.appendChild(td(nf.format(d.aSuivre), 'n')); tr.appendChild(td(nf.format(d.vus), 'n')); tr.appendChild(pctTd(d.aSuivre ? d.taux : null));
      [d.nonVus, d.nouveaux, d.suspects, d.confirmes].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); }); return tr;
    }).concat([(function () {
      var t = A.total, tr = document.createElement('tr'); tr.className = 'total'; tr.appendChild(td('Total 21 j', 'first'));
      tr.appendChild(td(nf.format(t.aSuivre), 'n')); tr.appendChild(td(nf.format(t.vus), 'n')); tr.appendChild(pctTd(t.aSuivre ? t.taux : null));
      [t.nonVus, t.nouveaux, t.suspects, t.confirmes].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); }); return tr;
    })()]));
  }

  var CAT_IDX = { 'Jamais vu': 0, 'Non vu ≥ 2 j': 1, 'Non vu 24 h': 2 };
  function renderRelancer() {
    if (state.tab !== 'relancer') return;
    var f = state.filters, ref = f.ref, q = MVEC.norm($('#q-relancer').value), all = state.result.relancer;
    $$('#q-cat button').forEach(function (b) { b.setAttribute('aria-pressed', b.dataset.c === state.cat ? 'true' : 'false'); });
    var list = all.filter(function (r) { return !state.cat || r.cat === state.cat; });
    if (q) list = list.filter(function (r) { return MVEC.norm([r.id, r.epid, r.aire, r.zone, r.enr].join(' ')).indexOf(q) >= 0; });
    var n = function (c) { return all.filter(function (r) { return r.cat === c; }).length; };
    $('#sub-relancer').textContent = nf.format(all.length) + ' contact' + (all.length > 1 ? 's' : '') + ' en suivi non vu' + (all.length > 1 ? 's' : '') + ' le ' + fmtDate(ref) +
      ' : ' + nf.format(n('Jamais vu')) + ' jamais vus, ' + nf.format(n('Non vu ≥ 2 j')) + ' non vus depuis ≥ 2 j, ' + nf.format(n('Non vu 24 h')) + ' vus la veille. Retrouvez-les dans Tracker Capture par leur ID contact.';
    var host = $('#list-relancer');
    host.replaceChildren.apply(host, list.slice(0, state.nRel).map(function (r) {
      var d = document.createElement('div'); d.className = 'item';
      var top = document.createElement('div'); top.className = 'top-line';
      var id = document.createElement('span'); id.className = 'idn'; id.textContent = r.id || '(sans ID contact)';
      var nb = document.createElement('span'); nb.className = 'nb cat-' + CAT_IDX[r.cat]; nb.textContent = r.cat;
      top.append(id, nb);
      var m = document.createElement('div'); m.className = 'meta';
      m.textContent = 'J' + (ref - (r.deb || 0)) + ' sur 21 · ' + (r.dernVu ? 'dernière visite le ' + fmtDate(r.dernVu) : 'aucune visite') + ' · listé le ' + fmtDate(r.dEnr) + ' · ' + (r.zone || '–') + ' › ' + (r.aire || '–');
      var tags = document.createElement('div'); tags.className = 'tags';
      [r.epid ? 'Cas source ' + r.epid : 'Sans N° Epid', r.sexe ? r.sexe + (r.age != null ? ', ' + (r.age < 1 ? Math.round(r.age * 12) + ' mois' : Math.floor(r.age) + ' ans') : '') : null, r.circ, r.rel]
        .filter(Boolean).forEach(function (x) { var s = document.createElement('span'); s.textContent = x; tags.appendChild(s); });
      d.append(top, m, tags); return d;
    }));
    if (!list.length) { var p = document.createElement('p'); p.className = 'sub'; p.textContent = q || state.cat ? 'Aucun contact ne correspond.' : 'Tous les contacts à suivre ont été vus ce jour.'; host.appendChild(p); }
    $('#more-relancer').hidden = list.length <= state.nRel;
    $('#more-relancer').textContent = 'Afficher plus (' + nf.format(list.length - state.nRel) + ' restants)';
  }

  /* ---------- Onglets et feuilles ---------- */
  var ONGLETS = ['synthese', 'graphiques', 'zones', 'analyse', 'relancer'];
  function showTab(name) {
    if (ONGLETS.indexOf(name) < 0) name = 'synthese';
    state.tab = name; lsSet('mvec-onglet', name);
    ONGLETS.forEach(function (t) { $('#tab-' + t).hidden = t !== name; });
    $$('#tabs button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.tab === name ? 'true' : 'false'); });
    window.scrollTo(0, 0);
    if (state.result) { renderGraphiques(); renderZones(); renderAnalyse(); renderRelancer(); }
  }
  function openSheet(id) { closeSheets(true); $('#toast').hidden = true; $('#sheet-bg').hidden = false; $(id).hidden = false; history.pushState({ sheet: id }, ''); }
  function closeSheets(silent) {
    var open = $$('.sheet').some(function (s) { return !s.hidden; });
    $$('.sheet').forEach(function (s) { s.hidden = true; }); $('#sheet-bg').hidden = true;
    if (open && !silent && history.state && history.state.sheet) history.back();
  }
  window.addEventListener('popstate', function () { $$('.sheet').forEach(function (s) { s.hidden = true; }); $('#sheet-bg').hidden = true; });

  /* ---------- Événements ---------- */
  $$('#tabs button').forEach(function (b) { b.addEventListener('click', function () { showTab(b.dataset.tab); }); });
  $('#sheet-bg').addEventListener('click', function () { closeSheets(); });
  $$('[data-close]').forEach(function (b) { b.addEventListener('click', function () { closeSheets(); }); });
  $('#btn-open-empty').addEventListener('click', function () { $('#file').click(); });
  $('#btn-open').addEventListener('click', function () { closeSheets(); $('#file').click(); });
  $('#file').addEventListener('change', function (e) { openFile(e.target.files[0]); e.target.value = ''; });
  $('#btn-menu').addEventListener('click', function () {
    var d = state.data, last = lsGet('mvec-sync-last');
    $('#menu-info').textContent = (d ? 'Données : ' + d.fichier + ', reçues le ' + new Date(d.chargeLe).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + '. ' : 'Aucune donnée chargée. ') +
      (sync.key ? 'Mise à jour automatique activée' + (last ? ', dernière vérification réussie le ' + new Date(last).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) : '') + '.' : 'Mise à jour automatique désactivée.') +
      (function () { var e = lsGet('mvec-sync-err'); return e && e.m ? ' Dernier échec (' + new Date(e.t).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + ') : ' + e.m : ''; })() +
      (d && d.publie ? ' Publication du ' + new Date(d.publie).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + '.' : '');
    $('#btn-clear').hidden = !d; setSyncUI(); openSheet('#sheet-menu');
  });
  $('#btn-help').addEventListener('click', function () { openSheet('#sheet-help'); });
  $('#btn-sync').addEventListener('click', function () { syncNow(true); });
  $('#btn-sync-menu').addEventListener('click', function () { closeSheets(); syncNow(true); });
  $('#btn-auto').addEventListener('click', openPassword);
  $('#btn-auto-empty').addEventListener('click', function () {
    cleSuivi().then(function (ok) { if (ok) { setSyncUI(); toast('Mise à jour automatique activée avec le mot de passe de « Suivi MVE ».'); } else openPassword(); });
  });
  $('#btn-auto-off').addEventListener('click', function () { sync.key = null; store.del('cle'); setSyncUI(); closeSheets(); toast('Mise à jour automatique désactivée sur ce téléphone.'); });
  $('#form-pw').addEventListener('submit', function (e) { e.preventDefault(); var pw = $('#pw').value; if (!pw) { $('#pw-msg').textContent = 'Saisissez le mot de passe d’équipe.'; return; } connect(pw); });
  $('#pw-show').addEventListener('change', function () { $('#pw').type = $('#pw-show').checked ? 'text' : 'password'; });
  $('#btn-clear').addEventListener('click', function () {
    var b = $('#btn-clear');
    if (b.dataset.confirm !== '1') { b.dataset.confirm = '1'; b.textContent = 'Confirmer : effacer les données du téléphone'; return; }
    b.dataset.confirm = ''; b.textContent = 'Effacer les données de ce téléphone';
    store.del('dataset').then(function () {
      state.data = null; state.result = null; closeSheets();
      $$('.tab').forEach(function (t) { t.hidden = true; }); $('#empty').hidden = false; $('#tabs').hidden = true; $('#chips').hidden = true; $('#src').textContent = 'Aucune donnée chargée';
      toast('Données effacées de ce téléphone.');
    });
  });
  $('#f-zone').addEventListener('change', function () { fillAire($('#f-aire'), $('#f-zone').value, ''); });
  $('#q-periode').addEventListener('click', function (e) {
    var p = e.target.dataset && e.target.dataset.p; if (!p) return; draft.periode = p;
    var rows = state.data.rows, r = inputToS($('#f-ref').value) || state.filters.ref;
    if (p === 'tout') { $('#f-debut').value = sToInput(minN(rows)); $('#f-fin').value = sToInput(Math.max(today(), maxN(rows))); }
    else { $('#f-fin').value = sToInput(r); $('#f-debut').value = sToInput(r - (+p) + 1); }
  });
  $('#q-ref').addEventListener('click', function (e) { var m = e.target.dataset && e.target.dataset.r; if (!m) return; draft.refMode = m; $('#f-ref').value = sToInput(today() - (+m)); });
  $('#f-debut').addEventListener('input', function () { draft.periode = 'fixe'; });
  $('#f-fin').addEventListener('input', function () { draft.periode = 'fixe'; });
  $('#f-ref').addEventListener('input', function () { draft.refMode = 'fixe'; });
  $('#btn-apply').addEventListener('click', applyFilters);
  $('#btn-reset').addEventListener('click', function () { lsSet('mvec-filtres', {}); closeSheets(); state.filters = resolveFilters(); render(); toast('Filtres réinitialisés.'); });
  $('#q-relancer').addEventListener('input', function () { state.nRel = 50; renderRelancer(); });
  $('#q-cat').addEventListener('click', function (e) { var c = e.target.dataset ? e.target.dataset.c : undefined; if (c === undefined) return; state.cat = c; state.nRel = 50; renderRelancer(); });
  $('#more-relancer').addEventListener('click', function () { state.nRel += 50; renderRelancer(); });
  var rz;
  window.addEventListener('resize', function () { clearTimeout(rz); rz = setTimeout(function () { if (state.result) renderGraphiques(); }, 150); });
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible' && Date.now() - (lsGet('mvec-sync-last') || 0) > 2 * 60000) syncNow(false); });
  window.addEventListener('online', function () { syncNow(false); });
  var installEvt = null;
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); installEvt = e; $('#btn-install').hidden = false; });
  $('#btn-install').addEventListener('click', function () { if (!installEvt) return; closeSheets(); installEvt.prompt(); installEvt.userChoice.finally(function () { installEvt = null; $('#btn-install').hidden = true; }); });

  /* ---------- Démarrage ---------- */
  var cle = store.get('cle').then(function (c) { if (c && c.key) { sync.key = c.key; sync.salt = c.salt; sync.iter = c.iter; } }).catch(function () { /* aucune clé */ });
  cle.then(function () {
    setSyncUI();
    return store.get('dataset').then(function (d) {
      if (d && d.contacts) setData(setFromTables(d.contacts, d.suivis || [], { fichier: d.fichier, publie: d.publie, chargeLe: d.chargeLe }), false);
    }).catch(function () { /* première ouverture */ });
  }).then(function () {
    if (sync.key) return syncNow(false);
    return cleSuivi().then(function (ok) { if (ok) { setSyncUI(); toast('Mise à jour automatique activée avec le mot de passe de « Suivi MVE ».'); } });
  });
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(function () { /* hors ligne indisponible */ });
})();
