/* Interface de l'application « Qualité Contacts MVE ». Calculs : engine-qc.js
   Données : les mêmes que « Contacts MVE » (../contacts/contacts.enc, même mot de passe). */
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };
  var state = { data: null, filters: null, tab: 'synthese', result: null, nC: 50, nF: 50 };
  var nf = new Intl.NumberFormat('fr-FR');
  var pf = new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 1 });
  var df1 = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 });

  /* ---------- Dates ---------- */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function sToDate(s) { return MVEQC.serialToUTC(s); }
  function fmtDate(s, short) { if (s == null) return '–'; var d = sToDate(s); return pad(d.getUTCDate()) + '/' + pad(d.getUTCMonth() + 1) + (short ? '' : '/' + d.getUTCFullYear()); }
  function sToInput(s) { var d = sToDate(s); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function inputToS(v) { return v ? MVEQC.toSerial(v) : null; }
  function today() { return Math.floor(MVEQC.nowSerial()); }

  /* ---------- Stockage local ---------- */
  function idb(name) {
    return new Promise(function (res, rej) {
      var rq = indexedDB.open(name || 'qualite-contacts-mve', 1);
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
  function setFromTables(contacts, suivis, meta, seuils) {
    var out = MVEQC.normalize(contacts, suivis, seuils);
    return Object.assign({ rows: out.rows, fiches: out.fiches, extrait: out.extrait, extraitS: out.extraitS, params: out.params, charge: out.charge, nbFiches: out.nbFiches, seuils: seuils || null, chargeLe: Date.now() }, meta);
  }
  // Paramètres!D2:D6 (délai de listage, âge maximal, retard de saisie, seuils vert / orange)
  function lireSeuils(wb) {
    var p = wb.Sheets['Paramètres']; if (!p) return null;
    var s = {};
    ['DelaiList', 'AgeMax', 'RetSaisie', 'Vert', 'Orange'].forEach(function (k, i) { var c = p['D' + (i + 2)]; if (c && typeof c.v === 'number') s[k] = c.v; });
    return s;
  }
  /* Seules les colonnes A à V de « Données » et A à J de « Suivis » sont gardées : les noms (colonnes W à Y) sont écartés dès la lecture. */
  function readWorkbook(buf, name, garder) {
    busy(true);
    return new Promise(function (r) { setTimeout(r, 30); }).then(function () {
      var wb = XLSX.read(buf, { type: 'array', sheets: ['Données', 'Suivis', 'Paramètres'], cellDates: false, cellFormula: false, cellHTML: false, dense: false });
      if (!wb.Sheets['Données']) throw new Error('Ce fichier ne contient pas la feuille « Données ».');
      var T = function (n) { return wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: false }) : []; };
      var contacts = MVEQC.epurer(T('Données'), MVEQC.COLS), suivis = MVEQC.epurer(T('Suivis'), MVEQC.SCOLS);
      var data = setFromTables(contacts, suivis, { fichier: name || 'fichier.xlsx' }, lireSeuils(wb));
      return store.set('dataset', { contacts: contacts, suivis: suivis, seuils: data.seuils, fichier: data.fichier, publie: null, chargeLe: data.chargeLe })
        .catch(function () { /* stockage plein */ }).then(function () { return data; });
    }).then(function (d) { busy(false); setData(d, !garder); toast(nf.format(d.charge) + ' contacts chargés.'); })
      .catch(function (e) { busy(false); toast('Lecture impossible : ' + (e && e.message ? e.message : e), 6000); });
  }
  // Programme d'ordinateur (dossier de classeurs) : lecture directe d'un classeur sans passer par le sélecteur de fichier
  window.MVE_lireClasseur = function (buf, name, garder) { return readWorkbook(buf, name, garder); };
  function openFile(f) { if (f) f.arrayBuffer().then(function (b) { return readWorkbook(new Uint8Array(b), f.name); }); }

  /* ---------- Mise à jour automatique (fichier de « Contacts MVE », même mot de passe que Suivi MVE) ---------- */
  var REMOTE = '../contacts/contacts.enc', sync = { key: null, salt: null, iter: 0, running: false };
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
    var sl = (cur && cur.seuils) || null;
    var d = setFromTables(p.contacts, p.suivis, { fichier: 'mise à jour automatique', publie: p.publie }, sl);
    return store.set('dataset', { contacts: p.contacts, suivis: p.suivis, seuils: sl, fichier: d.fichier, publie: p.publie, chargeLe: d.chargeLe })
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
      lsSet('mvqc-sync-last', Date.now()); lsSet('mvqc-sync-err', null);
      if (ch && !manual) toast('Données mises à jour automatiquement.'); else if (manual) toast(ch ? 'Données mises à jour.' : 'Les données sont déjà à jour.');
    }).catch(function (e) {
      var m = (e && e.message) || String(e);
      lsSet('mvqc-sync-err', { m: m, t: Date.now() });
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
            .then(function () { lsSet('mvqc-sync-last', Date.now()); setSyncUI(); toast('Mise à jour automatique activée.'); });
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
    var s = lsGet('mvqc-filtres') || {}, rows = state.data.rows, t = today();
    var f = { zone: s.zone || '', aire: s.aire || '', periode: s.periode || 'tout', debutFixe: s.debutFixe, finFixe: s.finFixe, refMode: s.refMode || '0', refFixe: s.refFixe };
    if (f.zone && MVEQC.zonesOf(rows).indexOf(f.zone) < 0) { f.zone = ''; f.aire = ''; }
    f.ref = f.refMode === 'fixe' ? f.refFixe : t - (+f.refMode || 0);
    if (f.periode === 'fixe') { f.debut = f.debutFixe; f.fin = f.finFixe; }
    else if (f.periode === 'tout') { f.debut = minN(rows); f.fin = Math.max(t, maxN(rows)); }
    else { f.fin = f.ref; f.debut = f.ref - (+f.periode) + 1; }
    return f;
  }
  function saveFilters(f) { lsSet('mvqc-filtres', { zone: f.zone, aire: f.aire, periode: f.periode, debutFixe: f.debutFixe, finFixe: f.finFixe, refMode: f.refMode, refFixe: f.refFixe }); }
  function fillZone(sel, v) { sel.replaceChildren(new Option('Toutes les zones', '')); MVEQC.zonesOf(state.data.rows).forEach(function (z) { sel.appendChild(new Option(z, z)); }); sel.value = v || ''; }
  function fillAire(sel, z, v) {
    sel.replaceChildren(new Option(z ? 'Toutes les aires de ' + z : 'Toutes les aires', ''));
    MVEQC.airesOf(state.data.rows, z).forEach(function (a) { sel.appendChild(new Option(a, a)); }); sel.value = v || ''; if (sel.value !== (v || '')) sel.value = '';
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
    saveFilters(f); closeSheets(); state.filters = resolveFilters(); state.nC = 50; state.nF = 50; render();
  }
  function setLieu(zone, aire, msg) {
    var s = lsGet('mvqc-filtres') || {}; s.zone = zone; s.aire = aire; lsSet('mvqc-filtres', s);
    state.filters = resolveFilters(); render(); toast(msg);
  }

  /* ---------- Rendu ---------- */
  function setData(d, fresh) {
    state.data = d; state.filters = resolveFilters();
    $('#empty').hidden = true; $('#tabs').hidden = false; $('#chips').hidden = false;
    showTab(fresh ? 'synthese' : (lsGet('mvqc-onglet') || 'synthese'));
    render();
  }
  function cls(v) { var P = state.data.params; return v == null ? '' : v >= P.Vert ? 'good' : v >= P.Orange ? 'mid' : 'bad'; }
  function render() {
    var d = state.data, f = state.filters;
    state.result = MVEQC.compute(d, f);
    var ex = d.extrait != null ? sToDate(d.extrait) : null;
    $('#src').textContent = 'DHIS2 SNIS · extrait le ' + (ex ? pad(ex.getUTCDate()) + '/' + pad(ex.getUTCMonth() + 1) + '/' + ex.getUTCFullYear() + ' à ' + pad(ex.getUTCHours()) + 'h' + pad(ex.getUTCMinutes()) : '(date inconnue)') +
      ' · ' + nf.format(d.charge) + ' contacts, ' + nf.format(d.nbFiches) + ' fiches de suivi';
    var age = d.extrait != null ? (Date.now() - (sToDate(d.extrait).getTime() + new Date().getTimezoneOffset() * 60000)) / 864e5 : null;
    $('#src').classList.toggle('stale', age != null && age > 2);
    renderChips(); renderSynthese(); renderGraphiques(); renderZones(); renderContacts(); renderFiches();
  }
  function renderChips() {
    var f = state.filters, c = $('#chips');
    var lieu = f.aire ? f.aire + ' (' + (f.zone || 'aire') + ')' : (f.zone || 'Toutes les zones');
    var per = f.periode === 'tout' ? 'Depuis le début' : (/^\d+$/.test(f.periode) ? f.periode + ' derniers jours' : fmtDate(f.debut) + ' → ' + fmtDate(f.fin));
    c.replaceChildren.apply(c, [['Lieu', lieu], ['Période', per], ['Réf.', fmtDate(f.ref)]].map(function (it) {
      var b = document.createElement('button'); b.className = 'chip'; var bb = document.createElement('b'); bb.textContent = it[1];
      b.append(it[0] + ' : ', bb); b.addEventListener('click', openFilters); return b;
    }));
  }

  /* Indicateurs (dictionnaire du classeur) : [clé, libellé, format, définition, lecture, filtres, coloration, hero] */
  var P_ = 'Zone + aire + période (date d’enregistrement du contact)', F_ = 'Zone + aire + période (date de la fiche de suivi)', R_ = 'Zone + aire + date de référence (toutes périodes)';
  function tilesDef() {
    var P = state.data.params, V = 'Vert ≥ ' + pf.format(P.Vert) + ', orange ≥ ' + pf.format(P.Orange) + ', rouge en dessous.';
    return {
      id: [
        ['listes', 'Contacts listés', 'n', 'Contacts enregistrés dans DHIS2 pendant la période, dans la zone et l’aire choisies.', 'Dénominateur des indicateurs de cette section.', P_, null, true],
        ['pSans', 'Contacts sans anomalie', 'p', 'Contacts sans aucune variable manquante ni incohérence ÷ contacts listés (score global).', V, P_, 'pct'],
        ['compl', 'Complétude des variables', 'p', '1 − (variables manquantes ÷ variables contrôlées). 9 variables : ID contact, N° Epid du cas source, date du dernier contact, sexe, relation avec le cas, circonstance, type de contact, notifié par, âge. « Inconnue » n’est pas comptée comme manquante.', V + ' Détail par variable dans « Graphiques ».', P_, 'pct'],
        ['pLies', 'Contacts liés à un cas source', 'p', 'Contacts avec N° Epid du cas source renseigné ÷ contacts listés.', 'Cible : 100 % (traçabilité de la chaîne de transmission).', P_, 'pct'],
        ['incoh', 'Contacts avec incohérence', 'n', '7 contrôles : dernier contact après l’enregistrement ; listé > ' + P.DelaiList + ' j après le dernier contact ; N° Epid non standard ; âge invalide (< 0, > ' + P.AgeMax + ' ans ou mois > 11) ; ID contact en double ; issue sans date ou datée avant le début du suivi ; symptômes signalés sans issue.', 'Cible : 0. Liste dans l’onglet « Contacts ».', P_, 'zero'],
        ['epidNS', 'N° Epid non standard', 'n', 'N° Epid du cas source qui ne suit pas le format RDC-HUE-XXX-26-0000000-MVE (26 caractères, sans espace).', 'Cible : 0. Ces contacts sont difficiles à relier à leur cas.', P_, 'zero']
      ],
      fi: [
        ['fiches', 'Fiches de suivi', 'n', 'Fiches de suivi quotidien (une ligne par contact et par jour) datées dans la période.', 'Dénominateur de « Fiches sans anomalie ».', F_, null, true],
        ['pFichesSans', 'Fiches sans anomalie', 'p', 'Fiches sans anomalie ÷ fiches. 10 contrôles : fiche sans contact ; hors J1–J21 ; datée avant l’enregistrement ; après l’issue ; date future ; « vu » non renseigné ; non vu sans raison ; vu avec raison de non-visite ; symptômes non renseignés (vu) ; plusieurs fiches le même jour.', V, F_, 'pct'],
        ['complFiches', 'Complétude des fiches de suivi', 'p', 'Jours saisis (plafonnés) ÷ jours attendus. Jours attendus = du début du suivi au plus tôt de J21, de la date de réf. et de la date d’issue.', V, P_, 'pct'],
        ['sansFiche', 'Contacts sans aucune fiche', 'n', 'Contacts listés dont le suivi a commencé (jours attendus > 0) mais avec 0 jour saisi.', 'Cible : 0. Contacts non suivis ou fiches non saisies.', P_, 'zero'],
        ['doubles', 'Fiches en double', 'n', 'Fiches dont le couple contact + date apparaît plus d’une fois.', 'Cible : 0.', F_, 'zero'],
        ['hors', 'Fiches hors période J1–J21', 'n', 'Fiches dont le jour de suivi est < 1 ou > 21.', 'Cible : 0. Souvent une erreur de date ou un contact enregistré après le début réel du suivi.', F_, 'zero']
      ],
      ret: [
        ['enSuivi', 'Contacts en suivi', 'n', 'Statut « En suivi » : listé, suivi commencé, J21 non dépassé et pas d’issue enregistrée.', 'Charge de suivi active.', R_, null, true],
        ['retard', 'En suivi sans fiche récente', 'n', 'Aucune fiche, ou dernière fiche plus ancienne que ' + P.RetSaisie + ' j avant la date de réf.', 'Cible : 0. À relancer en priorité.', R_, 'zero'],
        ['pAJour', 'Part des contacts en suivi à jour', 'p', '1 − (contacts en retard ÷ contacts en suivi).', V, R_, 'pct'],
        ['delai', 'Délai dernier contact → listage (j)', 'd1', 'Moyenne de (date d’enregistrement − date du dernier contact) ; délais négatifs et dates manquantes exclus.', 'Plus court = mieux. > ' + P.DelaiList + ' j = « listé trop tard ».', P_],
        ['derniereFiche', 'Dernière fiche saisie', 'date', 'Date de suivi la plus récente, sans dépasser la date d’extraction.', 'Doit être proche d’hier ; sinon la saisie accuse un retard.', 'Toutes les fiches'],
        ['fraicheur', 'Fraîcheur des données (h)', 'n', 'Heures écoulées depuis la dernière extraction DHIS2.', 'Plus bas = mieux ; les données sont publiées toutes les 30 minutes.', 'Toutes les données']
      ]
    };
  }
  function fmtVal(v, t) {
    if (v == null) return '–';
    if (t === 'p') return pf.format(v); if (t === 'd1') return df1.format(v); if (t === 'date') return fmtDate(v);
    return nf.format(v);
  }
  function tile(t, v, def) {
    var b = document.createElement('button');
    b.className = 'tile' + (t.hero ? ' hero' : '');
    if (t.col === 'pct' && v != null) b.className += ' q-' + cls(v);
    else if (t.col === 'zero' && v != null) b.className += v === 0 ? ' q-good' : ' q-bad';
    else if (t.key === 'derniereFiche' && v != null) b.className += state.filters.ref - v <= 1 ? ' q-good' : state.filters.ref - v <= 3 ? ' q-mid' : ' q-bad';
    else if (t.key === 'fraicheur' && v != null) b.className += v <= 24 ? ' q-good' : v <= 48 ? ' q-mid' : ' q-bad';
    var lab = document.createElement('span'); lab.className = 'lab'; lab.textContent = t.lab;
    var val = document.createElement('span'); val.className = 'val'; val.textContent = fmtVal(v, t.fmt);
    b.append(lab, val);
    b.addEventListener('click', function () {
      $('#t-def').textContent = t.lab; $('#def-val').textContent = fmtVal(v, t.fmt);
      $('#def-def').textContent = def[0]; $('#def-int').textContent = def[1] || ''; $('#def-act').textContent = def[2] || ''; openSheet('#sheet-def');
    });
    return b;
  }
  function renderSynthese() {
    var k = state.result.k, f = state.filters, T = tilesDef();
    Object.keys(T).forEach(function (g) {
      var host = $('#tiles-' + g);
      host.replaceChildren.apply(host, T[g].map(function (t) { return tile({ key: t[0], lab: t[1], fmt: t[2], col: t[6], hero: t[7] }, k[t[0]], [t[3], t[4], t[5]]); }));
    });
    $('#sub-id').textContent = 'Contacts enregistrés du ' + fmtDate(f.debut) + ' au ' + fmtDate(f.fin) + '.';
    $('#sub-fi').textContent = 'Fiches datées du ' + fmtDate(f.debut) + ' au ' + fmtDate(f.fin) + ' (une ligne par contact et par jour).';
    $('#sub-ret').textContent = 'Situation au ' + fmtDate(f.ref) + ', toutes périodes de listage.';
    var lieu = f.aire ? 'l’aire ' + de(f.aire) : f.zone ? 'la zone ' + de(f.zone) : 'le Haut-Uélé';
    var parts = ['Dans ' + lieu + ', ', ['b', nf.format(k.listes)], ' contact' + (k.listes > 1 ? 's' : '') + ' listé' + (k.listes > 1 ? 's' : '') + ' sur la période'];
    if (k.pSans != null) parts.push(', dont ', ['b', pf.format(k.pSans)], ' sans anomalie');
    parts.push('. ');
    if (k.complFiches != null) parts.push('Complétude des fiches de suivi : ', ['b', pf.format(k.complFiches)], ' ; ');
    parts.push(['b', nf.format(k.sansFiche)], ' contact' + (k.sansFiche > 1 ? 's' : '') + ' sans aucune fiche. ');
    if (k.enSuivi) parts.push(['b', nf.format(k.retard)], ' des ', ['b', nf.format(k.enSuivi)], ' contacts en suivi n’ont pas de fiche récente.');
    var lede = $('#lede');
    lede.replaceChildren.apply(lede, parts.map(function (p) { if (typeof p === 'string') return document.createTextNode(p); var e = document.createElement(p[0]); e.textContent = p[1]; return e; }));
  }
  function de(n) { return /^[aeiouyéèêâîôûh]/i.test(n) ? 'd’' + n : 'de ' + n; }

  /* ---------- Graphiques SVG ---------- */
  var NS = 'http://www.w3.org/2000/svg';
  function el(tag, a, p) { var e = document.createElementNS(NS, tag); for (var k in a) e.setAttribute(k, a[k]); if (p) p.appendChild(e); return e; }
  function tw(svg, t) { var x = el('text', { x: -999, y: -999 }, svg); x.textContent = t; var w = x.getComputedTextLength ? x.getComputedTextLength() : t.length * 6.5; svg.removeChild(x); return w || t.length * 6.5; }
  function nice(v) { if (v <= 5) return 5; var p = Math.pow(10, Math.floor(Math.log10(v))), m = v / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p; }
  function hbars(host, items, o) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), rowH = 26, H = items.length * rowH + 4;
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': o.aria }, host);
    var lw = Math.min(W * 0.55, Math.max.apply(null, items.map(function (i) { return tw(svg, i.label); })) + 10);
    var max = o.max || Math.max.apply(null, items.map(function (i) { return i.value || 0; }).concat([1]));
    var vw = tw(svg, o.fmt(max)) + 10, iw = W - lw - vw;
    items.forEach(function (it, i) {
      var y = i * rowH + 2, w = it.value == null ? 0 : (it.value / max) * iw;
      var t = el('text', { x: lw - 8, y: y + rowH / 2 - 1, 'text-anchor': 'end', dy: '0.32em' }, svg), lab = it.label;
      t.textContent = lab;
      while (lab.length > 4 && t.getComputedTextLength && t.getComputedTextLength() > lw - 10) { lab = lab.slice(0, -2); t.textContent = lab + '…'; }
      if (lab !== it.label) { var tt = el('title', {}, t); tt.textContent = it.label; }
      el('rect', { x: lw, y: y + 5, width: iw, height: rowH - 12, rx: 3, style: 'fill:var(--bg)' }, svg);
      if (w > 0) el('rect', { x: lw, y: y + 5, width: Math.max(2, w), height: rowH - 12, rx: 3, style: 'fill:' + (it.color || o.color) }, svg);
      var v = el('text', { x: lw + iw + 6, y: y + rowH / 2 - 1, dy: '0.32em', class: 'vlab' }, svg); v.textContent = it.value == null ? 's.o.' : o.fmt(it.value);
    });
  }
  function couleur(v) { var P = state.data.params; return v == null ? 'var(--line)' : v >= P.Vert ? 'var(--ok)' : v >= P.Orange ? 'var(--warn)' : 'var(--bad)'; }
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
  // Barres : contacts listés ; lignes : % sans anomalie et complétude des fiches (axe de droite 0–100 %)
  function hebdo(host, data) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), H = 220, svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Évolution hebdomadaire de la qualité' }, host);
    var maxN = nice(Math.max(1, Math.max.apply(null, data.map(function (d) { return d.listes; }))));
    var ml = tw(svg, nf.format(maxN)) + 8, mr = tw(svg, '100 %') + 8, mt = 8, mb = 24, iw = W - ml - mr, ih = H - mt - mb, bw = iw / data.length;
    var yN = function (v) { return mt + ih - v / maxN * ih; }, yP = function (v) { return mt + ih - Math.min(1, v) * ih; };
    [0, .5, 1].forEach(function (p) {
      el('line', { x1: ml, x2: W - mr, y1: mt + ih - p * ih, y2: mt + ih - p * ih, style: 'stroke:var(--line)' }, svg);
      var a = el('text', { x: ml - 6, y: mt + ih - p * ih, 'text-anchor': 'end', dy: '0.32em' }, svg); a.textContent = nf.format(maxN * p);
      var b = el('text', { x: W - mr + 6, y: mt + ih - p * ih, dy: '0.32em' }, svg); b.textContent = pf.format(p);
    });
    data.forEach(function (d, i) {
      var x = ml + i * bw;
      if (d.listes) el('rect', { x: x + bw * .15, width: bw * .7, y: yN(d.listes), height: ih - (yN(d.listes) - mt), rx: 2, style: 'fill:var(--line-2)' }, svg);
      if (i % 3 === 2 || i === data.length - 1) { var t = el('text', { x: x + bw / 2, y: H - 6, 'text-anchor': 'middle' }, svg); t.textContent = fmtDate(d.semaine, true); }
    });
    [['pSans', 'var(--ink-2)'], ['complFiches', 'var(--accent)']].forEach(function (s) {
      var pts = [];
      data.forEach(function (d, i) { if (d[s[0]] != null) pts.push((ml + i * bw + bw / 2) + ',' + yP(d[s[0]])); });
      if (pts.length) el('polyline', { points: pts.join(' '), fill: 'none', 'stroke-width': 2, style: 'stroke:' + s[1] }, svg);
      data.forEach(function (d, i) { if (d[s[0]] != null) el('circle', { cx: ml + i * bw + bw / 2, cy: yP(d[s[0]]), r: 2.5, style: 'fill:' + s[1] }, svg); });
    });
    tipOn(host, svg, W, ml, bw, data.length, function (i) {
      var d = data[i]; return ['Semaine du ' + fmtDate(d.semaine), 'Contacts listés : ' + nf.format(d.listes), '% sans anomalie : ' + (d.pSans == null ? '–' : pf.format(d.pSans)), 'Complétude des fiches : ' + (d.complFiches == null ? '–' : pf.format(d.complFiches))];
    });
  }
  function zonesBars(host, list) {
    host.replaceChildren();
    var items = list.filter(function (z) { return z.listes; }); if (!items.length) { host.textContent = 'Aucun contact listé sur la période.'; return; }
    var W = Math.max(280, host.clientWidth), rowH = 34, H = items.length * rowH + 4;
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Qualité par zone' }, host);
    var lw = Math.min(W * 0.4, Math.max.apply(null, items.map(function (i) { return tw(svg, i.lieu); })) + 10), vw = tw(svg, '100 %') + 10, iw = W - lw - vw;
    items.forEach(function (z, i) {
      var y = i * rowH + 2;
      var t = el('text', { x: lw - 8, y: y + rowH / 2, 'text-anchor': 'end', dy: '0.32em' }, svg); t.textContent = z.lieu;
      [['pSans', 'var(--ink-2)', 4], ['complFiches', 'var(--accent)', 17]].forEach(function (s) {
        var v = z[s[0]] || 0;
        el('rect', { x: lw, y: y + s[2], width: iw, height: 11, rx: 2, style: 'fill:var(--bg)' }, svg);
        if (v) el('rect', { x: lw, y: y + s[2], width: Math.max(2, v * iw), height: 11, rx: 2, style: 'fill:' + s[1] }, svg);
        var tv = el('text', { x: lw + iw + 6, y: y + s[2] + 5.5, dy: '0.32em', class: 'vlab' }, svg); tv.textContent = z[s[0]] == null ? '–' : pf.format(z[s[0]]);
      });
    });
  }
  function renderGraphiques() {
    if (state.tab !== 'graphiques') return;
    var r = state.result;
    hbars($('#ch-var'), r.variables.map(function (v) { return { label: v.lib, value: v.pct, color: couleur(v.pct) }; }), { max: 1, fmt: function (v) { return pf.format(v); }, aria: 'Complétude des variables' });
    hebdo($('#ch-sem'), r.semaines);
    hbars($('#ch-types'), r.types.slice().sort(function (a, b) { return b.n - a.n; }).map(function (t) { return { label: t.lib, value: t.n, color: /^Fiche/.test(t.lib) ? 'var(--accent)' : 'var(--bad)' }; }),
      { fmt: function (v) { return nf.format(v); }, aria: 'Anomalies par type' });
    zonesBars($('#ch-zones'), r.zones);
  }

  function td(t, c) { var e = document.createElement('td'); e.textContent = t; if (c) e.className = c; return e; }
  function pctTd(v) { var e = document.createElement('td'); e.className = 'n'; if (v == null) { e.textContent = '–'; return e; } var s = document.createElement('span'); s.className = 'pct ' + cls(v); s.textContent = pf.format(v); e.appendChild(s); return e; }
  function rougeTd(v) { var e = td(nf.format(v), 'n'); if (v > 0) { e.style.color = 'var(--bad)'; e.style.fontWeight = '600'; } return e; }
  function lieuRows(host, list, tot, clic) {
    var row = function (z, total) {
      var tr = document.createElement('tr'); if (total) tr.className = 'total'; else if (z.listes) tr.className = 'actif';
      tr.appendChild(td(z.lieu, 'first')); tr.appendChild(td(nf.format(z.listes), 'n'));
      tr.appendChild(pctTd(z.listes || total ? z.pSans : null)); tr.appendChild(pctTd(z.listes || total ? z.compl : null)); tr.appendChild(rougeTd(z.incoh));
      tr.appendChild(pctTd(z.complFiches)); tr.appendChild(rougeTd(z.sansFiche)); tr.appendChild(td(nf.format(z.fiches), 'n'));
      tr.appendChild(rougeTd(z.fichesAnom)); tr.appendChild(rougeTd(z.doubles)); tr.appendChild(rougeTd(z.retard));
      if (!total && clic) { tr.style.cursor = 'pointer'; tr.addEventListener('click', function () { clic(z); }); }
      return tr;
    };
    host.replaceChildren.apply(host, list.map(function (z) { return row(z); }).concat([row(tot, true)]));
  }
  function renderZones() {
    if (state.tab !== 'zones') return;
    var r = state.result, f = state.filters, P = state.data.params;
    lieuRows($('#tb-zones'), r.zones, r.totalZones, function (z) { setLieu(z.lieu, '', 'Filtre : zone ' + de(z.lieu) + '.'); });
    $('#t-aires').textContent = f.zone ? 'Qualité par aire de santé – zone ' + de(f.zone) : 'Qualité par aire de santé – toutes les zones';
    lieuRows($('#tb-aires'), r.aires, r.totalAires, function (a) { setLieu(f.zone, a.lieu, 'Filtre : aire ' + de(a.lieu) + '.'); });
    $('#note-zones').textContent = 'Fond bleu : contacts listés dans la période. Pourcentages : vert ≥ ' + pf.format(P.Vert) + ', orange ≥ ' + pf.format(P.Orange) + ', rouge en dessous. Ligne Total des zones = indicateurs du lieu choisi (comme le classeur). Touchez une ligne pour filtrer.';
  }

  function item(idTxt, nbTxt, metaTxt, tagsList) {
    var d = document.createElement('div'); d.className = 'item';
    var top = document.createElement('div'); top.className = 'top-line';
    var id = document.createElement('span'); id.className = 'idn'; id.textContent = idTxt;
    var nb = document.createElement('span'); nb.className = 'nb'; nb.textContent = nbTxt;
    top.append(id, nb);
    var m = document.createElement('div'); m.className = 'meta'; m.textContent = metaTxt;
    var tags = document.createElement('div'); tags.className = 'tags';
    tagsList.forEach(function (x) { var s = document.createElement('span'); s.className = /^manque/.test(x) ? '' : 'inc'; s.textContent = x; tags.appendChild(s); });
    d.append(top, m, tags); return d;
  }
  function idContact(r) { return r ? (r.id && String(r.id).trim() ? String(r.id) : '(sans ID) ' + r.enr) : null; }
  function renderContacts() {
    if (state.tab !== 'contacts') return;
    var q = MVEQC.norm($('#q-contacts').value), all = state.result.aCorriger;
    var list = q ? all.filter(function (r) { return MVEQC.norm([r.id, r.enr, r.epid, r.zone, r.aire].concat(r.detail).join(' ')).indexOf(q) >= 0; }) : all;
    $('#sub-contacts').textContent = nf.format(all.length) + ' contact' + (all.length > 1 ? 's' : '') + ' de la période avec au moins une anomalie (zone et aire choisies), les plus récents d’abord. « (sans ID) » : identifiant interne DHIS2.';
    var host = $('#list-contacts');
    host.replaceChildren.apply(host, list.slice(0, state.nC).map(function (r) {
      return item(idContact(r), r.nbAnom + ' anomalie' + (r.nbAnom > 1 ? 's' : ''),
        'Listé le ' + fmtDate(r.dEnr) + ' · ' + (r.zone || '–') + ' › ' + (r.aire || '–') + ' · cas source ' + (r.epid && String(r.epid).trim() ? r.epid : '–'), r.detail);
    }));
    if (!list.length) { var p = document.createElement('p'); p.className = 'sub'; p.textContent = q ? 'Aucun contact ne correspond.' : 'Aucun contact à corriger sur cette période.'; host.appendChild(p); }
    $('#more-contacts').hidden = list.length <= state.nC;
    $('#more-contacts').textContent = 'Afficher plus (' + nf.format(list.length - state.nC) + ' restants)';
  }
  function renderFiches() {
    if (state.tab !== 'fiches') return;
    var q = MVEQC.norm($('#q-fiches').value), all = state.result.fichesACorriger;
    var list = q ? all.filter(function (x) { return MVEQC.norm([x.c ? x.c.id : '', x.enr, x.zone, x.aire].concat(x.detail).join(' ')).indexOf(q) >= 0; }) : all;
    $('#sub-fiches').textContent = nf.format(all.length) + ' fiche' + (all.length > 1 ? 's' : '') + ' de suivi datée' + (all.length > 1 ? 's' : '') + ' dans la période avec au moins une anomalie, les plus récentes d’abord. Retrouvez le contact dans Tracker Capture puis corrigez la fiche du jour indiqué.';
    var host = $('#list-fiches');
    host.replaceChildren.apply(host, list.slice(0, state.nF).map(function (x) {
      var j = x.c && x.date != null ? ' · J' + (x.date - (x.c.deb || 0)) : '';
      return item(x.c ? idContact(x.c) : '(contact introuvable) ' + x.enr, x.nbAnom + ' anomalie' + (x.nbAnom > 1 ? 's' : ''),
        'Fiche du ' + fmtDate(x.date) + j + (x.vu ? ' · vu : ' + x.vu : '') + ' · ' + (x.zone || '–') + ' › ' + (x.aire || '–'), x.detail);
    }));
    if (!list.length) { var p = document.createElement('p'); p.className = 'sub'; p.textContent = q ? 'Aucune fiche ne correspond.' : 'Aucune fiche à corriger sur cette période.'; host.appendChild(p); }
    $('#more-fiches').hidden = list.length <= state.nF;
    $('#more-fiches').textContent = 'Afficher plus (' + nf.format(list.length - state.nF) + ' restantes)';
  }

  /* ---------- Onglets et feuilles ---------- */
  var ONGLETS = ['synthese', 'graphiques', 'zones', 'contacts', 'fiches'];
  function showTab(name) {
    if (ONGLETS.indexOf(name) < 0) name = 'synthese';
    state.tab = name; lsSet('mvqc-onglet', name);
    ONGLETS.forEach(function (t) { $('#tab-' + t).hidden = t !== name; });
    $$('#tabs button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.tab === name ? 'true' : 'false'); });
    window.scrollTo(0, 0);
    if (state.result) { renderGraphiques(); renderZones(); renderContacts(); renderFiches(); }
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
    var d = state.data, last = lsGet('mvqc-sync-last');
    $('#menu-info').textContent = (d ? 'Données : ' + d.fichier + ', reçues le ' + new Date(d.chargeLe).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + '. ' : 'Aucune donnée chargée. ') +
      (sync.key ? 'Mise à jour automatique activée' + (last ? ', dernière vérification réussie le ' + new Date(last).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) : '') + '.' : 'Mise à jour automatique désactivée.') +
      (function () { var e = lsGet('mvqc-sync-err'); return e && e.m ? ' Dernier échec (' + new Date(e.t).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + ') : ' + e.m : ''; })() +
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
  $('#btn-reset').addEventListener('click', function () { lsSet('mvqc-filtres', {}); closeSheets(); state.filters = resolveFilters(); render(); toast('Filtres réinitialisés.'); });
  $('#q-contacts').addEventListener('input', function () { state.nC = 50; renderContacts(); });
  $('#more-contacts').addEventListener('click', function () { state.nC += 50; renderContacts(); });
  $('#q-fiches').addEventListener('input', function () { state.nF = 50; renderFiches(); });
  $('#more-fiches').addEventListener('click', function () { state.nF += 50; renderFiches(); });
  var rz;
  window.addEventListener('resize', function () { clearTimeout(rz); rz = setTimeout(function () { if (state.result) renderGraphiques(); }, 150); });
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible' && Date.now() - (lsGet('mvqc-sync-last') || 0) > 2 * 60000) syncNow(false); });
  window.addEventListener('online', function () { syncNow(false); });
  var installEvt = null;
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); installEvt = e; $('#btn-install').hidden = false; });
  $('#btn-install').addEventListener('click', function () { if (!installEvt) return; closeSheets(); installEvt.prompt(); installEvt.userChoice.finally(function () { installEvt = null; $('#btn-install').hidden = true; }); });

  /* ---------- Démarrage ---------- */
  var cle = store.get('cle').then(function (c) { if (c && c.key) { sync.key = c.key; sync.salt = c.salt; sync.iter = c.iter; } }).catch(function () { /* aucune clé */ });
  cle.then(function () {
    setSyncUI();
    return store.get('dataset').then(function (d) {
      if (d && d.contacts) setData(setFromTables(d.contacts, d.suivis || [], { fichier: d.fichier, publie: d.publie, chargeLe: d.chargeLe }, d.seuils), false);
    }).catch(function () { /* première ouverture */ });
  }).then(function () {
    if (sync.key) return syncNow(false);
    return cleSuivi().then(function (ok) { if (ok) { setSyncUI(); toast('Mise à jour automatique activée avec le mot de passe de « Suivi MVE ».'); } });
  });
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(function () { /* hors ligne indisponible */ });
})();
