/* Interface de l'application « Vaccination MVE ». Calculs : engine-vacc.js */
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };
  var state = { data: null, filters: null, tab: 'synthese', result: null, nCorr: 50, nDoub: 30 };
  var nf = new Intl.NumberFormat('fr-FR');
  var pf = new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 1 });
  var df1 = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 });

  /* ---------- Dates ---------- */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function sToDate(s) { return MVEV.serialToUTC(s); }
  function fmtDate(s, short) { if (s == null) return '–'; var d = sToDate(s); return pad(d.getUTCDate()) + '/' + pad(d.getUTCMonth() + 1) + (short ? '' : '/' + d.getUTCFullYear()); }
  function sToInput(s) { var d = sToDate(s); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function inputToS(v) { return v ? MVEV.toSerial(v) : null; }
  function today() { return Math.floor(MVEV.nowSerial()); }

  /* ---------- Stockage local ---------- */
  function idb(name) {
    return new Promise(function (res, rej) {
      var rq = indexedDB.open(name || 'vaccination-mve', 1);
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
  // Paramètres!D2:D8 (début de campagne, âges, délai, seuils, écarts de dates) et F2:F31 (lots valides)
  function lireSeuils(wb) {
    var p = wb.Sheets['Paramètres']; if (!p) return null;
    var s = {}, v = function (a) { var c = p[a]; return c ? c.v : null; };
    ['DebutCampagne', 'AgeMin', 'AgeMax', 'DelaiMax', 'Vert', 'Orange'].forEach(function (k, i) { var x = v('D' + (i + 2)); if (typeof x === 'number') s[k] = x; });
    var inc = v('D8'); if (inc != null) s.InclDates = MVEV.norm(inc) === 'oui';
    var lots = []; for (var i = 2; i <= 31; i++) { var l = v('F' + i); if (l != null && String(l).trim()) lots.push(String(l).trim()); }
    if (lots.length) s.Lots = lots;
    return s;
  }
  function setFromTable(table, seuils, meta) {
    var out = MVEV.normalize(table, seuils);
    return Object.assign({ rows: out.rows, extrait: out.extrait, charge: out.charge, seuils: seuils || null, chargeLe: Date.now() }, meta);
  }
  /* Les noms ne sont jamais gardés : chaque clé « nom|sexe|âge » est remplacée par une
     empreinte HMAC-SHA256 (secret aléatoire, oublié aussitôt), puis la colonne « Nom » disparaît. */
  function anonymiser(table) {
    var h = (table[0] || []).map(MVEV.norm), iNom = h.indexOf(MVEV.norm('Nom du vacciné'));
    if (iNom < 0) return Promise.resolve(table);
    var iS = h.indexOf('sexe'), iA = h.indexOf(MVEV.norm('Âge (ans)'));
    var num = function (v) { return v == null || v === '' || isNaN(+v) ? null : +v; };
    var enc = new TextEncoder(), hex = function (b) { return Array.prototype.map.call(new Uint8Array(b).slice(0, 8), function (x) { return (x < 16 ? '0' : '') + x.toString(16); }).join(''); };
    return crypto.subtle.importKey('raw', crypto.getRandomValues(new Uint8Array(32)), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']).then(function (k) {
      var memo = {};
      return Promise.all(table.map(function (r, i) {
        var c = r.slice();
        if (i === 0) { c[iNom] = 'Clé doublon'; return c; }
        var cle = MVEV.cleNom(r[iNom], iS >= 0 ? r[iS] : '', iA >= 0 ? num(r[iA]) : null);
        if (!cle) { c[iNom] = null; return c; }
        memo[cle] = memo[cle] || crypto.subtle.sign('HMAC', k, enc.encode(cle)).then(hex);
        return memo[cle].then(function (x) { c[iNom] = x; return c; });
      }));
    });
  }
  function readWorkbook(buf, name, garder) {
    busy(true);
    return new Promise(function (r) { setTimeout(r, 30); }).then(function () {
      var wb = XLSX.read(buf, { type: 'array', sheets: ['Données', 'Paramètres'], cellDates: false, cellFormula: false, cellHTML: false, dense: false });
      if (!wb.Sheets['Données']) throw new Error('Ce fichier ne contient pas la feuille « Données ».');
      var table = XLSX.utils.sheet_to_json(wb.Sheets['Données'], { header: 1, raw: true, defval: null, blankrows: false });
      return anonymiser(table).then(function (t) { return { table: t, seuils: lireSeuils(wb) }; });
    }).then(function (o) {
      var table = o.table;
      var data = setFromTable(table, o.seuils, { fichier: name || 'fichier.xlsx' });
      return store.set('dataset', { table: table, seuils: data.seuils, fichier: data.fichier, publie: null, chargeLe: data.chargeLe })
        .catch(function () { /* stockage plein */ }).then(function () { return data; });
    }).then(function (d) { busy(false); setData(d, !garder); toast(nf.format(d.charge) + ' fiches chargées.'); })
      .catch(function (e) { busy(false); toast('Lecture impossible : ' + (e && e.message ? e.message : e), 6000); });
  }
  // Programme d'ordinateur (dossier de classeurs) : lecture directe d'un classeur sans passer par le sélecteur de fichier
  window.MVE_lireClasseur = function (buf, name, garder) { return readWorkbook(buf, name, garder); };
  function openFile(f) { if (f) f.arrayBuffer().then(function (b) { return readWorkbook(new Uint8Array(b), f.name); }); }

  /* ---------- Mise à jour automatique (vaccination.enc, même mot de passe que Suivi MVE) ---------- */
  var REMOTE = 'vaccination.enc', sync = { key: null, salt: null, iter: 0, running: false };
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
      if (r.status === 404) throw err('Aucune donnée de vaccination n’a encore été publiée.', 'absent');
      if (!r.ok) throw err('Le site de l’application ne répond pas (erreur ' + r.status + ').');
      return r.arrayBuffer();
    }, function () { throw err('Connexion impossible : les dernières données restent affichées.', 'reseau'); });
  }
  function applyRemote(p, force) {
    var cur = state.data;
    // On compare la date de publication (horodatage ISO de l'ordinateur), pas la date d'extraction :
    // une donnée locale (classeur ouvert à la main, horloge décalée) ne bloque plus les publications suivantes.
    if (!force && cur && cur.publie && p.publie && p.publie <= cur.publie) return Promise.resolve(false);
    var d = setFromTable(p.table, p.seuils, { fichier: 'mise à jour automatique', publie: p.publie });
    return store.set('dataset', { table: p.table, seuils: p.seuils, fichier: d.fichier, publie: p.publie, chargeLe: d.chargeLe })
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
      lsSet('mvev-sync-last', Date.now()); lsSet('mvev-sync-err', null);
      if (ch && !manual) toast('Données mises à jour automatiquement.'); else if (manual) toast(ch ? 'Données mises à jour.' : 'Les données sont déjà à jour.');
    }).catch(function (e) {
      var m = (e && e.message) || String(e);
      lsSet('mvev-sync-err', { m: m, t: Date.now() });
      if (e.code === 'cle') { sync.key = null; store.del('cle'); toast(m + ' Saisissez-le à nouveau.', 6000); openPassword(); }
      // Auparavant, les échecs automatiques étaient silencieux dès que des données étaient affichées.
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
            .then(function () { lsSet('mvev-sync-last', Date.now()); setSyncUI(); toast('Mise à jour automatique activée.'); });
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
  function minN(rows) { var m = Infinity; rows.forEach(function (r) { if (r.d != null && r.d < m) m = r.d; }); return m === Infinity ? today() : m; }
  function maxN(rows) { var m = 0; rows.forEach(function (r) { if (r.d != null && r.d > m) m = r.d; }); return m; }
  function seuils() { return Object.assign({}, MVEV.PARAMS_DEFAUT, (state.data && state.data.seuils) || {}); }
  function resolveFilters() {
    var s = lsGet('mvev-filtres') || {}, rows = state.data.rows, t = today();
    var f = { zone: s.zone || '', aire: s.aire || '', periode: s.periode || 'tout', debutFixe: s.debutFixe, finFixe: s.finFixe, refMode: s.refMode || '0', refFixe: s.refFixe };
    if (f.zone && MVEV.zonesOf(rows).indexOf(f.zone) < 0) { f.zone = ''; f.aire = ''; }
    f.ref = f.refMode === 'fixe' ? f.refFixe : t - (+f.refMode || 0);
    if (f.periode === 'fixe') { f.debut = f.debutFixe; f.fin = f.finFixe; }
    else if (f.periode === 'tout') { f.debut = minN(rows); f.fin = Math.max(t, maxN(rows)); }
    else { f.fin = f.ref; f.debut = f.ref - (+f.periode) + 1; }
    return f;
  }
  function saveFilters(f) { lsSet('mvev-filtres', { zone: f.zone, aire: f.aire, periode: f.periode, debutFixe: f.debutFixe, finFixe: f.finFixe, refMode: f.refMode, refFixe: f.refFixe }); }
  function fillZone(sel, v) { sel.replaceChildren(new Option('Toutes les zones', '')); MVEV.zonesOf(state.data.rows).forEach(function (z) { sel.appendChild(new Option(z, z)); }); sel.value = v || ''; }
  function fillAire(sel, z, v) {
    sel.replaceChildren(new Option(z ? 'Toutes les aires de ' + z : 'Toutes les aires', ''));
    MVEV.airesOf(state.data.rows, z).forEach(function (a) { sel.appendChild(new Option(a, a)); }); sel.value = v || ''; if (sel.value !== (v || '')) sel.value = '';
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
    saveFilters(f); closeSheets(); state.filters = resolveFilters(); state.nCorr = 50; state.nDoub = 30; render();
  }

  /* ---------- Rendu ---------- */
  function setData(d, fresh) {
    state.data = d; state.filters = resolveFilters();
    $('#empty').hidden = true; $('#tabs').hidden = false; $('#chips').hidden = false;
    showTab(fresh ? 'synthese' : (lsGet('mvev-onglet') || 'synthese'));
    render();
  }
  function cls(v) { var s = seuils(); return v == null ? '' : v >= s.Vert ? 'good' : v >= s.Orange ? 'mid' : 'bad'; }
  function render() {
    var d = state.data, f = state.filters;
    state.result = MVEV.compute(d.rows, f, seuils());
    var ex = d.extrait != null ? sToDate(d.extrait) : null, nv = d.rows.filter(function (r) { return r.vac; }).length;
    $('#src').textContent = 'DHIS2 SNIS · extrait le ' + (ex ? pad(ex.getUTCDate()) + '/' + pad(ex.getUTCMonth() + 1) + '/' + ex.getUTCFullYear() + ' à ' + pad(ex.getUTCHours()) + 'h' + pad(ex.getUTCMinutes()) : '(date inconnue)') +
      ' · ' + nf.format(nv) + ' vaccinations, ' + nf.format(d.charge) + ' inscrits';
    var age = d.extrait != null ? (Date.now() - (sToDate(d.extrait).getTime() + new Date().getTimezoneOffset() * 60000)) / 864e5 : null;
    $('#src').classList.toggle('stale', age != null && age > 2);
    renderChips(); renderSynthese(); renderQualite(); renderGraphiques(); renderZones(); renderCorriger(); renderDoublons();
  }
  function renderChips() {
    var f = state.filters, c = $('#chips');
    var lieu = f.aire ? f.aire + ' (' + (f.zone || 'aire') + ')' : (f.zone || 'Toutes les zones');
    var per = f.periode === 'tout' ? 'Toute la campagne' : (/^\d+$/.test(f.periode) ? f.periode + ' derniers jours' : fmtDate(f.debut) + ' → ' + fmtDate(f.fin));
    c.replaceChildren.apply(c, [['Lieu', lieu], ['Vaccinés', per], ['Réf.', fmtDate(f.ref)]].map(function (it) {
      var b = document.createElement('button'); b.className = 'chip'; var bb = document.createElement('b'); bb.textContent = it[1];
      b.append(it[0] + ' : ', bb); b.addEventListener('click', openFilters); return b;
    }));
  }

  /* Indicateurs (dictionnaire du classeur) : [clé, libellé, format, définition, lecture, remarque, coloration] */
  var P_ = 'Zone + aire + période (date de vaccination du/au)', R_ = 'Zone + aire + date de référence (ignore la période)';
  function tiles() {
    var s = seuils(), V = 'vert ≥ ' + pf.format(s.Vert) + ', orange ≥ ' + pf.format(s.Orange);
    return {
      vacc: [
        ['vaccines', 'Vaccinés (période)', 'n', 'Personnes ayant un événement « Vaccination », date de vaccination dans la période.', 'Volume de la campagne sur la période ; dénominateur des pourcentages.', P_, 'hero'],
        ['hommes', 'Hommes', 'n', 'Vaccinés (période) de sexe masculin.', 'Équilibre de couverture entre sexes.', P_],
        ['femmes', 'Femmes', 'n', 'Vaccinés (période) de sexe féminin.', 'Les fiches sans sexe ne sont comptées ni chez les hommes ni chez les femmes.', P_],
        ['enceintes', 'Femmes enceintes', 'n', 'Vaccinées déclarées enceintes.', 'Groupe nécessitant un suivi renforcé (issue de grossesse).', P_],
        ['allaitantes', 'Femmes allaitantes', 'n', 'Vaccinées déclarées allaitantes.', 'Suivi renforcé du nourrisson.', P_],
        ['pContact', '% contacts de cas', 'p', 'Contact d’un cas = Oui ÷ vaccinés (période).', 'Ciblage de la vaccination en anneau. Une part faible peut signaler une vaccination hors anneau ou un champ mal rempli.', P_],
        ['jour', 'Vaccinés – jour de réf.', 'n', 'Vaccinations dont la date est la date de référence.', 'Activité du jour (souvent incomplète le jour même : préférer « Hier »).', R_],
        ['j7', 'Vaccinés (7 derniers jours)', 'n', 'Vaccinations entre la date de référence − 6 jours et la date de référence.', 'Tendance récente de l’activité.', R_],
        ['moy7', 'Moyenne / jour (7 j)', 'd1', 'Vaccinés des 7 derniers jours ÷ 7.', 'Rythme journalier récent.', R_],
        ['cumul', 'Cumul depuis le début', 'n', 'Toutes les vaccinations jusqu’à la date de référence.', 'Progression globale de la campagne.', R_],
        ['ageMoyen', 'Âge moyen (ans)', 'd1', 'Âge moyen des vaccinés de la période.', 'Profil des personnes vaccinées.', P_],
        ['joursDepuis', 'Jours depuis la dernière vaccination', 'j', 'Date de référence − dernière date de vaccination.', 'Une valeur élevée signale une interruption de la vaccination ou de la saisie.', R_]
      ],
      secu: [
        ['mapi', 'MAPI notifiées', 'n', 'Manifestations post-vaccinales indésirables déclarées (MAPI = Oui).', 'Toute MAPI doit être suivie.', P_],
        ['mapiGrave', 'MAPI graves', 'n', 'MAPI déclarées graves.', 'Toute MAPI grave doit être investiguée et notifiée à la pharmacovigilance.', P_, 'zero'],
        ['pObs', '% observés 15 min', 'p', 'Observation de 15 minutes = Oui ÷ vaccinés.', 'Respect du protocole de surveillance immédiate. Vide ou Non = non documenté.', P_, true],
        ['pSuivi', '% avec ≥ 1 suivi post-vaccinal', 'p', 'Vaccinés ayant au moins un événement « Suivi après vaccination » ÷ vaccinés.', 'Une valeur basse peut refléter un suivi non saisi plutôt que non réalisé.', P_, true],
        ['pConsent', '% consentement « Oui »', 'p', 'Consentement = Oui ÷ vaccinés.', 'Exigence éthique : toute vaccination devrait avoir un consentement documenté.', P_, true],
        ['sansVacc', 'Inscrits sans vaccination', 'n', 'Personnes inscrites au registre sans événement « Vaccination ».', 'Inscription incomplète, refus ou fiche créée par erreur : à vérifier dans Tracker Capture.', 'Zone + aire (sans date)', 'zero']
      ],
      qual: [
        ['global', 'Score global de qualité', 'p', 'Moyenne des scores des 6 dimensions.', 'Synthèse de la qualité des données (' + V + ').', P_, true, 'hero'],
        ['dim_Complétude', 'Complétude', 'p', 'Moyenne des 11 indicateurs de complétude.', 'Les variables clés du formulaire sont-elles remplies ?', P_, true],
        ['dim_Validité', 'Validité', 'p', 'Moyenne des 4 indicateurs de validité.', 'Format et valeurs attendus (ID, lot, âge, téléphone).', P_, true],
        ['dim_Cohérence', 'Cohérence', 'p', 'Moyenne des 5 indicateurs de cohérence.', 'Les dates et variables liées sont-elles logiques ?', P_, true],
        ['dim_Conformité protocole', 'Conformité au protocole', 'p', 'Moyenne des 3 indicateurs de protocole.', 'Consentement, observation 15 min, suivi post-vaccinal documentés.', P_, true],
        ['dim_Unicité', 'Unicité', 'p', 'Moyenne des 2 indicateurs d’unicité.', 'Chaque personne vaccinée est-elle enregistrée une seule fois ?', P_, true],
        ['dim_Promptitude', 'Promptitude de saisie', 'p', 'Moyenne des 2 indicateurs de promptitude.', 'Les vaccinations sont-elles saisies rapidement dans DHIS2 ?', P_, true],
        ['avecAnom', 'Fiches avec ≥ 1 anomalie', 'n', 'Fiches présentant au moins une anomalie corrigeable (contrôles A1 à A17).', 'Charge de travail de correction : voir l’onglet « À corriger ».', 'Zone + aire + période (+ inscrits sans date)'],
        ['pSansAnom', '% fiches vaccinées sans anomalie', 'p', '1 − fiches vaccinées avec anomalie ÷ vaccinés.', 'Indicateur simple de qualité par fiche.', P_, true],
        ['doublons', 'Doublons probables', 'n', 'Fiches partageant le même nom (normalisé), sexe et âge qu’une autre fiche.', 'Risque de double comptage. Comparer les fiches avant toute suppression (homonymes possibles).', P_, 'zero'],
        ['delaiMoyen', 'Délai moyen de saisie (j)', 'd1', 'Moyenne de (date de saisie − date de vaccination), délais ≥ 0.', 'Rapidité de remontée des données.', P_],
        ['ecartDates', 'Date vacc. ≠ date événement', 'n', 'Fiches dont la date de vaccination du formulaire diffère de la date de l’événement DHIS2.', 'Souvent, la date de l’événement est la date de saisie : les analyses DHIS2 basées sur la date d’événement sont alors décalées.', P_]
      ]
    };
  }
  function fmtVal(v, t) {
    if (v == null) return '–';
    if (t === 'p') return pf.format(v); if (t === 'd1') return df1.format(v); if (t === 'j') return nf.format(v) + ' j';
    if (t === 'date') return fmtDate(v); return nf.format(v);
  }
  function renderSynthese() {
    var k = state.result.k, f = state.filters, T = tiles();
    Object.keys(T).forEach(function (g) {
      var host = $('#tiles-' + g);
      host.replaceChildren.apply(host, T[g].map(function (t) {
        var b = document.createElement('button'), v = k[t[0]];
        b.className = 'tile' + (t[6] === 'hero' || t[7] === 'hero' ? ' hero' : '');
        if (t[6] === true && v != null && t[7] !== 'hero') b.className += ' q-' + cls(v);
        else if (t[6] === true && t[7] === 'hero') b.className += ' h-' + (cls(v) || 'x');
        else if (t[6] === 'zero' && v != null) b.className += v === 0 ? ' q-good' : ' q-bad';
        else if (t[0] === 'joursDepuis' && v != null) b.className += v <= 1 ? ' q-good' : v <= 3 ? ' q-mid' : ' q-bad';
        var lab = document.createElement('span'); lab.className = 'lab'; lab.textContent = t[1];
        var val = document.createElement('span'); val.className = 'val'; val.textContent = fmtVal(v, t[2]);
        b.append(lab, val);
        if (t[0] === 'joursDepuis' && k.dernier) { var n = document.createElement('span'); n.className = 'note'; n.textContent = 'dernière le ' + fmtDate(k.dernier); b.appendChild(n); }
        b.addEventListener('click', function () {
          $('#t-def').textContent = t[1]; $('#def-val').textContent = fmtVal(v, t[2]);
          $('#def-def').textContent = t[3]; $('#def-int').textContent = t[4]; $('#def-act').textContent = t[5]; openSheet('#sheet-def');
        });
        return b;
      }));
    });
    $('#sub-vacc').textContent = 'Vaccinés du ' + fmtDate(f.debut) + ' au ' + fmtDate(f.fin) + ' ; « jour », « 7 j » et « cumul » à la date de référence (' + fmtDate(f.ref) + ').';
    var lieu = f.aire ? 'l’aire ' + de(f.aire) : f.zone ? 'la zone ' + de(f.zone) : 'le Haut-Uélé';
    var parts = ['Dans ', lieu, ', ', ['b', nf.format(k.vaccines)], ' personne' + (k.vaccines > 1 ? 's' : '') + ' vaccinée' + (k.vaccines > 1 ? 's' : '') + ' sur la période, dont ', ['b', nf.format(k.j7)], ' ces 7 derniers jours (cumul : ', ['b', nf.format(k.cumul)], ').'];
    if (k.mapi) parts.push(' ', ['b', nf.format(k.mapi)], ' MAPI notifiée' + (k.mapi > 1 ? 's' : '') + (k.mapiGrave ? ', dont ' + k.mapiGrave + ' grave' + (k.mapiGrave > 1 ? 's' : '') : '') + '.');
    if (k.global != null) parts.push(' Score de qualité : ', ['b', pf.format(k.global)], ' ; ', ['b', nf.format(k.avecAnom)], ' fiche' + (k.avecAnom > 1 ? 's' : '') + ' à corriger.');
    var lede = $('#lede');
    lede.replaceChildren.apply(lede, parts.map(function (p) { if (typeof p === 'string') return document.createTextNode(p); var e = document.createElement(p[0]); e.textContent = p[1]; return e; }));
  }
  function de(n) { return /^[aeiouyéèêâîôûh]/i.test(n) ? 'd’' + n : 'de ' + n; }

  function td(t, c) { var e = document.createElement('td'); e.textContent = t; if (c) e.className = c; return e; }
  function pctTd(v) { var e = document.createElement('td'); e.className = 'n'; if (v == null) { e.textContent = '–'; return e; } var s = document.createElement('span'); s.className = 'pct ' + cls(v); s.textContent = pf.format(v); e.appendChild(s); return e; }
  function renderQualite() {
    if (state.tab !== 'qualite') return;
    var r = state.result, s = seuils();
    var tb = $('#tb-dims');
    tb.replaceChildren.apply(tb, r.dims.map(function (d) {
      var tr = document.createElement('tr'); tr.appendChild(td(d.dim, 'first')); tr.appendChild(pctTd(d.score)); tr.appendChild(td(r.statut(d.score))); tr.appendChild(td(d.quoi, 'wrap')); return tr;
    }));
    var ti = $('#tb-ind'), dim = null, out = [];
    r.indicateurs.forEach(function (i) {
      if (i.dim !== dim) { dim = i.dim; var h = document.createElement('tr'); h.className = 'grp'; var c = td(dim, 'first'); c.colSpan = 5; h.appendChild(c); out.push(h); }
      var tr = document.createElement('tr');
      tr.appendChild(td(i.lib, 'first wrap')); tr.appendChild(td(nf.format(i.num), 'n')); tr.appendChild(td(nf.format(i.den), 'n')); tr.appendChild(pctTd(i.pct)); tr.appendChild(td(i.pct != null && i.pct < s.Vert ? i.action : '', 'wrap act'));
      out.push(tr);
    });
    ti.replaceChildren.apply(ti, out);
    $('#note-qual').textContent = 'Chaque indicateur = % de fiches conformes. Vert ≥ ' + pf.format(s.Vert) + ', orange ≥ ' + pf.format(s.Orange) + ', rouge en dessous. L’action recommandée s’affiche quand l’indicateur n’est pas « Bon ».';
  }

  /* ---------- Graphiques SVG ---------- */
  var NS = 'http://www.w3.org/2000/svg';
  function el(tag, a, p) { var e = document.createElementNS(NS, tag); for (var k in a) e.setAttribute(k, a[k]); if (p) p.appendChild(e); return e; }
  function tw(svg, t) { var x = el('text', { x: -999, y: -999 }, svg); x.textContent = t; var w = x.getComputedTextLength ? x.getComputedTextLength() : t.length * 6.5; svg.removeChild(x); return w || t.length * 6.5; }
  function nice(v) { if (v <= 5) return 5; var p = Math.pow(10, Math.floor(Math.log10(v))), m = v / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p; }
  function hbars(host, items, o) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), rowH = 26, H = items.length * rowH + 4;
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': o.aria }, host);
    var lw = Math.min(W * 0.5, Math.max.apply(null, items.map(function (i) { return tw(svg, i.label); })) + 10);
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
  function quotidien(host, data) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), H = 220, svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Vaccinations quotidiennes et cumul' }, host);
    var maxN = nice(Math.max(1, Math.max.apply(null, data.map(function (d) { return d.n; })))), maxC = nice(Math.max(1, data[data.length - 1].cumul));
    var ml = tw(svg, nf.format(maxN)) + 8, mr = tw(svg, nf.format(maxC)) + 8, mt = 8, mb = 24, iw = W - ml - mr, ih = H - mt - mb, bw = iw / data.length;
    var yN = function (v) { return mt + ih - v / maxN * ih; }, yC = function (v) { return mt + ih - v / maxC * ih; };
    [0, .5, 1].forEach(function (p) {
      el('line', { x1: ml, x2: W - mr, y1: mt + ih - p * ih, y2: mt + ih - p * ih, style: 'stroke:var(--line)' }, svg);
      var a = el('text', { x: ml - 6, y: mt + ih - p * ih, 'text-anchor': 'end', dy: '0.32em' }, svg); a.textContent = nf.format(maxN * p);
      var b = el('text', { x: W - mr + 6, y: mt + ih - p * ih, dy: '0.32em' }, svg); b.textContent = nf.format(maxC * p);
    });
    data.forEach(function (d, i) {
      var x = ml + i * bw;
      if (d.n) el('rect', { x: x + bw * .15, width: bw * .7, y: yN(d.n), height: ih - (yN(d.n) - mt), rx: 2, style: 'fill:var(--ink-2)' }, svg);
      if (i % 5 === 4 || i === data.length - 1) { var t = el('text', { x: x + bw / 2, y: H - 6, 'text-anchor': 'middle' }, svg); t.textContent = fmtDate(d.date, true); }
    });
    el('polyline', { points: data.map(function (d, i) { return (ml + i * bw + bw / 2) + ',' + yC(d.cumul); }).join(' '), fill: 'none', 'stroke-width': 2, style: 'stroke:var(--accent)' }, svg);
    tipOn(host, svg, W, ml, bw, data.length, function (i) { var d = data[i]; return [fmtDate(d.date), 'Vaccinés : ' + nf.format(d.n), 'Cumul : ' + nf.format(d.cumul)]; });
  }
  function pyramide(host, data) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), rowH = 24, H = data.length * rowH + 22, svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Vaccinés par tranche d’âge et sexe' }, host);
    var max = Math.max(1, Math.max.apply(null, data.map(function (d) { return Math.max(d.h, d.f); })));
    var lw = Math.max.apply(null, data.map(function (d) { return tw(svg, d.tranche); })) + 12, vw = tw(svg, nf.format(max)) + 8;
    var half = (W - lw) / 2 - vw, cx = lw / 2 + W / 2;
    data.forEach(function (d, i) {
      var y = i * rowH + 2, wh = d.h / max * half, wf = d.f / max * half;
      var t = el('text', { x: W / 2, y: y + rowH / 2, 'text-anchor': 'middle', dy: '0.32em' }, svg); t.textContent = d.tranche;
      if (wh) el('rect', { x: W / 2 - lw / 2 - wh, y: y + 4, width: wh, height: rowH - 8, rx: 3, style: 'fill:var(--men)' }, svg);
      if (wf) el('rect', { x: W / 2 + lw / 2, y: y + 4, width: wf, height: rowH - 8, rx: 3, style: 'fill:var(--women)' }, svg);
      var a = el('text', { x: W / 2 - lw / 2 - wh - 4, y: y + rowH / 2, 'text-anchor': 'end', dy: '0.32em', class: 'vlab' }, svg); a.textContent = d.h ? nf.format(d.h) : '';
      var b = el('text', { x: W / 2 + lw / 2 + wf + 4, y: y + rowH / 2, dy: '0.32em', class: 'vlab' }, svg); b.textContent = d.f ? nf.format(d.f) : '';
    });
    void cx;
  }
  function renderGraphiques() {
    if (state.tab !== 'graphiques') return;
    var r = state.result, s = seuils();
    quotidien($('#ch-jours'), r.jours);
    pyramide($('#ch-age'), r.ageSexe);
    hbars($('#ch-delai'), r.delais.map(function (d) { return { label: d.classe, value: d.n, color: d.classe === 'Négatif' || d.classe === '8 j et +' ? 'var(--bad)' : d.classe === '0 j' || d.classe === '1 j' ? 'var(--ok)' : 'var(--warn)' }; }),
      { fmt: function (v) { return nf.format(v); }, aria: 'Délai de saisie' });
    hbars($('#ch-compl'), r.indicateurs.filter(function (i) { return i.dim === 'Complétude'; }).map(function (i) {
      return { label: i.lib, value: i.pct, color: i.pct == null ? 'var(--line)' : i.pct >= s.Vert ? 'var(--ok)' : i.pct >= s.Orange ? 'var(--warn)' : 'var(--bad)' };
    }), { max: 1, fmt: function (v) { return pf.format(v); }, aria: 'Complétude par variable' });
    hbars($('#ch-anom'), r.anomTypes.slice().sort(function (a, b) { return b.n - a.n; }).map(function (t) { return { label: t.type, value: t.n }; }),
      { color: 'var(--bad)', fmt: function (v) { return nf.format(v); }, aria: 'Anomalies par type' });
  }

  function lieuRows(host, list, tot, clic) {
    var row = function (z, total) {
      var tr = document.createElement('tr'); if (total) tr.className = 'total'; else if (z.j7) tr.className = 'actif';
      tr.appendChild(td(z.lieu, 'first'));
      [z.periode, z.j7, z.jour, z.cumul, z.mapi].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); });
      tr.appendChild(pctTd(z.pSans)); tr.appendChild(pctTd(z.pLot)); tr.appendChild(td(nf.format(z.doublons), 'n'));
      tr.appendChild(td(z.delai == null ? '–' : df1.format(z.delai), 'n')); tr.appendChild(td(z.dernier ? fmtDate(z.dernier) : '–', 'n'));
      if (!total && clic) { tr.style.cursor = 'pointer'; tr.addEventListener('click', function () { clic(z); }); }
      return tr;
    };
    host.replaceChildren.apply(host, list.map(function (z) { return row(z); }).concat([row(tot, true)]));
  }
  function renderZones() {
    if (state.tab !== 'zones') return;
    var r = state.result, f = state.filters;
    lieuRows($('#tb-zones'), r.zones, r.totalZones, function (z) { var s = lsGet('mvev-filtres') || {}; s.zone = z.lieu; s.aire = ''; lsSet('mvev-filtres', s); state.filters = resolveFilters(); render(); toast('Filtre : zone ' + de(z.lieu) + '.'); });
    $('#t-aires').textContent = f.zone ? 'Situation par aire de santé – zone ' + de(f.zone) : 'Situation par aire de santé – toutes les zones';
    lieuRows($('#tb-aires'), r.aires, r.totalAires, function (a) { var s = lsGet('mvev-filtres') || {}; s.aire = a.lieu; lsSet('mvev-filtres', s); state.filters = resolveFilters(); render(); toast('Filtre : aire ' + de(a.lieu) + '.'); });
    var s = seuils();
    $('#note-zones').textContent = 'Fond bleu : vaccination dans les 7 jours précédant la date de référence. % sans anomalie et % lot conforme : vert ≥ ' + pf.format(s.Vert) + ', orange ≥ ' + pf.format(s.Orange) + ', rouge en dessous. Touchez une ligne pour filtrer.';
  }

  function meta(r) { return (r.d ? 'Vacciné le ' + fmtDate(r.d) : 'Sans date de vaccination') + (r.dSai ? ' · saisi le ' + fmtDate(r.dSai) : '') + ' · ' + (r.zone || '–') + ' › ' + (r.aire || '–'); }
  function renderCorriger() {
    if (state.tab !== 'corriger') return;
    var q = MVEV.norm($('#q-corriger').value), all = state.result.aCorriger;
    var list = q ? all.filter(function (r) { return MVEV.norm((r.id || '') + ' ' + (r.aire || '') + ' ' + (r.zone || '') + ' ' + r.detail.join(' ')).indexOf(q) >= 0; }) : all;
    $('#sub-corriger').textContent = nf.format(all.length) + ' fiche' + (all.length > 1 ? 's' : '') + ' avec au moins une anomalie (lieu et période choisis, plus les inscrits sans date), les plus récentes d’abord. Recherchez la fiche dans Tracker Capture par son ID de vaccination.';
    var host = $('#list-corriger');
    host.replaceChildren.apply(host, list.slice(0, state.nCorr).map(function (r) {
      var d = document.createElement('div'); d.className = 'item';
      var top = document.createElement('div'); top.className = 'top-line';
      var id = document.createElement('span'); id.className = 'idn'; id.textContent = r.id || '(sans ID de vaccination)';
      var nb = document.createElement('span'); nb.className = 'nb'; nb.textContent = r.nbAnom + ' anomalie' + (r.nbAnom > 1 ? 's' : '');
      top.append(id, nb);
      var m = document.createElement('div'); m.className = 'meta'; m.textContent = meta(r);
      var tags = document.createElement('div'); tags.className = 'tags';
      r.detail.forEach(function (x) { var s = document.createElement('span'); if (/manquant|absent/.test(x)) s.className = ''; else s.className = 'inc'; s.textContent = x; tags.appendChild(s); });
      d.append(top, m, tags); return d;
    }));
    if (!list.length) { var p = document.createElement('p'); p.className = 'sub'; p.textContent = q ? 'Aucune fiche ne correspond à la recherche.' : 'Aucune fiche à corriger sur cette période.'; host.appendChild(p); }
    $('#more-corriger').hidden = list.length <= state.nCorr;
    $('#more-corriger').textContent = 'Afficher plus (' + nf.format(list.length - state.nCorr) + ' restantes)';
  }
  function renderDoublons() {
    if (state.tab !== 'doublons') return;
    var g = state.result.doublons, n = g.reduce(function (t, x) { return t + x.membres.length; }, 0);
    $('#sub-doublons').textContent = nf.format(n) + ' fiche' + (n > 1 ? 's' : '') + ' dans ' + nf.format(g.length) + ' groupe' + (g.length > 1 ? 's' : '') +
      ' : même nom, même sexe et même âge. Les noms ne sont pas affichés : retrouvez les fiches dans Tracker Capture par leur ID de vaccination et comparez-les avant toute suppression (homonymes possibles).';
    var host = $('#list-doublons');
    host.replaceChildren.apply(host, g.slice(0, state.nDoub).map(function (grp, i) {
      var box = document.createElement('div'); box.className = 'group';
      var h = document.createElement('h3'), a = document.createElement('span'), b = document.createElement('span'), m0 = grp.membres[0];
      a.textContent = 'Groupe ' + (i + 1); b.textContent = (m0.fem ? 'Féminin' : m0.hom ? 'Masculin' : 'sexe ?') + ' · ' + (m0.age != null ? m0.age + ' ans' : 'âge ?') + ' · ' + grp.total + ' fiches au total';
      h.append(a, b); box.appendChild(h);
      grp.membres.forEach(function (r) {
        var m = document.createElement('div'); m.className = 'mem';
        var id = document.createElement('b'); id.textContent = r.id || '(sans ID de vaccination)';
        var mt = document.createElement('span'); mt.style.color = 'var(--muted)'; mt.textContent = meta(r);
        m.append(id, mt); box.appendChild(m);
      });
      if (grp.total > grp.membres.length) { var o = document.createElement('div'); o.className = 'mem'; o.style.color = 'var(--muted)'; o.textContent = (grp.total - grp.membres.length) + ' autre(s) fiche(s) du groupe hors du filtre.'; box.appendChild(o); }
      return box;
    }));
    if (!g.length) { var p = document.createElement('p'); p.className = 'sub'; p.textContent = 'Aucun doublon probable sur cette période.'; host.appendChild(p); }
    $('#more-doublons').hidden = g.length <= state.nDoub;
  }

  /* ---------- Onglets et feuilles ---------- */
  function showTab(name) {
    state.tab = name; lsSet('mvev-onglet', name);
    ['synthese', 'qualite', 'graphiques', 'zones', 'corriger', 'doublons'].forEach(function (t) { $('#tab-' + t).hidden = t !== name; });
    $$('#tabs button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.tab === name ? 'true' : 'false'); });
    window.scrollTo(0, 0);
    if (state.result) { renderQualite(); renderGraphiques(); renderZones(); renderCorriger(); renderDoublons(); }
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
    var d = state.data, last = lsGet('mvev-sync-last');
    $('#menu-info').textContent = (d ? 'Données : ' + d.fichier + ', reçues le ' + new Date(d.chargeLe).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + '. ' : 'Aucune donnée chargée. ') +
      (sync.key ? 'Mise à jour automatique activée' + (last ? ', dernière vérification réussie le ' + new Date(last).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) : '') + '.' : 'Mise à jour automatique désactivée.') +
      (function () { var e = lsGet('mvev-sync-err'); return e && e.m ? ' Dernier échec (' + new Date(e.t).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + ') : ' + e.m : ''; })() +
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
  $('#btn-reset').addEventListener('click', function () { lsSet('mvev-filtres', {}); closeSheets(); state.filters = resolveFilters(); render(); toast('Filtres réinitialisés.'); });
  $('#q-corriger').addEventListener('input', function () { state.nCorr = 50; renderCorriger(); });
  $('#more-corriger').addEventListener('click', function () { state.nCorr += 50; renderCorriger(); });
  $('#more-doublons').addEventListener('click', function () { state.nDoub += 30; renderDoublons(); });
  var rz;
  window.addEventListener('resize', function () { clearTimeout(rz); rz = setTimeout(function () { if (state.result) renderGraphiques(); }, 150); });
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible' && Date.now() - (lsGet('mvev-sync-last') || 0) > 2 * 60000) syncNow(false); });
  window.addEventListener('online', function () { syncNow(false); });
  var installEvt = null;
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); installEvt = e; $('#btn-install').hidden = false; });
  $('#btn-install').addEventListener('click', function () { if (!installEvt) return; closeSheets(); installEvt.prompt(); installEvt.userChoice.finally(function () { installEvt = null; $('#btn-install').hidden = true; }); });

  /* ---------- Démarrage ---------- */
  var cle = store.get('cle').then(function (c) { if (c && c.key) { sync.key = c.key; sync.salt = c.salt; sync.iter = c.iter; } }).catch(function () { /* aucune clé */ });
  cle.then(function () {
    setSyncUI();
    return store.get('dataset').then(function (d) {
      if (d && d.table) setData(setFromTable(d.table, d.seuils, { fichier: d.fichier, publie: d.publie, chargeLe: d.chargeLe }), false);
    }).catch(function () { /* première ouverture */ });
  }).then(function () {
    if (sync.key) return syncNow(false);
    // Pas encore de clé ici : on réutilise celle de « Suivi MVE » si elle existe (même mot de passe d'équipe)
    return cleSuivi().then(function (ok) { if (ok) { setSyncUI(); toast('Mise à jour automatique activée avec le mot de passe de « Suivi MVE ».'); } });
  });
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(function () { /* hors ligne indisponible */ });
})();
