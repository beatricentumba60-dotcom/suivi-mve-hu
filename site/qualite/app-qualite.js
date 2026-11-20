/* Interface de l'application « Qualité des données MVE ». Calculs : engine-qualite.js */
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
  function sToDate(s) { return MVEQ.serialToUTC(s); }
  function fmtDate(s, short) { if (s == null) return '–'; var d = sToDate(s); return pad(d.getUTCDate()) + '/' + pad(d.getUTCMonth() + 1) + (short ? '' : '/' + d.getUTCFullYear()); }
  function sToInput(s) { var d = sToDate(s); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function inputToS(v) { return v ? MVEQ.toSerial(v) : null; }
  function today() { return Math.floor(MVEQ.nowSerial()); }

  /* ---------- Stockage local ---------- */
  function idb(name) {
    return new Promise(function (res, rej) {
      var rq = indexedDB.open(name || 'qualite-mve', 1);
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
  function lireSeuils(wb) {
    var p = wb.Sheets['Paramètres']; if (!p) return null;
    var cles = ['DDSmax', 'AgeMax', 'CiblePrelev', 'RetInv', 'RetLabo', 'RetClassif', 'Vert', 'Orange', 'Fenetre'], s = {};
    cles.forEach(function (k, i) { var c = p['D' + (i + 2)]; if (c && typeof c.v === 'number') s[k] = c.v; });
    return s;
  }
  function setFromTable(table, seuils, meta) {
    var out = MVEQ.normalize(table, seuils);
    return Object.assign({ rows: out.rows, extrait: out.extrait, sansDate: out.sansDate, charge: out.charge, seuils: seuils || null, chargeLe: Date.now() }, meta);
  }
  /* Les noms ne sont jamais gardés : chaque clé « nom|sexe|âge » est remplacée par une
     empreinte HMAC-SHA256 (secret aléatoire, oublié aussitôt), puis la colonne « Nom » disparaît. */
  function anonymiser(table) {
    var h = (table[0] || []).map(MVEQ.norm), iNom = h.indexOf('nom');
    if (iNom < 0) return Promise.resolve(table);
    var iS = h.indexOf('sexe'), iA = h.indexOf(MVEQ.norm('Âge (ans)')), iM = h.indexOf(MVEQ.norm('Âge (mois)'));
    var num = function (v) { return v == null || v === '' || isNaN(+v) ? null : +v; };
    var enc = new TextEncoder(), hex = function (b) { return Array.prototype.map.call(new Uint8Array(b).slice(0, 8), function (x) { return (x < 16 ? '0' : '') + x.toString(16); }).join(''); };
    return crypto.subtle.importKey('raw', crypto.getRandomValues(new Uint8Array(32)), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']).then(function (k) {
      var memo = {};
      return Promise.all(table.map(function (r, i) {
        var c = r.slice();
        if (i === 0) { c[iNom] = 'Clé doublon'; return c; }
        var cle = MVEQ.cleNom(r[iNom], iS >= 0 ? r[iS] : '', iA >= 0 ? num(r[iA]) : null, iM >= 0 ? num(r[iM]) : null);
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
    }).then(function (d) { busy(false); setData(d, !garder); toast(nf.format(d.charge) + ' alertes chargées.'); })
      .catch(function (e) { busy(false); toast('Lecture impossible : ' + (e && e.message ? e.message : e), 6000); });
  }
  // Programme d'ordinateur (dossier de classeurs) : lecture directe d'un classeur sans passer par le sélecteur de fichier
  window.MVE_lireClasseur = function (buf, name, garder) { return readWorkbook(buf, name, garder); };
  function openFile(f) { if (f) f.arrayBuffer().then(function (b) { return readWorkbook(new Uint8Array(b), f.name); }); }

  /* ---------- Mise à jour automatique (qualite.enc, même mot de passe que Suivi MVE) ---------- */
  var REMOTE = 'qualite.enc', sync = { key: null, salt: null, iter: 0, running: false };
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
      if (r.status === 404) throw err('Aucun contrôle de qualité n’a encore été publié depuis l’ordinateur.', 'absent');
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
      lsSet('mveq-sync-last', Date.now()); lsSet('mveq-sync-err', null);
      if (ch && !manual) toast('Données mises à jour automatiquement.'); else if (manual) toast(ch ? 'Données mises à jour.' : 'Les données sont déjà à jour.');
    }).catch(function (e) {
      var m = (e && e.message) || String(e);
      lsSet('mveq-sync-err', { m: m, t: Date.now() });
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
            .then(function () { lsSet('mveq-sync-last', Date.now()); setSyncUI(); toast('Mise à jour automatique activée.'); });
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
  function minN(rows) { var m = Infinity; rows.forEach(function (r) { if (r.n < m) m = r.n; }); return m; }
  function maxN(rows) { var m = 0; rows.forEach(function (r) { if (r.n > m) m = r.n; }); return m; }
  function seuils() { return Object.assign({}, MVEQ.SEUILS_DEFAUT, (state.data && state.data.seuils) || {}); }
  function resolveFilters() {
    var s = lsGet('mveq-filtres') || {}, rows = state.data.rows, t = today();
    var f = { zone: s.zone || '', aire: s.aire || '', periode: s.periode || String(seuils().Fenetre), debutFixe: s.debutFixe, finFixe: s.finFixe, refMode: s.refMode || '0', refFixe: s.refFixe };
    if (f.zone && MVEQ.zonesOf(rows).indexOf(f.zone) < 0) { f.zone = ''; f.aire = ''; }
    f.ref = f.refMode === 'fixe' ? f.refFixe : t - (+f.refMode || 0);
    if (f.periode === 'fixe') { f.debut = f.debutFixe; f.fin = f.finFixe; }
    else if (f.periode === 'tout') { f.debut = minN(rows); f.fin = Math.max(t, maxN(rows)); }
    else { f.fin = f.ref; f.debut = f.ref - (+f.periode) + 1; }
    return f;
  }
  function saveFilters(f) { lsSet('mveq-filtres', { zone: f.zone, aire: f.aire, periode: f.periode, debutFixe: f.debutFixe, finFixe: f.finFixe, refMode: f.refMode, refFixe: f.refFixe }); }
  function fillZone(sel, v) { sel.replaceChildren(new Option('Toutes les zones', '')); MVEQ.zonesOf(state.data.rows).forEach(function (z) { sel.appendChild(new Option(z, z)); }); sel.value = v || ''; }
  function fillAire(sel, z, v) {
    sel.replaceChildren(new Option(z ? 'Toutes les aires de ' + z : 'Toutes les aires', ''));
    MVEQ.airesOf(state.data.rows, z).forEach(function (a) { sel.appendChild(new Option(a, a)); }); sel.value = v || ''; if (sel.value !== (v || '')) sel.value = '';
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
    showTab(fresh ? 'synthese' : (lsGet('mveq-onglet') || 'synthese'));
    render();
  }
  function cls(v) { var s = seuils(); return v == null ? '' : v >= s.Vert ? 'good' : v >= s.Orange ? 'mid' : 'bad'; }
  function render() {
    var d = state.data, f = state.filters, s = seuils();
    state.result = MVEQ.compute(d.rows, f, s);
    var ex = d.extrait != null ? sToDate(d.extrait) : null;
    $('#src').textContent = 'DHIS2 SNIS · extrait le ' + (ex ? pad(ex.getUTCDate()) + '/' + pad(ex.getUTCMonth() + 1) + '/' + ex.getUTCFullYear() + ' à ' + pad(ex.getUTCHours()) + 'h' + pad(ex.getUTCMinutes()) : '(date inconnue)') +
      ' · ' + nf.format(d.charge) + ' alertes' + (d.sansDate ? ' (dont ' + nf.format(d.sansDate) + ' sans date de notification)' : '');
    var age = d.extrait != null ? (Date.now() - (sToDate(d.extrait).getTime() + new Date().getTimezoneOffset() * 60000)) / 864e5 : null;
    $('#src').classList.toggle('stale', age != null && age > 2);
    renderChips(); renderSynthese(); renderGraphiques(); renderZones(); renderCorriger(); renderDoublons();
  }
  function renderChips() {
    var f = state.filters, c = $('#chips');
    var lieu = f.aire ? f.aire + ' (' + (f.zone || 'aire') + ')' : (f.zone || 'Toutes les zones');
    var per = f.periode === 'tout' ? 'Toute la période' : (/^\d+$/.test(f.periode) ? f.periode + ' derniers jours' : fmtDate(f.debut) + ' → ' + fmtDate(f.fin));
    c.replaceChildren.apply(c, [['Lieu', lieu], ['Période', per], ['Réf.', fmtDate(f.ref)]].map(function (it) {
      var b = document.createElement('button'); b.className = 'chip'; var bb = document.createElement('b'); bb.textContent = it[1];
      b.append(it[0] + ' : ', bb); b.addEventListener('click', openFilters); return b;
    }));
  }

  /* Indicateurs : [clé, libellé, format, définition, lecture, action, coloration] */
  function tiles() {
    var s = seuils(), V = 'vert ≥ ' + pf.format(s.Vert) + ', orange ≥ ' + pf.format(s.Orange);
    return {
      vue: [
        ['alertes', 'Alertes analysées', 'n', 'Alertes notifiées dans la période (zone et aire choisies).', 'Volume d’activité. Avec peu d’alertes, les pourcentages varient beaucoup.', 'Dénominateur des pourcentages de la vue d’ensemble.'],
        ['pSans', 'Alertes sans anomalie', 'p', 'Part des alertes sans aucune variable clé manquante ni incohérence.', 'Score global de qualité (' + V + '). Rouge : données peu fiables, corrections prioritaires.', 'Plus c’est haut, mieux c’est.', true],
        ['avecAnom', 'Alertes avec ≥ 1 anomalie', 'n', 'Alertes ayant au moins une variable manquante ou une incohérence.', 'Charge de travail de correction dans DHIS2.', 'Voir l’onglet « À corriger ».'],
        ['completude', 'Complétude des variables clés', 'p', '1 − manquantes ÷ applicables, sur les 10 variables clés.', 'Mesure si les informations essentielles sont saisies (' + V + '). Une variable non applicable n’est pas comptée.', 'Voir les variables les moins remplies dans « Graphiques ».', true],
        ['avecIncoh', 'Alertes avec incohérence', 'n', 'Alertes avec au moins une incohérence logique (14 contrôles).', 'Erreurs de saisie (dates inversées, statuts contradictoires) qui faussent l’analyse.', 'Vérifier dates, âge, statuts et classification.'],
        ['doublons', 'Doublons probables', 'n', 'Alertes partageant le même nom (casse, espaces et tirets ignorés), le même sexe et le même âge.', 'Les doublons gonflent le nombre d’alertes et de cas. Toute valeur > 0 doit être vérifiée.', 'Voir l’onglet « Doublons ».']
      ],
      delais: [
        ['delNotif', 'Délai moyen DDS → notification', 'd', 'Moyenne (notification − début des symptômes), délais ≥ 0.', 'Réactivité de la détection : un délai élevé laisse circuler les cas sans détection.', 'Plus court = détection plus rapide.'],
        ['pPrelCible', 'Prélevées dans le délai cible', 'p', 'Alertes prélevées entre 0 et ' + s.CiblePrelev + ' j après la notification ÷ alertes avec un délai calculable.', 'Rapidité de la riposte après une alerte.', 'Plus c’est haut, mieux c’est.', true],
        ['delLabo', 'Délai moyen prélèvement → analyse', 'd', 'Moyenne (analyse − prélèvement), délais ≥ 0.', 'Performance de la chaîne laboratoire.', 'Un délai long signale un problème d’acheminement ou de laboratoire.'],
        ['discord', 'Discordance « confirmé »', 'n', '« Cas confirmés » = 1 alors que la classification n’est pas « Cas confirmé », ou l’inverse.', 'Le nombre de cas confirmés diffère selon le champ utilisé.', 'Doit être 0 : harmoniser les deux champs dans DHIS2.', 'zero'],
        ['statutManq', 'Statut final manquant (confirmés)', 'n', 'Cas confirmés sans statut final (vivant / décédé).', 'Sans statut final, la létalité est mal calculée.', 'Doit être 0.', 'zero'],
        ['fraicheur', 'Fraîcheur des données', 'h', 'Heures écoulées depuis la dernière extraction DHIS2.', 'Au-delà de 24 h, les chiffres ne reflètent plus la situation actuelle.', 'Publier de nouveau depuis l’ordinateur si la valeur dépasse ~24 h.']
      ],
      retards: [
        ['retInv', 'Investigations en retard', 'n', 'Conclusion vide ou « En cours… » depuis plus de ' + s.RetInv + ' j après la notification.', 'Risque de cas non isolés.', 'Relancer les équipes d’investigation.'],
        ['retLabo', 'Échantillons non analysés', 'n', 'Prélevés sans date d’analyse depuis plus de ' + s.RetLabo + ' j.', 'Cas potentiellement positifs non confirmés.', 'Relancer le laboratoire / le transport.'],
        ['retClassif', 'Validées non classées', 'n', 'Validées sans classification finale (ou « Cas suspect ») depuis plus de ' + s.RetClassif + ' j.', 'Le nombre réel de cas est sous-estimé ou incertain.', 'Classer dans DHIS2.'],
        ['sansConcl', 'Alertes sans conclusion', 'n', 'Alertes dont le champ « Conclusion alerte » est vide (toutes périodes).', 'Issue inconnue : impossible de savoir si l’alerte est à suivre ou à écarter.', 'Compléter dans DHIS2.'],
        ['derniere', 'Dernière notification', 'date', 'Date de notification la plus récente (zone et aire choisies).', 'Dernière activité de saisie dans ce périmètre.', 'Permet de repérer une zone qui ne saisit plus.'],
        ['joursSans', 'Jours sans saisie', 'j', 'Date de référence − dernière notification.', '0–1 j : saisie régulière. Plusieurs jours : silence épidémiologique ou rupture de remontée.', 'Vérifier auprès de la zone.']
      ]
    };
  }
  function fmtVal(v, t) {
    if (v == null) return '–';
    if (t === 'p') return pf.format(v); if (t === 'd') return df1.format(v) + ' j'; if (t === 'j') return nf.format(v) + ' j';
    if (t === 'h') return nf.format(v) + ' h'; if (t === 'date') return fmtDate(v); return nf.format(v);
  }
  function renderSynthese() {
    var k = Object.assign({}, state.result.k), f = state.filters, s = seuils();
    k.fraicheur = state.data.extrait != null ? Math.round((MVEQ.nowSerial() - state.data.extrait) * 24) : null;
    var T = tiles();
    Object.keys(T).forEach(function (g) {
      var host = $('#tiles-' + g);
      host.replaceChildren.apply(host, T[g].map(function (t) {
        var b = document.createElement('button'), v = k[t[0]];
        b.className = 'tile';
        if (t[6] === true) b.className += ' q-' + (cls(v) || 'x');
        else if (t[6] === 'zero' && v != null) b.className += v === 0 ? ' q-good' : ' q-bad';
        else if (t[0] === 'fraicheur' && v != null) b.className += v <= 24 ? ' q-good' : v <= 48 ? ' q-mid' : ' q-bad';
        else if (/^ret/.test(t[0]) || t[0] === 'sansConcl') b.className += v ? ' q-bad' : ' q-good';
        var lab = document.createElement('span'); lab.className = 'lab'; lab.textContent = t[1];
        var val = document.createElement('span'); val.className = 'val'; val.textContent = fmtVal(v, t[2]);
        b.append(lab, val);
        b.addEventListener('click', function () {
          $('#t-def').textContent = t[1]; $('#def-val').textContent = fmtVal(v, t[2]);
          $('#def-def').textContent = t[3]; $('#def-int').textContent = t[4]; $('#def-act').textContent = t[5]; openSheet('#sheet-def');
        });
        return b;
      }));
    });
    $('#sub-periode').textContent = 'Alertes notifiées du ' + fmtDate(f.debut) + ' au ' + fmtDate(f.fin) + '.';
    $('#sub-retards').textContent = 'Toutes périodes, à la date de référence (' + fmtDate(f.ref) + ').';
    var lieu = f.aire ? 'l’aire ' + de(f.aire) : f.zone ? 'la zone ' + de(f.zone) : 'le Haut-Uélé';
    var parts = ['Dans ', lieu, ', ', ['b', nf.format(k.alertes)], ' alertes analysées'];
    if (k.alertes) parts.push(' : ', ['b', pf.format(k.pSans)], ' sans anomalie, complétude ', ['b', k.completude == null ? '–' : pf.format(k.completude)], '. ', ['b', nf.format(k.avecAnom)], ' alerte' + (k.avecAnom > 1 ? 's' : '') + ' à corriger dans DHIS2.');
    else parts.push('.');
    var ret = k.retInv + k.retLabo + k.retClassif;
    if (ret) parts.push(' ', ['b', nf.format(ret)], ' traitement' + (ret > 1 ? 's' : '') + ' en retard.');
    var lede = $('#lede');
    lede.replaceChildren.apply(lede, parts.map(function (p) { if (typeof p === 'string') return document.createTextNode(p); var e = document.createElement(p[0]); e.textContent = p[1]; return e; }));
    $('#note-seuils').textContent = 'Vert ≥ ' + pf.format(s.Vert) + ', orange ≥ ' + pf.format(s.Orange) + ', rouge en dessous. Les zones sans alerte sur la période restent vides.';
  }
  function de(n) { return /^[aeiouyéèêâîôûh]/i.test(n) ? 'd’' + n : 'de ' + n; }

  /* ---------- Graphiques SVG ---------- */
  var NS = 'http://www.w3.org/2000/svg';
  function el(tag, a, p) { var e = document.createElementNS(NS, tag); for (var k in a) e.setAttribute(k, a[k]); if (p) p.appendChild(e); return e; }
  function tw(svg, t) { var x = el('text', { x: -999, y: -999 }, svg); x.textContent = t; var w = x.getComputedTextLength ? x.getComputedTextLength() : t.length * 6.5; svg.removeChild(x); return w || t.length * 6.5; }
  function hbars(host, items, o) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), rowH = 26, H = items.length * rowH + 4;
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': o.aria }, host);
    var lw = Math.min(W * 0.48, Math.max.apply(null, items.map(function (i) { return tw(svg, i.label); })) + 10);
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
  function semaines(host, data) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), H = 220, svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Alertes et qualité par semaine' }, host);
    var maxA = Math.max(1, Math.max.apply(null, data.map(function (d) { return d.alertes; })));
    var nice = function (v) { if (v <= 5) return 5; var p = Math.pow(10, Math.floor(Math.log10(v))), m = v / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 5 ? 5 : 10) * p; };
    maxA = nice(maxA);
    var ml = tw(svg, nf.format(maxA)) + 8, mr = tw(svg, '100 %') + 8, mt = 8, mb = 24, iw = W - ml - mr, ih = H - mt - mb, bw = iw / data.length;
    var yA = function (v) { return mt + ih - v / maxA * ih; }, yP = function (v) { return mt + ih - v * ih; };
    [0, .5, 1].forEach(function (p) {
      el('line', { x1: ml, x2: W - mr, y1: yP(p), y2: yP(p), style: 'stroke:var(--line)' }, svg);
      var a = el('text', { x: ml - 6, y: yP(p), 'text-anchor': 'end', dy: '0.32em' }, svg); a.textContent = nf.format(maxA * p);
      var b = el('text', { x: W - mr + 6, y: yP(p), dy: '0.32em' }, svg); b.textContent = pf.format(p);
    });
    data.forEach(function (d, i) {
      var x = ml + i * bw;
      if (d.alertes) el('rect', { x: x + bw * .15, width: bw * .7, y: yA(d.alertes), height: ih - (yA(d.alertes) - mt), rx: 2, style: 'fill:var(--alert)' }, svg);
      if (i % 2 === 0) { var t = el('text', { x: x + bw / 2, y: H - 6, 'text-anchor': 'middle' }, svg); t.textContent = fmtDate(d.semaine, true); }
    });
    [['pSans', 'var(--ok)'], ['completude', 'var(--men)']].forEach(function (s) {
      var seg = [], segs = [];
      data.forEach(function (d, i) { if (d[s[0]] == null) { if (seg.length) segs.push(seg); seg = []; } else seg.push((ml + i * bw + bw / 2) + ',' + yP(d[s[0]])); });
      if (seg.length) segs.push(seg);
      segs.forEach(function (sg) { el('polyline', { points: sg.join(' '), fill: 'none', 'stroke-width': 2, style: 'stroke:' + s[1] }, svg); });
      data.forEach(function (d, i) { if (d[s[0]] != null) el('circle', { cx: ml + i * bw + bw / 2, cy: yP(d[s[0]]), r: 2.8, style: 'fill:' + s[1] }, svg); });
    });
    // Infobulle au toucher
    var tip = document.createElement('div'); tip.className = 'tip'; tip.hidden = true; host.appendChild(tip);
    var hit = el('rect', { x: ml, y: mt, width: iw, height: ih, fill: 'transparent' }, svg);
    var at = function (ev) {
      var r = svg.getBoundingClientRect(), sx = (ev.clientX - r.left) * (W / r.width), i = Math.floor((sx - ml) / bw);
      if (i < 0 || i >= data.length) return; var d = data[i];
      tip.replaceChildren(); var b = document.createElement('b'); b.textContent = 'Semaine du ' + fmtDate(d.semaine); tip.appendChild(b);
      ['Alertes : ' + nf.format(d.alertes), 'Sans anomalie : ' + (d.pSans == null ? '–' : pf.format(d.pSans)), 'Complétude : ' + (d.completude == null ? '–' : pf.format(d.completude))]
        .forEach(function (l) { var x = document.createElement('div'); x.textContent = l; tip.appendChild(x); });
      tip.hidden = false; var w = tip.offsetWidth;
      tip.style.left = Math.max(0, Math.min(host.clientWidth - w, (ml + i * bw + bw / 2) * (r.width / W) - w / 2)) + 'px'; tip.style.top = '0px';
    };
    hit.addEventListener('pointerdown', at); hit.addEventListener('pointermove', at);
    hit.addEventListener('pointerleave', function (e) { if (e.pointerType === 'mouse') tip.hidden = true; });
  }
  function renderGraphiques() {
    if (state.tab !== 'graphiques') return;
    var r = state.result, s = seuils();
    hbars($('#ch-compl'), r.parVariable.map(function (v) {
      return { label: v.variable, value: v.taux, color: v.taux == null ? 'var(--line)' : v.taux >= s.Vert ? 'var(--ok)' : v.taux >= s.Orange ? 'var(--warn)' : 'var(--bad)' };
    }), { max: 1, fmt: function (v) { return pf.format(v); }, aria: 'Complétude des variables clés' });
    hbars($('#ch-incoh'), r.parType.map(function (t) { return { label: t.type, value: t.alertes }; }),
      { color: 'var(--bad)', fmt: function (v) { return nf.format(v); }, aria: 'Incohérences par type' });
    semaines($('#ch-sem'), r.semaines);
  }

  function td(t, c) { var e = document.createElement('td'); e.textContent = t; if (c) e.className = c; return e; }
  function pctTd(v) { var e = document.createElement('td'); e.className = 'n'; if (v == null) { e.textContent = ''; return e; } var s = document.createElement('span'); s.className = 'pct ' + cls(v); s.textContent = pf.format(v); e.appendChild(s); return e; }
  function renderZones() {
    if (state.tab !== 'zones') return;
    var r = state.result, tb = $('#tb-zones');
    var row = function (z, total) {
      var tr = document.createElement('tr'); if (total) tr.className = 'total';
      tr.appendChild(td(z.zone, 'first')); tr.appendChild(td(nf.format(z.alertes), 'n'));
      tr.appendChild(pctTd(z.alertes ? z.pSans : null)); tr.appendChild(pctTd(z.alertes ? z.completude : null));
      tr.appendChild(td(nf.format(z.avecIncoh), 'n')); tr.appendChild(td(nf.format(z.doublons), 'n'));
      tr.appendChild(td(z.delNotif == null ? '–' : df1.format(z.delNotif), 'n')); tr.appendChild(td(z.delLabo == null ? '–' : df1.format(z.delLabo), 'n'));
      [z.retInv, z.retLabo, z.retClassif].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); });
      if (!total) {
        tr.style.cursor = 'pointer';
        tr.addEventListener('click', function () { var s = lsGet('mveq-filtres') || {}; s.zone = z.zone; s.aire = ''; lsSet('mveq-filtres', s); state.filters = resolveFilters(); render(); toast('Filtre : zone ' + de(z.zone) + '.'); });
      }
      return tr;
    };
    tb.replaceChildren.apply(tb, r.zones.map(function (z) { return row(z); }).concat([row(r.total, true)]));
  }

  function item(r, detail) {
    var d = document.createElement('div'); d.className = 'item';
    var top = document.createElement('div'); top.className = 'top-line';
    var id = document.createElement('span'); id.className = 'idn'; id.textContent = r.id || '(sans N° d’alerte)';
    top.appendChild(id);
    if (detail) { var nb = document.createElement('span'); nb.className = 'nb'; nb.textContent = r.nbAnom + ' anomalie' + (r.nbAnom > 1 ? 's' : ''); top.appendChild(nb); }
    var meta = document.createElement('div'); meta.className = 'meta';
    meta.textContent = 'Notifiée le ' + fmtDate(r.n) + ' · ' + (r.zone || '–') + ' › ' + (r.aire || '–') + (r.conclusion ? ' · ' + r.conclusion : '') + (r.classif ? ' · ' + r.classif : '');
    d.append(top, meta);
    if (detail) {
      var tags = document.createElement('div'); tags.className = 'tags';
      detail.forEach(function (x) { var s = document.createElement('span'); if (!/^manque/.test(x)) s.className = 'inc'; s.textContent = x; tags.appendChild(s); });
      d.appendChild(tags);
    }
    return d;
  }
  function renderCorriger() {
    if (state.tab !== 'corriger') return;
    var q = MVEQ.norm($('#q-corriger').value), all = state.result.aCorriger;
    var list = q ? all.filter(function (a) { return MVEQ.norm((a.r.id || '') + ' ' + (a.r.aire || '') + ' ' + (a.r.zone || '')).indexOf(q) >= 0; }) : all;
    $('#sub-corriger').textContent = nf.format(all.length) + ' alerte' + (all.length > 1 ? 's' : '') + ' avec au moins une variable manquante ou une incohérence (période et lieu choisis), les plus récentes d’abord. Corrigez-les dans DHIS2 à partir du N° d’alerte.';
    var host = $('#list-corriger');
    host.replaceChildren.apply(host, list.slice(0, state.nCorr).map(function (a) { return item(a.r, a.detail); }));
    if (!list.length) { var p = document.createElement('p'); p.className = 'sub'; p.textContent = q ? 'Aucune alerte ne correspond à la recherche.' : 'Aucune alerte à corriger sur cette période.'; host.appendChild(p); }
    $('#more-corriger').hidden = list.length <= state.nCorr;
    $('#more-corriger').textContent = 'Afficher plus (' + nf.format(list.length - state.nCorr) + ' restantes)';
  }
  function renderDoublons() {
    if (state.tab !== 'doublons') return;
    var g = state.result.doublons, n = g.reduce(function (t, x) { return t + x.membres.length; }, 0);
    $('#sub-doublons').textContent = nf.format(n) + ' alerte' + (n > 1 ? 's' : '') + ' dans ' + nf.format(g.length) + ' groupe' + (g.length > 1 ? 's' : '') +
      ' : même nom, même sexe et même âge. Les noms ne sont pas affichés : retrouvez les fiches dans DHIS2 par leur N° d’alerte, vérifiez, puis supprimez ou fusionnez les doublons.';
    var host = $('#list-doublons');
    host.replaceChildren.apply(host, g.slice(0, state.nDoub).map(function (grp, i) {
      var box = document.createElement('div'); box.className = 'group';
      var h = document.createElement('h3'); var a = document.createElement('span'); a.textContent = 'Groupe ' + (i + 1);
      var m0 = grp.membres[0], b = document.createElement('span');
      b.textContent = (m0.sexe || 'sexe ?') + ' · ' + (m0.ageAns != null ? m0.ageAns + ' ans' : m0.ageMois != null ? m0.ageMois + ' mois' : 'âge ?') + ' · ' + grp.total + ' alertes au total';
      h.append(a, b); box.appendChild(h);
      grp.membres.forEach(function (r) {
        var m = document.createElement('div'); m.className = 'mem';
        var id = document.createElement('b'); id.textContent = r.id || '(sans N° d’alerte)';
        var meta = document.createElement('span'); meta.className = 'meta'; meta.style.color = 'var(--muted)';
        meta.textContent = 'Notifiée le ' + fmtDate(r.n) + ' · ' + (r.zone || '–') + ' › ' + (r.aire || '–') + (r.conclusion ? ' · ' + r.conclusion : '') + (r.classif ? ' · ' + r.classif : '');
        m.append(id, meta); box.appendChild(m);
      });
      if (grp.total > grp.membres.length) { var o = document.createElement('div'); o.className = 'mem'; o.style.color = 'var(--muted)'; o.textContent = (grp.total - grp.membres.length) + ' autre(s) alerte(s) du groupe hors du filtre.'; box.appendChild(o); }
      return box;
    }));
    if (!g.length) { var p = document.createElement('p'); p.className = 'sub'; p.textContent = 'Aucun doublon probable sur cette période.'; host.appendChild(p); }
    $('#more-doublons').hidden = g.length <= state.nDoub;
  }

  /* ---------- Onglets et feuilles ---------- */
  function showTab(name) {
    state.tab = name; lsSet('mveq-onglet', name);
    ['synthese', 'graphiques', 'zones', 'corriger', 'doublons'].forEach(function (t) { $('#tab-' + t).hidden = t !== name; });
    $$('#tabs button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.tab === name ? 'true' : 'false'); });
    window.scrollTo(0, 0);
    if (state.result) { renderGraphiques(); renderZones(); renderCorriger(); renderDoublons(); }
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
    var d = state.data, last = lsGet('mveq-sync-last');
    $('#menu-info').textContent = (d ? 'Données : ' + d.fichier + ', reçues le ' + new Date(d.chargeLe).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + '. ' : 'Aucune donnée chargée. ') +
      (sync.key ? 'Mise à jour automatique activée' + (last ? ', dernière vérification réussie le ' + new Date(last).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) : '') + '.' : 'Mise à jour automatique désactivée.') +
      (function () { var e = lsGet('mveq-sync-err'); return e && e.m ? ' Dernier échec (' + new Date(e.t).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + ') : ' + e.m : ''; })() +
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
  $('#btn-reset').addEventListener('click', function () { lsSet('mveq-filtres', {}); closeSheets(); state.filters = resolveFilters(); render(); toast('Filtres réinitialisés.'); });
  $('#q-corriger').addEventListener('input', function () { state.nCorr = 50; renderCorriger(); });
  $('#more-corriger').addEventListener('click', function () { state.nCorr += 50; renderCorriger(); });
  $('#more-doublons').addEventListener('click', function () { state.nDoub += 30; renderDoublons(); });
  var rz;
  window.addEventListener('resize', function () { clearTimeout(rz); rz = setTimeout(function () { if (state.result) renderGraphiques(); }, 150); });
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible' && Date.now() - (lsGet('mveq-sync-last') || 0) > 2 * 60000) syncNow(false); });
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
