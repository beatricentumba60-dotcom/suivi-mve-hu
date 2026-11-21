/* Interface de l'application « Suivi MVE ». Les calculs sont dans engine.js. */
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };

  var state = { data: null, filters: null, tab: 'synthese', result: null, carte: null, carteSel: null };
  var nf = new Intl.NumberFormat('fr-FR');
  var pf = new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 1 });
  var df1 = new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 });

  /* ---------- Dates (numéros de série Excel) ---------- */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function sToDate(s) { return MVE.serialToUTC(s); }
  function fmtDate(s, short) {
    if (s == null) return '–';
    var d = sToDate(s);
    return pad(d.getUTCDate()) + '/' + pad(d.getUTCMonth() + 1) + (short ? '' : '/' + d.getUTCFullYear());
  }
  function sToInput(s) { var d = sToDate(s); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate()); }
  function inputToS(v) { return v ? MVE.toSerial(v) : null; }
  function today() { return MVE.todaySerial(); }

  /* ---------- Stockage local (IndexedDB, reste sur le téléphone) ---------- */
  function idb() {
    return new Promise(function (res, rej) {
      var rq = indexedDB.open('suivi-mve', 1);
      rq.onupgradeneeded = function () { rq.result.createObjectStore('kv'); };
      rq.onsuccess = function () { res(rq.result); };
      rq.onerror = function () { rej(rq.error); };
    });
  }
  function kv(mode, fn) {
    return idb().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction('kv', mode), st = tx.objectStore('kv'), out = fn(st);
        tx.oncomplete = function () { res(out && out.result !== undefined ? out.result : undefined); };
        tx.onerror = function () { rej(tx.error); };
      });
    });
  }
  var store = {
    get: function (k) { return kv('readonly', function (s) { return s.get(k); }); },
    set: function (k, v) { return kv('readwrite', function (s) { s.put(v, k); }); },
    del: function (k) { return kv('readwrite', function (s) { s.delete(k); }); }
  };
  function lsGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* stockage indisponible */ } }

  /* ---------- Messages ---------- */
  var toastTimer;
  function toast(msg, ms) {
    var t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, ms || 3500);
  }
  function busy(on) { $('#busy').hidden = !on; }

  /* ---------- Lecture du fichier Excel ---------- */
  function readWorkbook(buf, name, garder) {
    busy(true);
    return new Promise(function (res) { setTimeout(res, 30); }).then(function () {
      var wb = XLSX.read(buf, { type: 'array', sheets: ['Données'], cellDates: false, cellFormula: false, cellHTML: false, cellText: false, dense: true });
      var sheet = wb.Sheets['Données'];
      if (!sheet) {
        // Autre nom de feuille : on cherche celle qui contient « Date notification »
        wb = XLSX.read(buf, { type: 'array', cellDates: false, cellFormula: false, cellHTML: false, cellText: false, dense: true, sheetRows: 2 });
        var found = wb.SheetNames.filter(function (n) {
          var h = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1 })[0] || [];
          return h.some(function (c) { return MVE.norm(c) === 'date notification'; });
        })[0];
        if (!found) throw new Error('Ce fichier ne contient pas la feuille « Données » de la liste linéaire MVE.');
        wb = XLSX.read(buf, { type: 'array', sheets: [found], cellDates: false, dense: true });
        sheet = wb.Sheets[found];
      }
      var table = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
      var out = MVE.normalize(table);
      var data = { rows: out.rows, extrait: out.extrait, fichier: name || 'fichier.xlsx', chargeLe: Date.now() };
      return store.set('dataset', data).catch(function () { /* stockage plein : on garde en mémoire */ }).then(function () { return data; });
    }).then(function (data) {
      busy(false);
      setData(data, !garder);
      toast(nf.format(data.rows.length) + ' alertes chargées.');
    }).catch(function (err) {
      busy(false);
      toast('Lecture impossible : ' + (err && err.message ? err.message : err), 6000);
    });
  }
  /* ---------- Mise à jour automatique ----------
     L'ordinateur publie « donnees.enc » à côté de l'application : la liste linéaire,
     compressée puis chiffrée (AES-256-GCM, clé tirée du mot de passe d'équipe par
     PBKDF2-SHA256). Le téléphone garde seulement la clé (non exportable), pas le mot de passe. */
  var REMOTE = 'donnees.enc';
  var sync = { key: null, salt: null, iter: 0, running: false };
  function b64(u8) { var s = ''; for (var i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]); return btoa(s); }
  function err(msg, code) { var e = new Error(msg); e.code = code; return e; }
  function parseHeader(buf) {
    var u8 = new Uint8Array(buf);
    if (u8.length < 40 || String.fromCharCode(u8[0], u8[1], u8[2], u8[3]) !== 'MVE1') throw err('Le fichier de données publié est illisible.');
    return { iter: new DataView(buf).getUint32(4), salt: u8.slice(8, 24), iv: u8.slice(24, 36), data: u8.slice(36) };
  }
  function deriveKey(pw, salt, iter) {
    return crypto.subtle.importKey('raw', new TextEncoder().encode(pw), 'PBKDF2', false, ['deriveKey']).then(function (base) {
      return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
    });
  }
  function decryptPayload(h, key) {
    return crypto.subtle.decrypt({ name: 'AES-GCM', iv: h.iv }, key, h.data).then(function (plain) {
      var ds = new Blob([plain]).stream().pipeThrough(new DecompressionStream('gzip'));
      return new Response(ds).text();
    }, function () { throw err('Mot de passe incorrect.', 'cle'); }).then(JSON.parse);
  }
  function fetchRemote() {
    return fetch(REMOTE + '?t=' + Date.now(), { cache: 'no-store' }).then(function (res) {
      if (res.status === 404) throw err('Aucune donnée n’a encore été publiée depuis l’ordinateur.', 'absent');
      if (!res.ok) throw err('Le site de l’application ne répond pas (erreur ' + res.status + ').');
      return res.arrayBuffer();
    }, function () { throw err('Connexion impossible : les dernières données restent affichées.', 'reseau'); });
  }
  function applyRemote(payload, force) {
    var out = MVE.normalize(payload.table), cur = state.data;
    if (!force && cur && cur.publie === payload.publie) return Promise.resolve(false);
    if (!force && cur && cur.extrait != null && out.extrait != null && out.extrait < cur.extrait) return Promise.resolve(false);
    var data = { rows: out.rows, extrait: out.extrait, fichier: 'mise à jour automatique', publie: payload.publie, chargeLe: Date.now() };
    return store.set('dataset', data).catch(function () { /* stockage plein */ }).then(function () { setData(data, !cur); return true; });
  }
  function setSyncUI() {
    $('#btn-sync').hidden = !sync.key;
    $('#btn-sync').classList.toggle('spin', sync.running);
    $('#btn-sync-menu').hidden = !sync.key;
    $('#btn-auto').textContent = sync.key ? 'Changer le mot de passe d’équipe' : 'Activer la mise à jour automatique';
    $('#btn-auto-off').hidden = !sync.key;
  }
  function syncNow(manual) {
    if (!sync.key || sync.running) return Promise.resolve();
    sync.running = true; setSyncUI();
    return fetchRemote().then(function (buf) {
      var h = parseHeader(buf);
      if (b64(h.salt) !== sync.salt || h.iter !== sync.iter) throw err('Le mot de passe d’équipe a changé.', 'cle');
      return decryptPayload(h, sync.key);
    }).then(function (p) { return applyRemote(p, false); })
      .then(function (changed) {
        lsSet('mve-sync-last', Date.now());
        if (changed) chargerCarte();
        if (changed && !manual) toast('Données mises à jour automatiquement.');
        else if (manual) toast(changed ? 'Données mises à jour.' : 'Les données sont déjà à jour.');
      })
      .catch(function (e) {
        if (e.code === 'cle') { sync.key = null; store.del('cle'); toast(e.message + ' Saisissez-le à nouveau.', 6000); openPassword(); }
        else if (manual || !state.data) toast(e.message, 5000);
      })
      .then(function () { sync.running = false; setSyncUI(); });
  }
  function openPassword() {
    $('#pw').value = ''; $('#pw-msg').textContent = ''; openSheet('#sheet-pass');
    setTimeout(function () { $('#pw').focus(); }, 50);
  }
  function connect(pw) {
    $('#pw-msg').textContent = 'Connexion…'; $('#btn-pw').disabled = true;
    fetchRemote().then(function (buf) {
      var h = parseHeader(buf);
      return deriveKey(pw, h.salt, h.iter).then(function (key) {
        return decryptPayload(h, key).then(function (p) {
          sync.key = key; sync.salt = b64(h.salt); sync.iter = h.iter;
          return store.set('cle', { key: key, salt: sync.salt, iter: h.iter }).catch(function () { /* clé gardée en mémoire */ })
            .then(function () { closeSheets(); return applyRemote(p, true); })
            .then(function () { lsSet('mve-sync-last', Date.now()); setSyncUI(); toast('Mise à jour automatique activée.'); });
        });
      });
    }).catch(function (e) { $('#pw-msg').textContent = e.message; })
      .then(function () { $('#btn-pw').disabled = false; });
  }

  // Programme d'ordinateur (dossier de classeurs) : lecture directe d'un classeur sans passer par le sélecteur de fichier
  window.MVE_lireClasseur = function (buf, name, garder) { return readWorkbook(buf, name, garder); };
  function openFile(file) {
    if (!file) return;
    file.arrayBuffer().then(function (b) { return readWorkbook(new Uint8Array(b), file.name); });
  }

  /* ---------- Filtres ---------- */
  function minNotif(rows) { var m = Infinity; rows.forEach(function (r) { if (r.n < m) m = r.n; }); return m; }
  function maxNotif(rows) { var m = 0; rows.forEach(function (r) { if (r.n > m) m = r.n; }); return m; }
  function resolveFilters() {
    var saved = lsGet('mve-filtres') || {}, rows = state.data.rows, t = today();
    var f = {
      zone: saved.zone || '', aire: saved.aire || '',
      periode: saved.periode || 'tout', debutFixe: saved.debutFixe, finFixe: saved.finFixe,
      refMode: saved.refMode || '0', refFixe: saved.refFixe
    };
    if (f.zone && MVE.zonesOf(rows).indexOf(f.zone) < 0) { f.zone = ''; f.aire = ''; }
    f.ref = f.refMode === 'fixe' ? f.refFixe : f.refMode === 'data' ? maxNotif(rows) : t - (+f.refMode || 0);
    if (f.periode === 'fixe') { f.debut = f.debutFixe; f.fin = f.finFixe; }
    else if (f.periode === 'tout') { f.debut = minNotif(rows); f.fin = Math.max(t, maxNotif(rows)); }
    else { f.fin = f.ref; f.debut = f.ref - (+f.periode) + 1; }
    return f;
  }
  function saveFilters(f) {
    lsSet('mve-filtres', { zone: f.zone, aire: f.aire, periode: f.periode, debutFixe: f.debutFixe, finFixe: f.finFixe, refMode: f.refMode, refFixe: f.refFixe });
  }

  function fillZoneSelect(sel, value) {
    sel.replaceChildren(new Option('Toutes les zones', ''));
    MVE.zonesOf(state.data.rows).forEach(function (z) { sel.appendChild(new Option(z, z)); });
    sel.value = value || '';
  }
  function fillAireSelect(sel, zone, value) {
    sel.replaceChildren(new Option(zone ? 'Toutes les aires de ' + zone : 'Toutes les aires', ''));
    MVE.airesOf(state.data.rows, zone).forEach(function (a) { sel.appendChild(new Option(a, a)); });
    sel.value = value || '';
    if (sel.value !== (value || '')) sel.value = '';
  }

  var draft = null;
  function openFilters() {
    var f = state.filters;
    draft = JSON.parse(JSON.stringify(f));
    fillZoneSelect($('#f-zone'), f.zone);
    fillAireSelect($('#f-aire'), f.zone, f.aire);
    $('#f-debut').value = sToInput(f.debut); $('#f-fin').value = sToInput(f.fin); $('#f-ref').value = sToInput(f.ref);
    openSheet('#sheet-filtres');
  }
  function applyFilters() {
    var f = draft;
    f.zone = $('#f-zone').value; f.aire = $('#f-aire').value;
    var d = inputToS($('#f-debut').value), e = inputToS($('#f-fin').value), r = inputToS($('#f-ref').value);
    if (d == null || e == null || r == null) { toast('Renseignez les trois dates.'); return; }
    if (d > e) { toast('La date de début est après la date de fin.'); return; }
    if (f.periode === 'fixe') { f.debutFixe = d; f.finFixe = e; }
    if (f.refMode === 'fixe') f.refFixe = r;
    saveFilters(f);
    closeSheets();
    state.filters = resolveFilters();
    render();
  }

  /* ---------- Rendu ---------- */
  function setData(data, fresh) {
    state.data = data;
    state.filters = resolveFilters();
    $('#empty').hidden = true; $('#tabs').hidden = false; $('#chips').hidden = false;
    showTab(fresh ? 'synthese' : (lsGet('mve-onglet') || 'synthese'));
    render();
  }

  function render() {
    var d = state.data, f = state.filters;
    state.result = MVE.compute(d.rows, f);
    var ex = d.extrait != null ? sToDate(d.extrait) : null;
    $('#src').textContent = 'DHIS2 SNIS · extrait le ' + (ex ? pad(ex.getUTCDate()) + '/' + pad(ex.getUTCMonth() + 1) + '/' + ex.getUTCFullYear() + ' à ' + pad(ex.getUTCHours()) + 'h' + pad(ex.getUTCMinutes()) : '(date inconnue)') +
      ' · ' + nf.format(d.rows.length) + ' alertes';
    // Alerte de fraîcheur : extraction de plus de 2 jours
    var age = d.extrait != null ? (Date.now() - (sToDate(d.extrait).getTime() + new Date().getTimezoneOffset() * 60000)) / 864e5 : null;
    $('#src').classList.toggle('stale', age != null && age > 2);
    renderChips();
    renderSynthese();
    renderCourbes();
    renderZones();
    renderCarteTab();
    renderProfessions();
  }

  function renderChips() {
    var f = state.filters, c = $('#chips');
    var lieu = f.aire ? f.aire + ' (' + (f.zone || 'aire') + ')' : (f.zone || 'Toutes les zones');
    var per = f.periode === 'tout' ? 'Toute l’épidémie' : fmtDate(f.debut) + ' → ' + fmtDate(f.fin);
    var items = [['Lieu', lieu], ['Période', per], ['Réf.', fmtDate(f.ref)]];
    c.replaceChildren.apply(c, items.map(function (it) {
      var b = document.createElement('button'); b.className = 'chip';
      var bb = document.createElement('b'); bb.textContent = it[1];
      b.append(it[0] + ' : ', bb); b.addEventListener('click', openFilters); return b;
    }));
  }

  /* Indicateurs : libellé, valeur, définition (dictionnaire du classeur) */
  var P = 'Zone + aire + période (date de notification)';
  var R = 'Zone + aire + date de référence (ignore la période)';
  var TILES = {
    cas: [
      ['cas', 'Cas (conf. + prob.)', 'n', 'Nombre total de cas MVE : confirmés + probables.', 'Ampleur de l’épidémie sur la période. Dénominateur de la létalité, du % vaccinés et du % contact connu.', P, 'hero'],
      ['conf', 'Confirmés', 'n', 'Cas confirmés au laboratoire (« Cas confirmés » = 1).', 'Mesure la plus fiable : à privilégier pour comparer les zones et les périodes.', P, 'k-conf'],
      ['prob', 'Probables', 'n', 'Classification finale = « Cas probable ».', 'Cas non testés, souvent des décès communautaires. Une part élevée signale des décès non prélevés.', P, 'k-prob'],
      ['deces', 'Décès (cas)', 'n', 'Décès confirmés (indicateur DHIS2 « Décès confirmés ») + cas probables décédés (statut final « Décédé », nature « Décès », statut avant admission « Décédé » ou date de décès renseignée).', 'Gravité et retard de prise en charge. Les décès d’alertes non-cas ne sont pas comptés.', P, 'k-dec'],
      ['decesConf', 'Décès confirmés', 'n', 'Indicateur DHIS2 « Décès confirmés » (Ox52yq3XZUX) : alerte validée, N° Epid et sexe renseignés, classification « Cas confirmé », et statut final « Décédé », issue « Décédé(e) » à la sortie ou nature « Décès ». Même chiffre que la tuile « Cumul Décès Confirmés » du tableau de bord DHIS2.', 'Décès parmi les cas confirmés. DHIS2 compte par date d’inscription, l’application par date de notification : sur une période partielle, un petit écart reste possible.', P, 'k-dec'],
      ['letalite', 'Létalité', 'p', 'Décès (cas) ÷ Cas (conf. + prob.).', 'Létalité brute : les cas récents dont l’issue n’est pas encore connue la font sous-estimer.', P],
      ['casJour', 'Nouveaux cas – jour de réf.', 'n', 'Cas dont la date de notification est la date de référence.', 'À lire avec la tendance sur 7 jours et la courbe quotidienne.', R],
      ['cas7', 'Nouveaux cas (7 j)', 'n', 'Cas notifiés entre la date de référence − 6 jours et la date de référence.', 'Principal indicateur d’évolution : hausse = transmission active ; baisse durable = contrôle en cours.', R],
      ['joursDepuis', 'Jours depuis le dernier cas confirmé', 'j', 'Date de référence − date de notification du dernier cas confirmé.', '< 21 j : transmission possible ; ≥ 21 j : une période d’incubation sans cas ; ≥ 42 j : critère OMS de fin d’épidémie.', R],
      ['delai', 'Délai moyen DDS → notification', 'd', 'Moyenne de (notification − début des symptômes) chez les cas confirmés, délais ≥ 0.', 'Rapidité de détection : plus il est court, plus les cas sont isolés tôt. Les fiches sans DDS sont exclues.', P],
      ['pContact', '% cas avec contact Ebola connu', 'p', 'Cas avec « Contact avec malade Ebola » = Oui ÷ Cas.', 'Une part faible révèle des chaînes de transmission inconnues : renforcer le suivi des contacts.', P],
      ['decesCom', 'Décès communautaires (cas)', 'n', 'Cas dont la nature de l’alerte est « Décès ».', 'Décès hors structure : détection tardive, risque d’enterrements non sécurisés.', P],
      ['pVacc', '% cas vaccinés MVE', 'p', 'Cas avec « Vacciné MVE » = Oui ÷ Cas.', 'Les réponses Non, Inconnu et vides restent au dénominateur.', P],
      ['cte', 'Admis au CTE sans sortie', 'n', 'Date d’admission CTE renseignée, sans date de sortie et non décédé.', 'Charge actuelle des CTE (lits occupés), surestimée si les sorties ne sont pas saisies.', 'Zone + aire uniquement (ignore période et date de référence)']
    ],
    alertes: [
      ['alertes', 'Alertes notifiées', 'n', 'Toutes les alertes dont la date de notification est dans la période.', 'Sensibilité de la surveillance : une chute brutale peut signaler un relâchement.', P],
      ['alertesJour', 'Alertes – jour de réf.', 'n', 'Alertes notifiées le jour de référence.', 'Activité de surveillance du jour.', R],
      ['validees', 'Alertes validées', 'n', 'Conclusion de l’alerte = « Validée ».', 'Alertes répondant à la définition de cas suspect, à prélever et isoler.', P],
      ['pValidees', '% alertes validées', 'p', 'Alertes validées ÷ alertes notifiées.', 'Un taux très bas ou très élevé interroge la qualité du tri des alertes.', P],
      ['attente', 'Validées non classées / suspects', 'n', 'Conclusion « Validée » et classification vide ou « Cas suspect ».', 'Charge en attente de résultat de laboratoire ou de saisie : à résorber rapidement.', P],
      ['prelevees', 'Alertes prélevées', 'n', 'Date de prélèvement renseignée.', 'Idéalement toutes les alertes validées sont prélevées.', P],
      ['prelevees', 'Échantillons reçus', 'n', 'Date de prélèvement renseignée (même calcul que « Alertes prélevées » : aucune date de réception n’est importée).', 'Volume arrivé au laboratoire.', P],
      ['analysees', 'Échantillons analysés', 'n', 'Date d’analyse labo renseignée.', 'L’écart avec les prélevées mesure le retard d’analyse.', P],
      ['noncas', 'Non-cas', 'n', 'Indicateur DHIS2 « Non Cas » (mkyT4tONdVc) : alerte validée avec N° Epid et sexe, résultat de laboratoire négatif sans aucun positif, ou classification finale « Non cas ». Même chiffre que la tuile « Cumul Non Cas » du tableau de bord DHIS2.', 'Alertes écartées : un nombre élevé témoigne d’une surveillance sensible.', P],
      ['alertesDeces', 'Alertes décès', 'n', 'Alertes dont la nature est « Décès », qu’elles soient des cas ou non.', 'Décès signalés à la surveillance : chacun doit être investigué, prélevé (post-mortem) et inhumé de façon digne et sécurisée.', P]
    ],
    groupes: [
      ['enceintes', 'Femmes enceintes (cas)', 'n', 'Cas avec « Enceinte » = Oui.', 'Groupe à risque élevé : prise en charge spécifique et suivi du nouveau-né.', P],
      ['allaitantes', 'Femmes allaitantes (cas)', 'n', 'Cas avec « Allaite » = Oui.', 'Suivre l’enfant comme contact à haut risque.', P],
      ['ps', 'Personnel de santé infectés', 'n', 'Cas dont la profession est « Personnel de santé ».', 'Signal d’alerte sur la prévention et le contrôle des infections. Sous-estimé si la profession n’est pas renseignée.', P],
      ['ppl', 'PPL infectés', 'n', 'Cas dont la profession est « PPL » (personnel de première ligne).', 'Signal d’alerte sur la PCI et le port des équipements de protection.', P],
      ['psppl', 'Total PS + PPL infectés', 'n', 'Personnel de santé + PPL infectés.', 'Charge totale d’infections chez les intervenants de la santé et de la riposte.', P]
    ]
  };
  function fmtVal(v, type) {
    if (v == null) return '–';
    if (type === 'p') return pf.format(v);
    if (type === 'd') return df1.format(v) + ' j';
    if (type === 'j') return nf.format(v) + ' j';
    return nf.format(v);
  }
  function joursClass(j) { return j == null ? '' : j < 21 ? 'bad' : j < 42 ? 'warn' : 'ok'; }

  function renderSynthese() {
    var k = state.result.k, f = state.filters;
    Object.keys(TILES).forEach(function (grp) {
      var host = $('#tiles-' + grp);
      host.replaceChildren.apply(host, TILES[grp].map(function (t) {
        var b = document.createElement('button');
        b.className = 'tile' + (t[6] ? ' ' + t[6] : '');
        if (t[0] === 'joursDepuis' && k.joursDepuis != null) b.className += ' s-' + joursClass(k.joursDepuis);
        var lab = document.createElement('span'); lab.className = 'lab'; lab.textContent = t[1];
        var val = document.createElement('span'); val.className = 'val'; val.textContent = fmtVal(k[t[0]], t[2]);
        b.append(lab, val);
        if (t[6] && t[6].indexOf('k-') === 0) { var tag = document.createElement('i'); tag.className = 'tag'; b.appendChild(tag); }
        if (t[0] === 'joursDepuis' && k.dernierConf) { var n = document.createElement('span'); n.className = 'note'; n.textContent = 'dernier le ' + fmtDate(k.dernierConf); b.appendChild(n); }
        if (t[0] === 'cas') { var n2 = document.createElement('span'); n2.className = 'note'; n2.textContent = nf.format(k.conf) + ' confirmés · ' + nf.format(k.prob) + ' probables'; b.appendChild(n2); }
        b.addEventListener('click', function () { showDef(t, fmtVal(k[t[0]], t[2])); });
        return b;
      }));
    });
    // Phrase de synthèse
    var de = function (n) { return /^[aeiouyéèêâîôûh]/i.test(n) ? 'd’' + n : 'de ' + n; };
    var lede = $('#lede'), lieu = f.aire ? 'l’aire ' + de(f.aire) : f.zone ? 'la zone ' + de(f.zone) : 'le Haut-Uélé';
    var parts = [
      'Dans ', lieu, ', ', f.periode === 'tout' ? 'depuis le début' : 'du ' + fmtDate(f.debut) + ' au ' + fmtDate(f.fin), ' : ',
      ['b', nf.format(k.cas)], ' cas, dont ', ['b', nf.format(k.deces)], ' décès (létalité ', ['b', pf.format(k.letalite)], '). ',
      ['b', nf.format(k.cas7)], ' nouveaux cas sur les 7 jours précédant le ' + fmtDate(f.ref) + '. '
    ];
    if (k.joursDepuis == null) parts.push('Aucun cas confirmé à cette date.');
    else parts.push('Dernier cas confirmé il y a ', ['b', nf.format(k.joursDepuis) + ' jour' + (k.joursDepuis > 1 ? 's' : '')], '.');
    lede.replaceChildren.apply(lede, parts.map(function (p) {
      if (typeof p === 'string') return document.createTextNode(p);
      var e = document.createElement(p[0]); e.textContent = p[1]; return e;
    }));
  }

  function showDef(t, val) {
    $('#t-def').textContent = t[1]; $('#def-val').textContent = val;
    $('#def-def').textContent = t[3]; $('#def-int').textContent = t[4]; $('#def-fil').textContent = t[5];
    openSheet('#sheet-def');
  }

  /* ---------- Graphiques SVG ---------- */
  var NS = 'http://www.w3.org/2000/svg';
  function svgEl(tag, attrs, parent) {
    var e = document.createElementNS(NS, tag);
    for (var a in attrs) e.setAttribute(a, attrs[a]);
    if (parent) parent.appendChild(e);
    return e;
  }
  function niceMax(v) {
    if (v <= 4) return Math.max(v, 1) <= 2 ? 2 : 4;
    var p = Math.pow(10, Math.floor(Math.log10(v))), m = v / p;
    var n = m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10;
    return n * p;
  }
  function textWidth(svg, txt) {
    var t = svgEl('text', { x: -999, y: -999 }, svg); t.textContent = txt;
    var w = t.getComputedTextLength ? t.getComputedTextLength() : txt.length * 6.5;
    svg.removeChild(t); return w || txt.length * 6.5;
  }
  function makeTip(host) {
    var tip = document.createElement('div'); tip.className = 'tip'; tip.hidden = true; host.appendChild(tip);
    return {
      show: function (x, y, title, lines) {
        tip.replaceChildren(); var b = document.createElement('b'); b.textContent = title; tip.appendChild(b);
        lines.forEach(function (l) { var s = document.createElement('div'); s.textContent = l; tip.appendChild(s); });
        tip.hidden = false;
        var w = tip.offsetWidth, W = host.clientWidth;
        tip.style.left = Math.max(0, Math.min(W - w, x - w / 2)) + 'px';
        tip.style.top = Math.max(0, y - tip.offsetHeight - 10) + 'px';
      },
      hide: function () { tip.hidden = true; }
    };
  }

  /* Colonnes verticales empilées, avec courbes optionnelles */
  function columns(host, o) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), H = o.height || 210;
    var svg = svgEl('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': o.aria }, host);
    var data = o.data;
    var stackMax = Math.max.apply(null, data.map(function (d) { return o.bars.reduce(function (t, s) { return t + (d[s.key] || 0); }, 0); }).concat((o.lines || []).map(function (s) { return Math.max.apply(null, data.map(function (d) { return d[s.key] || 0; })); })).concat([0]));
    var max = niceMax(stackMax || 1);
    var ml = Math.ceil(textWidth(svg, nf.format(max))) + 8, mr = 6, mt = 8, mb = 24;
    var iw = W - ml - mr, ih = H - mt - mb, bw = iw / data.length;
    var y = function (v) { return mt + ih - (v / max) * ih; };
    var g = svgEl('g', { class: 'grid' }, svg);
    [0, .25, .5, .75, 1].forEach(function (p) {
      var v = max * p; if (max <= 4 && v % 1) return;
      svgEl('line', { x1: ml, x2: W - mr, y1: y(v), y2: y(v) }, g);
      var t = svgEl('text', { x: ml - 6, y: y(v), 'text-anchor': 'end', dy: '0.32em' }, svg); t.textContent = nf.format(v);
    });
    // étiquettes de l'axe X espacées pour ne pas se chevaucher
    var lw = textWidth(svg, o.xLabel(data[0])) + 10, every = Math.max(1, Math.ceil(lw / bw));
    data.forEach(function (d, i) {
      var x = ml + i * bw, y0 = mt + ih;
      o.bars.forEach(function (s) {
        var v = d[s.key] || 0; if (!v) return;
        var h = (v / max) * ih;
        svgEl('rect', { x: x + bw * 0.12, width: Math.max(1, bw * 0.76), y: y0 - h, height: h, rx: Math.min(2, bw * 0.2), style: 'fill:' + s.color }, svg);
        y0 -= h;
      });
      if (i % every === 0 && (data.length - i) >= every * 0.6) {
        var t = svgEl('text', { x: x + bw / 2, y: H - 6, 'text-anchor': 'middle' }, svg); t.textContent = o.xLabel(d);
      }
    });
    (o.lines || []).forEach(function (s) {
      var pts = data.map(function (d, i) { return (ml + i * bw + bw / 2) + ',' + y(d[s.key] || 0); }).join(' ');
      svgEl('polyline', { points: pts, fill: 'none', style: 'stroke:' + s.color, 'stroke-width': 2, 'stroke-linejoin': 'round' }, svg);
      data.forEach(function (d, i) { if (d[s.key]) svgEl('circle', { cx: ml + i * bw + bw / 2, cy: y(d[s.key]), r: 2.6, style: 'fill:' + s.color }, svg); });
    });
    var tip = makeTip(host), sel = null;
    var hl = svgEl('rect', { y: mt, height: ih, width: bw, fill: 'currentColor', opacity: 0, 'pointer-events': 'none' }, svg);
    var hit = svgEl('rect', { x: ml, y: mt, width: iw, height: ih, fill: 'transparent' }, svg);
    function at(ev) {
      var r = svg.getBoundingClientRect(), sx = (ev.clientX - r.left) * (W / r.width);
      var i = Math.floor((sx - ml) / bw); if (i < 0 || i >= data.length) return;
      sel = i; hl.setAttribute('x', ml + i * bw); hl.setAttribute('opacity', 0.06);
      var d = data[i];
      tip.show((ml + i * bw + bw / 2) * (r.width / W), y(o.bars.reduce(function (t, s) { return t + (d[s.key] || 0); }, 0)) * (r.height / H), o.tipTitle(d),
        o.bars.concat(o.lines || []).map(function (s) { return s.name + ' : ' + nf.format(d[s.key] || 0); }));
    }
    hit.addEventListener('pointerdown', at); hit.addEventListener('pointermove', at);
    hit.addEventListener('pointerleave', function (ev) { if (ev.pointerType === 'mouse') { tip.hide(); hl.setAttribute('opacity', 0); } });
  }

  /* Barres horizontales groupées (cas et décès par lieu) */
  function hbars(host, rows) {
    host.replaceChildren();
    if (!rows.length) { host.textContent = 'Aucune donnée.'; return; }
    var W = Math.max(280, host.clientWidth), rowH = 30, H = rows.length * rowH + 6;
    var svg = svgEl('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Cas et décès par lieu' }, host);
    var lw = Math.max.apply(null, rows.map(function (r) { return textWidth(svg, r.lieu); })) + 10;
    lw = Math.min(lw, W * 0.42);
    var max = Math.max.apply(null, rows.map(function (r) { return r.cas; }).concat([1]));
    var vw = textWidth(svg, nf.format(max)) + 8, iw = W - lw - vw;
    rows.forEach(function (r, i) {
      var y = i * rowH + 3;
      var t = svgEl('text', { x: lw - 8, y: y + rowH / 2 - 2, 'text-anchor': 'end', dy: '0.32em' }, svg); t.textContent = r.lieu;
      var wc = (r.cas / max) * iw, wd = (r.dec / max) * iw;
      svgEl('rect', { x: lw, y: y + 3, width: Math.max(r.cas ? 2 : 0, wc), height: 11, rx: 2, style: 'fill:var(--conf)' }, svg);
      svgEl('rect', { x: lw, y: y + 15, width: Math.max(r.dec ? 2 : 0, wd), height: 8, rx: 2, style: 'fill:var(--dec)' }, svg);
      var v = svgEl('text', { x: lw + wc + 5, y: y + 9, dy: '0.32em', class: 'vlab' }, svg); v.textContent = nf.format(r.cas);
      if (r.dec) { var v2 = svgEl('text', { x: lw + wd + 5, y: y + 19.5, dy: '0.32em' }, svg); v2.textContent = nf.format(r.dec); }
    });
  }

  /* Pyramide des âges */
  function pyramid(host, age) {
    host.replaceChildren();
    var W = Math.max(280, host.clientWidth), rowH = 28, H = age.length * rowH + 22;
    var svg = svgEl('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': 'Pyramide des cas par âge et sexe' }, host);
    var max = Math.max.apply(null, age.map(function (a) { return Math.max(a.h, a.f); }).concat([1]));
    var mid = W / 2, lab = 46, vw = textWidth(svg, nf.format(max)) + 8, half = mid - lab / 2 - vw;
    age.slice().reverse().forEach(function (a, i) {
      var y = i * rowH + 4, wh = (a.h / max) * half, wf = (a.f / max) * half;
      svgEl('rect', { x: mid - lab / 2 - wh, y: y, width: wh, height: rowH - 8, rx: 2, style: 'fill:var(--men)' }, svg);
      svgEl('rect', { x: mid + lab / 2, y: y, width: wf, height: rowH - 8, rx: 2, style: 'fill:var(--women)' }, svg);
      var t = svgEl('text', { x: mid, y: y + (rowH - 8) / 2, 'text-anchor': 'middle', dy: '0.32em' }, svg); t.textContent = a.tranche;
      var th = svgEl('text', { x: mid - lab / 2 - wh - 5, y: y + (rowH - 8) / 2, 'text-anchor': 'end', dy: '0.32em', class: 'vlab' }, svg); th.textContent = nf.format(a.h);
      var tf = svgEl('text', { x: mid + lab / 2 + wf + 5, y: y + (rowH - 8) / 2, dy: '0.32em', class: 'vlab' }, svg); tf.textContent = nf.format(a.f);
    });
    var sh = svgEl('text', { x: mid - lab / 2, y: H - 4, 'text-anchor': 'end' }, svg);
    sh.textContent = 'Hommes ' + nf.format(age.reduce(function (t, a) { return t + a.h; }, 0));
    var sf = svgEl('text', { x: mid + lab / 2, y: H - 4 }, svg);
    sf.textContent = 'Femmes ' + nf.format(age.reduce(function (t, a) { return t + a.f; }, 0));
  }

  function renderCourbes() {
    if (state.tab !== 'courbes') return;
    var r = state.result;
    columns($('#ch-courbe'), {
      data: r.courbe, aria: 'Courbe épidémique par semaine de début des symptômes',
      bars: [{ key: 'conf', name: 'Confirmés', color: 'var(--conf)' }, { key: 'prob', name: 'Probables', color: 'var(--prob)' }],
      xLabel: function (d) { return fmtDate(d.semaine, true); },
      tipTitle: function (d) { return 'Semaine du ' + fmtDate(d.semaine); }
    });
    columns($('#ch-notif'), {
      data: r.parNotif, aria: 'Cas par semaine de notification',
      bars: [{ key: 'conf', name: 'Confirmés', color: 'var(--conf)' }, { key: 'prob', name: 'Probables', color: 'var(--prob)' }],
      xLabel: function (d) { return fmtDate(d.semaine, true); },
      tipTitle: function (d) { return 'Semaine du ' + fmtDate(d.semaine); }
    });
    $('#sub-jours').textContent = 'Par jour de notification, du ' + fmtDate(state.filters.ref - 29) + ' au ' + fmtDate(state.filters.ref) + ' (date de référence).';
    columns($('#ch-jours'), {
      data: r.jours, aria: 'Alertes, cas confirmés et décès des 30 derniers jours',
      bars: [{ key: 'alertes', name: 'Alertes', color: 'var(--alert)' }],
      lines: [{ key: 'conf', name: 'Cas confirmés', color: 'var(--conf)' }, { key: 'dec', name: 'Décès (cas)', color: 'var(--dec)' }],
      xLabel: function (d) { return fmtDate(d.jour, true); },
      tipTitle: function (d) { return fmtDate(d.jour); }
    });
    pyramid($('#ch-age'), r.age);
    var note = $('#age-note');
    note.hidden = !r.ageInconnu;
    note.textContent = r.ageInconnu ? nf.format(r.ageInconnu) + ' cas sans âge ou sans sexe ne figurent pas dans la pyramide.' : '';
  }

  function td(txt, cls) { var e = document.createElement('td'); e.textContent = txt; if (cls) e.className = cls; return e; }
  function pillTd(j) {
    var e = document.createElement('td'); e.className = 'n';
    if (j == null) { e.textContent = '–'; return e; }
    var s = document.createElement('span'); s.className = 'pill ' + joursClass(j); s.textContent = nf.format(j); e.appendChild(s); return e;
  }

  /* ---------- Carte des aires de santé et aires actives (carte.js) ---------- */
  var CARTE_URL = 'carte.json';
  function chargerCarte() {
    var local = store.get('carte').then(function (c) { if (c && c.aires && !state.carte) { state.carte = c; if (state.result) renderCarteTab(); } })
      .catch(function () { /* première ouverture */ });
    if (location.protocol === 'file:') return local;
    return local.then(function () {
      return fetch(CARTE_URL, { cache: 'no-cache' }).then(function (res) { return res.ok ? res.json() : null; }, function () { return null; });
    }).then(function (c) {
      if (!c || !c.aires || (state.carte && state.carte.genere === c.genere)) return;
      state.carte = c; store.set('carte', c).catch(function () { /* stockage plein */ });
      if (state.result) renderCarteTab();
    });
  }
  function indCarte() { var v = lsGet('mve-carte-ind'); return MVECarte.INDICATEURS[v] ? v : 'conf'; }
  function renderCarte() {
    var f = state.filters, d = state.data, C = state.carte, ind = indCarte();
    var A = MVECarte.stats(d.rows, f, C), S = MVECarte.synthese(A, f.zone);
    if (!state.carteSel && f.aire) state.carteSel = MVECarte.cle(f.zone, f.aire);
    if (state.carteSel && (!A[state.carteSel] || (f.zone && MVECarte.norm(A[state.carteSel].zone) !== MVECarte.norm(f.zone)))) state.carteSel = null;
    $$('#seg-carte button').forEach(function (b) { b.setAttribute('aria-pressed', b.dataset.ind === ind ? 'true' : 'false'); });
    $('#carte-titre').textContent = f.zone ? 'Carte des aires de santé – ' + f.zone : 'Carte des aires de santé';
    var per = f.periode === 'tout' ? 'toute l’épidémie' : fmtDate(f.debut) + ' → ' + fmtDate(f.fin);
    $('#carte-sub').textContent = MVECarte.INDICATEURS[ind].aide + (ind === 'statut' ? ' Référence : ' + fmtDate(f.ref) + '.' : ' Période : ' + per + '.');
    var host = $('#ch-carte'), note = [];
    if (!C) {
      host.replaceChildren(); var p = document.createElement('p'); p.className = 'chart-note';
      p.textContent = 'Le fond de carte n’est pas encore sur ce téléphone : il arrive avec la prochaine publication (connexion Internet nécessaire une fois).';
      host.appendChild(p); $('#leg-carte').replaceChildren();
    } else {
      MVECarte.dessiner(host, C, A, {
        indicateur: ind, zone: f.zone, selection: state.carteSel,
        hauteurMax: Math.max(320, Math.min(620, window.innerHeight * 0.62)),
        onSelect: function (k) { state.carteSel = state.carteSel === k ? null : k; renderCarte(); }
      });
      MVECarte.legende($('#leg-carte'), ind);
      var hors = S.touchees.filter(function (o) { return !o.contour; });
      if (hors.length) note.push('Sans contour sur la carte (comptées dans les tableaux) : ' + hors.map(function (o) { return o.aire + ' (' + o.zone + ')'; }).join(', ') + '.');
      if (ind === 'taux' && C.popAnnee) note.push('Population ' + C.popAnnee + ' : ' + (C.popSource || 'DHIS2 « Population de l’Aire de Santé »') + '.');
    }
    $('#note-carte').textContent = note.join(' ');
    $('#note-carte').hidden = !note.length;
    // Détail de l'aire touchée
    var info = $('#info-carte'), o = state.carteSel ? A[state.carteSel] : null;
    info.replaceChildren();
    if (!o) {
      info.textContent = C ? 'Touchez une aire pour voir son détail.' + (f.zone ? '' : ' Choisissez une zone dans les filtres (ou touchez-la dans le tableau plus bas) pour l’agrandir.') : '';
      info.hidden = !C;
    } else {
      info.hidden = false;
      var t = document.createElement('b'); t.className = 't'; t.textContent = o.aire + ' · zone de ' + o.zone; info.appendChild(t);
      var kv = document.createElement('div'); kv.className = 'kv';
      [['Cas (période)', nf.format(o.cas) + ' (' + nf.format(o.conf) + ' conf., ' + nf.format(o.prob) + ' prob.)'],
        ['Décès (période)', nf.format(o.dec)],
        ['Population', o.pop ? nf.format(o.pop) : 'inconnue'],
        ['Taux d’attaque', o.taux == null ? '–' : df1.format(o.taux) + ' / 100 000'],
        ['Dernier confirmé', o.dernierConf == null ? 'aucun' : fmtDate(o.dernierConf) + ' (' + nf.format(o.jours) + ' j)'],
        ['Statut', MVECarte.STATUTS[o.statut].lib]].forEach(function (x) {
        var dv = document.createElement('div'), sp = document.createElement('span'); sp.textContent = x[0] + ' : ';
        dv.append(sp, x[1]); kv.appendChild(dv);
      });
      info.appendChild(kv);
      var btn = document.createElement('button'); btn.type = 'button';
      if (!f.zone) {
        btn.textContent = 'Agrandir la zone de ' + o.zone;
        btn.addEventListener('click', function () { filtrerZone(o.zone); });
      } else {
        btn.textContent = 'Voir toute la province';
        btn.addEventListener('click', function () { filtrerZone(''); });
      }
      info.appendChild(btn);
    }
    // Aires actives
    var tiles = [
      ['active', 'Aires actives', S.n.active, 'Dernier cas confirmé il y a moins de 21 jours.'],
      ['surveillance', 'Aires sous surveillance', S.n.surveillance, 'Dernier cas confirmé il y a 21 à 41 jours.'],
      ['calme', 'Aires calmes (≥ 42 j)', S.n.calme, 'Aires touchées, sans cas confirmé depuis au moins 42 jours.'],
      ['jamais', 'Aires sans cas confirmé', S.n.jamais, 'Aucun cas confirmé depuis le début de l’épidémie.']
    ];
    var th = $('#tiles-aires');
    th.replaceChildren.apply(th, tiles.map(function (x) {
      var b = document.createElement('div'); b.className = 'tile flat st-' + x[0];
      var tag = document.createElement('span'); tag.className = 'tag';
      var lab = document.createElement('span'); lab.className = 'lab'; lab.textContent = x[1];
      var val = document.createElement('span'); val.className = 'val'; val.textContent = nf.format(x[2]);
      var nt = document.createElement('span'); nt.className = 'note';
      nt.textContent = S.n.total ? pf.format(x[2] / S.n.total) + ' des ' + nf.format(S.n.total) + ' aires' : x[3];
      b.title = x[3]; b.append(tag, lab, val, nt); return b;
    }).concat(f.zone ? [] : [(function () {
      var b = document.createElement('div'); b.className = 'tile flat st-active';
      var tag = document.createElement('span'); tag.className = 'tag';
      var lab = document.createElement('span'); lab.className = 'lab'; lab.textContent = 'Zones avec une aire active';
      var val = document.createElement('span'); val.className = 'val'; val.textContent = nf.format(S.zonesActives.length);
      var nt = document.createElement('span'); nt.className = 'note'; nt.textContent = S.zonesActives.join(', ') || 'aucune';
      b.append(tag, lab, val, nt); return b;
    })()]));
    $('#aires-sub').textContent = 'Statut selon les jours écoulés depuis le dernier cas confirmé, au ' + fmtDate(f.ref) + (f.zone ? ' – zone de ' + f.zone : ' – toutes les zones') + '.';
    var tb = $('#tb-aires');
    var recentes = S.touchees.filter(function (o) { return o.statut === 'active' || o.statut === 'surveillance'; });
    var montrees = state.toutesAires ? S.touchees : recentes;
    var trs = montrees.map(function (o) {
      var tr = document.createElement('tr'); if (o.statut === 'active') tr.className = 'active';
      tr.appendChild(td(o.aire, 'first')); tr.appendChild(td(o.zone));
      [o.cas, o.conf, o.dec].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); });
      tr.appendChild(td(o.taux == null ? '–' : df1.format(o.taux), 'n'));
      tr.appendChild(td(fmtDate(o.dernierConf))); tr.appendChild(pillTd(o.jours));
      tr.style.cursor = 'pointer';
      tr.addEventListener('click', function () {
        state.carteSel = o.cle; renderCarte();
        var c = $('#card-carte'); if (c.scrollIntoView) c.scrollIntoView({ behavior: 'smooth', block: 'start' });
      });
      return tr;
    });
    if (!trs.length) {
      var tr = document.createElement('tr');
      var c = td(S.touchees.length ? 'Aucune aire active ni sous surveillance.' : 'Aucune aire avec un cas confirmé.', 'first'); c.colSpan = 8; tr.appendChild(c); trs.push(tr);
    }
    tb.replaceChildren.apply(tb, trs);
    var bt = $('#btn-aires');
    bt.hidden = S.touchees.length === recentes.length;
    bt.textContent = state.toutesAires ? 'Afficher seulement les aires actives et sous surveillance (' + nf.format(recentes.length) + ')'
      : 'Afficher toutes les aires touchées (' + nf.format(S.touchees.length) + ')';
  }
  /* Onglet Carte : carte, aires actives, propagation, foyers */
  function renderCarteTab() {
    if (state.tab !== 'carte' || !state.result) return;
    renderCarte(); renderPropagation(); renderFoyers();
  }

  /* ---------- 3. Propagation semaine par semaine ---------- */
  var prop = { idx: null, timer: null, P: null, cleP: null };
  function stopAnim() {
    if (prop.timer) { clearInterval(prop.timer); prop.timer = null; }
    $('#prop-play-ico').setAttribute('d', 'M7 4.5v15l12-7.5z'); $('#prop-play').setAttribute('aria-label', 'Lancer l’animation');
  }
  function renderPropagation(seulementCarte) {
    var f = state.filters, C = state.carte, d = state.data;
    var k = [d.publie, d.chargeLe, f.ref, f.zone].join('|');
    if (prop.cleP !== k) { prop.P = MVECarte.propagation(d.rows, C, f.ref, f.zone); prop.cleP = k; prop.idx = null; stopAnim(); }
    var S = prop.P.semaines, host = $('#ch-prop');
    $('#prop-titre').textContent = 'Propagation semaine par semaine' + (f.zone ? ' – ' + f.zone : '');
    if (!C || !S.length) {
      host.replaceChildren(); $('#leg-prop').replaceChildren(); $('#ch-prop-serie').replaceChildren();
      $('#prop-txt').textContent = !C ? 'Le fond de carte n’est pas encore sur ce téléphone.' : 'Aucun cas confirmé jusqu’à la date de référence.';
      $('#prop-w').disabled = true; $('#prop-play').disabled = true; $('#prop-lab').textContent = '';
      return;
    }
    $('#prop-w').disabled = false; $('#prop-play').disabled = false;
    if (prop.idx == null || prop.idx >= S.length) prop.idx = S.length - 1;
    var w = S[prop.idx], rg = $('#prop-w');
    rg.max = S.length - 1; rg.value = prop.idx;
    $('#prop-lab').textContent = 'Sem. du ' + fmtDate(w.semaine, true);
    MVECarte.dessiner(host, C, {}, {
      zone: f.zone, ariaLabel: 'Propagation – semaine du ' + fmtDate(w.semaine), hauteurMax: Math.max(300, Math.min(560, window.innerHeight * 0.55)),
      couleur: function (cle) { return MVECarte.ETATS[MVECarte.etatSemaine(prop.P.parAire[cle], w.semaine, w.fin)].coul; }
    });
    var leg = $('#leg-prop'); leg.replaceChildren();
    ['nouvelle', 'semaine', 'active', 'touchee', 'jamais'].forEach(function (e) {
      var sp = document.createElement('span'), i = document.createElement('i'); i.style.background = MVECarte.ETATS[e].coul;
      sp.append(i, MVECarte.ETATS[e].lib); leg.appendChild(sp);
    });
    var t = $('#prop-txt'); t.replaceChildren();
    var parts = [['Semaine du ' + fmtDate(w.semaine) + ' au ' + fmtDate(w.fin) + ' : ', null], [nf.format(w.cas), 1],
      [' cas confirmé' + (w.cas > 1 ? 's' : '') + ' dans ', null], [nf.format(w.airesCas), 1], [' aire' + (w.airesCas > 1 ? 's' : '') + ', dont ', null],
      [nf.format(w.nouvelles), 1], [' touchée' + (w.nouvelles > 1 ? 's' : '') + ' pour la première fois. Fin de semaine : ', null],
      [nf.format(w.actives), 1], [' aire' + (w.actives > 1 ? 's' : '') + ' active' + (w.actives > 1 ? 's' : '') + ', ', null],
      [nf.format(w.touchees), 1], [' touchée' + (w.touchees > 1 ? 's' : '') + ' depuis le début.', null]];
    parts.forEach(function (p) { if (p[1]) { var b = document.createElement('b'); b.textContent = p[0]; t.appendChild(b); } else t.appendChild(document.createTextNode(p[0])); });
    if (seulementCarte) return;
    columns($('#ch-prop-serie'), {
      data: S, height: 170, aria: 'Aires touchées par semaine',
      bars: [{ key: 'nouvelles', name: 'Aires touchées pour la 1re fois', color: 'var(--c4)' }],
      lines: [{ key: 'airesCas', name: 'Aires avec des cas', color: 'var(--ink-2)' }],
      xLabel: function (x) { return fmtDate(x.semaine, true); },
      tipTitle: function (x) { return 'Semaine du ' + fmtDate(x.semaine); }
    });
    // Toucher une barre affiche cette semaine sur la carte
    var hit = $('#ch-prop-serie svg rect[fill="transparent"]');
    if (hit) hit.addEventListener('pointerdown', function (ev) {
      var r = hit.getBoundingClientRect(), i = Math.floor((ev.clientX - r.left) / r.width * S.length);
      if (i >= 0 && i < S.length) { stopAnim(); prop.idx = i; renderPropagation(true); }
    });
  }

  /* ---------- 4. Foyers (balayage spatial de Kulldorff) ---------- */
  var foy = { cle: null, R: null, sel: null };
  function fenFoyers() { var v = lsGet('mve-foyers-fen'); return v === '42' || v === 'periode' ? v : '21'; }
  function renderFoyers() {
    var f = state.filters, C = state.carte, d = state.data, fen = fenFoyers();
    $$('#seg-foyers button').forEach(function (b) { b.setAttribute('aria-pressed', b.dataset.fen === fen ? 'true' : 'false'); });
    var debut = fen === 'periode' ? f.debut : f.ref - (+fen) + 1, fin = fen === 'periode' ? f.fin : f.ref;
    $('#foy-sub').textContent = 'Groupes d’aires voisines où les cas confirmés sont nettement plus nombreux qu’attendu au vu de leur population. Du ' +
      fmtDate(debut) + ' au ' + fmtDate(fin) + ', toute la province.';
    var host = $('#ch-foyers'), liste = $('#liste-foyers');
    if (!C) { host.replaceChildren(); liste.replaceChildren(); $('#leg-foyers').replaceChildren(); $('#foy-note').textContent = 'Le fond de carte n’est pas encore sur ce téléphone.'; return; }
    var k = [d.publie, d.chargeLe, debut, fin, C.genere].join('|');
    if (foy.cle !== k) {
      var A = MVECarte.stats(d.rows, { debut: debut, fin: fin, ref: f.ref }, C), cas = {};
      Object.keys(A).forEach(function (x) { if (A[x].conf) cas[x] = A[x].conf; });
      foy.R = MVECarte.foyers(C, cas, {}); foy.cle = k; foy.sel = null; foy.A = A;
    }
    var R = foy.R, groupes = R.foyers.map(function (F, i) {
      return { cles: F.cles, classe: i === 0 ? 'foyer' : 'foyer2', etiquette: String(i + 1) };
    });
    var dansFoyer = {}; R.foyers.forEach(function (F, i) { Object.keys(F.cles).forEach(function (c) { dansFoyer[c] = i; }); });
    MVECarte.dessiner(host, C, foy.A, {
      zone: f.zone, groupes: groupes, ariaLabel: 'Foyers de cas', hauteurMax: Math.max(300, Math.min(560, window.innerHeight * 0.55)),
      selection: foy.sel != null && R.foyers[foy.sel] ? R.foyers[foy.sel].centre.cle : null,
      couleur: function (cle, st) {
        if (dansFoyer[cle] != null) return st.conf ? 'var(--c4)' : 'var(--c2)';
        return st.conf ? 'var(--c1)' : 'var(--c0)';
      }
    });
    var leg = $('#leg-foyers'); leg.replaceChildren();
    [['var(--c4)', 'Aire d’un foyer, avec cas'], ['var(--c2)', 'Aire d’un foyer, sans cas'], ['var(--c1)', 'Cas hors foyer'], ['var(--c0)', 'Aucun cas']].forEach(function (x) {
      var sp = document.createElement('span'), i = document.createElement('i'); i.style.background = x[0]; sp.append(i, x[1]); leg.appendChild(sp);
    });
    liste.replaceChildren();
    if (R.N < 3) {
      var p0 = document.createElement('p'); p0.className = 'chart-note';
      p0.textContent = R.N ? 'Trop peu de cas confirmés (' + R.N + ') sur cette période pour rechercher des foyers.' : 'Aucun cas confirmé sur cette période.';
      liste.appendChild(p0);
    } else if (!R.foyers.length) {
      var p1 = document.createElement('p'); p1.className = 'carte-info';
      p1.textContent = 'Aucun foyer significatif : les ' + nf.format(R.N) + ' cas confirmés de la période ne sont pas plus regroupés que ne le voudrait le hasard, compte tenu de la population des aires.';
      liste.appendChild(p1);
    }
    R.foyers.forEach(function (F, i) {
      var c = document.createElement('div'); c.className = 'foyer-c' + (foy.sel === i ? ' sel' : '');
      var zs = {}; F.aires.forEach(function (a) { zs[a.zone] = 1; });
      var t = document.createElement('b'); t.className = 't'; t.textContent = 'Foyer ' + (i + 1) + ' · ' + Object.keys(zs).join(', ');
      var l1 = document.createElement('div');
      l1.textContent = nf.format(F.cas) + ' cas confirmés pour ' + df1.format(F.attendus) + ' attendus : risque ' + df1.format(F.rr) +
        ' fois plus élevé que dans le reste de la province (p ' + (F.p <= 0.001 ? '≤ 0,001' : '= ' + String(Math.round(F.p * 1000) / 1000).replace('.', ',')) + ').';
      var l2 = document.createElement('div'); l2.className = 'm';
      l2.textContent = F.aires.length + ' aire' + (F.aires.length > 1 ? 's' : '') + ' (' + nf.format(F.pop) + ' hab.) : ' +
        F.aires.map(function (a) { return a.aire + (a.cas ? ' (' + a.cas + ')' : ''); }).join(', ') + '.';
      c.append(t, l1, l2);
      c.addEventListener('click', function () { foy.sel = foy.sel === i ? null : i; renderFoyers(); });
      liste.appendChild(c);
    });
    var note = 'Méthode : balayage spatial de Kulldorff (modèle de Poisson, comme le logiciel SaTScan). Chaque aire et ses voisines les plus proches (jusqu’à 15 aires et 20 % de la population) sont comparées au reste de la province ; ' +
      nf.format(R.simulations) + ' simulations de Monte-Carlo donnent la valeur p. Seuls les foyers avec p < 0,05 et au moins 3 cas sont montrés. Population ' + (C.popAnnee || '') + '.';
    if (R.hors) note += ' ' + nf.format(R.hors) + ' cas d’unités sans contour ou sans population (ex. CTE) ne sont pas pris en compte.';
    $('#foy-note').textContent = note;
  }

  function filtrerZone(z) {
    var saved = lsGet('mve-filtres') || {}; saved.zone = z; saved.aire = ''; lsSet('mve-filtres', saved);
    state.filters = resolveFilters(); render(); toast(z ? 'Filtre : zone de ' + z + '.' : 'Filtre : toutes les zones.');
  }

  function renderZones() {
    if (state.tab !== 'zones') return;
    var r = state.result, f = state.filters;
    $('#lieu-title').textContent = f.zone ? 'Cas et décès par aire de santé – ' + f.zone : 'Cas et décès par zone de santé';
    hbars($('#ch-lieu'), r.parLieu);
    var tb = $('#tb-zones');
    var trs = r.zones.map(function (z) {
      var tr = document.createElement('tr'); if (z.cas7 > 0) tr.className = 'active';
      var first = td('', 'first'); var sw = document.createElement('span'); sw.className = 'sw cls' + MVE.classeConf(z.conf);
      first.append(sw, z.zone); tr.appendChild(first);
      [z.cas, z.conf, z.dec].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); });
      tr.appendChild(td(pf.format(z.letalite), 'n'));
      [z.cas7, z.casJour, z.alertes7, z.attente].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); });
      tr.appendChild(td(fmtDate(z.dernierConf)));
      tr.appendChild(pillTd(z.joursDepuis));
      tr.style.cursor = 'pointer';
      tr.addEventListener('click', function () { filtrerZone(z.zone); });
      return tr;
    });
    var t = r.totalZones, tot = document.createElement('tr'); tot.className = 'total';
    tot.appendChild(td('Total', 'first'));
    [t.cas, t.conf, t.dec].forEach(function (v) { tot.appendChild(td(nf.format(v), 'n')); });
    tot.appendChild(td(pf.format(t.letalite), 'n'));
    [t.cas7, t.casJour, t.alertes7, t.attente].forEach(function (v) { tot.appendChild(td(nf.format(v), 'n')); });
    tot.appendChild(td('')); tot.appendChild(td(''));
    trs.push(tot);
    tb.replaceChildren.apply(tb, trs);
    // Classes de cas confirmés
    var labels = ['0 cas', '1 – 10', '11 – 50', '51 – 100', '> 100'];
    var host = $('#classes');
    host.replaceChildren.apply(host, r.zones.slice().sort(function (a, b) { return b.conf - a.conf; }).map(function (z) {
      var d = document.createElement('div'); var sw = document.createElement('span'); sw.className = 'sw cls' + MVE.classeConf(z.conf);
      d.append(sw, z.zone + ' · ' + nf.format(z.conf)); return d;
    }).concat(labels.map(function (l, i) {
      var d = document.createElement('div'); d.style.color = 'var(--muted)';
      var n = r.zones.filter(function (z) { return MVE.classeConf(z.conf) === i; }).length;
      var sw = document.createElement('span'); sw.className = 'sw cls' + i; d.append(sw, l + ' : ' + n + ' ZS'); return d;
    })));
  }

  function renderProfessions() {
    if (state.tab !== 'professions') return;
    var r = state.result, tb = $('#tb-prof');
    var trs = r.professions.map(function (p) {
      var tr = document.createElement('tr');
      tr.appendChild(td(p.profession, 'first'));
      [p.alertes, p.cas, p.conf, p.dec].forEach(function (v) { tr.appendChild(td(nf.format(v), 'n')); });
      tr.appendChild(td(pf.format(p.letalite), 'n')); tr.appendChild(td(pf.format(p.pCas), 'n'));
      tr.appendChild(td(nf.format(p.cas7), 'n')); tr.appendChild(td(nf.format(p.alertes7), 'n'));
      tr.appendChild(td(fmtDate(p.dernierConf))); tr.appendChild(pillTd(p.joursDepuis));
      return tr;
    });
    var sumK = function (k) { return r.professions.reduce(function (t, p) { return t + p[k]; }, 0); };
    var tot = document.createElement('tr'); tot.className = 'total'; tot.appendChild(td('Total', 'first'));
    ['alertes', 'cas', 'conf', 'dec'].forEach(function (k) { tot.appendChild(td(nf.format(sumK(k)), 'n')); });
    tot.appendChild(td(pf.format(sumK('cas') ? sumK('dec') / sumK('cas') : 0), 'n')); tot.appendChild(td(pf.format(sumK('cas') ? 1 : 0), 'n'));
    tot.appendChild(td(nf.format(sumK('cas7')), 'n')); tot.appendChild(td(nf.format(sumK('alertes7')), 'n'));
    tot.appendChild(td('')); tot.appendChild(td(''));
    trs.push(tot);
    tb.replaceChildren.apply(tb, trs);
  }

  /* ---------- Onglets et feuilles ---------- */
  function showTab(name) {
    state.tab = name; lsSet('mve-onglet', name);
    if (name !== 'carte' && prop.timer) stopAnim();
    ['synthese', 'courbes', 'carte', 'zones', 'professions'].forEach(function (t) { $('#tab-' + t).hidden = t !== name; });
    $$('#tabs button').forEach(function (b) { b.setAttribute('aria-selected', b.dataset.tab === name ? 'true' : 'false'); });
    window.scrollTo(0, 0);
    if (state.result) { renderCourbes(); renderCarteTab(); renderZones(); renderProfessions(); }
  }
  function openSheet(id) {
    closeSheets(true); $('#toast').hidden = true;
    $('#sheet-bg').hidden = false; $(id).hidden = false;
    history.pushState({ sheet: id }, '');
  }
  function closeSheets(silent) {
    var open = $$('.sheet').some(function (s) { return !s.hidden; });
    $$('.sheet').forEach(function (s) { s.hidden = true; });
    $('#sheet-bg').hidden = true;
    if (open && !silent && history.state && history.state.sheet) history.back();
  }
  window.addEventListener('popstate', function () { $$('.sheet').forEach(function (s) { s.hidden = true; }); $('#sheet-bg').hidden = true; });

  /* ---------- Événements ---------- */
  $$('#tabs button').forEach(function (b) { b.addEventListener('click', function () { showTab(b.dataset.tab); }); });
  $('#prop-w').addEventListener('input', function () { stopAnim(); prop.idx = +$('#prop-w').value; renderPropagation(true); });
  $('#prop-play').addEventListener('click', function () {
    if (prop.timer) { stopAnim(); return; }
    var S = prop.P && prop.P.semaines; if (!S || !S.length) return;
    if (prop.idx == null || prop.idx >= S.length - 1) prop.idx = 0; else prop.idx++;
    renderPropagation(true);
    $('#prop-play-ico').setAttribute('d', 'M7 5h3.5v14H7zM13.5 5H17v14h-3.5z'); $('#prop-play').setAttribute('aria-label', 'Arrêter l’animation');
    prop.timer = setInterval(function () {
      if (state.tab !== 'carte' || prop.idx >= S.length - 1) { stopAnim(); return; }
      prop.idx++; renderPropagation(true);
    }, 900);
  });
  $$('#seg-foyers button').forEach(function (b) {
    b.addEventListener('click', function () { lsSet('mve-foyers-fen', b.dataset.fen); if (state.result) renderFoyers(); });
  });
  $('#btn-aires').addEventListener('click', function () { state.toutesAires = !state.toutesAires; if (state.result) renderCarte(); });
  $$('#seg-carte button').forEach(function (b) {
    b.addEventListener('click', function () { lsSet('mve-carte-ind', b.dataset.ind); if (state.result) renderCarte(); });
  });
  $('#sheet-bg').addEventListener('click', function () { closeSheets(); });
  $$('[data-close]').forEach(function (b) { b.addEventListener('click', function () { closeSheets(); }); });
  $('#btn-open-empty').addEventListener('click', function () { $('#file').click(); });
  $('#btn-open').addEventListener('click', function () { closeSheets(); $('#file').click(); });
  $('#file').addEventListener('change', function (e) { openFile(e.target.files[0]); e.target.value = ''; });
  $('#btn-menu').addEventListener('click', function () {
    var d = state.data;
    var last = lsGet('mve-sync-last');
    $('#menu-info').textContent = (d ? 'Données : ' + d.fichier + ', reçues le ' + new Date(d.chargeLe).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) + '. ' : 'Aucune donnée chargée. ') +
      (sync.key ? 'Mise à jour automatique activée' + (last ? ', dernière vérification le ' + new Date(last).toLocaleString('fr-FR', { dateStyle: 'short', timeStyle: 'short' }) : '') + '.' : 'Mise à jour automatique désactivée.');
    setSyncUI();
    $('#btn-clear').hidden = !d;
    openSheet('#sheet-menu');
  });
  $('#btn-help').addEventListener('click', function () { openSheet('#sheet-help'); });
  $('#btn-sync').addEventListener('click', function () { syncNow(true); });
  $('#btn-sync-menu').addEventListener('click', function () { closeSheets(); syncNow(true); });
  $('#btn-auto').addEventListener('click', openPassword);
  $('#btn-auto-empty').addEventListener('click', openPassword);
  $('#btn-auto-off').addEventListener('click', function () {
    sync.key = null; store.del('cle'); setSyncUI(); closeSheets(); toast('Mise à jour automatique désactivée sur ce téléphone.');
  });
  $('#form-pw').addEventListener('submit', function (e) {
    e.preventDefault();
    var pw = $('#pw').value;
    if (!pw) { $('#pw-msg').textContent = 'Saisissez le mot de passe d’équipe.'; return; }
    connect(pw);
  });
  $('#pw-show').addEventListener('change', function () { $('#pw').type = $('#pw-show').checked ? 'text' : 'password'; });
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'visible' && Date.now() - (lsGet('mve-sync-last') || 0) > 10 * 60000) syncNow(false);
  });
  window.addEventListener('online', function () { syncNow(false); });
  $('#btn-clear').addEventListener('click', function () {
    var b = $('#btn-clear');
    if (b.dataset.confirm !== '1') { b.dataset.confirm = '1'; b.textContent = 'Confirmer : effacer la liste linéaire du téléphone'; return; }
    b.dataset.confirm = ''; b.textContent = 'Effacer les données de ce téléphone';
    store.del('dataset').then(function () {
      state.data = null; state.result = null; closeSheets();
      $$('.tab').forEach(function (t) { t.hidden = true; });
      $('#empty').hidden = false; $('#tabs').hidden = true; $('#chips').hidden = true; $('#src').textContent = 'Aucune donnée chargée';
      toast('Données effacées de ce téléphone.');
    });
  });
  $('#f-zone').addEventListener('change', function () { fillAireSelect($('#f-aire'), $('#f-zone').value, ''); });
  $('#q-periode').addEventListener('click', function (e) {
    var p = e.target.dataset && e.target.dataset.p; if (!p) return;
    draft.periode = p; draft._periodeTouched = true;
    var rows = state.data.rows, r = inputToS($('#f-ref').value) || state.filters.ref;
    if (p === 'tout') { $('#f-debut').value = sToInput(minNotif(rows)); $('#f-fin').value = sToInput(Math.max(today(), maxNotif(rows))); }
    else { $('#f-fin').value = sToInput(r); $('#f-debut').value = sToInput(r - (+p) + 1); }
  });
  $('#q-ref').addEventListener('click', function (e) {
    var m = e.target.dataset && e.target.dataset.r; if (!m) return;
    draft.refMode = m; draft._refTouched = true;
    $('#f-ref').value = sToInput(m === 'data' ? maxNotif(state.data.rows) : today() - (+m));
  });
  $('#f-debut').addEventListener('input', function () { draft.periode = 'fixe'; draft._periodeTouched = true; });
  $('#f-fin').addEventListener('input', function () { draft.periode = 'fixe'; draft._periodeTouched = true; });
  $('#f-ref').addEventListener('input', function () { draft.refMode = 'fixe'; draft._refTouched = true; });
  $('#btn-apply').addEventListener('click', applyFilters);
  $('#btn-reset').addEventListener('click', function () {
    lsSet('mve-filtres', {}); closeSheets(); state.filters = resolveFilters(); render(); toast('Filtres réinitialisés.');
  });
  var resizeT;
  window.addEventListener('resize', function () { clearTimeout(resizeT); resizeT = setTimeout(function () { if (state.result) { renderCourbes(); renderCarteTab(); renderZones(); } }, 150); });

  /* Installation (Chrome Android) */
  var installEvt = null;
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); installEvt = e; $('#btn-install').hidden = false; });
  $('#btn-install').addEventListener('click', function () {
    if (!installEvt) return; closeSheets(); installEvt.prompt();
    installEvt.userChoice.finally(function () { installEvt = null; $('#btn-install').hidden = true; });
  });
  window.addEventListener('appinstalled', function () { toast('Application installée. Elle est maintenant sur l’écran d’accueil.'); });

  /* ---------- Démarrage ---------- */
  function boot() {
    var params = new URLSearchParams(location.search);
    var shared = params.get('partage') === '1' && 'caches' in window
      ? caches.open('partage').then(function (c) {
        return c.match('fichier-partage').then(function (res) {
          if (!res) return false;
          var name = decodeURIComponent(res.headers.get('X-Nom') || 'fichier.xlsx');
          return res.arrayBuffer().then(function (b) {
            c.delete('fichier-partage'); history.replaceState(null, '', './');
            return readWorkbook(new Uint8Array(b), name).then(function () { return true; });
          });
        });
      }) : Promise.resolve(false);
    var cle = store.get('cle').then(function (c) {
      if (c && c.key) { sync.key = c.key; sync.salt = c.salt; sync.iter = c.iter; }
    }).catch(function () { /* pas de clé enregistrée */ });
    Promise.all([shared, cle]).then(function (res) {
      setSyncUI();
      if (res[0]) return;
      return store.get('dataset').then(function (d) { if (d && d.rows) setData(d, false); }).catch(function () { /* première ouverture */ });
    }).then(function () { chargerCarte(); syncNow(false); });
    if ('launchQueue' in window) {
      window.launchQueue.setConsumer(function (p) { if (p.files && p.files[0]) p.files[0].getFile().then(openFile); });
    }
    if ('serviceWorker' in navigator && location.protocol !== 'file:') {
      navigator.serviceWorker.register('sw.js').catch(function () { /* hors ligne indisponible */ });
    }
  }
  boot();
})();
