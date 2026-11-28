/* Moteur de calcul « Qualité Contacts MVE ».
   Reproduit les feuilles Contrôles, ContrôlesSuivis, Calculs et Tableau de bord du classeur
   « Tableau_de_bord_Qualite_Donnees_Contacts_MVE_DHIS2 ». Mêmes données que « Contacts MVE »
   (Données A à V, Suivis A à J). Dates = numéros de série Excel. Aucun nom n'est lu. */
(function (root) {
  'use strict';

  var ZONES = ['Aba', 'Boma Mangbetu', 'Doruma', 'Dungu', 'Faradje', 'Gombari', 'Isiro',
    'Makoro', 'Niangara', 'Pawa', 'Rungu', 'Wamba', 'Watsa'];
  var EPOCH = Date.UTC(1899, 11, 30);
  // Paramètres!D2:D6
  var PARAMS_DEFAUT = { DelaiList: 21, AgeMax: 110, RetSaisie: 1, Vert: 0.9, Orange: 0.8 };
  var COLS = {
    enr: 'Enrollment', id: 'ID contact', epid: 'N° Epid cas source', dEnr: 'Date enregistrement', dDern: 'Date dernier contact',
    zone: 'Zone de santé', aire: 'Aire de santé', sexe: 'Sexe', age: 'Âge (ans)', ageM: 'Âge (mois)', rel: 'Relation avec le cas',
    circ: 'Circonstance du contact', type: 'Type de contact', notif: 'Notifié par', deb: 'Début suivi', fin: 'Fin suivi (J21)',
    issue: 'Issue', dIss: 'Date issue', nbJ: 'Jours de suivi saisis', sympt: 'Symptômes signalés', prem: '1er contact du cas source', extrait: 'Extrait le'
  };
  var SCOLS = {
    enr: 'Enrollment', date: 'Date suivi', vu: 'Vu', raison: 'Raison non vu', sympt: 'Symptômes', res: 'Résultat du suivi',
    issue: 'Issue', dernier: 'Dernier jour de suivi', nb: 'Nb fiches', extrait: 'Extrait le'
  };
  // Contrôles des contacts : [clé, libellé (graphique / détail)]
  var MANQ = [['mID', 'ID contact', 'ID contact'], ['mEpid', 'N° Epid du cas source', 'N° Epid du cas source'], ['mDern', 'Date du dernier contact', 'date du dernier contact'],
    ['mSexe', 'Sexe', 'sexe'], ['mRel', 'Relation avec le cas', 'relation avec le cas'], ['mCirc', 'Circonstance du contact', 'circonstance du contact'],
    ['mType', 'Type de contact', 'type de contact'], ['mNotif', 'Notifié par', 'notifié par'], ['mAge', 'Âge', 'âge']];
  var INCOH = [['iDernApres', 'dernier contact après l\'enregistrement'], ['iListTard', 'listé trop tard après le dernier contact'],
    ['iEpid', 'N° Epid non standard'], ['iAge', 'âge invalide'], ['iDoublon', 'ID contact en double'],
    ['iIssue', 'issue sans date ou hors période de suivi'], ['iSympt', 'symptômes signalés sans issue']];
  // Saisies rétrospectives : une fiche peut être datée avant l'enregistrement du contact dans DHIS2 ; ce n'est pas une anomalie.
  var FICHE = [['aOrph', 'fiche sans contact'], ['aHors', 'hors période J1–J21'],
    ['aApres', 'après l\'issue du contact'], ['aFutur', 'date dans le futur'], ['aVu', '« vu » non renseigné'], ['aRaison', 'non vu sans raison'],
    ['aVuRaison', 'vu avec une raison de non-visite'], ['aSympt', 'symptômes non renseignés (vu)'], ['aDouble', 'plusieurs fiches le même jour']];

  function norm(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase(); }
  function vide(v) { return v == null || String(v).replace(/ +/g, '') === ''; } // TRIM(x&"")=""
  function eqx(v, t) { return v != null && String(v).toLowerCase() === String(t).toLowerCase(); } // égalité Excel (sans casse)
  function eq(v, t) { return !vide(v) && norm(v) === norm(t); }
  function num(v) { return typeof v === 'number' && !isNaN(v) ? v : (typeof v === 'string' && v.trim() !== '' && !isNaN(+v) ? +v : null); }
  function toSerial(v, keepTime) {
    if (v == null || (typeof v === 'string' && v.trim() === '')) return null;
    var s = null, m, t;
    if (typeof v === 'number') s = v;
    else if (v instanceof Date) s = (v.getTime() - v.getTimezoneOffset() * 60000 - EPOCH) / 864e5;
    else if ((m = (t = String(v).trim()).match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/))) s = (Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) - EPOCH) / 864e5;
    else if ((m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) s = (Date.UTC(+m[3], +m[2] - 1, +m[1]) - EPOCH) / 864e5;
    else if (!isNaN(+t)) s = +t;
    if (s == null || isNaN(s)) return null;
    return keepTime ? s : Math.floor(s + 1e-9);
  }
  function serialToUTC(s) { return new Date(EPOCH + s * 864e5); }
  function weekday1(s) { return ((serialToUTC(s).getUTCDay() + 6) % 7) + 1; }
  function nowSerial() { var d = new Date(); return (d.getTime() - d.getTimezoneOffset() * 60000 - EPOCH) / 864e5; }
  function index(table, cols) {
    if (!table || !table.length) return null;
    var h = table[0].map(norm), idx = {};
    Object.keys(cols).forEach(function (k) { idx[k] = h.indexOf(norm(cols[k])); });
    return idx;
  }
  function epurer(table, cols) {
    var idx = index(table, cols); if (!idx) return [];
    var keys = Object.keys(cols);
    return [keys.map(function (k) { return cols[k]; })].concat(table.slice(1).map(function (r) {
      return keys.map(function (k) { var v = idx[k] >= 0 && r ? r[idx[k]] : null; return v === undefined ? null : v; });
    }));
  }
  function epidStandard(e) {
    var s = String(e), t = s.replace(/^ +| +$/g, '').replace(/ {2,}/g, ' '), u = t.toUpperCase();
    return t.length === 26 && u.slice(0, 8) === 'RDC-HUE-' && u.slice(11, 15) === '-26-' && u.slice(-4) === '-MVE' && s === t;
  }

  function normalize(contacts, suivis, params) {
    var P = Object.assign({}, PARAMS_DEFAUT, params || {});
    var ic = index(contacts, COLS), is = index(suivis, SCOLS);
    if (!ic) throw new Error('La feuille « Données » est vide.');
    var miss = ['enr', 'dEnr', 'zone', 'aire', 'deb', 'fin'].filter(function (k) { return ic[k] < 0; }).map(function (k) { return COLS[k]; });
    if (miss.length) throw new Error('Colonnes introuvables : ' + miss.join(', ') + '. Est-ce bien un classeur « Contacts MVE » ?');
    var extrait = null, extraitS = null, rows = [], byEnr = {}, nId = {};
    for (var i = 1; i < contacts.length; i++) {
      var t = contacts[i]; if (!t) continue;
      var r = {};
      Object.keys(COLS).forEach(function (k) { r[k] = ic[k] >= 0 ? t[ic[k]] : null; if (r[k] === undefined) r[k] = null; });
      if (r.enr == null || r.enr === '') continue; // Contrôles!A
      ['dEnr', 'dDern', 'deb', 'fin', 'dIss'].forEach(function (k) { r[k] = toSerial(r[k]); });
      var ex = toSerial(r.extrait, true); if (ex != null && (extrait == null || ex > extrait)) extrait = ex;
      r.age = num(r.age); r.ageM = num(r.ageM); r.nbJ = num(r.nbJ) || 0;
      r.rang = rows.length; r.enr = String(r.enr);
      if (!vide(r.id)) { var key = String(r.id).toLowerCase(); nId[key] = (nId[key] || 0) + 1; }
      rows.push(r);
      if (!(r.enr in byEnr)) byEnr[r.enr] = r;
    }
    if (!rows.length) throw new Error('Aucun contact dans le fichier.');
    rows.forEach(function (r) {
      r.mID = vide(r.id) ? 1 : 0; r.mEpid = vide(r.epid) ? 1 : 0; r.mDern = r.dDern == null ? 1 : 0; r.mSexe = vide(r.sexe) ? 1 : 0;
      r.mRel = vide(r.rel) ? 1 : 0; r.mCirc = vide(r.circ) ? 1 : 0; r.mType = vide(r.type) ? 1 : 0; r.mNotif = vide(r.notif) ? 1 : 0;
      r.mAge = r.age == null && r.ageM == null ? 1 : 0;
      r.iDernApres = r.dDern != null && r.dEnr != null && r.dDern > r.dEnr ? 1 : 0;
      r.iEpid = vide(r.epid) ? 0 : (epidStandard(r.epid) ? 0 : 1);
      r.iAge = (r.age != null && (r.age < 0 || r.age > P.AgeMax)) || (r.ageM != null && (r.ageM < 0 || r.ageM > 11)) ? 1 : 0;
      r.iDoublon = !vide(r.id) && nId[String(r.id).toLowerCase()] > 1 ? 1 : 0;
      r.iIssue = vide(r.issue) ? 0 : (r.dIss == null || r.dIss < (r.deb || 0) ? 1 : 0);
      r.iSympt = eqx(r.sympt, 'Oui') && vide(r.issue) ? 1 : 0;
      r.sem = (r.dEnr || 0) - weekday1(r.dEnr || 0) + 1;
      r.fiches = [];
    });
    // Fiches de suivi (ContrôlesSuivis)
    var fiches = [], today = Math.floor(nowSerial());
    if (is) for (var j = 1; j < suivis.length; j++) {
      var s = suivis[j]; if (!s) continue;
      var g = function (k) { var v = is[k] >= 0 ? s[is[k]] : null; return v === undefined ? null : v; };
      var enr = g('enr'); if (enr == null || enr === '') continue;
      var d = toSerial(g('date')), exs = toSerial(g('extrait'), true);
      if (exs != null && (extraitS == null || exs > extraitS)) extraitS = exs;
      var c = byEnr[String(enr)] || null, vu = g('vu'), raison = g('raison'), sy = g('sympt'), dd = d || 0;
      var f = {
        rang: fiches.length, enr: String(enr), c: c, date: d, vu: vide(vu) ? null : String(vu), zone: c ? (c.zone || '') : '', aire: c ? (c.aire || '') : '',
        aOrph: c ? 0 : 1,
        aHors: c && (dd - (c.deb || 0) < 1 || dd - (c.deb || 0) > 21) ? 1 : 0,
        aApres: c && c.dIss != null && dd > c.dIss ? 1 : 0,
        aFutur: d != null && d > (exs != null ? Math.floor(exs) : today) ? 1 : 0,
        aVu: vide(vu) ? 1 : 0,
        aRaison: eqx(vu, 'Non') && vide(raison) ? 1 : 0,
        aVuRaison: eqx(vu, 'Oui') && !vide(raison) ? 1 : 0,
        aSympt: eqx(vu, 'Oui') && vide(sy) ? 1 : 0,
        aDouble: (num(g('nb')) || 0) > 1 ? 1 : 0
      };
      f.nbAnom = FICHE.reduce(function (t2, m) { return t2 + f[m[0]]; }, 0);
      f.detail = FICHE.filter(function (m) { return f[m[0]]; }).map(function (m) { return m[1]; });
      fiches.push(f);
      if (c && d != null) c.fiches.push(d);
    }
    // Date de listage effective : saisie rétrospective = fiche de suivi antérieure à l'enregistrement DHIS2.
    // Le contact était alors déjà connu et suivi : on retient la plus ancienne fiche plausible (pas avant le dernier contact).
    rows.forEach(function (r) {
      var plancher = r.dDern != null ? r.dDern : (r.deb != null ? r.deb : -Infinity), premiere = null;
      r.fiches.forEach(function (d) { if (d >= plancher && (premiere == null || d < premiere)) premiere = d; });
      r.dListe = r.dEnr != null && premiere != null && premiere < r.dEnr ? premiere : r.dEnr;
      r.retro = r.dEnr != null && r.dListe != null && r.dListe < r.dEnr ? 1 : 0;
      r.iListTard = r.dDern != null && r.dListe != null && r.dListe - r.dDern > P.DelaiList ? 1 : 0;
      r.nbManq = MANQ.reduce(function (s, m) { return s + r[m[0]]; }, 0);
      r.nbIncoh = INCOH.reduce(function (s, m) { return s + r[m[0]]; }, 0);
      r.nbAnom = r.nbManq + r.nbIncoh;
      r.delai = r.dListe != null && r.dDern != null ? r.dListe - r.dDern : null;
      r.detail = MANQ.filter(function (m) { return r[m[0]]; }).map(function (m) { return 'manque ' + m[2]; })
        .concat(INCOH.filter(function (m) { return r[m[0]]; }).map(function (m) { return m[1]; }));
    });
    return { rows: rows, fiches: fiches, extrait: extrait, extraitS: extraitS, charge: rows.length, nbFiches: fiches.length, params: P };
  }

  function zonesOf(rows) {
    var seen = {}, list = ZONES.slice();
    list.forEach(function (z) { seen[norm(z)] = 1; });
    rows.forEach(function (r) { if (r.zone && !seen[norm(r.zone)]) { seen[norm(r.zone)] = 1; list.push(r.zone); } });
    return list;
  }
  function airesOf(rows, zone) {
    var s = {};
    rows.forEach(function (r) { if (!vide(r.aire) && (!zone || eq(r.zone, zone))) s[r.aire] = 1; });
    return Object.keys(s).sort(function (a, b) { return a.localeCompare(b, 'fr'); });
  }
  function cnt(rows, f) { var n = 0; for (var i = 0; i < rows.length; i++) if (f(rows[i])) n++; return n; }
  function sum(rows, f) { var n = 0; for (var i = 0; i < rows.length; i++) { var v = f(rows[i]); if (typeof v === 'number') n += v; } return n; }
  function avg(v) { return v.length ? v.reduce(function (t, x) { return t + x; }, 0) / v.length : null; }
  function ratio(a, b) { return b ? a / b : null; }

  /* Colonnes W à AB de Contrôles à la date de référence */
  function etat(r, ref, P) {
    var st;
    var connu = r.dListe != null ? r.dListe : (r.dEnr || 0); // saisie rétrospective : connu dès la 1re fiche
    if (connu > ref) st = 'Non encore listé';
    else if (!vide(r.issue) && r.dIss != null && r.dIss <= ref) st = r.issue;
    else if (ref > (r.fin || 0)) st = 'Sorti (21 j terminés)';
    else if (ref <= (r.deb || 0)) st = 'En attente (J0)';
    else st = 'En suivi';
    r.statut = st;
    if (connu > ref) { r.jAtt = null; r.jSais = null; r.sansFiche = null; }
    else {
      var fin = r.fin || 0;
      r.jAtt = Math.max(0, Math.min(fin, ref, r.dIss != null ? r.dIss : fin) - (r.deb || 0));
      r.jSais = Math.min(r.nbJ, r.jAtt);
      r.sansFiche = r.jAtt === 0 ? null : (r.nbJ === 0 ? 1 : 0);
    }
    r.retard = null; r.derniere = null;
    if (st === 'En suivi') {
      var m = 0; r.fiches.forEach(function (d) { if (d <= ref && d > m) m = d; });
      r.derniere = m || null; r.retard = !m || ref - m > P.RetSaisie ? 1 : 0;
    }
  }

  /* f = {zone, aire, debut, fin, ref} */
  function compute(D, f) {
    var P = D.params, rows = D.rows, ref = f.ref;
    rows.forEach(function (r) { etat(r, ref, P); });
    var zOk = function (z) { return f.zone ? eqx(z, f.zone) : !vide(z); };         // COUNTIFS : « * » = zone renseignée
    var aOk = function (a) { return !f.aire || eqx(a, f.aire); };
    var inPer = function (r) { return r.dEnr != null && r.dEnr >= f.debut && r.dEnr <= f.fin; };
    var geo = rows.filter(function (r) { return zOk(r.zone) && aOk(r.aire); }), per = geo.filter(inPer);
    var fOk = function (x) { return (!f.zone || eqx(x.zone, f.zone)) && (!f.aire || eqx(x.aire, f.aire)) && x.date != null && x.date >= f.debut && x.date <= f.fin; };
    var fil = D.fiches.filter(fOk);
    var c = cnt, L = per.length;
    var k = {
      listes: L, pSans: ratio(c(per, function (r) { return r.nbAnom === 0; }), L),
      compl: sum(per, function () { return 9; }) ? 1 - sum(per, function (r) { return r.nbManq; }) / (9 * L) : null,
      pLies: ratio(c(per, function (r) { return r.mEpid === 0; }), L),
      incoh: c(per, function (r) { return r.nbIncoh > 0; }), epidNS: c(per, function (r) { return r.iEpid; }),
      fiches: fil.length, pFichesSans: ratio(c(fil, function (x) { return x.nbAnom === 0; }), fil.length),
      complFiches: ratio(sum(per, function (r) { return r.jSais; }), sum(per, function (r) { return r.jAtt; })),
      sansFiche: c(per, function (r) { return r.sansFiche === 1; }), doubles: c(fil, function (x) { return x.aDouble; }), hors: c(fil, function (x) { return x.aHors; }),
      enSuivi: c(geo, function (r) { return r.statut === 'En suivi'; }), retard: c(geo, function (r) { return r.retard === 1; }),
      retro: c(per, function (r) { return r.retro; }),
      delai: avg(per.filter(function (r) { return r.delai != null && r.delai >= 0; }).map(function (r) { return r.delai; }))
    };
    k.pAJour = k.enSuivi ? 1 - k.retard / k.enSuivi : null;
    var lim = D.extraitS != null ? D.extraitS : Math.floor(nowSerial()), dm = 0;
    D.fiches.forEach(function (x) { if (x.date != null && x.date <= lim && x.date > dm) dm = x.date; });
    k.derniereFiche = dm || null;
    k.fraicheur = D.extrait != null ? Math.round((nowSerial() - D.extrait) * 24) : null;

    // Graphiques (Calculs)
    var variables = MANQ.map(function (m) { return { lib: m[1], n: L, manq: c(per, function (r) { return r[m[0]]; }), pct: L ? 1 - c(per, function (r) { return r[m[0]]; }) / L : null }; });
    var types = INCOH.map(function (m) { return { lib: 'Contact – ' + m[1], n: c(per, function (r) { return r[m[0]]; }) }; })
      .concat(FICHE.map(function (m) { return { lib: 'Fiche – ' + m[1], n: c(fil, function (x) { return x[m[0]]; }) }; }));
    var lundi = ref - weekday1(ref) + 1, semaines = [];
    for (var w = 11; w >= 0; w--) {
      var s = lundi - 7 * w, rs = geo.filter(function (r) { return r.sem === s; }), sa = sum(rs, function (r) { return r.jAtt; });
      semaines.push({ semaine: s, listes: rs.length, sans: c(rs, function (r) { return r.nbAnom === 0; }), pSans: rs.length ? c(rs, function (r) { return r.nbAnom === 0; }) / rs.length : null,
        complFiches: sa ? sum(rs, function (r) { return r.jSais; }) / sa : null });
    }

    // Tableau par zone (sans filtre d'aire) et par aire (zone choisie)
    var ligne = function (rs, fs, rsGeo, lieu) {
      var n = rs.length, sa = sum(rs, function (r) { return r.jAtt; });
      return {
        lieu: lieu, listes: n, pSans: n ? c(rs, function (r) { return r.nbAnom === 0; }) / n : null, compl: n ? 1 - sum(rs, function (r) { return r.nbManq; }) / (9 * n) : null,
        incoh: c(rs, function (r) { return r.nbIncoh > 0; }), complFiches: n && sa ? sum(rs, function (r) { return r.jSais; }) / sa : null,
        sansFiche: c(rs, function (r) { return r.sansFiche === 1; }), fiches: fs.length, fichesAnom: c(fs, function (x) { return x.nbAnom > 0; }),
        doubles: c(fs, function (x) { return x.aDouble; }), retard: c(rsGeo, function (r) { return r.retard === 1; })
      };
    };
    var finPer = function (x) { return x.date != null && x.date >= f.debut && x.date <= f.fin; };
    var zones = zonesOf(rows).map(function (z) {
      var zr = rows.filter(function (r) { return eqx(r.zone, z); });
      return ligne(zr.filter(inPer), D.fiches.filter(function (x) { return eqx(x.zone, z) && finPer(x); }), zr, z);
    });
    var tot = function (list) {
      var t = { lieu: 'Total' };
      ['listes', 'incoh', 'sansFiche', 'fiches', 'fichesAnom', 'doubles', 'retard'].forEach(function (x) { t[x] = list.reduce(function (s2, z) { return s2 + z[x]; }, 0); });
      t.pSans = k.pSans; t.compl = k.compl; t.complFiches = k.complFiches; return t;
    };
    var zr2 = rows.filter(function (r) { return zOk(r.zone); });
    var aires = airesOf(rows, f.zone).map(function (a) {
      var ar = zr2.filter(function (r) { return eqx(r.aire, a); });
      return ligne(ar.filter(inPer), D.fiches.filter(function (x) { return (!f.zone || eqx(x.zone, f.zone)) && eqx(x.aire, a) && finPer(x); }), ar, a);
    });
    var totA = (function () {
      var t = { lieu: 'Total' };
      ['listes', 'incoh', 'sansFiche', 'fiches', 'fichesAnom', 'doubles', 'retard'].forEach(function (x) { t[x] = aires.reduce(function (s2, z) { return s2 + z[x]; }, 0); });
      var rs = zr2.filter(function (r) { return !vide(r.aire) && inPer(r); }), sa = sum(rs, function (r) { return r.jAtt; });
      t.pSans = rs.length ? c(rs, function (r) { return r.nbAnom === 0; }) / rs.length : null; t.compl = rs.length ? 1 - sum(rs, function (r) { return r.nbManq; }) / (9 * rs.length) : null;
      t.complFiches = sa ? sum(rs, function (r) { return r.jSais; }) / sa : null; return t;
    })();

    // Contacts à corriger (Contrôles!AE:AG : zone « * » = toutes, y compris sans zone), les plus récents d'abord
    var dansFiltre = function (r) { return (!f.zone || eqx(r.zone, f.zone)) && (!f.aire || eqx(r.aire, f.aire)) && (r.dEnr || 0) >= f.debut && (r.dEnr || 0) <= f.fin; };
    var aCorriger = rows.filter(function (r) { return dansFiltre(r) && r.nbAnom > 0; }).reverse();
    var fichesACorriger = fil.filter(function (x) { return x.nbAnom > 0; }).sort(function (a, b) { return (b.date || 0) - (a.date || 0) || b.rang - a.rang; });

    return { k: k, variables: variables, types: types, semaines: semaines, zones: zones, totalZones: tot(zones), aires: aires, totalAires: totA,
      aCorriger: aCorriger, fichesACorriger: fichesACorriger, P: P };
  }

  var api = { ZONES: ZONES, PARAMS_DEFAUT: PARAMS_DEFAUT, COLS: COLS, SCOLS: SCOLS, normalize: normalize, compute: compute, epurer: epurer,
    zonesOf: zonesOf, airesOf: airesOf, toSerial: toSerial, serialToUTC: serialToUTC, nowSerial: nowSerial, norm: norm };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.MVEQC = api;
})(this);
