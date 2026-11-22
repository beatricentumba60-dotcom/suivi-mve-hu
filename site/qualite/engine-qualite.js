/* Moteur de calcul « Qualité des données MVE ».
   Reproduit les feuilles Contrôles, Calculs et Tableau de bord du classeur
   « Tableau_de_bord_Qualite_Donnees_MVE_DHIS2 ». Dates = numéros de série Excel.
   Les noms de patients ne sont jamais conservés : seule une clé de doublon
   (opaque si les données viennent de l'ordinateur) est utilisée. */
(function (root) {
  'use strict';

  var ZONES = ['Aba', 'Boma Mangbetu', 'Doruma', 'Dungu', 'Faradje', 'Gombari', 'Isiro',
    'Makoro', 'Niangara', 'Pawa', 'Rungu', 'Wamba', 'Watsa'];

  var SEUILS_DEFAUT = { DDSmax: 60, AgeMax: 110, CiblePrelev: 1, RetInv: 2, RetLabo: 3, RetClassif: 7, Vert: 0.9, Orange: 0.8, Fenetre: 7 };

  var COLS = {
    id: 'ID alerte', notif: 'Date notification', dds: 'Date début symptômes', zone: 'Zone de santé', aire: 'Aire de santé',
    sexe: 'Sexe', ageAns: 'Âge (ans)', ageMois: 'Âge (mois)', nature: 'Nature alerte', conclusion: 'Conclusion alerte',
    classif: 'Classification finale', statut: 'Statut final', dateDeces: 'Date décès', contact: 'Contact avec malade Ebola',
    prelev: 'Date prélèvement', analyse: 'Date analyse labo', confirmation: 'Date confirmation', admission: 'Date admission CTE',
    sortie: 'Date sortie CTE', vaccin: 'Vacciné MVE', extrait: 'Extrait le', casConf: 'Cas confirmés', nom: 'Nom',
    cleDoublon: 'Clé doublon', saisie: 'Date saisie',
    noncasPI: 'Non cas (DHIS2)' // indicateur DHIS2 « Non Cas » (mkyT4tONdVc) ; absent d'un classeur Excel
  };
  var DATES = ['notif', 'dds', 'dateDeces', 'prelev', 'analyse', 'confirmation', 'admission', 'sortie', 'saisie'];

  // Variables clés (complétude) et incohérences, dans l'ordre du classeur
  var MANQUES = [
    ['mID', 'N° alerte', 'manque N° alerte'],
    ['mDDS', 'Date début symptômes', 'manque date début symptômes'],
    ['mNature', 'Nature alerte', 'manque nature alerte'],
    ['mConcl', 'Conclusion alerte', 'manque conclusion alerte'],
    ['mAge', 'Âge', 'manque âge'],
    ['mPrelev', 'Date prélèvement (validées)', 'manque date prélèvement (validées)'],
    ['mClassif', 'Classification (analysées)', 'manque classification (analysées)'],
    ['mContact', 'Contact Ebola (validées)', 'manque contact Ebola (validées)'],
    ['mVaccin', 'Statut vaccinal (validées)', 'manque statut vaccinal (validées)'],
    ['mStatut', 'Statut final (confirmés)', 'manque statut final (confirmés)']
  ];
  var INCOH = [
    ['iDDSNotif', 'DDS après notification'],
    ['iDDSAnc', 'DDS trop ancienne'],
    ['iPrelDDS', 'Prélèvement avant DDS'],
    ['iAnaPrel', 'Analyse avant prélèvement'],
    ['iDecDDS', 'Décès avant DDS'],
    ['iSortie', 'Sortie CTE avant admission'],
    ['iAge', 'Âge invalide'],
    ['iConf', '« Cas confirmés » ≠ classification'],
    ['iDeces', 'Statut de décès'],
    ['iFutur', 'Date dans le futur'],
    ['iInval', 'Invalidée mais confirmée'],
    ['iConfNC', 'Confirmée et non-cas à la fois'],
    ['iSansAna', 'Classée sans date d’analyse'],
    ['iDoublon', 'Doublon probable (nom, sexe, âge)']
  ];

  function norm(s) {
    return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  }
  function blank(v) { return v == null || (typeof v === 'string' && v.trim() === ''); }
  function eq(v, t) { return !blank(v) && norm(v) === norm(t); }
  function starts(v, t) { return !blank(v) && norm(v).slice(0, norm(t).length) === norm(t); }
  function num(v) { return typeof v === 'number' && !isNaN(v) ? v : (typeof v === 'string' && v.trim() !== '' && !isNaN(+v) ? +v : null); }

  var EPOCH = Date.UTC(1899, 11, 30);
  function toSerial(v, keepTime) {
    if (blank(v)) return null;
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

  // Clé de doublon probable : nom normalisé | sexe | âge (ans) | âge (mois)
  function cleNom(nom, sexe, ans, mois) {
    var n = String(nom == null ? '' : nom).trim().toUpperCase().replace(/[ -]/g, '');
    if (!n) return null;
    var f = function (v) { return v == null ? '' : String(v); };
    return n + '|' + norm(sexe) + '|' + f(ans) + '|' + f(mois);
  }

  /* Lignes brutes (1re ligne = en-têtes) → alertes contrôlées. */
  function normalize(table, seuils) {
    var S = Object.assign({}, SEUILS_DEFAUT, seuils || {});
    if (!table || !table.length) throw new Error('La feuille « Données » est vide.');
    var header = table[0].map(norm), idx = {};
    Object.keys(COLS).forEach(function (k) { idx[k] = header.indexOf(norm(COLS[k])); });
    var missing = ['notif', 'zone', 'aire', 'conclusion', 'classif'].filter(function (k) { return idx[k] < 0; }).map(function (k) { return COLS[k]; });
    if (missing.length) throw new Error('Colonnes introuvables : ' + missing.join(', ') + '. Est-ce bien le classeur Qualité des données ?');
    var raw = [], sansDate = 0, extrait = null;
    for (var i = 1; i < table.length; i++) {
      var t = table[i]; if (!t || !t.some(function (c) { return !blank(c); })) continue;
      var r = {};
      Object.keys(COLS).forEach(function (k) { r[k] = idx[k] >= 0 ? t[idx[k]] : null; if (r[k] === undefined) r[k] = null; });
      DATES.forEach(function (k) { r[k] = toSerial(r[k]); });
      var ex = toSerial(r.extrait, true);
      if (ex != null && (extrait == null || ex > extrait)) extrait = ex;
      r.ex = ex;
      if (r.notif == null) { sansDate++; continue; } // comme Contrôles!A : alerte exclue
      r.ageAns = num(r.ageAns); r.ageMois = num(r.ageMois);
      r.cle = !blank(r.cleDoublon) ? String(r.cleDoublon) : cleNom(r.nom, r.sexe, r.ageAns, r.ageMois);
      raw.push(r);
    }
    if (!raw.length) throw new Error('Aucune alerte avec une date de notification dans le fichier.');
    var cpt = {};
    raw.forEach(function (r) { if (r.cle) cpt[r.cle] = (cpt[r.cle] || 0) + 1; });
    var rows = raw.map(function (r) {
      var conf = num(r.casConf) === 1, valide = eq(r.conclusion, 'Validée');
      // Non-cas : indicateur DHIS2 « Non Cas » s'il est publié, sinon classification finale « Non cas »
      var noncas = num(r.noncasPI) === 1 || eq(r.classif, 'Non cas');
      var dec = !blank(r.statut) && eq(r.statut, 'Décédé');
      var lim = r.ex != null ? Math.floor(r.ex) : Math.floor(nowSerial());
      var a = {
        id: blank(r.id) ? null : String(r.id).trim(), n: r.notif, saisie: r.saisie, dds: r.dds, zone: blank(r.zone) ? null : String(r.zone),
        aire: blank(r.aire) ? null : String(r.aire), sexe: blank(r.sexe) ? '' : String(r.sexe), ageAns: r.ageAns, ageMois: r.ageMois,
        conclusion: blank(r.conclusion) ? '' : String(r.conclusion), classif: blank(r.classif) ? '' : String(r.classif),
        cle: r.cle, nGroupe: r.cle ? cpt[r.cle] : 0,
        // Complétude : 1 = manquant, 0 = renseigné, null = non applicable
        mID: blank(r.id) ? 1 : 0,
        mDDS: r.dds == null ? 1 : 0,
        mNature: blank(r.nature) ? 1 : 0,
        mConcl: blank(r.conclusion) ? 1 : 0,
        mAge: r.ageAns == null && r.ageMois == null && blank(r.ageAns) ? 1 : 0,
        mPrelev: valide ? (r.prelev == null ? 1 : 0) : null,
        mClassif: r.analyse != null ? (blank(r.classif) ? 1 : 0) : null,
        mContact: valide ? (blank(r.contact) ? 1 : 0) : null,
        mVaccin: valide ? (blank(r.vaccin) ? 1 : 0) : null,
        mStatut: conf ? (blank(r.statut) ? 1 : 0) : null,
        // Incohérences
        iDDSNotif: r.dds != null && r.dds > r.notif ? 1 : 0,
        iDDSAnc: r.dds != null && r.notif - r.dds > S.DDSmax ? 1 : 0,
        iPrelDDS: r.prelev != null && r.dds != null && r.prelev < r.dds ? 1 : 0,
        iAnaPrel: r.analyse != null && r.prelev != null && r.analyse < r.prelev ? 1 : 0,
        iDecDDS: r.dateDeces != null && r.dds != null && r.dateDeces < r.dds ? 1 : 0,
        iSortie: r.sortie != null && r.admission != null && r.sortie < r.admission ? 1 : 0,
        iAge: (r.ageAns != null && (r.ageAns < 0 || r.ageAns > S.AgeMax)) || (r.ageMois != null && (r.ageMois < 0 || r.ageMois > 11)) ? 1 : 0,
        iConf: (conf && !eq(r.classif, 'Cas confirmé')) || (eq(r.classif, 'Cas confirmé') && !conf) ? 1 : 0,
        iDeces: (dec && r.dateDeces == null) || (r.dateDeces != null && eq(r.statut, 'Vivant')) || (eq(r.nature, 'Décès') && eq(r.statut, 'Vivant')) ? 1 : 0,
        iFutur: ['notif', 'dds', 'prelev', 'analyse', 'confirmation', 'dateDeces', 'admission', 'sortie'].some(function (k) { return r[k] != null && r[k] > lim; }) ? 1 : 0,
        iInval: starts(r.conclusion, 'Invalidé') && (conf || eq(r.classif, 'Cas confirmé')) ? 1 : 0,
        // Compté à la fois dans « Cas confirmés » et dans « Non-cas » (résultats de labo ou classification contradictoires)
        iConfNC: conf && noncas ? 1 : 0,
        // Classée confirmée ou non-cas (classement qui suppose une analyse) sans « Date analyse labo »
        iSansAna: (conf || noncas || eq(r.classif, 'Cas confirmé')) && r.analyse == null ? 1 : 0,
        iDoublon: r.cle && cpt[r.cle] > 1 ? 1 : 0,
        delNotif: r.dds != null ? r.notif - r.dds : null,
        delPrel: r.prelev != null ? r.prelev - r.notif : null,
        delLabo: r.analyse != null && r.prelev != null ? r.analyse - r.prelev : null,
        sem: r.notif - weekday1(r.notif) + 1,
        prelev: r.prelev, analyse: r.analyse
      };
      // Retards (calculés à la date de référence, voir compute)
      a.pInv = blank(r.conclusion) || starts(r.conclusion, 'En cours');
      a.pLabo = r.prelev != null && r.analyse == null;
      a.pClassif = valide && (blank(r.classif) || eq(r.classif, 'Cas suspect'));
      a.nbManq = 0; a.nbAppl = 0;
      MANQUES.forEach(function (m) { var v = a[m[0]]; if (v != null) { a.nbAppl++; if (v === 1) a.nbManq++; } });
      a.nbIncoh = INCOH.reduce(function (t, c) { return t + a[c[0]]; }, 0);
      a.nbAnom = a.nbManq + a.nbIncoh;
      return a;
    });
    // Ordre du classeur : date de notification puis date de saisie
    rows.sort(function (x, y) { return x.n - y.n || ((x.saisie == null ? -Infinity : x.saisie) - (y.saisie == null ? -Infinity : y.saisie)); });
    rows.forEach(function (r, i) { r.rang = i; });
    return { rows: rows, extrait: extrait, sansDate: sansDate, charge: raw.length + sansDate };
  }

  function zonesOf(rows) {
    var seen = {}, list = ZONES.slice();
    list.forEach(function (z) { seen[norm(z)] = 1; });
    rows.forEach(function (r) { if (r.zone && !seen[norm(r.zone)]) { seen[norm(r.zone)] = 1; list.push(r.zone); } });
    return list;
  }
  function airesOf(rows, zone) {
    var s = {};
    rows.forEach(function (r) { if (r.aire && (!zone || eq(r.zone, zone))) s[r.aire] = 1; });
    return Object.keys(s).sort(function (a, b) { return a.localeCompare(b, 'fr'); });
  }
  function sum(rs, f) { var t = 0; for (var i = 0; i < rs.length; i++) t += f(rs[i]); return t; }
  function moy(rs, f) {
    var t = 0, n = 0; rs.forEach(function (r) { var v = f(r); if (v != null && v >= 0) { t += v; n++; } });
    return n ? t / n : null;
  }
  function completude(rs) { var a = sum(rs, function (r) { return r.nbAppl; }); return a ? 1 - sum(rs, function (r) { return r.nbManq; }) / a : null; }
  function detail(r) {
    return MANQUES.filter(function (m) { return r[m[0]] === 1; }).map(function (m) { return m[2]; })
      .concat(INCOH.filter(function (c) { return r[c[0]] === 1; }).map(function (c) { return c[1].replace('Doublon probable (nom, sexe, âge)', 'doublon probable'); }));
  }

  /* f = {zone, aire, debut, fin, ref}, S = seuils */
  function compute(rows, f, seuils) {
    var S = Object.assign({}, SEUILS_DEFAUT, seuils || {});
    var inGeo = function (r) { return (!f.zone ? r.zone != null : eq(r.zone, f.zone)) && (!f.aire || eq(r.aire, f.aire)); };
    var inPer = function (r) { return r.n >= f.debut && r.n <= f.fin; };
    var rInv = function (r) { return r.pInv && f.ref - r.n > S.RetInv; };
    var rLabo = function (r) { return r.pLabo && f.ref - r.prelev > S.RetLabo; };
    var rClassif = function (r) { return r.pClassif && f.ref - r.n > S.RetClassif; };
    var geo = rows.filter(inGeo), per = geo.filter(inPer), n = per.length;
    var avecPrel = per.filter(function (r) { return r.delPrel != null && r.delPrel >= 0; });
    var dern = geo.length ? Math.max.apply(null, geo.map(function (r) { return r.n; })) : null;
    var k = {
      alertes: n,
      pSans: n ? sum(per, function (r) { return r.nbAnom ? 0 : 1; }) / n : null,
      avecAnom: sum(per, function (r) { return r.nbAnom ? 1 : 0; }),
      completude: completude(per),
      avecIncoh: sum(per, function (r) { return r.nbIncoh ? 1 : 0; }),
      doublons: sum(per, function (r) { return r.iDoublon; }),
      delNotif: moy(per, function (r) { return r.delNotif; }),
      pPrelCible: avecPrel.length ? sum(avecPrel, function (r) { return r.delPrel <= S.CiblePrelev ? 1 : 0; }) / avecPrel.length : null,
      delLabo: moy(per, function (r) { return r.delLabo; }),
      discord: sum(per, function (r) { return r.iConf; }),
      statutManq: sum(per, function (r) { return r.mStatut === 1 ? 1 : 0; }),
      retInv: sum(geo, function (r) { return rInv(r) ? 1 : 0; }),
      retLabo: sum(geo, function (r) { return rLabo(r) ? 1 : 0; }),
      retClassif: sum(geo, function (r) { return rClassif(r) ? 1 : 0; }),
      sansConcl: sum(geo, function (r) { return r.mConcl; }),
      derniere: dern,
      joursSans: dern == null ? null : f.ref - dern
    };
    var parVariable = MANQUES.map(function (m) {
      var appl = per.filter(function (r) { return r[m[0]] != null; }), manq = sum(appl, function (r) { return r[m[0]]; });
      return { variable: m[1], applicables: appl.length, manquantes: manq, taux: appl.length ? 1 - manq / appl.length : null };
    });
    var parType = INCOH.map(function (c) { return { type: c[1], alertes: sum(per, function (r) { return r[c[0]]; }) }; });
    var wRef = f.ref - weekday1(f.ref) + 1, semaines = [];
    for (var w = wRef - 7 * 11; w <= wRef; w += 7) {
      var rw = geo.filter(function (r) { return r.sem === w; });
      semaines.push({ semaine: w, alertes: rw.length, sans: sum(rw, function (r) { return r.nbAnom ? 0 : 1; }),
        pSans: rw.length ? sum(rw, function (r) { return r.nbAnom ? 0 : 1; }) / rw.length : null, completude: completude(rw) });
    }
    var zones = zonesOf(rows).map(function (z) {
      var zr = rows.filter(function (r) { return eq(r.zone, z); }), zp = zr.filter(inPer);
      return {
        zone: z, alertes: zp.length,
        pSans: zp.length ? sum(zp, function (r) { return r.nbAnom ? 0 : 1; }) / zp.length : null,
        completude: zp.length ? completude(zp) : null,
        avecIncoh: sum(zp, function (r) { return r.nbIncoh ? 1 : 0; }),
        doublons: sum(zp, function (r) { return r.iDoublon; }),
        delNotif: moy(zp, function (r) { return r.delNotif; }), delLabo: moy(zp, function (r) { return r.delLabo; }),
        retInv: sum(zr, function (r) { return rInv(r) ? 1 : 0; }), retLabo: sum(zr, function (r) { return rLabo(r) ? 1 : 0; }),
        retClassif: sum(zr, function (r) { return rClassif(r) ? 1 : 0; })
      };
    });
    var all = rows.filter(inPer);
    var total = {
      zone: 'Total', alertes: sum(zones, function (z) { return z.alertes; }),
      avecIncoh: sum(zones, function (z) { return z.avecIncoh; }), doublons: sum(zones, function (z) { return z.doublons; }),
      retInv: sum(zones, function (z) { return z.retInv; }), retLabo: sum(zones, function (z) { return z.retLabo; }),
      retClassif: sum(zones, function (z) { return z.retClassif; }),
      completude: all.length ? completude(all) : null,
      delNotif: moy(all, function (r) { return r.delNotif; }), delLabo: moy(all, function (r) { return r.delLabo; })
    };
    total.pSans = total.alertes ? sum(zones, function (z) { return (z.pSans || 0) * z.alertes; }) / total.alertes : null;
    // Alertes à corriger : les plus récentes d'abord
    var aCorriger = per.filter(function (r) { return r.nbAnom > 0; }).sort(function (a, b) { return b.rang - a.rang; })
      .map(function (r) { return { r: r, detail: detail(r) }; });
    // Doublons probables dans le filtre, regroupés
    var groupes = {}, ordre = [];
    per.forEach(function (r) { if (r.iDoublon) { if (!groupes[r.cle]) { groupes[r.cle] = []; ordre.push(r.cle); } groupes[r.cle].push(r); } });
    var doublons = ordre.map(function (c) { return { total: groupes[c][0].nGroupe, membres: groupes[c].sort(function (a, b) { return a.rang - b.rang; }) }; })
      .sort(function (a, b) { return a.membres[0].rang - b.membres[0].rang; });
    return { k: k, parVariable: parVariable, parType: parType, semaines: semaines, zones: zones, total: total, aCorriger: aCorriger, doublons: doublons };
  }

  var api = { normalize: normalize, compute: compute, zonesOf: zonesOf, airesOf: airesOf, toSerial: toSerial, serialToUTC: serialToUTC,
    nowSerial: nowSerial, cleNom: cleNom, norm: norm, SEUILS_DEFAUT: SEUILS_DEFAUT, MANQUES: MANQUES, INCOH: INCOH };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.MVEQ = api;
})(this);
