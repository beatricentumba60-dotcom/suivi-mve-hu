/* Moteur de calcul du tableau de bord MVE.
   Reproduit les formules du classeur Excel « Tableau_de_bord_MVE_DHIS2 »
   (feuilles Lignes, Calculs et Tableau de bord). Les dates sont des numéros
   de série Excel (jours entiers), comme dans le classeur. */
(function (root) {
  'use strict';

  var ZONES_HAUT_UELE = ['Aba', 'Boma Mangbetu', 'Doruma', 'Dungu', 'Faradje', 'Gombari', 'Isiro',
    'Makoro', 'Niangara', 'Pawa', 'Rungu', 'Wamba', 'Watsa'];
  var TRANCHES = ['0-4', '5-14', '15-24', '25-44', '45-64', '65+'];

  // En-têtes de la feuille « Données » (liste linéaire DHIS2)
  var COLS = {
    id: 'ID alerte', notif: 'Date notification', dds: 'Date début symptômes', zone: 'Zone de santé',
    aire: 'Aire de santé', sexe: 'Sexe', ageAns: 'Âge (ans)', ageMois: 'Âge (mois)', nature: 'Nature alerte',
    conclusion: 'Conclusion alerte', classif: 'Classification finale', statut: 'Statut final', dateDeces: 'Date décès',
    statutAvant: 'Statut avant admission', contact: 'Contact avec malade Ebola', prelev: 'Date prélèvement',
    analyse: 'Date analyse labo', admission: 'Date admission CTE', sortie: 'Date sortie CTE', vaccin: 'Vacciné MVE',
    extrait: 'Extrait le', allaite: 'Allaite', enceinte: 'Enceinte', casConf: 'Cas confirmés', profession: 'Profession',
    // Indicateurs de programme DHIS2 (Ox52yq3XZUX, mkyT4tONdVc) : absents si la source est un classeur Excel
    decPI: 'Décès confirmés (DHIS2)', noncasPI: 'Non cas (DHIS2)'
  };
  var DATE_FIELDS = ['notif', 'dds', 'dateDeces', 'prelev', 'analyse', 'admission', 'sortie'];
  var REQUIRED = ['notif', 'zone', 'aire', 'casConf', 'classif'];

  function norm(s) {
    return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/\s+/g, ' ').trim().toLowerCase();
  }
  function blank(v) { return v == null || (typeof v === 'string' && v.trim() === ''); }
  // Comparaison texte comme COUNTIFS : insensible à la casse (et ici aux accents)
  function eq(v, txt) { return !blank(v) && norm(v) === norm(txt); }

  var EPOCH = Date.UTC(1899, 11, 30);
  function serialFromDate(d) { return (d.getTime() - d.getTimezoneOffset() * 60000 - EPOCH) / 864e5; }
  function toSerial(v, keepTime) {
    if (blank(v)) return null;
    var s = null;
    if (typeof v === 'number') s = v;
    else if (v instanceof Date) s = serialFromDate(v);
    else {
      var t = String(v).trim(), m;
      if ((m = t.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/))) {
        s = (Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) - EPOCH) / 864e5;
      } else if ((m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: (\d{1,2})[:h](\d{2}))?/))) {
        s = (Date.UTC(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)) - EPOCH) / 864e5;
      } else if (!isNaN(+t)) s = +t;
    }
    if (s == null || isNaN(s)) return null;
    return keepTime ? s : Math.floor(s + 1e-9);
  }
  function serialToUTC(s) { return new Date(EPOCH + s * 864e5); }
  function weekday1(s) { return ((serialToUTC(s).getUTCDay() + 6) % 7) + 1; } // lundi = 1
  function todaySerial() { var d = new Date(); return Math.floor(serialFromDate(d)); }

  var OUI = ['oui', '1', 'yes', 'true', 'vrai'];

  /* Transforme le tableau brut (1re ligne = en-têtes) en lignes normalisées. */
  function normalize(table) {
    if (!table || !table.length) throw new Error('La feuille « Données » est vide.');
    var header = table[0].map(norm), idx = {}, missing = [];
    Object.keys(COLS).forEach(function (k) { idx[k] = header.indexOf(norm(COLS[k])); });
    REQUIRED.forEach(function (k) { if (idx[k] < 0) missing.push(COLS[k]); });
    if (missing.length) throw new Error('Colonnes introuvables : ' + missing.join(', ') + '.');
    // Les indicateurs DHIS2 ne servent que si la colonne existe ET contient des valeurs
    // (sinon : anciennes règles du classeur)
    var rempli = function (k) {
      if (idx[k] < 0) return false;
      for (var j = 1; j < table.length; j++) if (table[j] && !blank(table[j][idx[k]])) return true;
      return false;
    };
    var avecDecPI = rempli('decPI'), avecNcPI = rempli('noncasPI');
    var rows = [], extrait = null;
    for (var i = 1; i < table.length; i++) {
      var raw = table[i];
      if (!raw) continue;
      var get = function (k) { return idx[k] >= 0 ? raw[idx[k]] : null; };
      var notif = toSerial(get('notif'));
      if (notif == null) continue; // comme Lignes!A : ligne ignorée sans date de notification
      var r = {};
      Object.keys(COLS).forEach(function (k) { r[k] = get(k); });
      DATE_FIELDS.forEach(function (k) { r[k] = toSerial(r[k]); });
      var ex = toSerial(r.extrait, true);
      if (ex != null && (extrait == null || ex > extrait)) extrait = ex;
      ['zone', 'aire', 'sexe', 'nature', 'conclusion', 'classif', 'statut', 'statutAvant', 'contact', 'vaccin']
        .forEach(function (k) { r[k] = blank(r[k]) ? null : String(r[k]).trim(); });
      // Variables calculées (feuille Lignes)
      var conf = Number(r.casConf) === 1;
      var cas = conf || eq(r.classif, 'Cas probable');
      var decesRegle = eq(r.statut, 'Décédé') || eq(r.nature, 'Décès') || eq(r.statutAvant, 'Décédé') || r.dateDeces != null;
      var decPI = Number(r.decPI) === 1;
      // Cas confirmés : indicateur DHIS2 « Décès confirmés » ; cas probables : règle du classeur (inchangée)
      var deces = conf && avecDecPI ? decPI : decesRegle;
      var decConf = avecDecPI ? decPI : (conf && decesRegle);
      var courbe = r.dds != null ? r.dds : notif;
      var age = !blank(r.ageAns) && !isNaN(+r.ageAns) ? +r.ageAns : (!blank(r.ageMois) && !isNaN(+r.ageMois) ? +r.ageMois / 12 : null);
      var tranche = age == null ? null : age < 5 ? '0-4' : age < 15 ? '5-14' : age < 25 ? '15-24' : age < 45 ? '25-44' : age < 65 ? '45-64' : '65+';
      var prof = blank(r.profession) ? 'Non renseignée' : String(r.profession).replace(/\s+/g, ' ').trim();
      rows.push({
        n: notif, zone: r.zone, aire: r.aire,
        sexe: eq(r.sexe, 'Masculin') ? 'M' : (eq(r.sexe, 'Feminin') ? 'F' : null),
        conf: conf ? 1 : 0, cas: cas ? 1 : 0, dec: deces ? 1 : 0,
        prob: eq(r.classif, 'Cas probable') ? 1 : 0,
        noncas: (avecNcPI ? Number(r.noncasPI) === 1 : eq(r.classif, 'Non cas')) ? 1 : 0,
        decConf: decConf ? 1 : 0,
        valid: eq(r.conclusion, 'Validée') ? 1 : 0,
        attente: eq(r.conclusion, 'Validée') && (blank(r.classif) || eq(r.classif, 'Cas suspect')) ? 1 : 0,
        contact: eq(r.contact, 'Oui') ? 1 : 0,
        vacc: eq(r.vaccin, 'Oui') ? 1 : 0,
        natDec: eq(r.nature, 'Décès') ? 1 : 0,
        prel: r.prelev != null ? 1 : 0, anal: r.analyse != null ? 1 : 0,
        cte: r.admission != null && r.sortie == null && !deces ? 1 : 0,
        sem: courbe - weekday1(courbe) + 1, j: courbe,
        delai: r.dds != null ? notif - r.dds : null,
        tr: tranche,
        enc: OUI.indexOf(norm(r.enceinte)) >= 0 ? 1 : 0,
        all: OUI.indexOf(norm(r.allaite)) >= 0 ? 1 : 0,
        prof: prof
      });
    }
    if (!rows.length) throw new Error('Aucune alerte avec une date de notification dans le fichier.');
    return { rows: rows, extrait: extrait, sourcesDHIS2: { deces: avecDecPI, noncas: avecNcPI } };
  }

  function zonesOf(rows) {
    var seen = {}, list = ZONES_HAUT_UELE.slice();
    list.forEach(function (z) { seen[norm(z)] = 1; });
    rows.forEach(function (r) { if (r.zone && !seen[norm(r.zone)]) { seen[norm(r.zone)] = 1; list.push(r.zone); } });
    return list;
  }
  function airesOf(rows, zone) {
    var s = {};
    rows.forEach(function (r) { if (r.aire && (!zone || eq(r.zone, zone))) s[r.aire] = 1; });
    return Object.keys(s).sort(function (a, b) { return a.localeCompare(b, 'fr'); });
  }

  function sum(rows, f) { var t = 0; for (var i = 0; i < rows.length; i++) t += f(rows[i]); return t; }
  function ratio(a, b) { return b ? a / b : 0; }
  function lastConf(rows, ref) {
    var m = 0; rows.forEach(function (r) { if (r.conf && r.n <= ref && r.n > m) m = r.n; });
    return m || null;
  }

  /* Calcule tous les indicateurs.
     f = {zone: '' (toutes) ou nom, aire: '' ou nom, debut, fin, ref} (numéros de série) */
  function compute(rows, f) {
    var inGeo = function (r) {
      return (!f.zone ? r.zone != null : eq(r.zone, f.zone)) && (!f.aire || eq(r.aire, f.aire));
    };
    var geo = rows.filter(inGeo);
    var per = geo.filter(function (r) { return r.n >= f.debut && r.n <= f.fin; });
    var d7 = function (r) { return r.n > f.ref - 7 && r.n <= f.ref; };
    var cas = per.filter(function (r) { return r.cas; });
    var nCas = cas.length;
    var lc = lastConf(geo, f.ref);
    var delais = per.filter(function (r) { return r.conf && r.delai != null && r.delai >= 0; });
    var k = {
      cas: nCas,
      conf: sum(per, function (r) { return r.conf; }),
      prob: sum(per, function (r) { return r.prob; }),
      deces: sum(cas, function (r) { return r.dec; }),
      // Comme la tuile DHIS2 « Cumul Décès Confirmés » (indicateur Ox52yq3XZUX)
      decesConf: sum(per, function (r) { return r.decConf; }),
      casJour: sum(geo, function (r) { return r.cas && r.n === f.ref ? 1 : 0; }),
      cas7: sum(geo, function (r) { return r.cas && d7(r) ? 1 : 0; }),
      alertes: per.length,
      alertesJour: sum(geo, function (r) { return r.n === f.ref ? 1 : 0; }),
      attente: sum(per, function (r) { return r.attente; }),
      pContact: ratio(sum(cas, function (r) { return r.contact; }), nCas),
      dernierConf: lc,
      joursDepuis: lc == null ? null : f.ref - lc,
      validees: sum(per, function (r) { return r.valid; }),
      prelevees: sum(per, function (r) { return r.prel; }),
      analysees: sum(per, function (r) { return r.anal; }),
      delai: delais.length ? sum(delais, function (r) { return r.delai; }) / delais.length : null,
      decesCom: sum(cas, function (r) { return r.natDec; }),
      cte: sum(geo, function (r) { return r.cte; }),
      pVacc: ratio(sum(cas, function (r) { return r.vacc; }), nCas),
      noncas: sum(per, function (r) { return r.noncas; }),
      // Comme Tableau de bord!L19 : toutes les alertes (cas ou non) dont la nature est « Décès »
      alertesDeces: sum(per, function (r) { return r.natDec; }),
      enceintes: sum(cas, function (r) { return r.enc; }),
      allaitantes: sum(cas, function (r) { return r.all; }),
      ps: sum(cas, function (r) { return r.prof === 'Personnel de santé' ? 1 : 0; }),
      ppl: sum(cas, function (r) { return r.prof === 'PPL' ? 1 : 0; })
    };
    k.letalite = ratio(k.deces, nCas);
    k.pValidees = ratio(k.validees, k.alertes);
    k.psppl = k.ps + k.ppl;

    // Courbe épidémique : semaine (lundi) de la DDS, sinon de la notification
    var semaines = {};
    cas.forEach(function (r) {
      var w = semaines[r.sem] || (semaines[r.sem] = { conf: 0, prob: 0, dec: 0 });
      if (r.conf) w.conf++; if (r.prob) w.prob++; if (r.dec) w.dec++;
    });
    var courbe = weeksRange(Object.keys(semaines).map(Number), f.fin).map(function (w) {
      var v = semaines[w] || { conf: 0, prob: 0, dec: 0 };
      return { semaine: w, conf: v.conf, prob: v.prob, dec: v.dec };
    });

    // Cas par semaine de notification (même base que la tuile « Cas »)
    var notifW = {};
    per.forEach(function (r) {
      if (!r.conf && !r.prob) return;
      var w = r.n - weekday1(r.n) + 1, o = notifW[w] || (notifW[w] = { conf: 0, prob: 0 });
      if (r.conf) o.conf++; if (r.prob) o.prob++;
    });
    // « Toute l'épidémie » : l'axe part de la 1re semaine avec des cas (comme la courbe épidémique),
    // pour qu'une date de notification aberrante (ex. 2024) n'écrase pas le graphique.
    // Période choisie (7 j, 30 j, dates fixes) : l'axe couvre toute la période.
    var debutAxe = f.periode === 'tout' && Object.keys(notifW).length ? null : f.debut;
    var parNotif = weeksRange(Object.keys(notifW).map(Number), f.fin, debutAxe == null ? undefined : debutAxe).map(function (w) {
      var v = notifW[w] || { conf: 0, prob: 0 }; return { semaine: w, conf: v.conf, prob: v.prob };
    });

    // 30 derniers jours jusqu'à la date de référence
    var jours = [];
    for (var d = f.ref - 29; d <= f.ref; d++) jours.push({ jour: d, alertes: 0, conf: 0, dec: 0 });
    geo.forEach(function (r) {
      var i = r.n - (f.ref - 29);
      if (i < 0 || i > 29) return;
      jours[i].alertes++; if (r.conf) jours[i].conf++; if (r.cas && r.dec) jours[i].dec++;
    });

    // Cas et décès par zone (ou par aire si une zone est choisie) – filtre période
    var parLieu;
    var perAll = rows.filter(function (r) { return r.n >= f.debut && r.n <= f.fin; });
    if (!f.zone) {
      parLieu = zonesOf(rows).map(function (z) {
        var rs = perAll.filter(function (r) { return eq(r.zone, z) && r.cas; });
        return { lieu: z, cas: rs.length, dec: sum(rs, function (r) { return r.dec; }) };
      });
    } else {
      parLieu = airesOf(rows, f.zone).map(function (a) {
        var rs = perAll.filter(function (r) { return eq(r.zone, f.zone) && eq(r.aire, a) && r.cas; });
        return { lieu: a, cas: rs.length, dec: sum(rs, function (r) { return r.dec; }) };
      });
    }
    parLieu.sort(function (a, b) { return b.cas - a.cas || b.dec - a.dec; });

    // Pyramide des âges
    var age = TRANCHES.map(function (t) {
      return {
        tranche: t,
        h: sum(cas, function (r) { return r.tr === t && r.sexe === 'M' ? 1 : 0; }),
        f: sum(cas, function (r) { return r.tr === t && r.sexe === 'F' ? 1 : 0; })
      };
    });
    var ageInconnu = sum(cas, function (r) { return r.tr == null || r.sexe == null ? 1 : 0; });

    // Situation par zone (ignore les filtres zone et aire)
    var zones = zonesOf(rows).map(function (z) {
      var zr = rows.filter(function (r) { return eq(r.zone, z); });
      var zp = zr.filter(function (r) { return r.n >= f.debut && r.n <= f.fin; });
      var zc = zp.filter(function (r) { return r.cas; });
      var l = lastConf(zr, f.ref);
      var o = {
        zone: z, cas: zc.length, conf: sum(zp, function (r) { return r.conf; }),
        dec: sum(zc, function (r) { return r.dec; }),
        cas7: sum(zr, function (r) { return r.cas && d7(r) ? 1 : 0; }),
        casJour: sum(zr, function (r) { return r.cas && r.n === f.ref ? 1 : 0; }),
        alertes7: sum(zr, function (r) { return d7(r) ? 1 : 0; }),
        attente: sum(zp, function (r) { return r.attente; }),
        dernierConf: l, joursDepuis: l == null ? null : f.ref - l
      };
      o.letalite = ratio(o.dec, o.cas);
      return o;
    });
    var totalZones = ['cas', 'conf', 'dec', 'cas7', 'casJour', 'alertes7', 'attente'].reduce(function (t, key) {
      t[key] = sum(zones, function (z) { return z[key]; }); return t;
    }, { zone: 'Total' });
    totalZones.letalite = ratio(totalZones.dec, totalZones.cas);

    // Situation par profession (filtres zone, aire et période)
    var profs = {};
    rows.forEach(function (r) { profs[r.prof] = 1; });
    var professions = Object.keys(profs).map(function (p) {
      var pg = geo.filter(function (r) { return r.prof === p; });
      var pp = pg.filter(function (r) { return r.n >= f.debut && r.n <= f.fin; });
      var pc = pp.filter(function (r) { return r.cas; });
      var l = lastConf(pg, f.ref);
      var o = {
        profession: p, alertes: pp.length, cas: pc.length, conf: sum(pp, function (r) { return r.conf; }),
        dec: sum(pc, function (r) { return r.dec; }),
        cas7: sum(pg, function (r) { return r.cas && d7(r) ? 1 : 0; }),
        alertes7: sum(pg, function (r) { return d7(r) ? 1 : 0; }),
        dernierConf: l, joursDepuis: l == null ? null : f.ref - l
      };
      o.letalite = ratio(o.dec, o.cas);
      o.pCas = ratio(o.cas, nCas);
      return o;
    }).filter(function (o) { return o.alertes || o.alertes7; })
      .sort(function (a, b) { return b.cas - a.cas || b.alertes - a.alertes || a.profession.localeCompare(b.profession, 'fr'); });

    return {
      k: k, courbe: courbe, parNotif: parNotif, jours: jours, parLieu: parLieu, lieuType: f.zone ? 'aire' : 'zone',
      age: age, ageInconnu: ageInconnu, zones: zones, totalZones: totalZones, professions: professions
    };
  }

  // Liste continue de lundis, du plus ancien présent jusqu'à la semaine de « fin »
  function weeksRange(weeks, fin, debut) {
    var lastW = fin - weekday1(fin) + 1;
    var first = weeks.length ? Math.min.apply(null, weeks) : (debut != null ? debut - weekday1(debut) + 1 : lastW);
    if (debut != null) first = Math.min(first, debut - weekday1(debut) + 1);
    var maxW = weeks.length ? Math.max(lastW, Math.max.apply(null, weeks)) : lastW;
    var out = [];
    for (var w = first; w <= maxW; w += 7) out.push(w);
    return out;
  }

  // Classes de la carte des cas confirmés (Calculs!B41:E43)
  function classeConf(n) {
    if (!n) return 0; if (n <= 10) return 1; if (n <= 50) return 2; if (n <= 100) return 3; return 4;
  }

  var api = {
    normalize: normalize, compute: compute, zonesOf: zonesOf, airesOf: airesOf, toSerial: toSerial,
    serialToUTC: serialToUTC, todaySerial: todaySerial, classeConf: classeConf, norm: norm, TRANCHES: TRANCHES
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MVE = api;
})(this);
