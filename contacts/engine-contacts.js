/* Moteur de calcul « Contacts MVE ».
   Reproduit les feuilles Lignes, LignesSuivis, Calculs et Tableau de bord du classeur
   « Tableau_de_bord_Contacts_MVE_DHIS2 ». Dates = numéros de série Excel (jours).
   Aucun nom n'est lu ni conservé : seules les colonnes A à V de « Données » et A à J de « Suivis » servent. */
(function (root) {
  'use strict';

  var ZONES = ['Aba', 'Boma Mangbetu', 'Doruma', 'Dungu', 'Faradje', 'Gombari', 'Isiro',
    'Makoro', 'Niangara', 'Pawa', 'Rungu', 'Wamba', 'Watsa'];
  var EPOCH = Date.UTC(1899, 11, 30);
  var PARAMS_DEFAUT = { Semaine1: (Date.UTC(2026, 5, 22) - EPOCH) / 864e5, Objectif: 0.9, Orange: 0.8 }; // Paramètres!C2, seuils du dictionnaire
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
  var DATES = ['dEnr', 'dDern', 'deb', 'fin', 'dIss'];
  var CIRC = ['Foyer', 'Hôpital', 'Funérailles', 'Lieu du travail'];
  var TRANCHES = ['0-4', '5-14', '15-24', '25-44', '45-64', '65+', 'Âge inconnu'];

  function norm(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase(); }
  function blank(v) { return v == null || (typeof v === 'string' && v.trim() === ''); }
  function eq(v, t) { return !blank(v) && norm(v) === norm(t); }
  function num(v) { return typeof v === 'number' && !isNaN(v) ? v : (typeof v === 'string' && v.trim() !== '' && !isNaN(+v) ? +v : null); }
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
  function tranche(a) { return a == null ? 'Âge inconnu' : a < 5 ? '0-4' : a < 15 ? '5-14' : a < 25 ? '15-24' : a < 45 ? '25-44' : a < 65 ? '45-64' : '65+'; }
  function txt(v) { return blank(v) ? null : String(v).trim(); }
  function index(table, cols, nom) {
    if (!table || !table.length) return null;
    var h = table[0].map(norm), idx = {};
    Object.keys(cols).forEach(function (k) { idx[k] = h.indexOf(norm(cols[k])); });
    return idx;
  }

  /* Ne garde que les colonnes connues (les noms des colonnes W à Y ne sont jamais conservés). */
  function epurer(table, cols) {
    var idx = index(table, cols); if (!idx) return [];
    var keys = Object.keys(cols);
    return [keys.map(function (k) { return cols[k]; })].concat(table.slice(1).map(function (r) {
      return keys.map(function (k) { var v = idx[k] >= 0 && r ? r[idx[k]] : null; return v === undefined ? null : v; });
    }));
  }

  function normalize(contacts, suivis) {
    var ic = index(contacts, COLS), is = index(suivis, SCOLS);
    if (!ic) throw new Error('La feuille « Données » est vide.');
    var miss = ['enr', 'dEnr', 'zone', 'aire', 'deb', 'fin'].filter(function (k) { return ic[k] < 0; }).map(function (k) { return COLS[k]; });
    if (miss.length) throw new Error('Colonnes introuvables : ' + miss.join(', ') + '. Est-ce bien le classeur « Contacts MVE » ?');
    var extrait = null;
    var rows = [], byEnr = {};
    for (var i = 1; i < contacts.length; i++) {
      var t = contacts[i]; if (!t) continue;
      var r = {};
      Object.keys(COLS).forEach(function (k) { r[k] = ic[k] >= 0 ? t[ic[k]] : null; if (r[k] === undefined) r[k] = null; });
      if (blank(r.enr)) continue; // Lignes!A
      DATES.forEach(function (k) { r[k] = toSerial(r[k]); });
      var ex = toSerial(r.extrait, true); if (ex != null && (extrait == null || ex > extrait)) extrait = ex;
      ['enr', 'id', 'epid', 'zone', 'aire', 'sexe', 'rel', 'circ', 'type', 'notif', 'issue', 'sympt', 'prem'].forEach(function (k) { r[k] = txt(r[k]); });
      r.age = num(r.age); r.ageM = num(r.ageM); r.nbJ = num(r.nbJ) || 0;
      var a = r.age != null ? r.age : (r.ageM != null ? r.ageM / 12 : null);
      var o = {
        rang: rows.length, enr: r.enr, id: r.id, epid: r.epid, dEnr: r.dEnr, dDern: r.dDern, zone: r.zone, aire: r.aire,
        hom: eq(r.sexe, 'Masculin'), fem: eq(r.sexe, 'Feminin'), sexe: r.sexe, age: a, tr: tranche(a), circ: r.circ, type: r.type, rel: r.rel,
        deb: r.deb, fin: r.fin, issue: r.issue, dIss: r.dIss, nbJ: r.nbJ, sympt: eq(r.sympt, 'Oui'), prem: eq(r.prem, 'Oui'),
        sem: r.dEnr == null ? null : r.dEnr - weekday1(r.dEnr) + 1,
        delai: r.dDern == null ? null : (r.dEnr || 0) - r.dDern,
        vus: [] // dates des visites « vu = Oui »
      };
      rows.push(o);
      if (!(o.enr in byEnr)) byEnr[o.enr] = o; // EQUIV : 1re ligne du contact
    }
    if (!rows.length) throw new Error('Aucun contact dans le fichier.');
    // Feuille Suivis + LignesSuivis
    var visites = [];
    if (is) for (var j = 1; j < suivis.length; j++) {
      var s = suivis[j]; if (!s) continue;
      var enr = txt(is.enr >= 0 ? s[is.enr] : null), d = toSerial(is.date >= 0 ? s[is.date] : null);
      if (!enr || d == null) continue;
      var exs = toSerial(is.extrait >= 0 ? s[is.extrait] : null, true); if (exs != null && (extrait == null || exs > extrait)) extrait = exs;
      if (!eq(is.vu >= 0 ? s[is.vu] : null, 'Oui')) continue;
      var c = byEnr[enr]; if (!c) continue;
      c.vus.push(d);
      var jour = d - (c.deb || 0);
      var compte = !!c.zone && jour >= 1 && jour <= 21 && d >= (c.dEnr || 0) && (c.dIss == null || d < c.dIss);
      visites.push({ c: c, date: d, compte: compte });
    }
    return { rows: rows, visites: visites, extrait: extrait, charge: rows.length, nbSuivis: is ? suivis.length - 1 : 0 };
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
  function cnt(rows, f) { var n = 0; for (var i = 0; i < rows.length; i++) if (f(rows[i])) n++; return n; }
  function avg(v) { return v.length ? v.reduce(function (t, x) { return t + x; }, 0) / v.length : null; }
  function ratio(a, b) { return b ? a / b : null; }

  /* Lignes!B:E à une date de référence */
  function situation(r, ref) {
    var st;
    if ((r.dEnr || 0) > ref) st = 'Non encore listé';
    else if (r.issue && (r.dIss || 0) <= ref) st = r.issue;
    else if (ref > (r.fin || 0)) st = 'Sorti (21 j terminés)';
    else if (ref <= (r.deb || 0)) st = 'En attente (J0)';
    else st = 'En suivi';
    var o = { statut: st, vu: null, dernVu: null, cat: null };
    if (st === 'En suivi') {
      var dv = 0, vu = false;
      r.vus.forEach(function (d) { if (d === ref) vu = true; if (d <= ref && d > dv) dv = d; });
      o.vu = vu; o.dernVu = dv;
      o.cat = dv === ref ? 'Vu' : dv === 0 ? 'Jamais vu' : ref - dv >= 2 ? 'Non vu ≥ 2 j' : 'Non vu 24 h';
    }
    return o;
  }
  /* Calculs!I : contacts à suivre le jour h (formule « listé, entre J1 et J21, sans issue ») */
  function aSuivre(r, h) {
    return r.dEnr != null && r.dEnr <= h && r.deb != null && r.deb < h && r.fin != null && r.fin >= h && !(r.dIss != null && r.dIss <= h);
  }

  /* f = {zone, aire, debut, fin, ref} */
  function compute(D, f, params) {
    var P = Object.assign({}, PARAMS_DEFAUT, params || {});
    var rows = D.rows, ref = f.ref;
    rows.forEach(function (r) { var s = situation(r, ref); r.statut = s.statut; r.vu = s.vu; r.dernVu = s.dernVu; r.cat = s.cat; });
    // Critères COUNTIFS : zone « * » = zone renseignée ; aire « <>§ » = tout
    var inGeo = function (r) { return (f.zone ? eq(r.zone, f.zone) : !!r.zone) && (!f.aire || eq(r.aire, f.aire)); };
    var inPer = function (r) { return r.dEnr != null && r.dEnr >= f.debut && r.dEnr <= f.fin; };
    var geo = rows.filter(inGeo), per = geo.filter(inPer);
    var vGeo = D.visites.filter(function (v) { return v.compte && inGeo(v.c); });
    var c = cnt;
    var k = {
      listes: per.length,
      aSuivre: c(geo, function (r) { return r.statut === 'En suivi'; }),
      vus: c(geo, function (r) { return r.vu === true; }),
      nouveauxJour: c(geo, function (r) { return r.dEnr === ref; }),
      nouveaux7: c(geo, function (r) { return r.dEnr != null && r.dEnr > ref - 7 && r.dEnr <= ref; }),
      jamaisVus: c(geo, function (r) { return r.cat === 'Jamais vu'; }),
      nonVus2: c(geo, function (r) { return r.cat === 'Non vu ≥ 2 j'; }),
      nonVus24: c(geo, function (r) { return r.cat === 'Non vu 24 h'; }),
      suspects: c(per, function (r) { return eq(r.issue, 'Devenu suspect'); }),
      confirmes: c(per, function (r) { return eq(r.issue, 'Confirmé'); }),
      sortis: c(per, function (r) { return r.statut === 'Sorti (21 j terminés)'; }),
      casSources: c(per, function (r) { return r.prem; }),
      symptomes: c(per, function (r) { return r.sympt; }),
      decedes: c(per, function (r) { return eq(r.issue, 'Décédé'); }),
      vusUneFois: c(per, function (r) { return r.vus.length > 0; })
    };
    var lies = c(per, function (r) { return !!r.epid; });
    k.taux = ratio(k.vus, k.aSuivre); k.nonVus = k.aSuivre - k.vus;
    k.parCas = ratio(lies, k.casSources); k.pLies = ratio(lies, k.listes);
    k.delai = avg(per.filter(function (r) { return r.delai != null && r.delai >= 0; }).map(function (r) { return r.delai; }));
    k.pVusUneFois = ratio(k.vusUneFois, k.listes);

    // Calculs!A:F – listés par semaine et circonstance (30 semaines à partir de Paramètres!C2)
    var semaines = [];
    for (var w = 0; w < 30; w++) {
      var lu = P.Semaine1 + 7 * w, rs = per.filter(function (r) { return r.sem === lu; }), o = { semaine: lu, total: rs.length };
      CIRC.forEach(function (ci) { o[ci] = c(rs, function (r) { return eq(r.circ, ci); }); });
      o['Non précisé'] = o.total - CIRC.reduce(function (t, ci) { return t + o[ci]; }, 0);
      semaines.push(o);
    }
    // Calculs!H:K – 30 jours jusqu'à la date de référence
    var jours = [];
    for (var j = 29; j >= 0; j--) {
      var h = ref - j, as = c(geo, function (r) { return aSuivre(r, h); }), vv = c(vGeo, function (v) { return v.date === h; });
      jours.push({ date: h, aSuivre: as, vus: vv, taux: as ? vv / as : 0 });
    }
    // Calculs!R:T – tranche d'âge et sexe (période)
    var ageSexe = TRANCHES.map(function (t) { return { tranche: t, h: c(per, function (r) { return r.tr === t && r.hom; }), f: c(per, function (r) { return r.tr === t && r.fem; }) }; });

    // Tableau par zone (toutes zones, sans filtre d'aire) et par aire (zone choisie)
    var ligne = function (rs, lieu) {
      var p = rs.filter(inPer), as = c(rs, function (r) { return r.statut === 'En suivi'; }), vu = c(rs, function (r) { return r.vu === true; });
      return {
        lieu: lieu, listes: p.length, n7: c(rs, function (r) { return r.dEnr != null && r.dEnr > ref - 7 && r.dEnr <= ref; }),
        aSuivre: as, vus: vu, taux: as ? vu / as : null, nonVus2: c(rs, function (r) { return r.cat === 'Non vu ≥ 2 j'; }),
        jamais: c(rs, function (r) { return r.cat === 'Jamais vu'; }), suspects: c(p, function (r) { return eq(r.issue, 'Devenu suspect'); }),
        confirmes: c(p, function (r) { return eq(r.issue, 'Confirmé'); }), sortis: c(p, function (r) { return r.statut === 'Sorti (21 j terminés)'; })
      };
    };
    var total = function (list, lieu) {
      var t = { lieu: lieu };
      ['listes', 'n7', 'aSuivre', 'vus', 'nonVus2', 'jamais', 'suspects', 'confirmes', 'sortis'].forEach(function (x) { t[x] = list.reduce(function (s, z) { return s + z[x]; }, 0); });
      t.taux = t.aSuivre ? t.vus / t.aSuivre : null; return t;
    };
    var zones = zonesOf(rows).map(function (z) { return ligne(rows.filter(function (r) { return eq(r.zone, z); }), z); });
    var zr = rows.filter(function (r) { return f.zone ? eq(r.zone, f.zone) : !!r.zone; });
    var aires = airesOf(rows, f.zone).map(function (a) { return ligne(zr.filter(function (r) { return eq(r.aire, a); }), a); });

    // Contacts à relancer (en suivi, non vus le jour de réf.) : jamais vus, puis non vus ≥ 2 j, puis 24 h
    var ordre = { 'Jamais vu': 0, 'Non vu ≥ 2 j': 1, 'Non vu 24 h': 2 };
    var relancer = geo.filter(function (r) { return r.statut === 'En suivi' && r.cat !== 'Vu'; }).sort(function (a, b) {
      return ordre[a.cat] - ordre[b.cat] || (a.dernVu || 0) - (b.dernVu || 0) || (a.deb || 0) - (b.deb || 0) || a.rang - b.rang;
    });

    return {
      k: k, semaines: semaines, jours: jours, ageSexe: ageSexe, zones: zones, totalZones: total(zones, 'Total'),
      aires: aires, totalAires: total(aires, 'Total'), relancer: relancer, analyse21: analyse21(D, f, P)
    };
  }

  /* Calculs!W:BD – analyse des 21 derniers jours complets (toutes zones, fin = veille de la date de référence) */
  function analyse21(D, f, P) {
    var rows = D.rows.filter(function (r) { return !f.zone || eq(r.zone, f.zone); });
    rows = rows.filter(function (r) { return !f.aire || eq(r.aire, f.aire); });
    var fin = f.ref - 1, deb = fin - 20, c = cnt;
    var dans = function (r) { return r.dEnr != null && r.dEnr >= deb && r.dEnr <= fin; };
    var p = rows.filter(dans), vis = D.visites.filter(function (v) { return v.compte && (!f.zone || eq(v.c.zone, f.zone)) && (!f.aire || eq(v.c.aire, f.aire)); });
    var jours = [];
    for (var h = deb; h <= fin; h++) {
      var as = c(rows, function (r) { return aSuivre(r, h); }), vu = c(vis, (function (hh) { return function (v) { return v.date === hh; }; })(h));
      jours.push({
        date: h, aSuivre: as, vus: vu, taux: as ? vu / as : 0, nonVus: as - vu,
        nouveaux: c(rows, (function (hh) { return function (r) { return r.dEnr === hh; }; })(h)),
        suspects: c(rows, (function (hh) { return function (r) { return eq(r.issue, 'Devenu suspect') && r.dIss === hh; }; })(h)),
        confirmes: c(rows, (function (hh) { return function (r) { return eq(r.issue, 'Confirmé') && r.dIss === hh; }; })(h))
      });
    }
    var sum = function (l, k) { return l.reduce(function (t, x) { return t + x[k]; }, 0); };
    var tot = { aSuivre: sum(jours, 'aSuivre'), vus: sum(jours, 'vus'), nonVus: sum(jours, 'nonVus'), nouveaux: sum(jours, 'nouveaux'), suspects: sum(jours, 'suspects'), confirmes: sum(jours, 'confirmes') };
    tot.taux = tot.aSuivre ? tot.vus / tot.aSuivre : 0;
    var mn = jours[0], mx = jours[0];
    jours.forEach(function (j) { if (j.taux < mn.taux) mn = j; if (j.taux > mx.taux) mx = j; });
    var nouv = p.length, cas = c(p, function (r) { return r.prem; }), lies = c(p, function (r) { return !!r.epid; });
    var cles = {
      nouveaux: nouv, casSources: cas, parCas: ratio(lies, cas), pLies: ratio(lies, nouv),
      delai: avg(p.filter(function (r) { return r.delai != null && r.delai >= 0; }).map(function (r) { return r.delai; })),
      cjSuivre: tot.aSuivre, visites: tot.vus, taux: tot.taux, tauxMin: mn.taux, dateMin: mn.date, tauxMax: mx.taux, dateMax: mx.date,
      joursObjectif: c(jours, function (j) { return j.taux >= P.Objectif; }), suspects: tot.suspects, confirmes: tot.confirmes,
      symptomes: c(p, function (r) { return r.sympt; }),
      termineSansIssue: c(rows, function (r) { return r.fin != null && r.fin >= deb && r.fin <= fin && !r.issue; }),
      enSuivi: c(rows, function (r) { return r.statut === 'En suivi'; }), jamaisVus: c(rows, function (r) { return r.cat === 'Jamais vu'; }),
      nonVus2: c(rows, function (r) { return r.cat === 'Non vu ≥ 2 j'; })
    };
    var semaines = [0, 1, 2].map(function (i) {
      var l = jours.slice(i * 7, i * 7 + 7), a = sum(l, 'aSuivre'), v = sum(l, 'vus');
      return { n: i + 1, du: l[0].date, au: l[6].date, aSuivre: a, vus: v, taux: a ? v / a : 0, nouveaux: sum(l, 'nouveaux'), suspects: sum(l, 'suspects') };
    });
    var zones = zonesOf(D.rows).filter(function (z) { return !f.zone || eq(z, f.zone); }).map(function (z) {
      var zr = rows.filter(function (r) { return eq(r.zone, z); }), zv = vis.filter(function (v) { return eq(v.c.zone, z); });
      var as = 0, vu = 0, as7 = 0, vu7 = 0;
      for (var h = deb; h <= fin; h++) {
        var a1 = c(zr, (function (hh) { return function (r) { return aSuivre(r, hh); }; })(h)), v1 = c(zv, (function (hh) { return function (v) { return v.date === hh; }; })(h));
        as += a1; vu += v1; if (h > fin - 7) { as7 += a1; vu7 += v1; }
      }
      return {
        lieu: z, listes: c(zr, dans), aSuivre: as, vus: vu, taux: as ? vu / as : null, taux7: as7 ? vu7 / as7 : null, _as7: as7, _vu7: vu7,
        enSuivi: c(zr, function (r) { return r.statut === 'En suivi'; }), jamais: c(zr, function (r) { return r.cat === 'Jamais vu'; })
      };
    });
    var tz = { lieu: 'Total' };
    ['listes', 'aSuivre', 'vus', 'enSuivi', 'jamais', '_as7', '_vu7'].forEach(function (x) { tz[x] = zones.reduce(function (t, z) { return t + z[x]; }, 0); });
    tz.taux = tz.aSuivre ? tz.vus / tz.aSuivre : null; tz.taux7 = tz._as7 ? tz._vu7 / tz._as7 : null;
    return { debut: deb, fin: fin, cles: cles, jours: jours, total: tot, semaines: semaines, zones: zones, totalZones: tz };
  }

  var api = { ZONES: ZONES, PARAMS_DEFAUT: PARAMS_DEFAUT, COLS: COLS, SCOLS: SCOLS, CIRC: CIRC, TRANCHES: TRANCHES, normalize: normalize, compute: compute,
    epurer: epurer, zonesOf: zonesOf, airesOf: airesOf, toSerial: toSerial, serialToUTC: serialToUTC, nowSerial: nowSerial, norm: norm };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.MVEC = api;
})(this);
