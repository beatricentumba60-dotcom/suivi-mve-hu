/* Moteur de calcul « Vaccination MVE ».
   Reproduit les feuilles Lignes, Calculs, Anomalies et Tableau de bord du classeur
   « Tableau_de_bord_Vaccination_MVE_DHIS2 ». Dates = numéros de série Excel (jours).
   Aucun nom n'est conservé : seule une clé de doublon anonyme est utilisée. */
(function (root) {
  'use strict';

  var ZONES = ['Aba', 'Boma Mangbetu', 'Doruma', 'Dungu', 'Faradje', 'Gombari', 'Isiro',
    'Makoro', 'Niangara', 'Pawa', 'Rungu', 'Wamba', 'Watsa'];
  var EPOCH = Date.UTC(1899, 11, 30);
  // Paramètres!D2:D8 et F2:F31 du classeur
  var PARAMS_DEFAUT = {
    DebutCampagne: (Date.UTC(2026, 9, 1) - EPOCH) / 864e5, AgeMin: 1, AgeMax: 100, DelaiMax: 1,
    Vert: 0.95, Orange: 0.8, InclDates: false,
    Lots: ['H24P0781AC', 'H24P07B1AC', 'H24P0181AC', 'H24P06A1AC', 'N24P07B1AC']
  };
  var COLS = {
    id: 'ID vaccination', dVac: 'Date vaccination', dEvt: 'Date événement', dIns: 'Date inscription', dSai: 'Date saisie',
    zone: 'Zone de santé', aire: 'Aire de santé', sexe: 'Sexe', age: 'Âge (ans)', contact: 'Contact d\'un cas',
    consent: 'Consentement', allaite: 'Femme allaitante', enceinte: 'Femme enceinte', vaccin: 'Vaccin', lot: 'Numéro de lot',
    obs15: 'Observation 15 min', mapi: 'MAPI', mapiGrave: 'MAPI grave', statut: 'Statut événement', nomRens: 'Nom renseigné',
    tel: 'Téléphone', village: 'Avenue/Village renseigné', suivis: 'Nb suivis post-vaccinaux', nbNom: 'Nb fiches même nom/sexe/âge',
    extrait: 'Extrait le', cleDoublon: 'Clé doublon', nom: 'Nom du vacciné'
  };
  var DATES = ['dVac', 'dEvt', 'dIns', 'dSai'];

  // Contrôles par fiche (Lignes!K:AA) : [clé, libellé court (Anomalies), libellé graphique]
  var ANOM = [
    ['a1', 'ID invalide', 'ID invalide'], ['a2', 'ID en double', 'ID en double'], ['a3', 'doublon probable', 'Doublon probable'],
    ['a4', 'sans vaccination', 'Sans vaccination'], ['a5', 'date vacc. manquante', 'Date vacc. manquante'],
    ['a6', 'date vacc. ≠ date événement', 'Date vacc. ≠ date événement'], ['a7', 'date hors campagne', 'Date hors campagne'],
    ['a8', 'vacc. avant inscription', 'Vacc. avant inscription'], ['a9', 'saisie avant vacc.', 'Saisie avant vacc.'],
    ['a10', 'sexe manquant', 'Sexe manquant'], ['a11', 'âge manquant/hors plage', 'Âge manquant/hors plage'],
    ['a12', 'grossesse/allaitement chez homme', 'Grossesse/allaitement chez homme'],
    ['a13', 'statut grossesse/allaitement manquant', 'Statut grossesse/allaitement manquant'],
    ['a14', 'consentement ≠ Oui', 'Consentement ≠ Oui'], ['a15', 'lot manquant/non conforme', 'Lot manquant/non conforme'],
    ['a16', 'observation 15 min ≠ Oui', 'Observation 15 min ≠ Oui'], ['a17', 'téléphone invalide', 'Téléphone invalide']
  ];
  var DIMS = [
    ['Complétude', 'Les variables clés du formulaire sont-elles remplies ?'],
    ['Validité', 'Les valeurs saisies respectent-elles le format et les valeurs attendus (ID, lot, âge, téléphone) ?'],
    ['Cohérence', 'Les dates et variables liées sont-elles logiques entre elles ?'],
    ['Conformité protocole', 'Le protocole de vaccination est-il documenté (consentement, observation 15 min, suivi) ?'],
    ['Unicité', 'Chaque personne vaccinée est-elle enregistrée une seule fois ?'],
    ['Promptitude', 'Les vaccinations sont-elles saisies rapidement dans DHIS2 ?']
  ];
  var TRANCHES = ['0-14', '15-24', '25-34', '35-44', '45-54', '55-64', '65 et +', 'Non renseigné'];
  var CLASSES = ['Négatif', '0 j', '1 j', '2-3 j', '4-7 j', '8 j et +'];

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
  // Clé de doublon, comme la requête Power Query : mots du nom (lettres seules, triés) | sexe | âge
  function cleNom(nom, sexe, age) {
    if (blank(nom)) return null;
    var mots = String(nom).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z ]/g, '').split(' ').filter(Boolean).sort();
    if (!mots.length) return null;
    return mots.join('') + '|' + norm(sexe) + '|' + (age == null ? '' : String(age));
  }
  function tranche(a) {
    if (a == null) return 'Non renseigné';
    return a < 15 ? '0-14' : a < 25 ? '15-24' : a < 35 ? '25-34' : a < 45 ? '35-44' : a < 55 ? '45-54' : a < 65 ? '55-64' : '65 et +';
  }
  function classe(d) { return d == null ? null : d < 0 ? 'Négatif' : d === 0 ? '0 j' : d === 1 ? '1 j' : d <= 3 ? '2-3 j' : d <= 7 ? '4-7 j' : '8 j et +'; }

  /* Feuille « Données » (1re ligne = en-têtes) → fiches contrôlées (feuille Lignes). */
  function normalize(table, params) {
    var P = Object.assign({}, PARAMS_DEFAUT, params || {});
    if (!table || !table.length) throw new Error('La feuille « Données » est vide.');
    var header = table[0].map(norm), idx = {};
    Object.keys(COLS).forEach(function (k) { idx[k] = header.indexOf(norm(COLS[k])); });
    var missing = ['dEvt', 'dIns', 'zone', 'aire', 'sexe'].filter(function (k) { return idx[k] < 0; }).map(function (k) { return COLS[k]; });
    if (missing.length) throw new Error('Colonnes introuvables : ' + missing.join(', ') + '. Est-ce bien le classeur « Vaccination MVE » ?');
    var raw = [], extrait = null, nId = {}, nCle = {};
    for (var i = 1; i < table.length; i++) {
      var t = table[i]; if (!t) continue;
      var r = {};
      Object.keys(COLS).forEach(function (k) { r[k] = idx[k] >= 0 ? t[idx[k]] : null; if (r[k] === undefined) r[k] = null; });
      DATES.forEach(function (k) { r[k] = toSerial(r[k]); });
      if (r.dEvt == null && r.dIns == null) continue; // comme Lignes!A
      var ex = toSerial(r.extrait, true); r.ex = ex;
      if (ex != null && (extrait == null || ex > extrait)) extrait = ex;
      ['id', 'zone', 'aire', 'sexe', 'contact', 'consent', 'allaite', 'enceinte', 'lot', 'obs15', 'mapi', 'mapiGrave', 'tel', 'village']
        .forEach(function (k) { r[k] = blank(r[k]) ? null : String(r[k]).trim(); });
      r.age = num(r.age); r.suivis = num(r.suivis) || 0; r.nbNom = num(r.nbNom);
      r.cle = !blank(r.cleDoublon) ? String(r.cleDoublon) : cleNom(r.nom, r.sexe, r.age);
      if (r.id) nId[r.id.toUpperCase()] = (nId[r.id.toUpperCase()] || 0) + 1;
      if (r.cle) nCle[r.cle] = (nCle[r.cle] || 0) + 1;
      raw.push(r);
    }
    if (!raw.length) throw new Error('Aucune fiche de vaccination dans le fichier.');
    var lots = {}; (P.Lots || []).forEach(function (l) { lots[norm(l)] = 1; });
    var today = Math.floor(nowSerial());
    var rows = raw.map(function (r, i) {
      var vac = r.dEvt != null, d = r.dVac != null ? r.dVac : (r.dEvt != null ? r.dEvt : null);
      var delai = d != null && r.dSai != null ? r.dSai - d : null;
      var fem = eq(r.sexe, 'Féminin'), hom = eq(r.sexe, 'Masculin');
      var lim = r.ex != null ? Math.floor(r.ex) : today;
      var nb = r.nbNom != null ? r.nbNom : (r.cle ? nCle[r.cle] : 1);
      var a = {
        rang: i, id: r.id, zone: r.zone, aire: r.aire, vac: vac, d: d, dSai: r.dSai, sem: d == null ? null : d - weekday1(d) + 1,
        fem: fem, hom: hom, age: r.age, tr: tranche(r.age), delai: delai, cl: classe(delai), cle: r.cle,
        contact: eq(r.contact, 'Oui'), consentOui: eq(r.consent, 'Oui'), consentRens: r.consent != null,
        enc: eq(r.enceinte, 'Oui'), all: eq(r.allaite, 'Oui'), encRens: r.enceinte != null, allRens: r.allaite != null,
        lotRens: r.lot != null, obsOui: eq(r.obs15, 'Oui'), obsRens: r.obs15 != null, mapi: eq(r.mapi, 'Oui'), mapiGrave: eq(r.mapiGrave, 'Oui'),
        suivi: r.suivis > 0, telRens: !eq(r.tel, 'Absent'), telValide: eq(r.tel, 'Valide'), village: eq(r.village, 'Oui'),
        sexeRens: r.sexe != null, ageRens: r.age != null, idRens: r.id != null, dVacRens: r.dVac != null,
        a1: !r.id || r.id.length !== 33 || r.id.slice(0, 8).toUpperCase() !== 'RDC-HUE-' || r.id.slice(-4).toUpperCase() !== '-VAC' ? 1 : 0,
        a2: r.id && nId[r.id.toUpperCase()] > 1 ? 1 : 0,
        a3: nb > 1 ? 1 : 0,
        a4: vac ? 0 : 1,
        a5: vac && r.dVac == null ? 1 : 0,
        a6: r.dVac != null && r.dEvt != null && r.dVac !== r.dEvt ? 1 : 0,
        a7: d != null && (d < P.DebutCampagne || d > lim) ? 1 : 0,
        a8: d != null && r.dIns != null && d < r.dIns ? 1 : 0,
        a9: delai != null && delai < 0 ? 1 : 0,
        a10: r.sexe == null ? 1 : 0,
        a11: r.age == null || r.age < P.AgeMin || r.age > P.AgeMax ? 1 : 0,
        a12: hom && (eq(r.allaite, 'Oui') || eq(r.enceinte, 'Oui')) ? 1 : 0,
        a13: fem && (r.allaite == null || r.enceinte == null) ? 1 : 0,
        a14: eq(r.consent, 'Oui') ? 0 : 1,
        a15: vac && (r.lot == null || !lots[norm(r.lot)]) ? 1 : 0,
        a16: vac && !eq(r.obs15, 'Oui') ? 1 : 0,
        a17: eq(r.tel, 'Invalide') ? 1 : 0
      };
      a.nbAnom = a.a1 + a.a2 + a.a3 + a.a4 + a.a5 + a.a7 + a.a9 + a.a10 + a.a11 + a.a12 + a.a13 + a.a14 + a.a15 + a.a16 + a.a17 + (P.InclDates ? a.a6 + a.a8 : 0);
      a.detail = ANOM.filter(function (x) { return a[x[0]] && (P.InclDates || (x[0] !== 'a6' && x[0] !== 'a8')); }).map(function (x) { return x[1]; });
      return a;
    });
    return { rows: rows, extrait: extrait, charge: rows.length };
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
  function avg(vals) { return vals.length ? vals.reduce(function (t, v) { return t + v; }, 0) / vals.length : null; }
  function ratio(a, b) { return b ? a / b : null; }

  // Indicateurs détaillés (Tableau de bord!B31:K57) : [libellé, dimension, numérateur, dénominateur, action]
  function indicateurs(ret, filt, P) {
    var R = ret.length, c = function (f) { return cnt(ret, f); };
    var fem = c(function (r) { return r.fem; }), telR = c(function (r) { return r.telRens; });
    var delaiOK = c(function (r) { return r.delai != null; });
    var L = [
      ['ID de vaccination renseigné', 'Complétude', c(function (r) { return r.idRens; }), R, 'Rendre l’ID obligatoire ; compléter les fiches.'],
      ['Date de vaccination (formulaire) renseignée', 'Complétude', c(function (r) { return r.dVacRens; }), R, 'Compléter la date sur la fiche papier / DHIS2.'],
      ['Sexe renseigné', 'Complétude', c(function (r) { return r.sexeRens; }), R, 'Compléter le sexe à partir du registre papier.'],
      ['Âge renseigné', 'Complétude', c(function (r) { return r.ageRens; }), R, 'Compléter l’âge ; vérifier l’unité (ans).'],
      ['Consentement renseigné', 'Complétude', c(function (r) { return r.consentRens; }), R, 'Documenter systématiquement le consentement.'],
      ['Numéro de lot renseigné', 'Complétude', c(function (r) { return r.lotRens; }), R, 'Recopier le lot depuis la fiche de pointage.'],
      ['Observation 15 min renseignée', 'Complétude', c(function (r) { return r.obsRens; }), R, 'Rappeler le champ aux vaccinateurs.'],
      ['Téléphone renseigné', 'Complétude', telR, R, 'Collecter un numéro pour le suivi post-vaccinal.'],
      ['Avenue / village renseigné', 'Complétude', c(function (r) { return r.village; }), R, 'Noter l’adresse pour retrouver la personne.'],
      ['Femmes : statut grossesse et allaitement renseigné', 'Complétude', c(function (r) { return r.fem && r.encRens && r.allRens; }), fem, 'Poser systématiquement les 2 questions aux femmes.'],
      ['Inscrits ayant un événement de vaccination', 'Complétude', cnt(filt, function (r) { return r.vac; }), filt.length, 'Vérifier les inscriptions sans vaccination.'],
      ['Format de l\'ID valide', 'Validité', R - c(function (r) { return r.a1; }), R, 'Corriger l’ID selon le format standard.'],
      ['Numéro de lot conforme à la liste', 'Validité', R - c(function (r) { return r.a15; }), R, 'Corriger les fautes de frappe (O/0, B/8) ; valider la liste.'],
      ['Âge plausible', 'Validité', R - c(function (r) { return r.a11; }), R, 'Vérifier l’âge et l’unité.'],
      ['Téléphone au bon format (parmi renseignés)', 'Validité', c(function (r) { return r.telValide; }), telR, 'Saisir 10 chiffres (08/09…).'],
      ['Date de vaccination = date de l\'événement', 'Cohérence', R - c(function (r) { return r.a6; }), R, 'Saisir la date de l’événement = date de vaccination.'],
      ['Vaccination le jour ou après l\'inscription', 'Cohérence', R - c(function (r) { return r.a8; }), R, 'Vérifier la date d’inscription.'],
      ['Saisie le jour ou après la vaccination', 'Cohérence', R - c(function (r) { return r.a9; }), R, 'Corriger la date de vaccination (future ?).'],
      ['Date dans la période de campagne', 'Cohérence', R - c(function (r) { return r.a7; }), R, 'Corriger les dates erronées.'],
      ['Grossesse / allaitement cohérent avec le sexe', 'Cohérence', R - c(function (r) { return r.a12; }), R, 'Corriger le sexe ou le statut.'],
      ['Consentement = « Oui »', 'Conformité protocole', c(function (r) { return r.consentOui; }), R, 'Investiguer toute vaccination sans consentement.'],
      ['Observation 15 min = « Oui »', 'Conformité protocole', c(function (r) { return r.obsOui; }), R, 'Renforcer la surveillance immédiate.'],
      ['Au moins 1 suivi post-vaccinal saisi', 'Conformité protocole', c(function (r) { return r.suivi; }), R, 'Planifier et saisir les suivis.'],
      ['Fiches sans ID en double', 'Unicité', R - c(function (r) { return r.a2; }), R, 'Corriger les ID dupliqués.'],
      ['Fiches sans doublon probable (nom/sexe/âge)', 'Unicité', R - c(function (r) { return r.a3; }), R, 'Comparer et fusionner/supprimer les doublons.'],
      ['Saisie dans le délai acceptable', 'Promptitude', c(function (r) { return r.delai != null && r.delai >= 0 && r.delai <= P.DelaiMax; }), delaiOK, 'Saisir le jour même (ou J+1).'],
      ['Saisie le jour même', 'Promptitude', c(function (r) { return r.delai === 0; }), delaiOK, 'Organiser la saisie quotidienne.']
    ];
    return L.map(function (x) { return { lib: x[0], dim: x[1], num: x[2], den: x[3], pct: ratio(x[2], x[3]), action: x[4] }; });
  }
  function statut(v, P) { return v == null ? '–' : v >= P.Vert ? 'Bon' : v >= P.Orange ? 'À améliorer' : 'Insuffisant'; }

  /* f = {zone:'' ou nom, aire:'' ou nom, debut, fin, ref} ; P = paramètres */
  function compute(rows, f, params) {
    var P = Object.assign({}, PARAMS_DEFAUT, params || {});
    var inGeo = function (r) { return (!f.zone || eq(r.zone, f.zone)) && (!f.aire || eq(r.aire, f.aire)); };
    var inPer = function (r) { return r.d != null && r.d >= f.debut && r.d <= f.fin; };
    var filt = rows.filter(inGeo);
    var ret = filt.filter(function (r) { return r.vac && inPer(r); });
    var vacF = filt.filter(function (r) { return r.vac && r.d != null; });
    var perim = filt.filter(function (r) { return r.d == null || inPer(r); });
    var R = ret.length;
    var c = function (rs, fn) { return cnt(rs, fn); };
    var dern = 0; vacF.forEach(function (r) { if (r.d <= f.ref && r.d > dern) dern = r.d; });
    var ages = ret.filter(function (r) { return r.age != null; }).map(function (r) { return r.age; });
    var delais = ret.filter(function (r) { return r.delai != null && r.delai >= 0; }).map(function (r) { return r.delai; });
    var k = {
      vaccines: R, hommes: c(ret, function (r) { return r.hom; }), femmes: c(ret, function (r) { return r.fem; }),
      enceintes: c(ret, function (r) { return r.enc; }), allaitantes: c(ret, function (r) { return r.all; }),
      pContact: ratio(c(ret, function (r) { return r.contact; }), R) || 0,
      jour: c(vacF, function (r) { return r.d === f.ref; }), j7: c(vacF, function (r) { return r.d > f.ref - 7 && r.d <= f.ref; }),
      cumul: c(vacF, function (r) { return r.d <= f.ref; }), ageMoyen: avg(ages), dernier: dern || null, joursDepuis: dern ? f.ref - dern : null,
      mapi: c(ret, function (r) { return r.mapi; }), mapiGrave: c(ret, function (r) { return r.mapiGrave; }),
      pObs: ratio(c(ret, function (r) { return r.obsOui; }), R) || 0, pSuivi: ratio(c(ret, function (r) { return r.suivi; }), R) || 0,
      pConsent: ratio(c(ret, function (r) { return r.consentOui; }), R) || 0, sansVacc: c(filt, function (r) { return !r.vac; }),
      avecAnom: c(perim, function (r) { return r.nbAnom > 0; }),
      pSansAnom: ratio(R - c(ret, function (r) { return r.nbAnom > 0; }), R),
      doublons: c(ret, function (r) { return r.a3; }), delaiMoyen: avg(delais), ecartDates: c(ret, function (r) { return r.a6; })
    };
    k.moy7 = k.j7 / 7;
    var ind = indicateurs(ret, filt, P);
    var dims = DIMS.map(function (d) {
      var v = ind.filter(function (i) { return i.dim === d[0] && i.pct != null; }).map(function (i) { return i.pct; });
      return { dim: d[0], quoi: d[1], score: avg(v) };
    });
    var gv = dims.filter(function (d) { return d.score != null; }).map(function (d) { return d.score; });
    k.global = avg(gv); dims.forEach(function (d) { k['dim_' + d.dim] = d.score; });

    // Graphiques (feuille Calculs)
    var jours = [];
    for (var j = 29; j >= 0; j--) { var dj = f.ref - j; jours.push({ date: dj, n: c(vacF, function (r) { return r.d === dj; }), cumul: c(vacF, function (r) { return r.d <= dj; }) }); }
    var ageSexe = TRANCHES.map(function (t) { return { tranche: t, h: c(ret, function (r) { return r.tr === t && r.hom; }), f: c(ret, function (r) { return r.tr === t && r.fem; }) }; });
    var delaisCl = CLASSES.map(function (cl) { return { classe: cl, n: c(ret, function (r) { return r.cl === cl; }) }; });
    var anomTypes = ANOM.map(function (a) { return { type: a[2], n: c(perim, function (r) { return r[a[0]]; }) }; });

    // Situation par zone (toutes les zones) et par aire (zone choisie)
    var ligne = function (rs, lieu) {
      var v = rs.filter(function (r) { return r.vac && r.d != null; }), p = v.filter(inPer), n = p.length;
      var dl = p.filter(function (r) { return r.delai != null && r.delai >= 0; }).map(function (r) { return r.delai; });
      var last = 0; v.forEach(function (r) { if (r.d <= f.ref && r.d > last) last = r.d; });
      return {
        lieu: lieu, periode: n, j7: c(v, function (r) { return r.d > f.ref - 7 && r.d <= f.ref; }), jour: c(v, function (r) { return r.d === f.ref; }),
        cumul: c(v, function (r) { return r.d <= f.ref; }), mapi: c(p, function (r) { return r.mapi; }),
        pSans: n ? 1 - c(p, function (r) { return r.nbAnom > 0; }) / n : null, pLot: n ? 1 - c(p, function (r) { return r.a15; }) / n : null,
        doublons: c(p, function (r) { return r.a3; }), delai: avg(dl), dernier: last || null
      };
    };
    var total = function (list, lieu) {
      var t = { lieu: lieu, periode: 0, j7: 0, jour: 0, cumul: 0, mapi: 0, doublons: 0, delai: null, dernier: null, pSans: null, pLot: null }, ws = 0, wl = 0;
      list.forEach(function (z) {
        t.periode += z.periode; t.j7 += z.j7; t.jour += z.jour; t.cumul += z.cumul; t.mapi += z.mapi; t.doublons += z.doublons;
        if (z.pSans != null) ws += z.periode * z.pSans; if (z.pLot != null) wl += z.periode * z.pLot;
        if (z.dernier && (!t.dernier || z.dernier > t.dernier)) t.dernier = z.dernier;
      });
      if (t.periode) { t.pSans = ws / t.periode; t.pLot = wl / t.periode; }
      return t;
    };
    var zones = zonesOf(rows).map(function (z) { return ligne(rows.filter(function (r) { return eq(r.zone, z); }), z); });
    var zr = rows.filter(function (r) { return !f.zone || eq(r.zone, f.zone); });
    var aires = airesOf(rows, f.zone).map(function (a) { return ligne(zr.filter(function (r) { return eq(r.aire, a); }), a); });

    // Fiches à corriger (feuille Anomalies) et doublons
    var aCorriger = perim.filter(function (r) { return r.nbAnom > 0; })
      .sort(function (a, b) { return (b.d == null ? Infinity : b.d) - (a.d == null ? Infinity : a.d) || b.rang - a.rang; });
    var grp = {};
    rows.forEach(function (r) { if (r.a3 && r.cle) (grp[r.cle] = grp[r.cle] || []).push(r); });
    var doublons = Object.keys(grp).map(function (key) {
      var all = grp[key], m = all.filter(function (r) { return inGeo(r) && (r.d == null || inPer(r)); });
      return { membres: m, total: all.length, last: Math.max.apply(null, m.map(function (r) { return r.d || 0; })) };
    }).filter(function (g) { return g.membres.length; }).sort(function (a, b) { return b.last - a.last; });

    return {
      k: k, indicateurs: ind, dims: dims, jours: jours, ageSexe: ageSexe, delais: delaisCl, anomTypes: anomTypes,
      zones: zones, totalZones: total(zones, 'Total'), aires: aires, totalAires: total(aires, 'Total'),
      aCorriger: aCorriger, doublons: doublons, statut: function (v) { return statut(v, P); }
    };
  }

  var api = { ZONES: ZONES, PARAMS_DEFAUT: PARAMS_DEFAUT, COLS: COLS, ANOM: ANOM, DIMS: DIMS, normalize: normalize, compute: compute,
    zonesOf: zonesOf, airesOf: airesOf, toSerial: toSerial, serialToUTC: serialToUTC, nowSerial: nowSerial, cleNom: cleNom, norm: norm, statut: statut };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.MVEV = api;
})(this);
