/* Moteur « Chaînes de transmission MVE ».
   Relie les cas sources à leurs contacts (N° Epid du cas source saisi sur la fiche contact),
   puis les contacts devenus cas à leur propre alerte, par rapprochement du nom, du sexe, de l'âge
   et des dates (les noms servent seulement au calcul, en mémoire : ils ne sont ni affichés ni gardés).
   Sources : classeur « Qualité des données MVE » (alertes, feuille Données) et classeur
   « Contacts MVE » (feuille Données, colonnes de noms comprises). Dates = numéros de série Excel. */
(function (root) {
  'use strict';
  var EPOCH = Date.UTC(1899, 11, 30);
  function norm(s) { return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase(); }
  function blank(v) { return v == null || String(v).trim() === ''; }
  function num(v) { return typeof v === 'number' && !isNaN(v) ? v : (typeof v === 'string' && v.trim() !== '' && !isNaN(+v) ? +v : null); }
  function toSerial(v) {
    if (blank(v)) return null;
    if (typeof v === 'number') return Math.floor(v + 1e-9);
    if (v instanceof Date) return Math.floor((v.getTime() - v.getTimezoneOffset() * 60000 - EPOCH) / 864e5);
    var m = String(v).trim().match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return (Date.UTC(+m[1], +m[2] - 1, +m[3]) - EPOCH) / 864e5;
    m = String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return (Date.UTC(+m[3], +m[2] - 1, +m[1]) - EPOCH) / 864e5;
    return null;
  }
  function serialToUTC(s) { return new Date(EPOCH + s * 864e5); }
  function mots(n) { // mots du nom (lettres seules, au moins 2 lettres)
    if (blank(n)) return [];
    var s = String(n).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z ]/g, ' ');
    var vu = {}; return s.split(' ').filter(function (w) { if (w.length < 2 || vu[w]) return false; vu[w] = 1; return true; });
  }
  function epid(v) { return blank(v) ? null : String(v).trim().toUpperCase(); }
  function court(id) { // RDC-HUE-ISI-26-0098545-MVE → ISI-0098545
    var m = /^RDC-HUE-([A-Z]{3})-\d{2}-(\d+)-MVE$/.exec(id || ''); return m ? m[1] + '-' + m[2] : (id || '?');
  }
  function colonnes(table, defs) {
    var h = (table[0] || []).map(norm), idx = {};
    Object.keys(defs).forEach(function (k) {
      var d = [].concat(defs[k]).map(norm);
      idx[k] = h.findIndex(function (x) { return d.some(function (y) { return x === y || (y.charAt(0) === '~' && x.indexOf(y.slice(1)) >= 0); }); });
    });
    return idx;
  }
  var ALERTES = { id: 'ID alerte', dn: 'Date notification', zone: 'Zone de santé', aire: 'Aire de santé', sexe: 'Sexe', age: 'Âge (ans)', ageM: 'Âge (mois)',
    classif: 'Classification finale', statut: 'Statut final', deces: 'Date décès', casConf: 'Cas confirmés', nom: ['Nom', '~nom du cas'] };
  var CONTACTS = { enr: 'Enrollment', id: 'ID contact', src: 'N° Epid cas source', dEnr: 'Date enregistrement', dDern: 'Date dernier contact',
    zone: 'Zone de santé', aire: 'Aire de santé', sexe: 'Sexe', age: 'Âge (ans)', issue: 'Issue', dIss: 'Date issue', nbJ: 'Jours de suivi saisis',
    nom: '~nom et prenom du contact' };

  function charger(alertesT, contactsT) {
    var ia = colonnes(alertesT, ALERTES), ic = colonnes(contactsT, CONTACTS);
    if (ia.id < 0 || ia.classif < 0) throw new Error('Classeur des alertes non reconnu (colonnes « ID alerte » et « Classification finale » attendues).');
    if (ia.nom < 0) throw new Error('Le classeur des alertes n’a pas de colonne « Nom » : utilisez le classeur « Qualité des données MVE ».');
    if (ic.src < 0 || ic.dEnr < 0) throw new Error('Classeur des contacts non reconnu (colonnes « N° Epid cas source » et « Date enregistrement » attendues).');
    var alertes = {}, cas = [], parMot = {};
    alertesT.slice(1).forEach(function (r) {
      var id = epid(r[ia.id]); if (!id || alertes[id]) return;
      var cl = norm(r[ia.classif]), age = num(r[ia.age]);
      if (age == null && ia.ageM >= 0 && num(r[ia.ageM]) != null) age = num(r[ia.ageM]) / 12;
      var a = { id: id, court: court(id), dn: toSerial(r[ia.dn]), zone: blank(r[ia.zone]) ? null : String(r[ia.zone]).trim(), aire: ia.aire >= 0 && !blank(r[ia.aire]) ? String(r[ia.aire]).trim() : null,
        sexe: norm(r[ia.sexe]), age: age,
        // Comme « Suivi MVE » : confirmé = indicateur DHIS2 « Cas confirmés » = 1 (à défaut, la classification) ; probable = classification « Cas probable »
        cas: (ia.casConf >= 0 ? num(r[ia.casConf]) === 1 : cl.indexOf('confirm') >= 0) ? 'Confirmé' : cl.indexOf('probable') >= 0 ? 'Probable' : null,
        classif: blank(r[ia.classif]) ? 'Non classée' : String(r[ia.classif]).trim(),
        decede: (ia.deces >= 0 && !blank(r[ia.deces])) || /dece|mort/.test(norm(r[ia.statut])) };
      alertes[id] = a;
      if (a.cas) { cas.push(a); a.mots = mots(r[ia.nom]); a.mots.forEach(function (w) { (parMot[w] = parMot[w] || []).push(a); }); }
    });
    var contacts = [];
    contactsT.slice(1).forEach(function (r, i) {
      if (blank(r[ic.enr >= 0 ? ic.enr : ic.src]) && blank(r[ic.src])) return;
      contacts.push({ rang: i, id: ic.id >= 0 && !blank(r[ic.id]) ? String(r[ic.id]).trim() : null, src: epid(r[ic.src]), dEnr: toSerial(r[ic.dEnr]),
        dDern: ic.dDern >= 0 ? toSerial(r[ic.dDern]) : null, zone: ic.zone >= 0 && !blank(r[ic.zone]) ? String(r[ic.zone]).trim() : null,
        aire: ic.aire >= 0 && !blank(r[ic.aire]) ? String(r[ic.aire]).trim() : null, sexe: ic.sexe >= 0 ? norm(r[ic.sexe]) : '', age: ic.age >= 0 ? num(r[ic.age]) : null,
        issue: ic.issue >= 0 && !blank(r[ic.issue]) ? String(r[ic.issue]).trim() : null, vu: ic.nbJ >= 0 && (num(r[ic.nbJ]) || 0) > 0,
        mots: ic.nom >= 0 ? mots(r[ic.nom]) : [] });
    });
    // Rapprochement contact → alerte (cas confirmé ou probable)
    contacts.forEach(function (c) {
      var M = c.mots; c.lien = null;
      if (M.length < 2) { delete c.mots; return; }
      var vus = {}, meilleur = null;
      M.forEach(function (w) { (parMot[w] || []).forEach(function (a) { vus[a.id] = a; }); });
      Object.keys(vus).forEach(function (k) {
        var a = vus[k]; if (a.id === c.src) return;
        var inter = M.filter(function (w) { return a.mots.indexOf(w) >= 0; }).length, sc = inter / Math.max(M.length, a.mots.length);
        if (inter < 2 || sc < 0.66) return;
        if (c.sexe && a.sexe && c.sexe !== a.sexe) return;
        if (c.age != null && a.age != null && Math.abs(c.age - a.age) > 3) return;
        var d0 = c.dDern != null ? c.dDern : c.dEnr, d1 = c.dEnr != null ? c.dEnr : c.dDern;
        if (a.dn != null && d0 != null && (a.dn < d0 - 3 || a.dn > d1 + 42)) return;
        var devenu = /confirm|suspect/i.test(c.issue || ''), memeZone = c.zone && a.zone && norm(c.zone) === norm(a.zone);
        var niveau = sc === 1 || (devenu && sc >= 0.66) ? 'fort' : memeZone ? 'probable' : null;
        if (!niveau) return;
        var ecart = a.dn != null && d0 != null ? Math.abs(a.dn - d0) : 99;
        if (!meilleur || sc > meilleur.sc || (sc === meilleur.sc && ecart < meilleur.ecart)) meilleur = { a: a, sc: sc, niveau: niveau, ecart: ecart };
      });
      if (meilleur) c.lien = { cible: meilleur.a.id, niveau: meilleur.niveau, score: meilleur.sc };
      delete c.mots;
    });
    cas.forEach(function (a) { delete a.mots; });
    return construire(alertes, cas, contacts);
  }

  function construire(alertes, cas, contacts) {
    var noeuds = {};
    function noeud(id) {
      if (!noeuds[id]) {
        var a = alertes[id];
        noeuds[id] = { id: id, court: court(id), alerte: a || null, cas: a ? a.cas : null, dn: a ? a.dn : null, zone: a ? a.zone : null, aire: a ? a.aire : null,
          decede: a ? a.decede : false, classif: a ? a.classif : 'Hors liste des alertes', contacts: [], enfants: [], parent: null, lienParent: null };
      }
      return noeuds[id];
    }
    contacts.forEach(function (c) { if (c.src) noeud(c.src).contacts.push(c); });
    // Arêtes cas source → cas secondaire (le plus fort lien par couple)
    var aretes = {};
    contacts.forEach(function (c) {
      if (!c.src || !c.lien) return;
      var k = c.src + '>' + c.lien.cible, ex = aretes[k];
      if (!ex || (ex.niveau === 'probable' && c.lien.niveau === 'fort')) aretes[k] = { de: c.src, vers: c.lien.cible, niveau: c.lien.niveau, contact: c };
    });
    var liste = Object.keys(aretes).map(function (k) { return aretes[k]; }).sort(function (x, y) { return (x.niveau === 'fort' ? 0 : 1) - (y.niveau === 'fort' ? 0 : 1); });
    function ancetre(n, cible) { for (var p = n; p; p = p.parent) if (p === cible) return true; return false; }
    liste.forEach(function (e) {
      var de = noeud(e.de), vers = noeud(e.vers);
      if (vers.parent || de === vers || ancetre(de, vers)) { e.ignore = true; return; } // un seul cas source par cas ; pas de boucle
      vers.parent = de; vers.lienParent = e; de.enfants.push(vers);
    });
    // Chaînes : arbres à partir des racines
    var chaines = [];
    Object.keys(noeuds).forEach(function (k) {
      var r = noeuds[k]; if (r.parent) return;
      var membres = [], gen = 0;
      (function parcours(n, g) { n.gen = g; membres.push(n); if (g > gen) gen = g; n.enfants.sort(function (a, b) { return (a.dn || 0) - (b.dn || 0); }).forEach(function (e) { parcours(e, g + 1); }); })(r, 0);
      var cs = [], dates = [], zones = {};
      membres.forEach(function (m) { m.contacts.forEach(function (c) { cs.push(c); }); if (m.dn != null) dates.push(m.dn); if (m.zone) zones[m.zone] = 1; });
      chaines.push({
        racine: r, membres: membres, generations: gen + 1, nbCas: membres.filter(function (m) { return m.cas; }).length, nbSecondaires: membres.length - 1,
        contacts: cs.length, vus: cs.filter(function (c) { return c.vu; }).length, devenus: cs.filter(function (c) { return /confirm/i.test(c.issue || ''); }).length,
        debut: dates.length ? Math.min.apply(null, dates) : null, fin: dates.length ? Math.max.apply(null, dates) : null, zones: Object.keys(zones).sort(),
        liensProbables: membres.filter(function (m) { return m.lienParent && m.lienParent.niveau === 'probable'; }).length
      });
    });
    chaines.sort(function (a, b) { return b.generations - a.generations || b.nbSecondaires - a.nbSecondaires || b.contacts - a.contacts; });
    chaines.forEach(function (ch, i) { ch.num = i + 1; ch.membres.forEach(function (m) { m.chaine = ch; }); });
    return { alertes: alertes, cas: cas, contacts: contacts, noeuds: noeuds, aretes: liste, chaines: chaines };
  }

  /* f = { zone, debut, fin } : chaînes ayant au moins un cas dans la zone, première date dans la période */
  function synthese(D, f) {
    var dans = function (ch) {
      if (f.zone && !ch.membres.some(function (m) { return m.zone && norm(m.zone) === norm(f.zone); })) return false;
      var d = ch.debut; if (f.debut != null && (d == null || d < f.debut)) return false; if (f.fin != null && (d == null || d > f.fin)) return false;
      return true;
    };
    var chaines = D.chaines.filter(function (ch) { return ch.contacts > 0 || ch.nbSecondaires > 0; }).filter(dans);
    var membres = []; chaines.forEach(function (ch) { ch.membres.forEach(function (m) { membres.push(m); }); });
    var casZ = D.cas.filter(function (a) { return (!f.zone || (a.zone && norm(a.zone) === norm(f.zone))) && (f.debut == null || (a.dn != null && a.dn >= f.debut)) && (f.fin == null || (a.dn != null && a.dn <= f.fin)); });
    var avecContacts = casZ.filter(function (a) { return D.noeuds[a.id] && D.noeuds[a.id].contacts.length; }).length;
    var relies = casZ.filter(function (a) { var n = D.noeuds[a.id]; return n && (n.parent || n.enfants.length); }).length;
    var contacts = 0, devenus = 0, devenusSansAlerte = [];
    chaines.forEach(function (ch) { contacts += ch.contacts; devenus += ch.devenus; ch.membres.forEach(function (m) { m.contacts.forEach(function (c) { if (/confirm/i.test(c.issue || '') && !c.lien) devenusSansAlerte.push({ c: c, source: m }); }); }); });
    var liens = membres.filter(function (m) { return m.lienParent; });
    var k = {
      chaines: chaines.length, multi: chaines.filter(function (ch) { return ch.generations >= 2; }).length,
      maxGen: chaines.reduce(function (t, ch) { return Math.max(t, ch.generations); }, 0),
      sources: membres.filter(function (m) { return m.contacts.length; }).length, contacts: contacts,
      parSource: null, devenus: devenus, liensForts: liens.filter(function (m) { return m.lienParent.niveau === 'fort'; }).length,
      liensProbables: liens.filter(function (m) { return m.lienParent.niveau === 'probable'; }).length,
      casPeriode: casZ.length, casConf: casZ.filter(function (a) { return a.cas === 'Confirmé'; }).length, casProb: casZ.filter(function (a) { return a.cas === 'Probable'; }).length, pAvecContacts: casZ.length ? avecContacts / casZ.length : null, sansContacts: casZ.length - avecContacts,
      pRelies: casZ.length ? relies / casZ.length : null, horsListe: membres.filter(function (m) { return !m.alerte && m.contacts.length; }).length
    };
    k.parSource = k.sources ? contacts / k.sources : null;
    var aValider = liens.filter(function (m) { return m.lienParent.niveau === 'probable'; }).map(function (m) { return m.lienParent; })
      .concat(liens.filter(function (m) { return m.lienParent.niveau === 'fort'; }).map(function (m) { return m.lienParent; }));
    var sansContactsListe = casZ.filter(function (a) { return !(D.noeuds[a.id] && D.noeuds[a.id].contacts.length); }).sort(function (a, b) { return (b.dn || 0) - (a.dn || 0); });
    return { k: k, chaines: chaines, aValider: aValider, devenusSansAlerte: devenusSansAlerte, sansContacts: sansContactsListe };
  }


  /* Données publiées pour les téléphones : aucun nom, seulement les liens déjà calculés */
  function exporter(D) {
    var ids = {};
    D.cas.forEach(function (a) { ids[a.id] = 1; });
    D.contacts.forEach(function (c) { if (c.src) ids[c.src] = 1; if (c.lien) ids[c.lien.cible] = 1; });
    return {
      alertes: Object.keys(ids).filter(function (k) { return D.alertes[k]; }).map(function (k) { var a = D.alertes[k]; return [a.id, a.dn, a.zone, a.aire, a.cas, a.classif, a.decede ? 1 : 0]; }),
      contacts: D.contacts.filter(function (c) { return c.src || c.lien; }).map(function (c) {
        return [c.id, c.src, c.dEnr, c.dDern, c.zone, c.issue, c.vu ? 1 : 0, c.lien ? c.lien.cible : null, c.lien ? c.lien.niveau : null];
      })
    };
  }
  function assembler(p) {
    var alertes = {}, cas = [];
    (p.alertes || []).forEach(function (r) {
      var a = { id: r[0], court: court(r[0]), dn: r[1], zone: r[2], aire: r[3], cas: r[4], classif: r[5], decede: !!r[6] };
      alertes[a.id] = a; if (a.cas) cas.push(a);
    });
    var contacts = (p.contacts || []).map(function (r, i) {
      return { rang: i, id: r[0], src: r[1], dEnr: r[2], dDern: r[3], zone: r[4], issue: r[5], vu: !!r[6], lien: r[7] ? { cible: r[7], niveau: r[8] } : null };
    });
    return construire(alertes, cas, contacts);
  }

  var api = { exporter: exporter, assembler: assembler, charger: charger, synthese: synthese, toSerial: toSerial, serialToUTC: serialToUTC, court: court, norm: norm };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.MVECH = api;
})(this);
