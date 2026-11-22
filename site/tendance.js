/* Onglet « Tendance » de Suivi MVE : transmission (Rt), projection à court terme et compte à rebours des 42 jours.
   Calculs faits sur le téléphone, à partir de la liste linéaire déjà chargée (aucune donnée supplémentaire).

   1. Rt (nombre de reproduction effectif) : méthode de Cori et al. (Am J Epidemiol 2013, package EpiEstim).
      Incidence quotidienne des cas (confirmés + probables) par date de début des symptômes (date de notification
      si absente). Intervalle sériel d'Ebola : loi gamma de moyenne 15,3 j et d'écart-type 9,3 j (WHO Ebola Response
      Team, NEJM 2014). Fenêtre glissante de 14 jours, a priori gamma (moyenne 5, écart-type 5).
      Les derniers jours, incomplets (délai entre début des symptômes et notification), sont exclus : l'estimation
      s'arrête à la date de référence moins le délai de notification du 3e quartile observé (entre 3 et 14 jours).
   2. Projection : processus de branchement de Poisson (comme le package « projections ») ; 1 000 simulations,
      Rt tiré dans sa loi a posteriori et supposé constant sur l'horizon.
   3. Fin de l'épidémie : 42 jours (deux périodes d'incubation maximales) sans nouveau cas confirmé. */
(function (root) {
  'use strict';

  var SI_MOY = 15.3, SI_ET = 9.3, SI_MAX = 80, FENETRE = 14, A_PRIOR = 1, B_PRIOR = 5, NSIM = 1000;

  /* ---------- Lois et fonctions spéciales ---------- */
  function lgamma(x) { // Lanczos
    var g = 7, c = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
      12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
    if (x < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * x)) - lgamma(1 - x);
    x -= 1; var a = c[0], t = x + g + 0.5;
    for (var i = 1; i < 9; i++) a += c[i] / (x + i);
    return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a);
  }
  function gammaP(a, x) { // fonction gamma incomplète régularisée P(a, x)
    if (x <= 0) return 0;
    if (x < a + 1) {
      var s = 1 / a, t = s;
      for (var n = 1; n < 500; n++) { t *= x / (a + n); s += t; if (Math.abs(t) < Math.abs(s) * 1e-12) break; }
      return s * Math.exp(-x + a * Math.log(x) - lgamma(a));
    }
    var b = x + 1 - a, c = 1e300, d = 1 / b, h = d;
    for (var i = 1; i < 500; i++) {
      var an = -i * (i - a); b += 2; d = an * d + b; if (Math.abs(d) < 1e-300) d = 1e-300;
      c = b + an / c; if (Math.abs(c) < 1e-300) c = 1e-300; d = 1 / d; var del = d * c; h *= del; if (Math.abs(del - 1) < 1e-12) break;
    }
    return 1 - Math.exp(-x + a * Math.log(x) - lgamma(a)) * h;
  }
  function gammaQuantile(p, forme, echelle) {
    var lo = 0, hi = Math.max(1, forme * 10);
    while (gammaP(forme, hi) < p) hi *= 2;
    for (var i = 0; i < 80; i++) { var m = (lo + hi) / 2; if (gammaP(forme, m) < p) lo = m; else hi = m; }
    return (lo + hi) / 2 * echelle;
  }
  /* Intervalle sériel discrétisé : w(k) = P(k − 0,5 < X ≤ k + 0,5), w(0) = 0, puis normalisé (moyenne conservée) */
  function intervalleSeriel() {
    var forme = Math.pow(SI_MOY / SI_ET, 2), ech = SI_ET * SI_ET / SI_MOY, w = [0];
    var F = function (x) { return x <= 0 ? 0 : gammaP(forme, x / ech); };
    for (var k = 1; k <= SI_MAX; k++) w.push(Math.max(0, F(k + 0.5) - F(k - 0.5)));
    var s = w.reduce(function (t, v) { return t + v; }, 0);
    return w.map(function (v) { return v / s; });
  }
  var W = intervalleSeriel();

  /* ---------- Générateur aléatoire reproductible ---------- */
  function alea(graine) { var x = graine >>> 0 || 1; return function () { x ^= x << 13; x ^= x >>> 17; x ^= x << 5; return ((x >>> 0) + 0.5) / 4294967296; }; }
  function normale(u) { return Math.sqrt(-2 * Math.log(u())) * Math.cos(2 * Math.PI * u()); }
  function tirageGamma(forme, u) { // Marsaglia et Tsang (échelle 1)
    if (forme < 1) return tirageGamma(forme + 1, u) * Math.pow(u(), 1 / forme);
    var d = forme - 1 / 3, c = 1 / Math.sqrt(9 * d);
    for (;;) {
      var x = normale(u), v = 1 + c * x; if (v <= 0) continue;
      v = v * v * v; var uu = u();
      if (uu < 1 - 0.0331 * x * x * x * x || Math.log(uu) < 0.5 * x * x + d * (1 - v + Math.log(v))) return d * v;
    }
  }
  function tiragePoisson(l, u) {
    if (l <= 0) return 0;
    if (l > 40) return Math.max(0, Math.round(l + Math.sqrt(l) * normale(u)));
    var L = Math.exp(-l), k = 0, p = 1;
    do { k++; p *= u(); } while (p > L);
    return k - 1;
  }

  /* ---------- Séries ---------- */
  function lambda(inc, t) { var s = 0; for (var k = 1; k <= SI_MAX && t - k >= 0; k++) s += inc[t - k] * W[k]; return s; }
  function qr(arr, p) { return Math.round(quantile(arr, p)); }
  function quantile(arr, p) { var a = arr.slice().sort(function (x, y) { return x - y; }); var i = (a.length - 1) * p, lo = Math.floor(i); return a[lo] + (a[Math.min(lo + 1, a.length - 1)] - a[lo]) * (i - lo); }

  /* Délai début des symptômes → notification (3e quartile), borné à [3, 14] jours */
  function delaiNotification(rows) {
    var d = rows.filter(function (r) { return r.cas && r.delai != null && r.delai >= 0 && r.delai <= 60; }).map(function (r) { return r.delai; });
    if (d.length < 5) return 7;
    return Math.max(3, Math.min(14, Math.round(quantile(d, 0.75))));
  }

  /* Rt sur la fenêtre [t − 13, t] (a posteriori gamma : forme a, échelle b) */
  function rtA(inc, lam, t) {
    var sI = 0, sL = 0;
    for (var s = t - FENETRE + 1; s <= t; s++) { if (s < 0) continue; sI += inc[s]; sL += lam[s]; }
    var a = A_PRIOR + sI, b = 1 / (1 / B_PRIOR + sL);
    return { forme: a, echelle: b, cas: sI, moy: a * b, bas: gammaQuantile(0.025, a, b), haut: gammaQuantile(0.975, a, b),
      pSup1: 1 - gammaP(a, 1 / b), fiable: sI >= 6 && sL > 1 };
  }

  /* Analyse complète pour un ensemble de lignes (déjà limité à la zone choisie) */
  function analyser(rows, ref, opts) {
    opts = opts || {};
    var cas = rows.filter(function (r) { return r.cas && r.j != null && r.j <= ref; });
    var res = { ref: ref, nbCas: cas.length, delai: delaiNotification(rows), si: { moy: SI_MOY, et: SI_ET }, fenetre: FENETRE };
    var conf = rows.filter(function (r) { return r.conf && r.n != null && r.n <= ref; });
    res.dernierConf = conf.length ? Math.max.apply(null, conf.map(function (r) { return r.n; })) : null;
    res.joursSans = res.dernierConf == null ? null : ref - res.dernierConf;
    res.finPossible = res.dernierConf == null ? null : res.dernierConf + 42;
    if (!cas.length) { res.vide = true; return res; }
    var d0 = Math.min.apply(null, cas.map(function (r) { return r.j; })), T = ref - res.delai;
    res.debut = d0; res.T = T;
    var n = ref - d0 + 1, inc = new Array(n).fill(0);
    cas.forEach(function (r) { inc[r.j - d0]++; });
    res.inc = inc;
    var iT = T - d0;
    if (iT < FENETRE) { res.court = true; return res; }
    var lam = inc.map(function (v, t) { return lambda(inc, t); });
    // Rt jour par jour
    res.serie = [];
    for (var t = FENETRE; t <= iT; t++) { var e = rtA(inc, lam, t); e.jour = d0 + t; res.serie.push(e); }
    res.rt = res.serie[res.serie.length - 1];
    if (opts.sansProjection) return res;
    // Projection : jours iT+1 … (ref + horizon), Rt tiré dans la loi a posteriori de la dernière fenêtre
    var lundi = function (s) { return s - ((((new Date(Date.UTC(1899, 11, 30) + s * 864e5)).getUTCDay() + 6) % 7)); };
    var L0 = lundi(ref), H = opts.horizon || 28;
    var fin = Math.max(ref + H, L0 + 34) - d0, u = alea(opts.graine || 20261122);
    var nJ = fin - iT, sims = [];
    for (var k = 0; k < NSIM; k++) {
      var R = tirageGamma(res.rt.forme, u) * res.rt.echelle, x = inc.slice(0, iT + 1), out = [];
      for (var t2 = iT + 1; t2 <= fin; t2++) { var v = tiragePoisson(R * lambda(x, t2), u); x.push(v); out.push(v); }
      sims.push(out);
    }
    // Totaux sur les périodes après la date de référence (par date de début des symptômes)
    var off = (ref - d0) - iT; // index, dans « out », du jour de référence
    function total(a, b) { var v = sims.map(function (o) { var s = 0; for (var i = off + a - 1; i <= off + b - 1; i++) s += o[i]; return s; }); return { med: qr(v, 0.5), b50: qr(v, 0.25), h50: qr(v, 0.75), b95: qr(v, 0.025), h95: qr(v, 0.975), p0: v.filter(function (x) { return x === 0; }).length / v.length }; }
    res.proj14 = total(1, 14); res.proj28 = total(1, 28);
    // Semaines (lundi → dimanche) : 8 semaines observées + 4 semaines projetées
    var semaines = [];
    for (var w = -8; w <= 4; w++) {
      var a = L0 + 7 * w, b = a + 6, obs = 0;
      for (var dd = a; dd <= Math.min(b, ref); dd++) if (dd >= d0) obs += inc[dd - d0];
      var simW = null;
      if (b > T) { // semaine incomplète ou future : distribution simulée (jours observés fiables + jours simulés)
        var v = sims.map(function (o) {
          var s = 0;
          for (var d2 = a; d2 <= b; d2++) { if (d2 <= T) { if (d2 >= d0) s += inc[d2 - d0]; } else s += o[d2 - d0 - iT - 1]; }
          return s;
        });
        simW = { med: qr(v, 0.5), b50: qr(v, 0.25), h50: qr(v, 0.75), b95: qr(v, 0.025), h95: qr(v, 0.975) };
      }
      semaines.push({ debut: a, fin: b, obs: a <= ref ? obs : null, future: a > ref, partielle: a <= ref && b > T, sim: simW });
    }
    res.semaines = semaines;
    return res;
  }

  function lecture(rt) {
    if (!rt) return { classe: '', texte: 'Données insuffisantes' };
    if (!rt.fiable) return { classe: 'mid', texte: 'Trop peu de cas pour conclure' };
    if (rt.haut < 1) return { classe: 'good', texte: 'En recul (Rt < 1)' };
    if (rt.bas > 1) return { classe: 'bad', texte: 'En progression (Rt > 1)' };
    return { classe: 'mid', texte: 'Incertain (l’intervalle contient 1)' };
  }

  var api = { analyser: analyser, lecture: lecture, intervalleSeriel: intervalleSeriel, gammaQuantile: gammaQuantile, gammaP: gammaP, W: W,
    PARAMS: { SI_MOY: SI_MOY, SI_ET: SI_ET, FENETRE: FENETRE, NSIM: NSIM } };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.MVETendance = api;
})(this);
