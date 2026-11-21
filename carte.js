/* Analyse spatiale de « Suivi MVE » : carte des aires de santé et aires actives.
   - stats()   : indicateurs par aire à partir des lignes calculées par engine.js ;
   - dessiner(): carte SVG (contours DHIS2 du fichier carte.json, sans fond de carte en ligne).
   Aucune donnée de patient n'est dans carte.json : seulement les contours et la population des aires. */
(function (root) {
  'use strict';

  var NS = 'http://www.w3.org/2000/svg';
  function norm(s) {
    return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  }
  function cle(zone, aire) { return norm(zone) + '|' + norm(aire); }

  // Seuils (jours depuis le dernier cas confirmé), comme le tableau « Situation par zone »
  var ACTIVE = 21, FIN = 42;
  var STATUTS = {
    active: { lib: 'Active (< 21 j)', coul: 'var(--bad)' },
    surveillance: { lib: 'Sous surveillance (21 – 41 j)', coul: 'var(--warn)' },
    calme: { lib: 'Sans cas depuis ≥ 42 j', coul: 'var(--ok)' },
    jamais: { lib: 'Aucun cas confirmé', coul: 'var(--c0)' }
  };
  function statutDe(j) { return j == null ? 'jamais' : j < ACTIVE ? 'active' : j < FIN ? 'surveillance' : 'calme'; }

  // Indicateurs de la carte : classes 0 (aucun) à 4
  var INDICATEURS = {
    conf: { lib: 'Cas confirmés', seuils: [1, 5, 15, 50], unite: '', aide: 'Cas confirmés notifiés sur la période choisie.' },
    taux: { lib: 'Taux d’attaque', seuils: [0.0001, 10, 50, 200], unite: ' / 100 000', aide: 'Cas (confirmés + probables) de la période pour 100 000 habitants de l’aire.' },
    dec: { lib: 'Décès', seuils: [1, 3, 10, 25], unite: '', aide: 'Décès parmi les cas (confirmés + probables) de la période.' },
    statut: { lib: 'Transmission', aide: 'Jours depuis le dernier cas confirmé, à la date de référence (toute l’épidémie).' }
  };
  function classe(v, seuils) {
    if (v == null) return null;
    if (!(v > 0)) return 0;
    for (var i = seuils.length - 1; i >= 0; i--) if (v >= seuils[i]) return i + 1;
    return 1;
  }
  function etiquettes(ind) {
    var s = INDICATEURS[ind].seuils, f = function (x) { return String(x).replace('.', ','); };
    if (ind === 'taux') return ['0', '< 10', '10 – 49', '50 – 199', '≥ 200'];
    return ['0', f(s[0]) + ' – ' + f(s[1] - 1), f(s[1]) + ' – ' + f(s[2] - 1), f(s[2]) + ' – ' + f(s[3] - 1), '≥ ' + f(s[3])];
  }

  /* Indicateurs par aire.
     Comptes (cas, confirmés, décès) : filtre période. Statut : dernier cas confirmé jusqu'à la date de référence. */
  function stats(rows, f, carte) {
    var A = {};
    function aire(z, a) {
      var k = cle(z, a);
      return A[k] || (A[k] = { cle: k, zone: z, aire: a, cas: 0, conf: 0, prob: 0, dec: 0, dernierConf: null, pop: null, contour: false });
    }
    if (carte) {
      carte.aires.forEach(function (a) { var o = aire(a.z, a.n); o.pop = a.pop == null ? null : a.pop; o.contour = true; });
      (carte.sansContour || []).forEach(function (a) { var o = aire(a.z, a.n); if (o.pop == null && a.pop != null) o.pop = a.pop; });
    }
    rows.forEach(function (r) {
      if (!r.zone || !r.aire) return;
      var o = aire(r.zone, r.aire);
      if (r.n >= f.debut && r.n <= f.fin && r.cas) { o.cas++; if (r.conf) o.conf++; if (r.prob) o.prob++; if (r.dec) o.dec++; }
      if (r.conf && r.n <= f.ref && (o.dernierConf == null || r.n > o.dernierConf)) o.dernierConf = r.n;
    });
    Object.keys(A).forEach(function (k) {
      var o = A[k];
      o.jours = o.dernierConf == null ? null : f.ref - o.dernierConf;
      o.statut = statutDe(o.jours);
      o.taux = o.pop ? o.cas / o.pop * 1e5 : null;
    });
    return A;
  }

  /* Synthèse « aires actives » (filtre zone ; l'aire choisie n'est pas un filtre ici). */
  function synthese(A, zone) {
    var liste = Object.keys(A).map(function (k) { return A[k]; })
      .filter(function (o) { return !zone || norm(o.zone) === norm(zone); });
    var n = { active: 0, surveillance: 0, calme: 0, jamais: 0, total: 0 };
    var zones = {};
    liste.forEach(function (o) {
      if (!o.contour && o.statut === 'jamais') return; // unité sans contour et sans cas (ex. CTE) : ignorée
      n[o.statut]++; n.total++;
      if (o.statut === 'active') zones[norm(o.zone)] = o.zone;
    });
    var touchees = liste.filter(function (o) { return o.dernierConf != null; })
      .sort(function (a, b) { return a.jours - b.jours || b.conf - a.conf || a.aire.localeCompare(b.aire, 'fr'); });
    return { n: n, zonesActives: Object.keys(zones).map(function (k) { return zones[k]; }).sort(), touchees: touchees };
  }

  /* ---------- Dessin ---------- */
  function el(tag, attrs, parent) {
    var e = document.createElementNS(NS, tag);
    for (var a in attrs) e.setAttribute(a, attrs[a]);
    if (parent) parent.appendChild(e);
    return e;
  }
  function bbox(listes) {
    var b = [Infinity, Infinity, -Infinity, -Infinity];
    listes.forEach(function (g) {
      g.forEach(function (r) {
        for (var i = 0; i < r.length; i += 2) {
          if (r[i] < b[0]) b[0] = r[i]; if (r[i] > b[2]) b[2] = r[i];
          if (r[i + 1] < b[1]) b[1] = r[i + 1]; if (r[i + 1] > b[3]) b[3] = r[i + 1];
        }
      });
    });
    return b;
  }
  // Point d'étiquette : centre de gravité de l'anneau le plus grand
  function centre(g) {
    var best = null, bestA = 0;
    g.forEach(function (r) {
      var a = 0, cx = 0, cy = 0;
      for (var i = 0; i < r.length - 2; i += 2) {
        var x0 = r[i], y0 = r[i + 1], x1 = r[i + 2], y1 = r[i + 3], c = x0 * y1 - x1 * y0;
        a += c; cx += (x0 + x1) * c; cy += (y0 + y1) * c;
      }
      if (Math.abs(a) > bestA) { bestA = Math.abs(a); best = [cx / (3 * a), cy / (3 * a)]; }
    });
    return best;
  }

  /* o = { indicateur, zone, selection (clé), onSelect(cle), hauteurMax } */
  function dessiner(host, carte, A, o) {
    host.replaceChildren();
    var zoneN = o.zone ? norm(o.zone) : null;
    var aires = carte.aires.filter(function (a) { return !zoneN || norm(a.z) === zoneN; });
    if (o.ariaLabel) o._aria = o.ariaLabel;
    if (!aires.length) {
      var p = document.createElement('p'); p.className = 'chart-note';
      p.textContent = 'Aucun contour disponible pour cette zone.'; host.appendChild(p); return null;
    }
    var zones = carte.zones.filter(function (z) { return !zoneN || norm(z.n) === zoneN; });
    var b = bbox(aires.map(function (a) { return a.g; }));
    var kx = Math.cos(((b[1] + b[3]) / 2) * Math.PI / 180);
    var gw = (b[2] - b[0]) * kx, gh = b[3] - b[1], pad = 6;
    var W = Math.max(280, host.clientWidth || 600), hMax = o.hauteurMax || 560;
    var s = Math.min((W - 2 * pad) / gw, (hMax - 2 * pad) / gh);
    var H = Math.round(gh * s + 2 * pad), ox = (W - gw * s) / 2;
    var X = function (lon) { return ox + (lon - b[0]) * kx * s; };
    var Y = function (lat) { return pad + (b[3] - lat) * s; };
    var chemin = function (g) {
      return g.map(function (r) {
        var d = '';
        for (var i = 0; i < r.length; i += 2) d += (i ? 'L' : 'M') + X(r[i]).toFixed(1) + ',' + Y(r[i + 1]).toFixed(1);
        return d + 'Z';
      }).join('');
    };
    var svg = el('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': o._aria || ('Carte des aires de santé – ' + (INDICATEURS[o.indicateur] || { lib: '' }).lib) }, host);
    svg.style.touchAction = 'manipulation';
    var gA = el('g', {}, svg), gZ = el('g', {}, svg), gL = el('g', {}, svg), gS = el('g', {}, svg);
    var ind = INDICATEURS[o.indicateur] || INDICATEURS.conf;
    aires.forEach(function (a) {
      var st = A[cle(a.z, a.n)] || { cas: 0, conf: 0, dec: 0, taux: null, statut: 'jamais', pop: a.pop };
      var fill;
      if (o.couleur) fill = o.couleur(cle(a.z, a.n), st, a);
      else if (o.indicateur === 'statut') fill = STATUTS[st.statut].coul;
      else {
        var v = o.indicateur === 'taux' ? (a.pop ? st.taux || 0 : null) : st[o.indicateur];
        var c = classe(v, ind.seuils);
        fill = c == null ? 'var(--nd)' : 'var(--c' + c + ')';
      }
      var path = el('path', { d: chemin(a.g), 'fill-rule': 'evenodd', class: 'aire', style: 'fill:' + fill }, gA);
      var t = el('title', {}, path); t.textContent = a.n + ' (' + a.z + ')';
      path.addEventListener('click', function () { if (o.onSelect) o.onSelect(cle(a.z, a.n)); });
      if (o.selection && o.selection === cle(a.z, a.n)) {
        el('path', { d: chemin(a.g), 'fill-rule': 'evenodd', class: 'aire-sel' }, gS);
      }
    });
    // Contours de groupes d'aires (foyers) : chaque aire du groupe est entourée
    (o.groupes || []).forEach(function (gr) {
      aires.forEach(function (a) { if (gr.cles[cle(a.z, a.n)]) el('path', { d: chemin(a.g), 'fill-rule': 'evenodd', class: gr.classe || 'foyer' }, gS); });
      if (gr.etiquette) {
        var pts = aires.filter(function (a) { return gr.cles[cle(a.z, a.n)]; }).map(function (a) { return centre(a.g); }).filter(Boolean);
        if (pts.length) {
          var mx = pts.reduce(function (t, p) { return t + p[0]; }, 0) / pts.length, my = pts.reduce(function (t, p) { return t + p[1]; }, 0) / pts.length;
          var t = el('text', { x: X(mx).toFixed(1), y: Y(my).toFixed(1), class: 'flab', 'text-anchor': 'middle', dy: '0.35em' }, gS); t.textContent = gr.etiquette;
        }
      }
    });
    var etiq = [];
    zones.forEach(function (z) {
      el('path', { d: chemin(z.g), 'fill-rule': 'evenodd', class: 'zone' }, gZ);
      if (!zoneN) {
        var c = centre(z.g), bz = bbox([z.g]);
        if (c) etiq.push({ n: z.n, x: X(c[0]), y: Y(c[1]), aire: (bz[2] - bz[0]) * (bz[3] - bz[1]) });
      }
    });
    // Noms de zones : les plus grandes d'abord ; un nom qui en chevaucherait un autre n'est pas écrit
    var poses = [];
    etiq.sort(function (a, b) { return b.aire - a.aire; }).forEach(function (e) {
      var w = e.n.length * 6.4 + 4, h = 13, r = [e.x - w / 2, e.y - h / 2, e.x + w / 2, e.y + h / 2];
      if (poses.some(function (p) { return r[0] < p[2] && r[2] > p[0] && r[1] < p[3] && r[3] > p[1]; })) return;
      poses.push(r);
      var t = el('text', { x: e.x.toFixed(1), y: e.y.toFixed(1), class: 'zlab', 'text-anchor': 'middle', dy: '0.35em' }, gL); t.textContent = e.n;
    });
    return { W: W, H: H };
  }

  function legende(host, ind) {
    host.replaceChildren();
    var items = ind === 'statut'
      ? ['active', 'surveillance', 'calme', 'jamais'].map(function (k) { return [STATUTS[k].coul, STATUTS[k].lib]; })
      : etiquettes(ind).map(function (l, i) { return ['var(--c' + i + ')', l + (i ? INDICATEURS[ind].unite : '')]; });
    if (ind === 'taux') items.push(['var(--nd)', 'Population inconnue']);
    items.forEach(function (it) {
      var s = document.createElement('span'), i = document.createElement('i');
      i.style.background = it[0]; s.append(i, it[1]); host.appendChild(s);
    });
  }


  /* ---------- 3. Propagation dans le temps ----------
     Semaines du lundi au dimanche (date de notification), du premier cas confirmé à la date de référence. */
  function lundi(s) { return s - (((s - 2) % 7) + 7) % 7; } // numéro de série Excel : 2 = lundi 01/01/1900
  function propagation(rows, carte, ref, zone) {
    var zoneN = zone ? norm(zone) : null, parAire = {}, premier = Infinity;
    rows.forEach(function (r) {
      if (!r.conf || !r.zone || !r.aire || r.n > ref) return;
      if (zoneN && norm(r.zone) !== zoneN) return;
      var k = cle(r.zone, r.aire);
      (parAire[k] || (parAire[k] = [])).push(r.n);
      if (r.n < premier) premier = r.n;
    });
    if (premier === Infinity) return { semaines: [], parAire: parAire };
    Object.keys(parAire).forEach(function (k) { parAire[k].sort(function (a, b) { return a - b; }); });
    var semaines = [];
    for (var w = lundi(premier); w <= ref; w += 7) {
      var fin = Math.min(w + 6, ref), o = { semaine: w, fin: fin, cas: 0, airesCas: 0, nouvelles: 0, actives: 0, touchees: 0 };
      Object.keys(parAire).forEach(function (k) {
        var d = parAire[k], n = 0, dernier = null;
        for (var i = 0; i < d.length && d[i] <= fin; i++) { dernier = d[i]; if (d[i] >= w) n++; }
        if (dernier == null) return;
        o.touchees++;
        if (n) { o.cas += n; o.airesCas++; if (d[0] >= w) o.nouvelles++; }
        if (fin - dernier < ACTIVE) o.actives++;
      });
      semaines.push(o);
    }
    return { semaines: semaines, parAire: parAire };
  }
  // État d'une aire à la fin de la semaine w : 'semaine' (cas cette semaine), 'active' (< 21 j), 'touchee', 'jamais'
  function etatSemaine(dates, w, fin) {
    if (!dates) return 'jamais';
    var dernier = null, n = 0;
    for (var i = 0; i < dates.length && dates[i] <= fin; i++) { dernier = dates[i]; if (dates[i] >= w) n++; }
    if (dernier == null) return 'jamais';
    if (n) return dates[0] >= w ? 'nouvelle' : 'semaine';
    return fin - dernier < ACTIVE ? 'active' : 'touchee';
  }
  var ETATS = {
    nouvelle: { lib: 'Touchée pour la 1re fois cette semaine', coul: 'var(--c4)' },
    semaine: { lib: 'Cas cette semaine', coul: 'var(--c3)' },
    active: { lib: 'Cas dans les 21 jours', coul: 'var(--c2)' },
    touchee: { lib: 'Touchée auparavant', coul: 'var(--past)' },
    jamais: { lib: 'Pas encore touchée', coul: 'var(--c0)' }
  };

  /* ---------- 4. Foyers : statistique de balayage spatial de Kulldorff (modèle de Poisson, comme SaTScan) ----------
     Fenêtres : chaque aire et ses voisines les plus proches (centres), jusqu'à maxAires aires et maxPop de la population.
     Significativité : maxRep simulations de Monte-Carlo (cas répartis au hasard selon la population). */
  function alea(graine) { var x = graine >>> 0 || 1; return function () { x ^= x << 13; x ^= x >>> 17; x ^= x << 5; return (x >>> 0) / 4294967296; }; }
  function foyers(carte, cas, opts) {
    opts = opts || {};
    var maxAires = opts.maxAires || 15, maxPop = opts.maxPop || 0.2, R = opts.simulations || 999, minCas = opts.minCas || 3;
    var U = carte.aires.filter(function (a) { return a.pop > 0; }).map(function (a) {
      var c = centre(a.g); return { cle: cle(a.z, a.n), aire: a.n, zone: a.z, pop: a.pop, x: c ? c[0] : 0, y: c ? c[1] : 0, cas: cas[cle(a.z, a.n)] || 0 };
    });
    var hors = Object.keys(cas).filter(function (k) { return cas[k] && !U.some(function (u) { return u.cle === k; }); })
      .reduce(function (t, k) { return t + cas[k]; }, 0);
    var n = U.length, N = U.reduce(function (t, u) { return t + u.cas; }, 0), P = U.reduce(function (t, u) { return t + u.pop; }, 0);
    var res = { N: N, P: P, hors: hors, foyers: [], fenetres: 0, simulations: R };
    if (N < minCas || n < 2) return res;
    var kx = Math.cos((U.reduce(function (t, u) { return t + u.y; }, 0) / n) * Math.PI / 180);
    // Fenêtres = listes d'indices (centre + voisines par distance croissante)
    var fen = [];
    U.forEach(function (u, i) {
      var ord = U.map(function (v, j) { var dx = (v.x - u.x) * kx, dy = v.y - u.y; return [dx * dx + dy * dy, j]; })
        .sort(function (a, b) { return a[0] - b[0]; });
      var liste = [], pop = 0;
      for (var k = 0; k < ord.length && liste.length < maxAires; k++) {
        var j = ord[k][1]; if (pop + U[j].pop > maxPop * P && liste.length) break;
        liste.push(j); pop += U[j].pop; fen.push({ centre: i, idx: liste.slice(), pop: pop });
      }
    });
    res.fenetres = fen.length;
    function llr(O, E) {
      if (O <= E) return 0;
      var a = O * Math.log(O / E), b = N - O > 0 ? (N - O) * Math.log((N - O) / (N - E)) : 0;
      return a + b;
    }
    function balayer(c) {
      var out = new Float64Array(fen.length);
      for (var f = 0; f < fen.length; f++) {
        var O = 0, L = fen[f].idx; for (var q = 0; q < L.length; q++) O += c[L[q]];
        out[f] = llr(O, N * fen[f].pop / P);
      }
      return out;
    }
    var obs = U.map(function (u) { return u.cas; });
    var L0 = balayer(obs);
    // Simulations sous l'hypothèse nulle (répartition des N cas proportionnelle à la population)
    var cum = [], s = 0; U.forEach(function (u) { s += u.pop; cum.push(s / P); });
    var rnd = alea(opts.graine || 20261119), maxSim = new Float64Array(R), c = new Int32Array(n);
    for (var r = 0; r < R; r++) {
      c.fill(0);
      for (var m = 0; m < N; m++) { var x = rnd(), lo = 0, hi = n - 1; while (lo < hi) { var mid = (lo + hi) >> 1; if (cum[mid] < x) lo = mid + 1; else hi = mid; } c[lo]++; }
      var best = 0;
      for (var f = 0; f < fen.length; f++) {
        var O = 0, Lf = fen[f].idx; for (var q = 0; q < Lf.length; q++) O += c[Lf[q]];
        var v = llr(O, N * fen[f].pop / P); if (v > best) best = v;
      }
      maxSim[r] = best;
    }
    // Foyers : fenêtres par LLR décroissant, sans chevauchement (foyers secondaires comme SaTScan)
    var ordre = Array.prototype.map.call(L0, function (v, f) { return [v, f]; }).filter(function (x) { return x[0] > 0; })
      .sort(function (a, b) { return b[0] - a[0]; });
    var pris = new Uint8Array(n);
    for (var t = 0; t < ordre.length && res.foyers.length < (opts.maxFoyers || 5); t++) {
      var F = fen[ordre[t][1]];
      if (F.idx.some(function (i) { return pris[i]; })) continue;
      var Obs = F.idx.reduce(function (a, i) { return a + U[i].cas; }, 0);
      if (Obs < minCas) continue;
      var rang = 1; for (var r2 = 0; r2 < R; r2++) if (maxSim[r2] >= ordre[t][0]) rang++;
      var pval = rang / (R + 1);
      if (pval > (opts.seuil || 0.05)) break; // les suivants ont un LLR plus faible : non significatifs
      F.idx.forEach(function (i) { pris[i] = 1; });
      var E = N * F.pop / P;
      res.foyers.push({
        aires: F.idx.map(function (i) { return U[i]; }), centre: U[F.centre], cas: Obs, attendus: E, pop: F.pop,
        rr: (Obs / E) / ((N - Obs) / (N - E || 1)), llr: ordre[t][0], p: pval,
        cles: F.idx.reduce(function (o, i) { o[U[i].cle] = 1; return o; }, {})
      });
    }
    return res;
  }

  var api = { propagation: propagation, etatSemaine: etatSemaine, ETATS: ETATS, lundi: lundi, foyers: foyers, stats: stats, synthese: synthese, dessiner: dessiner, legende: legende, cle: cle, norm: norm,
    INDICATEURS: INDICATEURS, STATUTS: STATUTS, statutDe: statutDe, ACTIVE: ACTIVE, FIN: FIN };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MVECarte = api;
})(this);
