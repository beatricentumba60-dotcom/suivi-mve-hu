#!/usr/bin/env node
/* =====================================================================
   Suivi MVE – publication des données pour les téléphones
   ---------------------------------------------------------------------
   1. Lit la feuille « Données » du classeur Excel actualisé.
   2. Garde seulement les colonnes utiles au tableau de bord (sans l'ID
      de l'alerte), compresse et CHIFFRE le tout (AES-256-GCM, clé tirée
      du mot de passe d'équipe par PBKDF2-SHA256, 600 000 itérations).
   3. Publie l'application + le fichier chiffré « donnees.enc » sur le
      site Netlify. Les téléphones le récupèrent à l'ouverture.

   Utilisation :  node publier.js            (publication)
                  node publier.js --config   (refaire la configuration)
                  node publier.js --test     (prépare sans publier)
   Nécessite Node.js 18 ou plus récent. Aucune autre installation.
   ===================================================================== */
'use strict';
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const crypto = require('crypto');
const readline = require('readline');

const ICI = __dirname;
const SITE = path.join(ICI, '..', 'site');
const CONFIG = path.join(ICI, 'config.json');
const XLSX = require(path.join(SITE, 'xlsx.mini.min.js'));
const ITERATIONS = 600000;

// Colonnes transmises aux téléphones (celles qu'utilise le tableau de bord)
const COLONNES = ['Date notification', 'Date début symptômes', 'Zone de santé', 'Aire de santé', 'Sexe',
  'Âge (ans)', 'Âge (mois)', 'Nature alerte', 'Conclusion alerte', 'Classification finale', 'Statut final',
  'Date décès', 'Statut avant admission', 'Contact avec malade Ebola', 'Date prélèvement', 'Date analyse labo',
  'Date admission CTE', 'Date sortie CTE', 'Vacciné MVE', 'Extrait le', 'Allaite', 'Enceinte', 'Cas confirmés', 'Profession',
  'Décès confirmés (DHIS2)', 'Non cas (DHIS2)'];

const args = process.argv.slice(2);
const MODE_TEST = args.includes('--test');

function norm(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
}
function log(m) { console.log(m); }
function stop(m) { console.error('\n✗ ' + m + '\n'); process.exit(1); }

/* ---------- Questions dans la console ---------- */
function ask(q, { hidden = false } = {}) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout, terminal: true });
    if (hidden) {
      rl._writeToOutput = function (s) { if (s.includes(q)) rl.output.write(s); else rl.output.write('*'); };
    }
    rl.question(q, (a) => { rl.close(); if (hidden) process.stdout.write('\n'); resolve(a.trim()); });
  });
}

/* ---------- Netlify (API publique, jeton personnel) ---------- */
async function netlify(token, method, url, body, type) {
  const res = await fetch('https://api.netlify.com/api/v1' + url, {
    method,
    headers: Object.assign({ Authorization: 'Bearer ' + token }, type ? { 'Content-Type': type } : {}),
    body
  });
  const text = await res.text();
  let json = null; try { json = JSON.parse(text); } catch (e) { /* réponse non JSON */ }
  if (!res.ok) {
    const msg = (json && (json.message || json.errors && JSON.stringify(json.errors))) || text.slice(0, 200);
    const e = new Error('Netlify a répondu ' + res.status + ' : ' + msg); e.status = res.status; throw e;
  }
  return json;
}

/* ---------- GitHub Pages (API REST, jeton personnel « fine-grained ») ---------- */
async function github(token, method, url, body) {
  // Connexion parfois instable : jusqu'à 5 essais (coupure réseau, délai dépassé ou erreur 5xx de GitHub)
  let res;
  for (let essai = 1; ; essai++) {
    try {
      res = await fetch('https://api.github.com' + url, {
        method,
        headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28',
          'User-Agent': 'suivi-mve-publier', ...(body ? { 'Content-Type': 'application/json' } : {}) },
        body: body ? JSON.stringify(body) : undefined,
        signal: AbortSignal.timeout(300000)
      });
      if (res.status < 500 || essai >= 5) break;
    } catch (e) {
      if (essai >= 5) {
        const err = new Error('Pas de connexion Internet (ou accès à api.github.com bloqué) après 5 essais' + (e.cause && e.cause.code ? ' (' + e.cause.code + ')' : e.name === 'TimeoutError' ? ' (délai dépassé)' : '') + '. Réessayez plus tard.');
        err.reseau = true; throw err;
      }
    }
    if (essai === 1) log('  (connexion à GitHub instable : nouvel essai…)');
    await new Promise((r) => setTimeout(r, 4000 * essai));
  }
  const text = await res.text();
  let json = null; try { json = JSON.parse(text); } catch (e) { /* réponse non JSON */ }
  if (!res.ok) {
    const e = new Error('GitHub a répondu ' + res.status + ' : ' + ((json && json.message) || text.slice(0, 200))); e.status = res.status; throw e;
  }
  return json;
}
function adresseGithub(depot) {
  const [compte, nom] = depot.split('/');
  return nom.toLowerCase() === compte.toLowerCase() + '.github.io'
    ? 'https://' + nom.toLowerCase() + '/'
    : 'https://' + compte.toLowerCase() + '.github.io/' + nom + '/';
}
// Remplace tout le contenu de la branche par les fichiers fournis, sans garder d'historique
// (commit « orphelin » poussé en force) : les anciennes données chiffrées ne restent pas dans le dépôt.
async function publierGithub(c, fichiers, publie, opts = {}) {
  const d = '/repos/' + c.github_depot, b = c.github_branche || 'main';
  const tous = opts.sansNojekyll ? fichiers : fichiers.concat([{ nom: '.nojekyll', data: Buffer.alloc(0) }]);
  const arbre = [];
  for (const f of tous) {
    const blob = await github(c.github_jeton, 'POST', d + '/git/blobs', { content: f.data.toString('base64'), encoding: 'base64' });
    arbre.push({ path: f.nom, mode: '100644', type: 'blob', sha: blob.sha });
  }
  const tree = await github(c.github_jeton, 'POST', d + '/git/trees', { tree: arbre });
  const commit = await github(c.github_jeton, 'POST', d + '/git/commits', { message: opts.message || ('Publication du ' + publie), tree: tree.sha, parents: [] });
  try {
    await github(c.github_jeton, 'PATCH', d + '/git/refs/heads/' + b, { sha: commit.sha, force: true });
  } catch (e) {
    if (!(e.status === 404 || (e.status === 422 && /not exist|not found/i.test(e.message)))) throw e;
    await github(c.github_jeton, 'POST', d + '/git/refs', { ref: 'refs/heads/' + b, sha: commit.sha }); // branche encore absente
  }
  if (c.sansAttente) return 'ok';
  // Attente de la mise en ligne (facultatif : nécessite l'autorisation « Pages » du jeton)
  for (let i = 0; i < 40; i++) {
    await new Promise((r) => setTimeout(r, 3000));
    let s;
    try { s = await github(c.github_jeton, 'GET', d + '/pages/builds/latest'); } catch (e) { return 'inconnu'; }
    if (s && s.commit === commit.sha && s.status === 'built') return 'ok';
    if (s && s.commit === commit.sha && s.status === 'errored') return 'erreur';
  }
  return 'en cours';
}
async function activerPages(c) {
  const d = '/repos/' + c.github_depot, b = c.github_branche || 'main';
  try {
    const p = await github(c.github_jeton, 'GET', d + '/pages');
    if (p.source && (p.source.branch !== b || p.source.path !== '/')) await github(c.github_jeton, 'PUT', d + '/pages', { source: { branch: b, path: '/' } });
    return true;
  } catch (e) {
    if (e.status !== 404) return false;
  }
  try { await github(c.github_jeton, 'POST', d + '/pages', { source: { branch: b, path: '/' } }); return true; } catch (e) { return false; }
}

/* ---------- Configuration guidée (première utilisation) ---------- */
async function configurer(ancien) {
  const c = Object.assign({}, ancien);
  log('\n=== Configuration de la publication Suivi MVE ===\n');
  log('D’où viennent les données ?');
  log('  1 = directement de DHIS2 (recommandé : pas besoin d’Excel)');
  log('  2 = du classeur Excel actualisé à la main');
  const choix = await ask('Votre choix' + (c.source ? ' [' + (c.source === 'excel' ? '2' : '1') + ']' : ' [1]') + ' : ');
  c.source = choix === '2' ? 'excel' : choix === '1' ? 'dhis2' : (c.source || 'dhis2');

  if (c.source === 'excel') {
    let f = await ask('\nChemin complet du classeur Excel actualisé' + (c.classeur ? ' [' + c.classeur + ']' : '') + ' :\n> ');
    f = f.replace(/^"|"$/g, '') || c.classeur;
    if (!f || !fs.existsSync(f)) stop('Fichier introuvable : ' + f);
    c.classeur = f;
    let fq = await ask('\nChemin du classeur « Qualité des données » (facultatif, Entrée pour ignorer)' + (c.classeur_qualite ? ' [' + c.classeur_qualite + ']' : '') + ' :\n> ');
    fq = fq.replace(/^"|"$/g, '') || c.classeur_qualite || '';
    if (fq && !fs.existsSync(fq)) stop('Fichier introuvable : ' + fq);
    c.classeur_qualite = fq || undefined;
  } else {
    c.dhis2_url = c.dhis2_url || 'https://ezd.snisrdc.com/dhis/';
    c.programme = c.programme || 'wDFB7FFlYQM';
    c.date_debut = c.date_debut || '2026-05-01';
    log('\nServeur DHIS2 : ' + c.dhis2_url + '  (programme ' + c.programme + ', alertes depuis le ' + c.date_debut + ')');
    log('Ce sont les réglages de la requête Power Query du classeur. Pour les changer, modifiez config.json.');
    const u = await ask('Nom d’utilisateur SNIS (le même que dans Excel)' + (c.dhis2_utilisateur ? ' [' + c.dhis2_utilisateur + ']' : '') + ' :\n> ');
    c.dhis2_utilisateur = u || c.dhis2_utilisateur;
    const p = await ask('Mot de passe SNIS' + (c.dhis2_mot_de_passe ? ' [laisser vide pour garder l’actuel]' : '') + ' :\n> ', { hidden: true });
    c.dhis2_mot_de_passe = p || c.dhis2_mot_de_passe;
    if (!c.dhis2_utilisateur || !c.dhis2_mot_de_passe) stop('L’identifiant et le mot de passe SNIS sont nécessaires.');
    log('Vérification de la connexion à DHIS2…');
    const moi = await dhis2(c, 'api/me.json', { fields: 'name,username' });
    log('✓ Connecté à DHIS2 en tant que ' + (moi.name || moi.username));
  }

  log('\nOù publier l’application ?');
  log('  1 = GitHub Pages (gratuit, recommandé : pas de limite de crédits par publication)');
  log('  2 = Netlify (le plan gratuit bloque les publications après environ 20 par mois)');
  const h = await ask('Votre choix' + (c.hebergeur ? ' [' + (c.hebergeur === 'netlify' ? '2' : '1') + ']' : ' [1]') + ' : ');
  c.hebergeur = h === '2' ? 'netlify' : h === '1' ? 'github' : (c.hebergeur || 'github');

  if (c.hebergeur === 'github') {
    log('\nAvant de continuer (voir LISEZ-MOI, section A bis) :');
    log('  - un dépôt PUBLIC sur github.com, par exemple « suivi-mve-hu », créé avec « Add a README file » ;');
    log('  - un jeton « fine-grained » limité à ce dépôt, avec les autorisations');
    log('    Contents : Read and write  et  Pages : Read and write.');
    let dep = await ask('Dépôt GitHub (compte/nom-du-depot)' + (c.github_depot ? ' [' + c.github_depot + ']' : '') + ' :\n> ');
    dep = (dep || c.github_depot || '').trim().replace(/^https?:\/\/github\.com\//i, '').replace(/\.git$|\/$/g, '');
    if (!/^[\w.-]+\/[\w.-]+$/.test(dep)) stop('Dépôt invalide : écrivez compte/nom-du-depot, par exemple adam-mve/suivi-mve-hu.');
    c.github_depot = dep;
    const gt = await ask('Jeton GitHub' + (c.github_jeton ? ' [laisser vide pour garder l’actuel]' : '') + ' :\n> ', { hidden: true });
    c.github_jeton = gt || c.github_jeton;
    if (!c.github_jeton) stop('Le jeton GitHub est nécessaire.');
    c.github_branche = c.github_branche || 'main';
    let r;
    try { r = await github(c.github_jeton, 'GET', '/repos/' + c.github_depot); } catch (e) {
      if (e.status === 401) stop('Jeton GitHub refusé. Vérifiez-le puis relancez « node publier.js --config ».');
      if (e.status === 404) stop('Dépôt « ' + c.github_depot + ' » introuvable, ou le jeton n’y a pas accès.');
      throw e;
    }
    if (r.private) stop('Le dépôt est privé : GitHub Pages gratuit exige un dépôt public (les données publiées restent chiffrées).');
    if (r.permissions && !r.permissions.push) stop('Le jeton ne peut pas écrire dans ce dépôt : donnez-lui « Contents : Read and write ».');
    c.adresse = adresseGithub(c.github_depot);
    log('✓ Dépôt GitHub accessible. Adresse de l’application : ' + c.adresse);
  } else {
  log('\nJeton Netlify : sur netlify.com, menu de votre compte › User settings › Applications ›');
  log('Personal access tokens › New access token. Copiez-le ici (il reste sur cet ordinateur).');
  const t = await ask('Jeton Netlify' + (c.jeton ? ' [laisser vide pour garder l’actuel]' : '') + ' :\n> ', { hidden: true });
  c.jeton = t || c.jeton;
  if (!c.jeton) stop('Le jeton Netlify est nécessaire.');
  }

  if (c.hebergeur === 'netlify' && !c.site_id) {
    for (;;) {
      const nom = (await ask('\nNom souhaité pour le site (lettres, chiffres, tirets), par ex. suivi-mve-hu :\n> ')).toLowerCase();
      try {
        const site = await netlify(c.jeton, 'POST', '/sites', JSON.stringify(nom ? { name: nom } : {}), 'application/json');
        c.site_id = site.id; c.adresse = site.ssl_url || site.url;
        log('✓ Site créé : ' + c.adresse);
        break;
      } catch (e) {
        if (e.status === 422) { log('Ce nom est déjà pris, essayez-en un autre.'); continue; }
        if (e.status === 401) stop('Jeton Netlify refusé. Vérifiez-le et relancez « node publier.js --config ».');
        throw e;
      }
    }
  }

  const changer = !c.mot_de_passe || (await ask('\nChanger le mot de passe d’équipe ? (o/N) ')).toLowerCase().startsWith('o');
  if (changer) {
    log('\nMot de passe d’équipe : au moins 12 caractères. Il sera demandé une fois sur chaque téléphone.');
    log('Conseil : une phrase de 4 ou 5 mots, par exemple « riposte-isiro-baobab-2026-ciel ».');
    for (;;) {
      const p1 = await ask('Mot de passe : ', { hidden: true });
      if (p1.length < 12) { log('Trop court (12 caractères minimum).'); continue; }
      const p2 = await ask('Confirmez : ', { hidden: true });
      if (p1 !== p2) { log('Les deux saisies sont différentes.'); continue; }
      c.mot_de_passe = p1;
      c.sel = crypto.randomBytes(16).toString('base64'); // nouveau sel : les téléphones redemanderont le mot de passe
      break;
    }
  }
  fs.writeFileSync(CONFIG, JSON.stringify(c, null, 2), { mode: 0o600 });
  log('\n✓ Configuration enregistrée dans ' + CONFIG + '\n  (gardez ce fichier sur cet ordinateur uniquement).\n');
  return c;
}

/* ---------- Lecture directe dans DHIS2 (même requête que le Power Query du classeur) ---------- */
async function dhis2(c, chemin, params, opts = {}) {
  const fin = opts.jeter ? (m) => { throw new Error(m); } : stop;
  const base = c.dhis2_url.replace(/\/?$/, '/');
  const url = new URL(chemin, base);
  for (const [k, v] of Object.entries(params || {})) {
    for (const x of [].concat(v)) url.searchParams.append(k, x);
  }
  const auth = /^d2p(at)?_/.test(c.dhis2_mot_de_passe) // jetons DHIS2 : « d2p_… » (versions récentes) ou « d2pat_… »
    ? 'ApiToken ' + c.dhis2_mot_de_passe
    : 'Basic ' + Buffer.from(c.dhis2_utilisateur + ':' + c.dhis2_mot_de_passe, 'utf8').toString('base64');
  // Jusqu'à 4 essais : une connexion coupée en cours de réponse (« terminated ») ou une erreur 5xx est réessayée
  let res, text, derniere;
  for (let essai = 1; essai <= 4; essai++) {
    try {
      res = await fetch(url, { headers: { Authorization: auth, Accept: 'application/json' }, redirect: 'manual', signal: AbortSignal.timeout(10 * 60000) });
      if (res.status >= 500 && essai < 4) throw new Error('erreur ' + res.status);
      text = await res.text();
      derniere = null; break;
    } catch (e) {
      derniere = e;
      if (essai < 4) await new Promise((r) => setTimeout(r, 5000 * essai));
    }
  }
  if (derniere) fin('Impossible de lire la réponse de DHIS2 (' + url.host + ') : ' + ((derniere.cause && derniere.cause.code) || derniere.message) + '. Vérifiez la connexion Internet.');
  if (res.status === 401 || res.status === 302) fin('DHIS2 refuse l’identifiant ou le mot de passe SNIS. Relancez « node publier.js --config ».');
  if (res.status === 403) fin('Votre compte SNIS n’a pas accès à cette information (erreur 403).');
  if (!res.ok) fin('DHIS2 a répondu ' + res.status + ' : ' + text.slice(0, 300));
  try { return JSON.parse(text); } catch (e) { fin('Réponse inattendue de DHIS2 (pas du JSON). L’adresse du serveur est-elle correcte ? ' + base); }
}

// Colonne de la réponse analytics → colonne du tableau de bord
const DIMENSIONS = {
  'Date début symptômes': 'uZZDLTjD15t', 'Sexe': 'a4OXn5u2eJy', 'Âge (ans)': 'ANjgCylgyBb', 'Âge (mois)': 'g8Rqc96GGGi',
  'Nature alerte': 'njs9IDEFVtC.kdOYmDgoyAA', 'Conclusion alerte': 'njs9IDEFVtC.KhsBtTYkFZd',
  'Classification finale': 'r7nrCHTBI5P.D6kduc7OZnS', 'Statut final': 'kOyiPgabAuY.Za0cx3pmcWW',
  'Date décès': 'kOyiPgabAuY.x1aazi4fgKO', 'Statut avant admission': 'rMvKKqab4bW.AawdHKqaXcj',
  'Contact avec malade Ebola': 'fqrtWpCo7za.Tzr3SapM9je', 'Date prélèvement': 'GO2aLxqhDIS.CxQAC5LkMtn',
  'Date analyse labo': 'r7nrCHTBI5P.BTMKxJvLTer', 'Date admission CTE': 'rMvKKqab4bW.KGsTJ4jV7Fb',
  'Date sortie CTE': 'rMvKKqab4bW.Xy5J5MGpaZ7', 'Vacciné MVE': 'fqrtWpCo7za.t4RcYSXmYgW',
  'Allaite': 'fqrtWpCo7za.dp2w0au7Qq7', 'Enceinte': 'fqrtWpCo7za.ICpmsUy8ros', 'Cas confirmés': 'rBeoKuRdbtw',
  // Indicateurs de programme du tableau de bord DHIS2 « MVE - Situation épidémiologique (17e EPIDEMIE) »
  'Décès confirmés (DHIS2)': 'Ox52yq3XZUX', 'Non cas (DHIS2)': 'mkyT4tONdVc'
};
const COL_DATES = ['Date notification', 'Date début symptômes', 'Date décès', 'Date prélèvement', 'Date analyse labo', 'Date admission CTE', 'Date sortie CTE'];
const COL_NOMBRES = ['Âge (ans)', 'Âge (mois)', 'Cas confirmés', 'Décès confirmés (DHIS2)', 'Non cas (DHIS2)'];

function niveau(h, i, suffixe) { // comme la fonction « Niveau » du Power Query
  const p = h ? String(h).split(' / ') : [];
  let v = p.length > i ? p[i] : null;
  if (v != null && v.length > 3 && v[2] === ' ' && v.slice(0, 2).toLowerCase() === v.slice(0, 2)) v = v.slice(3);
  return v == null ? null : v.replace(suffixe, '').trim();
}
function enSerie(t) { // « aaaa-mm-jj… » → numéro de série Excel (jour)
  const m = t && String(t).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? (Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 864e5 : null;
}
function maintenantSerie() { // heure locale, comme DateTime.LocalNow()
  const d = new Date();
  return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - Date.UTC(1899, 11, 30)) / 864e5;
}

async function lireDHIS2(c) {
  log('Interrogation de DHIS2 (' + new URL(c.dhis2_url).host + ')…');
  const meta = await dhis2(c, 'api/programs/' + c.programme + '.json', {
    fields: 'programStages[id,name,programStageDataElements[dataElement[id,name,shortName,code,formName,valueType]]],programTrackedEntityAttributes[trackedEntityAttribute[id,name,shortName,code,formName,valueType]]'
  });
  const elements = [];
  for (const s of meta.programStages || []) for (const e of s.programStageDataElements || []) elements.push(Object.assign({ col: s.id + '.' + e.dataElement.id }, e.dataElement));
  for (const a of meta.programTrackedEntityAttributes || []) elements.push(Object.assign({ col: a.trackedEntityAttribute.id }, a.trackedEntityAttribute));
  // Variable « MVE-N-Profession » (tableau de bord Suivi)
  const cible = 'MVE-N-PROFESSION';
  const p = elements.find((o) => [o.name, o.shortName, o.code, o.formName].some((x) => x && String(x).trim().toUpperCase() === cible));
  const dimProf = p ? p.col : null;
  if (!dimProf) log('  (variable « MVE-N-Profession » introuvable : la colonne Profession restera vide)');
  // Vraie date de notification (tableau de bord Qualité, comme la « Requête PQ corrigée »)
  const nI = (t) => String(t || '').toLowerCase().replace(/i/g, 'l');
  const parId = elements.filter((o) => nI(o.id) === nI(c.element_notif || 'PlM95bVTK26'));
  const parNom = elements.filter((o) => ['DATE', 'DATETIME'].includes(o.valueType) && /notification/i.test(o.name || ''));
  const colNotifQ = parId.length ? parId[0].col : parNom.length === 1 ? parNom[0].col : null;
  if (!colNotifQ) log('  (date de notification de l’étape introuvable : les contrôles de qualité ne seront pas publiés)');
  // Date de notification du tableau de bord Suivi MVE : variable « MPOX-N-Date et heure de notification de l'alerte »,
  // comme la colonne B de la « Requête corrigée » du classeur (et non plus la date d'inscription DHIS2).
  const nomNotif = String(c.var_notif || 'MPOX-N-Date et heure de notification de l\'alerte').trim().toUpperCase();
  const pNotif = elements.find((o) => [o.name, o.shortName, o.code, o.formName].some((x) => x && String(x).trim().toUpperCase() === nomNotif));
  const colNotifS = pNotif ? pNotif.col : null;
  if (!colNotifS) log('  (variable « ' + (c.var_notif || 'MPOX-N-Date et heure de notification de l\'alerte') + ' » introuvable : la date d’inscription sert de date de notification)');

  const demain = new Date(Date.now() + 864e5);
  const fin = demain.getFullYear() + '-' + String(demain.getMonth() + 1).padStart(2, '0') + '-' + String(demain.getDate()).padStart(2, '0');
  const dims = Array.from(new Set(['ou:USER_ORGUNIT'].concat(Object.values(DIMENSIONS), Object.values(DIM_QUALITE),
    dimProf ? [dimProf] : [], colNotifQ ? [colNotifQ] : [], colNotifS ? [colNotifS] : [])));
  const R = await dhis2(c, 'api/analytics/enrollments/query/' + c.programme + '.json', {
    dimension: dims, startDate: c.date_debut, endDate: fin, paging: 'false', dataIdScheme: 'NAME', displayProperty: 'NAME'
  });
  const noms = (R.headers || []).map((h) => h.name);
  const col = (n) => noms.indexOf(n);
  if (col('enrollmentdate') < 0) stop('Réponse DHIS2 inattendue : colonne « enrollmentdate » absente.');
  const extrait = maintenantSerie();
  const vide = (v) => (v === '' || v === undefined ? null : v);
  const brut = (r, id) => (id && col(id) >= 0 ? vide(r[col(id)]) : null);
  const typer = (nom, v) => {
    if (v == null) return null;
    if (COL_DATES.includes(nom) || ['Date confirmation', 'Date saisie'].includes(nom)) return enSerie(v);
    if (COL_NOMBRES.includes(nom)) { const x = Number(String(v).replace(',', '.')); return isNaN(x) ? null : x; }
    return String(v);
  };
  const zone = (r) => niveau(r[col('ounamehierarchy')], 2, ' Zone de Santé');
  const aire = (r) => niveau(r[col('ounamehierarchy')], 3, ' Aire de Santé');

  // 1. Tableau de bord Suivi MVE (date de notification de l'alerte, comme la « Requête corrigée » du classeur ;
  //    les alertes sans cette date sont écartées, comme dans le classeur)
  const lignes = (R.rows || []).map((r) => COLONNES.map((nom) => {
    if (nom === 'Date notification') return typer(nom, brut(r, colNotifS || 'enrollmentdate'));
    if (nom === 'Zone de santé') return zone(r);
    if (nom === 'Aire de santé') return aire(r);
    if (nom === 'Extrait le') return extrait;
    if (nom === 'Profession') return typer(nom, brut(r, dimProf));
    return typer(nom, brut(r, DIMENSIONS[nom]));
  })).filter((r) => r[0] != null).sort((a, b) => a[0] - b[0]);
  if (!lignes.length) stop('DHIS2 n’a renvoyé aucune alerte pour la période (depuis le ' + c.date_debut + ').');
  const ecartees = (R.rows || []).length - lignes.length;
  if (colNotifS && ecartees > 0) log('  (' + ecartees + ' alerte(s) sans date de notification de l’alerte, écartées du tableau Suivi comme dans le classeur)');

  // 2. Qualité des données (date de notification de l'étape, nom → clé anonyme)
  let qualite = null, alertesChaines = null;
  if (colNotifQ) {
    const source = Object.assign({}, DIMENSIONS, DIM_QUALITE, { 'Date notification': colNotifQ, 'Date saisie': 'enrollmentdate' });
    const lignesQ = (R.rows || []).map((r) => {
      const o = {};
      for (const nom of QCOLONNES) {
        if (nom === 'Zone de santé') o[nom] = zone(r);
        else if (nom === 'Aire de santé') o[nom] = aire(r);
        else if (nom === 'Extrait le') o[nom] = extrait;
        else if (nom === 'Clé doublon') o[nom] = null;
        else o[nom] = typer(nom, brut(r, source[nom]));
      }
      o.__nom = brut(r, DIM_QUALITE['Nom']);
      return o;
    });
    // Pour les chaînes de transmission : copie en mémoire avec les noms (jamais publiée telle quelle)
    const ACH = ['ID alerte', 'Date notification', 'Zone de santé', 'Aire de santé', 'Sexe', 'Âge (ans)', 'Âge (mois)', 'Classification finale', 'Statut final', 'Date décès', 'Cas confirmés', 'Nom'];
    alertesChaines = [ACH].concat(lignesQ.map((o) => ACH.map((k) => (k === 'Nom' ? o.__nom : (o[k] === undefined ? null : o[k])))));
    qualite = construireQualite(lignesQ, extrait, c.seuils_qualite);
  }
  return { table: [COLONNES].concat(lignes), n: lignes.length, extrait, qualite, alertesChaines };
}

/* ---------- Qualité : colonnes publiées et anonymisation des noms ---------- */
const DIM_QUALITE = { 'ID alerte': 'otfO0yxrBj5', 'Date confirmation': 'kOyiPgabAuY.L1jSKwWq5uI', 'Nom': 'GK0zoBqrpAx' };
const QCOLONNES = ['ID alerte', 'Date notification', 'Date début symptômes', 'Zone de santé', 'Aire de santé', 'Sexe', 'Âge (ans)',
  'Âge (mois)', 'Nature alerte', 'Conclusion alerte', 'Classification finale', 'Statut final', 'Date décès', 'Contact avec malade Ebola',
  'Date prélèvement', 'Date analyse labo', 'Date confirmation', 'Date admission CTE', 'Date sortie CTE', 'Vacciné MVE', 'Extrait le',
  'Cas confirmés', 'Clé doublon', 'Date saisie', 'Non cas (DHIS2)'];
const MVEQ = require(path.join(SITE, 'qualite', 'engine-qualite.js'));
// lignes : objets {colonne: valeur} + __nom (jamais publié)
function construireQualite(lignes, extrait, seuils) {
  const secret = crypto.randomBytes(32); // nouveau à chaque publication : les clés ne servent qu'à comparer
  for (const o of lignes) {
    const cle = MVEQ.cleNom(o.__nom, o['Sexe'], o['Âge (ans)'], o['Âge (mois)']);
    o['Clé doublon'] = cle ? crypto.createHmac('sha256', secret).update(cle, 'utf8').digest('hex').slice(0, 16) : null;
    delete o.__nom;
  }
  lignes.sort((a, b) => (a['Date notification'] == null ? Infinity : a['Date notification']) - (b['Date notification'] == null ? Infinity : b['Date notification'])
    || (a['Date saisie'] || 0) - (b['Date saisie'] || 0));
  const table = [QCOLONNES].concat(lignes.map((o) => QCOLONNES.map((k) => (o[k] === undefined ? null : o[k]))));
  const n = lignes.filter((o) => o['Date notification'] != null).length;
  return { table, n, sansDate: lignes.length - n, extrait, seuils: Object.assign({}, MVEQ.SEUILS_DEFAUT, seuils || {}) };
}
function lireClasseurQualite(fichier) {
  let buf;
  try { buf = fs.readFileSync(fichier); } catch (e) { stop('Impossible de lire le classeur Qualité (' + e.code + '). Fermez-le dans Excel puis réessayez.'); }
  const wb = XLSX.read(buf, { type: 'buffer', sheets: ['Données', 'Paramètres'], cellDates: false, cellFormula: false, cellHTML: false });
  if (!wb.Sheets['Données']) stop('La feuille « Données » est absente du classeur Qualité.');
  const t = XLSX.utils.sheet_to_json(wb.Sheets['Données'], { header: 1, raw: true, defval: null, blankrows: false });
  const h = (t[0] || []).map(norm);
  const idx = (n) => h.indexOf(norm(n));
  if (idx('Date notification') < 0 || idx('Nom') < 0) stop('Le classeur Qualité n’a pas les colonnes « Date notification » et « Nom ».');
  const lignes = t.slice(1).filter((r) => r && r.some((x) => x != null && x !== '')).map((r) => {
    const o = {};
    for (const k of QCOLONNES) o[k] = idx(k) >= 0 ? (r[idx(k)] === undefined ? null : r[idx(k)]) : null;
    o.__nom = r[idx('Nom')];
    return o;
  });
  const p = wb.Sheets['Paramètres'], seuils = {};
  if (p) ['DDSmax', 'AgeMax', 'CiblePrelev', 'RetInv', 'RetLabo', 'RetClassif', 'Vert', 'Orange', 'Fenetre']
    .forEach((k, i) => { const c = p['D' + (i + 2)]; if (c && typeof c.v === 'number') seuils[k] = c.v; });
  const ex = Math.max(0, ...lignes.map((o) => (typeof o['Extrait le'] === 'number' ? o['Extrait le'] : 0)));
  return construireQualite(lignes, ex || null, seuils);
}

/* ---------- Registre de vaccination (application « Vaccination MVE ») ----------
   Même lecture que la requête Power Query du classeur « Tableau de bord Vaccination MVE » :
   API tracker (données en temps réel), une ligne par événement de vaccination (ou par inscription
   sans vaccination). Les noms, téléphones et adresses ne quittent jamais l'ordinateur / GitHub :
   seuls des indicateurs de présence et une clé de doublon anonyme sont publiés. */
const VACC_DEFAUT = { programme: 'ebHbRt6ZxSq', etape_vacc: 'wf3RbrmLorl', etape_suivi: 'LJhBI8GIJE9', racine: 'wy1lwIP18SL' };
const VCOLONNES = ['ID vaccination', 'Date vaccination', 'Date événement', 'Date inscription', 'Date saisie', 'Zone de santé', 'Aire de santé',
  'Sexe', 'Âge (ans)', 'Contact d\'un cas', 'Consentement', 'Femme allaitante', 'Femme enceinte', 'Vaccin', 'Numéro de lot', 'Observation 15 min',
  'MAPI', 'MAPI grave', 'Statut événement', 'Nom renseigné', 'Téléphone', 'Avenue/Village renseigné', 'Nb suivis post-vaccinaux',
  'Nb fiches même nom/sexe/âge', 'Extrait le', 'Clé doublon'];
async function lireVaccination(c) {
  const V = Object.assign({}, VACC_DEFAUT, c.vaccination || {});
  if (V.desactive) return null;
  log('Lecture du registre de vaccination (programme ' + V.programme + ')…');
  const J = { jeter: true };
  const rep = await dhis2(c, 'api/tracker/trackedEntities', {
    program: V.programme, orgUnits: V.racine, orgUnitMode: 'DESCENDANTS', paging: 'false',
    fields: 'trackedEntity,attributes[attribute,value],enrollments[enrolledAt,orgUnit,events[programStage,occurredAt,createdAt,status,orgUnit,deleted,dataValues[dataElement,value]]]'
  }, J);
  const personnes = rep.trackedEntities || rep.instances || [];
  const ous = await dhis2(c, 'api/organisationUnits.json', { filter: 'path:like:' + V.racine, fields: 'id,name,level,ancestors[name,level]', paging: 'false' }, J);
  const nettoie = (t) => (t == null ? null : String(t).replace(/^hu /, '').replace(' Zone de Santé', '').replace(' Aire de Santé', '').trim());
  const niv = (o, n) => (o.level === n ? o.name : ((o.ancestors || []).find((a) => a.level === n) || {}).name);
  const OU = {};
  for (const o of ous.organisationUnits || []) OU[o.id] = { z: nettoie(niv(o, 3)), a: nettoie(niv(o, 4)) };
  const val = (l, k, id) => { const x = (l || []).find((y) => y[k] === id); return x && x.value != null && x.value !== '' ? x.value : null; };
  const oui = (v) => (v === 'true' ? 'Oui' : v === 'false' ? 'Non' : null);
  const date = (t) => enSerie(t);
  const tel = (p) => {
    const d = p ? String(p).replace(/\D/g, '') : '';
    if (!d) return 'Absent';
    return (d.length === 10 && /^0[89]/.test(d)) || (d.length === 12 && /^243[89]/.test(d)) || (d.length === 9 && /^[89]/.test(d)) ? 'Valide' : 'Invalide';
  };
  const cleNom = (n, s, a) => {
    if (!n) return null;
    const mots = String(n).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z ]/g, '').split(' ').filter(Boolean).sort();
    return mots.length ? mots.join('') + '|' + (s || '') + '|' + (a || '') : null;
  };
  const extrait = maintenantSerie(), lignes = [];
  for (const p of personnes) {
    const A = (id) => val(p.attributes, 'attribute', id);
    for (const e of p.enrollments || []) {
      const evs = (e.events || []).filter((v) => v.deleted !== true);
      const vacs = evs.filter((v) => v.programStage === V.etape_vacc);
      const nbS = evs.filter((v) => v.programStage === V.etape_suivi).length;
      for (const v of vacs.length ? vacs : [null]) {
        const DV = (id) => (v ? val(v.dataValues, 'dataElement', id) : null);
        const ou = OU[v ? v.orgUnit : e.orgUnit] || { z: null, a: null };
        const sx = A('CklPZdOd6H1'), age = A('WDvjwflOAJg'), st = v ? v.status : null;
        lignes.push({
          'ID vaccination': A('O3dgLbMEHkr'), 'Date vaccination': date(DV('t5EZnDkNcTI')), 'Date événement': v ? date(v.occurredAt) : null,
          'Date inscription': date(e.enrolledAt), 'Date saisie': v ? date(v.createdAt) : null, 'Zone de santé': ou.z, 'Aire de santé': ou.a,
          'Sexe': sx === 'MALE' ? 'Masculin' : sx === 'FEMALE' ? 'Féminin' : sx, 'Âge (ans)': age == null || isNaN(+age) ? null : +age,
          'Contact d\'un cas': oui(A('L6pXunhryH7')), 'Consentement': oui(A('DaJt8Pg6ZuA')), 'Femme allaitante': oui(DV('ZWOSpZsZvBm')),
          'Femme enceinte': oui(DV('jot4RqggkhF')), 'Vaccin': DV('RQqJPmtvX7T'), 'Numéro de lot': DV('Zz30BstEKsk'), 'Observation 15 min': oui(DV('L296IDDLaq1')),
          'MAPI': oui(DV('uNpMVWZMLab')), 'MAPI grave': oui(DV('w2ILcg807sJ')), 'Statut événement': st === 'COMPLETED' ? 'Complété' : st === 'ACTIVE' ? 'Actif' : st,
          'Nom renseigné': A('MfDgmmCkBis') ? 'Oui' : 'Non', 'Téléphone': tel(A('fctSQp5nAYl')), 'Avenue/Village renseigné': A('KlltlZWwIe9') ? 'Oui' : 'Non',
          'Nb suivis post-vaccinaux': nbS, 'Extrait le': extrait, __cle: cleNom(A('MfDgmmCkBis'), sx, age)
        });
      }
    }
  }
  if (!lignes.length) { log('  (registre de vaccination vide ou inaccessible avec ce compte : application Vaccination non mise à jour)'); return null; }
  const nb = {}, secret = crypto.randomBytes(32);
  for (const o of lignes) if (o.__cle) nb[o.__cle] = (nb[o.__cle] || 0) + 1;
  for (const o of lignes) {
    o['Nb fiches même nom/sexe/âge'] = o.__cle ? nb[o.__cle] : 1;
    o['Clé doublon'] = o.__cle ? crypto.createHmac('sha256', secret).update(o.__cle, 'utf8').digest('hex').slice(0, 16) : null;
    delete o.__cle;
  }
  lignes.sort((a, b) => ((a['Date vaccination'] ?? a['Date événement'] ?? Infinity) - (b['Date vaccination'] ?? b['Date événement'] ?? Infinity)));
  return { table: [VCOLONNES].concat(lignes.map((o) => VCOLONNES.map((k) => (o[k] === undefined ? null : o[k])))), n: lignes.length,
    nVac: lignes.filter((o) => o['Date événement'] != null).length, parametres: c.parametres_vaccination || {} };
}

/* Lecture page par page (les grosses réponses d'un seul bloc sont coupées par le serveur ou le réseau) */
async function dhis2Pages(c, chemin, params, cles, opts = {}) {
  const taille = 1000, tout = [];
  for (let page = 1; page <= 500; page++) {
    const rep = await dhis2(c, chemin, Object.assign({}, params, { paging: 'true', page, pageSize: taille }), opts);
    const l = cles.map((k) => rep[k]).find(Array.isArray) || [];
    tout.push(...l);
    if (l.length < taille) break;
  }
  return tout;
}

/* ---------- Fond de carte de « Suivi MVE » (contours des zones et aires de santé + population) ----------
   Données administratives publiques de DHIS2 : aucune donnée de patient. Publié en clair (carte.json),
   le téléphone le garde pour fonctionner sans Internet. */
/* Construction du fond de carte (contours des zones et des aires de santé + population).
   Données administratives publiques : aucune donnée de patient. */
function construireCarte(ous, pops, tol) {
  tol = tol || 0.002; // ≈ 200 m : invisible à l'échelle d'un téléphone
  var court = function (n, suffixe) {
    var v = String(n || '');
    if (v.length > 3 && v[2] === ' ' && v.slice(0, 2).toLowerCase() === v.slice(0, 2)) v = v.slice(3);
    return v.replace(suffixe, '').trim();
  };
  function dp(pts, eps) { // Douglas-Peucker
    if (pts.length <= 4) return pts;
    var keep = new Uint8Array(pts.length); keep[0] = keep[pts.length - 1] = 1;
    var pile = [[0, pts.length - 1]];
    while (pile.length) {
      var s = pile.pop(), a = s[0], b = s[1], ax = pts[a][0], ay = pts[a][1], bx = pts[b][0], by = pts[b][1];
      var dx = bx - ax, dy = by - ay, L = dx * dx + dy * dy, dmax = -1, imax = -1;
      for (var i = a + 1; i < b; i++) {
        var px = pts[i][0] - ax, py = pts[i][1] - ay, d;
        if (L === 0) d = px * px + py * py;
        else { var t = Math.max(0, Math.min(1, (px * dx + py * dy) / L)); var qx = px - t * dx, qy = py - t * dy; d = qx * qx + qy * qy; }
        if (d > dmax) { dmax = d; imax = i; }
      }
      if (dmax > eps * eps) { keep[imax] = 1; pile.push([a, imax], [imax, b]); }
    }
    return pts.filter(function (p, i) { return keep[i]; });
  }
  function anneaux(g) {
    if (!g || !g.coordinates) return null;
    var polys = g.type === 'Polygon' ? [g.coordinates] : g.type === 'MultiPolygon' ? g.coordinates : null;
    if (!polys) return null;
    var out = [];
    polys.forEach(function (poly) {
      poly.forEach(function (ring, k) {
        var r = dp(ring, tol);
        if (r.length < 4) { if (k > 0) return; r = ring.length > 40 ? dp(ring, tol / 10) : ring; } // ne jamais perdre un contour extérieur
        if (r.length < 4) return;
        var flat = [];
        r.forEach(function (p) { flat.push(Math.round(p[0] * 1e4) / 1e4, Math.round(p[1] * 1e4) / 1e4); });
        out.push(flat);
      });
    });
    return out.length ? out : null;
  }
  var zones = [], aires = [], sansContour = [];
  ous.forEach(function (o) {
    var g = anneaux(o.geometry);
    if (o.level === 3) { if (g) zones.push({ n: court(o.name, ' Zone de Santé'), g: g }); }
    else if (o.level === 4) {
      var a = { n: court(o.name, ' Aire de Santé'), z: court(o.parent && o.parent.name, ' Zone de Santé') };
      if (pops && pops[o.id] != null) a.pop = Math.round(pops[o.id]);
      if (g) { a.g = g; aires.push(a); } else sansContour.push(a);
    }
  });
  var tri = function (a, b) { return (a.z || '').localeCompare(b.z || '', 'fr') || a.n.localeCompare(b.n, 'fr'); };
  zones.sort(tri); aires.sort(tri);
  return { v: 1, zones: zones, aires: aires, sansContour: sansContour };
}

const CARTE_DEFAUT = { racine: 'wy1lwIP18SL', population: 'Z3OgyjnbyhL', tolerance: 0.002, population_fichier: 'population-aires.xlsx' };
/* Population par aire depuis un classeur (ordinateur\population-aires.xlsx) : colonnes « Zone de santé »,
   « Aire de santé » et « Population AAAA ». Renvoie { annee, valeurs: { 'zone|aire' normalisé : population } }. */
function lirePopulationFichier(nom) {
  const f = path.isAbsolute(nom) ? nom : path.join(ICI, nom);
  if (!fs.existsSync(f)) return null;
  const wb = XLSX.read(fs.readFileSync(f), { type: 'buffer' });
  for (const sn of wb.SheetNames) {
    const t = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null, blankrows: false });
    for (let h = 0; h < Math.min(t.length, 10); h++) {
      const en = (t[h] || []).map((x) => norm(x));
      const iz = en.indexOf('zone de sante'), ia = en.indexOf('aire de sante');
      const ip = en.findIndex((x) => /^population( \d{4})?$/.test(x));
      if (iz < 0 || ia < 0 || ip < 0) continue;
      const m = String(t[h][ip]).match(/(\d{4})/), valeurs = {};
      let n = 0;
      for (const r of t.slice(h + 1)) {
        const v = r && Number(r[ip]);
        if (!r || !r[iz] || !r[ia] || !(v > 0)) continue;
        valeurs[norm(r[iz]) + '|' + norm(r[ia])] = Math.round(v); n++;
      }
      if (n) return { annee: m ? m[1] : null, valeurs, n, fichier: path.basename(f) };
    }
  }
  throw new Error('colonnes « Zone de santé », « Aire de santé » et « Population » introuvables dans ' + path.basename(f));
}
async function lireCarte(c) {
  const K = Object.assign({}, CARTE_DEFAUT, c.carte || {});
  if (K.desactive) return null;
  const J = { jeter: true };
  log('Lecture des contours des zones et aires de santé…');
  const r = await dhis2(c, 'api/organisationUnits.json', {
    filter: ['path:like:' + K.racine, 'level:in:[3,4]'], fields: 'id,name,level,parent[name],geometry', paging: 'false'
  }, J);
  const ous = r.organisationUnits || [];
  // Population de l'aire : année en cours, sinon les 3 années précédentes
  let pops = null, annee = null;
  for (let a = new Date().getFullYear(); a >= new Date().getFullYear() - 3 && !pops; a--) {
    try {
      const p = await dhis2(c, 'api/analytics.json', { dimension: ['dx:' + K.population, 'ou:' + K.racine + ';LEVEL-4'], filter: 'pe:' + a, skipMeta: 'true' }, J);
      if (p.rows && p.rows.length) { pops = {}; annee = String(a); for (const x of p.rows) pops[x[1]] = Number(x[2]); }
    } catch (e) { /* population non disponible pour cette année */ }
  }
  const carte = construireCarte(ous, pops, K.tolerance);
  if (!carte.aires.length) throw new Error('aucun contour d’aire de santé dans DHIS2');
  carte.genere = new Date().toISOString();
  carte.popAnnee = annee;
  carte.popSource = annee ? 'DHIS2 « Population de l’Aire de Santé »' : null;
  // Classeur de population : utilisé s'il est plus récent que DHIS2 (ou à année égale)
  let P = null;
  try { P = K.population_fichier ? lirePopulationFichier(K.population_fichier) : null; } catch (e) { log('  ⚠ Population du classeur ignorée (' + e.message + ')'); }
  if (P && (!annee || !P.annee || Number(P.annee) >= Number(annee))) {
    const cleA = (a) => norm(a.z) + '|' + norm(a.n);
    let trouvees = 0;
    for (const a of carte.aires.concat(carte.sansContour)) {
      const v = P.valeurs[cleA(a)];
      if (v != null) { a.pop = v; trouvees++; } else delete a.pop;
    }
    carte.popAnnee = P.annee; carte.popSource = 'classeur « ' + P.fichier + ' »';
    log('  Population ' + (P.annee || '') + ' du classeur ' + P.fichier + ' : ' + trouvees + ' aires sur ' + carte.aires.length +
      (trouvees < carte.aires.length ? ' (les autres : population inconnue)' : ''));
  }
  return carte;
}

/* ---------- Suivi des contacts (application « Contacts MVE ») ----------
   Même traitement que les requêtes Power Query « DHIS2_Contacts » et « DHIS2_Suivis » du classeur
   « Tableau de bord Contacts MVE », mais lu dans l'API tracker (temps réel). Aucun nom, téléphone ni
   village n'est lu : seulement l'identifiant interne, l'ID contact, le N° Epid du cas source, les dates,
   la zone/aire, le sexe, l'âge et les variables épidémiologiques. */
const CONT_DEFAUT = { programme: 'wyLS271t8Id', etape: 'YIebrtM7xO2', racine: 'wy1lwIP18SL', date_debut: '2026-05-01', duree: 21 };
const CCOLONNES = ['Enrollment', 'ID contact', 'N° Epid cas source', 'Date enregistrement', 'Date dernier contact', 'Zone de santé', 'Aire de santé',
  'Sexe', 'Âge (ans)', 'Âge (mois)', 'Relation avec le cas', 'Circonstance du contact', 'Type de contact', 'Notifié par', 'Début suivi',
  'Fin suivi (J21)', 'Issue', 'Date issue', 'Jours de suivi saisis', 'Symptômes signalés', '1er contact du cas source', 'Extrait le'];
const SCOLONNES = ['Enrollment', 'Date suivi', 'Vu', 'Raison non vu', 'Symptômes', 'Résultat du suivi', 'Issue', 'Dernier jour de suivi', 'Nb fiches', 'Extrait le'];
async function lireContacts(c) {
  const K = Object.assign({}, CONT_DEFAUT, c.contacts || {});
  if (K.desactive) return null;
  log('Lecture du suivi des contacts (programme ' + K.programme + ')…');
  const J = { jeter: true };
  // Libellés des options (l'API analytics du classeur renvoie les noms ; l'API tracker renvoie les codes)
  const meta = await dhis2(c, 'api/programs/' + K.programme + '.json', {
    fields: 'programTrackedEntityAttributes[trackedEntityAttribute[id,optionSet[options[code,name]]]]'
  }, J);
  const OPT = {};
  for (const a of meta.programTrackedEntityAttributes || []) {
    const t = a.trackedEntityAttribute || {}, m = {};
    for (const o of (t.optionSet && t.optionSet.options) || []) m[o.code] = o.name;
    OPT[t.id] = m;
  }
  const personnes = await dhis2Pages(c, 'api/tracker/trackedEntities', {
    program: K.programme, orgUnits: K.racine, orgUnitMode: 'DESCENDANTS',
    fields: 'attributes[attribute,value],enrollments[enrollment,program,enrolledAt,orgUnit,deleted]'
  }, ['trackedEntities', 'instances'], J);
  log('  ' + personnes.length + ' personnes lues ; lecture des fiches de suivi…');
  const fiches = await dhis2Pages(c, 'api/tracker/events', {
    program: K.programme, programStage: K.etape, ouMode: 'ACCESSIBLE', occurredAfter: K.date_debut,
    fields: 'event,enrollment,occurredAt,status,dataValues[dataElement,value]'
  }, ['events', 'instances'], J);
  const ous = await dhis2(c, 'api/organisationUnits.json', { filter: 'path:like:' + K.racine, fields: 'id,name,level,ancestors[name,level]', paging: 'false' }, J);
  const nettoie = (t) => (t == null ? null : String(t).replace(/^[a-z]{2} /, '').replace(' Zone de Santé', '').replace(' Aire de Santé', '').trim());
  const niv = (o, n) => (o.level === n ? o.name : ((o.ancestors || []).find((a) => a.level === n) || {}).name);
  const OU = {};
  for (const o of ous.organisationUnits || []) OU[o.id] = { z: nettoie(niv(o, 3)), a: nettoie(niv(o, 4)) };
  const val = (l, k, id) => { const x = (l || []).find((y) => y[k] === id); return x && x.value != null && x.value !== '' ? x.value : null; };
  const lib = (codes, v) => (v == null || v === '' ? null : (Object.prototype.hasOwnProperty.call(codes, String(v)) ? codes[String(v)] : String(v)));
  const prio = (vals, ordre) => ordre.find((o) => vals.includes(o)) ?? null;
  const OUINON = { true: 'Oui', false: 'Non' }, SYMPT = { O: 'Oui', N: 'Non', JNCP: 'Je ne sais pas' };
  const RAISON = { Ab: 'Absence', PDV: 'Perdu de vue (non vu 48 h)', D: 'Déplacé', JV: 'Jamais vu', RR: 'Réticence/Refus', DCD: 'Décédé(e)', TZS: 'Transféré vers une autre ZS' };
  const RESULTAT = { NC: 'Non cas', CS: 'Contact devenu suspect', CM: 'Confirmé', RC: 'Recyclé', DB: 'Doublon', Autre: 'Autre' };
  const extrait = maintenantSerie();

  // --- Requête DHIS2_Suivis : une ligne par contact et par jour ---
  const groupes = new Map();
  for (const e of fiches) {
    if (e.deleted === true || !e.enrollment || e.status === 'SCHEDULE' || e.status === 'SKIPPED') continue;
    const dv = e.dataValues || [], V = (id) => val(dv, 'dataElement', id);
    const date = enSerie(V('YErKiJ9fvXj')) ?? enSerie(e.occurredAt);
    if (date == null) continue;
    const res = lib(RESULTAT, V('xJ4Lfaz8i6R')), raison = lib(RAISON, V('HLR2jTQTUJN'));
    const f = {
      vu: lib(OUINON, V('a9PXWLKqQQG')), raison, sympt: lib(SYMPT, V('L0lvdD4XWzk')), res,
      issue: res === 'Confirmé' ? 'Confirmé' : res === 'Contact devenu suspect' ? 'Devenu suspect' : raison === 'Décédé(e)' ? 'Décédé'
        : raison === 'Transféré vers une autre ZS' ? 'Transféré autre ZS' : res === 'Doublon' ? 'Doublon' : null,
      dernier: lib(OUINON, V('RLo66HizupL'))
    };
    const cle = e.enrollment + '|' + date;
    if (!groupes.has(cle)) groupes.set(cle, { enr: e.enrollment, date, l: [] });
    groupes.get(cle).l.push(f);
  }
  const suivis = [...groupes.values()].map((g) => {
    const L = (k) => g.l.map((x) => x[k]);
    return {
      'Enrollment': g.enr, 'Date suivi': g.date, 'Vu': prio(L('vu'), ['Oui', 'Non']), 'Raison non vu': L('raison').find((x) => x != null) ?? null,
      'Symptômes': prio(L('sympt'), ['Oui', 'Je ne sais pas', 'Non']),
      'Résultat du suivi': prio(L('res'), ['Confirmé', 'Contact devenu suspect', 'Doublon', 'Recyclé', 'Autre', 'Non cas']),
      'Issue': prio(L('issue'), ['Confirmé', 'Devenu suspect', 'Décédé', 'Transféré autre ZS', 'Doublon']),
      'Dernier jour de suivi': prio(L('dernier'), ['Oui', 'Non']), 'Nb fiches': g.l.length, 'Extrait le': extrait
    };
  }).sort((a, b) => (a.Enrollment < b.Enrollment ? -1 : a.Enrollment > b.Enrollment ? 1 : a['Date suivi'] - b['Date suivi']));
  const parEnr = new Map();
  for (const s of suivis) { if (!parEnr.has(s.Enrollment)) parEnr.set(s.Enrollment, []); parEnr.get(s.Enrollment).push(s); }

  // --- Requête DHIS2_Contacts : une ligne par contact ---
  const debut = enSerie(K.date_debut), fin = Math.floor(extrait) + 1;
  const nombre = (t) => (t == null || t === '' || isNaN(+t) ? null : +t);
  const sansNumero = (t) => { if (t == null || t === '') return null; const s = String(t).trim(), p = s.indexOf('.'); return p > 0 && p <= 2 && !isNaN(+s.slice(0, p)) ? s.slice(p + 1).trim() : s; };
  const lignes = [];
  for (const p of personnes) {
    const A = (id) => { const v = val(p.attributes, 'attribute', id); return v == null ? null : lib(OPT[id] || {}, v); };
    for (const e of p.enrollments || []) {
      if (e.deleted === true || (e.program && e.program !== K.programme)) continue;
      const de = enSerie(e.enrolledAt);
      if (de == null || de < debut || de > fin) continue;
      const dc = enSerie(A('rFE7EHPn5S8')), ou = OU[e.orgUnit] || { z: null, a: null };
      const deb = dc == null || (de != null && dc > de) ? de : dc;
      const S = parEnr.get(e.enrollment) || [];
      const prem = S.filter((s) => s.Issue != null).sort((a, b) => a['Date suivi'] - b['Date suivi'])[0];
      const rel = A('doL36yFlRUG'), epid = A('tBvV2SfZxOf');
      lignes.push({
        'Enrollment': e.enrollment, 'ID contact': A('nRK12xQqtxf'), 'N° Epid cas source': epid == null ? null : String(epid).trim().toUpperCase() || null,
        'Date enregistrement': de, 'Date dernier contact': dc, 'Zone de santé': ou.z, 'Aire de santé': ou.a, 'Sexe': A('a4OXn5u2eJy'),
        'Âge (ans)': nombre(A('ANjgCylgyBb')), 'Âge (mois)': nombre(A('g8Rqc96GGGi')), 'Relation avec le cas': rel == null ? null : String(rel).trim(),
        'Circonstance du contact': sansNumero(A('HVRjyIknBzc')), 'Type de contact': sansNumero(A('wmhOOHWkk0l')), 'Notifié par': A('h0U3OJRhR3Z'),
        'Début suivi': deb, 'Fin suivi (J21)': deb == null ? null : deb + K.duree, 'Issue': prem ? prem.Issue : null, 'Date issue': prem ? prem['Date suivi'] : null,
        'Jours de suivi saisis': S.length, 'Symptômes signalés': S.some((s) => s['Symptômes'] === 'Oui') ? 'Oui' : 'Non', '1er contact du cas source': 'Non', 'Extrait le': extrait,
        'Nom et prénom du contact': A('luaR4AtJPBC') // en mémoire seulement (chaînes de transmission) : jamais publié
      });
    }
  }
  if (!lignes.length) { log('  (aucun contact lisible avec ce compte : application Contacts non mise à jour)'); return null; }
  lignes.sort((a, b) => a['Date enregistrement'] - b['Date enregistrement'] || (a.Enrollment < b.Enrollment ? -1 : a.Enrollment > b.Enrollment ? 1 : 0));
  const vus = new Set();
  for (const o of lignes) { const k = o['N° Epid cas source']; if (k && !vus.has(k)) { vus.add(k); o['1er contact du cas source'] = 'Oui'; } }
  const T = (cols, l) => [cols].concat(l.map((o) => cols.map((k) => (o[k] === undefined ? null : o[k]))));
  return { contacts: T(CCOLONNES, lignes), suivis: T(SCOLONNES, suivis), n: lignes.length, nS: suivis.length,
    contactsNoms: T(CCOLONNES.concat(['Nom et prénom du contact']), lignes) };
}

/* ---------- Lecture du classeur ---------- */
function lireClasseur(fichier) {
  let buf;
  try { buf = fs.readFileSync(fichier); } catch (e) { stop('Impossible de lire le classeur (' + e.code + '). Fermez-le dans Excel puis réessayez.'); }
  let wb = XLSX.read(buf, { type: 'buffer', sheets: ['Données'], cellDates: false, cellFormula: false, cellHTML: false, cellText: false, dense: true });
  const feuille = wb.Sheets['Données'];
  if (!feuille) stop('La feuille « Données » est absente du classeur.');
  const table = XLSX.utils.sheet_to_json(feuille, { header: 1, raw: true, defval: null, blankrows: false });
  const entete = (table[0] || []).map(norm);
  const idx = COLONNES.map((c) => entete.indexOf(norm(c)));
  const manquantes = COLONNES.filter((c, i) => idx[i] < 0 && ['Date notification', 'Zone de santé', 'Aire de santé', 'Classification finale', 'Cas confirmés'].includes(c));
  if (manquantes.length) stop('Colonnes introuvables dans « Données » : ' + manquantes.join(', '));
  const iNotif = idx[0];
  const lignes = table.slice(1).filter((r) => r && r[iNotif] != null && r[iNotif] !== '')
    .map((r) => idx.map((i) => (i < 0 ? null : r[i] === undefined ? null : r[i])));
  if (!lignes.length) stop('Aucune alerte dans la feuille « Données ». Le classeur a-t-il été actualisé et enregistré ?');
  const iEx = COLONNES.indexOf('Extrait le');
  const extrait = Math.max(...lignes.map((r) => (typeof r[iEx] === 'number' ? r[iEx] : 0)));
  return { table: [COLONNES].concat(lignes), n: lignes.length, extrait };
}
function serieEnTexte(s) {
  if (!s) return 'inconnue';
  const d = new Date(Date.UTC(1899, 11, 30) + s * 864e5), p = (n) => String(n).padStart(2, '0');
  return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear() + ' à ' + p(d.getUTCHours()) + 'h' + p(d.getUTCMinutes());
}

/* ---------- Chiffrement (format lu par l'application : MVE1 | itérations | sel | iv | données+tag) ---------- */
function chiffrer(objet, motDePasse, selB64) {
  const sel = Buffer.from(selB64, 'base64');
  const cle = crypto.pbkdf2Sync(Buffer.from(motDePasse, 'utf8'), sel, ITERATIONS, 32, 'sha256');
  const iv = crypto.randomBytes(12);
  const clair = zlib.gzipSync(Buffer.from(JSON.stringify(objet), 'utf8'), { level: 9 });
  const c = crypto.createCipheriv('aes-256-gcm', cle, iv);
  const corps = Buffer.concat([c.update(clair), c.final(), c.getAuthTag()]);
  const tete = Buffer.alloc(8); tete.write('MVE1', 0, 'ascii'); tete.writeUInt32BE(ITERATIONS, 4);
  return Buffer.concat([tete, sel, iv, corps]);
}

/* ---------- Archive zip (pour le déploiement Netlify) ---------- */
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(b) { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function zip(fichiers) {
  const locaux = [], centraux = []; let offset = 0;
  const now = new Date();
  const dosT = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const dosD = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  for (const { nom, data } of fichiers) {
    const n = Buffer.from(nom, 'utf8'), comp = zlib.deflateRawSync(data), crc = crc32(data);
    const loc = Buffer.alloc(30);
    loc.writeUInt32LE(0x04034b50, 0); loc.writeUInt16LE(20, 4); loc.writeUInt16LE(0x0800, 6); loc.writeUInt16LE(8, 8);
    loc.writeUInt16LE(dosT, 10); loc.writeUInt16LE(dosD, 12); loc.writeUInt32LE(crc, 14); loc.writeUInt32LE(comp.length, 18); loc.writeUInt32LE(data.length, 22);
    loc.writeUInt16LE(n.length, 26); loc.writeUInt16LE(0, 28);
    const cen = Buffer.alloc(46);
    cen.writeUInt32LE(0x02014b50, 0); cen.writeUInt16LE(20, 4); cen.writeUInt16LE(20, 6); cen.writeUInt16LE(0x0800, 8); cen.writeUInt16LE(8, 10);
    cen.writeUInt16LE(dosT, 12); cen.writeUInt16LE(dosD, 14); cen.writeUInt32LE(crc, 16); cen.writeUInt32LE(comp.length, 20); cen.writeUInt32LE(data.length, 24);
    cen.writeUInt16LE(n.length, 28); cen.writeUInt32LE(offset, 42);
    locaux.push(loc, n, comp); centraux.push(cen, n);
    offset += 30 + n.length + comp.length;
  }
  const dir = Buffer.concat(centraux), fin = Buffer.alloc(22);
  fin.writeUInt32LE(0x06054b50, 0); fin.writeUInt16LE(fichiers.length, 8); fin.writeUInt16LE(fichiers.length, 10);
  fin.writeUInt32LE(dir.length, 12); fin.writeUInt32LE(offset, 16);
  return Buffer.concat(locaux.concat([dir, fin]));
}
function fichiersDuSite(dossier, base = '') {
  let out = [];
  for (const e of fs.readdirSync(dossier, { withFileTypes: true })) {
    if (e.name.startsWith('.') || e.name.endsWith('.enc')) continue;
    const p = path.join(dossier, e.name), rel = base ? base + '/' + e.name : e.name;
    if (e.isDirectory()) out = out.concat(fichiersDuSite(p, rel));
    else out.push({ nom: rel, data: fs.readFileSync(p) });
  }
  return out;
}

/* ---------- Publication automatique dans le nuage (GitHub Actions) ----------
   Sur GitHub, le programme tourne sans config.json : les réglages non secrets sont lus dans
   config-nuage.json (déposé dans le dépôt) et les secrets dans les variables d'environnement
   fournies par les « secrets » du dépôt. */
const EN_NUAGE = process.env.GITHUB_ACTIONS === 'true' && !args.includes('--local');
const INSTALLER_NUAGE = args.includes('--installer-nuage');
const BRANCHE_SITE = 'gh-pages';
const CONFIG_NUAGE = path.join(ICI, 'config-nuage.json');
const CLES_NUAGE = ['vaccination', 'parametres_vaccination', 'contacts', 'carte', 'source', 'hebergeur', 'dhis2_url', 'programme', 'date_debut', 'element_notif', 'var_notif',
  'seuils_qualite', 'github_depot', 'github_branche', 'sel'];
function configNuage() {
  if (!fs.existsSync(CONFIG_NUAGE)) stop('config-nuage.json absent du dépôt : relancez « node publier.js --installer-nuage » sur l’ordinateur.');
  const c = JSON.parse(fs.readFileSync(CONFIG_NUAGE, 'utf8'));
  const env = process.env, manque = [];
  c.mot_de_passe = env.MVE_MOT_DE_PASSE || '';
  if (!c.mot_de_passe) manque.push('MVE_MOT_DE_PASSE');
  if (env.MVE_DHIS2_JETON) { c.dhis2_utilisateur = 'jeton'; c.dhis2_mot_de_passe = env.MVE_DHIS2_JETON.trim(); }
  else if (env.MVE_DHIS2_UTILISATEUR && env.MVE_DHIS2_MOT_DE_PASSE) { c.dhis2_utilisateur = env.MVE_DHIS2_UTILISATEUR; c.dhis2_mot_de_passe = env.MVE_DHIS2_MOT_DE_PASSE; }
  else manque.push('MVE_DHIS2_JETON');
  c.github_jeton = (env.MVE_GITHUB_JETON || env.GITHUB_TOKEN || '').trim();
  if (!c.github_jeton) manque.push('MVE_GITHUB_JETON');
  if (c.dhis2_mot_de_passe && !/^d2p(at)?_/.test(c.dhis2_mot_de_passe) && env.MVE_DHIS2_JETON) stop('Le secret MVE_DHIS2_JETON doit être un jeton d’accès personnel DHIS2 (il commence par d2p_ ou d2pat_).');
  if (manque.length) stop('Secrets manquants dans le dépôt (Settings › Secrets and variables › Actions) : ' + manque.join(', '));
  return c;
}
const WORKFLOW = `# Publication automatique Suivi MVE / Qualité MVE (générée par « node publier.js --installer-nuage »)
name: Publication Suivi MVE
on:
  schedule:
    - cron: '7,37 * * * *'   # toutes les 30 minutes (à h07 et h37, pour éviter l'encombrement de l'heure pile)
  workflow_dispatch: {}       # bouton « Run workflow » pour publier à la demande
permissions:
  contents: write
  pages: read
concurrency:
  group: publication-mve
  cancel-in-progress: false
jobs:
  publier:
    runs-on: ubuntu-latest
    timeout-minutes: 30
    env:
      TZ: Africa/Lubumbashi
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: '20'
      - name: Lecture de DHIS2 et publication
        run: node ordinateur/publier.js
        env:
          MVE_MOT_DE_PASSE: \${{ secrets.MVE_MOT_DE_PASSE }}
          MVE_DHIS2_JETON: \${{ secrets.MVE_DHIS2_JETON }}
          MVE_GITHUB_JETON: \${{ secrets.MVE_GITHUB_JETON }}
          GITHUB_TOKEN: \${{ secrets.GITHUB_TOKEN }}
`;
const LISEZ_DEPOT = `# Suivi MVE – publication automatique

Ce dépôt contient :
- la branche **main** : le programme de publication et la tâche planifiée (.github/workflows/publier.yml) ;
- la branche **gh-pages** : le site publié (applications Suivi MVE et Qualité MVE), régénéré à chaque publication.

Toutes les 30 minutes, GitHub lit DHIS2, chiffre les données avec le mot de passe d'équipe et met le site à jour.
Les secrets (jeton DHIS2, mot de passe d'équipe) sont dans Settings › Secrets and variables › Actions.
Publier à la demande : onglet Actions › Publication Suivi MVE › Run workflow.
Aucune donnée en clair n'est déposée ici.
`;
async function installerNuage(c) {
  const d = '/repos/' + c.github_depot;
  log('2/3 Le site est maintenant publié depuis la branche « ' + BRANCHE_SITE + ' »…');
  let pagesOk = false;
  try {
    const p = await github(c.github_jeton, 'GET', d + '/pages');
    if (p.source && p.source.branch === BRANCHE_SITE) pagesOk = true;
    else { await github(c.github_jeton, 'PUT', d + '/pages', { source: { branch: BRANCHE_SITE, path: '/' } }); pagesOk = true; }
  } catch (e) { pagesOk = false; }
  if (!pagesOk) {
    log('\n⚠ Étape manuelle (une fois) : ouvrez https://github.com/' + c.github_depot + '/settings/pages');
    log('  › Branch : choisissez « ' + BRANCHE_SITE + ' » puis « / (root) » › Save.');
    log('  Ensuite relancez :  node publier.js --installer-nuage');
    log('  (Le programme ne remplace pas la branche « main » tant que le site n’est pas servi depuis « ' + BRANCHE_SITE + ' ».)\n');
    return;
  }
  log('3/3 Dépôt du programme et de la tâche planifiée sur la branche « main »…');
  const cfg = {};
  for (const k of CLES_NUAGE) if (c[k] !== undefined) cfg[k] = c[k];
  const lire = (p) => fs.readFileSync(p);
  const fichiers = fichiersDuSite(SITE).map((f) => ({ nom: 'site/' + f.nom, data: f.data })).concat([
    { nom: 'ordinateur/publier.js', data: lire(__filename) },
    { nom: 'ordinateur/config-nuage.json', data: Buffer.from(JSON.stringify(cfg, null, 2)) },
    ...(fs.existsSync(path.join(ICI, 'population-aires.xlsx')) ? [{ nom: 'ordinateur/population-aires.xlsx', data: lire(path.join(ICI, 'population-aires.xlsx')) }] : []),
    { nom: '.github/workflows/publier.yml', data: Buffer.from(WORKFLOW) },
    { nom: 'README.md', data: Buffer.from(LISEZ_DEPOT) }
  ]);
  if (fs.existsSync(path.join(ICI, '..', 'LISEZ-MOI.txt'))) fichiers.push({ nom: 'LISEZ-MOI.txt', data: lire(path.join(ICI, '..', 'LISEZ-MOI.txt')) });
  try {
    await publierGithub(Object.assign({}, c, { github_branche: 'main', sansAttente: true }), fichiers, new Date().toISOString(), { sansNojekyll: true, message: 'Programme de publication automatique' });
  } catch (e) {
    if (e.status === 403 || e.status === 404 || /workflow/i.test(e.message)) stop('GitHub refuse le dépôt de la tâche planifiée (' + e.message + ').\n  Ajoutez au jeton GitHub l’autorisation « Workflows : Read and write » (github.com › Settings › Developer settings ›\n  Fine-grained tokens › votre jeton › Edit), puis relancez :  node publier.js --installer-nuage');
    throw e;
  }
  log('\n✓ Publication automatique installée.');
  log('  Dernière étape (une fois) : https://github.com/' + c.github_depot + '/settings/secrets/actions › New repository secret :');
  log('    MVE_DHIS2_JETON   = votre jeton d’accès personnel DHIS2 (commence par d2p_)');
  log('    MVE_MOT_DE_PASSE  = le mot de passe d’équipe (le même que sur les téléphones)');
  log('  Puis : onglet Actions › « Publication Suivi MVE » › Run workflow, pour vérifier.');
  log('  Ensuite, GitHub publie seul toutes les 30 minutes : désactivez la tâche planifiée Windows.\n');
}

/* ---------- Programme principal ---------- */
(async () => {
  const vNode = +process.versions.node.split('.')[0];
  if (vNode < 18) stop('Node.js ' + process.versions.node + ' est trop ancien : installez la version LTS depuis nodejs.org.');

  let config;
  if (EN_NUAGE) {
    config = configNuage();
  } else {
    config = fs.existsSync(CONFIG) ? JSON.parse(fs.readFileSync(CONFIG, 'utf8')) : null;
    if (config && !config.hebergeur) config.hebergeur = 'netlify'; // ancienne configuration
  }
  const incomplet = (c) => !c || !c.source || !c.mot_de_passe ||
    (c.hebergeur === 'github' ? !c.github_depot || !c.github_jeton : !c.jeton || !c.site_id) ||
    (c.source === 'excel' && !c.classeur) || (c.source === 'dhis2' && (!c.dhis2_utilisateur || !c.dhis2_mot_de_passe));
  if (INSTALLER_NUAGE) {
    if (incomplet(config)) stop('Configurez d’abord la publication sur cet ordinateur (« node publier.js --config »).');
    if (config.hebergeur !== 'github') stop('La publication automatique sur GitHub nécessite l’hébergeur GitHub Pages (« node publier.js --config », choix 1).');
    if (config.source !== 'dhis2') stop('La publication automatique sur GitHub nécessite la source DHIS2 (« node publier.js --config », choix 1) : GitHub ne peut pas lire un classeur Excel de cet ordinateur.');
    // Le site sera désormais publié sur la branche « gh-pages » ; « main » contiendra le programme et la tâche planifiée.
    config.github_branche = BRANCHE_SITE;
    fs.writeFileSync(CONFIG, JSON.stringify(config, null, 2), { mode: 0o600 });
    config.sansAttente = true; // la bascule de Pages est faite juste après par installerNuage (non enregistré)
    log('\n=== Installation de la publication automatique sur GitHub ===\n1/3 Publication du site sur la branche « ' + BRANCHE_SITE + ' »…');
  }
  if (!EN_NUAGE && (args.includes('--config') || incomplet(config))) {
    if (!process.stdout.isTTY || !process.stdin.isTTY) stop('Configuration à compléter : double-cliquez une fois sur « Publier.bat » et répondez aux questions.');
    config = await configurer(config || {});
  }

  let lu;
  if (config.source === 'dhis2') {
    lu = await lireDHIS2(config);
    log('✓ ' + lu.n + ' alertes lues dans DHIS2 (' + serieEnTexte(lu.extrait) + ')');
  } else {
    log('Lecture de « ' + path.basename(config.classeur) + ' »…');
    lu = lireClasseur(config.classeur);
    log('✓ ' + lu.n + ' alertes, extraction DHIS2 du ' + serieEnTexte(lu.extrait));
  }
  const { table } = lu;
  const publie = new Date().toISOString();

  const enc = chiffrer({ v: 1, publie, table }, config.mot_de_passe, config.sel);
  log('✓ Données Suivi MVE chiffrées (' + Math.round(enc.length / 1024) + ' Ko)');
  const fichiers = fichiersDuSite(SITE).concat([{ nom: 'donnees.enc', data: enc }]);

  // Fond de carte (carte des aires de santé de « Suivi MVE ») : un échec ne bloque pas la publication
  if (config.source === 'dhis2') {
    try {
      const carte = await lireCarte(config);
      if (carte) {
        const json = Buffer.from(JSON.stringify(carte));
        fichiers.push({ nom: 'carte.json', data: json });
        log('✓ Fond de carte : ' + carte.zones.length + ' zones, ' + carte.aires.length + ' aires de santé' +
          (carte.popAnnee ? ', population ' + carte.popAnnee : ', sans population') + ' (' + Math.round(json.length / 1024) + ' Ko)');
      }
    } catch (e) {
      log('  ⚠ Fond de carte non publié (' + (e.message || e) + ') : les téléphones gardent la dernière carte reçue.');
    }
  }

  // Contrôles de qualité (application « Qualité MVE »)
  let q = lu.qualite || null;
  if (!q && config.source === 'excel' && config.classeur_qualite) {
    log('Lecture de « ' + path.basename(config.classeur_qualite) + ' »…');
    q = lireClasseurQualite(config.classeur_qualite);
  }
  let encQ = null;
  if (q) {
    encQ = chiffrer({ v: 1, publie, seuils: q.seuils, table: q.table }, config.mot_de_passe, config.sel);
    fichiers.push({ nom: 'qualite/qualite.enc', data: encQ });
    log('✓ Contrôles de qualité chiffrés : ' + q.n + ' alertes' + (q.sansDate ? ' (+ ' + q.sansDate + ' sans date de notification)' : '') +
      ', noms remplacés par une clé anonyme (' + Math.round(encQ.length / 1024) + ' Ko)');
  } else {
    log('  (contrôles de qualité non publiés' + (config.source === 'excel' ? ' : aucun classeur Qualité configuré' : '') + ')');
  }

  // Registre de vaccination (application « Vaccination MVE ») : un échec ne bloque pas les autres publications
  if (config.source === 'dhis2') {
    try {
      const v = await lireVaccination(config);
      if (v) {
        const encV = chiffrer({ v: 1, publie, seuils: v.parametres, table: v.table }, config.mot_de_passe, config.sel);
        fichiers.push({ nom: 'vaccination/vaccination.enc', data: encV });
        log('✓ Registre de vaccination chiffré : ' + v.nVac + ' vaccinations, ' + v.n + ' fiches, sans nom ni téléphone (' + Math.round(encV.length / 1024) + ' Ko)');
      }
    } catch (e) {
      log('  ⚠ Registre de vaccination non publié (' + (e.message || e) + ') : les autres applications sont publiées normalement.');
    }
  }

  // Suivi des contacts (application « Contacts MVE ») : un échec ne bloque pas les autres publications
  let kContacts = null;
  if (config.source === 'dhis2') {
    try {
      const k = await lireContacts(config);
      kContacts = k;
      if (k) {
        const encK = chiffrer({ v: 1, publie, contacts: k.contacts, suivis: k.suivis }, config.mot_de_passe, config.sel);
        fichiers.push({ nom: 'contacts/contacts.enc', data: encK });
        log('✓ Suivi des contacts chiffré : ' + k.n + ' contacts, ' + k.nS + ' jours de suivi, sans nom ni téléphone (' + Math.round(encK.length / 1024) + ' Ko)');
      }
    } catch (e) {
      log('  ⚠ Suivi des contacts non publié (' + (e.message || e) + ') : les autres applications sont publiées normalement.');
    }
  }

  // Chaînes de transmission : liens calculés ici (noms en mémoire), publiés sans aucun nom
  if (kContacts && lu && lu.alertesChaines) {
    try {
      const CH = require(path.join(SITE, 'chaines', 'engine-chaines.js'));
      const D = CH.charger(lu.alertesChaines, kContacts.contactsNoms), P = CH.exporter(D);
      const encC = chiffrer(Object.assign({ v: 1, publie }, P), config.mot_de_passe, config.sel);
      fichiers.push({ nom: 'chaines/chaines.enc', data: encC });
      const liens = D.aretes.filter((e) => !e.ignore).length;
      log('✓ Chaînes de transmission chiffrées : ' + liens + ' liens entre cas, ' + D.chaines.filter((x) => x.generations >= 2).length + ' chaînes de 2 générations ou plus, sans aucun nom (' + Math.round(encC.length / 1024) + ' Ko)');
    } catch (e) {
      log('  ⚠ Chaînes de transmission non publiées (' + (e.message || e) + ') : les autres applications sont publiées normalement.');
    }
  }

  const archive = zip(fichiers);
  if (MODE_TEST) {
    fs.writeFileSync(path.join(ICI, 'test-donnees.enc'), enc);
    if (encQ) fs.writeFileSync(path.join(ICI, 'test-qualite.enc'), encQ);
    fs.writeFileSync(path.join(ICI, 'test-site.zip'), archive);
    log('Mode test : rien n’a été publié (fichiers test-donnees.enc et test-site.zip créés).');
    return;
  }
  if (config.hebergeur === 'github') {
    log('Publication sur GitHub Pages (' + config.github_depot + ')…');
    let etat;
    try {
      etat = await publierGithub(config, fichiers, publie);
    } catch (e) {
      if (e.reseau) stop(e.message);
      if (e.status === 401) stop('Jeton GitHub refusé ou expiré. Créez-en un nouveau puis relancez « node publier.js --config ».');
      if (e.status === 409) stop('Le dépôt « ' + config.github_depot + ' » est totalement vide : sur github.com, ajoutez-y un fichier (bouton « Add a README ») puis relancez.');
      if (e.status === 403 || e.status === 404) stop('GitHub refuse l’écriture dans « ' + config.github_depot + ' » (' + e.message + '). Vérifiez l’autorisation « Contents : Read and write » du jeton.');
      stop(e.message);
    }
    if (etat === 'erreur') stop('GitHub Pages a signalé une erreur de mise en ligne. Ouvrez l’onglet « Actions » du dépôt pour le détail.');
    if (etat !== 'ok') {
      const actif = await activerPages(config);
      if (!actif) log('\n⚠ Vérifiez une fois que GitHub Pages est activé : dépôt › Settings › Pages › Source « Deploy from a branch »,\n  branche « ' + (config.github_branche || 'main') + ' », dossier « / (root) ».');
    }
    log('\n✓ Publié' + (etat === 'ok' ? '' : ' (mise en ligne en cours : une à deux minutes)') + '. Les téléphones recevront ces données à leur prochaine ouverture.');
    log('  Adresse de l’application : ' + adresseGithub(config.github_depot));
    log('  Application Qualité : ' + adresseGithub(config.github_depot) + 'qualite/');
    log('  Application Vaccination : ' + adresseGithub(config.github_depot) + 'vaccination/');
    log('  Application Contacts : ' + adresseGithub(config.github_depot) + 'contacts/');
    log('  Application Qualité Contacts : ' + adresseGithub(config.github_depot) + 'qualite-contacts/');
    log('  Chaînes de transmission : ' + adresseGithub(config.github_depot) + 'chaines/\n');
    if (INSTALLER_NUAGE) await installerNuage(config);
    return;
  }
  log('Publication sur Netlify…');
  try {
    const dep = await netlify(config.jeton, 'POST', '/sites/' + config.site_id + '/deploys', archive, 'application/zip');
    // attendre la mise en ligne
    let etat = dep.state, essais = 0;
    while (!['ready', 'error'].includes(etat) && essais++ < 30) {
      await new Promise((r) => setTimeout(r, 2000));
      etat = (await netlify(config.jeton, 'GET', '/deploys/' + dep.id)).state;
    }
    if (etat === 'error') stop('Netlify a signalé une erreur de publication. Réessayez dans quelques minutes.');
    log('\n✓ Publié. Les téléphones recevront ces données à leur prochaine ouverture.');
    log('  Adresse de l’application : ' + (config.adresse || dep.ssl_url || dep.url) + '\n');
  } catch (e) {
    if (e.status === 401) stop('Jeton Netlify refusé ou expiré. Relancez « node publier.js --config ».');
    if (e.status === 403 && /credit/i.test(e.message)) stop('Crédits Netlify épuisés : publications bloquées. Passez à GitHub Pages avec « node publier.js --config » (choix 1).');
    if (e.status === 404) stop('Site Netlify introuvable. Relancez « node publier.js --config » après avoir vidé "site_id" dans config.json.');
    if (e.cause || /fetch failed/i.test(e.message)) stop('Pas de connexion Internet (ou accès à api.netlify.com bloqué). Réessayez plus tard.');
    stop(e.message);
  }
})().catch((e) => stop(e.message || String(e)));
