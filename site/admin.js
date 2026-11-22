/* Actions réservées à l'administrateur (les six applications).
   Sur un téléphone où la mise à jour automatique est activée, ces commandes demandent le code administrateur :
     - « Changer le mot de passe d'équipe »
     - « Désactiver la mise à jour automatique »
     - « Effacer les données de ce téléphone »
   L'activation reste libre quand aucune clé n'est enregistrée (première installation, ou après un
   changement du mot de passe d'équipe : l'application le redemande alors d'elle-même).
   Le code n'est jamais stocké en clair : admin-config.js contient seulement son empreinte PBKDF2-SHA-256,
   produite sur l'ordinateur par « node code-admin.js » (ce fichier n'est pas livré dans le zip, pour ne pas
   écraser votre code quand vous remplacez le dossier site). Après un code correct, les actions restent
   ouvertes 5 minutes dans l'application. */
(function () {
  'use strict';
  var CFG = window.MVE_ADMIN || null;
  var ouvertJusqua = 0, DUREE = 5 * 60000, essais = 0, bloqueJusqua = 0;

  function b64(buf) { var s = ''; new Uint8Array(buf).forEach(function (x) { s += String.fromCharCode(x); }); return btoa(s); }
  function deb64(t) { var s = atob(t), u = new Uint8Array(s.length); for (var i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; }
  function verifier(code) {
    if (!CFG || !window.crypto || !crypto.subtle) return Promise.resolve(false);
    return crypto.subtle.importKey('raw', new TextEncoder().encode(code), 'PBKDF2', false, ['deriveBits']).then(function (k) {
      return crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: deb64(CFG.sel), iterations: CFG.iter }, k, 256);
    }).then(function (bits) { return b64(bits) === CFG.empreinte; }).catch(function () { return false; });
  }

  var css = '#adm-bg{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:99999;display:flex;align-items:center;justify-content:center;padding:16px}' +
    '#adm{background:#fff;color:#1b1f23;border-radius:14px;max-width:360px;width:100%;padding:18px 18px 14px;font:15px/1.4 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;box-shadow:0 10px 40px rgba(0,0,0,.3)}' +
    '#adm h3{margin:0 0 6px;font-size:17px}#adm p{margin:0 0 10px;color:#4a5560;font-size:14px}' +
    '#adm input[type=password],#adm input[type=text]{width:100%;box-sizing:border-box;font-size:16px;padding:10px 12px;border:1px solid #b8c2cc;border-radius:9px;margin:2px 0 8px}' +
    '#adm label{font-size:13.5px;color:#4a5560;display:flex;gap:6px;align-items:center}' +
    '#adm .m{color:#b3261e;font-size:13.5px;min-height:18px;margin:4px 0}' +
    '#adm .b{display:flex;gap:8px;justify-content:flex-end;margin-top:6px}' +
    '#adm button{font:inherit;font-weight:600;border-radius:9px;padding:9px 14px;border:0;cursor:pointer}' +
    '#adm .ok{background:#12393d;color:#fff}#adm .no{background:#e8edf0;color:#1b1f23}' +
    '@media (prefers-color-scheme: dark){#adm{background:#1f2629;color:#e8edf0}#adm p,#adm label{color:#aab4bb}' +
    '#adm input[type=password],#adm input[type=text]{background:#12181a;color:#e8edf0;border-color:#3a4549}#adm .no{background:#2d3538;color:#e8edf0}#adm .ok{background:#2f8f83}}';

  function dialogue(action, suite) {
    if (!document.getElementById('adm-css')) { var st = document.createElement('style'); st.id = 'adm-css'; st.textContent = css; document.head.appendChild(st); }
    var bg = document.createElement('div'); bg.id = 'adm-bg';
    bg.innerHTML = '<form id="adm" autocomplete="off"><h3>Action réservée à l’administrateur</h3>' +
      '<p></p><input type="password" id="adm-code" autocapitalize="off" spellcheck="false" aria-label="Code administrateur">' +
      '<label><input type="checkbox" id="adm-voir"> Afficher le code</label><div class="m" role="alert"></div>' +
      '<div class="b"><button type="button" class="no">Annuler</button><button type="submit" class="ok">Valider</button></div></form>';
    var f = bg.querySelector('form'), inp = bg.querySelector('#adm-code'), msg = bg.querySelector('.m');
    bg.querySelector('p').textContent = CFG
      ? '« ' + action + ' » nécessite le code administrateur. Si vous n’êtes pas l’administrateur, adressez-vous au gestionnaire des données.'
      : '« ' + action + ' » est réservée à l’administrateur, mais aucun code administrateur n’est encore configuré (LISEZ-MOI, section W).';
    if (!CFG) { inp.hidden = true; bg.querySelector('label').hidden = true; bg.querySelector('.ok').hidden = true; }
    function fermer() { bg.remove(); }
    bg.querySelector('.no').addEventListener('click', fermer);
    bg.addEventListener('click', function (e) { if (e.target === bg) fermer(); });
    bg.querySelector('#adm-voir').addEventListener('change', function (e) { inp.type = e.target.checked ? 'text' : 'password'; });
    f.addEventListener('submit', function (e) {
      e.preventDefault();
      if (Date.now() < bloqueJusqua) { msg.textContent = 'Trop d’essais. Réessayez dans une minute.'; return; }
      if (!inp.value) { msg.textContent = 'Saisissez le code administrateur.'; return; }
      msg.textContent = 'Vérification…';
      verifier(inp.value).then(function (ok) {
        if (ok) { essais = 0; ouvertJusqua = Date.now() + DUREE; fermer(); suite(); return; }
        essais++; if (essais >= 5) { essais = 0; bloqueJusqua = Date.now() + 60000; }
        msg.textContent = 'Code incorrect.'; inp.select();
      });
    });
    document.body.appendChild(bg);
    setTimeout(function () { if (!inp.hidden) inp.focus(); }, 30);
  }

  // Repère visible dans le menu : confirme que la protection est chargée sur ce téléphone.
  function repere() {
    var off = document.getElementById('btn-auto-off');
    if (!off || document.getElementById('adm-repere')) return;
    var p = document.createElement('p'); p.id = 'adm-repere';
    p.style.cssText = 'font-size:12.5px;opacity:.75;margin:8px 2px 4px';
    p.textContent = CFG ? 'Changer le mot de passe, désactiver la mise à jour et effacer les données : code administrateur requis.'
      : 'Code administrateur non configuré : ces commandes sont bloquées (LISEZ-MOI, section W).';
    off.parentNode.insertBefore(p, off);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', repere); else repere();

  // Interception avant les gestionnaires des applications (phase de capture).
  document.addEventListener('click', function (e) {
    var b = e.target && e.target.closest ? e.target.closest('#btn-auto, #btn-auto-off, #btn-clear') : null;
    if (!b || Date.now() < ouvertJusqua) return;
    var libelle = (b.textContent || '').trim();
    if (b.id === 'btn-auto' && !/^Changer/i.test(libelle)) return; // activation libre quand aucune clé n'est enregistrée
    if (b.id === 'btn-clear' && /^Confirmer/i.test(libelle)) return;  // 2e appui de confirmation, déjà autorisé
    e.preventDefault(); e.stopImmediatePropagation();
    dialogue(libelle.replace(/^Confirmer\s*:\s*/i, ''), function () { b.click(); });
  }, true);
})();
