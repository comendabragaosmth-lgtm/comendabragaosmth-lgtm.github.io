/* App instalável da Comenda de Braga (PWA)
 * 1. Regista o service worker (abrir rápido e sem rede).
 * 2. Convida a instalar a app no telemóvel (Android: botão; iPhone: instruções).
 * 3. «Entrar com o meu email»: código de 6 dígitos enviado por email, para quem
 *    abre a app instalada (no iPhone a app instalada não partilha o acesso do Safari).
 * Ficheiro independente: não altera o funcionamento da página principal.
 */
(function () {
  'use strict';
  var TK = 'comenda_token', LK = 'comenda_lang', DK = 'comenda_pwa_dispensa';
  var ls = {
    g: function (k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } },
    s: function (k, v) { try { v ? localStorage.setItem(k, v) : localStorage.removeItem(k); } catch (e) {} }
  };
  var ua = navigator.userAgent || '';
  var IOS = /iPhone|iPad|iPod/.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  var ANDROID = /Android/i.test(ua);
  var MOVEL = IOS || ANDROID;
  var INSTALADA = (window.matchMedia && matchMedia('(display-mode: standalone)').matches) || navigator.standalone === true;
  var BASE = location.origin + location.pathname;
  var promptInstalar = null;

  function EN() {
    var q = new URLSearchParams(location.search).get('lang');
    var l = q || ls.g(LK) || (/^pt/i.test(navigator.language || 'pt') ? 'pt' : 'en');
    return l === 'en' && new URLSearchParams(location.search).get('v') !== 'admin';
  }
  function T(pt, en) { return EN() ? en : pt; }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

  /* ---------- service worker ---------- */
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () {}); });
  }

  /* ---------- servidor ---------- */
  function chamar(fn) {
    var args = Array.prototype.slice.call(arguments, 1);
    return fetch(window.API, { method: 'POST', credentials: 'omit', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ fn: fn, args: args }) })
      .then(function (r) { if (!r.ok) throw new Error(); return r.json(); })
      .catch(function () { throw new Error(T('Sem ligação ao servidor. Verifique a internet e tente de novo.', 'No connection to the server. Please check your internet and try again.')); })
      .then(function (j) { if (!j.ok) throw new Error(String(j.error || '').replace(/^Error:\s*/, '')); return j.data; });
  }

  /* ---------- estilos ---------- */
  var css = document.createElement('style');
  css.textContent =
    '.pwa-bar{position:fixed;left:12px;right:12px;bottom:calc(12px + env(safe-area-inset-bottom,0px));z-index:45;background:var(--surface,#fff);color:var(--ink,#1d1a1b);border:1px solid var(--line,#e3dcdd);border-radius:14px;box-shadow:var(--shadow,0 10px 30px rgba(40,10,12,.12));padding:12px 14px;display:flex;gap:12px;align-items:center;max-width:520px;margin:0 auto}' +
    '.pwa-bar img{width:40px;height:40px;border-radius:10px;flex:none}.pwa-bar div{flex:1;min-width:0;font-size:.92rem;line-height:1.35}.pwa-bar b{display:block}' +
    '.pwa-x{border:0;background:none;font-size:1.3rem;line-height:1;color:var(--muted,#6a6264);cursor:pointer;padding:6px}' +
    '.pwa-modal{position:fixed;inset:0;z-index:70;background:rgba(10,5,6,.45);display:flex;align-items:flex-end;justify-content:center;padding:16px}' +
    '@media (min-width:600px){.pwa-modal{align-items:center}}' +
    '.pwa-box{background:var(--surface,#fff);color:var(--ink,#1d1a1b);border-radius:16px;width:min(460px,100%);padding:22px 20px calc(20px + env(safe-area-inset-bottom,0px));display:flex;flex-direction:column;gap:12px;position:relative;max-height:90vh;overflow:auto}' +
    '.pwa-box h2{font-size:1.15rem;margin:0;padding-right:28px}.pwa-box p{margin:0}.pwa-box .pwa-x{position:absolute;top:10px;right:10px}' +
    '.pwa-steps{margin:0;padding-left:22px;display:flex;flex-direction:column;gap:8px}.pwa-ico{display:inline-block;vertical-align:-4px;width:20px;height:20px}' +
    '.pwa-code{font-size:1.6rem;letter-spacing:.4em;text-align:center;font-variant-numeric:tabular-nums}' +
    '.pwa-err{color:var(--red,#b0001e);font-weight:600;font-size:.9rem}' +
    '.pwa-entrar{border-color:var(--gold,#8a6a1f)}';
  document.head.appendChild(css);

  var ICONE_PARTILHAR = '<svg class="pwa-ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3v12"/><path d="M8 7l4-4 4 4"/><path d="M5 11v9h14v-9"/></svg>';

  /* ---------- janela genérica ---------- */
  function abrir(html, aoAbrir) {
    fechar();
    var m = document.createElement('div');
    m.className = 'pwa-modal'; m.id = 'pwa-modal';
    m.innerHTML = '<div class="pwa-box" role="dialog" aria-modal="true"><button class="pwa-x" data-pwa="fechar" aria-label="' + T('Fechar', 'Close') + '">×</button>' + html + '</div>';
    m.addEventListener('click', function (e) { if (e.target === m) fechar(); });
    document.body.appendChild(m);
    if (aoAbrir) aoAbrir(m);
    var f = m.querySelector('input'); if (f) setTimeout(function () { f.focus(); }, 50);
  }
  function fechar() { var m = document.getElementById('pwa-modal'); if (m) m.remove(); }

  /* ---------- instalar ---------- */
  window.addEventListener('beforeinstallprompt', function (e) { e.preventDefault(); promptInstalar = e; mostrarBarra(); });
  window.addEventListener('appinstalled', function () { tirarBarra(); ls.s(DK, String(Date.now())); });

  function instalar() {
    if (promptInstalar) {
      promptInstalar.prompt();
      promptInstalar.userChoice.finally(function () { promptInstalar = null; tirarBarra(); });
      return;
    }
    if (IOS) {
      abrir('<h2>' + T('Instalar no iPhone / iPad', 'Install on iPhone / iPad') + '</h2>' +
        '<ol class="pwa-steps"><li>' + T('Carregue no botão <b>Partilhar</b> ', 'Tap the <b>Share</b> button ') + ICONE_PARTILHAR + T(' (em baixo no Safari, ou no topo no iPad).', ' (at the bottom in Safari, or at the top on iPad).') + '</li>' +
        '<li>' + T('Deslize e escolha <b>«Adicionar ao ecrã principal»</b>.', 'Scroll down and choose <b>“Add to Home Screen”</b>.') + '</li>' +
        '<li>' + T('Carregue em <b>Adicionar</b>. O ícone da Comenda aparece no ecrã do telemóvel.', 'Tap <b>Add</b>. The Comenda icon appears on your home screen.') + '</li>' +
        '<li>' + T('Abra a app pelo ícone e carregue em <b>«Entrar com o meu email»</b> para ficar com o seu acesso pessoal.', 'Open the app from the icon and tap <b>“Sign in with my email”</b> to get your personal access.') + '</li></ol>' +
        '<p class="small muted">' + T('Se não vir a opção, abra esta página no Safari. É preciso iOS 16.4 ou mais recente.', 'If you do not see the option, open this page in Safari. iOS 16.4 or later is required.') + '</p>' +
        '<button class="btn primary" data-pwa="fechar">' + T('Percebi', 'Got it') + '</button>');
      return;
    }
    abrir('<h2>' + T('Instalar a app', 'Install the app') + '</h2>' +
      '<ol class="pwa-steps"><li>' + T('Abra o menu do browser (<b>⋮</b>, no canto superior direito).', 'Open the browser menu (<b>⋮</b>, top right).') + '</li>' +
      '<li>' + T('Escolha <b>«Instalar app»</b> ou <b>«Adicionar ao ecrã principal»</b>.', 'Choose <b>“Install app”</b> or <b>“Add to Home screen”</b>.') + '</li>' +
      '<li>' + T('Confirme. O ícone da Comenda aparece no telemóvel.', 'Confirm. The Comenda icon appears on your phone.') + '</li></ol>' +
      '<p class="small muted">' + T('No computador, use o ícone de instalação na barra de endereço do Chrome ou do Edge.', 'On a computer, use the install icon in the Chrome or Edge address bar.') + '</p>' +
      '<button class="btn primary" data-pwa="fechar">' + T('Percebi', 'Got it') + '</button>');
  }

  function dispensado() { var t = +ls.g(DK); return t && Date.now() - t < 30 * 864e5; }
  function mostrarBarra() {
    if (INSTALADA || dispensado() || document.getElementById('pwa-bar')) return;
    if (!MOVEL && !promptInstalar) return;
    if (new URLSearchParams(location.search).get('v') === 'admin') return;
    var b = document.createElement('div');
    b.className = 'pwa-bar'; b.id = 'pwa-bar';
    b.innerHTML = '<img src="icon-192.png" alt=""><div><b>' + T('Instale a app da Comenda', 'Install the Comenda app') + '</b>' + T('Abre num toque, como as outras apps do telemóvel.', 'Opens in one tap, like your other apps.') + '</div>' +
      '<button class="btn primary sm" data-pwa="instalar">' + T('Instalar', 'Install') + '</button><button class="pwa-x" data-pwa="dispensar" aria-label="' + T('Agora não', 'Not now') + '">×</button>';
    document.body.appendChild(b);
  }
  function tirarBarra() { var b = document.getElementById('pwa-bar'); if (b) b.remove(); }

  /* ---------- entrar com código ---------- */
  var estado = { email: '' };
  function janelaEmail(msgErro) {
    abrir('<h2>' + T('Entrar com o meu email', 'Sign in with my email') + '</h2>' +
      '<p class="muted small">' + T('Use o email em que recebe os convites da Comenda. Vamos enviar-lhe um código de 6 dígitos.', 'Use the email where you receive the Comenda invitations. We will send you a 6-digit code.') + '</p>' +
      '<form class="stack" data-pwa-form="email"><label>Email<input type="email" id="pwa-email" required autocomplete="email" value="' + esc(estado.email) + '"></label>' +
      (msgErro ? '<p class="pwa-err">' + esc(msgErro) + '</p>' : '') +
      '<button class="btn primary">' + T('Enviar código', 'Send code') + '</button></form>');
  }
  function janelaCodigo(info, msgErro) {
    abrir('<h2>' + T('Introduza o código', 'Enter the code') + '</h2>' +
      '<p class="muted small">' + esc(info || '') + '</p>' +
      '<form class="stack" data-pwa-form="codigo"><label>' + T('Código de 6 dígitos', '6-digit code') + '<input type="text" id="pwa-codigo" class="pwa-code" required inputmode="numeric" pattern="[0-9]{6}" maxlength="6" autocomplete="one-time-code"></label>' +
      (msgErro ? '<p class="pwa-err">' + esc(msgErro) + '</p>' : '') +
      '<button class="btn primary">' + T('Entrar', 'Sign in') + '</button>' +
      '<p class="small"><a href="#" class="lnk" data-pwa="outro-email">' + T('Usar outro email ou pedir novo código', 'Use another email or request a new code') + '</a></p></form>');
  }
  function ocupado(form, sim) { var b = form.querySelector('button.primary'); if (b) { b.disabled = sim; } }

  document.addEventListener('submit', function (e) {
    var f = e.target.closest && e.target.closest('[data-pwa-form]');
    if (!f) return;
    e.preventDefault(); e.stopPropagation();
    var tipo = f.getAttribute('data-pwa-form');
    if (tipo === 'email') {
      estado.email = (document.getElementById('pwa-email').value || '').trim().toLowerCase();
      ocupado(f, true);
      chamar('apiEntrarPedirCodigo', estado.email, EN() ? 'en' : 'pt')
        .then(function (msg) { janelaCodigo(msg); })
        .catch(function (err) { janelaEmail(err.message); });
    } else if (tipo === 'codigo') {
      var cod = (document.getElementById('pwa-codigo').value || '').replace(/\D/g, '');
      ocupado(f, true);
      chamar('apiEntrarCodigo', estado.email, cod, EN() ? 'en' : 'pt')
        .then(function (r) {
          ls.s(TK, r.token);
          if (r.lingua === 'en' || r.lingua === 'pt') { if (!ls.g(LK)) ls.s(LK, r.lingua); }
          location.href = BASE + (INSTALADA ? '?app=1' : '');
        })
        .catch(function (err) { janelaCodigo(T('Enviámos o código para ', 'We sent the code to ') + estado.email + '.', err.message); });
    }
  }, true);

  document.addEventListener('click', function (e) {
    var el = e.target.closest && e.target.closest('[data-pwa]');
    if (!el) return;
    e.preventDefault(); e.stopPropagation();
    var a = el.getAttribute('data-pwa');
    if (a === 'fechar') fechar();
    else if (a === 'instalar') instalar();
    else if (a === 'dispensar') { ls.s(DK, String(Date.now())); tirarBarra(); }
    else if (a === 'entrar') janelaEmail();
    else if (a === 'outro-email') janelaEmail();
    else if (a === 'push-on') ativarPush();
    else if (a === 'push-off') desativarPush();
  }, true);

  /* ---------- avisos no telemóvel (push) ---------- */
  var PK = 'comenda_push', CK = 'comenda_push_cfg', FB = 'https://www.gstatic.com/firebasejs/10.14.1/';
  var PUSH_OK = 'serviceWorker' in navigator && 'PushManager' in window && 'Notification' in window;
  var push = { cfg: null, cfgLida: false, ocupado: false, msg: '' };
  function script(src) {
    return new Promise(function (ok, erro) {
      if (document.querySelector('script[src="' + src + '"]')) return ok();
      var s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = function () { erro(new Error(T('Não foi possível carregar o serviço de avisos.', 'Could not load the notification service.'))); };
      document.head.appendChild(s);
    });
  }
  function lerCfg() {
    if (push.cfgLida) return Promise.resolve(push.cfg);
    var c = null; try { c = JSON.parse(ls.g(CK) || 'null'); } catch (e) {}
    if (c) { push.cfg = c; push.cfgLida = true; }
    return chamar('apiPushConfig').then(function (r) { push.cfg = r; push.cfgLida = true; ls.s(CK, r ? JSON.stringify(r) : ''); return r; })
      .catch(function () { push.cfgLida = !!c; return c; });
  }
  function tokenFcm() {
    return lerCfg().then(function (cfg) {
      if (!cfg) throw new Error(T('Os avisos ainda não estão disponíveis.', 'Notifications are not available yet.'));
      return script(FB + 'firebase-app-compat.js').then(function () { return script(FB + 'firebase-messaging-compat.js'); }).then(function () {
        var fb = window.firebase;
        if (!fb.apps.length) fb.initializeApp({ apiKey: cfg.apiKey, projectId: cfg.projectId, messagingSenderId: cfg.messagingSenderId, appId: cfg.appId, authDomain: cfg.authDomain || (cfg.projectId + '.firebaseapp.com') });
        return navigator.serviceWorker.ready.then(function (reg) { return fb.messaging().getToken({ vapidKey: cfg.vapid, serviceWorkerRegistration: reg }); });
      });
    });
  }
  function plataforma() { return (IOS ? 'iOS' : ANDROID ? 'Android' : 'Computador') + (INSTALADA ? ' · app' : ''); }
  function ativarPush() {
    if (push.ocupado) return;
    push.ocupado = true; push.msg = ''; desenharPush();
    Notification.requestPermission().then(function (p) {
      if (p !== 'granted') throw new Error(T('Os avisos foram bloqueados. Pode autorizá-los nas definições do telemóvel (Notificações → Comenda).', 'Notifications were blocked. You can allow them in your phone settings (Notifications → Comenda).'));
      return tokenFcm();
    }).then(function (fcm) {
      return chamar('apiPushRegistar', ls.g(TK), fcm, plataforma(), EN() ? 'en' : 'pt').then(function (msg) { ls.s(PK, fcm); push.agora = true; push.msg = msg; });
    }).catch(function (e) { push.msg = e.message || String(e); })
      .then(function () { push.ocupado = false; desenharPush(); });
  }
  function desativarPush() {
    var fcm = ls.g(PK);
    ls.s(PK, ''); push.agora = false;
    push.msg = T('Avisos desativados neste telemóvel.', 'Notifications turned off on this phone.');
    desenharPush();
    if (fcm) chamar('apiPushRemover', ls.g(TK), fcm).catch(function () {});
    try { if (window.firebase && firebase.apps.length) firebase.messaging().deleteToken().catch(function () {}); } catch (e) {}
  }
  /* ao abrir a app: confirma que o token continua válido (o Firebase pode renová-lo) */
  function renovarPush() {
    var antigo = ls.g(PK);
    if (!PUSH_OK || !antigo || !ls.g(TK) || Notification.permission !== 'granted') return;
    tokenFcm().then(function (fcm) {
      if (fcm && fcm !== antigo) return chamar('apiPushRegistar', ls.g(TK), fcm, plataforma(), EN() ? 'en' : 'pt').then(function () { ls.s(PK, fcm); });
    }).catch(function () {});
  }
  function htmlPush() {
    var ativo = PUSH_OK && ls.g(PK) && (push.agora || Notification.permission === 'granted');
    var h = '<h2>🔔 ' + T('Avisos no telemóvel', 'Phone notifications') + '</h2>';
    if (!PUSH_OK && IOS && !INSTALADA) {
      h += '<p class="muted small" style="margin:0 0 12px">' + T('No iPhone, os avisos funcionam com a app instalada no ecrã principal.', 'On iPhone, notifications work once the app is installed on the home screen.') + '</p><button class="btn primary" data-pwa="instalar">' + T('Instalar a app', 'Install the app') + '</button>';
    } else if (!PUSH_OK) {
      h += '<p class="muted small" style="margin:0">' + T('Este browser não suporta avisos. Continua a receber tudo por email.', 'This browser does not support notifications. You will keep receiving everything by email.') + '</p>';
    } else if (ativo) {
      h += '<p class="small" style="margin:0 0 12px"><span class="calok">✓ ' + T('Ativos neste telemóvel', 'On for this phone') + '</span> <span class="muted">' + T('— convites, lembretes e alterações chegam aqui e por email.', '— invitations, reminders and changes arrive here and by email.') + '</span></p><button class="btn sm" data-pwa="push-off">' + T('Desativar', 'Turn off') + '</button>';
    } else {
      h += '<p class="muted small" style="margin:0 0 12px">' + T('Receba os convites, lembretes e alterações dos eventos como notificação, além do email.', 'Get event invitations, reminders and changes as notifications, as well as by email.') + '</p><button class="btn primary" data-pwa="push-on"' + (push.ocupado ? ' disabled' : '') + '>' + (push.ocupado ? T('A ativar…', 'Turning on…') : T('Ativar avisos', 'Turn on notifications')) + '</button>';
    }
    if (push.msg) h += '<p class="small" style="margin:10px 0 0">' + esc(push.msg) + '</p>';
    return h;
  }
  function desenharPush() { var s = document.getElementById('pwa-push'); if (s) s.innerHTML = htmlPush(); }

  /* ---------- acrescentos à página (refeitos a cada redesenho) ---------- */
  function acrescentar() {
    var app = document.getElementById('app');
    if (!app) return;
    var sub = app.querySelector('form[data-form="subscribe"]');
    if (sub && !document.getElementById('pwa-entrar')) {
      var painel = sub.closest('section');
      var s = document.createElement('section');
      s.className = 'panel pwa-entrar'; s.id = 'pwa-entrar';
      s.innerHTML = '<h2>' + T('Já recebe os convites?', 'Already receiving invitations?') + '</h2>' +
        '<p class="muted small" style="margin:0 0 12px">' + T('Entre com o seu email para ver as suas inscrições e os eventos reservados a membros.', 'Sign in with your email to see your registrations and members-only events.') + '</p>' +
        '<button class="btn primary" data-pwa="entrar">' + T('Entrar com o meu email', 'Sign in with my email') + '</button>';
      if (painel && painel.parentNode) painel.parentNode.insertBefore(s, painel);
    }
    var sair = app.querySelector('[data-act="forget"]');
    if (sair && !document.getElementById('pwa-push') && ls.g(TK) && (push.cfg || !push.cfgLida)) {
      var meu = sair.closest('section');
      var ps = document.createElement('section');
      ps.className = 'panel'; ps.id = 'pwa-push';
      if (meu && meu.parentNode) {
        meu.parentNode.insertBefore(ps, meu.nextSibling);
        if (push.cfgLida) desenharPush();
        else { ps.hidden = true; lerCfg().then(function (c) { var el = document.getElementById('pwa-push'); if (!el) return; if (c) { el.hidden = false; desenharPush(); } else el.remove(); }); }
      }
    }
    var foot = document.querySelector('footer.foot');
    if (foot && !INSTALADA && !document.getElementById('pwa-foot')) {
      var l = document.createElement('a');
      l.href = '#'; l.className = 'lnk'; l.id = 'pwa-foot'; l.setAttribute('data-pwa', 'instalar');
      l.textContent = T('📱 Instalar a app', '📱 Install the app');
      foot.insertBefore(l, document.getElementById('admlink'));
    }
  }

  function iniciar() {
    acrescentar();
    var app = document.getElementById('app');
    if (app && window.MutationObserver) new MutationObserver(acrescentar).observe(app, { childList: true });
    setTimeout(mostrarBarra, 2500);
    setTimeout(renovarPush, 6000);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar); else iniciar();
})();
