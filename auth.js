/*
 * auth.js — control de acceso del portal Supervisor (COBRA).
 *
 * IMPORTANTE — alcance real de esta protección:
 * El sitio es estático (GitHub Pages), así que esto es una BARRERA, no
 * seguridad de verdad. Frena el acceso casual (alguien a quien le reenvían
 * el link), pero una persona técnica puede leer este archivo, ver los hashes
 * y abrir el contenido. Si el dato pasa a ser sensible, hay que mover el
 * sitio detrás de algo tipo Cloudflare Access.
 *
 * Las claves NO se guardan en texto plano: se guarda el SHA-256 de
 * "usuario:clave". Para cambiar/agregar usuarios, calcular el hash con:
 *   node -e "console.log(require('crypto').createHash('sha256').update('usuario:clave').digest('hex'))"
 * y editar el objeto USUARIOS de abajo.
 */
(function () {
  'use strict';

  var SESSION_KEY = 'supAuth_v1';

  // usuario -> SHA-256 hex de (usuario + ':' + clave)
  var USUARIOS = {
    rcerda:      '91d50f2230e30ebe4bdb8bbb1e4cae66f6049353a0960a1443a5a9e6c9dca4bd',
    tcontreras:  '84a063ef1d39ad6df9578d0389a3eb6d472d2769973d91494aa6473df028fee8',
    jvodnizza:   '165bc0849067edd4c6cad7f79fabbdc93cfe4169738fdf74e8caf8996dd48f20',
    amanrriquez: '5e3020f9f2c77ad69039767c6921ae00760010ffc26dc13400a0c2e592cf0e05',
    jbaez:       '4c444b9fe4bb2ece193c6ddafadca08b0c1e226cc087fc171f68c2e33544baf1',
    cquiroz:     '790cc118f6e5c7a78b7b4016feb1212978f42fb048f05634373b9e7cb8c24ccb',
    fflores:     'ae89c196e6c42f781261f0a910c2a59c1b4d938a3de19279eeed58862978d0eb'
  };

  function ocultar()  { try { document.documentElement.style.visibility = 'hidden'; } catch (e) {} }
  function revelar()  { try { document.documentElement.style.visibility = '';       } catch (e) {} }

  function sesionValida() {
    try {
      var d = JSON.parse(sessionStorage.getItem(SESSION_KEY) || 'null');
      return !!(d && d.u && USUARIOS[d.u] && d.t === USUARIOS[d.u]);
    } catch (e) { return false; }
  }

  function cerrarSesion() {
    try { sessionStorage.removeItem(SESSION_KEY); } catch (e) {}
    location.reload();
  }
  window.supCerrarSesion = cerrarSesion;

  function sha256Hex(str) {
    var bytes = new TextEncoder().encode(str);
    return crypto.subtle.digest('SHA-256', bytes).then(function (buf) {
      var out = '';
      var arr = new Uint8Array(buf);
      for (var i = 0; i < arr.length; i++) out += ('0' + arr[i].toString(16)).slice(-2);
      return out;
    });
  }

  function inyectarBotonSalir(usuario) {
    if (document.getElementById('supLogoutBtn')) return;
    var b = document.createElement('button');
    b.id = 'supLogoutBtn';
    b.type = 'button';
    b.textContent = 'Cerrar sesión' + (usuario ? ' · ' + usuario : '');
    b.style.cssText = [
      'position:fixed', 'top:10px', 'right:12px', 'z-index:2147483646',
      'background:rgba(255,255,255,.92)', 'border:1px solid #e0e5ea', 'border-radius:20px',
      'padding:5px 12px', 'font:600 11.5px \'Segoe UI\',Arial,sans-serif', 'color:#003c71',
      'cursor:pointer', 'box-shadow:0 2px 8px rgba(20,50,80,.12)'
    ].join(';');
    b.addEventListener('click', cerrarSesion);
    document.body.appendChild(b);
  }

  function construirLogin() {
    if (document.getElementById('supLoginOverlay')) return;
    var ov = document.createElement('div');
    ov.id = 'supLoginOverlay';
    ov.style.cssText = [
      'visibility:visible', 'position:fixed', 'inset:0', 'z-index:2147483647',
      'background:linear-gradient(120deg,#eef1f4 0%,#e8f6fd 55%,#dcf1fb 100%)',
      'display:flex', 'align-items:center', 'justify-content:center', 'padding:20px',
      'font-family:\'Segoe UI\',Arial,sans-serif'
    ].join(';');
    ov.innerHTML =
      '<form id="supLoginForm" autocomplete="on" style="background:#fff;padding:32px 30px;border-radius:16px;' +
        'box-shadow:0 20px 50px rgba(0,60,113,.18);width:min(92vw,360px);">' +
        '<div style="font-size:12.5px;letter-spacing:.14em;text-transform:uppercase;color:#29a9e0;font-weight:800;">Calidad &amp; Capacitacion</div>' +
        '<h1 style="margin:4px 0 2px;font-size:24px;color:#003c71;font-weight:800;">Supervisor</h1>' +
        '<p style="margin:0 0 20px;font-size:13px;color:#6b7a8c;line-height:1.5;">Ingresa tus credenciales para ver los informes.</p>' +
        '<label for="supU" style="display:block;font-size:12px;font-weight:700;color:#22303f;margin-bottom:4px;">Usuario</label>' +
        '<input id="supU" name="username" autocomplete="username" autocapitalize="none" autocorrect="off" spellcheck="false" ' +
          'style="width:100%;padding:10px 12px;border:1px solid #e0e5ea;border-radius:9px;font-size:14px;margin-bottom:14px;">' +
        '<label for="supP" style="display:block;font-size:12px;font-weight:700;color:#22303f;margin-bottom:4px;">Clave</label>' +
        '<input id="supP" name="password" type="password" autocomplete="current-password" ' +
          'style="width:100%;padding:10px 12px;border:1px solid #e0e5ea;border-radius:9px;font-size:14px;margin-bottom:8px;">' +
        '<div id="supErr" role="alert" style="min-height:18px;font-size:12.5px;color:#c0392b;margin-bottom:8px;"></div>' +
        '<button type="submit" id="supBtn" style="width:100%;padding:11px;border:none;border-radius:9px;background:#0071ce;' +
          'color:#fff;font-size:14px;font-weight:800;cursor:pointer;">Entrar</button>' +
      '</form>' +
      '<img src="logo-cobra-t.png" alt="Cobra" style="position:absolute;top:22px;left:26px;height:34px;width:auto;">';
    document.body.appendChild(ov);

    var form = document.getElementById('supLoginForm');
    var errEl = document.getElementById('supErr');
    var btn = document.getElementById('supBtn');

    form.addEventListener('submit', function (ev) {
      ev.preventDefault();
      errEl.textContent = '';
      var u = (document.getElementById('supU').value || '').trim().toLowerCase();
      var p = document.getElementById('supP').value || '';
      if (!u || !p) { errEl.textContent = 'Completa usuario y clave.'; return; }
      if (!USUARIOS[u]) { errEl.textContent = 'Usuario o clave incorrectos.'; return; }
      btn.disabled = true;
      sha256Hex(u + ':' + p).then(function (hex) {
        if (hex === USUARIOS[u]) {
          try {
            sessionStorage.setItem(SESSION_KEY, JSON.stringify({ u: u, t: USUARIOS[u], ts: Date.now() }));
          } catch (e) {}
          ov.remove();
          revelar();
          inyectarBotonSalir(u);
        } else {
          btn.disabled = false;
          errEl.textContent = 'Usuario o clave incorrectos.';
        }
      }).catch(function () {
        btn.disabled = false;
        errEl.textContent = 'No se pudo validar. Reintenta.';
      });
    });

    document.getElementById('supU').focus();
  }

  function alDOM(fn) {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', fn);
    else fn();
  }

  if (sesionValida()) {
    revelar();
    var usuarioActual;
    try { usuarioActual = JSON.parse(sessionStorage.getItem(SESSION_KEY)).u; } catch (e) {}
    alDOM(function () { inyectarBotonSalir(usuarioActual); });
  } else {
    ocultar();
    alDOM(construirLogin);
  }
})();
