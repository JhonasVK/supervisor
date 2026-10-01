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
    fflores:     'bc541b07a9bb72311d725b4e847b811c1286f1aa3aeba5c1ad50773a53d48e0c',
    egonzalez:   '8d46510246affb8a24b5235f7e2153d2368af3870e31183b97083a334fe94d69'
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
      'background:#ffffff', 'border:1px solid #d9dbe0', 'border-radius:2px',
      'padding:7px 12px', 'font:600 11px \'Helvetica Neue\',\'Hanken Grotesk\',Helvetica,Arial,sans-serif', 'color:#003575',
      'letter-spacing:.1em', 'text-transform:uppercase', 'cursor:pointer'
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
      'background:linear-gradient(90deg,rgba(0,53,117,.96) 0%,rgba(0,53,117,.9) 40%,rgba(0,53,117,.55) 100%),url(fondo-portada.jpg) 55% 25%/cover no-repeat #003575',
      'display:flex', 'align-items:center', 'justify-content:center', 'padding:20px',
      'font-family:\'Helvetica Neue\',\'Hanken Grotesk\',Helvetica,Arial,sans-serif'
    ].join(';');
    var LBL = 'display:block;font-size:11px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#c9d7ee;margin-bottom:7px;';
    var INP = 'width:100%;padding:13px 14px;border:1px solid rgba(255,255,255,.3);border-radius:2px;background:rgba(255,255,255,.95);color:#334155;font-size:17px;font-family:inherit;';
    ov.innerHTML =
      '<form id="supLoginForm" autocomplete="on" style="background:rgba(0,34,77,.6);padding:34px 30px 30px;border:1px solid rgba(255,255,255,.18);' +
        'border-top:3px solid #f0a500;-webkit-backdrop-filter:blur(6px);backdrop-filter:blur(6px);width:min(92vw,420px);color:#fff;">' +
        '<div style="font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:#9fc2f0;font-weight:600;">Calidad &amp; Capacitaci&oacute;n</div>' +
        '<h1 style="margin:12px 0 8px;font-size:34px;line-height:1.1;color:#fff;font-weight:200;">Portal del <b style="font-weight:700;">Supervisor</b></h1>' +
        '<p style="margin:0 0 24px;font-size:16px;color:#c9d7ee;font-weight:300;line-height:1.5;">Ingresa tus credenciales para ver los informes.</p>' +
        '<label for="supU" style="' + LBL + '">Usuario</label>' +
        '<input id="supU" name="username" autocomplete="username" autocapitalize="none" autocorrect="off" spellcheck="false" ' +
          'style="' + INP + 'margin-bottom:18px;">' +
        '<label for="supP" style="' + LBL + '">Clave</label>' +
        '<input id="supP" name="password" type="password" autocomplete="current-password" ' +
          'style="' + INP + 'margin-bottom:10px;">' +
        '<div id="supErr" role="alert" style="min-height:20px;font-size:14px;color:#ffe2a8;margin-bottom:10px;"></div>' +
        '<button type="submit" id="supBtn" style="width:100%;padding:14px 22px;border:none;border-radius:2px;background:#f0a500;' +
          'color:#1d1a12;font-size:16px;font-weight:600;font-family:inherit;cursor:pointer;">Ingresar &rarr;</button>' +
      '</form>' +
      '<img src="logo-cobra.png" alt="Cobra" style="position:absolute;top:18px;left:22px;height:40px;width:auto;background:#fff;padding:6px 12px;">';
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
