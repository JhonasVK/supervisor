/*
 * auth.js — control de acceso del portal Supervisor (COBRA).
 *
 * Los usuarios y sus claves viven en Supabase (proyecto "portal-supervisor"),
 * ya no en este archivo. Cada usuario parte con la clave inicial y debe crear
 * la suya en el primer ingreso. Usuarios nuevos y restablecer clave: desde el
 * SQL Editor de Supabase:
 *   select public.crear_usuario('usuario', 'Nombre Apellido');
 *   select public.restablecer_clave('usuario');
 *
 * IMPORTANTE — alcance real de esta protección:
 * El sitio es estático (GitHub Pages), así que esto sigue siendo una BARRERA:
 * Supabase valida usuario y clave, pero los informes son archivos públicos.
 * Si el dato pasa a ser sensible, hay que mover el sitio detrás de algo tipo
 * Cloudflare Access.
 */
(function () {
  'use strict';

  var SESSION_KEY = 'supAuth_v2';
  // URL y clave publicable del proyecto: son públicas por diseño.
  var SB_URL = 'https://enoclwynrxaizjgvqhim.supabase.co';
  var SB_KEY = 'sb_publishable_aoJKG0x8aVtY6GSyWIwfCw_awWdTKKN';
  var DOMINIO = '@supervisores.portal.invalid';
  var CLAVE_MIN = 8;

  function ocultar()  { try { document.documentElement.style.visibility = 'hidden'; } catch (e) {} }
  function revelar()  { try { document.documentElement.style.visibility = '';       } catch (e) {} }

  function sesion() {
    try {
      var d = JSON.parse(sessionStorage.getItem(SESSION_KEY) || 'null');
      return d && d.u ? d : null;
    } catch (e) { return null; }
  }

  function cerrarSesion() {
    try { sessionStorage.removeItem(SESSION_KEY); } catch (e) {}
    location.reload();
  }
  window.supCerrarSesion = cerrarSesion;

  // Llamada a Supabase. Devuelve { ok, estado, datos }.
  function sb(metodo, ruta, cuerpo, token) {
    var cab = { 'apikey': SB_KEY, 'Content-Type': 'application/json' };
    if (token) cab['Authorization'] = 'Bearer ' + token;
    return fetch(SB_URL + ruta, { method: metodo, headers: cab, body: cuerpo ? JSON.stringify(cuerpo) : undefined })
      .then(function (r) {
        return r.text().then(function (t) {
          var datos = null;
          try { datos = t ? JSON.parse(t) : null; } catch (e) {}
          return { ok: r.ok, estado: r.status, datos: datos };
        });
      });
  }

  function dispositivo() {
    var ua = navigator.userAgent;
    return /iPad|Tablet/i.test(ua) ? 'Tablet' : /Mobi|Android|iPhone/i.test(ua) ? 'Celular' : 'Computador';
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

  /* ---------- Registro de ingresos (solo admin) ---------- */
  // Entrega un token vigente del admin; lo renueva si ya venció.
  function tokenAdmin() {
    var d = sesion();
    if (!d || !d.at) return Promise.reject(new Error('sesion'));
    if (Date.now() < d.exp - 60000) return Promise.resolve(d.at);
    return sb('POST', '/auth/v1/token?grant_type=refresh_token', { refresh_token: d.rt }).then(function (r) {
      if (!r.ok || !r.datos || !r.datos.access_token) throw new Error('sesion');
      d.at = r.datos.access_token; d.rt = r.datos.refresh_token; d.exp = Date.now() + (r.datos.expires_in || 3600) * 1000;
      try { sessionStorage.setItem(SESSION_KEY, JSON.stringify(d)); } catch (e) {}
      return d.at;
    });
  }

  function fechaHora(iso) {
    try { return new Date(iso).toLocaleString('es-CL', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); }
    catch (e) { return ''; }
  }

  function descargarCSV(filas) {
    var lineas = ['Fecha y hora;Usuario;Nombre;Dispositivo'].concat(filas.map(function (f) {
      return [fechaHora(f.fecha), f.usuario, f.nombre, f.dispositivo].map(function (v) { return '"' + String(v || '').replace(/"/g, '""') + '"'; }).join(';');
    }));
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob(['﻿' + lineas.join('\r\n')], { type: 'text/csv;charset=utf-8' }));
    a.download = 'ingresos-portal-supervisor-' + new Date().toISOString().slice(0, 10) + '.csv';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  var FUENTE = '\'Helvetica Neue\',\'Hanken Grotesk\',Helvetica,Arial,sans-serif';

  // En la portada (cabecera blanca) el usuario, Salir y Administrar van dentro de la cabecera,
  // como en la Academia Técnica. En los informes, que no tienen esa cabecera, quedan flotando.
  function inyectarBarra(d) {
    var admin = d.r === 'admin' && !!d.at;
    var fila = document.querySelector('header.top .brand-row');
    if (!fila) { inyectarBotonSalir(d.u); if (admin) inyectarBotonIngresos(); return; }
    if (document.getElementById('supBarra')) return;
    estilosBotones();
    fila.style.flexWrap = 'wrap';
    var barra = document.createElement('div');
    barra.id = 'supBarra';
    barra.innerHTML =
      '<nav class="sup-tabs" aria-label="Secciones"><a href="index.html" aria-current="page">Inicio</a>' +
        (admin ? '<button type="button" id="supIngresosBtn">Administrar</button>' : '') + '</nav>' +
      '<div class="sup-who"><span class="me"><b></b><small></small></span><button type="button" id="supLogoutBtn">Salir</button></div>';
    barra.querySelector('.me b').textContent = d.n || d.u;
    barra.querySelector('.me small').textContent = d.n ? d.u : '';
    fila.appendChild(barra);
    document.getElementById('supLogoutBtn').addEventListener('click', cerrarSesion);
    if (admin) document.getElementById('supIngresosBtn').addEventListener('click', abrirIngresos);
  }

  // Botones con el mismo formato de la Academia Técnica (.btn y .btn.primary).
  function estilosBotones() {
    if (document.getElementById('supBtnCss')) return;
    var st = document.createElement('style');
    st.id = 'supBtnCss';
    st.textContent =
      '.sup-btn{border:1px solid #003575;background:#ffffff;color:#003575;padding:12px 20px;border-radius:2px;' +
        'font:600 14px/1.2 \'Helvetica Neue\',\'Hanken Grotesk\',Helvetica,Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase;cursor:pointer}' +
      '.sup-btn:hover{background:#f2f5fa}' +
      '.sup-btn.primary{background:#003575;color:#ffffff}' +
      '.sup-btn.primary:hover{background:#00428f}' +
      '.sup-btn.chico{padding:8px 14px;font-size:12px}' +
      '.sup-btn:disabled{opacity:.55;cursor:default}' +
      '.sup-btn:focus-visible{outline:2px solid #f0a500;outline-offset:2px}' +
      // Barra de la cabecera, igual a la de la Academia Técnica (pestañas + usuario + Salir).
      '#supBarra{margin-left:auto;display:flex;flex-wrap:wrap;align-items:center;justify-content:flex-end;gap:10px 32px}' +
      '.sup-tabs{display:flex;gap:4px 26px;flex-wrap:wrap}' +
      '.sup-tabs a,.sup-tabs button{border:0;background:none;padding:8px 0;font:500 14px/1 ' + FUENTE + ';letter-spacing:.08em;' +
        'text-transform:uppercase;color:#62646e;border-bottom:2px solid transparent;text-decoration:none;cursor:pointer}' +
      '.sup-tabs a:hover,.sup-tabs button:hover{color:#003575}' +
      '.sup-tabs [aria-current="page"]{color:#003575;border-bottom-color:#f0a500}' +
      '.sup-who{display:flex;align-items:center;gap:14px}' +
      '.sup-who .me{display:flex;flex-direction:column;line-height:1.25;text-align:right}' +
      '.sup-who .me b{color:#003575;font:600 14px/1.25 ' + FUENTE + '}' +
      '.sup-who .me small{font:500 11px/1.3 "IBM Plex Mono",ui-monospace,Consolas,monospace;color:#62646e}' +
      '.sup-who button{border:1px solid #d9dbe0;background:#ffffff;color:#003575;padding:7px 12px;border-radius:2px;' +
        'font:600 11px/1 ' + FUENTE + ';letter-spacing:.1em;text-transform:uppercase;cursor:pointer}' +
      '.sup-who button:hover{border-color:#003575}';
    document.head.appendChild(st);
  }

  function abrirIngresos() {
    if (document.getElementById('supIngresos')) return;
    estilosBotones();
    var F = '\'Helvetica Neue\',\'Hanken Grotesk\',Helvetica,Arial,sans-serif';
    var ov = document.createElement('div');
    ov.id = 'supIngresos';
    ov.style.cssText = 'position:fixed;inset:0;z-index:2147483647;background:rgba(0,26,58,.6);display:flex;align-items:flex-start;justify-content:center;padding:40px 16px;overflow:auto;font-family:' + F;
    ov.innerHTML =
      '<div role="dialog" aria-modal="true" aria-labelledby="supIngTit" style="background:#fff;width:min(100%,860px);border-top:3px solid #f0a500;padding:28px 28px 24px;color:#334155;">' +
        '<div style="display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end;justify-content:space-between;margin-bottom:6px;">' +
          '<h2 id="supIngTit" style="margin:0;font-size:28px;font-weight:200;color:#003575;">Administración del <b style="font-weight:700;">portal</b></h2>' +
          '<div style="display:flex;gap:8px;">' +
            '<button type="button" id="supIngCsv" class="sup-btn primary" disabled>Descargar ingresos</button>' +
            '<button type="button" id="supIngCerrar" class="sup-btn">Cerrar</button>' +
          '</div></div>' +
        '<h3 style="margin:22px 0 4px;font-size:19px;font-weight:700;color:#003575;">Usuarios y claves</h3>' +
        '<p id="supUsuInfo" style="margin:0 0 14px;font-size:15px;color:#64748b;">Cargando…</p>' +
        '<p id="supUsuAviso" role="status" style="display:none;margin:0 0 14px;padding:10px 14px;border-left:3px solid #f0a500;background:#fff8e6;font-size:15px;"></p>' +
        '<div style="overflow:auto;"><table style="width:100%;border-collapse:collapse;font-size:15px;"><thead><tr id="supUsuCab"></tr></thead><tbody id="supUsuCuerpo"></tbody></table></div>' +
        '<h3 style="margin:30px 0 4px;font-size:19px;font-weight:700;color:#003575;">Registro de ingresos</h3>' +
        '<p id="supIngInfo" style="margin:0 0 14px;font-size:15px;color:#64748b;">Cargando…</p>' +
        '<div style="overflow:auto;"><table style="width:100%;border-collapse:collapse;font-size:15px;"><thead><tr id="supIngCab"></tr></thead><tbody id="supIngCuerpo"></tbody></table></div>' +
      '</div>';
    document.body.appendChild(ov);
    function cerrar() { ov.remove(); document.removeEventListener('keydown', tecla); }
    function tecla(e) { if (e.key === 'Escape') cerrar(); }
    document.addEventListener('keydown', tecla);
    document.getElementById('supIngCerrar').addEventListener('click', cerrar);
    ov.addEventListener('click', function (e) { if (e.target === ov) cerrar(); });
    document.getElementById('supIngCerrar').focus();

    var TH = 'text-align:left;padding:10px 12px;border-bottom:2px solid #003575;font-size:11px;letter-spacing:.1em;text-transform:uppercase;color:#003575;';
    var TD = 'padding:10px 12px;border-bottom:1px solid #e5e7eb;';
    function celda(tr, texto, estilo) { var td = document.createElement('td'); td.style.cssText = TD + (estilo || ''); td.textContent = texto; tr.appendChild(td); return td; }
    function cabecera(id, cols) {
      var fila = document.getElementById(id); fila.textContent = '';
      cols.forEach(function (c) { var th = document.createElement('th'); th.style.cssText = TH; th.textContent = c; fila.appendChild(th); });
    }

    // ----- Usuarios y claves -----
    var usuInfo = document.getElementById('supUsuInfo');
    var usuAviso = document.getElementById('supUsuAviso');
    var yo = (sesion() || {}).u;
    function avisar(texto) { usuAviso.textContent = texto; usuAviso.style.display = texto ? 'block' : 'none'; }
    function restablecer(u, boton) {
      if (!confirm('¿Restablecer la clave de ' + u + '?\n\nVolverá a la clave inicial y deberá crear una nueva al ingresar.')) return;
      boton.disabled = true;
      tokenAdmin().then(function (t) {
        return sb('POST', '/rest/v1/rpc/restablecer_clave', { p_usuario: u }, t);
      }).then(function (r) {
        if (!r.ok) throw new Error('rpc');
        avisar('Listo: ' + u + ' ya puede ingresar con la clave inicial y deberá crear una nueva.');
        cargarUsuarios();
      }).catch(function () {
        boton.disabled = false;
        avisar('No se pudo restablecer la clave de ' + u + '. Revisa tu conexión y reintenta.');
      });
    }
    function cargarUsuarios() {
      tokenAdmin().then(function (t) {
        return sb('GET', '/rest/v1/usuarios?select=usuario,nombre,rol,debe_cambiar_clave&order=usuario', null, t);
      }).then(function (r) {
        if (!r.ok) throw new Error('lectura');
        var lista = r.datos || [];
        var pendientes = lista.filter(function (x) { return x.debe_cambiar_clave; }).length;
        usuInfo.textContent = lista.length + ' usuarios · ' + pendientes + ' todavía con la clave inicial. Restablecer devuelve al usuario a la clave inicial; nadie puede ver las claves personales.';
        cabecera('supUsuCab', ['Usuario', 'Nombre', 'Rol', 'Clave', '']);
        var cuerpo = document.getElementById('supUsuCuerpo'); cuerpo.textContent = '';
        lista.forEach(function (x) {
          var tr = document.createElement('tr');
          celda(tr, x.usuario); celda(tr, x.nombre || '—'); celda(tr, x.rol === 'admin' ? 'Administrador' : 'Supervisor');
          celda(tr, x.debe_cambiar_clave ? 'Inicial (aún no la cambia)' : 'Personal', x.debe_cambiar_clave ? 'color:#9a6700;' : '');
          var td = celda(tr, '');
          if (x.usuario !== yo) {
            var b = document.createElement('button');
            b.type = 'button'; b.textContent = 'Restablecer';
            b.className = 'sup-btn chico';
            b.addEventListener('click', function () { restablecer(x.usuario, b); });
            td.appendChild(b);
          }
          cuerpo.appendChild(tr);
        });
      }).catch(function (e) {
        usuInfo.textContent = e && e.message === 'sesion' ? 'Tu sesión venció. Cierra sesión y vuelve a ingresar.' : 'No se pudo cargar la lista de usuarios. Revisa tu conexión y reintenta.';
      });
    }
    cargarUsuarios();

    // ----- Registro de ingresos -----
    var info = document.getElementById('supIngInfo');
    tokenAdmin().then(function (t) {
      return sb('GET', '/rest/v1/ingresos?select=fecha,dispositivo,usuarios(usuario,nombre)&order=fecha.desc&limit=2000', null, t);
    }).then(function (r) {
      if (!r.ok) throw new Error('lectura');
      var filas = (r.datos || []).map(function (x) {
        return { fecha: x.fecha, dispositivo: x.dispositivo || '', usuario: (x.usuarios && x.usuarios.usuario) || '', nombre: (x.usuarios && x.usuarios.nombre) || '' };
      });
      if (!filas.length) { info.textContent = 'Todavía no hay ingresos registrados.'; return; }
      info.textContent = 'Últimos ' + filas.length + ' ingresos, del más reciente al más antiguo. Solo los administradores ven este registro.';
      cabecera('supIngCab', ['Fecha y hora', 'Usuario', 'Nombre', 'Dispositivo']);
      var cuerpo = document.getElementById('supIngCuerpo');
      filas.forEach(function (f) {
        var tr = document.createElement('tr');
        [fechaHora(f.fecha), f.usuario, f.nombre || '—', f.dispositivo].forEach(function (v) { celda(tr, v); });
        cuerpo.appendChild(tr);
      });
      var csv = document.getElementById('supIngCsv');
      csv.disabled = false;
      csv.addEventListener('click', function () { descargarCSV(filas); });
    }).catch(function (e) {
      info.textContent = e && e.message === 'sesion' ? 'Tu sesión venció. Cierra sesión y vuelve a ingresar para ver el registro.' : 'No se pudo cargar el registro. Revisa tu conexión y reintenta.';
    });
  }

  function inyectarBotonIngresos() {
    if (document.getElementById('supIngresosBtn')) return;
    var b = document.createElement('button');
    b.id = 'supIngresosBtn';
    b.type = 'button';
    b.textContent = 'Administrar';
    estilosBotones();
    b.style.cssText = [
      'position:fixed', 'top:46px', 'right:12px', 'z-index:2147483646',
      'background:#ffffff', 'border:1px solid #d9dbe0', 'border-radius:2px',
      'padding:7px 12px', 'font:600 11px \'Helvetica Neue\',\'Hanken Grotesk\',Helvetica,Arial,sans-serif', 'color:#003575',
      'letter-spacing:.1em', 'text-transform:uppercase', 'cursor:pointer'
    ].join(';');
    b.addEventListener('click', abrirIngresos);
    document.body.appendChild(b);
  }

  var LBL = 'display:block;font-size:11px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:#c9d7ee;margin-bottom:7px;';
  var INP = 'width:100%;padding:13px 14px;border:1px solid rgba(255,255,255,.3);border-radius:2px;background:rgba(255,255,255,.95);color:#334155;font-size:17px;font-family:inherit;';
  var FORM = 'background:rgba(0,34,77,.6);padding:34px 30px 30px;border:1px solid rgba(255,255,255,.18);' +
    'border-top:3px solid #f0a500;-webkit-backdrop-filter:blur(6px);backdrop-filter:blur(6px);width:min(92vw,420px);color:#fff;';
  var BTN = 'width:100%;padding:14px 22px;border:none;border-radius:2px;background:#f0a500;' +
    'color:#1d1a12;font-size:16px;font-weight:600;font-family:inherit;cursor:pointer;';
  var ERR = 'min-height:20px;font-size:14px;color:#ffe2a8;margin-bottom:10px;';
  var EYEBROW = '<div style="font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:#9fc2f0;font-weight:600;">Calidad &amp; Capacitaci&oacute;n</div>';
  var H1 = 'margin:12px 0 8px;font-size:34px;line-height:1.1;color:#fff;font-weight:200;';
  var P = 'margin:0 0 24px;font-size:16px;color:#c9d7ee;font-weight:300;line-height:1.5;';

  function htmlLogin() {
    return '<form id="supLoginForm" autocomplete="on" style="' + FORM + '">' + EYEBROW +
        '<h1 style="' + H1 + '">Portal del <b style="font-weight:700;">Supervisor</b></h1>' +
        '<p style="' + P + '">Ingresa tus credenciales para ver los informes.</p>' +
        '<label for="supU" style="' + LBL + '">Usuario</label>' +
        '<input id="supU" name="username" autocomplete="username" autocapitalize="none" autocorrect="off" spellcheck="false" ' +
          'style="' + INP + 'margin-bottom:18px;">' +
        '<label for="supP" style="' + LBL + '">Clave</label>' +
        '<input id="supP" name="password" type="password" autocomplete="current-password" ' +
          'style="' + INP + 'margin-bottom:10px;">' +
        '<div id="supErr" role="alert" style="' + ERR + '"></div>' +
        '<button type="submit" id="supBtn" style="' + BTN + '">Ingresar &rarr;</button>' +
      '</form>';
  }

  function htmlCambio(nombre) {
    return '<form id="supCambioForm" autocomplete="on" style="' + FORM + '">' + EYEBROW +
        '<h1 style="' + H1 + '">Crea tu <b style="font-weight:700;">clave</b></h1>' +
        '<p id="supHola" style="' + P + '"></p>' +
        '<label for="supN1" style="' + LBL + '">Clave nueva</label>' +
        '<input id="supN1" name="new-password" type="password" autocomplete="new-password" ' +
          'style="' + INP + 'margin-bottom:18px;">' +
        '<label for="supN2" style="' + LBL + '">Repite la clave nueva</label>' +
        '<input id="supN2" name="new-password-2" type="password" autocomplete="new-password" ' +
          'style="' + INP + 'margin-bottom:10px;">' +
        '<div id="supErr" role="alert" style="' + ERR + '"></div>' +
        '<button type="submit" id="supBtn" style="' + BTN + '">Guardar y entrar &rarr;</button>' +
      '</form>';
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
    var caja = document.createElement('div');
    var logo = '<img src="logo-cobra.png" alt="Cobra" style="position:absolute;top:18px;left:22px;height:40px;width:auto;background:#fff;padding:6px 12px;">';
    ov.appendChild(caja);
    ov.insertAdjacentHTML('beforeend', logo);
    document.body.appendChild(ov);

    function entrar(perfil, token, uid) {
      // Registro de ingresos: si falla, no impide entrar.
      sb('POST', '/rest/v1/ingresos', { usuario_id: uid, dispositivo: dispositivo() }, token).catch(function () {});
      var d = { u: perfil.usuario, n: perfil.nombre || '', r: perfil.rol, ts: Date.now() };
      // Solo el admin conserva su sesión de Supabase, para poder leer el registro de ingresos.
      if (perfil.rol === 'admin') { d.at = token; d.rt = refresco; d.exp = vence; }
      try { sessionStorage.setItem(SESSION_KEY, JSON.stringify(d)); } catch (e) {}
      ov.remove();
      revelar();
      inyectarBarra(d);
    }
    var refresco = null, vence = 0;

    function pantallaCambio(perfil, token, uid, claveActual) {
      caja.innerHTML = htmlCambio();
      caja.style.display = 'contents';
      document.getElementById('supHola').textContent = 'Hola, ' + (perfil.nombre || perfil.usuario) +
        '. Es tu primer ingreso: crea una clave personal de al menos ' + CLAVE_MIN + ' caracteres.';
      var errEl = document.getElementById('supErr');
      var btn = document.getElementById('supBtn');
      document.getElementById('supCambioForm').addEventListener('submit', function (ev) {
        ev.preventDefault();
        errEl.textContent = '';
        var n1 = document.getElementById('supN1').value || '';
        var n2 = document.getElementById('supN2').value || '';
        if (n1.length < CLAVE_MIN) { errEl.textContent = 'La clave debe tener al menos ' + CLAVE_MIN + ' caracteres.'; return; }
        if (n1 !== n2) { errEl.textContent = 'Las dos claves no coinciden.'; return; }
        if (n1 === claveActual) { errEl.textContent = 'La clave nueva debe ser distinta de la inicial.'; return; }
        btn.disabled = true;
        sb('PUT', '/auth/v1/user', { password: n1 }, token).then(function (r) {
          if (!r.ok) throw new Error('clave');
          return sb('POST', '/rest/v1/rpc/marcar_clave_cambiada', {}, token);
        }).then(function (r) {
          if (!r.ok) throw new Error('marca');
          entrar(perfil, token, uid);
        }).catch(function () {
          btn.disabled = false;
          errEl.textContent = 'No se pudo guardar la clave. Reintenta.';
        });
      });
      document.getElementById('supN1').focus();
    }

    function pantallaLogin() {
      caja.innerHTML = htmlLogin();
      var errEl = document.getElementById('supErr');
      var btn = document.getElementById('supBtn');
      document.getElementById('supLoginForm').addEventListener('submit', function (ev) {
        ev.preventDefault();
        errEl.textContent = '';
        var u = (document.getElementById('supU').value || '').trim().toLowerCase();
        var p = document.getElementById('supP').value || '';
        if (!u || !p) { errEl.textContent = 'Completa usuario y clave.'; return; }
        if (!/^[a-z0-9._-]{3,30}$/.test(u)) { errEl.textContent = 'Usuario o clave incorrectos.'; return; }
        btn.disabled = true;
        var token, uid;
        sb('POST', '/auth/v1/token?grant_type=password', { email: u + DOMINIO, password: p }).then(function (r) {
          if (r.estado === 400 || r.estado === 401) throw new Error('credenciales');
          if (!r.ok || !r.datos || !r.datos.access_token) throw new Error('red');
          token = r.datos.access_token; uid = r.datos.user.id;
          refresco = r.datos.refresh_token; vence = Date.now() + (r.datos.expires_in || 3600) * 1000;
          return sb('GET', '/rest/v1/usuarios?select=usuario,nombre,rol,debe_cambiar_clave&id=eq.' + uid, null, token);
        }).then(function (r) {
          var perfil = r.ok && r.datos && r.datos[0];
          if (!perfil) throw new Error('credenciales');
          if (perfil.debe_cambiar_clave) { pantallaCambio(perfil, token, uid, p); return; }
          entrar(perfil, token, uid);
        }).catch(function (e) {
          btn.disabled = false;
          errEl.textContent = e && e.message === 'credenciales' ? 'Usuario o clave incorrectos.' : 'No se pudo validar. Revisa tu conexión y reintenta.';
        });
      });
      document.getElementById('supU').focus();
    }

    pantallaLogin();
  }

  function alDOM(fn) {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', fn);
    else fn();
  }

  var actual = sesion();
  if (actual) {
    revelar();
    alDOM(function () { inyectarBarra(actual); });
  } else {
    ocultar();
    alDOM(construirLogin);
  }
})();
