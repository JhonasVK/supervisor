// Genera index.html: pagina de indice con acceso a los dos informes
// (Repetido Reparado / Averias de Infancia). Se regenera cada vez que
// corre Generar_Reporte_Reincidencias.bat, despues de los otros dos scripts.

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const carpeta = __dirname;
const carpetaBbdd = path.join(carpeta, 'bbdd');
const carpetaBaremos = path.join(carpeta, 'baremosTigo');
const META_PRODUCTIVIDAD = 5; // productos por dia trabajado por tecnico

const NOMBRES_MES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
function nombreMes(slug) {
  const [y, m] = slug.split('-');
  return NOMBRES_MES[parseInt(m, 10) - 1] + ' ' + y;
}

function listarMeses(prefijo) {
  const re = new RegExp('^' + prefijo + '_(\\d{4}-\\d{2})\\.html$');
  const slugs = fs.readdirSync(carpeta)
    .map((f) => f.match(re))
    .filter(Boolean)
    .map((m) => m[1])
    .sort()
    .reverse();
  return slugs.map((slug) => ({
    slug,
    label: nombreMes(slug),
    url: `${prefijo}_${slug}.html`,
  }));
}

// meses.json queda disponible para que cada dashboard (incluidos los archivados,
// que ya no se regeneran) puedan consultar en vivo la lista completa y actualizada
// de meses disponibles, en vez de depender de la lista que tenian embebida al nacer.
const mesesReincidencias = listarMeses('Dashboard_Reincidencias');
const mesesInfancia = listarMeses('Dashboard_Infancia');
fs.writeFileSync(
  path.join(carpeta, 'meses.json'),
  JSON.stringify({ reincidencias: mesesReincidencias, infancia: mesesInfancia }),
  'utf8'
);
console.log('meses.json generado:', mesesReincidencias.length, 'meses de reincidencias,', mesesInfancia.length, 'meses de infancia');

function fechaArchivo(nombre) {
  const p = path.join(carpeta, nombre);
  if (!fs.existsSync(p)) return null;
  return fs.statSync(p).mtime.toLocaleString('es-CL');
}

const actualizadoReincidencias = fechaArchivo('Dashboard_Reincidencias.html');
const actualizadoInfancia = fechaArchivo('Dashboard_Infancia.html');
const actualizadoProduccion = fechaArchivo('Dashboard_Produccion.html');

// ---------- Resumen por agencia (panel superior, solo lo ve el supervisor) ----------

// Extrae el objeto DATA embebido de un Dashboard ya generado. Tolerante a
// CRLF (Windows) o LF, que el checkout de git puede intercambiar.
function leerDataDashboard(archivo) {
  const p = path.join(carpeta, archivo);
  if (!fs.existsSync(p)) return null;
  try {
    const html = fs.readFileSync(p, 'utf8');
    const m = html.match(/const DATA = (\{[\s\S]*?\});\r?\n\r?\nfunction npsClass/);
    return m ? JSON.parse(m[1]) : null;
  } catch (e) {
    return null;
  }
}

// NPS por zona a partir de bbdd/nps-tecnicos.json (lo exporta el Informe NPS).
function npsPorZona() {
  const p = path.join(carpetaBbdd, 'nps-tecnicos.json');
  if (!fs.existsSync(p)) return null;
  try {
    const nps = JSON.parse(fs.readFileSync(p, 'utf8'));
    if (!Array.isArray(nps.tecnicos)) return null;
    const z = {};
    nps.tecnicos.forEach((t) => {
      const zona = (t.zona || '').toUpperCase();
      if (!z[zona]) z[zona] = { P: 0, D: 0, total: 0 };
      z[zona].P += t.P; z[zona].D += t.D; z[zona].total += t.total;
    });
    const out = { meta: nps.meta, periodo: nps.periodo, valores: {} };
    const tot = { P: 0, D: 0, total: 0 };
    Object.entries(z).forEach(([zona, x]) => {
      out.valores[zona] = x.total ? +(((x.P - x.D) / x.total) * 100).toFixed(1) : null;
      tot.P += x.P; tot.D += x.D; tot.total += x.total;
    });
    out.valores['TOTAL'] = tot.total ? +(((tot.P - tot.D) / tot.total) * 100).toFixed(1) : null;
    return out;
  } catch (e) {
    return null;
  }
}

// Produccion promedio del mes por tecnico, por agencia, desde el INF-09.
async function productividadPorAgencia() {
  if (!fs.existsSync(carpetaBaremos)) return null;
  const cand = fs.readdirSync(carpetaBaremos)
    .filter((f) => /^INF-09.*\.xlsx$/i.test(f) && !f.startsWith('~$'))
    .map((f) => ({ f, m: fs.statSync(path.join(carpetaBaremos, f)).mtimeMs }))
    .sort((a, b) => b.m - a.m)[0];
  if (!cand) return null;
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(carpetaBaremos, cand.f));
    const ws = wb.getWorksheet('BASE_MES_ PUNTA_ARENAS');
    if (!ws) return null;
    const ag = {};
    for (let k = 2; k <= ws.rowCount; k++) {
      const row = ws.getRow(k);
      if (!row.getCell(1).value) continue;
      const agencia = (row.getCell(2).value || '').toString().trim().toUpperCase();
      const prod = Number(row.getCell(5).value) || 0;
      const rut = (row.getCell(6).value || '').toString().trim();
      const dia = Number(row.getCell(10).value);
      const tipo = (row.getCell(14).value || '').toString().trim();
      for (const key of [agencia, 'TOTAL']) {
        if (!ag[key]) ag[key] = { total: 0, instala: 0, repara: 0, ruts: new Set(), td: new Set() };
        ag[key].total += prod;
        if (tipo === 'Instala') ag[key].instala += prod;
        else if (tipo === 'Repara') ag[key].repara += prod;
        if (rut) ag[key].ruts.add(rut);
        if (rut && Number.isFinite(dia)) ag[key].td.add(rut + '|' + dia);
      }
    }
    const m = cand.f.match(/(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+(\d{4})/i);
    const out = { periodo: m ? (m[1] + ' ' + m[2]) : null, valores: {} };
    Object.entries(ag).forEach(([agencia, x]) => {
      const nt = x.ruts.size, td = x.td.size;
      const par = (n) => ({ mes: nt ? +(n / nt).toFixed(1) : null, dia: td ? +(n / td).toFixed(2) : null });
      out.valores[agencia] = { total: par(x.total), instala: par(x.instala), repara: par(x.repara) };
    });
    return out;
  } catch (e) {
    return null;
  }
}

// Tasa por agencia + una entrada TOTAL que combina todas las agencias como si
// fueran una sola (se recalcula sobre los casos crudos, no promediando tasas).
function agenciasDe(data, campoCasos) {
  const o = {};
  let totC = 0, totT = 0;
  if (data && Array.isArray(data.agencias)) {
    data.agencias.forEach((a) => {
      o[(a.agencia || '').toUpperCase()] = a.tasa;
      totC += a[campoCasos] || 0;
      totT += a.total || 0;
    });
    o['TOTAL'] = totT ? +((totC / totT) * 100).toFixed(1) : null;
  }
  return o;
}

// Arma el HTML del panel de resumen (tabla: filas = indicadores, columnas =
// agencias). Dos agencias fijas: Punta Arenas y Coyhaique.
function panelResumen(resumen) {
  const AG = ['PUNTA ARENAS', 'COYHAIQUE'];
  const un = (v) => Number(v).toFixed(1); // siempre 1 decimal
  const fmt = (v, suf) => (v == null ? '—' : un(v) + (suf || ''));
  const cel = (txt, clase) => '<td class="' + (clase || '') + '">' + txt + '</td>';
  const tasaCls = (v, meta) => (v == null ? '' : (v <= meta ? 'ok' : 'bad'));
  const npsCls = (v, meta) => (v == null ? '' : (v >= meta ? 'ok' : 'bad'));

  const rr = resumen.reincidencias, inf = resumen.infancia, nps = resumen.nps, prod = resumen.productividad;
  const npsMeta = (nps && nps.meta) || 86;

  // Produccion: promedio del mes + promedio por dia. Solo el TOTAL se colorea
  // contra la meta de 5/dia (la meta es del total, no de cada tipo por separado).
  function celProd(k, campo, extraCls) {
    const v = prod && prod.valores[k] && prod.valores[k][campo];
    if (!v || v.mes == null) return cel('—', extraCls);
    let dia = '';
    if (v.dia != null) {
      const cls = campo === 'total' ? (v.dia >= META_PRODUCTIVIDAD ? 'ok' : 'bad') : '';
      dia = ' <span class="' + cls + '" style="font-size:12px;font-weight:700;color:' + (cls ? '' : 'var(--text-dim)') + ';">' + un(v.dia) + '/d</span>';
    }
    return cel(un(v.mes) + dia, extraCls);
  }

  const celTasa = (v, meta) => cel(fmt(v, '%'), tasaCls(v, meta));
  const celNps = (v) => cel(nps && nps.valores ? fmt(v, '%') : '—', nps && nps.valores ? npsCls(v, npsMeta) : '');

  const tablaProd = '<table class="kpi-table">'
    + '<tr class="fila-tit"><th>Produccion / tecnico</th><th>P. Arenas</th><th>Coyhaique</th><th class="col-tot">Total P.A.</th></tr>'
    + '<tr><td class="ind">Total <span class="sub">(mes &middot; /dia)</span></td>' + celProd('PUNTA ARENAS', 'total') + celProd('COYHAIQUE', 'total') + celProd('TOTAL', 'total', 'col-tot') + '</tr>'
    + '<tr><td class="ind sub2">Instala</td>' + celProd('PUNTA ARENAS', 'instala') + celProd('COYHAIQUE', 'instala') + celProd('TOTAL', 'instala', 'col-tot') + '</tr>'
    + '<tr><td class="ind sub2">Repara</td>' + celProd('PUNTA ARENAS', 'repara') + celProd('COYHAIQUE', 'repara') + celProd('TOTAL', 'repara', 'col-tot') + '</tr>'
    + '</table>';

  const tablaCalidad = '<table class="kpi-table">'
    + '<tr class="fila-tit"><th>Calidad</th><th>P. Arenas</th><th>Coyhaique</th><th class="col-tot">Total P.A.</th></tr>'
    + '<tr><td class="ind">Repetido Reparado</td>' + celTasa(rr['PUNTA ARENAS'], 4) + celTasa(rr['COYHAIQUE'], 4) + cel(fmt(rr['TOTAL'], '%'), tasaCls(rr['TOTAL'], 4) + ' col-tot') + '</tr>'
    + '<tr><td class="ind">Averias de Infancia</td>' + celTasa(inf['PUNTA ARENAS'], 2.5) + celTasa(inf['COYHAIQUE'], 2.5) + cel(fmt(inf['TOTAL'], '%'), tasaCls(inf['TOTAL'], 2.5) + ' col-tot') + '</tr>'
    + '<tr><td class="ind">NPS</td>' + celNps(nps && nps.valores && nps.valores['PUNTA ARENAS']) + celNps(nps && nps.valores && nps.valores['COYHAIQUE'])
      + cel(nps && nps.valores ? fmt(nps.valores['TOTAL'], '%') : '—', (nps && nps.valores ? npsCls(nps.valores['TOTAL'], npsMeta) : '') + ' col-tot') + '</tr>'
    + '</table>';

  const periodo = (prod && prod.periodo) || (nps && nps.periodo) || '';
  return '<section class="kpi-panel">'
    + '<div class="kpi-panel-head">Resumen por agencia' + (periodo ? ' &middot; ' + periodo : '') + '</div>'
    + '<div class="kpi-cols">' + tablaProd + tablaCalidad + '</div>'
    + '<div class="kpi-note">"Total P.A." = Punta Arenas + Coyhaique como una sola agencia. Produccion: promedio del mes por tecnico &middot; promedio por dia trabajado (meta ' + META_PRODUCTIVIDAD + '/dia). Metas: Repetido Reparado &le;4% &middot; Infancia &le;2.5% &middot; NPS &ge;' + npsMeta + '%</div>'
    + '</section>';
}

function tarjeta({ href, disponible, titulo, descripcion, meta, actualizado, meses }) {
  if (!disponible) {
    return `<div class="card disabled">
      <div class="card-icon">📄</div>
      <h2>${titulo}</h2>
      <p>${descripcion}</p>
      <div class="card-meta">Aun no generado</div>
    </div>`;
  }
  return `<a class="card" href="${href}">
    <div class="card-icon">📊</div>
    <h2>${titulo}</h2>
    <p>${descripcion}</p>
    <div class="card-meta">Meta: ${meta} &nbsp;•&nbsp; ${meses} mes${meses === 1 ? '' : 'es'} de historial &nbsp;•&nbsp; Actualizado: ${actualizado}</div>
    <div class="card-cta">Ver informe &rarr;</div>
  </a>`;
}

function tarjetaExterna({ href, icono, titulo, descripcion, nota }) {
  return `<a class="card" href="${href}" target="_blank" rel="noopener">
    <div class="card-icon">${icono}</div>
    <h2>${titulo}</h2>
    <p>${descripcion}</p>
    <div class="card-meta">${nota}</div>
    <div class="card-cta">Ver informe &rarr;</div>
  </a>`;
}

async function main() {

const resumen = {
  reincidencias: agenciasDe(leerDataDashboard('Dashboard_Reincidencias.html'), 'reincidencias'),
  infancia: agenciasDe(leerDataDashboard('Dashboard_Infancia.html'), 'infancia'),
  nps: npsPorZona(),
  productividad: await productividadPorAgencia(),
};
console.log('Resumen por agencia armado (productividad:', resumen.productividad ? 'si' : 'no', '| nps:', resumen.nps ? 'si' : 'no', ')');

const html = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Supervisor · COBRA</title>
<meta name="theme-color" content="#003c71">
<link rel="manifest" href="manifest.json">
<link rel="apple-touch-icon" href="icon-192.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="Supervisor">
<script>try{if(!sessionStorage.getItem('supAuth_v1'))document.documentElement.style.visibility='hidden'}catch(e){}</script>
<script src="auth.js"></script>
<style>
  :root{
    --bg:#eef1f4; --panel:#ffffff; --border:#e0e5ea; --text:#22303f; --text-dim:#6b7a8c;
    --cobra-navy:#003c71; --cobra-blue:#0071ce; --celeste:#29a9e0; --celeste-soft:#e8f6fd;
  }
  *{box-sizing:border-box;}
  body{ margin:0; font-family:'Segoe UI', Arial, sans-serif; background:var(--bg); color:var(--text); -webkit-font-smoothing:antialiased; }
  header.hero{
    background:linear-gradient(120deg,#ffffff 0%,var(--celeste-soft) 55%,#dcf1fb 100%);
    padding:34px 6vw 40px; position:relative; overflow:hidden; border-bottom:4px solid var(--celeste);
  }
  header.hero::after{
    content:""; position:absolute; right:-100px; top:-100px; width:340px; height:340px; border-radius:50%;
    background:radial-gradient(circle, rgba(41,169,224,0.18), transparent 70%);
  }
  .brand-row{ display:flex; align-items:center; gap:18px; margin-bottom:22px; }
  .brand-row img{ height:46px; cursor:pointer; }
  .easter-toast{ position:fixed; left:50%; bottom:30px; transform:translateX(-50%) translateY(20px); background:var(--cobra-navy); color:#fff; padding:12px 22px; border-radius:30px; font-size:14px; font-weight:700; box-shadow:0 8px 24px rgba(0,60,113,.3); opacity:0; transition:opacity .25s ease, transform .25s ease; z-index:9999; pointer-events:none; white-space:nowrap; }
  .easter-toast.show{ opacity:1; transform:translateX(-50%) translateY(0); }
  .easter-emoji{ position:fixed; font-size:22px; pointer-events:none; z-index:9999; animation:easterFloat 1.1s ease-out forwards; }
  @keyframes easterFloat{ 0%{ transform:translate(0,0) scale(.6); opacity:1; } 100%{ transform:translate(var(--dx),-90px) scale(1.3); opacity:0; } }
  .easter-game-overlay{ position:fixed; inset:0; background:rgba(10,20,35,.85); z-index:10000; color:#fff; text-align:center; overflow:hidden; }
  .easter-game-overlay .eg-cerrar{ position:absolute; top:18px; right:22px; background:none; border:none; color:#fff; font-size:26px; cursor:pointer; }
  .easter-game-hud{ position:absolute; top:20px; left:22px; font-size:15px; font-weight:700; }
  .easter-game-titulo{ position:absolute; top:60px; left:0; right:0; font-size:15px; }
  .easter-star{ position:absolute; width:44px; height:44px; display:flex; align-items:center; justify-content:center; font-size:26px; cursor:pointer; user-select:none; }
  .easter-game-final{ position:absolute; top:40%; left:50%; transform:translate(-50%,-50%); font-size:20px; font-weight:800; }
  .easter-game-final button{ margin-top:16px; padding:10px 22px; border-radius:20px; border:none; background:var(--celeste); color:var(--cobra-navy); font-weight:800; cursor:pointer; font-size:14px; }
  .brand-divider{ width:1px; height:34px; background:var(--border); }
  .eyebrow{ text-transform:uppercase; letter-spacing:.14em; font-size:12.5px; color:var(--celeste); font-weight:800; }
  h1{ margin:0 0 6px; font-size:clamp(26px,4vw,38px); font-weight:800; letter-spacing:-0.01em; color:var(--cobra-navy); }
  .subtitle{ color:#3a4a5c; font-size:15px; max-width:660px; line-height:1.55; }
  main{ padding:44px 6vw 80px; max-width:1100px; margin:0 auto; }
  .cards{ display:grid; grid-template-columns:repeat(auto-fit,minmax(300px,1fr)); gap:22px; }
  .card{
    display:block; background:var(--panel); border:1px solid var(--border); border-radius:16px;
    padding:26px 26px 22px; text-decoration:none; color:inherit; box-shadow:0 10px 24px rgba(20,50,80,.06);
    transition: transform .15s ease, box-shadow .15s ease;
  }
  .card:hover{ transform:translateY(-3px); box-shadow:0 14px 30px rgba(20,50,80,.12); }
  .card.disabled{ opacity:.55; cursor:default; }
  .card-icon{ font-size:28px; margin-bottom:10px; }
  .card h2{ margin:0 0 8px; font-size:19px; color:var(--cobra-navy); }
  .card p{ margin:0 0 14px; font-size:13.5px; color:var(--text-dim); line-height:1.55; }
  .card-meta{ font-size:11.5px; color:var(--text-dim); border-top:1px solid var(--border); padding-top:12px; }
  .card-cta{ margin-top:10px; font-size:13px; font-weight:700; color:var(--celeste); }

  .kpi-panel{ background:var(--panel); border:1px solid var(--border); border-radius:16px; padding:20px 24px 16px; box-shadow:0 10px 24px rgba(20,50,80,.06); margin-bottom:26px; }
  .kpi-panel-head{ font-size:12px; text-transform:uppercase; letter-spacing:.06em; font-weight:800; color:var(--text-dim); margin-bottom:12px; }
  .kpi-cols{ display:grid; grid-template-columns:1fr 1fr; gap:30px; align-items:start; }
  @media (max-width:760px){ .kpi-cols{ grid-template-columns:1fr; } }
  .kpi-table{ width:100%; border-collapse:collapse; font-size:13.5px; }
  .kpi-table th{ text-align:right; font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:var(--text-dim); font-weight:800; padding:4px 10px 8px; }
  .kpi-table th:first-child{ text-align:left; color:var(--cobra-navy); }
  .kpi-table td{ padding:5px 10px; text-align:right; font-weight:800; color:var(--cobra-navy); white-space:nowrap; }
  .kpi-table td.ind{ text-align:left; font-weight:600; color:var(--text); }
  .kpi-table td.ind.sub2{ padding-left:22px; color:var(--text-dim); font-weight:500; }
  .kpi-table td.ind .sub{ font-weight:400; color:var(--text-dim); font-size:11px; }
  .kpi-table td.ok{ color:#1fa971; }
  .kpi-table td.bad{ color:#e2523e; }
  .kpi-table td .ok{ color:#1fa971; }
  .kpi-table td .bad{ color:#e2523e; }
  .kpi-table .col-tot{ border-left:1px solid var(--border); padding-left:14px; background:#f6f9fc; }
  .kpi-table th.col-tot{ background:transparent; }
  .kpi-note{ margin-top:12px; font-size:11px; color:var(--text-dim); }

  footer{ text-align:center; padding:26px; color:var(--text-dim); font-size:12px; }
</style>
</head>
<body>

<header class="hero">
  <div class="brand-row">
    <img src="logo-cobra.png" alt="Cobra">
    <div class="brand-divider"></div>
    <div class="eyebrow">Calidad &amp; Capacitacion</div>
  </div>
  <h1>Supervisor</h1>
  <div class="subtitle">Elige el informe que quieres revisar.</div>
</header>

<main>
  ${panelResumen(resumen)}
  <div class="cards">
    ${tarjeta({
      href: 'Dashboard_Reincidencias.html',
      disponible: !!actualizadoReincidencias,
      titulo: 'Informe de Repetido Reparado',
      descripcion: 'Reparaciones que volvieron a fallar dentro de 30 dias: tasa por agencia/causa/tecnico, tiempo hasta la reiteracion y si la atendio el mismo tecnico.',
      meta: '4%',
      actualizado: actualizadoReincidencias,
      meses: mesesReincidencias.length,
    })}
    ${tarjeta({
      href: 'Dashboard_Infancia.html',
      disponible: !!actualizadoInfancia,
      titulo: 'Informe de Averias de Infancia',
      descripcion: 'Instalaciones que generaron una reparacion dentro de su periodo de infancia: tasa por agencia/producto/tecnico instalador, causas y tiempo hasta la falla.',
      meta: '2.5%',
      actualizado: actualizadoInfancia,
      meses: mesesInfancia.length,
    })}
    ${actualizadoProduccion ? `<a class="card" href="Dashboard_Produccion.html">
      <div class="card-icon">📊</div>
      <h2>Produccion por tecnicos</h2>
      <p>Produccion por tecnico de Punta Arenas y Coyhaique (INF-09): productos instala/repara, dias trabajados, productos por dia y puntos baremo.</p>
      <div class="card-meta">Meta: ${META_PRODUCTIVIDAD}/dia &nbsp;&bull;&nbsp; Actualizado: ${actualizadoProduccion}</div>
      <div class="card-cta">Ver informe &rarr;</div>
    </a>` : ''}
    ${tarjetaExterna({
      href: 'https://jhonasvk.github.io/dashboard-auditorias-tigo/',
      icono: '🧾',
      titulo: 'Auditorias de Terreno',
      descripcion: 'Dashboard de auditorias en terreno: nota promedio por tecnico, top hallazgos e incumplimientos por supervisor.',
      nota: 'Sitio externo',
    })}
    ${tarjetaExterna({
      href: 'https://jhonasvk.github.io/informe-nps/',
      icono: '⭐',
      titulo: 'Informe NPS',
      descripcion: 'Net Promoter Score de las intervenciones tecnicas: evolucion diaria/semanal, desglose por zona y ranking de tecnicos.',
      nota: 'Sitio externo',
    })}
  </div>
</main>

<script>
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(()=>{}));
}

// ---- Easter egg: click en el logo ----
(function () {
  var FRASES = [
    '¡Vas por buen camino! 💪',
    'Cada reparacion cuenta, sigue asi 🔧',
    'Hoy es un gran dia para superar tu meta ⭐',
    'El equipo esta orgulloso de tu trabajo 🙌',
    'Un dia a la vez, vas mejorando 📈',
    'Gracias por dejar todo en cada visita 🚀',
    'Tu esfuerzo se nota, sigue asi 🌟',
    'Pequenos pasos, grandes resultados 🏆',
    'Lo que haces hoy suma para todo el equipo 🤝',
    'Cada visita es una oportunidad de brillar ✨',
    'Tu trabajo mejora la conexion de muchas familias 📶',
    'La constancia gana, sigue enfocado 🎯',
    'Buen trabajo no siempre se ve, pero siempre se nota 👏',
    'Eres parte clave del equipo COBRA 🔵',
    'Sigue asi, los resultados van a llegar 🚀',
    'Un cliente satisfecho es la mejor recompensa 😊',
    'Nunca es tarde para un gran dia 🌅',
    'Confiamos en tu criterio y experiencia 🛠️',
    'Cada reto es una oportunidad de aprender 📚',
    'Tu actitud hace la diferencia 🔥',
  ];
  // Frases especiales por fecha (automatico segun el mes del visitante):
  // Fiestas Patrias en septiembre, Navidad/Ano Nuevo en diciembre.
  var FRASES_ESPECIALES = {
    9: [
      '¡Viva Chile! Que las Fiestas Patrias te recarguen de energia 🇨🇱🎉',
      'Dieciocho de septiembre: a celebrar como se merece, con toda la energia del pais 🇨🇱🥟',
      'Como buen chileno, sigue poniendole empanada y power a cada dia 🇨🇱💪',
      'Fiestas Patrias: buen momento para parar, celebrar, y volver con toda la energia 🇨🇱',
    ],
    12: [
      '🎄 Feliz Navidad, que este mes cierre con broche de oro',
      '¡Que el espiritu navideno te acompane en cada visita! 🎅',
      'Un fin de ano de excelentes resultados para ti y tu familia 🎆',
      '🎁 Diciembre es para cerrar el ano arriba, sigue asi',
    ],
  };
  var FRASES_MES = FRASES.concat(FRASES_ESPECIALES[new Date().getMonth() + 1] || []);
  var clicks = 0, clickTimer = null;

  function mostrarToast(texto) {
    var t = document.createElement('div');
    t.className = 'easter-toast';
    t.textContent = texto;
    document.body.appendChild(t);
    requestAnimationFrame(function () { t.classList.add('show'); });
    setTimeout(function () {
      t.classList.remove('show');
      setTimeout(function () { t.remove(); }, 300);
    }, 2200);
  }

  function lanzarEmojis(x, y) {
    var emojis = ['✨', '⭐', '🎉', '💙'];
    for (var i = 0; i < 6; i++) {
      var e = document.createElement('div');
      e.className = 'easter-emoji';
      e.textContent = emojis[Math.floor(Math.random() * emojis.length)];
      e.style.left = x + 'px';
      e.style.top = y + 'px';
      e.style.setProperty('--dx', (Math.random() * 140 - 70) + 'px');
      document.body.appendChild(e);
      (function (el) { setTimeout(function () { el.remove(); }, 1200); })(e);
    }
  }

  function iniciarMiniJuego() {
    var overlay = document.createElement('div');
    overlay.className = 'easter-game-overlay';
    overlay.innerHTML = '<button class="eg-cerrar">✕</button>'
      + '<div class="easter-game-hud">⭐ Puntaje: <span id="egScore">0</span> · ⏱ <span id="egTime">15</span>s</div>'
      + '<div class="easter-game-titulo">¡Atrapa las estrellas!</div>';
    document.body.appendChild(overlay);

    var score = 0, tiempo = 15;
    var scoreEl = overlay.querySelector('#egScore');
    var timeEl = overlay.querySelector('#egTime');

    var spawnInt = setInterval(function () {
      var star = document.createElement('div');
      star.className = 'easter-star';
      star.textContent = '⭐';
      star.style.left = (Math.random() * 80 + 5) + '%';
      star.style.top = '-40px';
      overlay.appendChild(star);
      var posY = -40;
      var fall = setInterval(function () {
        posY += 3;
        star.style.top = posY + 'px';
        if (posY > window.innerHeight) { star.remove(); clearInterval(fall); }
      }, 16);
      star.onclick = function () {
        score++;
        scoreEl.textContent = score;
        clearInterval(fall);
        star.remove();
      };
    }, 550);

    var timeInt = setInterval(function () {
      tiempo--;
      timeEl.textContent = tiempo;
      if (tiempo <= 0) {
        clearInterval(spawnInt);
        clearInterval(timeInt);
        var restantes = overlay.querySelectorAll('.easter-star');
        for (var i = 0; i < restantes.length; i++) restantes[i].remove();
        overlay.innerHTML = '<button class="eg-cerrar">✕</button>'
          + '<div class="easter-game-final">🎉 Puntaje final: ' + score + ' estrellas<br>'
          + '<span style="font-size:14px;font-weight:400;">' + FRASES_MES[Math.floor(Math.random() * FRASES_MES.length)] + '</span><br>'
          + '<button id="egCerrarFinal">Cerrar</button></div>';
        overlay.querySelector('#egCerrarFinal').addEventListener('click', function () { overlay.remove(); });
        overlay.querySelector('.eg-cerrar').addEventListener('click', function () { overlay.remove(); });
      }
    }, 1000);

    overlay.querySelector('.eg-cerrar').addEventListener('click', function () {
      clearInterval(spawnInt);
      clearInterval(timeInt);
      overlay.remove();
    });
  }

  var logo = document.querySelector('.brand-row img');
  if (logo) {
    logo.addEventListener('click', function (e) {
      clicks++;
      lanzarEmojis(e.clientX, e.clientY);
      mostrarToast(FRASES_MES[Math.floor(Math.random() * FRASES_MES.length)]);
      clearTimeout(clickTimer);
      clickTimer = setTimeout(function () { clicks = 0; }, 3000);
      if (clicks >= 5) {
        clicks = 0;
        iniciarMiniJuego();
      }
    });
  }
})();
</script>

</body>
</html>`;

fs.writeFileSync(path.join(carpeta, 'index.html'), html, 'utf8');
console.log('Indice generado:', path.join(carpeta, 'index.html'));

}

main().catch((err) => {
  console.error('ERROR generando el indice:', err.message);
  process.exitCode = 1;
});
