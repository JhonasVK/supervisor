// Formato visual: sistema de diseno de la Academia Tecnica (azul marino + ambar).
// Genera index.html: pagina de indice con acceso a los dos informes
// (Repetido Reparado / Averias de Infancia). Se regenera cada vez que
// corre Generar_Reporte_Reincidencias.bat, despues de los otros dos scripts.

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { execFileSync } = require('child_process');

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

// Formato corto y comun a todas las tarjetas: 29-09-2026 13:06
function fmtFecha(d) {
  const p = (n) => String(n).padStart(2, '0');
  return `${p(d.getDate())}-${p(d.getMonth() + 1)}-${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

function fechaArchivo(nombre) {
  const p = path.join(carpeta, nombre);
  if (!fs.existsSync(p)) return null;
  return fmtFecha(fs.statSync(p).mtime);
}

// Sitios externos: fecha del ultimo commit que publico su index.html en el repo local.
// El indice se genera despues de que esos sitios se publican en la manana
// (Auditorias TIGO 08:50, Informe NPS 10:00; este indice corre tras la MFT de las 11:00).
function fechaPublicacion(repoDir) {
  try {
    const iso = execFileSync('git', ['-C', repoDir, 'log', '-1', '--format=%cI', '--', 'index.html'], { encoding: 'utf8' }).trim();
    return iso ? fmtFecha(new Date(iso)) : null;
  } catch (e) {
    return null;
  }
}

const actualizadoReincidencias = fechaArchivo('Dashboard_Reincidencias.html');
const actualizadoInfancia = fechaArchivo('Dashboard_Infancia.html');
const actualizadoProduccion = fechaArchivo('Dashboard_Produccion.html');
const actualizadoAuditorias = fechaPublicacion('C:\\Bases_Tigo\\dashboard-auditorias-tigo');
const actualizadoNps = fechaPublicacion('C:\\Bases_Tigo\\NPS\\informe-nps');

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
    + '<div class="kpi-cols"><div class="kpi-scroll">' + tablaProd + '</div><div class="kpi-scroll">' + tablaCalidad + '</div></div>'
    + '<div class="kpi-note">"Total P.A." = Punta Arenas + Coyhaique como una sola agencia. Produccion: promedio del mes por tecnico &middot; promedio por dia trabajado (meta ' + META_PRODUCTIVIDAD + '/dia). Metas: Repetido Reparado &le;4% &middot; Infancia &le;2.5% &middot; NPS &ge;' + npsMeta + '%</div>'
    + '</section>';
}

function tarjeta({ href, disponible, titulo, descripcion, meta, actualizado, meses }) {
  if (!disponible) {
    return `<div class="card disabled">
      <div class="card-icon">Informe</div>
      <h2>${titulo}</h2>
      <p>${descripcion}</p>
      <div class="card-meta">Aun no generado</div>
    </div>`;
  }
  return `<a class="card" href="${href}">
    <div class="card-icon">Informe</div>
    <h2>${titulo}</h2>
    <p>${descripcion}</p>
    <div class="card-meta">Meta: ${meta} &nbsp;•&nbsp; ${meses} mes${meses === 1 ? '' : 'es'} de historial &nbsp;•&nbsp; Actualizado: ${actualizado}</div>
    <div class="card-cta">Ver informe &rarr;</div>
  </a>`;
}

// La fecha de un sitio externo se muestra primero con el valor del momento en que se
// genero este indice y, al abrir la pagina, se reemplaza por la que el sitio publica
// en su actualizado.txt (asi no queda atrasada si el sitio se publico despues).
function tarjetaExterna({ href, icono, titulo, descripcion, actualizado }) {
  return `<a class="card" href="${href}" target="_blank" rel="noopener">
    <div class="card-icon">Sitio externo</div>
    <h2>${titulo}</h2>
    <p>${descripcion}</p>
    <div class="card-meta">Sitio externo &nbsp;•&nbsp; Actualizado: <span class="fecha-viva" data-url="${href}actualizado.txt">${actualizado || '—'}</span></div>
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
<meta name="theme-color" content="#003575">
<link rel="manifest" href="manifest.json">
<link rel="apple-touch-icon" href="icon-192.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="Supervisor">
<script>try{if(!sessionStorage.getItem('supAuth_v1'))document.documentElement.style.visibility='hidden'}catch(e){}</script>
<script src="auth.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@200;300;400;500;600;700&family=IBM+Plex+Mono:wght@500&display=swap">
<style>
/* Formato de la Academia Tecnica (ver diseno-academia.json): azul marino + ambar,
     esquinas rectas, titulos livianos con la palabra clave en negrita. */
  :root{
    color-scheme: light;
    --navy:#003575; --blue:#3c62ac; --amber:#f0a500;
    --bg:#f2f2f2; --white:#ffffff; --ink:#334155; --muted:#62646e; --line:#d9dbe0;
    --ok:#1f7a45; --ok-bg:#e2f1e8; --warn:#9a5b00; --warn-bg:#fdf1d8;
    --f1:#1f5fbf; --f2:#e0701a; --f3:#2c8a45; --f4:#7a4b2a;
    --text-dim: var(--muted);
    --font:"Helvetica Neue","Hanken Grotesk",Helvetica,Arial,sans-serif;
    --mono:"IBM Plex Mono",ui-monospace,Consolas,monospace;
  }
  *{ box-sizing:border-box; }
  body{ margin:0; background:var(--bg); color:var(--ink); font:400 16px/1.6 var(--font); -webkit-font-smoothing:antialiased; }
  img{ max-width:100%; }
  :focus-visible{ outline:2px solid var(--amber); outline-offset:2px; }
  .wrap{ max-width:1180px; margin:0 auto; padding-inline:24px; }
  @media (max-width:480px){ .wrap{ padding-inline:16px; } }
  h1,h2{ text-wrap:balance; margin:0; }

  /* cabecera */
  header.top{ background:var(--white); border-bottom:1px solid var(--line); position:sticky; top:0; z-index:10; }
  .brand-row{ display:flex; align-items:center; gap:18px; min-height:72px; padding-block:8px; }
  .brand-row img{ height:40px; width:auto; display:block; cursor:pointer; }
  .brand-divider{ width:1px; height:34px; background:var(--line); }
  .brand-name{ font:300 18px/1.1 var(--font); letter-spacing:.12em; text-transform:uppercase; color:var(--blue); }
  .brand-name strong{ font-weight:700; color:var(--navy); }
  @media (max-width:560px){ .brand-row img{ height:30px; } .brand-name{ font-size:14px; } }

  /* portada */
  .hero{ position:relative; overflow:hidden; color:#ffffff; background:#003575; }
  .hero .photo{ position:absolute; top:0; right:0; bottom:0; width:74%; background:url("fondo-portada.jpg") 55% 22% / cover no-repeat;
    -webkit-mask-image:linear-gradient(90deg, transparent 0%, rgba(0,0,0,.35) 25%, rgba(0,0,0,.8) 48%, #000 65%); mask-image:linear-gradient(90deg, transparent 0%, rgba(0,0,0,.35) 25%, rgba(0,0,0,.8) 48%, #000 65%); }
  .hero .shade{ position:absolute; inset:0; pointer-events:none;
    background:linear-gradient(90deg, #003575 0%, #003575 26%, rgba(0,53,117,.9) 36%, rgba(0,53,117,.68) 46%, rgba(0,53,117,.4) 56%, rgba(0,53,117,.16) 67%, rgba(0,53,117,0) 82%),
               linear-gradient(0deg, rgba(0,34,77,.5) 0%, rgba(0,34,77,0) 38%); }
  .hero .fibers{ position:absolute; inset:0; width:100%; height:100%; pointer-events:none;
    -webkit-mask-image:linear-gradient(90deg, #000 0%, #000 38%, transparent 66%); mask-image:linear-gradient(90deg, #000 0%, #000 38%, transparent 66%); }
  .hero .fibers path{ fill:none; vector-effect:non-scaling-stroke; }
  .hero .fibers .strand path{ stroke:rgba(160,195,240,.1); stroke-width:1; }
  .hero .fibers .pulse{ stroke:rgba(255,190,60,.32); stroke-width:1.3; stroke-linecap:round; stroke-dasharray:70 1500; stroke-dashoffset:1570; }
  @media (prefers-reduced-motion:no-preference){
    .hero .fibers .pulse{ animation:pulse 6s linear infinite; }
    .hero .fibers .p1{ animation-delay:-2s; animation-duration:7s; } .hero .fibers .p2{ animation-delay:-4.5s; animation-duration:6.5s; }
    .hero .fibers .p3{ animation-delay:-1s; animation-duration:8s; } .hero .fibers .p4{ animation-delay:-5.5s; animation-duration:5.5s; }
    .hero .fibers .p5{ animation-delay:-3s; animation-duration:7.5s; } .hero .fibers .p6{ animation-delay:-6s; animation-duration:6.8s; }
  }
  @media (prefers-reduced-motion:reduce){ .hero .fibers .pulse{ display:none; } }
  @keyframes pulse{ from{ stroke-dashoffset:1570; } to{ stroke-dashoffset:0; } }
  .hero .wrap{ position:relative; display:grid; grid-template-columns:minmax(0,600px); align-content:center; min-height:340px; padding-block:56px 60px; }
  .eyebrow{ font:600 12px/1.2 var(--font); letter-spacing:.14em; text-transform:uppercase; color:#9fc2f0; margin:0; }
  .hero h1{ font:200 clamp(38px,5.6vw,64px)/1.05 var(--font); letter-spacing:.01em; margin:16px 0; }
  .hero h1 b{ font-weight:700; }
  .subtitle{ color:#c9d7ee; font-size:18px; font-weight:300; max-width:46ch; margin:0; }
  @media (max-width:860px){
    .hero .photo{ width:100%; bottom:auto; height:240px; background-position:46% 30%; -webkit-mask-image:none; mask-image:none; }
    .hero .shade{ background:linear-gradient(180deg, rgba(0,34,77,0) 0%, rgba(0,34,77,0) 90px, rgba(0,53,117,.9) 210px, #003575 240px); }
    .hero .fibers{ -webkit-mask-image:linear-gradient(180deg, transparent 0, transparent 180px, #000 300px); mask-image:linear-gradient(180deg, transparent 0, transparent 180px, #000 300px); }
    .hero .wrap{ min-height:0; padding-block:180px 44px; }
  }

  /* contenido */
  main{ padding-block:64px 88px; display:flex; flex-direction:column; gap:72px; }
  .sec-title{ display:flex; align-items:end; justify-content:space-between; gap:12px 24px; flex-wrap:wrap; margin-bottom:26px; }
  .sec-title h2{ font:200 32px/1.1 var(--font); color:var(--blue); text-transform:uppercase; letter-spacing:.04em; }
  .sec-title h2 b{ font-weight:700; color:var(--navy); }

  /* resumen por agencia */
  .kpi-panel{ background:var(--white); border-top:3px solid var(--navy); padding:clamp(20px,3vw,32px); }
  .kpi-panel-head{ font:600 12px/1.2 var(--font); letter-spacing:.14em; text-transform:uppercase; color:var(--blue); margin-bottom:18px; }
  .kpi-cols{ display:grid; grid-template-columns:1fr 1fr; gap:28px 36px; align-items:start; }
  @media (max-width:860px){ .kpi-cols{ grid-template-columns:1fr; } }
  .kpi-scroll{ overflow-x:auto; }
  .kpi-table{ width:100%; border-collapse:collapse; font-size:15px; }
  .kpi-table th{ background:var(--navy); color:#ffffff; font:600 11px/1.3 var(--font); text-transform:uppercase; letter-spacing:.1em; text-align:right; padding:11px 14px; }
  .kpi-table th:first-child{ text-align:left; }
  .kpi-table td{ padding:11px 14px; text-align:right; border-bottom:1px solid var(--line); font:500 14px/1.3 var(--mono); font-variant-numeric:tabular-nums; color:var(--navy); white-space:nowrap; }
  .kpi-table td.ind{ text-align:left; font:600 15px/1.3 var(--font); color:var(--ink); white-space:normal; }
  .kpi-table td.ind.sub2{ padding-left:28px; font-weight:400; color:var(--muted); }
  .kpi-table td.ind .sub{ font-weight:400; color:var(--muted); font-size:12.5px; }
  .kpi-table td.ok, .kpi-table td .ok{ color:var(--ok) !important; }
  .kpi-table td.bad, .kpi-table td .bad{ color:var(--warn) !important; }
  .kpi-table td.ok::after{ content:" ✓"; font-size:12px; }
  .kpi-table td.bad::after{ content:" ▲"; font-size:10px; }
  .kpi-table td span{ font-size:12px !important; font-weight:500 !important; margin-left:6px; }
  .kpi-table td.col-tot{ background:#f7f8fa; border-left:1px solid var(--line); }
  .kpi-note{ margin-top:18px; font-size:13.5px; color:var(--muted); border-left:3px solid var(--amber); background:var(--warn-bg); padding:10px 14px; }

  /* tarjetas de informes */
  .cards{ display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:24px; }
  @media (max-width:760px){ .cards{ grid-template-columns:1fr; } }
  .card{ --fc:var(--navy); background:var(--white); border-top:3px solid var(--fc); padding:28px 30px 24px; text-decoration:none; color:inherit; display:flex; flex-direction:column; gap:12px; min-width:0; transition:box-shadow .2s ease; }
  .card:nth-child(1){ --fc:var(--f1); } .card:nth-child(2){ --fc:var(--f2); } .card:nth-child(3){ --fc:var(--f3); } .card:nth-child(4){ --fc:var(--f4); }
  a.card:hover{ box-shadow:0 14px 30px -18px rgba(0,53,117,.45); }
  .card.disabled{ opacity:.55; cursor:default; }
  .card-icon{ font:600 11.5px/1 var(--font); letter-spacing:.14em; text-transform:uppercase; color:var(--blue); }
  .card-icon::after{ content:""; display:inline-block; width:8px; height:8px; border-radius:50%; background:var(--fc); margin-left:10px; }
  .card h2{ font:600 24px/1.2 var(--font); color:var(--navy); }
  .card p{ margin:0; color:var(--muted); font-size:15px; }
  .card-meta{ margin-top:auto; padding-top:14px; border-top:1px solid var(--line); font:500 12px/1.5 var(--mono); color:var(--muted); }
  .card-cta{ color:var(--blue); font-weight:600; font-size:14px; letter-spacing:.06em; text-transform:uppercase; }

  /* pie */
  footer.foot{ background:var(--navy); color:#c9d7ee; }
  footer.foot .wrap{ display:flex; flex-wrap:wrap; justify-content:space-between; align-items:center; gap:16px; padding-block:28px; }
  footer.foot img{ height:28px; width:auto; background:#ffffff; padding:5px 10px; display:block; }
  footer.foot span{ font-weight:300; letter-spacing:.1em; text-transform:uppercase; font-size:12px; }
  footer.foot .by{ font-size:9px; letter-spacing:.08em; opacity:.85; }
  footer.foot .by b{ text-transform:none; font-weight:700; color:#ffffff; letter-spacing:.08em; }

  /* easter egg del logo */
  .easter-toast{ position:fixed; left:50%; bottom:24px; transform:translateX(-50%) translateY(20px); background:#00224d; color:#fff; border-left:3px solid var(--amber); padding:14px 20px; font-size:15px; font-weight:600; box-shadow:0 18px 40px -16px rgba(0,20,50,.6); opacity:0; transition:opacity .25s ease, transform .25s ease; z-index:9999; pointer-events:none; max-width:calc(100vw - 32px); }
  .easter-toast.show{ opacity:1; transform:translateX(-50%) translateY(0); }
  .easter-emoji{ position:fixed; font-size:22px; pointer-events:none; z-index:9999; animation:easterFloat 1.1s ease-out forwards; }
  @keyframes easterFloat{ 0%{ transform:translate(0,0) scale(.6); opacity:1; } 100%{ transform:translate(var(--dx),-90px) scale(1.3); opacity:0; } }
  .easter-game-overlay{ position:fixed; inset:0; background:rgba(0,20,50,.9); z-index:10000; color:#fff; text-align:center; overflow:hidden; }
  .easter-game-overlay .eg-cerrar{ position:absolute; top:18px; right:22px; background:none; border:none; color:#fff; font-size:26px; cursor:pointer; }
  .easter-game-hud{ position:absolute; top:20px; left:22px; font-size:15px; font-weight:600; }
  .easter-game-titulo{ position:absolute; top:60px; left:0; right:0; font-size:15px; }
  .easter-star{ position:absolute; width:44px; height:44px; display:flex; align-items:center; justify-content:center; font-size:26px; cursor:pointer; user-select:none; }
  .easter-game-final{ position:absolute; top:40%; left:50%; transform:translate(-50%,-50%); font-size:20px; font-weight:700; }
  .easter-game-final button{ margin-top:16px; padding:12px 22px; border-radius:2px; border:none; background:var(--amber); color:#1d1a12; font-weight:600; cursor:pointer; font-size:14px; }
</style>
</head>
<body>

<header class="top">
  <div class="wrap brand-row">
    <img src="logo-cobra.png" alt="Cobra">
    <div class="brand-divider"></div>
    <div class="brand-name">Portal <strong>Supervisor</strong></div>
  </div>
</header>

<section class="hero" aria-label="Portada">
  <div class="photo" role="img" aria-label="Tecnico con EPP trabajando en un poste"></div>
  <div class="shade" aria-hidden="true"></div>
  <svg class="fibers" viewBox="0 0 1440 600" preserveAspectRatio="none" aria-hidden="true"><g class="strand"><path d="M-20 40 C339 12 744 101 1100 178"/><path d="M-20 76 C364 65 649 188 1100 187"/><path d="M-20 112 C304 107 651 122 1100 196"/><path d="M-20 148 C351 174 660 155 1100 205"/><path d="M-20 184 C375 220 732 195 1100 214"/><path d="M-20 220 C417 184 777 185 1100 223"/><path d="M-20 256 C317 225 689 289 1100 232"/><path d="M-20 292 C322 299 742 218 1100 241"/><path d="M-20 328 C366 293 650 197 1100 250"/><path d="M-20 364 C382 358 690 274 1100 259"/><path d="M-20 400 C354 384 767 304 1100 268"/><path d="M-20 436 C329 442 724 345 1100 277"/><path d="M-20 472 C388 455 797 217 1100 286"/><path d="M-20 508 C350 529 664 293 1100 295"/><path d="M-20 544 C305 557 762 317 1100 304"/><path d="M-20 580 C405 565 751 330 1100 313"/></g><g><path class="pulse p0" d="M-20 112 C304 107 651 122 1100 196"/><path class="pulse p1" d="M-20 292 C322 299 742 218 1100 241"/><path class="pulse p2" d="M-20 472 C388 455 797 217 1100 286"/><path class="pulse p3" d="M-20 40 C339 12 744 101 1100 178"/><path class="pulse p4" d="M-20 184 C375 220 732 195 1100 214"/><path class="pulse p5" d="M-20 364 C382 358 690 274 1100 259"/><path class="pulse p6" d="M-20 544 C305 557 762 317 1100 304"/></g></svg>
  <div class="wrap"><div>
    <p class="eyebrow">Calidad &amp; Capacitaci&oacute;n</p>
    <h1>Portal del <b>Supervisor</b></h1>
    <p class="subtitle">Elige el informe que quieres revisar.</p>
  </div></div>
</section>

<main class="wrap">
  <section>
    <div class="sec-title"><h2>Resumen por <b>agencia</b></h2></div>
    ${panelResumen(resumen)}
  </section>
  <section>
    <div class="sec-title"><h2>Informes <b>disponibles</b></h2></div>
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
      <div class="card-icon">Informe</div>
      <h2>Produccion por tecnicos</h2>
      <p>Produccion por tecnico de Punta Arenas y Coyhaique (INF-09): productos instala/repara, dias trabajados y productos por dia.</p>
      <div class="card-meta">Meta: ${META_PRODUCTIVIDAD}/dia &nbsp;&bull;&nbsp; Actualizado: ${actualizadoProduccion}</div>
      <div class="card-cta">Ver informe &rarr;</div>
    </a>` : ''}
    ${tarjetaExterna({
      href: 'https://jhonasvk.github.io/dashboard-auditorias-tigo/',
      icono: '🧾',
      titulo: 'Auditorias de Terreno',
      descripcion: 'Dashboard de auditorias en terreno: nota promedio por tecnico, top hallazgos e incumplimientos por supervisor.',
      actualizado: actualizadoAuditorias,
    })}
    ${tarjetaExterna({
      href: 'https://jhonasvk.github.io/informe-nps/',
      icono: '⭐',
      titulo: 'Informe NPS',
      descripcion: 'Net Promoter Score de las intervenciones tecnicas: evolucion diaria/semanal, desglose por zona y ranking de tecnicos.',
      actualizado: actualizadoNps,
    })}
  </div>
  </section>
</main>

<footer class="foot"><div class="wrap">
  <img src="logo-cobra.png" alt="Cobra">
  <span>Portal Supervisor &middot; Calidad &amp; Capacitaci&oacute;n</span>
  <span class="by">Desarrollado por <b>J.V.Soft...</b> &middot; Construyendo ideas</span>
</div></footer>

<script>
// Fecha en vivo de los sitios externos (Auditorias TIGO, Informe NPS)
document.querySelectorAll('.fecha-viva').forEach(function (el) {
  fetch(el.dataset.url, { cache: 'no-store' })
    .then(function (r) { return r.ok ? r.text() : ''; })
    .then(function (t) { t = (t || '').trim(); if (/^\\d{2}-\\d{2}-\\d{4} \\d{2}:\\d{2}$/.test(t)) el.textContent = t; })
    .catch(function () {});
});
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
