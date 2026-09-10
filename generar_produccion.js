// Genera Dashboard_Produccion.html a partir del INF-09 (produccion / baremos
// de Punta Arenas + Coyhaique) que copiamos a baremosTigo\. Es un informe
// SOLO para el supervisor (incluye baremos = datos de pago), enlazado desde
// el indice. Se regenera con Generar_Reporte_Reincidencias.bat, despues de
// copiar_baremos.bat.

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const carpeta = __dirname;
const carpetaBaremos = path.join(carpeta, 'baremosTigo');
const META_PRODUCTIVIDAD = 5; // productos por dia trabajado por tecnico

function rutInterno(r) {
  return (r || '').toString().trim().toUpperCase().replace(/\./g, '').replace(/-/g, '');
}
function titleCase(s) {
  return (s || '').split(' ').map((w) => (w ? w[0] + w.slice(1).toLowerCase() : w)).join(' ');
}
function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
const NOMBRES_MES = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

async function main() {
  if (!fs.existsSync(carpetaBaremos)) {
    console.log('AVISO: no existe la carpeta baremosTigo -- no se genera el Dashboard de Produccion.');
    return;
  }
  const cand = fs.readdirSync(carpetaBaremos)
    .filter((f) => /^INF-09.*\.xlsx$/i.test(f) && !f.startsWith('~$'))
    .map((f) => ({ f, m: fs.statSync(path.join(carpetaBaremos, f)).mtimeMs }))
    .sort((a, b) => b.m - a.m)[0];
  if (!cand) {
    console.log('AVISO: no se encontro ningun INF-09*.xlsx en baremosTigo -- no se genera el Dashboard de Produccion.');
    return;
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(carpetaBaremos, cand.f));
  const ws = wb.getWorksheet('BASE_MES_ PUNTA_ARENAS');
  if (!ws) {
    console.log('AVISO: no se encontro la hoja "BASE_MES_ PUNTA_ARENAS" -- no se genera el Dashboard de Produccion.');
    return;
  }

  // Columnas: 1 Folio | 2 Agencia | 4 Baremos | 5 Nro Productos | 6 RUT | 7 Nombre | 10 Dia | 14 Tipo | 18 Baremos a Pago
  const porRut = {};
  const porAgencia = {};
  const global = { ordenes: 0, productos: 0, instala: 0, repara: 0, baremos: 0, ruts: new Set(), td: new Set() };

  for (let i = 2; i <= ws.rowCount; i++) {
    const row = ws.getRow(i);
    if (!row.getCell(1).value) continue;
    const agencia = (row.getCell(2).value || '').toString().trim().toUpperCase();
    const prod = Number(row.getCell(5).value) || 0;
    const rut = rutInterno(row.getCell(6).value);
    const nombre = (row.getCell(7).value || '').toString().trim();
    const dia = Number(row.getCell(10).value);
    const tipo = (row.getCell(14).value || '').toString().trim();
    const baremo = Number(row.getCell(18).value) || 0; // columna R = punto baremo (a pago)
    if (!rut) continue;

    if (!porRut[rut]) porRut[rut] = { nombre, agencia, ordenes: 0, productos: 0, instala: 0, repara: 0, baremos: 0, dias: new Set() };
    const t = porRut[rut];
    t.ordenes += 1; t.productos += prod; t.baremos += baremo;
    if (tipo === 'Instala') t.instala += prod; else if (tipo === 'Repara') t.repara += prod;
    if (Number.isFinite(dia)) t.dias.add(dia);

    if (!porAgencia[agencia]) porAgencia[agencia] = { ordenes: 0, productos: 0, instala: 0, repara: 0, baremos: 0, ruts: new Set(), td: new Set() };
    const a = porAgencia[agencia];
    a.ordenes += 1; a.productos += prod; a.instala += (tipo === 'Instala' ? prod : 0); a.repara += (tipo === 'Repara' ? prod : 0);
    a.baremos += baremo; a.ruts.add(rut);
    if (Number.isFinite(dia)) a.td.add(rut + '|' + dia);

    global.ordenes += 1; global.productos += prod; global.instala += (tipo === 'Instala' ? prod : 0); global.repara += (tipo === 'Repara' ? prod : 0);
    global.baremos += baremo; global.ruts.add(rut);
    if (Number.isFinite(dia)) global.td.add(rut + '|' + dia);
  }

  const un = (v) => (v == null ? '-' : Number(v).toFixed(1));

  const tecnicos = Object.entries(porRut).map(([rut, t]) => {
    const dias = t.dias.size;
    return {
      nombre: titleCase(t.nombre), agencia: titleCase(t.agencia),
      ordenes: t.ordenes,
      instala: +t.instala.toFixed(1), repara: +t.repara.toFixed(1), productos: +t.productos.toFixed(1),
      dias, prodDia: dias ? +(t.productos / dias).toFixed(1) : null,
      baremos: +t.baremos.toFixed(1),
    };
  }).sort((a, b) => b.baremos - a.baremos);

  const agencias = Object.entries(porAgencia).map(([ag, a]) => ({
    agencia: titleCase(ag),
    tecnicos: a.ruts.size, ordenes: a.ordenes,
    instala: +a.instala.toFixed(1), repara: +a.repara.toFixed(1), productos: +a.productos.toFixed(1),
    baremos: +a.baremos.toFixed(1),
    prodDia: a.td.size ? +(a.productos / a.td.size).toFixed(1) : null,
    barDia: a.td.size ? +(a.baremos / a.td.size).toFixed(1) : null,
  })).sort((a, b) => b.baremos - a.baremos);

  const m = cand.f.match(/(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+(\d{4})/i);
  const periodo = m ? (m[1] + ' ' + m[2]) : cand.f;
  const gTecnicos = global.ruts.size;
  const gTd = global.td.size;

  const DATA = {
    periodo,
    generadoEl: new Date().toLocaleString('es-CL'),
    meta: META_PRODUCTIVIDAD,
    kpis: {
      productos: +global.productos.toFixed(1),
      baremos: +global.baremos.toFixed(1),
      tecnicos: gTecnicos,
      ordenes: global.ordenes,
      prodDia: gTd ? +(global.productos / gTd).toFixed(1) : null,
      barDia: gTd ? +(global.baremos / gTd).toFixed(1) : null,
    },
    agencias,
    tecnicos,
  };

  const json = JSON.stringify(DATA);
  if (json.indexOf('</scr' + 'ipt') !== -1) { console.log('ERROR: datos con secuencia insegura.'); return; }

  const html = plantilla(json);
  fs.writeFileSync(path.join(carpeta, 'Dashboard_Produccion.html'), html, 'utf8');
  console.log('Dashboard de Produccion generado:', path.join(carpeta, 'Dashboard_Produccion.html'), '(' + gTecnicos + ' tecnicos, periodo ' + periodo + ')');
}

function plantilla(json) {
  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Informe de Produccion · COBRA</title>
<meta name="theme-color" content="#003c71">
<link rel="manifest" href="manifest.json">
<link rel="apple-touch-icon" href="icon-192.png">
<script>try{if(!sessionStorage.getItem('supAuth_v1'))document.documentElement.style.visibility='hidden'}catch(e){}<\/script>
<script src="auth.js"><\/script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"><\/script>
<style>
  :root{ --bg:#eef1f4; --panel:#fff; --panel-2:#f5f7f9; --border:#e0e5ea; --text:#22303f; --text-dim:#6b7a8c;
    --cobra-navy:#003c71; --cobra-blue:#0071ce; --celeste:#29a9e0; --celeste-soft:#e8f6fd; --promotor:#1fa971; --neutro:#e2962e; --detractor:#e2523e; }
  *{box-sizing:border-box;}
  body{ margin:0; font-family:'Segoe UI', Arial, sans-serif; background:var(--bg); color:var(--text); -webkit-font-smoothing:antialiased; }
  header.hero{ background:linear-gradient(120deg,#fff 0%,var(--celeste-soft) 55%,#dcf1fb 100%); padding:34px 6vw 40px; position:relative; overflow:hidden; border-bottom:4px solid var(--celeste); }
  header.hero::after{ content:""; position:absolute; right:-100px; top:-100px; width:340px; height:340px; border-radius:50%; background:radial-gradient(circle, rgba(41,169,224,0.18), transparent 70%); }
  .back-link{ display:inline-flex; align-items:center; gap:6px; font-size:12.5px; font-weight:700; color:var(--cobra-navy); text-decoration:none; margin-bottom:14px; }
  .back-link:hover{ text-decoration:underline; }
  .brand-row{ display:flex; align-items:center; gap:18px; margin-bottom:22px; }
  .brand-row img{ height:46px; }
  .brand-divider{ width:1px; height:34px; background:var(--border); }
  .eyebrow{ text-transform:uppercase; letter-spacing:.14em; font-size:12.5px; color:var(--celeste); font-weight:800; }
  h1{ margin:0 0 6px; font-size:clamp(26px,4vw,38px); font-weight:800; letter-spacing:-0.01em; color:var(--cobra-navy); }
  .subtitle{ color:#3a4a5c; font-size:15px; max-width:680px; line-height:1.55; }
  .meta-row{ margin-top:20px; font-size:13px; color:#3a4a5c; }
  .meta-row b{ color:var(--cobra-navy); }
  main{ padding:36px 6vw 80px; max-width:1280px; margin:0 auto; }
  .kpi-grid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:16px; margin:-58px 0 34px; position:relative; z-index:2; }
  .kpi-card{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:20px 20px 18px; box-shadow:0 10px 24px rgba(20,50,80,.08); }
  .kpi-card .label{ font-size:11.5px; text-transform:uppercase; letter-spacing:.08em; color:var(--text-dim); font-weight:700; margin-bottom:8px; }
  .kpi-card .value{ font-size:28px; font-weight:800; line-height:1; color:var(--cobra-navy); }
  .kpi-card .value.hl{ color:var(--celeste); }
  .kpi-card .sub{ font-size:12px; color:var(--text-dim); margin-top:6px; }
  section{ margin-bottom:44px; }
  .section-title{ display:flex; align-items:baseline; gap:10px; margin-bottom:6px; }
  .section-title .num{ font-size:13px; font-weight:800; color:var(--celeste); background:rgba(41,169,224,0.12); border:1px solid rgba(41,169,224,.35); border-radius:6px; padding:2px 8px; }
  .section-title h2{ margin:0; font-size:20px; font-weight:750; color:var(--cobra-navy); }
  .section-desc{ color:var(--text-dim); font-size:13.5px; margin:0 0 18px; max-width:820px; line-height:1.55; }
  .panel{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:22px 24px 14px; box-shadow:0 4px 14px rgba(20,50,80,.05); }
  .grid-2{ display:grid; grid-template-columns:1.15fr 1fr; gap:18px; }
  @media (max-width:900px){ .grid-2{grid-template-columns:1fr;} }
  .tabla-wrap{ overflow-x:auto; }
  table{ width:100%; border-collapse:collapse; font-size:13px; }
  th{ text-align:right; color:var(--text-dim); font-weight:700; font-size:11px; text-transform:uppercase; letter-spacing:.04em; padding:9px 9px; border-bottom:1px solid var(--border); white-space:nowrap; }
  th:nth-child(2){ text-align:left; }
  td{ padding:8px 9px; border-bottom:1px solid var(--panel-2); text-align:right; white-space:nowrap; }
  td:nth-child(2){ text-align:left; color:var(--cobra-navy); font-weight:600; }
  td.ag{ text-align:left; color:var(--text-dim); }
  tr:hover td{ background:var(--panel-2); }
  td.b{ font-weight:800; color:var(--cobra-navy); }
  .val-ok{ color:var(--promotor); font-weight:800; }
  .val-bad{ color:var(--detractor); font-weight:800; }
  footer{ text-align:center; padding:26px; color:var(--text-dim); font-size:12px; border-top:1px solid var(--border); }
  canvas{ max-width:100%; }
</style>
</head>
<body>

<header class="hero">
  <a class="back-link" href="index.html">&larr; Volver al indice de informes</a>
  <div class="brand-row">
    <img src="logo-cobra.png" alt="Cobra">
    <div class="brand-divider"></div>
    <div class="eyebrow">Calidad &amp; Capacitacion &middot; Produccion</div>
  </div>
  <h1>Informe de Produccion</h1>
  <div class="subtitle">Produccion por tecnico de las agencias Punta Arenas y Coyhaique, a partir del INF-09: cantidad de productos (instala / repara), dias trabajados, productos por dia y puntos baremo.</div>
  <div class="meta-row" id="metaRow"></div>
</header>

<main>
  <div class="kpi-grid" id="kpiGrid"></div>

  <section>
    <div class="section-title"><span class="num">01</span><h2>Por agencia</h2></div>
    <p class="section-desc">Totales de la agencia y promedios por dia trabajado.</p>
    <div class="tabla-wrap panel"><table id="tablaAgencia"></table></div>
  </section>

  <section>
    <div class="section-title"><span class="num">02</span><h2>Top tecnicos por puntos baremo</h2></div>
    <div class="panel" style="height:340px;"><canvas id="chartTop"></canvas></div>
  </section>

  <section>
    <div class="section-title"><span class="num">03</span><h2>Ranking de tecnicos</h2></div>
    <p class="section-desc">Ordenado de mayor a menor puntos baremo. Prod/dia coloreado contra la meta de <span id="metaTxt"></span> productos/dia.</p>
    <div class="tabla-wrap panel"><table id="tablaTecnicos"></table></div>
  </section>
</main>

<footer id="footerText"></footer>

<script>
const DATA = ${json};

document.getElementById('metaRow').innerHTML = 'Periodo: <b>' + DATA.periodo + '</b> &nbsp;&middot;&nbsp; <b>' + DATA.kpis.tecnicos + '</b> tecnicos &nbsp;&middot;&nbsp; Generado ' + DATA.generadoEl;
document.getElementById('metaTxt').textContent = DATA.meta;
document.getElementById('footerText').textContent = 'Informe de Produccion COBRA · ' + DATA.periodo + ' · generado ' + DATA.generadoEl;

var k = DATA.kpis;
document.getElementById('kpiGrid').innerHTML = [
  '<div class="kpi-card"><div class="label">Productos totales</div><div class="value hl">' + k.productos + '</div><div class="sub">' + k.ordenes + ' ordenes</div></div>',
  '<div class="kpi-card"><div class="label">Puntos baremo</div><div class="value">' + k.baremos + '</div><div class="sub">total del periodo</div></div>',
  '<div class="kpi-card"><div class="label">Productos / dia</div><div class="value">' + (k.prodDia == null ? '-' : k.prodDia) + '</div><div class="sub">promedio por tecnico &middot; meta ' + DATA.meta + '</div></div>',
  '<div class="kpi-card"><div class="label">Baremo / dia</div><div class="value">' + (k.barDia == null ? '-' : k.barDia) + '</div><div class="sub">promedio por tecnico</div></div>',
].join('');

// ---- Tabla por agencia ----
(function () {
  var head = '<tr><th>Agencia</th><th>Tecnicos</th><th>Ordenes</th><th>Instala</th><th>Repara</th><th>Total prod.</th><th>Prod/dia</th><th>Puntos baremo</th><th>Baremo/dia</th></tr>';
  var rows = DATA.agencias.map(function (a) {
    var pc = a.prodDia == null ? '' : (a.prodDia >= DATA.meta ? 'val-ok' : 'val-bad');
    return '<tr><td class="ag">' + a.agencia + '</td><td>' + a.tecnicos + '</td><td>' + a.ordenes + '</td><td>' + a.instala + '</td><td>' + a.repara + '</td>'
      + '<td class="b">' + a.productos + '</td><td class="' + pc + '">' + (a.prodDia == null ? '-' : a.prodDia) + '</td>'
      + '<td class="b">' + a.baremos + '</td><td>' + (a.barDia == null ? '-' : a.barDia) + '</td></tr>';
  }).join('');
  document.getElementById('tablaAgencia').innerHTML = head + rows;
})();

// ---- Ranking de tecnicos ----
(function () {
  var head = '<tr><th>#</th><th>Tecnico</th><th>Agencia</th><th>Ordenes</th><th>Instala</th><th>Repara</th><th>Total prod.</th><th>Dias</th><th>Prod/dia</th><th>Puntos baremo</th></tr>';
  var rows = DATA.tecnicos.map(function (t, i) {
    var pc = t.prodDia == null ? '' : (t.prodDia >= DATA.meta ? 'val-ok' : 'val-bad');
    return '<tr><td>' + (i + 1) + '</td><td>' + t.nombre + '</td><td class="ag">' + t.agencia + '</td><td>' + t.ordenes + '</td>'
      + '<td>' + t.instala + '</td><td>' + t.repara + '</td><td class="b">' + t.productos + '</td><td>' + t.dias + '</td>'
      + '<td class="' + pc + '">' + (t.prodDia == null ? '-' : t.prodDia) + '</td><td class="b">' + t.baremos + '</td></tr>';
  }).join('');
  document.getElementById('tablaTecnicos').innerHTML = head + rows;
})();

// ---- Chart: top 10 por puntos baremo ----
(function () {
  var top = DATA.tecnicos.slice(0, 10);
  new Chart(document.getElementById('chartTop'), {
    type: 'bar',
    data: {
      labels: top.map(function (t) { return t.nombre; }),
      datasets: [{ label: 'Puntos baremo', data: top.map(function (t) { return t.baremos; }), backgroundColor: 'rgba(0,113,206,0.85)', borderRadius: 6, maxBarThickness: 34 }],
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { x: { title: { display: true, text: 'Puntos baremo' }, grid: { color: 'rgba(20,50,80,0.06)' } } },
    },
  });
})();
<\/script>

</body>
</html>`;
}

main().catch((err) => {
  console.error('ERROR generando el Dashboard de Produccion:', err.message);
  process.exitCode = 1;
});
