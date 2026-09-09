// Genera una imagen (PNG) con el resumen diario de Repetido Reparado y
// Averias de Infancia, pensada para copiar y pegar directo en el cuerpo de
// un correo. Lee los dos dashboards ya generados (deben correr ANTES en el
// mismo pipeline: generar_reincidencias.js y generar_infancia.js) y los
// renderiza con Chrome o Edge en modo headless (no necesita Puppeteer).
//
// La imagen final NO se sube a GitHub (ver .gitignore) -- es solo para uso
// local, para pegarla en un correo.

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const carpeta = __dirname;
const OUTPUT_PNG = path.join(carpeta, 'Resumen_Diario_Correo.png');
const TEMP_PNG = path.join(carpeta, '_resumen_correo_tmp.png');
const TEMP_HTML = path.join(carpeta, '_resumen_correo_tmp.html');

function getData(archivoHtml) {
  const html = fs.readFileSync(archivoHtml, 'utf8');
  // Tolerante a CRLF (Windows) o LF: el checkout de git puede convertir los
  // saltos de linea segun core.autocrlf, y no hay que depender de cual sea.
  const m = html.match(/const DATA = (\{[\s\S]*?\});\r?\n\r?\nfunction npsClass/);
  if (!m) throw new Error('No se pudo leer el DATA embebido de ' + archivoHtml);
  return JSON.parse(m[1]);
}

function titleCase(s) {
  return (s || '').split(' ').map((w) => (w ? w[0] + w.slice(1).toLowerCase() : w)).join(' ');
}

function filasVariacion(lista, tipo) {
  if (!lista || lista.length === 0) {
    return '<div class="fila"><span class="nombre" style="color:#8a97a6;font-style:italic;">Sin datos suficientes</span></div>';
  }
  return lista.slice(0, 3).map((v) => {
    const cls = tipo === 'mejoraron' ? 'down' : 'up';
    const signo = v.delta > 0 ? '+' : '';
    return '<div class="fila"><span class="nombre">' + titleCase(v.tecnico) + '</span><span class="num ' + cls + '">' + signo + v.delta + 'pts</span></div>';
  }).join('');
}

function claseEstado(tasa, meta) { return tasa <= meta ? 'ok' : 'bad'; }
function textoEstado(tasa, meta) { return tasa <= meta ? '✅ Cumple la meta' : '⚠️ No cumple la meta'; }

function encontrarNavegador() {
  const candidatos = [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
  ];
  return candidatos.find((p) => fs.existsSync(p)) || null;
}

function construirHtml(r, i, logoBase64, periodoActual, generadoEl) {
  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Resumen para correo</title>
<style>
  :root{
    --panel:#ffffff; --panel-2:#f5f7f9; --border:#e0e5ea; --text:#22303f; --text-dim:#6b7a8c;
    --cobra-navy:#003c71; --celeste:#29a9e0; --celeste-soft:#e8f6fd;
    --promotor:#1fa971; --promotor-bg:#e2f6ee; --detractor:#e2523e; --detractor-bg:#fce4e0;
  }
  *{box-sizing:border-box;}
  body{ margin:0; font-family:'Segoe UI', Arial, sans-serif; background:#ffffff; color:var(--text); }
  #lienzo{ width:760px; background:#ffffff; }
  header{ background:linear-gradient(120deg,#ffffff 0%,var(--celeste-soft) 60%,#dcf1fb 100%); padding:28px 32px 22px; border-bottom:4px solid var(--celeste); }
  .brand-row{ display:flex; align-items:center; gap:14px; margin-bottom:14px; }
  .brand-row img{ height:38px; }
  .brand-divider{ width:1px; height:28px; background:var(--border); }
  .eyebrow{ text-transform:uppercase; letter-spacing:.12em; font-size:11.5px; color:var(--celeste); font-weight:800; }
  h1{ margin:0 0 4px; font-size:24px; font-weight:800; color:var(--cobra-navy); }
  .subtitle{ color:#3a4a5c; font-size:13px; }
  main{ padding:26px 32px 30px; }
  .grid{ display:grid; grid-template-columns:1fr 1fr; gap:18px; }
  .card{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:20px 22px; box-shadow:0 4px 14px rgba(20,50,80,.05); }
  .card .cab{ display:flex; align-items:center; gap:8px; margin-bottom:2px; }
  .card .cab .icono{ font-size:20px; }
  .card h2{ margin:0; font-size:16px; color:var(--cobra-navy); }
  .card .periodo{ font-size:11.5px; color:var(--text-dim); margin-bottom:14px; }
  .valor-linea{ display:flex; align-items:baseline; gap:10px; margin-bottom:4px; }
  .valor{ font-size:32px; font-weight:800; }
  .valor.ok{ color:var(--promotor); }
  .valor.bad{ color:var(--detractor); }
  .meta-txt{ font-size:12px; color:var(--text-dim); }
  .pill{ display:inline-flex; align-items:center; gap:5px; padding:3px 11px; border-radius:20px; font-size:11.5px; font-weight:700; margin:6px 0 14px; }
  .pill.ok{ background:var(--promotor-bg); color:var(--promotor); }
  .pill.bad{ background:var(--detractor-bg); color:var(--detractor); }
  .bloque-tit{ font-size:11px; text-transform:uppercase; letter-spacing:.05em; font-weight:700; color:var(--text-dim); margin:12px 0 6px; }
  .fila{ display:flex; justify-content:space-between; font-size:12.5px; padding:4px 0; border-bottom:1px solid var(--panel-2); }
  .fila .nombre{ color:var(--text); }
  .fila .num.up{ color:var(--detractor); font-weight:700; }
  .fila .num.down{ color:var(--promotor); font-weight:700; }
  footer{ text-align:center; padding:16px; color:var(--text-dim); font-size:11px; border-top:1px solid var(--border); }
</style>
</head>
<body>
<div id="lienzo">
  <header>
    <div class="brand-row">
      <img src="data:image/png;base64,${logoBase64}" alt="Cobra">
      <div class="brand-divider"></div>
      <div class="eyebrow">Supervisor · COBRA</div>
    </div>
    <h1>Resumen Diario</h1>
    <div class="subtitle">Repetido Reparado y Averías de Infancia · ${periodoActual}</div>
  </header>
  <main>
    <div class="grid">
      <div class="card">
        <div class="cab"><span class="icono">🔁</span><h2>Repetido Reparado</h2></div>
        <div class="periodo">Reparaciones que vuelven a fallar dentro de 30 días</div>
        <div class="valor-linea"><span class="valor ${claseEstado(r.tasaGlobal, r.meta)}">${r.tasaGlobal}%</span><span class="meta-txt">Meta ${r.meta}%</span></div>
        <div class="pill ${claseEstado(r.tasaGlobal, r.meta)}">${textoEstado(r.tasaGlobal, r.meta)}</div>
        <div class="bloque-tit">🟢 Más mejoraron</div>
        ${filasVariacion(r.masMejoraron, 'mejoraron')}
        <div class="bloque-tit">🔴 Más empeoraron</div>
        ${filasVariacion(r.masEmpeoraron, 'empeoraron')}
      </div>
      <div class="card">
        <div class="cab"><span class="icono">🏠</span><h2>Averías de Infancia</h2></div>
        <div class="periodo">Instalaciones que fallan poco después de instaladas</div>
        <div class="valor-linea"><span class="valor ${claseEstado(i.tasaGlobal, i.meta)}">${i.tasaGlobal}%</span><span class="meta-txt">Meta ${i.meta}%</span></div>
        <div class="pill ${claseEstado(i.tasaGlobal, i.meta)}">${textoEstado(i.tasaGlobal, i.meta)}</div>
        <div class="bloque-tit">🟢 Más mejoraron</div>
        ${filasVariacion(i.masMejoraron, 'mejoraron')}
        <div class="bloque-tit">🔴 Más empeoraron</div>
        ${filasVariacion(i.masEmpeoraron, 'empeoraron')}
      </div>
    </div>
  </main>
  <footer>Generado ${generadoEl} · Datos de Reincidencias/Infancia COBRA (zona Punta Arenas / Coyhaique)</footer>
</div>
</body>
</html>`;
}

function main() {
  const rPath = path.join(carpeta, 'Dashboard_Reincidencias.html');
  const iPath = path.join(carpeta, 'Dashboard_Infancia.html');
  if (!fs.existsSync(rPath) || !fs.existsSync(iPath)) {
    console.log('AVISO: no se encontraron los dashboards de Reincidencias/Infancia -- se omite el resumen para correo (corre primero generar_reincidencias.js y generar_infancia.js).');
    return;
  }

  let r, i;
  try {
    r = getData(rPath);
    i = getData(iPath);
  } catch (err) {
    console.log('AVISO: no se pudo leer los datos de los dashboards (' + err.message + ') -- se omite el resumen para correo.');
    return;
  }

  const logoPath = path.join(carpeta, 'logo-cobra.png');
  const logoBase64 = fs.existsSync(logoPath) ? fs.readFileSync(logoPath).toString('base64') : '';
  const periodoActual = (r.archivos && r.archivos[0] && r.archivos[0].label) || r.periodo;
  const generadoEl = new Date().toLocaleString('es-CL');

  const html = construirHtml(r, i, logoBase64, periodoActual, generadoEl);
  fs.writeFileSync(TEMP_HTML, html, 'utf8');

  const navegador = encontrarNavegador();
  if (!navegador) {
    console.log('AVISO: no se encontro Chrome ni Edge instalado en las rutas habituales -- no se pudo generar la imagen del resumen.');
    return;
  }

  try {
    // Ventana generosamente alta: el contenido real siempre termina antes;
    // sharp().trim() recorta despues el margen blanco sobrante de abajo,
    // asi no hace falta calcular la altura exacta de antemano.
    execFileSync(navegador, [
      '--headless=new', '--disable-gpu', '--hide-scrollbars',
      '--window-size=760,1000',
      '--screenshot=' + TEMP_PNG,
      'file:///' + TEMP_HTML.replace(/\\/g, '/'),
    ], { stdio: 'ignore', timeout: 30000 });

    const sharp = require('sharp');
    sharp(TEMP_PNG).trim().toFile(OUTPUT_PNG).then(() => {
      fs.unlinkSync(TEMP_PNG);
      fs.unlinkSync(TEMP_HTML);
      console.log('==> Resumen para correo generado:', OUTPUT_PNG);
    }).catch((err) => {
      console.log('AVISO: fallo al recortar la imagen (' + err.message + '); se deja sin recortar.');
      fs.renameSync(TEMP_PNG, OUTPUT_PNG);
      fs.unlinkSync(TEMP_HTML);
    });
  } catch (err) {
    console.log('AVISO: fallo al generar la imagen del resumen (' + err.message + ').');
    if (fs.existsSync(TEMP_HTML)) fs.unlinkSync(TEMP_HTML);
  }
}

main();
