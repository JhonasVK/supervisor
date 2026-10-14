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

function filasVariacion(lista, tipo, limite = 3) {
  if (!lista || lista.length === 0) {
    return '<div class="fila"><span class="nombre" style="color:#8a97a6;font-style:italic;">Sin datos suficientes</span></div>';
  }
  return lista.slice(0, limite).map((v) => {
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
  #lienzo{ width:820px; background:#ffffff; }
  header{ display:flex; align-items:center; gap:8px; background:linear-gradient(120deg,#ffffff 0%,var(--celeste-soft) 60%,#dcf1fb 100%); padding:5px 18px; border-bottom:3px solid var(--celeste); }
  header img{ height:26px; }
  .brand-divider{ width:1px; height:20px; background:var(--border); }
  .eyebrow{ text-transform:uppercase; letter-spacing:.1em; font-size:11px; color:var(--celeste); font-weight:800; }
  .titulo{ font-size:18px; font-weight:800; color:var(--cobra-navy); line-height:1.1; }
  .titulo span{ font-weight:400; color:var(--text-dim); font-size:12.5px; }
  .fila-reporte{ display:flex; align-items:stretch; gap:0; padding:6px 18px; border-bottom:1px solid var(--border); }
  .fila-reporte:last-of-type{ border-bottom:none; }
  .bloque-izq{ display:flex; align-items:center; gap:8px; width:200px; flex:none; }
  .bloque-izq .icono{ font-size:28px; }
  .card-tit{ font-size:15px; font-weight:700; color:var(--cobra-navy); margin-bottom:0; line-height:1.05; }
  .valor-linea{ display:flex; align-items:baseline; gap:4px; }
  .valor{ font-size:30px; font-weight:800; line-height:1; }
  .valor.ok{ color:var(--promotor); }
  .valor.bad{ color:var(--detractor); }
  .meta-txt{ font-size:11px; color:var(--text-dim); }
  .pill{ display:inline-flex; align-items:center; gap:3px; padding:1px 8px; border-radius:20px; font-size:11px; font-weight:700; margin-top:1px; }
  .pill.ok{ background:var(--promotor-bg); color:var(--promotor); }
  .pill.bad{ background:var(--detractor-bg); color:var(--detractor); }
  .divisor{ width:1px; background:var(--panel-2); margin:0 8px; }
  .col-lista{ flex:0 0 260px; min-width:0; }
  .bloque-tit{ font-size:10.5px; text-transform:uppercase; letter-spacing:.05em; font-weight:700; color:var(--text-dim); margin-bottom:1px; line-height:1.05; }
  .fila{ display:flex; align-items:baseline; gap:6px; font-size:13.5px; padding:0; line-height:1.2; }
  .fila .nombre{ color:var(--text); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; flex:1; }
  .fila .num.up{ color:var(--detractor); font-weight:700; flex:none; }
  .fila .num.down{ color:var(--promotor); font-weight:700; flex:none; }
  footer{ text-align:center; padding:3px; color:var(--text-dim); font-size:10px; border-top:1px solid var(--border); }
</style>
</head>
<body>
<div id="lienzo">
  <header>
    <img src="data:image/png;base64,${logoBase64}" alt="Cobra">
    <div class="brand-divider"></div>
    <div>
      <div class="eyebrow">Supervisor · COBRA</div>
      <div class="titulo">Resumen Diario <span>· ${periodoActual}</span></div>
    </div>
  </header>

  <div class="fila-reporte">
    <div class="bloque-izq">
      <span class="icono">🔁</span>
      <div>
        <div class="card-tit">Repetido Reparado</div>
        <div class="valor-linea"><span class="valor ${claseEstado(r.tasaGlobal, r.meta)}">${r.tasaGlobal}%</span><span class="meta-txt">Meta ${r.meta}%</span></div>
        <div class="pill ${claseEstado(r.tasaGlobal, r.meta)}">${textoEstado(r.tasaGlobal, r.meta)}</div>
      </div>
    </div>
    <div class="divisor"></div>
    <div class="col-lista">
      <div class="bloque-tit">🟢 Más mejoraron</div>
      ${filasVariacion(r.masMejoraron, 'mejoraron', 2)}
    </div>
    <div class="divisor"></div>
    <div class="col-lista">
      <div class="bloque-tit">🔴 Más empeoraron</div>
      ${filasVariacion(r.masEmpeoraron, 'empeoraron', 2)}
    </div>
  </div>

  <div class="fila-reporte">
    <div class="bloque-izq">
      <span class="icono">🏠</span>
      <div>
        <div class="card-tit">Averías de Infancia</div>
        <div class="valor-linea"><span class="valor ${claseEstado(i.tasaGlobal, i.meta)}">${i.tasaGlobal}%</span><span class="meta-txt">Meta ${i.meta}%</span></div>
        <div class="pill ${claseEstado(i.tasaGlobal, i.meta)}">${textoEstado(i.tasaGlobal, i.meta)}</div>
      </div>
    </div>
    <div class="divisor"></div>
    <div class="col-lista">
      <div class="bloque-tit">🟢 Más mejoraron</div>
      ${filasVariacion(i.masMejoraron, 'mejoraron', 2)}
    </div>
    <div class="divisor"></div>
    <div class="col-lista">
      <div class="bloque-tit">🔴 Más empeoraron</div>
      ${filasVariacion(i.masEmpeoraron, 'empeoraron', 2)}
    </div>
  </div>

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
      '--window-size=820,700',
      '--screenshot=' + TEMP_PNG,
      'file:///' + TEMP_HTML.replace(/\\/g, '/'),
    ], { stdio: 'ignore', timeout: 30000 });

    const sharp = require('sharp');
    // Se recorta el margen blanco sobrante y despues se reduce a un ancho
    // fijo (ANCHO_FINAL) -- el diseno es ancho y bajo (franjas horizontales)
    // a proposito, para que se vea bien en un correo sin ocupar mucho alto.
    const ANCHO_FINAL = 820;
    sharp(TEMP_PNG).trim().resize({ width: ANCHO_FINAL }).toFile(OUTPUT_PNG).then(() => {
      fs.unlinkSync(TEMP_PNG);
      fs.unlinkSync(TEMP_HTML);
      console.log('==> Resumen para correo generado:', OUTPUT_PNG);
    }).catch((err) => {
      console.log('AVISO: fallo al recortar/reducir la imagen (' + err.message + '); se deja sin procesar.');
      fs.renameSync(TEMP_PNG, OUTPUT_PNG);
      fs.unlinkSync(TEMP_HTML);
    });
  } catch (err) {
    console.log('AVISO: fallo al generar la imagen del resumen (' + err.message + ').');
    if (fs.existsSync(TEMP_HTML)) fs.unlinkSync(TEMP_HTML);
  }
}

main();
