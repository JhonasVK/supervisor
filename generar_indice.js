// Genera index.html: pagina de indice con acceso a los dos informes
// (Repetido Reparado / Averias de Infancia). Se regenera cada vez que
// corre Generar_Reporte_Reincidencias.bat, despues de los otros dos scripts.

const fs = require('fs');
const path = require('path');

const carpeta = __dirname;

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

const html = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Informes COBRA · Indice</title>
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
  .brand-row img{ height:46px; }
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
  <h1>Informes COBRA</h1>
  <div class="subtitle">Elige el informe que quieres revisar.</div>
</header>

<main>
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
  </div>
</main>

<footer>Generado automaticamente por generar_indice.js</footer>

</body>
</html>`;

fs.writeFileSync(path.join(carpeta, 'index.html'), html, 'utf8');
console.log('Indice generado:', path.join(carpeta, 'index.html'));
