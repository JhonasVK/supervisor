// Construye (o reconstruye desde cero) las bases historicas a partir de los
// Excel individuales que ya existen en "Archivos excel\". Se usa UNA VEZ para
// poblar el historial retroactivo -- de ahi en adelante, cada corrida de
// generar_reincidencias.js / generar_infancia.js ya se encarga de agregar su
// propia fila via agregarFilaBase() (ver base_historica.js).
//
// Uso: node construir_base_historica.js

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { agregarFilaBase, leerFilaGlobal } = require('./base_historica');

const carpeta = __dirname;
const carpetaExcel = path.join(carpeta, 'Archivos excel');

function parsearTimestampDeNombre(nombre, prefijo) {
  // Ej: "Reincidencias_COBRA_2026-09-09_1045.xlsx" -> 2026-09-09 10:45
  const m = nombre.match(new RegExp('^' + prefijo + '_COBRA_(\\d{4})-(\\d{2})-(\\d{2})_(\\d{2})(\\d{2})\\.xlsx$', 'i'));
  if (!m) return null;
  const [, y, mo, d, h, mi] = m;
  return new Date(+y, +mo - 1, +d, +h, +mi);
}

async function procesar(prefijo, hojaGlobal, tipo) {
  const archivos = fs.readdirSync(carpetaExcel)
    .filter((f) => f.toLowerCase().startsWith((prefijo + '_COBRA_').toLowerCase()) && f.toLowerCase().endsWith('.xlsx'))
    .map((f) => ({ nombre: f, fecha: parsearTimestampDeNombre(f, prefijo) }))
    .filter((f) => f.fecha !== null)
    .sort((a, b) => a.fecha - b.fecha);

  console.log(tipo + ': ' + archivos.length + ' archivos encontrados en "Archivos excel\\".');

  let ok = 0, fallidos = 0;
  for (const { nombre, fecha } of archivos) {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.readFile(path.join(carpetaExcel, nombre));
      const ws = wb.getWorksheet(hojaGlobal);
      if (!ws) throw new Error('no se encontro la hoja "' + hojaGlobal + '"');
      const fila = leerFilaGlobal(ws, fecha);
      await agregarFilaBase(tipo, fila);
      ok++;
    } catch (err) {
      console.log('  AVISO: no se pudo leer ' + nombre + ' (' + err.message + ')');
      fallidos++;
    }
  }
  console.log(tipo + ': ' + ok + ' filas agregadas a la base' + (fallidos ? ', ' + fallidos + ' archivos con error' : '') + '.');
}

async function main() {
  if (!fs.existsSync(carpetaExcel)) {
    console.log('No existe la carpeta "Archivos excel" en ' + carpeta + '. Nada que importar.');
    return;
  }
  await procesar('Reincidencias', 'Estadistica', 'reincidencias');
  await procesar('Infancia', 'Estadistica', 'infancia');
  console.log('Listo.');
}

main().catch((err) => {
  console.error('ERROR:', err.message);
  process.exitCode = 1;
});
