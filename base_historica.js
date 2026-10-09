// Modulo compartido para la "base historica": dos Excel (uno para Repetido
// Reparado, otro para Averias de Infancia) guardados en "Archivos excel\",
// cada uno con UNA HOJA POR MES. Cada corrida de generar_reincidencias.js /
// generar_infancia.js agrega una fila-resumen (fecha, total, tasa, meta,
// resultado) a la hoja del mes que corresponda -- se crea sola si es la
// primera corrida de ese mes.
//
// construir_base_historica.js usa este mismo modulo para poblar el
// historial retroactivo a partir de los Excel individuales ya existentes.

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const carpeta = __dirname;
const carpetaExcel = path.join(carpeta, 'Archivos excel');

const CONFIG = {
  reincidencias: {
    archivoBase: 'Base_Reincidencias_COBRA.xlsx',
    tituloTotal: 'Total Reparaciones',
    tituloProblemas: 'Reincidencias',
  },
  infancia: {
    archivoBase: 'Base_Infancia_COBRA.xlsx',
    tituloTotal: 'Total Instalaciones',
    tituloProblemas: 'Averias Infancia',
  },
};

// La hoja "Estadistica" de cada Excel individual siempre trae la fila GLOBAL
// en la fila 5, columnas A-E: Total | Problemas | Tasa (0-1) | Meta (0-1) | Resultado.
function leerFilaGlobal(wsEstadistica, fecha) {
  const fila = wsEstadistica.getRow(5);
  const total = Number(fila.getCell(1).value);
  const problemas = Number(fila.getCell(2).value);
  const tasa = Number(fila.getCell(3).value);
  const meta = Number(fila.getCell(4).value);
  const resultado = String(fila.getCell(5).value || '');
  if (!Number.isFinite(total) || !Number.isFinite(tasa)) {
    throw new Error('la fila GLOBAL de "Estadistica" no tiene el formato esperado');
  }
  return {
    fecha,
    total,
    problemas,
    tasa: +(tasa * 100).toFixed(2),
    meta: +(meta * 100).toFixed(2),
    resultado,
  };
}

function nombreHojaMes(fecha) {
  return fecha.getFullYear() + '-' + String(fecha.getMonth() + 1).padStart(2, '0');
}

async function agregarFilaBase(tipo, fila) {
  const cfg = CONFIG[tipo];
  if (!cfg) throw new Error('tipo desconocido: ' + tipo);
  if (!fs.existsSync(carpetaExcel)) fs.mkdirSync(carpetaExcel, { recursive: true });
  const rutaBase = path.join(carpetaExcel, cfg.archivoBase);

  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(rutaBase)) {
    await wb.xlsx.readFile(rutaBase);
  }

  // IMPORTANTE: se usan filas/columnas por POSICION (array), no por "key" de
  // ws.columns -- ExcelJS no conserva el mapeo de keys al releer un archivo
  // ya guardado, y addRow({...}) por key falla ("Out of bounds") en cuanto
  // se reabre la base para agregar una segunda fila a una hoja existente.
  const nombreHoja = nombreHojaMes(fila.fecha);
  let ws = wb.getWorksheet(nombreHoja);
  if (!ws) {
    ws = wb.addWorksheet(nombreHoja, { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.getColumn(1).width = 20;
    ws.getColumn(2).width = 18;
    ws.getColumn(3).width = 16;
    ws.getColumn(4).width = 12;
    ws.getColumn(5).width = 12;
    ws.getColumn(6).width = 14;
    const headerRow = ws.addRow(['Fecha corrida', cfg.tituloTotal, cfg.tituloProblemas, 'Tasa %', 'Meta %', 'Resultado']);
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF003C71' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
    headerRow.height = 22;
  }

  const nuevaFila = ws.addRow([fila.fecha, fila.total, fila.problemas, fila.tasa, fila.meta, fila.resultado]);
  nuevaFila.getCell(1).numFmt = 'dd-mm-yyyy hh:mm';
  nuevaFila.getCell(6).font = {
    bold: true,
    color: { argb: fila.resultado === 'CUMPLE' ? 'FF1FA971' : 'FFE2523E' },
  };

  await wb.xlsx.writeFile(rutaBase);
}

module.exports = { leerFilaGlobal, agregarFilaBase, nombreHojaMes, carpetaExcel };
