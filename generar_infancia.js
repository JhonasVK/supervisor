// Analisis de Averias de Infancia (instalaciones que generaron una reparacion
// dentro de su periodo de infancia). Lee el CSV "p22-Averias-infancia..._COBRA_...csv"
// desde la carpeta bbdd\ (subcarpeta de esta misma carpeta) y genera un Excel +
// un dashboard HTML con: tasa de infancia por agencia/producto/tecnico instalador
// (meta 2.5%), distribucion de causas y claves de cierre de la reparacion,
// tiempo hasta la infancia, y si la reparacion la resolvio el mismo tecnico
// que hizo la instalacion.
//
// Para actualizar el analisis: sobreescribe el CSV dentro de la carpeta bbdd\
// con el nuevo mes y vuelve a ejecutar "Generar_Reporte_Reincidencias.bat".

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const carpeta = __dirname;
const carpetaBbdd = path.join(carpeta, 'bbdd');
const META = 0.025; // 2.5%

function encontrarCsvOrigen() {
  if (!fs.existsSync(carpetaBbdd)) {
    throw new Error('No existe la carpeta "bbdd" en ' + carpeta);
  }
  const candidatos = fs
    .readdirSync(carpetaBbdd)
    .filter((f) => /^p22-averias-infancia.*COBRA.*\.csv$/i.test(f))
    .map((f) => ({ nombre: f, mtime: fs.statSync(path.join(carpetaBbdd, f)).mtimeMs }))
    .sort((a, b) => b.mtime - a.mtime);

  if (candidatos.length === 0) {
    throw new Error(
      'No se encontro ningun archivo "p22-averias-infancia..._COBRA_....csv" en ' + carpetaBbdd
    );
  }
  return path.join(carpetaBbdd, candidatos[0].nombre);
}

function parseCsvLine(line) {
  const result = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"') {
        if (line[i + 1] === '"') { cur += '"'; i++; }
        else inQuotes = false;
      } else cur += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ';') { result.push(cur); cur = ''; }
      else cur += c;
    }
  }
  result.push(cur);
  return result;
}

function leerCsv(csvPath) {
  const raw = fs.readFileSync(csvPath, 'utf8');
  const lines = raw.split(/\r?\n/).filter((l) => l.length > 0);
  const headers = parseCsvLine(lines[0]);
  return lines.slice(1).map((l) => {
    const vals = parseCsvLine(l);
    const obj = {};
    headers.forEach((h, i) => (obj[h] = vals[i] !== undefined ? vals[i] : ''));
    return obj;
  });
}

function timestamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return (
    d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) +
    '_' + pad(d.getHours()) + pad(d.getMinutes())
  );
}

const NOMBRES_MES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
function nombreMes(slug) {
  const [y, m] = slug.split('-');
  return NOMBRES_MES[parseInt(m, 10) - 1] + ' ' + y;
}

// El "mes" del informe se toma de partition_date del CSV (la fecha del snapshot de datos),
// no de la fecha del sistema, para que el historial quede ligado a los datos reales.
function obtenerSlugDesdeCsv(rows) {
  const pd = (rows[0] && rows[0]['partition_date']) || '';
  if (/^\d{4}-\d{2}/.test(pd)) return pd.slice(0, 7);
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function listarMesesExistentes(carpeta, prefijo) {
  const re = new RegExp('^' + prefijo + '_(\\d{4}-\\d{2})\\.html$');
  return fs.readdirSync(carpeta)
    .map((f) => f.match(re))
    .filter(Boolean)
    .map((m) => m[1]);
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function mediana(arr) {
  if (arr.length === 0) return null;
  const s = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 === 0 ? (s[mid - 1] + s[mid]) / 2 : s[mid];
}
function promedio(arr) {
  if (arr.length === 0) return null;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

async function main() {
  const csvPath = encontrarCsvOrigen();
  console.log('Archivo origen:', csvPath);

  const rows = leerCsv(csvPath);
  // Las instalaciones (Alta) y los traslados (T, tambien involucran una reconexion
  // fisica) cuentan como base para la tasa de infancia. Las Reparaciones (R) no,
  // porque no son una instalacion y quedan fuera del calculo.
  const instalaciones = rows.filter((r) => ['A', 'T'].includes((r['vpi_tipo_trabajo_producto'] || '').trim()));
  const inf = instalaciones.filter((r) => (r['infancia'] || '').trim() === '1');

  inf.sort((a, b) => {
    const ta = a['toa_provider_name'] || '';
    const tb = b['toa_provider_name'] || '';
    if (ta !== tb) return ta.localeCompare(tb);
    return (a['toa_xa_order_creation_date'] || '').localeCompare(b['toa_xa_order_creation_date'] || '');
  });

  console.log('Total registros:', rows.length, '| Instalaciones (Alta+Traslado):', instalaciones.length, '| Averias de infancia:', inf.length);

  // ================== EXCEL ==================
  const columns = [
    { header: 'Folio Instalacion', key: 'folio1', width: 16 },
    { header: 'Fecha Instalacion', key: 'fecha1', width: 20 },
    { header: 'Tecnico Instalador', key: 'tec1', width: 32 },
    { header: 'RUT Tecnico Instalador', key: 'rut1', width: 16 },
    { header: 'Agencia', key: 'agencia1', width: 14 },
    { header: 'Producto', key: 'producto', width: 12 },
    { header: 'Folio Reparacion', key: 'folio2', width: 18 },
    { header: 'Fecha Reparacion', key: 'fecha2', width: 20 },
    { header: 'Tecnico Reparador', key: 'tec2', width: 32 },
    { header: 'RUT Tecnico Reparador', key: 'rut2', width: 16 },
    { header: 'Clave de Cierre', key: 'clave2', width: 16 },
    { header: 'Causa', key: 'causa2', width: 34 },
    { header: 'Resolucion (Nivel 2)', key: 'nivel2', width: 36 },
    { header: 'Mismo Tecnico?', key: 'mismo', width: 14 },
    { header: 'Dias hasta Infancia', key: 'dias', width: 16 },
  ];

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Analisis Averias de Infancia';
  wb.created = new Date();

  const ws = wb.addWorksheet('Infancia Detalle', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = columns;

  inf.forEach((r) => {
    const diasVal = r['q_dias_infancia'];
    ws.addRow({
      folio1: r['toa_appt_number'],
      fecha1: r['toa_eta'] || r['toa_xa_order_creation_date'],
      tec1: r['toa_provider_name'],
      rut1: r['toa_provider_external_id'],
      agencia1: r['toa_xa_original_agency'],
      producto: r['vpi_producto'],
      folio2: r['rmdy_folio_repara'],
      fecha2: r['rmdy_fecha_creacion_repara'],
      tec2: r['rmdy_nombre_tecnico'],
      rut2: r['rmdy_rut_tecnico'],
      clave2: r['rmdy_clave_cierre'],
      causa2: r['rmdy_causa'],
      nivel2: r['rmdy_nivel2_resolucion'],
      mismo: r['rmdy_nombre_tecnico'] === r['toa_provider_name'] ? 'SI' : 'NO',
      dias: diasVal === '' ? null : Number(diasVal),
    });
  });

  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' },
    };
  });
  headerRow.height = 30;

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    row.eachCell((cell) => {
      cell.font = { name: 'Arial', size: 10 };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        right: { style: 'thin', color: { argb: 'FFD9D9D9' } },
      };
    });
    if (row.getCell('mismo').value === 'SI') {
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4E4' } };
      });
    }
  });

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: inf.length + 1, column: columns.length } };

  // ---------- Hoja Estadistica: tasa de infancia (%) vs meta 2.5% ----------
  // Nota: toa_appt_number NO es unico (una misma cita puede tener varias filas,
  // una por producto instalado), asi que no se puede usar como clave de un Set
  // para "es infancia" -- eso duplicaria el conteo en citas con mas de una fila.
  // Se usa el campo "infancia" de cada fila directamente.
  function agruparPor(campo) {
    const grupos = {};
    instalaciones.forEach((r) => {
      const key = (r[campo] || '').trim() || '(sin dato)';
      if (!grupos[key]) grupos[key] = { total: 0, infancia: 0 };
      grupos[key].total += 1;
      if ((r['infancia'] || '').trim() === '1') grupos[key].infancia += 1;
    });
    return grupos;
  }
  function metaPorTecnico() {
    const info = {};
    instalaciones.forEach((r) => {
      const tec = (r['toa_provider_name'] || '').trim() || '(sin dato)';
      if (!info[tec]) info[tec] = { agencias: {}, ruts: {} };
      const ag = (r['toa_xa_original_agency'] || '').trim() || '(sin dato)';
      const rut = (r['toa_provider_external_id'] || '').trim();
      info[tec].agencias[ag] = (info[tec].agencias[ag] || 0) + 1;
      if (rut) info[tec].ruts[rut] = (info[tec].ruts[rut] || 0) + 1;
    });
    const masFrecuente = (obj) => Object.entries(obj).sort((a, b) => b[1] - a[1])[0]?.[0] || '';
    const resultado = {};
    Object.keys(info).forEach((tec) => {
      resultado[tec] = { agencia: masFrecuente(info[tec].agencias), rut: masFrecuente(info[tec].ruts) };
    });
    return resultado;
  }

  const statsGlobal = { total: instalaciones.length, infancia: inf.length };
  const statsPorAgencia = agruparPor('toa_xa_original_agency');
  const statsPorProducto = agruparPor('vpi_producto');
  const statsPorTecnico = agruparPor('toa_provider_name');
  const metaTec = metaPorTecnico();

  const est = wb.addWorksheet('Estadistica');
  est.columns = [{ width: 34 }, { width: 16 }, { width: 20 }, { width: 20 }, { width: 14 }, { width: 16 }];

  function estilizarTitulo(cell, texto) {
    cell.value = texto;
    cell.font = { name: 'Arial', bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
    cell.alignment = { vertical: 'middle' };
  }
  function estilizarHeader(row) {
    row.eachCell((cell) => {
      cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' },
      };
    });
  }
  function pintarTasa(cell, tasa) {
    cell.value = tasa;
    cell.numFmt = '0.0%';
    cell.font = { name: 'Arial', bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: tasa > META ? 'FFFCE4E4' : 'FFE2EFDA' } };
    cell.font.color = { argb: tasa > META ? 'FFC00000' : 'FF375623' };
  }
  function pintarCumple(cell, tasa) {
    const cumple = tasa <= META;
    cell.value = cumple ? 'CUMPLE' : 'NO CUMPLE';
    cell.font = { name: 'Arial', bold: true, color: { argb: cumple ? 'FF375623' : 'FFC00000' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cumple ? 'FFE2EFDA' : 'FFFCE4E4' } };
    cell.alignment = { horizontal: 'center' };
  }

  est.mergeCells('A1:F1');
  estilizarTitulo(est.getCell('A1'), 'TASA DE AVERIAS DE INFANCIA (%) — Meta: no superar el ' + (META * 100).toFixed(1) + '%');
  est.getRow(1).height = 24;

  est.addRow([]);
  est.mergeCells('A3:F3');
  est.getCell('A3').value = 'GLOBAL';
  est.getCell('A3').font = { name: 'Arial', bold: true, size: 12 };
  const headerGlobal = est.addRow(['Total Instalaciones', 'Averias de Infancia', 'Tasa Global', 'Meta', 'Resultado', '']);
  estilizarHeader(headerGlobal);
  const rowGlobal = est.addRow([statsGlobal.total, statsGlobal.infancia, null, META, null, '']);
  rowGlobal.getCell(1).font = { name: 'Arial' };
  rowGlobal.getCell(2).font = { name: 'Arial' };
  pintarTasa(rowGlobal.getCell(3), statsGlobal.total ? statsGlobal.infancia / statsGlobal.total : 0);
  rowGlobal.getCell(4).numFmt = '0.0%';
  rowGlobal.getCell(4).font = { name: 'Arial' };
  pintarCumple(rowGlobal.getCell(5), statsGlobal.total ? statsGlobal.infancia / statsGlobal.total : 0);

  function bloqueTabla(titulo, headers, statsObj, keyName) {
    est.addRow([]);
    est.addRow([]);
    const filaTitulo = est.rowCount + 1;
    est.mergeCells(`A${filaTitulo}:F${filaTitulo}`);
    est.getCell(`A${filaTitulo}`).value = titulo;
    est.getCell(`A${filaTitulo}`).font = { name: 'Arial', bold: true, size: 12 };
    const headerRowT = est.addRow(headers);
    estilizarHeader(headerRowT);
    Object.entries(statsObj)
      .sort((a, b) => (b[1].infancia / b[1].total) - (a[1].infancia / a[1].total))
      .forEach(([key, s]) => {
        const tasa = s.total ? s.infancia / s.total : 0;
        const row = est.addRow([key, s.total, s.infancia, null, META, null]);
        row.getCell(1).font = { name: 'Arial' };
        row.getCell(2).font = { name: 'Arial' };
        row.getCell(3).font = { name: 'Arial' };
        pintarTasa(row.getCell(4), tasa);
        row.getCell(5).numFmt = '0.0%';
        row.getCell(5).font = { name: 'Arial' };
        pintarCumple(row.getCell(6), tasa);
      });
  }

  bloqueTabla('POR AGENCIA', ['Agencia', 'Total Instalaciones', 'Averias de Infancia', 'Tasa', 'Meta', 'Resultado'], statsPorAgencia);
  bloqueTabla('POR PRODUCTO', ['Producto', 'Total Instalaciones', 'Averias de Infancia', 'Tasa', 'Meta', 'Resultado'], statsPorProducto);

  // Por Tecnico (con columnas extra Agencia/RUT)
  est.addRow([]);
  est.addRow([]);
  const filaTituloTecnico = est.rowCount + 1;
  est.mergeCells(`A${filaTituloTecnico}:F${filaTituloTecnico}`);
  est.getCell(`A${filaTituloTecnico}`).value = 'POR TECNICO INSTALADOR';
  est.getCell(`A${filaTituloTecnico}`).font = { name: 'Arial', bold: true, size: 12 };
  const headerTecnico = est.addRow(['Tecnico', 'RUT', 'Agencia', 'Total Instalaciones', 'Averias de Infancia', 'Tasa', 'Meta', 'Resultado']);
  est.getColumn(7).width = 12;
  est.getColumn(8).width = 14;
  estilizarHeader(headerTecnico);
  Object.entries(statsPorTecnico)
    .sort((a, b) => (b[1].infancia / b[1].total) - (a[1].infancia / a[1].total))
    .forEach(([tecnico, s]) => {
      const tasa = s.total ? s.infancia / s.total : 0;
      const meta = metaTec[tecnico] || { agencia: '', rut: '' };
      const row = est.addRow([tecnico, meta.rut, meta.agencia, s.total, s.infancia, null, META, null]);
      row.getCell(1).font = { name: 'Arial' };
      row.getCell(2).font = { name: 'Arial' };
      row.getCell(3).font = { name: 'Arial' };
      row.getCell(4).font = { name: 'Arial' };
      row.getCell(5).font = { name: 'Arial' };
      pintarTasa(row.getCell(6), tasa);
      row.getCell(7).numFmt = '0.0%';
      row.getCell(7).font = { name: 'Arial' };
      pintarCumple(row.getCell(8), tasa);
    });

  est.getColumn(1).width = 44;
  est.getColumn(2).width = 14;
  est.getColumn(3).width = 16;
  est.getColumn(4).width = 18;
  est.getColumn(5).width = 14;
  est.getColumn(6).width = 12;

  // ---------- Hoja Tiempo hasta la Infancia ----------
  const diasInf = inf.map((r) => Number(r['q_dias_infancia'])).filter((n) => !Number.isNaN(n));
  function agruparDiasPor(campo, esGlobalRow) {
    const grupos = {};
    inf.forEach((r) => {
      const key = (r[campo] || '').trim() || '(sin dato)';
      const dias = Number(r['q_dias_infancia']);
      if (Number.isNaN(dias)) return;
      if (!grupos[key]) grupos[key] = [];
      grupos[key].push(dias);
    });
    return grupos;
  }
  const diasPorTecnico = agruparDiasPor('toa_provider_name');
  const diasPorCausa = agruparDiasPor('rmdy_causa');

  const tr = wb.addWorksheet('Tiempo Infancia');
  tr.columns = [{ width: 34 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }];
  tr.mergeCells('A1:F1');
  estilizarTitulo(tr.getCell('A1'), 'TIEMPO HASTA LA INFANCIA (dias entre instalacion y reparacion)');
  tr.getRow(1).height = 24;

  tr.addRow([]);
  tr.mergeCells('A3:F3');
  tr.getCell('A3').value = 'GLOBAL';
  tr.getCell('A3').font = { name: 'Arial', bold: true, size: 12 };
  const headerGlobalT = tr.addRow(['Casos con dato', 'Promedio (dias)', 'Mediana (dias)', 'Minimo', 'Maximo', '% dentro de 7 dias']);
  estilizarHeader(headerGlobalT);
  const dentro7 = diasInf.filter((d) => d <= 7).length;
  const rowGlobalT = tr.addRow([
    diasInf.length,
    diasInf.length ? Number(promedio(diasInf).toFixed(1)) : null,
    diasInf.length ? mediana(diasInf) : null,
    diasInf.length ? Math.min(...diasInf) : null,
    diasInf.length ? Math.max(...diasInf) : null,
    diasInf.length ? dentro7 / diasInf.length : null,
  ]);
  rowGlobalT.eachCell((cell) => { cell.font = { name: 'Arial' }; });
  rowGlobalT.getCell(6).numFmt = '0.0%';

  function bloqueTiempo(titulo, statsMap) {
    tr.addRow([]);
    tr.addRow([]);
    const fila = tr.rowCount + 1;
    tr.mergeCells(`A${fila}:F${fila}`);
    tr.getCell(`A${fila}`).value = titulo;
    tr.getCell(`A${fila}`).font = { name: 'Arial', bold: true, size: 12 };
    const h = tr.addRow(['Categoria', 'Casos', 'Promedio (dias)', 'Mediana (dias)', 'Minimo', 'Maximo']);
    estilizarHeader(h);
    Object.entries(statsMap)
      .sort((a, b) => promedio(a[1]) - promedio(b[1]))
      .forEach(([key, dias]) => {
        const row = tr.addRow([key, dias.length, Number(promedio(dias).toFixed(1)), mediana(dias), Math.min(...dias), Math.max(...dias)]);
        row.eachCell((cell) => { cell.font = { name: 'Arial' }; });
      });
  }
  bloqueTiempo('POR CAUSA', diasPorCausa);
  bloqueTiempo('POR TECNICO INSTALADOR', diasPorTecnico);

  // ---------- Hoja Mismo Tecnico (instalador = reparador) ----------
  const mismoTotal = inf.length;
  const esMismoTecnico = (r) => r['rmdy_nombre_tecnico'] === r['toa_provider_name'];
  const mismoSi = inf.filter(esMismoTecnico).length;

  function agruparMismoPor(campo) {
    const grupos = {};
    inf.forEach((r) => {
      const key = (r[campo] || '').trim() || '(sin dato)';
      if (!grupos[key]) grupos[key] = { total: 0, mismo: 0 };
      grupos[key].total += 1;
      if (esMismoTecnico(r)) grupos[key].mismo += 1;
    });
    return grupos;
  }
  const mismoPorCausa = agruparMismoPor('rmdy_causa');

  const mt = wb.addWorksheet('Mismo Tecnico');
  mt.columns = [{ width: 34 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 14 }];
  mt.mergeCells('A1:E1');
  estilizarTitulo(mt.getCell('A1'), 'REPARACION DE INFANCIA ATENDIDA POR EL MISMO TECNICO QUE INSTALO, vs. OTRO');
  mt.getRow(1).height = 24;

  mt.addRow([]);
  mt.mergeCells('A3:E3');
  mt.getCell('A3').value = 'GLOBAL';
  mt.getCell('A3').font = { name: 'Arial', bold: true, size: 12 };
  const headerGlobalM = mt.addRow(['Casos', 'Mismo Tecnico (SI)', 'Distinto Tecnico (NO)', '% Mismo Tecnico', '% Distinto Tecnico']);
  estilizarHeader(headerGlobalM);
  const rowGlobalM = mt.addRow([mismoTotal, mismoSi, mismoTotal - mismoSi, mismoTotal ? mismoSi / mismoTotal : null, mismoTotal ? (mismoTotal - mismoSi) / mismoTotal : null]);
  rowGlobalM.eachCell((cell) => { cell.font = { name: 'Arial' }; });
  rowGlobalM.getCell(4).numFmt = '0.0%';
  rowGlobalM.getCell(5).numFmt = '0.0%';

  mt.addRow([]);
  mt.addRow([]);
  const filaTituloCausaM = mt.rowCount + 1;
  mt.mergeCells(`A${filaTituloCausaM}:E${filaTituloCausaM}`);
  mt.getCell(`A${filaTituloCausaM}`).value = 'POR CAUSA';
  mt.getCell(`A${filaTituloCausaM}`).font = { name: 'Arial', bold: true, size: 12 };
  const headerCausaM = mt.addRow(['Causa', 'Casos', 'Mismo Tecnico (SI)', 'Distinto Tecnico (NO)', '% Mismo Tecnico']);
  estilizarHeader(headerCausaM);
  Object.entries(mismoPorCausa)
    .sort((a, b) => (b[1].mismo / b[1].total) - (a[1].mismo / a[1].total))
    .forEach(([causa, s]) => {
      const row = mt.addRow([causa, s.total, s.mismo, s.total - s.mismo, s.total ? s.mismo / s.total : null]);
      row.eachCell((cell) => { cell.font = { name: 'Arial' }; });
      row.getCell(5).numFmt = '0.0%';
    });

  // ---------- Notas ----------
  const notas = wb.addWorksheet('Notas');
  notas.columns = [{ width: 100 }];
  const notasTexto = [
    'ANALISIS DE AVERIAS DE INFANCIA - COBRA',
    '',
    'Archivo origen: ' + path.basename(csvPath),
    'Generado: ' + new Date().toLocaleString('es-CL'),
    '',
    'Una "averia de infancia" es una instalacion o traslado (Alta o Traslado) que genero una reparacion (rmdy_*) dentro de su periodo de infancia, segun el campo "infancia" del archivo origen (1 = tuvo reparacion de infancia). El archivo tambien trae Reparaciones (R), que no son una instalacion y quedan fuera de este calculo (no se usan como base de la tasa).',
    '',
    '- "Mismo Tecnico?": indica si el tecnico que atendio la reparacion de infancia es el mismo que hizo la instalacion original.',
    '- Filas resaltadas en ROJO en "Infancia Detalle": la reparacion de infancia la atendio el MISMO tecnico que instalo (posible senal de instalacion mal hecha).',
    '- "Dias hasta Infancia": dias transcurridos entre la instalacion y la reparacion de infancia (campo q_dias_infancia del archivo origen).',
    '',
    'Hoja "Estadistica": tasa de averias de infancia (averias / total de instalaciones) calculada Global, por Agencia, por Producto y por Tecnico Instalador. Meta: no superar el ' + (META * 100).toFixed(1) + '%. Verde = cumple la meta, Rojo = no cumple.',
    '',
    'Hoja "Tiempo Infancia": promedio y mediana de dias transcurridos entre la instalacion y la reparacion de infancia, Global, por Causa y por Tecnico Instalador.',
    '',
    'Hoja "Mismo Tecnico": de las averias de infancia, indica si la atendio el MISMO tecnico que instalo o uno DISTINTO. Un % alto sugiere que el propio instalador detecto y corrigio su error; un % bajo sugiere que otro tecnico tuvo que corregirlo.',
    '',
    'Total de registros en el archivo origen (Altas + Reparaciones + Traslados): ' + rows.length,
    'Total de instalaciones (Altas + Traslados, base del calculo): ' + instalaciones.length,
    'Total de averias de infancia detectadas: ' + inf.length,
    '',
    'Para actualizar: reemplaza/sobreescribe el CSV en la carpeta bbdd\\ con los datos nuevos y vuelve a ejecutar Generar_Reporte_Reincidencias.bat. Cada ejecucion genera un archivo Excel nuevo (no sobreescribe reportes anteriores) y actualiza Dashboard_Infancia.html.',
  ];
  notasTexto.forEach((t, i) => {
    const row = notas.addRow([t]);
    row.getCell(1).font = i === 0 ? { name: 'Arial', bold: true, size: 12 } : { name: 'Arial', size: 10 };
    row.getCell(1).alignment = { wrapText: true, vertical: 'top' };
  });

  const outName = 'Infancia_COBRA_' + timestamp() + '.xlsx';
  const outPath = path.join(carpeta, outName);
  await wb.xlsx.writeFile(outPath);
  console.log('Excel generado:', outPath);

  // ================== DASHBOARD HTML (mismo formato del Informe NPS) ==================
  function aEntries(statsObj) {
    return Object.entries(statsObj).map(([label, s]) => ({
      label, total: s.total, infancia: s.infancia, tasa: s.total ? s.infancia / s.total : 0,
    }));
  }

  const agenciaEntries = aEntries(statsPorAgencia).sort((a, b) => b.tasa - a.tasa);
  const productoEntries = aEntries(statsPorProducto).sort((a, b) => b.tasa - a.tasa);
  const tecnicoEntries = aEntries(statsPorTecnico).sort((a, b) => b.tasa - a.tasa);

  const tasaGlobal = statsGlobal.total ? statsGlobal.infancia / statsGlobal.total : 0;
  const promedioGlobalDias = promedio(diasInf);
  const pctDentro7 = diasInf.length ? diasInf.filter((d) => d <= 7).length / diasInf.length : 0;
  const mismoPctGlobal = mismoTotal ? mismoSi / mismoTotal : 0;

  // Distribucion de causas y claves dentro del subconjunto de infancia (no es tasa vs poblacion)
  function distribucion(campo) {
    const counts = {};
    inf.forEach((r) => {
      const key = (r[campo] || '').trim() || '(sin dato)';
      counts[key] = (counts[key] || 0) + 1;
    });
    return Object.entries(counts)
      .map(([label, count]) => ({ label, count, pct: inf.length ? (count / inf.length) * 100 : 0 }))
      .sort((a, b) => b.count - a.count);
  }
  const causasDist = distribucion('rmdy_causa');
  const clavesDist = distribucion('rmdy_clave_cierre');

  const buckets = [
    { label: '0-1 dias', min: 0, max: 1 },
    { label: '2-3 dias', min: 2, max: 3 },
    { label: '4-7 dias', min: 4, max: 7 },
    { label: '8-15 dias', min: 8, max: 15 },
    { label: '16-30 dias', min: 16, max: 30 },
  ].map((b) => {
    const count = diasInf.filter((d) => d >= b.min && d <= b.max).length;
    return { label: b.label, count, pct: diasInf.length ? (count / diasInf.length) * 100 : 0 };
  });

  const causaPorTecnicoInf = {};
  inf.forEach((r) => {
    const t = (r['toa_provider_name'] || '').trim() || '(sin dato)';
    const c = (r['rmdy_causa'] || '').trim() || '(sin dato)';
    if (!causaPorTecnicoInf[t]) causaPorTecnicoInf[t] = {};
    causaPorTecnicoInf[t][c] = (causaPorTecnicoInf[t][c] || 0) + 1;
  });
  function causaFrecuenteTecnico(t) {
    const obj = causaPorTecnicoInf[t];
    if (!obj) return null;
    return Object.entries(obj).sort((a, b) => b[1] - a[1])[0];
  }

  const fechasAll = instalaciones.map((r) => r['toa_eta'] || r['toa_xa_order_creation_date']).filter(Boolean).sort();
  const periodoStr = fechasAll.length ? `${fechasAll[0].slice(0, 10)} al ${fechasAll[fechasAll.length - 1].slice(0, 10)}` : 'N/D';
  const agenciasLista = [...new Set(instalaciones.map((r) => (r['toa_xa_original_agency'] || '').trim()).filter(Boolean))].join(' y ');

  const elegiblesHighlight = tecnicoEntries.filter((e) => e.total >= 10);
  const highlightsTop = [...elegiblesHighlight].sort((a, b) => a.tasa - b.tasa).slice(0, 4).map((e) => {
    const cf = causaFrecuenteTecnico(e.label);
    return {
      tecnico: e.label, tasa: +(e.tasa * 100).toFixed(1), total: e.total, infancia: e.infancia,
      nota: e.infancia === 0
        ? `Sin averias de infancia en ${e.total} instalaciones`
        : `Su causa mas frecuente: ${cf ? cf[0] : 'N/D'}${cf ? ` (${cf[1]} caso${cf[1] === 1 ? '' : 's'})` : ''}`,
    };
  });
  const highlightsBottom = [...elegiblesHighlight].sort((a, b) => b.tasa - a.tasa).slice(0, 4).map((e) => {
    const cf = causaFrecuenteTecnico(e.label);
    return {
      tecnico: e.label, tasa: +(e.tasa * 100).toFixed(1), total: e.total, infancia: e.infancia,
      nota: `Su causa mas frecuente: ${cf ? cf[0] : 'N/D'}${cf ? ` (${cf[1]} caso${cf[1] === 1 ? '' : 's'})` : ''}`,
    };
  });

  const peorAgencia = agenciaEntries[0];
  const peorProducto = productoEntries[0];
  const peorTecnico = [...tecnicoEntries].filter((e) => e.total >= 10).sort((a, b) => b.tasa - a.tasa)[0];
  const mejorTecnico = [...tecnicoEntries].filter((e) => e.total >= 10).sort((a, b) => a.tasa - b.tasa)[0];
  const causaMasFrecuente = causasDist[0];

  const conclusiones = [
    `Tasa global de averias de infancia: <strong>${(tasaGlobal * 100).toFixed(1)}%</strong> — ${tasaGlobal <= META ? '<span style="color:#1fa971;font-weight:700;">CUMPLE</span>' : '<span style="color:#e2523e;font-weight:700;">NO CUMPLE</span>'} la meta del ${(META * 100).toFixed(1)}% (${statsGlobal.infancia} de ${statsGlobal.total} instalaciones tuvieron reparacion dentro de su infancia).`,
    peorAgencia ? `La agencia con mayor tasa es <strong>${escapeHtml(peorAgencia.label)}</strong> con ${(peorAgencia.tasa * 100).toFixed(1)}% (${peorAgencia.infancia}/${peorAgencia.total}).` : '',
    peorProducto ? `El producto con mayor tasa es <strong>${escapeHtml(peorProducto.label)}</strong> con ${(peorProducto.tasa * 100).toFixed(1)}%.` : '',
    causaMasFrecuente ? `La causa mas frecuente de la reparacion de infancia es <strong>${escapeHtml(causaMasFrecuente.label)}</strong> (${causaMasFrecuente.count} de ${inf.length} casos, ${causaMasFrecuente.pct.toFixed(0)}%).` : '',
    peorTecnico ? `El tecnico instalador con mayor tasa de infancia (minimo 10 instalaciones) es <strong>${escapeHtml(peorTecnico.label)}</strong> con ${(peorTecnico.tasa * 100).toFixed(1)}% (${peorTecnico.infancia}/${peorTecnico.total}).` : '',
    mejorTecnico ? `El mejor desempeno (minimo 10 instalaciones) es de <strong>${escapeHtml(mejorTecnico.label)}</strong> con ${(mejorTecnico.tasa * 100).toFixed(1)}%.` : '',
    diasInf.length ? `En promedio, la reparacion de infancia ocurre a los <strong>${promedioGlobalDias.toFixed(1)} dias</strong> de la instalacion; el ${(pctDentro7 * 100).toFixed(0)}% ocurre dentro de la primera semana.` : '',
    mismoTotal ? `El <strong>${(mismoPctGlobal * 100).toFixed(0)}%</strong> de las averias de infancia fueron atendidas por el mismo tecnico que hizo la instalacion original.` : '',
  ].filter(Boolean);

  const periodoSlug = obtenerSlugDesdeCsv(rows);
  const slugsExistentes = new Set(listarMesesExistentes(carpeta, 'Dashboard_Infancia'));
  slugsExistentes.add(periodoSlug);
  const archivos = [...slugsExistentes].sort().reverse().map((s) => ({
    slug: s,
    label: nombreMes(s),
    url: s === periodoSlug ? 'Dashboard_Infancia.html' : `Dashboard_Infancia_${s}.html`,
  }));

  const DATA = {
    archivoOrigen: path.basename(csvPath),
    generadoEl: new Date().toLocaleString('es-CL'),
    periodo: periodoStr,
    periodoSlug,
    archivos,
    agencias_lista: agenciasLista,
    meta: +(META * 100).toFixed(1),
    totalInstalaciones: statsGlobal.total,
    totalInfancia: statsGlobal.infancia,
    tasaGlobal: +(tasaGlobal * 100).toFixed(1),
    diasPromedio: diasInf.length ? +promedioGlobalDias.toFixed(1) : null,
    diasMediana: diasInf.length ? mediana(diasInf) : null,
    pctDentro7: +(pctDentro7 * 100).toFixed(1),
    mismoPct: mismoTotal ? +(mismoPctGlobal * 100).toFixed(1) : null,
    mismoTotal,
    agencias: agenciaEntries.map((e) => ({ agencia: e.label, total: e.total, infancia: e.infancia, tasa: +(e.tasa * 100).toFixed(1) })),
    productos: productoEntries.map((e) => ({ producto: e.label, total: e.total, infancia: e.infancia, tasa: +(e.tasa * 100).toFixed(1) })),
    causas: causasDist.map((c) => ({ causa: c.label, count: c.count, pct: +c.pct.toFixed(1) })),
    claves: clavesDist.map((c) => ({ clave: c.label, count: c.count, pct: +c.pct.toFixed(1) })),
    diasBuckets: buckets.map((b) => ({ label: b.label, count: b.count, pct: +b.pct.toFixed(1) })),
    mismoPorCausa: Object.entries(mismoPorCausa).map(([causa, s]) => ({ causa, total: s.total, mismo: s.mismo, tasa: s.total ? +((s.mismo / s.total) * 100).toFixed(1) : 0 })),
    tecnicos: tecnicoEntries.map((e) => ({
      tecnico: e.label, agencia: (metaTec[e.label] || {}).agencia || 'sin informacion',
      total: e.total, infancia: e.infancia, tasa: +(e.tasa * 100).toFixed(1),
    })),
    highlights: { top: highlightsTop, bottom: highlightsBottom },
    conclusiones,
  };

  const html = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Informe de Averias de Infancia · COBRA</title>
<meta name="theme-color" content="#003c71">
<link rel="manifest" href="manifest.json">
<link rel="apple-touch-icon" href="icon-192.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="Supervisor">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"><\/script>
<style>
  :root{
    --bg:#eef1f4;
    --panel:#ffffff;
    --panel-2:#f5f7f9;
    --border:#e0e5ea;
    --text:#22303f;
    --text-dim:#6b7a8c;
    --cobra-navy:#003c71;
    --cobra-blue:#0071ce;
    --celeste:#29a9e0;
    --celeste-soft:#e8f6fd;
    --promotor:#1fa971;
    --neutro:#e2962e;
    --detractor:#e2523e;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;
    font-family:'Segoe UI', Arial, sans-serif;
    background:var(--bg);
    color:var(--text);
    -webkit-font-smoothing:antialiased;
  }
  header.hero{
    background:linear-gradient(120deg,#ffffff 0%,var(--celeste-soft) 55%,#dcf1fb 100%);
    padding:34px 6vw 40px;
    position:relative;
    overflow:hidden;
    border-bottom:4px solid var(--celeste);
  }
  header.hero::after{
    content:"";
    position:absolute; right:-100px; top:-100px;
    width:340px; height:340px; border-radius:50%;
    background:radial-gradient(circle, rgba(41,169,224,0.18), transparent 70%);
  }
  .brand-row{ display:flex; align-items:center; gap:18px; margin-bottom:22px; }
  .brand-row img{ height:46px; }
  .brand-divider{ width:1px; height:34px; background:var(--border); }
  .eyebrow{ text-transform:uppercase; letter-spacing:.14em; font-size:12.5px; color:var(--celeste); font-weight:800; }
  h1{ margin:0 0 6px; font-size:clamp(26px,4vw,38px); font-weight:800; letter-spacing:-0.01em; color:var(--cobra-navy); }
  .subtitle{ color:#3a4a5c; font-size:15px; max-width:660px; line-height:1.55; }
  .meta-row{ display:flex; gap:22px; flex-wrap:wrap; margin-top:22px; font-size:13px; color:#3a4a5c; }
  .meta-row span b{color:var(--cobra-navy);}
  .back-link{ display:inline-flex; align-items:center; gap:6px; font-size:12.5px; font-weight:700; color:var(--cobra-navy); text-decoration:none; margin-bottom:14px; }
  .back-link:hover{ text-decoration:underline; }
  .archive-row{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin-top:16px; }
  .archive-row .archive-label{ font-size:11.5px; text-transform:uppercase; letter-spacing:.06em; font-weight:700; color:var(--text-dim); margin-right:2px; }
  .archive-pill{ display:inline-block; font-size:12.5px; font-weight:700; padding:4px 12px; border-radius:20px; text-decoration:none; border:1px solid var(--border); color:var(--cobra-navy); background:#fff; transition:background .15s ease; }
  .archive-pill:hover{ background:var(--celeste-soft); }
  .archive-pill.current{ background:var(--cobra-navy); color:#fff; border-color:var(--cobra-navy); cursor:default; }
  main{ padding:36px 6vw 80px; max-width:1280px; margin:0 auto; }

  .kpi-grid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:16px; margin:-58px 0 34px; position:relative; z-index:2; }
  .kpi-card{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:20px 20px 18px; box-shadow:0 10px 24px rgba(20,50,80,.08); }
  .kpi-card .label{ font-size:11.5px; text-transform:uppercase; letter-spacing:.08em; color:var(--text-dim); font-weight:700; margin-bottom:8px; }
  .kpi-card .value{ font-size:30px; font-weight:800; line-height:1; color:var(--cobra-navy); }
  .kpi-card .value.hl{ color:var(--celeste); }
  .kpi-card .sub{ font-size:12px; color:var(--text-dim); margin-top:6px; }

  section{ margin-bottom:44px; }
  .section-title{ display:flex; align-items:baseline; gap:10px; margin-bottom:6px; }
  .section-title .num{ font-size:13px; font-weight:800; color:var(--celeste); background:rgba(41,169,224,0.12); border:1px solid rgba(41,169,224,.35); border-radius:6px; padding:2px 8px; }
  .section-title h2{ margin:0; font-size:20px; font-weight:750; color:var(--cobra-navy); }
  .section-desc{ color:var(--text-dim); font-size:13.5px; margin:0 0 18px; max-width:820px; line-height:1.55;}

  .panel{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:22px 24px 14px; box-shadow:0 4px 14px rgba(20,50,80,.05); }
  .grid-2{ display:grid; grid-template-columns:1.35fr .95fr; gap:18px; }
  @media (max-width:880px){ .grid-2{grid-template-columns:1fr;} }

  table{ width:100%; border-collapse:collapse; font-size:13.5px; }
  th{ text-align:left; color:var(--text-dim); font-weight:700; font-size:11.5px; text-transform:uppercase; letter-spacing:.05em; padding:10px 10px; border-bottom:1px solid var(--border); }
  td{ padding:10px; border-bottom:1px solid var(--panel-2); }
  tr:hover td{ background:var(--panel-2); }
  .badge{ display:inline-block; padding:2px 9px; border-radius:20px; font-size:11.5px; font-weight:700; }
  .badge.hi{ background:rgba(31,169,113,.12); color:var(--promotor); }
  .badge.mid{ background:rgba(226,150,46,.14); color:var(--neutro); }
  .badge.lo{ background:rgba(226,82,62,.12); color:var(--detractor); }

  .callout{ border-left:3px solid var(--celeste); background:linear-gradient(90deg, var(--celeste-soft), transparent); padding:14px 18px; border-radius:0 10px 10px 0; font-size:13.5px; color:#2c3e50; line-height:1.6; }
  .callout.warn{ border-left-color:var(--neutro); background:linear-gradient(90deg, rgba(226,150,46,.10), transparent); }

  .rec-list{ display:grid; gap:10px; margin-top:8px; }
  .rec-item{ display:flex; gap:12px; align-items:flex-start; background:var(--panel-2); border:1px solid var(--border); border-radius:10px; padding:14px 16px; }
  .rec-item .idx{ flex:none; width:26px; height:26px; border-radius:50%; background:var(--celeste); color:#fff; font-weight:800; font-size:12.5px; display:flex; align-items:center; justify-content:center; }
  .rec-item p{ margin:0; font-size:13.5px; color:#334252; line-height:1.55; }
  .rec-item strong{ color:var(--cobra-navy); }

  .perf-grid{ display:grid; grid-template-columns:1fr 1fr; gap:18px; }
  @media (max-width:880px){ .perf-grid{grid-template-columns:1fr;} }
  .perf-col-title{ font-size:13px; font-weight:800; text-transform:uppercase; letter-spacing:.06em; margin-bottom:12px; display:flex; align-items:center; gap:8px; }
  .perf-col-title.top{ color:var(--promotor); }
  .perf-col-title.bottom{ color:var(--detractor); }
  .tech-card{ background:var(--panel); border:1px solid var(--border); border-radius:12px; padding:16px 18px; margin-bottom:12px; box-shadow:0 3px 10px rgba(20,50,80,.04); }
  .tech-card.bottom-card{ border-left:4px solid var(--detractor); }
  .tech-card.top-card{ border-left:4px solid var(--promotor); }
  .tech-head{ display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; }
  .tech-name{ font-weight:750; font-size:14px; color:var(--cobra-navy); }
  .tech-nps{ font-weight:800; font-size:15px; }
  .tech-nps.lo{ color:var(--detractor); }
  .tech-nps.hi{ color:var(--promotor); }
  .tech-meta{ font-size:12px; color:var(--text-dim); margin-bottom:8px; }
  .tech-quote{ font-size:12.5px; color:#4a5a6b; font-style:italic; margin-top:8px; background:var(--panel-2); border-radius:8px; padding:8px 10px; border-left:2px solid var(--border); }

  footer{ text-align:center; padding:26px; color:var(--text-dim); font-size:12px; border-top:1px solid var(--border); }
  canvas{ max-width:100%; }

  .export-pdf-btn{
    position:fixed; top:18px; right:18px; z-index:50;
    display:flex; align-items:center; gap:8px;
    background:var(--cobra-navy); color:#fff; border:none;
    padding:10px 16px; border-radius:24px; font-size:13.5px; font-weight:700;
    font-family:inherit; cursor:pointer;
    box-shadow:0 6px 16px rgba(0,60,113,.28);
    transition:background .15s ease, transform .15s ease;
  }
  .export-pdf-btn:hover{ background:var(--cobra-blue); transform:translateY(-1px); }
  .export-pdf-btn svg{ width:15px; height:15px; flex:none; }
  @media (max-width:640px){ .export-pdf-btn span{ display:none; } .export-pdf-btn{ padding:10px; } }

  @media print{
    .no-print{ display:none !important; }
    body{ background:#fff; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    header.hero{ background:#fff; border-bottom:2px solid var(--celeste); }
    header.hero::after{ display:none; }
    .kpi-grid{ margin-top:14px; }
    .panel, .kpi-card, .tech-card, .rec-item{ box-shadow:none; border:1px solid var(--border); }
    section{ break-inside:avoid-page; }
    .kpi-card, .tech-card, .rec-item{ break-inside:avoid; }
    .grid-2, .perf-grid{ break-inside:avoid; }
  }
</style>
</head>
<body>

<button class="export-pdf-btn no-print" onclick="window.print()" type="button">
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V2h9l3 3v4"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><path d="M6 14h12v8H6z"/></svg>
  <span>Exportar a PDF</span>
</button>

<header class="hero">
  <a class="back-link no-print" href="index.html">&larr; Volver al indice de informes</a>
  <div class="brand-row">
    <img src="logo-cobra.png" alt="Cobra">
    <div class="brand-divider"></div>
    <div class="eyebrow">Calidad &amp; Capacitacion · Analisis de Averias de Infancia</div>
  </div>
  <h1>Informe de Averias de Infancia</h1>
  <div class="subtitle">Instalaciones de Fibra Optica que generaron una reparacion dentro de su periodo de infancia (los dias siguientes a la instalacion), a partir de los registros de COBRA. Incluye tasa de infancia por agencia, producto y tecnico instalador, distribucion de causas y claves de cierre de la reparacion, tiempos hasta la falla, y si la resolvio el mismo tecnico que instalo.</div>
  <div class="meta-row" id="metaRow"></div>
  <div class="archive-row no-print" id="archiveRow"></div>
</header>

<main>
  <div class="kpi-grid" id="kpiGrid"></div>

  <section>
    <div class="section-title"><span class="num">01</span><h2>Tasa de Infancia por Agencia</h2></div>
    <p class="section-desc">Porcentaje de instalaciones que generaron una reparacion dentro de su periodo de infancia, comparado contra la meta institucional.</p>
    <div class="grid-2">
      <div class="panel" style="height:300px;"><canvas id="chartAgencia"></canvas></div>
      <div class="panel">
        <div style="font-size:13px;color:var(--text-dim);font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;">Detalle por agencia</div>
        <table id="tablaAgencia"></table>
        <div class="callout" style="margin-top:16px;" id="agenciaCallout"></div>
      </div>
    </div>
  </section>

  <section>
    <div class="section-title"><span class="num">02</span><h2>Tasa de Infancia por Producto</h2></div>
    <p class="section-desc">Que producto instalado reincide con mas frecuencia dentro del periodo de infancia.</p>
    <div class="grid-2">
      <div class="panel" style="height:280px;"><canvas id="chartProducto"></canvas></div>
      <div class="panel">
        <div style="font-size:13px;color:var(--text-dim);font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;">Detalle por producto</div>
        <table id="tablaProducto"></table>
        <div class="callout" style="margin-top:16px;" id="productoCallout"></div>
      </div>
    </div>
  </section>

  <section>
    <div class="section-title"><span class="num">03</span><h2>Causas y Claves de Cierre de la Reparacion de Infancia</h2></div>
    <p class="section-desc">Distribucion dentro del grupo de averias de infancia (no es una tasa contra el total de instalaciones).</p>
    <div class="grid-2">
      <div class="panel" style="height:320px;"><canvas id="chartCausa"></canvas></div>
      <div class="panel">
        <div style="font-size:13px;color:var(--text-dim);font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;">Clave de cierre</div>
        <table id="tablaClave"></table>
      </div>
    </div>
  </section>

  <section>
    <div class="section-title"><span class="num">04</span><h2>Tiempo hasta la Infancia</h2></div>
    <p class="section-desc">Cuantos dias pasan entre la instalacion y la reparacion de infancia. Mientras mas cerca de 0 dias, mas probable que la instalacion original tuviera un defecto de origen.</p>
    <div class="grid-2">
      <div class="panel" style="height:300px;"><canvas id="chartDias"></canvas></div>
      <div class="panel">
        <div style="font-size:13px;color:var(--text-dim);font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;">Resumen</div>
        <table id="tablaDias"></table>
        <div class="callout" style="margin-top:16px;" id="diasCallout"></div>
      </div>
    </div>
  </section>

  <section>
    <div class="section-title"><span class="num">05</span><h2>Reparacion Atendida por el Mismo Tecnico que Instalo</h2></div>
    <p class="section-desc">De las averias de infancia, si volvio el mismo tecnico que instalo o si tuvo que intervenir otro.</p>
    <div class="grid-2">
      <div class="panel" style="height:300px;"><canvas id="chartMismo"></canvas></div>
      <div class="panel">
        <div style="font-size:13px;color:var(--text-dim);font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;">Por causa</div>
        <table id="tablaMismo"></table>
        <div class="callout" style="margin-top:16px;" id="mismoCallout"></div>
      </div>
    </div>
  </section>

  <section>
    <div class="section-title"><span class="num">06</span><h2>Ranking de Tecnicos Instaladores por Tasa de Infancia (min. 10 instalaciones)</h2></div>
    <p class="section-desc">Ordenado de mayor a menor tasa dentro de cada agencia. Base para reconocimiento y para focalizar coaching/capacitacion.</p>
    <div id="tecnicosPorAgencia"></div>
  </section>

  <section>
    <div class="section-title"><span class="num">07</span><h2>Mejor y Menor Desempeno</h2></div>
    <p class="section-desc">Tecnicos instaladores con minimo 10 instalaciones, ordenados por su tasa de infancia.</p>
    <div class="perf-grid">
      <div>
        <div class="perf-col-title top">🟢 Mejor desempeno</div>
        <div id="topTecnicos"></div>
      </div>
      <div>
        <div class="perf-col-title bottom">🔴 Menor desempeno</div>
        <div id="bottomTecnicos"></div>
      </div>
    </div>
  </section>

  <section>
    <div class="section-title"><span class="num">08</span><h2>Conclusiones y Recomendaciones</h2></div>
    <div class="rec-list" id="recList"></div>
  </section>
</main>

<footer id="footerText"></footer>

<script>
const DATA = ${JSON.stringify(DATA)};

function npsClass(v, meta){ return v<=meta ? 'hi' : (v<=meta*2 ? 'mid' : 'lo'); }
function titleCase(s){ return s.split(' ').map(w=>w?w[0]+w.slice(1).toLowerCase():w).join(' '); }

document.getElementById('metaRow').innerHTML = \`
  <span>📅 Periodo analizado: <b>\${DATA.periodo}</b></span>
  <span>📍 Agencias: <b>\${DATA.agencias_lista}</b></span>
  <span>🧾 Base: <b>\${DATA.totalInstalaciones} instalaciones</b> · \${DATA.totalInfancia} averias de infancia</span>
  <span>🧮 Formula: <b>Averias de Infancia / Total de instalaciones</b></span>
\`;

// ---- Historial de meses ----
function renderArchiveRow(lista) {
  if (!lista || lista.length < 2) return;
  const pills = lista.map(a =>
    a.slug === DATA.periodoSlug
      ? \`<span class="archive-pill current">\${a.label}</span>\`
      : \`<a class="archive-pill" href="\${a.url}">\${a.label}</a>\`
  ).join('');
  document.getElementById('archiveRow').innerHTML = \`<span class="archive-label">📁 Meses disponibles:</span>\${pills}\`;
}
renderArchiveRow(DATA.archivos);
fetch('meses.json', { cache: 'no-store' })
  .then(r => r.ok ? r.json() : null)
  .then(data => { if (data && data.infancia) renderArchiveRow(data.infancia); })
  .catch(() => {});

const fontColor = '#6b7a8c';
Chart.defaults.color = fontColor;
Chart.defaults.borderColor = 'rgba(20,50,80,0.06)';
Chart.defaults.font.family = "'Segoe UI', Arial, sans-serif";

// Plugin liviano: dibuja el valor al final de cada barra (sin depender de librerias externas).
// El formatter va como closure (no dentro de chart.options): Chart.js "resuelve" cualquier
// funcion guardada dentro de options.plugins.<id> llamandola con un objeto de contexto interno
// (para soportar "opciones dinamicas"), lo que rompe un formatter normal si se guarda ahi.
function crearEtiquetasPlugin(formatter, datasetIndex) {
  return {
    id: 'valueLabels',
    afterDatasetsDraw(chart) {
      const idx = datasetIndex || 0;
      const meta = chart.getDatasetMeta(idx);
      if (!meta || meta.hidden) return;
      const ctx = chart.ctx;
      ctx.save();
      ctx.font = '600 11px "Segoe UI", Arial, sans-serif';
      ctx.fillStyle = '#22303f';
      meta.data.forEach((bar, i) => {
        const raw = chart.data.datasets[idx].data[i];
        if (raw === null || raw === undefined) return;
        const label = String(formatter(raw));
        if (chart.options.indexAxis === 'y') {
          ctx.textAlign = 'left';
          ctx.textBaseline = 'middle';
          ctx.fillText(label, bar.x + 6, bar.y);
        } else {
          ctx.textAlign = 'center';
          ctx.textBaseline = 'bottom';
          ctx.fillText(label, bar.x, bar.y - 4);
        }
      });
      ctx.restore();
    }
  };
}

const metaGap = +(DATA.tasaGlobal - DATA.meta).toFixed(1);
const metaSub = metaGap <= 0
  ? \`<span style="color:var(--promotor)">▼ \${Math.abs(metaGap)} pts bajo la meta (\${DATA.meta}%) — cumple</span>\`
  : \`<span style="color:var(--detractor)">▲ \${metaGap} pts sobre la meta (\${DATA.meta}%) — no cumple</span>\`;
document.getElementById('kpiGrid').innerHTML = \`
  <div class="kpi-card"><div class="label">Tasa Global de Infancia</div><div class="value hl">\${DATA.tasaGlobal}%</div><div class="sub">\${metaSub}</div></div>
  <div class="kpi-card"><div class="label">Total Instalaciones</div><div class="value">\${DATA.totalInstalaciones}</div><div class="sub">Periodo \${DATA.periodo}</div></div>
  <div class="kpi-card"><div class="label">Averias de Infancia</div><div class="value">\${DATA.totalInfancia}</div><div class="sub">Instalaciones que fallaron temprano</div></div>
  <div class="kpi-card"><div class="label">Dias Prom. hasta la Infancia</div><div class="value">\${DATA.diasPromedio ?? '-'}</div><div class="sub">Mediana: \${DATA.diasMediana ?? '-'} dias · \${DATA.pctDentro7}% dentro de 7 dias</div></div>
  <div class="kpi-card"><div class="label">% Mismo Tecnico</div><div class="value">\${DATA.mismoPct ?? '-'}%</div><div class="sub">\${DATA.mismoTotal} casos</div></div>
\`;

new Chart(document.getElementById('chartAgencia'), {
  type: 'bar',
  plugins: [crearEtiquetasPlugin((v)=>v+'%')],
  data: {
    labels: DATA.agencias.map(a=>a.agencia),
    datasets: [
      { label:'Tasa de infancia (%)', data: DATA.agencias.map(a=>a.tasa), backgroundColor: DATA.agencias.map(a=> a.tasa<=DATA.meta ? 'rgba(31,169,113,0.85)' : 'rgba(226,82,62,0.85)'), borderRadius:6, maxBarThickness:60 },
      { label:'Meta ('+DATA.meta+'%)', data: DATA.agencias.map(()=>DATA.meta), type:'line', borderColor:'#e2523e', borderDash:[5,4], borderWidth:1.5, pointRadius:0, tension:0 }
    ]
  },
  options: { responsive:true, maintainAspectRatio:false, layout:{ padding:{ top:20 } }, plugins:{ legend:{ position:'top', labels:{boxWidth:10} } }, scales:{ y:{ min:0, title:{display:true,text:'Tasa %'}, grid:{color:'rgba(20,50,80,0.06)'} } } }
});
let rowsAg = '<tr><th>Agencia</th><th>Tasa</th><th>vs. Meta</th><th>Instalaciones</th></tr>';
DATA.agencias.slice().sort((a,b)=>a.tasa-b.tasa).forEach(a=>{
  const gap = +(a.tasa - DATA.meta).toFixed(1);
  const gapTxt = gap<=0 ? \`<span style="color:var(--promotor)">\${gap}</span>\` : \`<span style="color:var(--detractor)">+\${gap}</span>\`;
  rowsAg += \`<tr><td>\${a.agencia}</td><td><span class="badge \${npsClass(a.tasa,DATA.meta)}">\${a.tasa}%</span></td><td>\${gapTxt}</td><td>\${a.total}</td></tr>\`;
});
document.getElementById('tablaAgencia').innerHTML = rowsAg;
const agSobreMeta = DATA.agencias.filter(a=>a.tasa > DATA.meta);
document.getElementById('agenciaCallout').innerHTML = agSobreMeta.length
  ? \`<b>Sobre la meta:</b> \${agSobreMeta.map(a=>\`\${a.agencia} (\${a.tasa}%)\`).join(' y ')} supera el \${DATA.meta}% institucional.\`
  : \`<b>Todas las agencias dentro de la meta</b> de \${DATA.meta}% este periodo.\`;

new Chart(document.getElementById('chartProducto'), {
  type: 'bar',
  plugins: [crearEtiquetasPlugin((v)=>v+'%')],
  data: {
    labels: DATA.productos.map(p=>p.producto),
    datasets: [
      { label:'Tasa de infancia (%)', data: DATA.productos.map(p=>p.tasa), backgroundColor: DATA.productos.map(p=> p.tasa<=DATA.meta ? 'rgba(31,169,113,0.85)' : 'rgba(226,82,62,0.85)'), borderRadius:6, maxBarThickness:50 },
      { label:'Meta ('+DATA.meta+'%)', data: DATA.productos.map(()=>DATA.meta), type:'line', borderColor:'#e2523e', borderDash:[5,4], borderWidth:1.5, pointRadius:0, tension:0 }
    ]
  },
  options: { responsive:true, maintainAspectRatio:false, layout:{ padding:{ top:20 } }, plugins:{ legend:{ position:'top', labels:{boxWidth:10} } }, scales:{ y:{ min:0, title:{display:true,text:'Tasa %'}, grid:{color:'rgba(20,50,80,0.06)'} } } }
});
let rowsPr = '<tr><th>Producto</th><th>Tasa</th><th>Infancia</th><th>Total</th></tr>';
DATA.productos.slice().sort((a,b)=>b.tasa-a.tasa).forEach(p=>{
  rowsPr += \`<tr><td>\${p.producto}</td><td><span class="badge \${npsClass(p.tasa,DATA.meta)}">\${p.tasa}%</span></td><td>\${p.infancia}</td><td>\${p.total}</td></tr>\`;
});
document.getElementById('tablaProducto').innerHTML = rowsPr;
const peorPr = DATA.productos.slice().sort((a,b)=>b.tasa-a.tasa)[0];
document.getElementById('productoCallout').innerHTML = peorPr ? \`<b>Foco principal:</b> \${peorPr.producto} tiene la mayor tasa de infancia (\${peorPr.tasa}%).\` : '';

new Chart(document.getElementById('chartCausa'), {
  type: 'bar',
  plugins: [crearEtiquetasPlugin((v)=>v+'%')],
  data: {
    labels: DATA.causas.map(c=>c.causa),
    datasets: [{ label:'% de averias de infancia', data: DATA.causas.map(c=>c.pct), backgroundColor:'rgba(0,113,206,0.85)', borderRadius:5, maxBarThickness:26 }]
  },
  options: { indexAxis:'y', responsive:true, maintainAspectRatio:false, layout:{ padding:{ right:40 } }, plugins:{ legend:{ display:false } }, scales:{ x:{ min:0, title:{display:true,text:'% de casos'}, grid:{color:'rgba(20,50,80,0.06)'} } } }
});
let rowsCl = '<tr><th>Clave</th><th>Casos</th><th>%</th></tr>';
DATA.claves.forEach(c=>{ rowsCl += \`<tr><td>\${c.clave}</td><td>\${c.count}</td><td>\${c.pct}%</td></tr>\`; });
document.getElementById('tablaClave').innerHTML = rowsCl;

new Chart(document.getElementById('chartDias'), {
  type: 'bar',
  plugins: [crearEtiquetasPlugin((v)=>v)],
  data: { labels: DATA.diasBuckets.map(b=>b.label), datasets: [{ label:'Casos', data: DATA.diasBuckets.map(b=>b.count), backgroundColor:'rgba(0,113,206,0.85)', borderRadius:6, maxBarThickness:60 }] },
  options: { responsive:true, maintainAspectRatio:false, layout:{ padding:{ top:20 } }, plugins:{ legend:{ display:false } }, scales:{ y:{ min:0, title:{display:true,text:'N° de casos'}, grid:{color:'rgba(20,50,80,0.06)'} } } }
});
document.getElementById('tablaDias').innerHTML = \`
  <tr><th>Indicador</th><th>Valor</th></tr>
  <tr><td>Promedio</td><td>\${DATA.diasPromedio ?? '-'} dias</td></tr>
  <tr><td>Mediana</td><td>\${DATA.diasMediana ?? '-'} dias</td></tr>
  <tr><td>Dentro de 7 dias</td><td><span class="badge \${DATA.pctDentro7>=60?'lo':'mid'}">\${DATA.pctDentro7}%</span></td></tr>
\`;
document.getElementById('diasCallout').innerHTML =
  \`<b>\${DATA.pctDentro7}%</b> de las averias de infancia ocurre dentro de la primera semana. Cuanto mas alto este numero, mas senal de un defecto de origen en la instalacion (no de una falla nueva independiente).\`;

new Chart(document.getElementById('chartMismo'), {
  type: 'doughnut',
  data: { labels: ['Mismo tecnico', 'Tecnico distinto'], datasets: [{ data: [DATA.mismoPct, +(100-DATA.mismoPct).toFixed(1)], backgroundColor: ['#e2523e','#1fa971'], borderWidth:0 }] },
  options: { responsive:true, maintainAspectRatio:false, plugins:{ legend:{ position:'bottom', labels:{boxWidth:10} } }, cutout: '62%' }
});
let rowsMi = '<tr><th>Causa</th><th>% Mismo Tecnico</th><th>Casos</th></tr>';
DATA.mismoPorCausa.forEach(m=>{ rowsMi += \`<tr><td>\${m.causa}</td><td><span class="badge \${m.tasa>=50?'lo':'hi'}">\${m.tasa}%</span></td><td>\${m.mismo}/\${m.total}</td></tr>\`; });
document.getElementById('tablaMismo').innerHTML = rowsMi;
document.getElementById('mismoCallout').innerHTML =
  \`<b>\${DATA.mismoPct}%</b> de las averias de infancia fueron atendidas por el mismo tecnico que hizo la instalacion original.\`;

const agenciasOrden = DATA.agencias.slice().sort((a,b)=>b.total-a.total).map(a=>a.agencia);
let tecnicosHtml = '';
agenciasOrden.forEach(agencia => {
  const tecsAg = DATA.tecnicos.filter(t=>t.agencia===agencia && t.total>=10).sort((a,b)=>b.tasa-a.tasa);
  if (!tecsAg.length) return;
  let rowsTec = '<tr><th>#</th><th>Tecnico</th><th>Tasa</th><th>Instalaciones</th><th>Infancia</th></tr>';
  tecsAg.forEach((t,i)=>{
    rowsTec += \`<tr><td>\${i+1}</td><td>\${titleCase(t.tecnico)}</td><td><span class="badge \${npsClass(t.tasa,DATA.meta)}">\${t.tasa}%</span></td><td>\${t.total}</td><td>\${t.infancia}</td></tr>\`;
  });
  tecnicosHtml += \`<div class="panel" style="margin-bottom:16px;">
    <div style="font-size:13px;color:var(--text-dim);font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;">\${agencia} · \${tecsAg.length} tecnico\${tecsAg.length===1?'':'s'} (min. 10 instalaciones)</div>
    <table>\${rowsTec}</table>
  </div>\`;
});
document.getElementById('tecnicosPorAgencia').innerHTML = tecnicosHtml;

function techCard(t, kind){
  const cls = kind==='top' ? 'hi' : 'lo';
  return \`<div class="tech-card \${kind==='top'?'top-card':'bottom-card'}">
    <div class="tech-head"><span class="tech-name">\${titleCase(t.tecnico)}</span><span class="tech-nps \${cls}">\${t.tasa}%</span></div>
    <div class="tech-meta">\${t.total} instalaciones · \${t.infancia} averia\${t.infancia===1?'':'s'} de infancia</div>
    <div class="tech-quote">\${t.nota}</div>
  </div>\`;
}
document.getElementById('topTecnicos').innerHTML = DATA.highlights.top.map(t=>techCard(t,'top')).join('');
document.getElementById('bottomTecnicos').innerHTML = DATA.highlights.bottom.map(t=>techCard(t,'bottom')).join('');

document.getElementById('recList').innerHTML = DATA.conclusiones.map((c,i)=>\`
  <div class="rec-item"><div class="idx">\${i+1}</div><p>\${c}</p></div>
\`).join('');

document.getElementById('footerText').innerHTML =
  \`Informe generado a partir de \${DATA.archivoOrigen} · \${DATA.agencias_lista} · Generado \${DATA.generadoEl}\`;
<\/script>

<script>
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(()=>{}));
}
<\/script>

</body>
</html>`;

  const htmlPath = path.join(carpeta, 'Dashboard_Infancia.html');
  fs.writeFileSync(htmlPath, html, 'utf8');
  const htmlArchivoPath = path.join(carpeta, `Dashboard_Infancia_${periodoSlug}.html`);
  fs.writeFileSync(htmlArchivoPath, html, 'utf8');
  console.log('Dashboard HTML generado:', htmlPath, 'y', htmlArchivoPath);
}

main().catch((err) => {
  console.error('ERROR:', err.message);
  process.exitCode = 1;
});
