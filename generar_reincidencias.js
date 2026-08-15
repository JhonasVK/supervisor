// Analisis de reparaciones reiteradas (reincidencias dentro de 30 dias)
// Lee el CSV "p23-averias-reiteradas..._COBRA_...csv" desde la carpeta bbdd\
// (subcarpeta de esta misma carpeta) y genera un Excel con el detalle:
// 1ra reparacion (tecnico, clave de cierre, causa) vs 2da reparacion / reitero
// (tecnico, clave de cierre, causa) y los dias transcurridos entre ambas.
//
// Para actualizar el analisis: sobreescribe el CSV dentro de la carpeta bbdd\
// con el nuevo mes y vuelve a ejecutar "Generar_Reporte_Reincidencias.bat".

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const carpeta = __dirname;
const carpetaBbdd = path.join(carpeta, 'bbdd');

function encontrarCsvOrigen() {
  if (!fs.existsSync(carpetaBbdd)) {
    throw new Error('No existe la carpeta "bbdd" en ' + carpeta);
  }
  const candidatos = fs
    .readdirSync(carpetaBbdd)
    .filter((f) => /^p23-averias-reiteradas.*COBRA.*\.csv$/i.test(f))
    .map((f) => ({ nombre: f, mtime: fs.statSync(path.join(carpetaBbdd, f)).mtimeMs }))
    .sort((a, b) => b.mtime - a.mtime);

  if (candidatos.length === 0) {
    throw new Error(
      'No se encontro ningun archivo "p23-averias-reiteradas..._COBRA_....csv" en ' + carpetaBbdd
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

function clave(r) {
  const stb = (r['toa_piv_clave_stb'] || '').trim();
  const tv = (r['toa_piv_clave_tv'] || '').trim();
  const ba = (r['toa_piv_clave_ba'] || '').trim();
  const parts = [];
  if (ba) parts.push('BA:' + ba);
  if (tv) parts.push('TV:' + tv);
  if (stb) parts.push('STB:' + stb);
  return parts.join(' / ');
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


async function main() {
  const csvPath = encontrarCsvOrigen();
  console.log('Archivo origen:', csvPath);

  const rows = leerCsv(csvPath);
  const byFolio = {};
  rows.forEach((r) => { byFolio[r['toa_piv_folio_toa']] = r; });

  const reit = rows.filter((r) => (r['rdy_prd_tiene_reitero_30d'] || '').trim() === '1');
  reit.sort((a, b) => {
    const ta = a['toa_piv_nombre_tecnico'] || '';
    const tb = b['toa_piv_nombre_tecnico'] || '';
    if (ta !== tb) return ta.localeCompare(tb);
    return (a['toa_piv_fecha_ingreso'] || '').localeCompare(b['toa_piv_fecha_ingreso'] || '');
  });

  console.log('Total registros:', rows.length, '| Reincidencias (30d):', reit.length);

  const columns = [
    { header: 'Folio 1ra Reparacion', key: 'folio1', width: 20 },
    { header: 'Fecha 1ra Reparacion', key: 'fecha1', width: 20 },
    { header: 'Tecnico 1', key: 'tec1', width: 32 },
    { header: 'RUT Tecnico 1', key: 'rut1', width: 14 },
    { header: 'Agencia 1', key: 'agencia1', width: 14 },
    { header: 'Clave de Cierre 1', key: 'clave1', width: 16 },
    { header: 'Causa 1', key: 'causa1', width: 22 },
    { header: 'Subcausa 1', key: 'subcausa1', width: 42 },
    { header: 'Folio 2da Reparacion (Reitero)', key: 'folio2', width: 22 },
    { header: 'Fecha 2da Reparacion', key: 'fecha2', width: 20 },
    { header: 'Tecnico 2', key: 'tec2', width: 32 },
    { header: 'RUT Tecnico 2', key: 'rut2', width: 14 },
    { header: 'Clave de Cierre 2', key: 'clave2', width: 16 },
    { header: 'Causa 2', key: 'causa2', width: 22 },
    { header: 'Subcausa 2', key: 'subcausa2', width: 42 },
    { header: 'Mismo Tecnico?', key: 'mismo', width: 14 },
    { header: 'Dias hasta Reiteracion', key: 'dias', width: 18 },
  ];

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Analisis Averias Reiteradas';
  wb.created = new Date();

  const ws = wb.addWorksheet('Reincidencias Detalle', {
    views: [{ state: 'frozen', ySplit: 1 }],
  });
  ws.columns = columns;

  reit.forEach((r) => {
    const refFolio = r['rdy_prd_reiterado'];
    const r2 = byFolio[refFolio];
    const diasVal = r['rdy_prd_q_dias_reitero'];
    ws.addRow({
      folio1: r['toa_piv_folio_toa'],
      fecha1: r['toa_piv_fecha_ingreso'],
      tec1: r['toa_piv_nombre_tecnico'],
      rut1: r['toa_piv_rut_tecnico'],
      agencia1: r['toa_piv_agencia'],
      clave1: clave(r),
      causa1: r['toa_piv_causa'],
      subcausa1: r['toa_piv_subcausa'],
      folio2: refFolio,
      fecha2: r2 ? r2['toa_piv_fecha_ingreso'] : '(fuera del archivo)',
      tec2: r2 ? r2['toa_piv_nombre_tecnico'] : '(fuera del archivo)',
      rut2: r2 ? r2['toa_piv_rut_tecnico'] : '',
      clave2: r2 ? clave(r2) : '',
      causa2: r2 ? r2['toa_piv_causa'] : '',
      subcausa2: r2 ? r2['toa_piv_subcausa'] : '',
      mismo: r2 ? (r2['toa_piv_nombre_tecnico'] === r['toa_piv_nombre_tecnico'] ? 'SI' : 'NO') : 'S/D',
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
    const tec2Cell = row.getCell('tec2');
    const mismoCell = row.getCell('mismo');
    if (tec2Cell.value === '(fuera del archivo)') {
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      });
    } else if (mismoCell.value === 'NO') {
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4E4' } };
      });
    }
  });

  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: reit.length + 1, column: columns.length },
  };

  // ---------- Hoja de Estadistica: tasa de reincidencia (%) vs meta 4% ----------
  const META = 0.04; // 4%

  const reitSet = new Set(reit.map((r) => r['toa_piv_folio_toa']));

  function agruparPor(campo) {
    const grupos = {};
    rows.forEach((r) => {
      const key = (r[campo] || '').trim() || '(sin dato)';
      if (!grupos[key]) grupos[key] = { total: 0, reincidencias: 0 };
      grupos[key].total += 1;
      if (reitSet.has(r['toa_piv_folio_toa'])) grupos[key].reincidencias += 1;
    });
    return grupos;
  }

  // Agencia y RUT mas frecuente por tecnico (para mostrar en la tabla)
  function metaPorTecnico() {
    const info = {};
    rows.forEach((r) => {
      const tec = (r['toa_piv_nombre_tecnico'] || '').trim() || '(sin dato)';
      if (!info[tec]) info[tec] = { agencias: {}, ruts: {} };
      const ag = (r['toa_piv_agencia'] || '').trim() || '(sin dato)';
      const rut = (r['toa_piv_rut_tecnico'] || '').trim();
      info[tec].agencias[ag] = (info[tec].agencias[ag] || 0) + 1;
      if (rut) info[tec].ruts[rut] = (info[tec].ruts[rut] || 0) + 1;
    });
    const masFrecuente = (obj) =>
      Object.entries(obj).sort((a, b) => b[1] - a[1])[0]?.[0] || '';
    const resultado = {};
    Object.keys(info).forEach((tec) => {
      resultado[tec] = {
        agencia: masFrecuente(info[tec].agencias),
        rut: masFrecuente(info[tec].ruts),
      };
    });
    return resultado;
  }

  const statsGlobal = { total: rows.length, reincidencias: reit.length };
  const statsPorTecnico = agruparPor('toa_piv_nombre_tecnico');
  const statsPorAgencia = agruparPor('toa_piv_agencia');
  const statsPorCausa = agruparPor('toa_piv_causa');
  const statsPorSubcausa = agruparPor('toa_piv_subcausa');
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
    cell.fill = {
      type: 'pattern', pattern: 'solid',
      fgColor: { argb: tasa > META ? 'FFFCE4E4' : 'FFE2EFDA' },
    };
    cell.font.color = { argb: tasa > META ? 'FFC00000' : 'FF375623' };
  }

  function pintarCumple(cell, tasa) {
    const cumple = tasa <= META;
    cell.value = cumple ? 'CUMPLE' : 'NO CUMPLE';
    cell.font = { name: 'Arial', bold: true, color: { argb: cumple ? 'FF375623' : 'FFC00000' } };
    cell.fill = {
      type: 'pattern', pattern: 'solid',
      fgColor: { argb: cumple ? 'FFE2EFDA' : 'FFFCE4E4' },
    };
    cell.alignment = { horizontal: 'center' };
  }

  // --- Titulo ---
  est.mergeCells('A1:F1');
  estilizarTitulo(est.getCell('A1'), 'TASA DE REINCIDENCIA (%) — Meta: no superar el ' + (META * 100).toFixed(0) + '%');
  est.getRow(1).height = 24;

  // --- Global ---
  est.addRow([]);
  est.mergeCells('A3:F3');
  est.getCell('A3').value = 'GLOBAL';
  est.getCell('A3').font = { name: 'Arial', bold: true, size: 12 };

  const headerGlobal = est.addRow(['Total Reparaciones', 'Reincidencias', 'Tasa Global', 'Meta', 'Resultado', '']);
  estilizarHeader(headerGlobal);
  const rowGlobal = est.addRow([statsGlobal.total, statsGlobal.reincidencias, null, META, null, '']);
  rowGlobal.getCell(1).font = { name: 'Arial' };
  rowGlobal.getCell(2).font = { name: 'Arial' };
  pintarTasa(rowGlobal.getCell(3), statsGlobal.total ? statsGlobal.reincidencias / statsGlobal.total : 0);
  rowGlobal.getCell(4).numFmt = '0.0%';
  rowGlobal.getCell(4).font = { name: 'Arial' };
  pintarCumple(rowGlobal.getCell(5), statsGlobal.total ? statsGlobal.reincidencias / statsGlobal.total : 0);

  // --- Por Agencia ---
  est.addRow([]);
  est.addRow([]);
  const filaTituloAgencia = est.rowCount + 1;
  est.mergeCells(`A${filaTituloAgencia}:F${filaTituloAgencia}`);
  est.getCell(`A${filaTituloAgencia}`).value = 'POR AGENCIA';
  est.getCell(`A${filaTituloAgencia}`).font = { name: 'Arial', bold: true, size: 12 };

  const headerAgencia = est.addRow(['Agencia', 'Total Reparaciones', 'Reincidencias', 'Tasa', 'Meta', 'Resultado']);
  estilizarHeader(headerAgencia);
  Object.entries(statsPorAgencia)
    .sort((a, b) => (b[1].reincidencias / b[1].total) - (a[1].reincidencias / a[1].total))
    .forEach(([agencia, s]) => {
      const tasa = s.total ? s.reincidencias / s.total : 0;
      const row = est.addRow([agencia, s.total, s.reincidencias, null, META, null]);
      row.getCell(1).font = { name: 'Arial' };
      row.getCell(2).font = { name: 'Arial' };
      row.getCell(3).font = { name: 'Arial' };
      pintarTasa(row.getCell(4), tasa);
      row.getCell(5).numFmt = '0.0%';
      row.getCell(5).font = { name: 'Arial' };
      pintarCumple(row.getCell(6), tasa);
    });

  // --- Por Causa ---
  est.addRow([]);
  est.addRow([]);
  const filaTituloCausa = est.rowCount + 1;
  est.mergeCells(`A${filaTituloCausa}:F${filaTituloCausa}`);
  est.getCell(`A${filaTituloCausa}`).value = 'POR CAUSA';
  est.getCell(`A${filaTituloCausa}`).font = { name: 'Arial', bold: true, size: 12 };

  const headerCausa = est.addRow(['Causa', 'Total Reparaciones', 'Reincidencias', 'Tasa', 'Meta', 'Resultado']);
  estilizarHeader(headerCausa);
  Object.entries(statsPorCausa)
    .sort((a, b) => (b[1].reincidencias / b[1].total) - (a[1].reincidencias / a[1].total))
    .forEach(([causa, s]) => {
      const tasa = s.total ? s.reincidencias / s.total : 0;
      const row = est.addRow([causa, s.total, s.reincidencias, null, META, null]);
      row.getCell(1).font = { name: 'Arial' };
      row.getCell(2).font = { name: 'Arial' };
      row.getCell(3).font = { name: 'Arial' };
      pintarTasa(row.getCell(4), tasa);
      row.getCell(5).numFmt = '0.0%';
      row.getCell(5).font = { name: 'Arial' };
      pintarCumple(row.getCell(6), tasa);
    });

  // --- Por Subcausa ---
  est.addRow([]);
  est.addRow([]);
  const filaTituloSubcausa = est.rowCount + 1;
  est.mergeCells(`A${filaTituloSubcausa}:F${filaTituloSubcausa}`);
  est.getCell(`A${filaTituloSubcausa}`).value = 'POR SUBCAUSA';
  est.getCell(`A${filaTituloSubcausa}`).font = { name: 'Arial', bold: true, size: 12 };

  const headerSubcausa = est.addRow(['Subcausa', 'Total Reparaciones', 'Reincidencias', 'Tasa', 'Meta', 'Resultado']);
  estilizarHeader(headerSubcausa);
  Object.entries(statsPorSubcausa)
    .filter(([, s]) => s.total >= 5) // omite subcausas con muy pocos casos (tasa poco representativa)
    .sort((a, b) => (b[1].reincidencias / b[1].total) - (a[1].reincidencias / a[1].total))
    .forEach(([subcausa, s]) => {
      const tasa = s.total ? s.reincidencias / s.total : 0;
      const row = est.addRow([subcausa, s.total, s.reincidencias, null, META, null]);
      row.getCell(1).font = { name: 'Arial' };
      row.getCell(2).font = { name: 'Arial' };
      row.getCell(3).font = { name: 'Arial' };
      pintarTasa(row.getCell(4), tasa);
      row.getCell(5).numFmt = '0.0%';
      row.getCell(5).font = { name: 'Arial' };
      pintarCumple(row.getCell(6), tasa);
    });

  // --- Por Tecnico ---
  est.addRow([]);
  est.addRow([]);
  const filaTituloTecnico = est.rowCount + 1;
  est.mergeCells(`A${filaTituloTecnico}:F${filaTituloTecnico}`);
  est.getCell(`A${filaTituloTecnico}`).value = 'POR TECNICO';
  est.getCell(`A${filaTituloTecnico}`).font = { name: 'Arial', bold: true, size: 12 };

  const headerTecnico = est.addRow(['Tecnico', 'RUT', 'Agencia', 'Total Reparaciones', 'Reincidencias', 'Tasa', 'Meta', 'Resultado']);
  est.getColumn(7).width = 12;
  est.getColumn(8).width = 14;
  estilizarHeader(headerTecnico);

  Object.entries(statsPorTecnico)
    .sort((a, b) => (b[1].reincidencias / b[1].total) - (a[1].reincidencias / a[1].total))
    .forEach(([tecnico, s]) => {
      const tasa = s.total ? s.reincidencias / s.total : 0;
      const meta = metaTec[tecnico] || { agencia: '', rut: '' };
      const row = est.addRow([tecnico, meta.rut, meta.agencia, s.total, s.reincidencias, null, META, null]);
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

  // ---------- Hoja Tiempo hasta Reiteracion (dias) ----------
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

  const diasReit = reit
    .map((r) => Number(r['rdy_prd_q_dias_reitero']))
    .filter((n) => !Number.isNaN(n));

  function agruparDiasPor(campo) {
    const grupos = {};
    reit.forEach((r) => {
      const key = (r[campo] || '').trim() || '(sin dato)';
      const dias = Number(r['rdy_prd_q_dias_reitero']);
      if (Number.isNaN(dias)) return;
      if (!grupos[key]) grupos[key] = [];
      grupos[key].push(dias);
    });
    return grupos;
  }
  const diasPorTecnico = agruparDiasPor('toa_piv_nombre_tecnico');
  const diasPorCausa = agruparDiasPor('toa_piv_causa');

  const tr = wb.addWorksheet('Tiempo Reiteracion');
  tr.columns = [{ width: 34 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }];

  tr.mergeCells('A1:F1');
  estilizarTitulo(tr.getCell('A1'), 'TIEMPO HASTA LA REITERACION (dias entre 1ra y 2da reparacion)');
  tr.getRow(1).height = 24;

  // --- Global ---
  tr.addRow([]);
  tr.mergeCells('A3:F3');
  tr.getCell('A3').value = 'GLOBAL';
  tr.getCell('A3').font = { name: 'Arial', bold: true, size: 12 };

  const headerGlobalT = tr.addRow(['Casos con dato', 'Promedio (dias)', 'Mediana (dias)', 'Minimo', 'Maximo', '% dentro de 7 dias']);
  estilizarHeader(headerGlobalT);
  const dentro7 = diasReit.filter((d) => d <= 7).length;
  const rowGlobalT = tr.addRow([
    diasReit.length,
    diasReit.length ? Number(promedio(diasReit).toFixed(1)) : null,
    diasReit.length ? mediana(diasReit) : null,
    diasReit.length ? Math.min(...diasReit) : null,
    diasReit.length ? Math.max(...diasReit) : null,
    diasReit.length ? dentro7 / diasReit.length : null,
  ]);
  rowGlobalT.eachCell((cell) => { cell.font = { name: 'Arial' }; });
  rowGlobalT.getCell(6).numFmt = '0.0%';

  // --- Por Causa ---
  tr.addRow([]);
  tr.addRow([]);
  const filaTituloCausaT = tr.rowCount + 1;
  tr.mergeCells(`A${filaTituloCausaT}:F${filaTituloCausaT}`);
  tr.getCell(`A${filaTituloCausaT}`).value = 'POR CAUSA';
  tr.getCell(`A${filaTituloCausaT}`).font = { name: 'Arial', bold: true, size: 12 };

  const headerCausaT = tr.addRow(['Causa', 'Casos', 'Promedio (dias)', 'Mediana (dias)', 'Minimo', 'Maximo']);
  estilizarHeader(headerCausaT);
  Object.entries(diasPorCausa)
    .sort((a, b) => promedio(a[1]) - promedio(b[1]))
    .forEach(([causa, dias]) => {
      const row = tr.addRow([
        causa, dias.length, Number(promedio(dias).toFixed(1)), mediana(dias), Math.min(...dias), Math.max(...dias),
      ]);
      row.eachCell((cell) => { cell.font = { name: 'Arial' }; });
    });

  // --- Por Tecnico ---
  tr.addRow([]);
  tr.addRow([]);
  const filaTituloTecT = tr.rowCount + 1;
  tr.mergeCells(`A${filaTituloTecT}:F${filaTituloTecT}`);
  tr.getCell(`A${filaTituloTecT}`).value = 'POR TECNICO (ordenado de mas rapido a mas lento en reincidir)';
  tr.getCell(`A${filaTituloTecT}`).font = { name: 'Arial', bold: true, size: 12 };

  const headerTecT = tr.addRow(['Tecnico', 'Casos', 'Promedio (dias)', 'Mediana (dias)', 'Minimo', 'Maximo']);
  estilizarHeader(headerTecT);
  Object.entries(diasPorTecnico)
    .sort((a, b) => promedio(a[1]) - promedio(b[1]))
    .forEach(([tecnico, dias]) => {
      const row = tr.addRow([
        tecnico, dias.length, Number(promedio(dias).toFixed(1)), mediana(dias), Math.min(...dias), Math.max(...dias),
      ]);
      row.eachCell((cell) => { cell.font = { name: 'Arial' }; });
      const promCell = row.getCell(3);
      if (promedio(dias) <= 2) {
        promCell.font = { name: 'Arial', bold: true, color: { argb: 'FFC00000' } };
        promCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4E4' } };
      }
    });

  // ---------- Hoja Mismo Tecnico: reitero atendido por el mismo tecnico vs otro ----------
  // Solo se puede evaluar cuando la 2da reparacion esta en el archivo origen.
  const reitConDato = reit.filter((r) => byFolio[r['rdy_prd_reiterado']]);
  const esMismoTecnico = (r) => byFolio[r['rdy_prd_reiterado']]['toa_piv_nombre_tecnico'] === r['toa_piv_nombre_tecnico'];

  function agruparMismoTecnicoPor(campo) {
    const grupos = {};
    reitConDato.forEach((r) => {
      const key = (r[campo] || '').trim() || '(sin dato)';
      if (!grupos[key]) grupos[key] = { total: 0, mismo: 0 };
      grupos[key].total += 1;
      if (esMismoTecnico(r)) grupos[key].mismo += 1;
    });
    return grupos;
  }

  const mismoPorTecnico = agruparMismoTecnicoPor('toa_piv_nombre_tecnico');
  const mismoPorCausa = agruparMismoTecnicoPor('toa_piv_causa');
  const mismoGlobalTotal = reitConDato.length;
  const mismoGlobalSi = reitConDato.filter(esMismoTecnico).length;

  const mt = wb.addWorksheet('Mismo Tecnico');
  mt.columns = [{ width: 34 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 14 }];

  mt.mergeCells('A1:E1');
  estilizarTitulo(mt.getCell('A1'), 'REITERO ATENDIDO POR EL MISMO TECNICO vs. OTRO TECNICO');
  mt.getRow(1).height = 24;

  // --- Global ---
  mt.addRow([]);
  mt.mergeCells('A3:E3');
  mt.getCell('A3').value = 'GLOBAL (solo casos donde la 2da reparacion esta en el archivo)';
  mt.getCell('A3').font = { name: 'Arial', bold: true, size: 12 };

  const headerGlobalM = mt.addRow(['Casos con dato', 'Mismo Tecnico (SI)', 'Distinto Tecnico (NO)', '% Mismo Tecnico', '% Distinto Tecnico']);
  estilizarHeader(headerGlobalM);
  const rowGlobalM = mt.addRow([
    mismoGlobalTotal,
    mismoGlobalSi,
    mismoGlobalTotal - mismoGlobalSi,
    mismoGlobalTotal ? mismoGlobalSi / mismoGlobalTotal : null,
    mismoGlobalTotal ? (mismoGlobalTotal - mismoGlobalSi) / mismoGlobalTotal : null,
  ]);
  rowGlobalM.eachCell((cell) => { cell.font = { name: 'Arial' }; });
  rowGlobalM.getCell(4).numFmt = '0.0%';
  rowGlobalM.getCell(5).numFmt = '0.0%';

  // --- Por Causa ---
  mt.addRow([]);
  mt.addRow([]);
  const filaTituloCausaM = mt.rowCount + 1;
  mt.mergeCells(`A${filaTituloCausaM}:E${filaTituloCausaM}`);
  mt.getCell(`A${filaTituloCausaM}`).value = 'POR CAUSA';
  mt.getCell(`A${filaTituloCausaM}`).font = { name: 'Arial', bold: true, size: 12 };

  const headerCausaM = mt.addRow(['Causa', 'Casos con dato', 'Mismo Tecnico (SI)', 'Distinto Tecnico (NO)', '% Mismo Tecnico']);
  estilizarHeader(headerCausaM);
  Object.entries(mismoPorCausa)
    .sort((a, b) => (b[1].mismo / b[1].total) - (a[1].mismo / a[1].total))
    .forEach(([causa, s]) => {
      const row = mt.addRow([causa, s.total, s.mismo, s.total - s.mismo, s.total ? s.mismo / s.total : null]);
      row.eachCell((cell) => { cell.font = { name: 'Arial' }; });
      row.getCell(5).numFmt = '0.0%';
    });

  // --- Por Tecnico ---
  mt.addRow([]);
  mt.addRow([]);
  const filaTituloTecM = mt.rowCount + 1;
  mt.mergeCells(`A${filaTituloTecM}:E${filaTituloTecM}`);
  mt.getCell(`A${filaTituloTecM}`).value = 'POR TECNICO (de sus propias reincidencias, cuantas volvio a atender el mismo)';
  mt.getCell(`A${filaTituloTecM}`).font = { name: 'Arial', bold: true, size: 12 };

  const headerTecM = mt.addRow(['Tecnico', 'Casos con dato', 'Mismo Tecnico (SI)', 'Distinto Tecnico (NO)', '% Mismo Tecnico']);
  estilizarHeader(headerTecM);
  Object.entries(mismoPorTecnico)
    .sort((a, b) => (b[1].mismo / b[1].total) - (a[1].mismo / a[1].total))
    .forEach(([tecnico, s]) => {
      const row = mt.addRow([tecnico, s.total, s.mismo, s.total - s.mismo, s.total ? s.mismo / s.total : null]);
      row.eachCell((cell) => { cell.font = { name: 'Arial' }; });
      const pctCell = row.getCell(5);
      pctCell.numFmt = '0.0%';
      const pct = s.total ? s.mismo / s.total : 0;
      if (pct >= 0.5) {
        pctCell.font = { name: 'Arial', bold: true, color: { argb: 'FFC00000' } };
        pctCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4E4' } };
      }
    });

  const notas = wb.addWorksheet('Notas');
  notas.columns = [{ width: 100 }];
  const totalConSegunda = reit.filter((r) => byFolio[r['rdy_prd_reiterado']]).length;
  const notasTexto = [
    'ANALISIS DE REPARACIONES REITERADAS - COBRA',
    '',
    'Archivo origen: ' + path.basename(csvPath),
    'Generado: ' + new Date().toLocaleString('es-CL'),
    '',
    'Cada fila representa un caso donde una reparacion (1ra Reparacion) genero una nueva reparacion (2da Reparacion / Reitero) dentro de los 30 dias siguientes, segun el campo rdy_prd_tiene_reitero_30d del archivo origen.',
    '',
    '- "Clave de Cierre": codigo de cierre de la orden segun el servicio afectado (BA/TV/STB), tomado de las columnas toa_piv_clave_ba / toa_piv_clave_tv / toa_piv_clave_stb.',
    '- "Mismo Tecnico?": indica si el mismo tecnico atendio ambas reparaciones (SI/NO). "S/D" = sin dato porque la 2da reparacion no esta incluida en el archivo origen.',
    '- Filas resaltadas en AMARILLO: la 2da reparacion (folio de reitero) no aparece en el archivo origen -> probablemente se registro en otro periodo/reporte. Se conserva el folio de referencia pero no hay detalle de tecnico/causa.',
    '- Filas resaltadas en ROJO: la 2da reparacion fue atendida por un tecnico DISTINTO al de la 1ra reparacion.',
    '- "Dias hasta Reiteracion": dias transcurridos entre la 1ra y la 2da reparacion (campo rdy_prd_q_dias_reitero del archivo origen).',
    '',
    'Hoja "Estadistica": tasa de reincidencia (reincidencias / total de reparaciones) calculada Global, por Agencia, por Causa, por Subcausa y por Tecnico. Meta: no superar el 4%. Verde = cumple la meta, Rojo = no cumple. En "Por Subcausa" se omiten las subcausas con menos de 5 reparaciones totales (muestra muy chica para que la tasa sea representativa).',
    '',
    'Hoja "Tiempo Reiteracion": promedio y mediana de dias transcurridos entre la 1ra reparacion y la reiteracion, Global, por Causa y por Tecnico (ordenado de mas rapido a mas lento en reincidir). En "Por Tecnico" se resalta en rojo el promedio cuando es de 2 dias o menos, ya que sugiere reparaciones que no resolvieron el problema de fondo (el cliente volvio a fallar casi de inmediato).',
    '',
    'Hoja "Mismo Tecnico": de las reincidencias donde la 2da reparacion esta en el archivo origen, indica si fue atendida por el MISMO tecnico que hizo la 1ra reparacion o por uno DISTINTO. Un % alto de "Mismo Tecnico" por tecnico (resaltado en rojo si es 50% o mas) sugiere que ese tecnico no esta resolviendo el problema de raiz y vuelve el mismo a repetir el trabajo; un % bajo sugiere que otro tecnico tuvo que ir a corregir su reparacion.',
    '',
    'Total de registros en el archivo origen: ' + rows.length,
    'Total de casos con reitero detectado: ' + reit.length,
    'Casos donde la 2da reparacion SI esta en el archivo: ' + totalConSegunda,
    'Casos donde la 2da reparacion NO esta en el archivo: ' + (reit.length - totalConSegunda),
    '',
    'Para actualizar: reemplaza/sobreescribe el CSV en esta carpeta con los datos nuevos y vuelve a ejecutar Generar_Reporte_Reincidencias.bat. Cada ejecucion genera un archivo Excel nuevo (no sobreescribe reportes anteriores).',
  ];
  notasTexto.forEach((t, i) => {
    const row = notas.addRow([t]);
    row.getCell(1).font = i === 0 ? { name: 'Arial', bold: true, size: 12 } : { name: 'Arial', size: 10 };
    row.getCell(1).alignment = { wrapText: true, vertical: 'top' };
  });

  const outName = 'Reincidencias_COBRA_' + timestamp() + '.xlsx';
  const outPath = path.join(carpeta, outName);
  await wb.xlsx.writeFile(outPath);
  console.log('Excel generado:', outPath);

  // ================== DASHBOARD HTML (pagina local) ==================
  function aEntries(statsObj) {
    return Object.entries(statsObj).map(([label, s]) => ({
      label, total: s.total, reincidencias: s.reincidencias,
      tasa: s.total ? s.reincidencias / s.total : 0,
    }));
  }
  function peor(list, minMuestra = 1) {
    return list.filter((e) => e.total >= minMuestra).sort((a, b) => b.tasa - a.tasa)[0];
  }
  function mejor(list, minMuestra = 1) {
    return list.filter((e) => e.total >= minMuestra).sort((a, b) => a.tasa - b.tasa)[0];
  }

  const agenciaEntries = aEntries(statsPorAgencia).sort((a, b) => b.tasa - a.tasa);
  const causaEntries = aEntries(statsPorCausa).sort((a, b) => b.tasa - a.tasa);
  const subcausaEntries = aEntries(statsPorSubcausa).filter((e) => e.total >= 5).sort((a, b) => b.tasa - a.tasa);
  const tecnicoEntries = aEntries(statsPorTecnico).sort((a, b) => b.tasa - a.tasa);

  const tasaGlobal = statsGlobal.total ? statsGlobal.reincidencias / statsGlobal.total : 0;
  const maxTasaAgCa = Math.max(META, ...agenciaEntries.map((e) => e.tasa), ...causaEntries.map((e) => e.tasa)) * 1.15;
  const maxTasaSub = Math.max(META, ...subcausaEntries.map((e) => e.tasa)) * 1.15;
  const maxTasaTec = Math.max(META, ...tecnicoEntries.map((e) => e.tasa)) * 1.15;

  const diasTecnicoEntries = Object.entries(diasPorTecnico)
    .map(([label, arr]) => ({ label, casos: arr.length, promedio: promedio(arr) }))
    .sort((a, b) => a.promedio - b.promedio);
  const diasCausaEntries = Object.entries(diasPorCausa)
    .map(([label, arr]) => ({ label, casos: arr.length, promedio: promedio(arr) }))
    .sort((a, b) => a.promedio - b.promedio);
  const maxDias = Math.max(...diasTecnicoEntries.map((e) => e.promedio), ...diasCausaEntries.map((e) => e.promedio)) * 1.1;

  const buckets = [
    { label: '0-1 dias', min: 0, max: 1 },
    { label: '2-3 dias', min: 2, max: 3 },
    { label: '4-7 dias', min: 4, max: 7 },
    { label: '8-15 dias', min: 8, max: 15 },
    { label: '16-30 dias', min: 16, max: 30 },
  ].map((b) => {
    const count = diasReit.filter((d) => d >= b.min && d <= b.max).length;
    return { label: b.label, count, pct: diasReit.length ? (count / diasReit.length) * 100 : 0 };
  });
  const maxBucket = Math.max(...buckets.map((b) => b.count), 1);

  const mismoCausaEntries = Object.entries(mismoPorCausa)
    .map(([label, s]) => ({ label, total: s.total, mismo: s.mismo, tasa: s.total ? s.mismo / s.total : 0 }))
    .filter((e) => e.total >= 5)
    .sort((a, b) => b.tasa - a.tasa);
  const mismoPctGlobal = mismoGlobalTotal ? mismoGlobalSi / mismoGlobalTotal : 0;
  const maxMismo = Math.max(1, ...mismoCausaEntries.map((e) => e.tasa)) * 1.05;

  const peorAgencia = peor(agenciaEntries);
  const peorCausa = peor(causaEntries);
  const peorSubcausa = peor(subcausaEntries);
  const peorTecnico = peor(tecnicoEntries, 10);
  const mejorTecnico = mejor(tecnicoEntries, 10);
  const promedioGlobalDias = promedio(diasReit);
  const pctDentro7 = diasReit.length ? diasReit.filter((d) => d <= 7).length / diasReit.length : 0;
  const tecnicoMasRapido = diasTecnicoEntries.filter((e) => e.casos >= 3)[0];
  const causaMasMismoTecnico = mismoCausaEntries[0];

  const observaciones = [
    `Tasa global de reincidencia: <strong>${(tasaGlobal * 100).toFixed(1)}%</strong> — ${tasaGlobal <= META ? '<span class="ok">CUMPLE</span>' : '<span class="bad-text">NO CUMPLE</span>'} la meta del ${(META * 100).toFixed(0)}% (${statsGlobal.reincidencias} de ${statsGlobal.total} reparaciones reincidieron dentro de 30 dias).`,
    peorAgencia ? `La agencia con mayor tasa es <strong>${escapeHtml(peorAgencia.label)}</strong> con ${(peorAgencia.tasa * 100).toFixed(1)}% (${peorAgencia.reincidencias}/${peorAgencia.total}).` : '',
    peorCausa ? `La causa que mas reincide es <strong>${escapeHtml(peorCausa.label)}</strong> con ${(peorCausa.tasa * 100).toFixed(1)}%.` : '',
    peorSubcausa ? `La subcausa mas critica (con al menos 5 casos) es <strong>${escapeHtml(peorSubcausa.label)}</strong> con ${(peorSubcausa.tasa * 100).toFixed(1)}%.` : '',
    peorTecnico ? `El tecnico con mayor tasa de reincidencia (minimo 10 reparaciones) es <strong>${escapeHtml(peorTecnico.label)}</strong> con ${(peorTecnico.tasa * 100).toFixed(1)}% (${peorTecnico.reincidencias}/${peorTecnico.total}).` : '',
    mejorTecnico ? `El mejor desempeno (minimo 10 reparaciones) es de <strong>${escapeHtml(mejorTecnico.label)}</strong> con ${(mejorTecnico.tasa * 100).toFixed(1)}%.` : '',
    diasReit.length ? `En promedio, una reincidencia ocurre a los <strong>${promedioGlobalDias.toFixed(1)} dias</strong>; el ${(pctDentro7 * 100).toFixed(0)}% ocurre dentro de la primera semana.` : '',
    tecnicoMasRapido ? `<strong>${escapeHtml(tecnicoMasRapido.label)}</strong> tiene el promedio mas bajo de dias hasta reincidir (${tecnicoMasRapido.promedio.toFixed(1)} dias en ${tecnicoMasRapido.casos} casos) — posible senal de reparaciones que no resuelven el problema de fondo.` : '',
    mismoGlobalTotal ? `De los casos con dato completo, el <strong>${(mismoPctGlobal * 100).toFixed(0)}%</strong> de las reincidencias fueron re-atendidas por el mismo tecnico que hizo la reparacion original.` : '',
    causaMasMismoTecnico ? `En <strong>${escapeHtml(causaMasMismoTecnico.label)}</strong>, el ${(causaMasMismoTecnico.tasa * 100).toFixed(0)}% de las reincidencias las resuelve el mismo tecnico que hizo la reparacion original.` : '',
  ].filter(Boolean);

  // ---- datos adicionales para el layout tipo "Informe NPS" ----
  const causaPorTecnicoReit = {};
  reit.forEach((r) => {
    const t = (r['toa_piv_nombre_tecnico'] || '').trim() || '(sin dato)';
    const c = (r['toa_piv_causa'] || '').trim() || '(sin dato)';
    if (!causaPorTecnicoReit[t]) causaPorTecnicoReit[t] = {};
    causaPorTecnicoReit[t][c] = (causaPorTecnicoReit[t][c] || 0) + 1;
  });
  function causaFrecuenteTecnico(t) {
    const obj = causaPorTecnicoReit[t];
    if (!obj) return null;
    return Object.entries(obj).sort((a, b) => b[1] - a[1])[0];
  }

  const fechasAll = rows.map((r) => r['toa_piv_fecha_ingreso']).filter(Boolean).sort();
  const periodoStr = fechasAll.length
    ? `${fechasAll[0].slice(0, 10)} al ${fechasAll[fechasAll.length - 1].slice(0, 10)}`
    : 'N/D';
  const agenciasLista = [...new Set(rows.map((r) => (r['toa_piv_agencia'] || '').trim()).filter(Boolean))].join(' y ');

  const elegiblesHighlight = tecnicoEntries.filter((e) => e.total >= 10);
  const highlightsTop = [...elegiblesHighlight].sort((a, b) => a.tasa - b.tasa).slice(0, 4).map((e) => {
    const cf = causaFrecuenteTecnico(e.label);
    return {
      tecnico: e.label, tasa: +(e.tasa * 100).toFixed(1), total: e.total, reincidencias: e.reincidencias,
      nota: e.reincidencias === 0
        ? `Sin reincidencias en ${e.total} reparaciones`
        : `Su reincidencia mas frecuente: ${cf ? cf[0] : 'N/D'}${cf ? ` (${cf[1]} caso${cf[1] === 1 ? '' : 's'})` : ''}`,
    };
  });
  const highlightsBottom = [...elegiblesHighlight].sort((a, b) => b.tasa - a.tasa).slice(0, 4).map((e) => {
    const cf = causaFrecuenteTecnico(e.label);
    return {
      tecnico: e.label, tasa: +(e.tasa * 100).toFixed(1), total: e.total, reincidencias: e.reincidencias,
      nota: `Su reincidencia mas frecuente: ${cf ? cf[0] : 'N/D'}${cf ? ` (${cf[1]} caso${cf[1] === 1 ? '' : 's'})` : ''}`,
    };
  });

  const periodoSlug = obtenerSlugDesdeCsv(rows);
  const slugsExistentes = new Set(listarMesesExistentes(carpeta, 'Dashboard_Reincidencias'));
  slugsExistentes.add(periodoSlug);
  const archivos = [...slugsExistentes].sort().reverse().map((s) => ({
    slug: s,
    label: nombreMes(s),
    url: s === periodoSlug ? 'Dashboard_Reincidencias.html' : `Dashboard_Reincidencias_${s}.html`,
  }));

  const DATA = {
    archivoOrigen: path.basename(csvPath),
    generadoEl: new Date().toLocaleString('es-CL'),
    periodo: periodoStr,
    periodoSlug,
    archivos,
    agencias_lista: agenciasLista,
    meta: +(META * 100).toFixed(0),
    totalReparaciones: statsGlobal.total,
    totalReincidencias: statsGlobal.reincidencias,
    tasaGlobal: +(tasaGlobal * 100).toFixed(1),
    diasPromedio: diasReit.length ? +promedioGlobalDias.toFixed(1) : null,
    diasMediana: diasReit.length ? mediana(diasReit) : null,
    pctDentro7: +(pctDentro7 * 100).toFixed(1),
    mismoPct: mismoGlobalTotal ? +(mismoPctGlobal * 100).toFixed(1) : null,
    mismoTotal: mismoGlobalTotal,
    agencias: agenciaEntries.map((e) => ({ agencia: e.label, total: e.total, reincidencias: e.reincidencias, tasa: +(e.tasa * 100).toFixed(1) })),
    causas: causaEntries.map((e) => ({ causa: e.label, total: e.total, reincidencias: e.reincidencias, tasa: +(e.tasa * 100).toFixed(1) })),
    subcausas: subcausaEntries.map((e) => ({ subcausa: e.label, total: e.total, reincidencias: e.reincidencias, tasa: +(e.tasa * 100).toFixed(1) })),
    diasBuckets: buckets.map((b) => ({ label: b.label, count: b.count, pct: +b.pct.toFixed(1) })),
    mismoPorCausa: mismoCausaEntries.map((e) => ({ causa: e.label, total: e.total, mismo: e.mismo, tasa: +(e.tasa * 100).toFixed(1) })),
    tecnicos: tecnicoEntries.map((e) => ({
      tecnico: e.label, agencia: (metaTec[e.label] || {}).agencia || 'sin informacion',
      total: e.total, reincidencias: e.reincidencias, tasa: +(e.tasa * 100).toFixed(1),
    })),
    highlights: { top: highlightsTop, bottom: highlightsBottom },
    conclusiones: observaciones,
  };

  const html = `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Informe de Repetido Reparado · COBRA</title>
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
    <div class="eyebrow">Calidad &amp; Capacitacion · Analisis de Averias Reiteradas</div>
  </div>
  <h1>Informe de Repetido Reparado — COBRA</h1>
  <div class="subtitle" id="heroSubtitle">Analisis de reparaciones de Fibra Optica que volvieron a fallar dentro de 30 dias, a partir de las reparaciones registradas por COBRA. Incluye tasa de reincidencia por agencia, causa y tecnico, tiempos hasta la reiteracion, y si el reitero lo resolvio el mismo tecnico u otro.</div>
  <div class="meta-row" id="metaRow"></div>
  <div class="archive-row no-print" id="archiveRow"></div>
</header>

<main>
  <div class="kpi-grid" id="kpiGrid"></div>

  <section>
    <div class="section-title"><span class="num">01</span><h2>Tasa de Reincidencia por Agencia</h2></div>
    <p class="section-desc">Porcentaje de reparaciones que volvieron a fallar dentro de 30 dias, comparado contra la meta institucional.</p>
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
    <div class="section-title"><span class="num">02</span><h2>Tasa de Reincidencia por Causa</h2></div>
    <p class="section-desc">Que tipo de falla reincide con mas frecuencia dentro de los 30 dias posteriores a la reparacion.</p>
    <div class="grid-2">
      <div class="panel" style="height:340px;"><canvas id="chartCausa"></canvas></div>
      <div class="panel">
        <div style="font-size:13px;color:var(--text-dim);font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;">Detalle por causa</div>
        <table id="tablaCausa"></table>
        <div class="callout" style="margin-top:16px;" id="causaCallout"></div>
      </div>
    </div>
  </section>

  <section>
    <div class="section-title"><span class="num">03</span><h2>Tasa de Reincidencia por Subcausa</h2></div>
    <p class="section-desc">Detalle mas fino de la causa. Solo se incluyen subcausas con 5 o mas reparaciones totales, para que la tasa sea representativa.</p>
    <div class="panel" style="height:380px;"><canvas id="chartSubcausa"></canvas></div>
  </section>

  <section>
    <div class="section-title"><span class="num">04</span><h2>Tiempo hasta la Reiteracion</h2></div>
    <p class="section-desc">Cuantos dias pasan entre la reparacion original y la reincidencia. Mientras mas cerca de 0 dias, mas probable que la reparacion original no haya resuelto el problema de fondo.</p>
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
    <div class="section-title"><span class="num">05</span><h2>Reitero Resuelto por el Mismo Tecnico</h2></div>
    <p class="section-desc">De las reincidencias con la 2da reparacion identificada, si volvio el mismo tecnico o si tuvo que intervenir otro.</p>
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
    <div class="section-title"><span class="num">06</span><h2>Ranking de Tecnicos por Tasa de Reincidencia (min. 10 reparaciones)</h2></div>
    <p class="section-desc">Ordenado de menor a mayor tasa dentro de cada agencia. Base para reconocimiento y para focalizar coaching/capacitacion.</p>
    <div id="tecnicosPorAgencia"></div>
  </section>

  <section>
    <div class="section-title"><span class="num">07</span><h2>Mejor y Menor Desempeno</h2></div>
    <p class="section-desc">Tecnicos con minimo 10 reparaciones, ordenados por su tasa de reincidencia.</p>
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
  <span>🧾 Base: <b>\${DATA.totalReparaciones} reparaciones</b> · \${DATA.totalReincidencias} reincidencias</span>
  <span>🧮 Formula: <b>Reincidencias / Total de reparaciones</b></span>
\`;

// ---- Historial de meses ----
// meses.json se pisa completo en cada corrida del script, asi que se consulta en vivo
// (fetch) en vez de usar DATA.archivos: un informe archivado (ej. Dashboard_Reincidencias_2026-07.html)
// nunca se vuelve a regenerar, asi que su DATA embebido quedaria congelado en el
// historial que existia el dia que se genero.
function renderArchiveRow(lista) {
  if (!lista || lista.length < 2) return;
  const pills = lista.map(a =>
    a.slug === DATA.periodoSlug
      ? \`<span class="archive-pill current">\${a.label}</span>\`
      : \`<a class="archive-pill" href="\${a.url}">\${a.label}</a>\`
  ).join('');
  document.getElementById('archiveRow').innerHTML = \`<span class="archive-label">📁 Meses disponibles:</span>\${pills}\`;
}
renderArchiveRow(DATA.archivos); // fallback inmediato mientras llega el fetch
fetch('meses.json', { cache: 'no-store' })
  .then(r => r.ok ? r.json() : null)
  .then(data => { if (data && data.reincidencias) renderArchiveRow(data.reincidencias); })
  .catch(() => {}); // sin conexion o meses.json no existe todavia: se queda con el fallback

const fontColor = '#6b7a8c';
Chart.defaults.color = fontColor;
Chart.defaults.borderColor = 'rgba(20,50,80,0.06)';
Chart.defaults.font.family = "'Segoe UI', Arial, sans-serif";

// ---- KPI cards ----
const metaGap = +(DATA.tasaGlobal - DATA.meta).toFixed(1);
const metaSub = metaGap <= 0
  ? \`<span style="color:var(--promotor)">▼ \${Math.abs(metaGap)} pts bajo la meta (\${DATA.meta}%) — cumple</span>\`
  : \`<span style="color:var(--detractor)">▲ \${metaGap} pts sobre la meta (\${DATA.meta}%) — no cumple</span>\`;
document.getElementById('kpiGrid').innerHTML = \`
  <div class="kpi-card"><div class="label">Tasa Global de Reincidencia</div><div class="value hl">\${DATA.tasaGlobal}%</div><div class="sub">\${metaSub}</div></div>
  <div class="kpi-card"><div class="label">Total Reparaciones</div><div class="value">\${DATA.totalReparaciones}</div><div class="sub">Periodo \${DATA.periodo}</div></div>
  <div class="kpi-card"><div class="label">Reincidencias (30d)</div><div class="value">\${DATA.totalReincidencias}</div><div class="sub">Reparaciones que volvieron a fallar</div></div>
  <div class="kpi-card"><div class="label">Dias Prom. hasta Reitero</div><div class="value">\${DATA.diasPromedio ?? '-'}</div><div class="sub">Mediana: \${DATA.diasMediana ?? '-'} dias · \${DATA.pctDentro7}% dentro de 7 dias</div></div>
  <div class="kpi-card"><div class="label">% Mismo Tecnico en Reitero</div><div class="value">\${DATA.mismoPct ?? '-'}%</div><div class="sub">\${DATA.mismoTotal} casos con dato completo</div></div>
\`;

// ---- Chart 01: Agencia ----
new Chart(document.getElementById('chartAgencia'), {
  type: 'bar',
  data: {
    labels: DATA.agencias.map(a=>a.agencia),
    datasets: [
      { label:'Tasa de reincidencia (%)', data: DATA.agencias.map(a=>a.tasa), backgroundColor: DATA.agencias.map(a=> a.tasa<=DATA.meta ? 'rgba(31,169,113,0.85)' : 'rgba(226,82,62,0.85)'), borderRadius:6, maxBarThickness:60 },
      { label:'Meta ('+DATA.meta+'%)', data: DATA.agencias.map(()=>DATA.meta), type:'line', borderColor:'#e2523e', borderDash:[5,4], borderWidth:1.5, pointRadius:0, tension:0 }
    ]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins:{ legend:{ position:'top', labels:{boxWidth:10} } },
    scales:{ y:{ min:0, title:{display:true,text:'Tasa %'}, grid:{color:'rgba(20,50,80,0.06)'} } }
  }
});
let rowsAg = '<tr><th>Agencia</th><th>Tasa</th><th>vs. Meta</th><th>Reparaciones</th></tr>';
DATA.agencias.slice().sort((a,b)=>a.tasa-b.tasa).forEach(a=>{
  const gap = +(a.tasa - DATA.meta).toFixed(1);
  const gapTxt = gap<=0 ? \`<span style="color:var(--promotor)">\${gap}</span>\` : \`<span style="color:var(--detractor)">+\${gap}</span>\`;
  rowsAg += \`<tr><td>\${a.agencia}</td><td><span class="badge \${npsClass(a.tasa,DATA.meta)}">\${a.tasa}%</span></td><td>\${gapTxt}</td><td>\${a.total}</td></tr>\`;
});
document.getElementById('tablaAgencia').innerHTML = rowsAg;
const agBajoMeta = DATA.agencias.filter(a=>a.tasa > DATA.meta);
document.getElementById('agenciaCallout').innerHTML = agBajoMeta.length
  ? \`<b>Sobre la meta:</b> \${agBajoMeta.map(a=>\`\${a.agencia} (\${a.tasa}%)\`).join(' y ')} supera el \${DATA.meta}% institucional.\`
  : \`<b>Todas las agencias dentro de la meta</b> de \${DATA.meta}% este periodo.\`;

// ---- Chart 02: Causa ----
new Chart(document.getElementById('chartCausa'), {
  type: 'bar',
  data: {
    labels: DATA.causas.map(c=>c.causa),
    datasets: [
      { label:'Tasa de reincidencia (%)', data: DATA.causas.map(c=>c.tasa), backgroundColor: DATA.causas.map(c=> c.tasa<=DATA.meta ? 'rgba(31,169,113,0.85)' : 'rgba(226,82,62,0.85)'), borderRadius:5, maxBarThickness:26 },
      { label:'Meta ('+DATA.meta+'%)', data: DATA.causas.map(()=>DATA.meta), type:'line', borderColor:'#e2523e', borderDash:[5,4], borderWidth:1.5, pointRadius:0, tension:0 }
    ]
  },
  options: {
    indexAxis:'y', responsive:true, maintainAspectRatio:false,
    plugins:{ legend:{ position:'top', labels:{boxWidth:10} } },
    scales:{ x:{ min:0, title:{display:true,text:'Tasa %'}, grid:{color:'rgba(20,50,80,0.06)'} } }
  }
});
let rowsCa = '<tr><th>Causa</th><th>Tasa</th><th>Reincidencias</th><th>Total</th></tr>';
DATA.causas.slice().sort((a,b)=>b.tasa-a.tasa).forEach(c=>{
  rowsCa += \`<tr><td>\${c.causa}</td><td><span class="badge \${npsClass(c.tasa,DATA.meta)}">\${c.tasa}%</span></td><td>\${c.reincidencias}</td><td>\${c.total}</td></tr>\`;
});
document.getElementById('tablaCausa').innerHTML = rowsCa;
const peorCausaC = DATA.causas.slice().sort((a,b)=>b.tasa-a.tasa)[0];
document.getElementById('causaCallout').innerHTML = peorCausaC
  ? \`<b>Foco principal:</b> \${peorCausaC.causa} tiene la mayor tasa de reincidencia (\${peorCausaC.tasa}%, \${peorCausaC.reincidencias} de \${peorCausaC.total} reparaciones).\`
  : '';

// ---- Chart 03: Subcausa ----
const subOrd = DATA.subcausas.slice().sort((a,b)=>b.tasa-a.tasa);
new Chart(document.getElementById('chartSubcausa'), {
  type: 'bar',
  data: {
    labels: subOrd.map(s=>s.subcausa),
    datasets: [
      { label:'Tasa de reincidencia (%)', data: subOrd.map(s=>s.tasa), backgroundColor: subOrd.map(s=> s.tasa<=DATA.meta ? 'rgba(31,169,113,0.85)' : 'rgba(226,82,62,0.85)'), borderRadius:5, maxBarThickness:22 },
      { label:'Meta ('+DATA.meta+'%)', data: subOrd.map(()=>DATA.meta), type:'line', borderColor:'#e2523e', borderDash:[5,4], borderWidth:1.5, pointRadius:0, tension:0 }
    ]
  },
  options: {
    indexAxis:'y', responsive:true, maintainAspectRatio:false,
    plugins:{ legend:{ position:'top', labels:{boxWidth:10} } },
    scales:{ x:{ min:0, title:{display:true,text:'Tasa %'}, grid:{color:'rgba(20,50,80,0.06)'} } }
  }
});

// ---- Chart 04: Dias hasta reiteracion ----
new Chart(document.getElementById('chartDias'), {
  type: 'bar',
  data: {
    labels: DATA.diasBuckets.map(b=>b.label),
    datasets: [
      { label:'Casos', data: DATA.diasBuckets.map(b=>b.count), backgroundColor:'rgba(0,113,206,0.85)', borderRadius:6, maxBarThickness:60 }
    ]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins:{ legend:{ display:false } },
    scales:{ y:{ min:0, title:{display:true,text:'N° de casos'}, grid:{color:'rgba(20,50,80,0.06)'} } }
  }
});
document.getElementById('tablaDias').innerHTML = \`
  <tr><th>Indicador</th><th>Valor</th></tr>
  <tr><td>Promedio</td><td>\${DATA.diasPromedio ?? '-'} dias</td></tr>
  <tr><td>Mediana</td><td>\${DATA.diasMediana ?? '-'} dias</td></tr>
  <tr><td>Dentro de 7 dias</td><td><span class="badge \${DATA.pctDentro7>=60?'lo':'mid'}">\${DATA.pctDentro7}%</span></td></tr>
\`;
document.getElementById('diasCallout').innerHTML =
  \`<b>\${DATA.pctDentro7}%</b> de las reincidencias ocurre dentro de la primera semana. Cuanto mas alto este numero, mas senal de que la reparacion original no resolvio el problema de fondo (no de una falla nueva).\`;

// ---- Chart 05: Mismo tecnico (donut) ----
new Chart(document.getElementById('chartMismo'), {
  type: 'doughnut',
  data: {
    labels: ['Mismo tecnico', 'Tecnico distinto'],
    datasets: [{ data: [DATA.mismoPct, +(100-DATA.mismoPct).toFixed(1)], backgroundColor: ['#e2523e','#1fa971'], borderWidth:0 }]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins:{ legend:{ position:'bottom', labels:{boxWidth:10} } },
    cutout: '62%'
  }
});
let rowsMi = '<tr><th>Causa</th><th>% Mismo Tecnico</th><th>Casos</th></tr>';
DATA.mismoPorCausa.forEach(m=>{
  rowsMi += \`<tr><td>\${m.causa}</td><td><span class="badge \${m.tasa>=50?'lo':'hi'}">\${m.tasa}%</span></td><td>\${m.mismo}/\${m.total}</td></tr>\`;
});
document.getElementById('tablaMismo').innerHTML = rowsMi;
document.getElementById('mismoCallout').innerHTML =
  \`<b>\${DATA.mismoPct}%</b> de las reincidencias con dato completo fueron re-atendidas por el mismo tecnico que hizo la reparacion original — señal de que no resolvio el problema a la primera.\`;

// ---- Tabla: Tecnicos por agencia ----
const agenciasOrden = DATA.agencias.slice().sort((a,b)=>b.total-a.total).map(a=>a.agencia);
let tecnicosHtml = '';
agenciasOrden.forEach(agencia => {
  const tecsAg = DATA.tecnicos.filter(t=>t.agencia===agencia && t.total>=10).sort((a,b)=>a.tasa-b.tasa);
  if (!tecsAg.length) return;
  let rowsTec = '<tr><th>#</th><th>Tecnico</th><th>Tasa</th><th>Reparaciones</th><th>Reincidencias</th></tr>';
  tecsAg.forEach((t,i)=>{
    rowsTec += \`<tr>
      <td>\${i+1}</td>
      <td>\${titleCase(t.tecnico)}</td>
      <td><span class="badge \${npsClass(t.tasa,DATA.meta)}">\${t.tasa}%</span></td>
      <td>\${t.total}</td>
      <td>\${t.reincidencias}</td>
    </tr>\`;
  });
  tecnicosHtml += \`<div class="panel" style="margin-bottom:16px;">
    <div style="font-size:13px;color:var(--text-dim);font-weight:700;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px;">\${agencia} · \${tecsAg.length} tecnico\${tecsAg.length===1?'':'s'} (min. 10 reparaciones)</div>
    <table>\${rowsTec}</table>
  </div>\`;
});
document.getElementById('tecnicosPorAgencia').innerHTML = tecnicosHtml;

// ---- Highlights: top / bottom tecnicos ----
function techCard(t, kind){
  const cls = kind==='top' ? 'hi' : 'lo';
  return \`<div class="tech-card \${kind==='top'?'top-card':'bottom-card'}">
    <div class="tech-head"><span class="tech-name">\${titleCase(t.tecnico)}</span><span class="tech-nps \${cls}">\${t.tasa}%</span></div>
    <div class="tech-meta">\${t.total} reparaciones · \${t.reincidencias} reincidencia\${t.reincidencias===1?'':'s'}</div>
    <div class="tech-quote">\${t.nota}</div>
  </div>\`;
}
document.getElementById('topTecnicos').innerHTML = DATA.highlights.top.map(t=>techCard(t,'top')).join('');
document.getElementById('bottomTecnicos').innerHTML = DATA.highlights.bottom.map(t=>techCard(t,'bottom')).join('');

// ---- Conclusiones ----
document.getElementById('recList').innerHTML = DATA.conclusiones.map((c,i)=>\`
  <div class="rec-item"><div class="idx">\${i+1}</div><p>\${c}</p></div>
\`).join('');

// ---- Footer ----
document.getElementById('footerText').innerHTML =
  \`Informe generado a partir de \${DATA.archivoOrigen} · \${DATA.agencias_lista} · Generado \${DATA.generadoEl}\`;
<\/script>

</body>
</html>`;

  const htmlPath = path.join(carpeta, 'Dashboard_Reincidencias.html');
  fs.writeFileSync(htmlPath, html, 'utf8');
  const htmlArchivoPath = path.join(carpeta, `Dashboard_Reincidencias_${periodoSlug}.html`);
  fs.writeFileSync(htmlArchivoPath, html, 'utf8');
  console.log('Dashboard HTML generado:', htmlPath, 'y', htmlArchivoPath);
}

main().catch((err) => {
  console.error('ERROR:', err.message);
  process.exitCode = 1;
});
