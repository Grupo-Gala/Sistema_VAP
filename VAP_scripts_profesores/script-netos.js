/**
 * =========================================================
 *  NETOS NÓMINA -> VINCULACIÓN GASTOS PERSONAL 2026 netos
 *  Match por: EMPRESA (destino col E) + Nº EMPLEADO (destino col F)
 *  Escribe NETO en: destino col T
 *
 *  Origen: Google Sheet convertido desde Excel
 *   - Busca columnas por CABECERA (no por letra)
 *   - Toma el ÚLTIMO registro por (empresa+número)
 *
 *  Control:
 *   - LOG de cambios (antes/después)
 *   - NO_ENCONTRADO_DESTINO
 *   - DUPLICADO_ORIGEN (informativo)
 *   - DUPLICADO_DESTINO (bloquea escritura)
 * =========================================================
 */

// ====== CONFIG REAL (ya rellenada con tus IDs) ======
const SOURCE_SPREADSHEET_ID = '1bUt5KbZ4efseIaB7CoE2bjh7mSVOJTMTIjgouSK7Fog';
const SOURCE_SHEET_NAME = 'Informe'; // si no coincide, ejecuta listarHojasOrigen()

const TARGET_SPREADSHEET_ID = '1HDwgihfz91atsLE99Qu0k3kTyBKpo3Iu3SaUtVGMPTs';
const TARGET_SHEET_NAME = 'Hoja 1';

// DESTINO: Empresa(E=5), Nº empleado(F=6), Neto(T=20)
const TGT_EMPRESA_COL = 5; // E
const TGT_NUM_COL = 6;     // F
const TGT_NETO_COL = 20;   // T

// Cabeceras en el ORIGEN (ajústalas si en tu hoja convertida aparecen distinto)
const SRC_HEADER_EMPRESA = 'Empresa';
const SRC_HEADER_NUM = 'Empleado -  Código';
const SRC_HEADER_NETO = 'Líquido a percibir';

// LOG
const LOG_SHEET_NAME = 'LOG NETOS';
const LOG_HEADERS = [
  'Timestamp',
  'Run ID',
  'Tipo',
  'Empresa',
  'Nº Empleado',
  'Fila origen',
  'Fila destino',
  'Neto origen',
  'Neto destino antes',
  'Neto destino después',
  'Detalle'
];

/**
 * FUNCIÓN PRINCIPAL (esta es la que debes ejecutar y la que debes usar en el trigger)
 */
function actualizarNetosDesdeNominas() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const runId = Utilities.getUuid();
  const now = new Date();

  try {
    const srcSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const srcSheet = srcSS.getSheetByName(SOURCE_SHEET_NAME);
    if (!srcSheet) throw new Error(`No existe la pestaña origen: ${SOURCE_SHEET_NAME}`);

    const tgtSS = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const tgtSheet = tgtSS.getSheetByName(TARGET_SHEET_NAME);
    if (!tgtSheet) throw new Error(`No existe la pestaña destino: ${TARGET_SHEET_NAME}`);

    const logSheet = getOrCreateLogSheet_(tgtSS, LOG_SHEET_NAME, LOG_HEADERS);
    const logRows = [];

    // ---- ORIGEN ----
    const srcLastRow = srcSheet.getLastRow();
    const srcLastCol = srcSheet.getLastColumn();
    if (srcLastRow < 2) {
      logRows.push(makeLogRow_(now, runId, 'RESUMEN', '', '', '', '', '', '', '', 'Origen sin datos.'));
      appendLogRows_(logSheet, logRows);
      return;
    }

    const srcValues = srcSheet.getRange(1, 1, srcLastRow, srcLastCol).getValues();

    const headerRowIndex = findHeaderRow_(srcValues, SRC_HEADER_EMPRESA);
    if (headerRowIndex === -1) throw new Error(`No encuentro la fila de cabeceras con "${SRC_HEADER_EMPRESA}"`);

    const headers = srcValues[headerRowIndex].map(h => (h || '').toString().trim());
    const srcEmpresaCol = findHeaderCol_(headers, SRC_HEADER_EMPRESA);
    const srcNumCol = findHeaderCol_(headers, SRC_HEADER_NUM);
    const srcNetoCol = findHeaderCol_(headers, SRC_HEADER_NETO);

    if (!srcEmpresaCol || !srcNumCol || !srcNetoCol) {
      throw new Error(
        `No encuentro cabeceras necesarias en origen.\n` +
        `Necesito: "${SRC_HEADER_EMPRESA}", "${SRC_HEADER_NUM}", "${SRC_HEADER_NETO}".`
      );
    }

    // srcMap: key -> {empresaDisp, numDisp, neto, srcRow}
    const srcMap = new Map();
    const srcDupes = new Map();

    for (let r = headerRowIndex + 1; r < srcValues.length; r++) {
      const row = srcValues[r];
      const empresaDisp = (row[srcEmpresaCol - 1] || '').toString().trim();
      const numDisp = (row[srcNumCol - 1] || '').toString().trim();
      const neto = row[srcNetoCol - 1];

      if (!empresaDisp || !numDisp) continue;
      if (!hasValue_(neto)) continue;

      const key = makeKey_(empresaDisp, numDisp);

      if (srcMap.has(key)) srcDupes.set(key, (srcDupes.get(key) || 1) + 1);

      // Última fila gana
      srcMap.set(key, { empresaDisp, numDisp, neto, srcRow: r + 1 });
    }

    srcDupes.forEach((count, key) => {
      const last = srcMap.get(key);
      logRows.push(makeLogRow_(now, runId, 'DUPLICADO_ORIGEN', last.empresaDisp, last.numDisp, last.srcRow, '', last.neto, '', '', `Aparece ${count} veces. Se usa la última fila.`));
    });

    if (srcMap.size === 0) {
      logRows.push(makeLogRow_(now, runId, 'RESUMEN', '', '', '', '', '', '', '', 'Origen sin valores de neto.'));
      appendLogRows_(logSheet, logRows);
      return;
    }

    // ---- DESTINO ----
    const tgtLastRow = tgtSheet.getLastRow();
    if (tgtLastRow < 2) {
      logRows.push(makeLogRow_(now, runId, 'RESUMEN', '', '', '', '', '', '', '', 'Destino sin filas.'));
      appendLogRows_(logSheet, logRows);
      return;
    }

    // Leemos hasta la columna T
    const tgtValues = tgtSheet.getRange(2, 1, tgtLastRow - 1, TGT_NETO_COL).getValues();

    const tgtRowByKey = new Map();
    const tgtDupes = new Map();

    for (let i = 0; i < tgtValues.length; i++) {
      const row = tgtValues[i];
      const empresaT = (row[TGT_EMPRESA_COL - 1] || '').toString().trim();
      const numT = (row[TGT_NUM_COL - 1] || '').toString().trim();
      if (!empresaT || !numT) continue;

      const key = makeKey_(empresaT, numT);
      const rowNum = i + 2;

      if (tgtRowByKey.has(key)) {
        const prev = tgtRowByKey.get(key);
        tgtDupes.set(key, (tgtDupes.get(key) || [prev]).concat([rowNum]));
      } else {
        tgtRowByKey.set(key, rowNum);
      }
    }

    tgtDupes.forEach((rowsArr, key) => {
      const src = srcMap.get(key);
      logRows.push(makeLogRow_(
        now, runId, 'DUPLICADO_DESTINO',
        src ? src.empresaDisp : '',
        src ? src.numDisp : '',
        src ? src.srcRow : '',
        rowsArr.join(','),
        src ? src.neto : '',
        '', '',
        `Clave repetida en destino (filas ${rowsArr.join(', ')}). NO se escribe hasta corregir.`
      ));
    });

    // ---- ESCRITURA + CONTROL CAMBIOS ----
    let updated = 0, omitted = 0, notFound = 0, blocked = 0;

    for (let i = 0; i < tgtValues.length; i++) {
      const row = tgtValues[i];
      const empresaT = (row[TGT_EMPRESA_COL - 1] || '').toString().trim();
      const numT = (row[TGT_NUM_COL - 1] || '').toString().trim();
      const before = row[TGT_NETO_COL - 1];

      const key = makeKey_(empresaT, numT);

      // Si duplicado en destino, bloquear
      if (tgtDupes.has(key)) {
        blocked++;
        continue;
      }

      // Si no hay dato en origen, no tocar
      const src = srcMap.get(key);
      if (!src) continue;

      // Si la celda destino ya tiene cualquier valor, no tocar
      if (hasValue_(before)) {
        omitted++;
        logRows.push(makeLogRow_(now, runId, 'OMITIDO_CELDA_CON_VALOR', src.empresaDisp, src.numDisp, src.srcRow, i + 2, src.neto, before, before, 'Celda T ya tiene valor; no se sobreescribe.'));
        continue;
      }

      // Celda vacía: escribir neto del origen celda a celda
      tgtSheet.getRange(i + 2, TGT_NETO_COL).setValue(src.neto);
      updated++;
      logRows.push(makeLogRow_(now, runId, 'ACTUALIZADO', src.empresaDisp, src.numDisp, src.srcRow, i + 2, src.neto, before, src.neto, ''));
    }

    // Origen que NO existe en destino
    srcMap.forEach((src, key) => {
      if (tgtDupes.has(key)) return;
      if (!tgtRowByKey.has(key)) {
        notFound++;
        logRows.push(makeLogRow_(now, runId, 'NO_ENCONTRADO_DESTINO', src.empresaDisp, src.numDisp, src.srcRow, '', src.neto, '', '', 'No existe clave (empresa+nº) en destino.'));
      }
    });

    logRows.push(makeLogRow_(
      now, runId, 'RESUMEN', '', '', '', '', '', '', '',
      `Actualizados=${updated} | OmitidosCeldaConValor=${omitted} | NoEncontrados=${notFound} | BloqueadosDuplicadoDestino=${blocked} | DuplicadosOrigen=${srcDupes.size} | DuplicadosDestino=${tgtDupes.size}`
    ));

    appendLogRows_(logSheet, logRows);

  } finally {
    lock.releaseLock();
  }
}

/**
 * Utilidad: lista las pestañas del ORIGEN por si SOURCE_SHEET_NAME no es "Informe"
 */
function listarHojasOrigen() {
  const ss = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  Logger.log(ss.getSheets().map(s => s.getName()).join(' | '));
}

/** ================= Helpers ================= */

function getOrCreateLogSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function appendLogRows_(logSheet, rows) {
  if (!rows || rows.length === 0) return;
  const startRow = logSheet.getLastRow() + 1;
  logSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function makeLogRow_(ts, runId, tipo, empresa, num, filaOrigen, filaDestino, netoOrigen, antes, despues, detalle) {
  return [
    ts,
    runId,
    tipo,
    empresa || '',
    num || '',
    filaOrigen || '',
    filaDestino || '',
    netoOrigen ?? '',
    antes ?? '',
    despues ?? '',
    detalle || ''
  ];
}

function findHeaderRow_(matrix, requiredHeader) {
  const needle = normalizeText_(requiredHeader);
  for (let r = 0; r < matrix.length; r++) {
    const row = matrix[r].map(x => normalizeText_((x || '').toString()));
    if (row.includes(needle)) return r;
  }
  return -1;
}

function findHeaderCol_(headers, headerName) {
  const needle = normalizeText_(headerName);
  for (let i = 0; i < headers.length; i++) {
    if (normalizeText_(headers[i]) === needle) return i + 1;
  }
  return null;
}

function makeKey_(empresa, numEmpleado) {
  return `${normalizeText_(empresa)}__${normalizeEmployeeNumber_(numEmpleado)}`;
}

function normalizeText_(s) {
  const cleaned = (s || '').toString().trim().replace(/\s+/g, ' ').toLowerCase();
  return cleaned.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function normalizeEmployeeNumber_(s) {
  return (s || '').toString().trim().replace(/\s+/g, '').replace(/-/g, '').toLowerCase();
}

function hasValue_(v) {
  if (v === null || v === undefined) return false;
  if (typeof v === 'string') return v.trim() !== '';
  return true;
}

function valuesEqual_(a, b) {
  if (a === b) return true;
  const na = toNumber_(a);
  const nb = toNumber_(b);
  if (na !== null && nb !== null) return na === nb;
  return (a ?? '').toString().trim() === (b ?? '').toString().trim();
}

function toNumber_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  const s = v.toString().trim();
  if (!s) return null;
  const normalized = s.replace(/\./g, '').replace(',', '.');
  const num = Number(normalized);
  return Number.isFinite(num) ? num : null;
}
