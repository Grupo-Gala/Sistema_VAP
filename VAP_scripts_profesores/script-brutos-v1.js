/**
 * =========================================================
 *  BRUTOS + IRPF -> VINCULACIÓN GASTOS PERSONAL 2026 brutos
 *  Match por: EMPRESA (destino col E) + Nº EMPLEADO (destino col F)
 *
 *  Origen: Google Sheet convertido desde Excel (mismo que netos)
 *   - Bruto: Col H
 *   - IRPF importe: Col K
 *   - IRPF %: Col L
 *
 *  Destino (Hoja 1):
 *   - Empresa: E
 *   - Nº empleado: F
 *   - Bruto: L
 *   - IRPF importe: R
 *   - IRPF %: S
 *
 *  Control:
 *   - LOG de cambios (antes/después)
 *   - NO_ENCONTRADO_DESTINO
 *   - DUPLICADO_ORIGEN (informativo)
 *   - DUPLICADO_DESTINO (bloquea escritura)
 * =========================================================
 */

// ====== CONFIG REAL ======
const SOURCE_SPREADSHEET_ID = '1bUt5KbZ4efseIaB7CoE2bjh7mSVOJTMTIjgouSK7Fog';
const SOURCE_SHEET_NAME = 'Informe'; // si no coincide, ejecuta listarHojasOrigen()

const TARGET_SPREADSHEET_ID = '1eTVPws0NVRMIlXgZYsUZANdfMyFYQtDlEEK5ujFt04A';
const TARGET_SHEET_NAME = 'Hoja 1';

// DESTINO: Empresa(E=5), Nº empleado(F=6)
const TGT_EMPRESA_COL = 5; // E
const TGT_NUM_COL = 6;     // F

// DESTINO: columnas de salida
const TGT_BRUTO_COL = 12;      // L
const TGT_IRPF_IMP_COL = 18;   // R
const TGT_IRPF_PCT_COL = 19;   // S

// ORIGEN: columnas por letra (en tu hoja origen)
const SRC_EMPRESA_COL = 0; // se detecta por cabecera, ver abajo
const SRC_NUM_COL = 0;     // se detecta por cabecera, ver abajo

// Cabeceras en el ORIGEN para emparejar (igual que en netos)
const SRC_HEADER_EMPRESA = 'Empresa';
const SRC_HEADER_NUM = 'Empleado -  Código';

// ORIGEN: columnas de datos (1-based)
const SRC_BRUTO_COL = 8;      // H
const SRC_IRPF_IMP_COL = 11;  // K
const SRC_IRPF_PCT_COL = 12;  // L

// LOG
const LOG_SHEET_NAME = 'LOG BRUTOS';
const LOG_HEADERS = [
  'Timestamp',
  'Run ID',
  'Tipo',
  'Empresa',
  'Nº Empleado',
  'Fila origen',
  'Fila destino',
  'Bruto origen',
  'Bruto antes',
  'Bruto después',
  'IRPF € origen',
  'IRPF € antes',
  'IRPF € después',
  'IRPF % origen',
  'IRPF % antes',
  'IRPF % después',
  'Detalle'
];

/**
 * FUNCIÓN PRINCIPAL (esta debe aparecer en el desplegable)
 */
function actualizarBrutosDesdeNominas() {
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
      logRows.push(makeLogRow_(now, runId, 'RESUMEN', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Origen sin datos.'));
      appendLogRows_(logSheet, logRows);
      return;
    }

    const srcValues = srcSheet.getRange(1, 1, srcLastRow, srcLastCol).getValues();

    const headerRowIndex = findHeaderRow_(srcValues, SRC_HEADER_EMPRESA);
    if (headerRowIndex === -1) throw new Error(`No encuentro la fila de cabeceras con "${SRC_HEADER_EMPRESA}"`);

    const headers = srcValues[headerRowIndex].map(h => (h || '').toString().trim());
    const srcEmpresaCol = findHeaderCol_(headers, SRC_HEADER_EMPRESA);
    const srcNumCol = findHeaderCol_(headers, SRC_HEADER_NUM);

    if (!srcEmpresaCol || !srcNumCol) {
      throw new Error(`No encuentro cabeceras de clave en origen: "${SRC_HEADER_EMPRESA}" y "${SRC_HEADER_NUM}".`);
    }

    // srcMap: key -> {empresaDisp, numDisp, bruto?, irpfImp?, irpfPct?, srcRow}
    const srcMap = new Map();
    const srcDupes = new Map();

    for (let r = headerRowIndex + 1; r < srcValues.length; r++) {
      const row = srcValues[r];
      const empresaDisp = (row[srcEmpresaCol - 1] || '').toString().trim();
      const numDisp = (row[srcNumCol - 1] || '').toString().trim();
      if (!empresaDisp || !numDisp) continue;

      const bruto = row[SRC_BRUTO_COL - 1];
      const irpfImp = row[SRC_IRPF_IMP_COL - 1];
      const irpfPct = row[SRC_IRPF_PCT_COL - 1];

      // Si los 3 están vacíos, no aporta nada
      if (!hasValue_(bruto) && !hasValue_(irpfImp) && !hasValue_(irpfPct)) continue;

      const key = makeKey_(empresaDisp, numDisp);

      if (srcMap.has(key)) srcDupes.set(key, (srcDupes.get(key) || 1) + 1);

      // Última fila gana (pero guardamos solo lo que tenga valor)
      const existing = srcMap.get(key) || { empresaDisp, numDisp, srcRow: r + 1 };
      const obj = {
        empresaDisp,
        numDisp,
        srcRow: r + 1,
        bruto: hasValue_(bruto) ? bruto : existing.bruto,
        irpfImp: hasValue_(irpfImp) ? irpfImp : existing.irpfImp,
        irpfPct: hasValue_(irpfPct) ? irpfPct : existing.irpfPct,
      };

      srcMap.set(key, obj);
    }

    srcDupes.forEach((count, key) => {
      const last = srcMap.get(key);
      logRows.push(makeLogRow_(
        now, runId, 'DUPLICADO_ORIGEN',
        last.empresaDisp, last.numDisp, last.srcRow, '',
        last.bruto ?? '', '', '',
        last.irpfImp ?? '', '', '',
        last.irpfPct ?? '', '', '',
        `Aparece ${count} veces. Se usa el último valor por campo (no vacío).`
      ));
    });

    if (srcMap.size === 0) {
      logRows.push(makeLogRow_(now, runId, 'RESUMEN', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Origen sin valores (bruto/irpf).'));
      appendLogRows_(logSheet, logRows);
      return;
    }

    // ---- DESTINO ----
    const tgtLastRow = tgtSheet.getLastRow();
    if (tgtLastRow < 2) {
      logRows.push(makeLogRow_(now, runId, 'RESUMEN', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Destino sin filas.'));
      appendLogRows_(logSheet, logRows);
      return;
    }

    // Necesitamos leer al menos hasta S (=19)
    const readCols = Math.max(TGT_IRPF_PCT_COL, TGT_NUM_COL);
    const tgtValues = tgtSheet.getRange(2, 1, tgtLastRow - 1, readCols).getValues();

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
        src ? (src.bruto ?? '') : '', '', '',
        src ? (src.irpfImp ?? '') : '', '', '',
        src ? (src.irpfPct ?? '') : '', '', '',
        `Clave repetida en destino (filas ${rowsArr.join(', ')}). NO se escribe hasta corregir.`
      ));
    });

    // ---- Actualizar fila a fila ----
    // Regla: solo se escribe en una celda destino si está VACÍA.
    // Si ya tiene cualquier valor, se omite esa celda (campo a campo, independientemente).
    let updated = 0, same = 0, notFound = 0, blocked = 0, omitted = 0;

    for (let i = 0; i < tgtValues.length; i++) {
      const row = tgtValues[i];
      const empresaT = (row[TGT_EMPRESA_COL - 1] || '').toString().trim();
      const numT = (row[TGT_NUM_COL - 1] || '').toString().trim();
      if (!empresaT || !numT) continue;

      const key = makeKey_(empresaT, numT);
      const destRowNum = i + 2;

      if (tgtDupes.has(key)) { blocked++; continue; }

      const src = srcMap.get(key);
      if (!src) continue;

      const beforeBruto = row[TGT_BRUTO_COL - 1];
      const beforeImp   = row[TGT_IRPF_IMP_COL - 1];
      const beforePct   = row[TGT_IRPF_PCT_COL - 1];

      // Por campo: ¿hay dato en origen Y destino está vacío? → se puede escribir
      // Por campo: ¿hay dato en origen Y destino ya tiene valor? → se omite
      const canWriteBruto = hasValue_(src.bruto)    && !hasValue_(beforeBruto);
      const canWriteImp   = hasValue_(src.irpfImp)  && !hasValue_(beforeImp);
      const canWritePct   = hasValue_(src.irpfPct)  && !hasValue_(beforePct);

      const omitBruto = hasValue_(src.bruto)   && hasValue_(beforeBruto);
      const omitImp   = hasValue_(src.irpfImp) && hasValue_(beforeImp);
      const omitPct   = hasValue_(src.irpfPct) && hasValue_(beforePct);

      const afterBruto = canWriteBruto ? src.bruto    : beforeBruto;
      const afterImp   = canWriteImp   ? src.irpfImp  : beforeImp;
      const afterPct   = canWritePct   ? src.irpfPct  : beforePct;

      const anyWritten = canWriteBruto || canWriteImp || canWritePct;
      const anyOmitted = omitBruto || omitImp || omitPct;

      if (anyWritten) {
        if (canWriteBruto) tgtSheet.getRange(destRowNum, TGT_BRUTO_COL).setValue(src.bruto);
        if (canWriteImp)   tgtSheet.getRange(destRowNum, TGT_IRPF_IMP_COL).setValue(src.irpfImp);
        if (canWritePct)   tgtSheet.getRange(destRowNum, TGT_IRPF_PCT_COL).setValue(src.irpfPct);

        const omitDetail = anyOmitted
          ? ' | Omitidos (celda con valor): ' + [omitBruto ? 'Bruto (L)' : null, omitImp ? 'IRPF € (R)' : null, omitPct ? 'IRPF % (S)' : null].filter(Boolean).join(', ')
          : '';

        updated++;
        logRows.push(makeLogRow_(
          now, runId, 'ACTUALIZADO',
          src.empresaDisp, src.numDisp, src.srcRow, destRowNum,
          src.bruto ?? '', beforeBruto, afterBruto,
          src.irpfImp ?? '', beforeImp, afterImp,
          src.irpfPct ?? '', beforePct, afterPct,
          omitDetail
        ));
      } else if (anyOmitted) {
        const omitFields = [omitBruto ? 'Bruto (L)' : null, omitImp ? 'IRPF € (R)' : null, omitPct ? 'IRPF % (S)' : null].filter(Boolean).join(', ');
        omitted++;
        logRows.push(makeLogRow_(
          now, runId, 'OMITIDO_CELDA_CON_VALOR',
          src.empresaDisp, src.numDisp, src.srcRow, destRowNum,
          src.bruto ?? '', beforeBruto, beforeBruto,
          src.irpfImp ?? '', beforeImp, beforeImp,
          src.irpfPct ?? '', beforePct, beforePct,
          `Celda(s) con valor existente, no se sobrescriben: ${omitFields}.`
        ));
      } else {
        same++;
        logRows.push(makeLogRow_(
          now, runId, 'SIN_CAMBIO',
          src.empresaDisp, src.numDisp, src.srcRow, destRowNum,
          src.bruto ?? '', beforeBruto, afterBruto,
          src.irpfImp ?? '', beforeImp, afterImp,
          src.irpfPct ?? '', beforePct, afterPct,
          ''
        ));
      }
    }

    // Origen que NO existe en destino
    srcMap.forEach((src, key) => {
      if (tgtDupes.has(key)) return;
      if (!tgtRowByKey.has(key)) {
        notFound++;
        logRows.push(makeLogRow_(
          now, runId, 'NO_ENCONTRADO_DESTINO',
          src.empresaDisp, src.numDisp, src.srcRow, '',
          src.bruto ?? '', '', '',
          src.irpfImp ?? '', '', '',
          src.irpfPct ?? '', '', '',
          'No existe clave (empresa+nº) en destino.'
        ));
      }
    });

    logRows.push(makeLogRow_(
      now, runId, 'RESUMEN', '', '', '', '',
      '', '', '',
      '', '', '',
      '', '', '',
      `Actualizados=${updated} | SinCambio=${same} | OmitidosCeldaConValor=${omitted} | NoEncontrados=${notFound} | BloqueadosDuplicadoDestino=${blocked} | DuplicadosOrigen=${srcDupes.size} | DuplicadosDestino=${tgtDupes.size}`
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

function makeLogRow_(ts, runId, tipo, empresa, num, filaOrigen, filaDestino,
                     brutoOri, brutoAntes, brutoDesp,
                     irpfImpOri, irpfImpAntes, irpfImpDesp,
                     irpfPctOri, irpfPctAntes, irpfPctDesp,
                     detalle) {
  return [
    ts, runId, tipo,
    empresa || '', num || '',
    filaOrigen || '', filaDestino || '',
    brutoOri ?? '', brutoAntes ?? '', brutoDesp ?? '',
    irpfImpOri ?? '', irpfImpAntes ?? '', irpfImpDesp ?? '',
    irpfPctOri ?? '', irpfPctAntes ?? '', irpfPctDesp ?? '',
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