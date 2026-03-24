
/**
 * =========================================================
 *  EXTRA SECRETARIAS 2026 -> VINCULACIÓN GASTOS PERSONAL
 *  (Instalar en el Spreadsheet ORIGEN y trigger "On form submit")
 * =========================================================
 *
 * ORIGEN:
 *  - Spreadsheet: "EXTRA SECRETARIAS 2026 (respuestas)"
 *  - Hoja: "Respuestas de formulario"
 *  - Nombre empleado: Columna E
 *  - Nº HORAS EXTRA: Columna G
 *  - € OTROS EXTRAS: Columna I
 *
 * DESTINO:
 *  - Spreadsheet: "VINCULACIÓN GASTOS PERSONAL"
 *  - Hoja: "Hoja 1"
 *  - Nombre empleado: Columna C
 *  - Volcado:
 *      - Nº HORAS EXTRA -> Columna AE
 *      - € OTROS EXTRAS -> Columna AH
 *
 * Reglas:
 *  - Si un empleado aparece varias veces en el origen, vale el ÚLTIMO registro.
 *  - Si el empleado no existe en destino, crea una nueva fila.
 *  - Ignora tildes/diacríticos, mayúsculas/minúsculas y espacios dobles al comparar nombres.
 */

/** === IDs (OBLIGATORIO) === */
const SOURCE_SPREADSHEET_ID = '1fW2bl9cVnDrTQei26S2qi4iZhT2OZnp-kNeFmVX5m0w';
const TARGET_SPREADSHEET_ID = '1qK3JeBdJV6nHMnFYZGgsRAC0t3rfvig17c5E0iYo8kc';

// Hojas
const SOURCE_SHEET_NAME = 'Respuestas de formulario';
const TARGET_SHEET_NAME = 'Hoja 1';

// Columnas clave (1-based)
const SOURCE_NAME_COL = 5; // E en origen
const TARGET_NAME_COL = 3; // C en destino

// Origen -> Destino (1-based)
const SRC_HORAS_EXTRA_COL = 7;   // G: Nº HORAS EXTRA
const SRC_OTROS_EXTRAS_COL = 9;  // I: € OTROS EXTRAS
const DST_HORAS_EXTRA_COL = 31;  // AE
const DST_OTROS_EXTRAS_COL = 34; // AH

/** Trigger "Al enviar formulario" */
function onFormSubmit(e) {
  syncExtrasSecretariasToVinculacion();
}

/**
 * Sincroniza tomando el último registro por empleado y volcando en destino.
 * Si el empleado no existe en destino, crea fila nueva al final.
 */
function syncExtrasSecretariasToVinculacion() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    // ORIGEN
    const srcSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const srcSheet = srcSS.getSheetByName(SOURCE_SHEET_NAME);
    if (!srcSheet) throw new Error(`No existe la pestaña origen: ${SOURCE_SHEET_NAME}`);

    const srcLastRow = srcSheet.getLastRow();
    if (srcLastRow < 2) return; // sin datos

    const srcLastCol = srcSheet.getLastColumn();
    const srcValues = srcSheet.getRange(2, 1, srcLastRow - 1, srcLastCol).getValues();

    // 1) Último registro por empleado (si se repite, gana el último)
    const latestByName = new Map(); // key -> {displayName, horas, otros}
    for (let i = 0; i < srcValues.length; i++) {
      const row = srcValues[i];

      const nameRaw = row[SOURCE_NAME_COL - 1];
      const displayName = (nameRaw || '').toString().trim();
      if (!displayName) continue;

      const key = normalizeKey(displayName);

      latestByName.set(key, {
        displayName,
        horas: row[SRC_HORAS_EXTRA_COL - 1],
        otros: row[SRC_OTROS_EXTRAS_COL - 1],
      });
    }

    if (latestByName.size === 0) return;

    // DESTINO
    const tgtSS = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const tgtSheet = tgtSS.getSheetByName(TARGET_SHEET_NAME);
    if (!tgtSheet) throw new Error(`No existe la pestaña destino: ${TARGET_SHEET_NAME}`);

    let tgtLastRow = tgtSheet.getLastRow();
    if (tgtLastRow < 2) {
      tgtSheet.insertRowsAfter(1, 1);
      tgtLastRow = tgtSheet.getLastRow();
    }

    // 2) Índice destino por nombre (col C)
    const tgtNames = tgtSheet.getRange(2, TARGET_NAME_COL, tgtLastRow - 1, 1).getValues();
    const tgtRowByName = new Map(); // key -> rowNumber
    for (let r = 0; r < tgtNames.length; r++) {
      const nameRaw = tgtNames[r][0];
      const name = (nameRaw || '').toString().trim();
      if (!name) continue;
      tgtRowByName.set(normalizeKey(name), r + 2);
    }

    // 3) Actualizar / crear
    latestByName.forEach(({ displayName, horas, otros }, key) => {
      let tgtRow = tgtRowByName.get(key);

      if (!tgtRow) {
        tgtRow = tgtSheet.getLastRow() + 1;
        tgtSheet.insertRowAfter(tgtSheet.getLastRow());
        // Escribir nombre en col C
        tgtSheet.getRange(tgtRow, TARGET_NAME_COL).setValue(displayName);
        tgtRowByName.set(key, tgtRow);
      }

      // Escribir AE y AF
      tgtSheet.getRange(tgtRow, DST_HORAS_EXTRA_COL).setValue(horas);
      tgtSheet.getRange(tgtRow, DST_OTROS_EXTRAS_COL).setValue(otros);
    });

  } finally {
    lock.releaseLock();
  }
}

/**
 * Normaliza nombres (ignora tildes):
 * - trim
 * - colapsa espacios múltiples
 * - minúsculas
 * - elimina diacríticos
 */
function normalizeKey(name) {
  const cleaned = name.toString().trim().replace(/\s+/g, ' ').toLowerCase();
  return cleaned.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}