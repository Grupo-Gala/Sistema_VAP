/****************************************************
 * VAP - SINCRONIZAR SECRETARIAS EN FORMULARIO REAL
 ****************************************************/

const FORM_ID = "1qIDdGFsy9oGD-cfuE_X_bS4Pj555hHJaQWszBTtFbyI";
const SPREADSHEET_ID = "1HQMxubwPv-vrslsXj5DvsB0gJ9OyewVNDmbPg1hjmHk";
const SHEET_SECRETARIAS = "VAP_Secretarias";
const QUESTION_TITLE = "SECRETARIA";
const RESPONSE_ID_HEADER = "ID_Secre";

/****************************************************
 * 1) SINCRONIZACIÓN PRINCIPAL DEL FORMULARIO
 ****************************************************/
function actualizarPreguntaSecretarias() {
  try {
    const opciones = obtenerSecretariasActivasUnicas_();

    if (opciones.length === 0) {
      throw new Error("No se han encontrado secretarias activas válidas");
    }

    const form = FormApp.openById(FORM_ID);
    const items = form.getItems(FormApp.ItemType.LIST);
    const item = items.find(i => i.getTitle() === QUESTION_TITLE);

    if (!item) {
      throw new Error(`No se encuentra la pregunta tipo lista con título "${QUESTION_TITLE}"`);
    }

    item.asListItem().setChoiceValues(opciones);

    Logger.log(`Formulario actualizado correctamente. Opciones cargadas: ${opciones.length}`);
  } catch (error) {
    Logger.log("ERROR actualizarPreguntaSecretarias: " + error.message);
    throw error;
  }
}

/****************************************************
 * 2) DISPARADOR POR EDICIÓN
 ****************************************************/
function onSpreadsheetEditTrigger(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (!sheet || sheet.getName() !== SHEET_SECRETARIAS) return;

    if (!puedeEjecutarSincronizacion_()) return;

    Logger.log("Cambio detectado en VAP_Secretarias (edición). Sincronizando formulario...");
    actualizarPreguntaSecretarias();

  } catch (error) {
    Logger.log("ERROR onSpreadsheetEditTrigger: " + error.message);
    throw error;
  }
}

/****************************************************
 * 3) DISPARADOR POR CAMBIOS ESTRUCTURALES
 ****************************************************/
function onSpreadsheetChangeTrigger(e) {
  try {
    if (!e) return;

    const tiposRelevantes = ["INSERT_ROW", "REMOVE_ROW", "INSERT_GRID", "REMOVE_GRID"];
    if (!tiposRelevantes.includes(e.changeType)) return;

    if (!puedeEjecutarSincronizacion_()) return;

    Logger.log(`Cambio estructural detectado (${e.changeType}). Sincronizando formulario...`);
    actualizarPreguntaSecretarias();

  } catch (error) {
    Logger.log("ERROR onSpreadsheetChangeTrigger: " + error.message);
    throw error;
  }
}

/****************************************************
 * 4) AL ENVIAR EL FORMULARIO
 * Crea / rellena ID_SECRETARIA
 ****************************************************/
function onFormSubmit(e) {
  try {
    Logger.log("1. Inicio onFormSubmit");

    const secretariaNombre = obtenerRespuestaPorTitulo_(e, QUESTION_TITLE);
    if (!secretariaNombre) {
      throw new Error(`No se encontró respuesta para la pregunta "${QUESTION_TITLE}"`);
    }
    Logger.log(`2. Secretaria recibida: ${secretariaNombre}`);

    const idSecre = buscarIdSecrePorNombre_(secretariaNombre);
    if (!idSecre) {
      throw new Error(`No se encontró ID_Secre para "${secretariaNombre}"`);
    }
    Logger.log(`3. ID encontrado: ${idSecre}`);

    const responseSheet = obtenerHojaRespuestasDelFormulario_();
    Logger.log(`4. Hoja de respuestas encontrada: ${responseSheet.getName()}`);

    const targetRow = localizarFilaRespuesta_(responseSheet, e.response);
    Logger.log(`5. Fila objetivo: ${targetRow}`);

    const lastCol = responseSheet.getLastColumn();
    const headers = responseSheet.getRange(1, 1, 1, lastCol).getValues()[0];

    let colId = headers.indexOf(RESPONSE_ID_HEADER) + 1;

    if (colId === 0) {
      colId = lastCol + 1;
      responseSheet.getRange(1, colId).setValue(RESPONSE_ID_HEADER);
      Logger.log(`6. Columna ${RESPONSE_ID_HEADER} creada en posición ${colId}`);
    } else {
      Logger.log(`6. Columna ${RESPONSE_ID_HEADER} ya existe en posición ${colId}`);
    }

    responseSheet.getRange(targetRow, colId).setValue(idSecre);
    Logger.log(`7. ${RESPONSE_ID_HEADER} escrito correctamente en fila ${targetRow}`);

  } catch (error) {
    Logger.log("ERROR onFormSubmit: " + error.message);
    throw error;
  }
}

/****************************************************
 * 5) OBTENER LISTA DE SECRETARIAS ACTIVAS
 ****************************************************/
function obtenerSecretariasActivasUnicas_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SECRETARIAS);

  if (!sheet) {
    throw new Error(`No existe la hoja "${SHEET_SECRETARIAS}"`);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return [];

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];

  const colId = headers.indexOf("ID_Secre");
  const colNombre = headers.indexOf("Nombre_Secretaria");
  const colAlta = headers.indexOf("Alta");

  if (colId === -1) throw new Error('No se encuentra la columna "ID_Secre"');
  if (colNombre === -1) throw new Error('No se encuentra la columna "Nombre_Secretaria"');
  if (colAlta === -1) throw new Error('No se encuentra la columna "Alta"');

  const nombresBase = data
    .slice(1)
    .filter(row => {
      const id = String(row[colId] || "").trim();
      const nombre = String(row[colNombre] || "").trim();
      const alta = row[colAlta] === true;
      return id && nombre && alta;
    })
    .map(row => String(row[colNombre] || "").trim().replace(/\s+/g, " "));

  return [...new Set(nombresBase)].sort((a, b) =>
    a.localeCompare(b, "es", { sensitivity: "base" })
  );
}

/****************************************************
 * 6) BUSCAR ID POR NOMBRE
 ****************************************************/
function buscarIdSecrePorNombre_(nombreBuscado) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SECRETARIAS);

  if (!sheet) {
    throw new Error(`No existe la hoja "${SHEET_SECRETARIAS}"`);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];

  const colId = headers.indexOf("ID_Secre");
  const colNombre = headers.indexOf("Nombre_Secretaria");
  const colAlta = headers.indexOf("Alta");

  if (colId === -1) throw new Error('No se encuentra la columna "ID_Secre"');
  if (colNombre === -1) throw new Error('No se encuentra la columna "Nombre_Secretaria"');
  if (colAlta === -1) throw new Error('No se encuentra la columna "Alta"');

  const nombreNormalizado = normalizarTexto_(nombreBuscado);

  const coincidencias = data.slice(1).filter(row => {
    const nombre = normalizarTexto_(row[colNombre]);
    const estaActiva = row[colAlta] === true;
    return nombre === nombreNormalizado && estaActiva;
  });

  if (coincidencias.length === 0) return "";
  if (coincidencias.length > 1) {
    throw new Error(`Hay varias secretarias activas con el mismo nombre: "${nombreBuscado}"`);
  }

  return String(coincidencias[0][colId] || "").trim();
}

/****************************************************
 * 7) HELPERS FORMULARIO / RESPUESTAS
 ****************************************************/
function obtenerRespuestaPorTitulo_(e, questionTitle) {
  const itemResponses = e.response.getItemResponses();

  for (let i = 0; i < itemResponses.length; i++) {
    const item = itemResponses[i].getItem();
    const title = item.getTitle();
    const response = itemResponses[i].getResponse();

    if (title === questionTitle) {
      return String(response || "").trim();
    }
  }

  return "";
}

function obtenerHojaRespuestasDelFormulario_() {
  const form = FormApp.openById(FORM_ID);
  const destinationId = form.getDestinationId();

  if (!destinationId) {
    throw new Error("El formulario no tiene hoja de respuestas vinculada");
  }

  const ss = SpreadsheetApp.openById(destinationId);
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) continue;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    if (headers.includes(QUESTION_TITLE) && headers.includes("Marca temporal")) {
      return sheet;
    }
  }

  throw new Error(`No se encontró la hoja de respuestas que contiene "${QUESTION_TITLE}"`);
}

function localizarFilaRespuesta_(responseSheet, formResponse) {
  const timestamp = formResponse.getTimestamp();
  const tz = Session.getScriptTimeZone();

  const timestampStr = Utilities.formatDate(timestamp, tz, "dd/MM/yyyy HH:mm:ss");

  const lastRow = responseSheet.getLastRow();
  const data = responseSheet.getRange(2, 1, Math.max(lastRow - 1, 0), 1).getDisplayValues();

  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === timestampStr) {
      return i + 2;
    }
  }

  return lastRow;
}

function normalizarTexto_(texto) {
  return String(texto || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

/****************************************************
 * 8) CONTROL DE FRECUENCIA
 ****************************************************/
function puedeEjecutarSincronizacion_() {
  const cache = CacheService.getScriptCache();
  const key = "vap_sync_secretarias_lock";
  const locked = cache.get(key);

  if (locked) {
    Logger.log("Sincronización omitida por control de frecuencia");
    return false;
  }

  cache.put(key, "1", 30);
  return true;
}

/****************************************************
 * 9) CREACIÓN DE TRIGGERS
 ****************************************************/
function crearTriggersProduccion() {
  borrarTriggers();

  const form = FormApp.openById(FORM_ID);

  ScriptApp.newTrigger("onFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("onSpreadsheetEditTrigger")
    .forSpreadsheet(SPREADSHEET_ID)
    .onEdit()
    .create();

  ScriptApp.newTrigger("onSpreadsheetChangeTrigger")
    .forSpreadsheet(SPREADSHEET_ID)
    .onChange()
    .create();

  Logger.log("Triggers de producción creados correctamente");
}

function crearTriggerHorarioRespaldo() {
  ScriptApp.newTrigger("actualizarPreguntaSecretarias")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("Trigger horario de respaldo creado correctamente");
}

function borrarTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log("Todos los triggers han sido eliminados");
}

function verTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    Logger.log(`Función: ${t.getHandlerFunction()} | Evento: ${t.getEventType()}`);
  });
}