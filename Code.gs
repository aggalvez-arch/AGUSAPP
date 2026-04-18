/**
 * AGUSAPP - Motor de análisis semiautomático con revisión por excepción.
 *
 * Hojas esperadas:
 * - RAW
 * - ANALISIS
 * - REVISION
 * - CONFIG
 */

const SHEETS = Object.freeze({
  RAW: 'RAW',
  ANALISIS: 'ANALISIS',
  REVISION: 'REVISION',
  CONFIG: 'CONFIG'
});

const STATUS = Object.freeze({
  NUEVO: 'NUEVO',
  EN_ANALISIS: 'EN_ANALISIS',
  AUTO_APROBADO: 'AUTO_APROBADO',
  REQUIERE_REVISION: 'REQUIERE_REVISION',
  APROBADO_MANUAL: 'APROBADO_MANUAL',
  RECHAZADO_MANUAL: 'RECHAZADO_MANUAL',
  ERROR: 'ERROR'
});

const DECISION = Object.freeze({
  AUTO_APROBAR: 'AUTO_APROBAR',
  REVISAR: 'REVISAR',
  BLOQUEAR: 'BLOQUEAR'
});

const RAW_HEADERS = [
  'id',
  'timestamp',
  'monto',
  'pais',
  'categoria',
  'descripcion',
  'origen',
  'status',
  'ultima_actualizacion',
  'error'
];

const ANALISIS_HEADERS = [
  'id',
  'timestamp',
  'monto',
  'pais',
  'categoria',
  'descripcion',
  'origen',
  'puntaje_confianza',
  'decision_sistema',
  'motivo_decision',
  'reglas_activadas',
  'requiere_revision_humana',
  'status_final',
  'procesado_en'
];

const REVISION_HEADERS = [
  'id',
  'timestamp',
  'monto',
  'pais',
  'categoria',
  'descripcion',
  'origen',
  'puntaje_confianza',
  'decision_sistema',
  'motivo_decision',
  'accion_humana',
  'revisor',
  'comentario_revision',
  'status_final',
  'ultima_actualizacion'
];

const CONFIG_HEADERS = [
  'clave',
  'valor',
  'descripcion'
];

const DEFAULT_CONFIG = [
  ['umbral_auto_aprobacion', '85', 'Puntaje mínimo para auto aprobar'],
  ['umbral_revision', '60', 'Puntaje mínimo para enviar a revisión'],
  ['monto_maximo_auto', '1000', 'Monto máximo permitido para auto aprobación'],
  ['paises_permitidos', 'AR,CL,MX,UY,PE,CO,US,ES', 'Lista de países permitidos (CSV)'],
  ['categorias_sensibles', 'crypto,casino,adulto', 'Categorías sensibles (CSV)'],
  ['origenes_confiables', 'portal_web,api,erp', 'Orígenes más confiables (CSV)'],
  ['limite_lote', '500', 'Cantidad máxima de filas a procesar por ejecución'],
  ['solo_procesar_nuevos', 'true', 'true para procesar solo status NUEVO']
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Procesamiento')
    .addItem('Inicializar estructura', 'setupProject')
    .addItem('Procesar todo (1 botón)', 'procesarTodo')
    .addItem('Aplicar decisiones manuales', 'aplicarDecisionesManuales')
    .addToUi();
}

/**
 * Crea/valida las hojas y encabezados mínimos.
 */
function setupProject() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeaders_(ss, SHEETS.RAW, RAW_HEADERS);
  ensureSheetWithHeaders_(ss, SHEETS.ANALISIS, ANALISIS_HEADERS);
  ensureSheetWithHeaders_(ss, SHEETS.REVISION, REVISION_HEADERS);
  ensureSheetWithHeaders_(ss, SHEETS.CONFIG, CONFIG_HEADERS);

  seedDefaultConfig_(ss.getSheetByName(SHEETS.CONFIG));

  SpreadsheetApp.getUi().alert('Estructura inicializada/validada correctamente.');
}

/**
 * Botón principal: procesa nuevos registros, decide automático y separa excepciones.
 */
function procesarTodo() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = getRequiredSheet_(ss, SHEETS.RAW);
    const analisisSheet = getRequiredSheet_(ss, SHEETS.ANALISIS);
    const revisionSheet = getRequiredSheet_(ss, SHEETS.REVISION);
    const configSheet = getRequiredSheet_(ss, SHEETS.CONFIG);

    const config = readConfig_(configSheet);
    const context = buildProcessingContext_(config);

    const rawData = getTableData_(rawSheet, RAW_HEADERS);
    const candidateRows = filterCandidateRows_(rawData.rows, context.soloProcesarNuevos, rawData.headerMap);
    const rowsToProcess = candidateRows.slice(0, context.limiteLote);

    if (!rowsToProcess.length) {
      SpreadsheetApp.getUi().alert('No hay filas para procesar.');
      return;
    }

    const now = new Date();
    const analisisRowsToAppend = [];
    const revisionRowsToAppend = [];
    const rawUpdates = [];

    rowsToProcess.forEach((item) => {
      const row = item.row;
      const rowIndex = item.rowIndex;

      try {
        const input = normalizeRawRow_(row, rawData.headerMap);
        const result = evaluarRegistro_(input, context);

        analisisRowsToAppend.push([
          input.id,
          input.timestamp,
          input.monto,
          input.pais,
          input.categoria,
          input.descripcion,
          input.origen,
          result.score,
          result.decision,
          result.reason,
          result.rules.join(' | '),
          result.requiereRevisionHumana ? 'SI' : 'NO',
          result.statusFinal,
          now
        ]);

        if (result.requiereRevisionHumana) {
          revisionRowsToAppend.push([
            input.id,
            input.timestamp,
            input.monto,
            input.pais,
            input.categoria,
            input.descripcion,
            input.origen,
            result.score,
            result.decision,
            result.reason,
            '',
            '',
            '',
            STATUS.REQUIERE_REVISION,
            now
          ]);
        }

        rawUpdates.push({
          rowIndex,
          status: result.statusFinal,
          ultimaActualizacion: now,
          error: ''
        });
      } catch (error) {
        rawUpdates.push({
          rowIndex,
          status: STATUS.ERROR,
          ultimaActualizacion: now,
          error: sanitizeError_(error)
        });
      }
    });

    appendRows_(analisisSheet, analisisRowsToAppend);
    appendRows_(revisionSheet, revisionRowsToAppend);
    applyRawUpdates_(rawSheet, rawData.headerMap, rawUpdates);

    SpreadsheetApp.getUi().alert(
      `Proceso completado. Analizados: ${rowsToProcess.length}. ` +
        `Auto-aprobados: ${rawUpdates.filter((u) => u.status === STATUS.AUTO_APROBADO).length}. ` +
        `A revisión: ${rawUpdates.filter((u) => u.status === STATUS.REQUIERE_REVISION).length}. ` +
        `Errores: ${rawUpdates.filter((u) => u.status === STATUS.ERROR).length}.`
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Aplica decisiones humanas ingresadas en REVISION.
 * accion_humana permitida: APROBAR o RECHAZAR.
 */
function aplicarDecisionesManuales() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const revisionSheet = getRequiredSheet_(ss, SHEETS.REVISION);
    const rawSheet = getRequiredSheet_(ss, SHEETS.RAW);

    const revisionData = getTableData_(revisionSheet, REVISION_HEADERS);
    const rawData = getTableData_(rawSheet, RAW_HEADERS);

    const now = new Date();
    const rawIdIndex = buildRawIdIndex_(rawData.rows, rawData.headerMap);
    const revisionUpdates = [];
    const rawUpdates = [];

    revisionData.rows.forEach((row, idx) => {
      const action = normalizeString_(row[revisionData.headerMap.accion_humana]);
      const statusFinal = normalizeString_(row[revisionData.headerMap.status_final]);
      const id = normalizeString_(row[revisionData.headerMap.id]);

      if (!id) return;
      if (statusFinal !== STATUS.REQUIERE_REVISION) return;
      if (!action) return;

      let nuevoStatus = '';
      if (action === 'APROBAR') nuevoStatus = STATUS.APROBADO_MANUAL;
      else if (action === 'RECHAZAR') nuevoStatus = STATUS.RECHAZADO_MANUAL;
      else return;

      revisionUpdates.push({
        rowIndex: idx + 2,
        statusFinal: nuevoStatus,
        ultimaActualizacion: now
      });

      const rawRowIndex = rawIdIndex[id];
      if (rawRowIndex) {
        rawUpdates.push({
          rowIndex: rawRowIndex,
          status: nuevoStatus,
          ultimaActualizacion: now,
          error: ''
        });
      }
    });

    applyRevisionUpdates_(revisionSheet, revisionData.headerMap, revisionUpdates);
    applyRawUpdates_(rawSheet, rawData.headerMap, rawUpdates);

    SpreadsheetApp.getUi().alert(`Decisiones manuales aplicadas: ${revisionUpdates.length}.`);
  } finally {
    lock.releaseLock();
  }
}

function evaluarRegistro_(input, context) {
  let score = 100;
  const rules = [];

  if (!input.id) {
    score -= 60;
    rules.push('ID_VACIO');
  }

  if (!input.timestamp || Object.prototype.toString.call(input.timestamp) !== '[object Date]') {
    score -= 30;
    rules.push('TIMESTAMP_INVALIDO');
  }

  if (typeof input.monto !== 'number' || Number.isNaN(input.monto) || input.monto < 0) {
    score -= 80;
    rules.push('MONTO_INVALIDO');
  } else {
    if (input.monto > context.montoMaximoAuto) {
      score -= 35;
      rules.push('MONTO_ALTO');
    }
    if (input.monto > context.montoMaximoAuto * 2) {
      score -= 20;
      rules.push('MONTO_MUY_ALTO');
    }
  }

  if (!context.paisesPermitidos.has(input.pais)) {
    score -= 25;
    rules.push('PAIS_NO_PERMITIDO');
  }

  if (context.categoriasSensibles.has(input.categoria)) {
    score -= 30;
    rules.push('CATEGORIA_SENSIBLE');
  }

  if (!context.origenesConfiables.has(input.origen)) {
    score -= 15;
    rules.push('ORIGEN_NO_CONFIABLE');
  }

  if (input.descripcion && input.descripcion.length < 5) {
    score -= 10;
    rules.push('DESCRIPCION_POBRE');
  }

  score = Math.max(0, Math.min(100, Math.round(score)));

  if (rules.includes('MONTO_INVALIDO') || rules.includes('ID_VACIO')) {
    return {
      score,
      decision: DECISION.BLOQUEAR,
      reason: 'Datos críticos inválidos.',
      rules,
      requiereRevisionHumana: true,
      statusFinal: STATUS.REQUIERE_REVISION
    };
  }

  if (score >= context.umbralAutoAprobacion && !rules.includes('CATEGORIA_SENSIBLE')) {
    return {
      score,
      decision: DECISION.AUTO_APROBAR,
      reason: 'Cumple criterios de confianza para aprobación automática.',
      rules,
      requiereRevisionHumana: false,
      statusFinal: STATUS.AUTO_APROBADO
    };
  }

  if (score >= context.umbralRevision) {
    return {
      score,
      decision: DECISION.REVISAR,
      reason: 'Cumple mínimo operativo pero requiere validación humana.',
      rules,
      requiereRevisionHumana: true,
      statusFinal: STATUS.REQUIERE_REVISION
    };
  }

  return {
    score,
    decision: DECISION.BLOQUEAR,
    reason: 'Bajo puntaje de confianza. Revisión obligatoria.',
    rules,
    requiereRevisionHumana: true,
    statusFinal: STATUS.REQUIERE_REVISION
  };
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const existingHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  const needsRewrite = headers.some((h, i) => normalizeString_(existingHeaders[i]) !== h);

  if (needsRewrite) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function seedDefaultConfig_(configSheet) {
  const lastRow = configSheet.getLastRow();
  const current = lastRow > 1 ? configSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(normalizeString_) : [];

  const toInsert = DEFAULT_CONFIG.filter((row) => !current.includes(row[0]));
  if (!toInsert.length) return;

  configSheet.getRange(configSheet.getLastRow() + 1, 1, toInsert.length, CONFIG_HEADERS.length).setValues(toInsert);
}

function getRequiredSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(`Falta la hoja requerida: ${name}. Ejecuta setupProject().`);
  }
  return sheet;
}

function getTableData_(sheet, expectedHeaders) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), expectedHeaders.length);

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeString_);
  const headerMap = {};
  expectedHeaders.forEach((h) => {
    const idx = headers.indexOf(h);
    if (idx === -1) {
      throw new Error(`La hoja ${sheet.getName()} no contiene el encabezado esperado: ${h}`);
    }
    headerMap[h] = idx;
  });

  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { rows, headerMap };
}

function filterCandidateRows_(rows, soloProcesarNuevos, headerMap) {
  if (!soloProcesarNuevos) {
    return rows.map((row, idx) => ({ row, rowIndex: idx + 2 }));
  }

  return rows
    .map((row, idx) => ({ row, rowIndex: idx + 2 }))
    .filter((item) => {
      const status = normalizeString_(item.row[headerMap.status]);
      return !status || status === STATUS.NUEVO || status === STATUS.ERROR;
    });
}

function normalizeRawRow_(row, headerMap) {
  const rawMonto = row[headerMap.monto];
  const monto = typeof rawMonto === 'number' ? rawMonto : Number(String(rawMonto).replace(',', '.'));

  return {
    id: normalizeString_(row[headerMap.id]),
    timestamp: row[headerMap.timestamp],
    monto,
    pais: normalizeString_(row[headerMap.pais]).toUpperCase(),
    categoria: normalizeString_(row[headerMap.categoria]).toLowerCase(),
    descripcion: normalizeString_(row[headerMap.descripcion]),
    origen: normalizeString_(row[headerMap.origen]).toLowerCase()
  };
}

function buildProcessingContext_(config) {
  return {
    umbralAutoAprobacion: getNumberConfig_(config, 'umbral_auto_aprobacion', 85),
    umbralRevision: getNumberConfig_(config, 'umbral_revision', 60),
    montoMaximoAuto: getNumberConfig_(config, 'monto_maximo_auto', 1000),
    paisesPermitidos: getCsvSetConfig_(config, 'paises_permitidos', true),
    categoriasSensibles: getCsvSetConfig_(config, 'categorias_sensibles', false),
    origenesConfiables: getCsvSetConfig_(config, 'origenes_confiables', false),
    limiteLote: Math.max(1, Math.round(getNumberConfig_(config, 'limite_lote', 500))),
    soloProcesarNuevos: getBooleanConfig_(config, 'solo_procesar_nuevos', true)
  };
}

function readConfig_(configSheet) {
  const lastRow = configSheet.getLastRow();
  if (lastRow <= 1) return {};

  const values = configSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return values.reduce((acc, row) => {
    const key = normalizeString_(row[0]);
    if (!key) return acc;
    acc[key] = row[1];
    return acc;
  }, {});
}

function getNumberConfig_(config, key, fallback) {
  const n = Number(config[key]);
  return Number.isFinite(n) ? n : fallback;
}

function getBooleanConfig_(config, key, fallback) {
  const value = normalizeString_(config[key]).toLowerCase();
  if (!value) return fallback;
  return value === 'true' || value === '1' || value === 'si';
}

function getCsvSetConfig_(config, key, uppercase, fallbackCsv) {
  const raw = normalizeString_(config[key] || fallbackCsv || '');
  if (!raw) return new Set();

  return new Set(
    raw
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean)
      .map((s) => (uppercase ? s.toUpperCase() : s.toLowerCase()))
  );
}

function appendRows_(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function applyRawUpdates_(rawSheet, headerMap, updates) {
  if (!updates.length) return;

  const statusCol = headerMap.status + 1;
  const updatedCol = headerMap.ultima_actualizacion + 1;
  const errorCol = headerMap.error + 1;
  const sorted = updates.slice().sort((a, b) => a.rowIndex - b.rowIndex);
  const startRow = sorted[0].rowIndex;
  const endRow = sorted[sorted.length - 1].rowIndex;
  const numRows = endRow - startRow + 1;

  const statusValues = Array.from({ length: numRows }, () => ['']);
  const updatedValues = Array.from({ length: numRows }, () => ['']);
  const errorValues = Array.from({ length: numRows }, () => ['']);

  sorted.forEach((u) => {
    const pos = u.rowIndex - startRow;
    statusValues[pos][0] = u.status;
    updatedValues[pos][0] = u.ultimaActualizacion;
    errorValues[pos][0] = u.error || '';
  });

  rawSheet.getRange(startRow, statusCol, numRows, 1).setValues(statusValues);
  rawSheet.getRange(startRow, updatedCol, numRows, 1).setValues(updatedValues);
  rawSheet.getRange(startRow, errorCol, numRows, 1).setValues(errorValues);
}

function applyRevisionUpdates_(revisionSheet, headerMap, updates) {
  if (!updates.length) return;

  const statusCol = headerMap.status_final + 1;
  const updatedCol = headerMap.ultima_actualizacion + 1;
  const sorted = updates.slice().sort((a, b) => a.rowIndex - b.rowIndex);
  const startRow = sorted[0].rowIndex;
  const endRow = sorted[sorted.length - 1].rowIndex;
  const numRows = endRow - startRow + 1;

  const statusValues = Array.from({ length: numRows }, () => ['']);
  const updatedValues = Array.from({ length: numRows }, () => ['']);

  sorted.forEach((u) => {
    const pos = u.rowIndex - startRow;
    statusValues[pos][0] = u.statusFinal;
    updatedValues[pos][0] = u.ultimaActualizacion;
  });

  revisionSheet.getRange(startRow, statusCol, numRows, 1).setValues(statusValues);
  revisionSheet.getRange(startRow, updatedCol, numRows, 1).setValues(updatedValues);
}

function buildRawIdIndex_(rawRows, headerMap) {
  return rawRows.reduce((acc, row, idx) => {
    const id = normalizeString_(row[headerMap.id]);
    if (id) acc[id] = idx + 2;
    return acc;
  }, {});
}

function normalizeString_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function sanitizeError_(error) {
  if (!error) return 'Error desconocido';
  const msg = error && error.message ? error.message : String(error);
  return msg.length > 500 ? msg.slice(0, 500) : msg;
}
