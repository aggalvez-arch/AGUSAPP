/*************************************************
 * PROYECTO UPG - ORQUESTADOR POR FASES
 * Compatible con Google Apps Script
 *************************************************/

var UPG = {
  SHEETS: {
    RAW_TXN: 'RAW_TXN',
    DISPLAY_OUT: 'DISPLAY_OUT',
    LOG1_RAW: 'LOG1_RAW',
    LOG1_ANALISIS: 'LOG1_ANALISIS',
    SABRE_PASS2: 'SABRE_PASS2',
    LOG2_RAW: 'LOG2_RAW',
    FINAL: 'FINAL',
    CONFIG: 'CONFIG',
    TRACE: 'TRACE'
  },
  STAGE: {
    PREP_DISPLAY: 1,
    PROCESS_LOG1: 2,
    BUILD_PASS2: 3,
    PROCESS_LOG2: 4,
    FINAL_RESULT: 5
  },
  STATUS: {
    OK: 'OK',
    REVIEW: 'REVISAR',
    NO_MATCH: 'SIN COINCIDENCIA',
    NO_REL: 'SIN RELACION',
    DOCV: 'VERIFY TRAVEL DOCUMENT DATA, THEN ADD DOCV',
    DUP: 'NOMBRE DUPLICADO',
    READY: 'LISTO'
  }
};

function doGet() {
  return HtmlService.createHtmlOutput('UPG por fases listo. Usa el menú UPG.');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('UPG')
    .addItem('Inicializar estructura', 'upgSetup')
    .addSeparator()
    .addItem('Ejecutar etapa actual (CONFIG)', 'upgRunCurrentStage')
    .addSeparator()
    .addItem('Etapa 1 - Preparar Display', 'upgStage1PrepareDisplay')
    .addItem('Etapa 2 - Procesar Log 1', 'upgStage2ProcessLog1')
    .addItem('Etapa 3 - Generar salida para Sabre (pasada 2)', 'upgStage3BuildPass2')
    .addItem('Etapa 4 - Resultado final equivalente (Log 2)', 'upgStage4ProcessLog2')
    .addItem('Etapa 5 - Cierre', 'upgStage5Finalize')
    .addItem('Validar equivalencia final', 'upgValidateFinalEquivalenceFromSheets')
    .addToUi();
}

function upgSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  upgEnsureStructure(ss);
  upgTrace('SETUP', 'Estructura inicializada', 0);
  SpreadsheetApp.getActive().toast('Estructura lista. Usa RAW_TXN/LOG1_RAW/LOG2_RAW según etapa.');
}

function upgRunCurrentStage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  upgEnsureStructure(ss);

  var cfg = upgReadConfig(ss.getSheetByName(UPG.SHEETS.CONFIG));
  var stage = Number(cfg.currentStage || 1);

  if (stage === UPG.STAGE.PREP_DISPLAY) return upgStage1PrepareDisplay();
  if (stage === UPG.STAGE.PROCESS_LOG1) return upgStage2ProcessLog1();
  if (stage === UPG.STAGE.BUILD_PASS2) return upgStage3BuildPass2();
  if (stage === UPG.STAGE.PROCESS_LOG2) return upgStage4ProcessLog2();
  if (stage === UPG.STAGE.FINAL_RESULT) return upgStage5Finalize();

  SpreadsheetApp.getActive().toast('CURRENT_STAGE inválido en CONFIG.');
}

/* =========================
   ETAPA 1
   ========================= */

function upgStage1PrepareDisplay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  upgEnsureStructure(ss);

  var rawSheet = ss.getSheetByName(UPG.SHEETS.RAW_TXN);
  var outSheet = ss.getSheetByName(UPG.SHEETS.DISPLAY_OUT);
  var cfgSheet = ss.getSheetByName(UPG.SHEETS.CONFIG);
  var cfg = upgReadConfig(cfgSheet);

  var last = rawSheet.getLastRow();
  if (last < 2) {
    upgClearDataRows(outSheet, 4);
    upgTrace('ETAPA_1', 'RAW_TXN vacío', 0);
    SpreadsheetApp.getActive().toast('RAW_TXN no tiene transacciones.');
    return;
  }

  var data = rawSheet.getRange(2, 1, last - 1, 1).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var raw = upgSafeTrim(data[i][0]);
    if (!raw) continue;

    rows.push([
      raw,
      upgBuildDisplay(raw),
      UPG.STATUS.READY,
      'Copiar DISPLAY_CMD a Sabre y pegar respuesta en LOG1_RAW'
    ]);
  }

  upgClearDataRows(outSheet, 4);
  if (rows.length) outSheet.getRange(2, 1, rows.length, 4).setValues(rows);

  upgSetCurrentStage(cfgSheet, cfg.autoAdvance ? UPG.STAGE.PROCESS_LOG1 : UPG.STAGE.PREP_DISPLAY);
  upgTrace('ETAPA_1', 'Display preparado', rows.length);
  SpreadsheetApp.getActive().toast('Etapa 1 lista: ' + rows.length + ' comandos.');
}

/* =========================
   ETAPA 2
   ========================= */

function upgStage2ProcessLog1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  upgEnsureStructure(ss);

  var cfgSheet = ss.getSheetByName(UPG.SHEETS.CONFIG);
  var cfg = upgReadConfig(cfgSheet);
  var log1Sheet = ss.getSheetByName(UPG.SHEETS.LOG1_RAW);
  var analysisSheet = ss.getSheetByName(UPG.SHEETS.LOG1_ANALISIS);

  var lines = upgReadSingleColumn(log1Sheet);
  if (!lines.length) {
    upgClearDataRows(analysisSheet, 14);
    upgTrace('ETAPA_2', 'LOG1_RAW vacío', 0);
    SpreadsheetApp.getActive().toast('Pega el primer .log en LOG1_RAW antes de etapa 2.');
    return;
  }

  var blocks = upgBuildBlocks(lines);
  var rows = upgAnalyzeBlocksForLog1(blocks, cfg);

  upgClearDataRows(analysisSheet, 14);
  if (rows.length) analysisSheet.getRange(2, 1, rows.length, 14).setValues(rows);
  upgPaintByStatus(analysisSheet, rows, 2, 14, 5);

  upgSetCurrentStage(cfgSheet, cfg.autoAdvance ? UPG.STAGE.BUILD_PASS2 : UPG.STAGE.PROCESS_LOG1);
  upgTrace('ETAPA_2', 'Log 1 procesado', rows.length);
  SpreadsheetApp.getActive().toast('Etapa 2 lista: ' + rows.length + ' bloques.');
}

/* =========================
   ETAPA 3
   ========================= */

function upgStage3BuildPass2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  upgEnsureStructure(ss);

  var cfgSheet = ss.getSheetByName(UPG.SHEETS.CONFIG);
  var cfg = upgReadConfig(cfgSheet);
  var analysisSheet = ss.getSheetByName(UPG.SHEETS.LOG1_ANALISIS);
  var pass2Sheet = ss.getSheetByName(UPG.SHEETS.SABRE_PASS2);

  var data = upgReadRangeData(analysisSheet, 2, 1, 14);
  if (!data.length) {
    upgClearDataRows(pass2Sheet, 5);
    upgTrace('ETAPA_3', 'LOG1_ANALISIS sin datos', 0);
    SpreadsheetApp.getActive().toast('Corre etapa 2 primero.');
    return;
  }

  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var status = data[i][5];
    if (status === UPG.STATUS.OK || status === UPG.STATUS.REVIEW) {
      rows.push([
        data[i][0],
        data[i][8],
        status === UPG.STATUS.OK ? UPG.STATUS.READY : UPG.STATUS.REVIEW,
        'Enviar COMANDO_SABRE_P2 a Sabre y pegar respuesta en LOG2_RAW',
        data[i][13]
      ]);
    }
  }

  upgClearDataRows(pass2Sheet, 5);
  if (rows.length) pass2Sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  upgSetCurrentStage(cfgSheet, cfg.autoAdvance ? UPG.STAGE.PROCESS_LOG2 : UPG.STAGE.BUILD_PASS2);
  upgTrace('ETAPA_3', 'Salida pasada 2 generada', rows.length);
  SpreadsheetApp.getActive().toast('Etapa 3 lista: ' + rows.length + ' comandos para Sabre.');
}

/* =========================
   ETAPA 4 (EQUIVALENTE FUNCIONAL A upgEstatusFinal ORIGINAL)
   ========================= */

function upgStage4ProcessLog2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  upgEnsureStructure(ss);

  var cfgSheet = ss.getSheetByName(UPG.SHEETS.CONFIG);
  var cfg = upgReadConfig(cfgSheet);
  var log2Sheet = ss.getSheetByName(UPG.SHEETS.LOG2_RAW);
  var finalSheet = ss.getSheetByName(UPG.SHEETS.FINAL);

  var lines = upgReadSingleColumn(log2Sheet);
  if (!lines.length) {
    upgClearDataRows(finalSheet, 6);
    upgTrace('ETAPA_4', 'LOG2_RAW vacío', 0);
    SpreadsheetApp.getActive().toast('Pega el segundo .log en LOG2_RAW antes de etapa 4.');
    return;
  }

  var mappings = upgReadFinalMappings(cfg);

  // Refactor estructural (equivalente esperado)
  var refactorResult = upgComputeFinalStatusRefactor(lines, mappings.findList, mappings.replaceList);
  var rows = upgBuildFinalSheetRows(refactorResult);

  upgClearDataRows(finalSheet, 6);
  if (rows.length) finalSheet.getRange(2, 1, rows.length, 6).setValues(rows);
  upgPaintByStatus(finalSheet, rows, 2, 6, 3);

  // Validación de equivalencia contra la traducción literal del algoritmo original
  var originalResult = upgComputeFinalStatusOriginalLiteral(lines, mappings.findList, mappings.replaceList);
  var validation = upgCompareFinalResults(originalResult, refactorResult);
  upgTrace('ETAPA_4_VALIDACION', validation.message, validation.mismatchCount);

  upgSetCurrentStage(cfgSheet, cfg.autoAdvance ? UPG.STAGE.FINAL_RESULT : UPG.STAGE.PROCESS_LOG2);
  upgTrace('ETAPA_4', 'Resultado final calculado', rows.length);
  SpreadsheetApp.getActive().toast('Etapa 4 lista. Validación equivalencia: ' + validation.message);
}

function upgReadFinalMappings(cfg) {
  return {
    findList: upgCsvArray(cfg.finalFindList),
    replaceList: upgCsvArray(cfg.finalReplaceList)
  };
}

/**
 * Traducción literal de upgEstatusFinal original:
 * - Toma líneas G que cumplan /^G-.*(J80|J20|J18)$/
 * - Busca en líneas siguientes la primera que contenga cualquier token de erroresF
 * - Si no encuentra, vacío
 * - Luego reemplaza solo si hay igualdad exacta con erroresF[k] por reemplazosG[k]
 */
function upgComputeFinalStatusOriginalLiteral(datosA, erroresF, reemplazosG) {
  var resultadosB = [];
  var resultadosC = [];
  var patronLinea = /^G-.*(J80|J20|J18)$/;

  for (var i = 0; i < datosA.length; i++) {
    var texto = datosA[i];
    if (!texto || !patronLinea.test(String(texto))) continue;

    resultadosB.push(String(texto));
    var encontrado = false;

    for (var j = i + 1; j < datosA.length && !encontrado; j++) {
      var textoDebajo = datosA[j];
      if (typeof textoDebajo !== 'string' || textoDebajo.trim() === '') continue;

      for (var k = 0; k < erroresF.length; k++) {
        if (String(textoDebajo).indexOf(erroresF[k]) !== -1) {
          resultadosC.push(String(textoDebajo));
          encontrado = true;
          break;
        }
      }
    }

    if (!encontrado) resultadosC.push('');
  }

  var resultadosCReemplazados = resultadosC.slice();
  for (i = 0; i < resultadosCReemplazados.length; i++) {
    for (j = 0; j < erroresF.length; j++) {
      if (resultadosCReemplazados[i] === erroresF[j]) {
        resultadosCReemplazados[i] = reemplazosG[j] !== undefined ? reemplazosG[j] : resultadosCReemplazados[i];
        break;
      }
    }
  }

  return {
    gLines: resultadosB,
    rawMatches: resultadosC,
    finalMatches: resultadosCReemplazados
  };
}

/**
 * Misma lógica funcional, pero organizada para trazabilidad.
 */
function upgComputeFinalStatusRefactor(datosA, erroresF, reemplazosG) {
  var gLines = [];
  var rawMatches = [];
  var finalMatches = [];
  var patronLinea = /^G-.*(J80|J20|J18)$/;

  for (var i = 0; i < datosA.length; i++) {
    var line = datosA[i];
    if (!line || !patronLinea.test(String(line))) continue;

    gLines.push(String(line));

    var match = '';
    var found = false;

    for (var j = i + 1; j < datosA.length && !found; j++) {
      var candidate = datosA[j];
      if (typeof candidate !== 'string' || candidate.trim() === '') continue;

      for (var k = 0; k < erroresF.length; k++) {
        if (String(candidate).indexOf(erroresF[k]) !== -1) {
          match = String(candidate);
          found = true;
          break;
        }
      }
    }

    rawMatches.push(match);

    var replaced = match;
    for (k = 0; k < erroresF.length; k++) {
      if (replaced === erroresF[k]) {
        replaced = reemplazosG[k] !== undefined ? reemplazosG[k] : replaced;
        break;
      }
    }

    finalMatches.push(replaced);
  }

  return {
    gLines: gLines,
    rawMatches: rawMatches,
    finalMatches: finalMatches
  };
}

function upgBuildFinalSheetRows(result) {
  var rows = [];

  for (var i = 0; i < result.gLines.length; i++) {
    var rawMatch = result.rawMatches[i] || '';
    var finalMatch = result.finalMatches[i] || '';
    var status = finalMatch ? UPG.STATUS.OK : UPG.STATUS.REVIEW;

    var confidence = finalMatch ? 100 : 0;
    var motivoDecision = finalMatch
      ? (rawMatch !== finalMatch ? 'Reemplazo exacto aplicado según mapeo original.' : 'Coincidencia encontrada por criterio original.')
      : 'Sin coincidencia por criterio original.';

    rows.push([
      result.gLines[i],
      rawMatch,
      finalMatch,
      status,
      confidence,
      motivoDecision
    ]);
  }

  return rows;
}

function upgCompareFinalResults(originalResult, refactorResult) {
  var mismatches = 0;

  if (originalResult.gLines.length !== refactorResult.gLines.length) mismatches++;

  var len = Math.max(originalResult.gLines.length, refactorResult.gLines.length);
  for (var i = 0; i < len; i++) {
    if ((originalResult.gLines[i] || '') !== (refactorResult.gLines[i] || '')) mismatches++;
    if ((originalResult.rawMatches[i] || '') !== (refactorResult.rawMatches[i] || '')) mismatches++;
    if ((originalResult.finalMatches[i] || '') !== (refactorResult.finalMatches[i] || '')) mismatches++;
  }

  return {
    mismatchCount: mismatches,
    message: mismatches === 0 ? 'OK (0 diferencias)' : ('DIFERENCIAS=' + mismatches)
  };
}

function upgValidateFinalEquivalenceFromSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  upgEnsureStructure(ss);

  var cfg = upgReadConfig(ss.getSheetByName(UPG.SHEETS.CONFIG));
  var lines = upgReadSingleColumn(ss.getSheetByName(UPG.SHEETS.LOG2_RAW));
  var mappings = upgReadFinalMappings(cfg);

  var originalResult = upgComputeFinalStatusOriginalLiteral(lines, mappings.findList, mappings.replaceList);
  var refactorResult = upgComputeFinalStatusRefactor(lines, mappings.findList, mappings.replaceList);
  var validation = upgCompareFinalResults(originalResult, refactorResult);

  upgTrace('VALIDACION_EQUIVALENCIA', validation.message, validation.mismatchCount);
  SpreadsheetApp.getActive().toast('Validación equivalencia final: ' + validation.message);
}

/* =========================
   ETAPA 5
   ========================= */

function upgStage5Finalize() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  upgEnsureStructure(ss);

  var cfgSheet = ss.getSheetByName(UPG.SHEETS.CONFIG);
  var cfg = upgReadConfig(cfgSheet);
  var finalSheet = ss.getSheetByName(UPG.SHEETS.FINAL);

  var data = upgReadRangeData(finalSheet, 2, 1, 6);
  var okCount = 0;
  var reviewCount = 0;

  for (var i = 0; i < data.length; i++) {
    if (data[i][3] === UPG.STATUS.OK) okCount++;
    else reviewCount++;
  }

  upgSetCurrentStage(cfgSheet, cfg.autoAdvance ? UPG.STAGE.PREP_DISPLAY : UPG.STAGE.FINAL_RESULT);
  upgTrace('ETAPA_5', 'Cierre OK=' + okCount + ' REVISAR=' + reviewCount, data.length);
  SpreadsheetApp.getActive().toast('Etapa 5 finalizada: OK=' + okCount + ' | REVISAR=' + reviewCount);
}

/* =========================
   ANALISIS LOG1
   ========================= */

function upgAnalyzeBlocksForLog1(blocks, cfg) {
  var rows = [];

  for (var i = 0; i < blocks.length; i++) {
    var block = blocks[i];
    var gName = upgExtraerNombreDesdeLineaG(block.gLine);
    var display = upgBuildDisplay(block.gLine);
    var evalResult = upgEvaluateLog1Block(block.lines, gName, cfg);

    rows.push([
      block.gLine,
      gName,
      evalResult.numero,
      evalResult.nombre,
      evalResult.sim,
      evalResult.status,
      evalResult.dup ? 'SI' : 'NO',
      display,
      upgInsertNumberIntoDisplay(display, evalResult.numero),
      block.lines.join('\n'),
      evalResult.matchLine,
      evalResult.note,
      evalResult.confidencia,
      evalResult.motivoDecision
    ]);
  }

  return rows;
}

function upgEvaluateLog1Block(lines, gName, cfg) {
  if (!gName) {
    return {
      numero: '',
      nombre: '',
      sim: 0,
      status: UPG.STATUS.NO_MATCH,
      dup: false,
      matchLine: '',
      note: 'No se pudo extraer nombre G.',
      confianza: 0,
      motivoDecision: 'Nombre G vacío o inválido.'
    };
  }

  var minSim = gName.length < 5 ? cfg.minShort : cfg.minLong;
  var best = { numero: '', nombre: '', sim: 0, line: '' };
  var dups = {};

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    if (line.indexOf(cfg.keywordDocv) !== -1) {
      return {
        numero: '',
        nombre: '',
        sim: 0,
        status: UPG.STATUS.DOCV,
        dup: false,
        matchLine: line,
        note: 'Caso DOCV.',
        confianza: 100,
        motivoDecision: 'Keyword DOCV detectada en bloque.'
      };
    }

    if (line.indexOf('‡') === 0 || line.indexOf('TICKETLESS') !== -1) {
      return {
        numero: '1',
        nombre: gName,
        sim: 1,
        status: UPG.STATUS.OK,
        dup: false,
        matchLine: line,
        note: 'Shortcut ticketless/‡.',
        confianza: 98,
        motivoDecision: 'Regla directa ticketless/‡.'
      };
    }

    if (!/^\d/.test(line)) continue;

    var num = upgExtraerNumero(line);
    var foundName = upgExtraerNombreDesdeLineaNumerada(line);
    if (!foundName) continue;

    var norm = upgNormalizarNombre(foundName);
    if (cfg.commonWords[norm]) continue;

    var dupKey = upgConstruirNombreBaseDuplicidad(foundName, cfg.blacklist);
    if (dupKey) dups[dupKey] = (dups[dupKey] || 0) + 1;

    var sim = upgCompararSimilitudNombreCompleto(gName, foundName);
    if (upgNormalizarNombre(gName) === norm) {
      best = { numero: num, nombre: foundName, sim: 1, line: line };
      break;
    }

    if (sim >= minSim && sim > best.sim) {
      best = { numero: num, nombre: foundName, sim: sim, line: line };
    }
  }

  var dup = Object.keys(dups).some(function (k) { return dups[k] > 1; });

  if (best.sim >= minSim) {
    var relation = upgValidarRelacionNombre(gName, best.nombre, cfg.minWord);
    if (dup && relation === UPG.STATUS.OK) relation = UPG.STATUS.REVIEW;

    var confidence = Math.max(0, Math.min(100, Math.round(best.sim * 100)));
    if (relation === UPG.STATUS.REVIEW) confidence = Math.min(confidence, 79);

    return {
      numero: best.numero,
      nombre: best.nombre,
      sim: best.sim,
      status: relation,
      dup: dup,
      matchLine: best.line,
      note: dup ? UPG.STATUS.DUP : '',
      confianza: confidence,
      motivoDecision: relation === UPG.STATUS.OK
        ? 'Coincidencia válida sobre umbral.'
        : (dup ? 'Coincidencia con posibles duplicados.' : 'Coincidencia parcial, requiere revisión.')
    };
  }

  return {
    numero: '',
    nombre: '',
    sim: best.sim,
    status: UPG.STATUS.NO_MATCH,
    dup: dup,
    matchLine: '',
    note: 'No alcanzó umbral.',
    confianza: Math.max(0, Math.min(100, Math.round((best.sim || 0) * 100))),
    motivoDecision: 'No hubo match confiable según umbral configurado.'
  };
}

/* =========================
   ESTRUCTURA / CONFIG
   ========================= */

function upgEnsureStructure(ss) {
  var rawTxn = upgGetOrCreateSheet(ss, UPG.SHEETS.RAW_TXN);
  var display = upgGetOrCreateSheet(ss, UPG.SHEETS.DISPLAY_OUT);
  var log1 = upgGetOrCreateSheet(ss, UPG.SHEETS.LOG1_RAW);
  var log1Analysis = upgGetOrCreateSheet(ss, UPG.SHEETS.LOG1_ANALISIS);
  var pass2 = upgGetOrCreateSheet(ss, UPG.SHEETS.SABRE_PASS2);
  var log2 = upgGetOrCreateSheet(ss, UPG.SHEETS.LOG2_RAW);
  var finalSheet = upgGetOrCreateSheet(ss, UPG.SHEETS.FINAL);
  var config = upgGetOrCreateSheet(ss, UPG.SHEETS.CONFIG);
  var trace = upgGetOrCreateSheet(ss, UPG.SHEETS.TRACE);

  upgEnsureHeaders(rawTxn, [['RAW_TRANSACCION']]);
  upgEnsureHeaders(display, [['RAW_TRANSACCION', 'DISPLAY_CMD', 'ESTADO', 'ACCION_OPERADOR']]);
  upgEnsureHeaders(log1, [['LOG1_LINEA']]);
  upgEnsureHeaders(log1Analysis, [[
    'G_LINE', 'NOMBRE_G', 'MATCH_NUMERO', 'MATCH_NOMBRE', 'SIMILITUD', 'ESTADO', 'DUPLICADO',
    'DISPLAY_BASE', 'DISPLAY_CON_NUMERO', 'BLOQUE_LINEAS', 'MATCH_LINEA', 'NOTA', 'CONFIDENCIA', 'MOTIVO_DECISION'
  ]]);
  upgEnsureHeaders(pass2, [['G_LINE', 'COMANDO_SABRE_P2', 'ESTADO', 'ACCION_OPERADOR', 'NOTA']]);
  upgEnsureHeaders(log2, [['LOG2_LINEA']]);
  upgEnsureHeaders(finalSheet, [['G_LINE', 'MATCH_ORIGINAL', 'MATCH_FINAL', 'ESTADO_FINAL', 'CONFIDENCIA', 'MOTIVO_DECISION']]);
  upgEnsureHeaders(trace, [['TIMESTAMP', 'ETAPA', 'EVENTO', 'CANTIDAD', 'USUARIO']]);

  upgEnsureConfig(config);
}

function upgEnsureConfig(sheet) {
  var defaults = {
    CURRENT_STAGE: '1',
    AUTO_ADVANCE_STAGE: 'TRUE',
    MIN_SIMILITUD_LARGA: '0.80',
    MIN_SIMILITUD_CORTA: '0.90',
    MIN_WORD_SIM: '0.93',
    PALABRAS_COMUNES: 'angel,joy,passenger,list,doc',
    TOKENS_BLACKLIST: 'gru,cwb,poa,for,ccp,anf,lsc,clo,bog,smr,nb,aci,sl,ff,plt,gld,blk,glp,sig,prch,docs,eti,et,ae,ob,bf,bg,ak,af,q,o,m,v,c,f,i',
    KEYWORD_DOCV: UPG.STATUS.DOCV,
    FINAL_FIND_LIST: '',
    FINAL_REPLACE_LIST: ''
  };

  var existing = {};
  if (sheet.getLastRow() > 1) {
    var current = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    current.forEach(function (r) {
      var k = upgSafeTrim(r[0]);
      if (k) existing[k] = r[1];
    });
  }

  var rows = [['KEY', 'VALUE']];
  Object.keys(defaults).forEach(function (key) {
    rows.push([key, existing[key] !== undefined ? existing[key] : defaults[key]]);
  });

  sheet.clear();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.setFrozenRows(1);
}

function upgReadConfig(sheet) {
  var cfg = {};
  if (sheet.getLastRow() > 1) {
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    values.forEach(function (r) {
      var k = upgSafeTrim(r[0]);
      if (k) cfg[k] = r[1];
    });
  }

  return {
    currentStage: Number(cfg.CURRENT_STAGE || 1),
    autoAdvance: String(cfg.AUTO_ADVANCE_STAGE || 'TRUE').toUpperCase() === 'TRUE',
    minLong: Number(cfg.MIN_SIMILITUD_LARGA || 0.8),
    minShort: Number(cfg.MIN_SIMILITUD_CORTA || 0.9),
    minWord: Number(cfg.MIN_WORD_SIM || 0.93),
    commonWords: upgCsvSet(cfg.PALABRAS_COMUNES),
    blacklist: upgCsvSet(cfg.TOKENS_BLACKLIST),
    keywordDocv: String(cfg.KEYWORD_DOCV || UPG.STATUS.DOCV),
    finalFindList: cfg.FINAL_FIND_LIST || '',
    finalReplaceList: cfg.FINAL_REPLACE_LIST || ''
  };
}

function upgSetCurrentStage(configSheet, stage) {
  var data = configSheet.getRange(2, 1, Math.max(configSheet.getLastRow() - 1, 1), 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === 'CURRENT_STAGE') {
      configSheet.getRange(i + 2, 2).setValue(String(stage));
      return;
    }
  }
}

/* =========================
   UTILITARIOS SHEET / TRACE
   ========================= */

function upgGetOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function upgEnsureHeaders(sheet, headers2D) {
  var neededCols = headers2D[0].length;
  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  }

  var current = sheet.getRange(1, 1, 1, neededCols).getValues()[0].join('|');
  var target = headers2D[0].join('|');
  if (current !== target) {
    sheet.getRange(1, 1, 1, neededCols).setValues(headers2D);
  }
  sheet.setFrozenRows(1);
}

function upgReadSingleColumn(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];

  return sheet.getRange(2, 1, last - 1, 1).getValues().map(function (r) {
    return upgSafeTrim(r[0]);
  }).filter(Boolean);
}

function upgReadRangeData(sheet, startRow, startCol, width) {
  var last = sheet.getLastRow();
  if (last < startRow) return [];

  var rows = sheet.getRange(startRow, startCol, last - startRow + 1, width).getValues();
  return rows.filter(function (r) {
    return r.join('').toString().trim() !== '';
  });
}

function upgClearDataRows(sheet, width) {
  var maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 1, maxRows, width).clearContent().setBackground(null);
}

function upgPaintByStatus(sheet, rows, startRow, width, statusColIndex) {
  if (!rows.length) return;

  var colors = [];
  for (var i = 0; i < rows.length; i++) {
    var status = rows[i][statusColIndex];
    var color = '#f4cccc';
    if (status === UPG.STATUS.OK) color = '#d9ead3';
    else if (status === UPG.STATUS.REVIEW) color = '#fff2cc';
    else if (status === UPG.STATUS.DOCV) color = '#cfe2f3';

    var rowColors = [];
    for (var j = 0; j < width; j++) rowColors.push(color);
    colors.push(rowColors);
  }

  sheet.getRange(startRow, 1, rows.length, width).setBackgrounds(colors);
}

function upgTrace(stage, event, count) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UPG.SHEETS.TRACE);
  if (!sheet) return;

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 5).setValues([[
    new Date(),
    stage,
    event,
    String(count || 0),
    Session.getActiveUser().getEmail()
  ]]);
}

/* =========================
   PARSEO / MATCHING
   ========================= */

function upgBuildBlocks(lines) {
  var blocks = [];
  var i = 0;

  while (i < lines.length) {
    var line = lines[i];
    if (!line || line.indexOf('G-') !== 0) {
      i++;
      continue;
    }

    var block = { gLine: line, lines: [] };
    i++;

    while (i < lines.length) {
      var next = lines[i];
      if (!next) {
        i++;
        continue;
      }

      if (next.indexOf('IGD') === 0) {
        i++; // consume IGD
        break;
      }

      if (next.indexOf('G-') === 0) {
        break; // siguiente bloque, no consumir
      }

      block.lines.push(next);
      i++;
    }

    blocks.push(block);
  }

  return blocks;
}

function upgBuildDisplay(rawText) {
  if (!rawText) return '';
  var text = String(rawText);
  var idx = text.indexOf('‡');
  var until = idx !== -1 ? text.substring(0, idx) : text;
  if (until.slice(-5) === '^EI^E') return until;
  return until.replace(/(\^)+$/, '') + '^EI^E';
}

function upgInsertNumberIntoDisplay(displayText, numero) {
  if (!displayText || !numero) return displayText || '';

  var parts = String(displayText).split('^');
  var before = parts[0];
  var after = parts.slice(1).join('^');

  var patched = after
    .replace(/\bEG-(?!\d)/g, 'EG-' + numero)
    .replace(/\bKG(?!\d|[-])/g, 'KG' + numero);

  return before + '^' + patched;
}

function upgExtraerNumero(linea) {
  var m = String(linea || '').match(/^\d+/);
  return m ? m[0] : '';
}

function upgExtraerNombreDesdeLineaG(gLine) {
  if (!gLine) return '';

  var parts = String(gLine).split('-');
  if (parts.length < 3) return '';

  var name = parts.slice(2).join('-');
  if (name.indexOf('‡') !== -1) name = name.substring(0, name.indexOf('‡'));
  if (name.indexOf('^') !== -1) name = name.substring(0, name.indexOf('^'));

  return name
    .replace(/[^\p{L}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function upgExtraerNombreDesdeLineaNumerada(linea) {
  if (!linea) return '';

  var noNum = String(linea).replace(/^\d+\s*/, '');
  return noNum
    .replace(/[^\p{L}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function upgConstruirNombreBaseDuplicidad(nombreCompleto, blacklistSet) {
  var tokens = upgNormalizarNombre(nombreCompleto).split(' ').filter(Boolean);
  if (tokens.length < 2) return '';

  var clean = tokens.filter(function (token) {
    return !upgEsTokenNoNombre(token, blacklistSet);
  });

  return clean.length >= 2 ? clean.join(' ') : '';
}

function upgEsTokenNoNombre(token, blacklistSet) {
  if (!token) return true;
  if (blacklistSet && blacklistSet[token]) return true;
  if (/^\d+[a-z]*$/i.test(token)) return true;
  if (/^[a-z]\d+[a-z]*$/i.test(token)) return true;
  if (/^\d+[a-z]+\*?$/i.test(token)) return true;
  if (/^[a-z]+\d+\*?$/i.test(token)) return true;
  return token.length === 1;
}

function upgNormalizarNombre(nombre) {
  return String(nombre || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z\s]/g, '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ');
}

function upgValidarRelacionNombre(nombreOriginal, nombreEncontrado, minWordSimilarity) {
  var original = upgNormalizarNombre(nombreOriginal).split(' ').filter(Boolean);
  var found = upgNormalizarNombre(nombreEncontrado).split(' ').filter(Boolean);
  var minWord = minWordSimilarity || 0.93;

  if (!original.length || !found.length) return UPG.STATUS.NO_MATCH;
  if (original.join(' ') === found.join(' ')) return UPG.STATUS.OK;

  var common = original.filter(function (p) { return found.indexOf(p) !== -1; });
  if (common.length >= 2) return UPG.STATUS.REVIEW;

  var strong = 0;
  for (var i = 0; i < original.length; i++) {
    for (var j = 0; j < found.length; j++) {
      if (upgJaroWinkler(original[i], found[j]) >= minWord) {
        strong++;
        break;
      }
    }
  }

  return strong >= 2 ? UPG.STATUS.REVIEW : UPG.STATUS.NO_REL;
}

function upgCompararSimilitudNombreCompleto(nombre1, nombre2) {
  var p1 = upgNormalizarNombre(nombre1).split(' ').filter(Boolean);
  var p2 = upgNormalizarNombre(nombre2).split(' ').filter(Boolean);

  if (!p1.length || !p2.length) return 0;

  if (nombre1.length < 5 || nombre2.length < 5) {
    return upgJaroWinkler(upgNormalizarNombre(nombre1), upgNormalizarNombre(nombre2));
  }

  var total = 0;
  for (var i = 0; i < p1.length; i++) {
    var best = 0;
    for (var j = 0; j < p2.length; j++) {
      var sim = upgJaroWinkler(p1[i], p2[j]);
      if (sim > best) best = sim;
    }
    total += best;
  }

  return total / p1.length;
}

function upgJaroWinkler(s1, s2) {
  var j = upgJaroDistance(s1, s2);
  var p = upgPrefixLength(s1, s2);
  return j + (0.1 * p * (1 - j));
}

function upgJaroDistance(s1, s2) {
  if (!s1 || !s2) return 0;
  if (s1 === s2) return 1;

  var mw = Math.floor(Math.max(s1.length, s2.length) / 2) - 1;
  var m1 = new Array(s1.length).fill(false);
  var m2 = new Array(s2.length).fill(false);
  var matches = 0;

  for (var i = 0; i < s1.length; i++) {
    var start = Math.max(0, i - mw);
    var end = Math.min(i + mw + 1, s2.length);

    for (var j = start; j < end; j++) {
      if (m2[j]) continue;
      if (s1[i] !== s2[j]) continue;
      m1[i] = true;
      m2[j] = true;
      matches++;
      break;
    }
  }

  if (!matches) return 0;

  var t = 0;
  var k = 0;
  for (i = 0; i < s1.length; i++) {
    if (!m1[i]) continue;
    while (!m2[k]) k++;
    if (s1[i] !== s2[k]) t++;
    k++;
  }

  return (matches / s1.length + matches / s2.length + (matches - t / 2) / matches) / 3;
}

function upgPrefixLength(s1, s2) {
  var len = 0;
  var maxLen = Math.min(4, s1.length, s2.length);

  for (var i = 0; i < maxLen; i++) {
    if (s1[i] !== s2[i]) break;
    len++;
  }

  return len;
}

function upgCsvSet(csv) {
  var set = {};
  String(csv || '')
    .split(',')
    .map(function (x) { return upgNormalizarNombre(x); })
    .filter(Boolean)
    .forEach(function (x) { set[x] = true; });
  return set;
}

function upgCsvArray(csv) {
  return String(csv || '')
    .split(',')
    .map(function (x) { return upgSafeTrim(x); })
    .filter(Boolean);
}

function upgSafeTrim(value) {
  return typeof value === 'string' ? value.trim() : '';
}
