/*************************************************
 * PROYECTO UPG - VERSION REESCRITA Y ORDENADA
 *************************************************/

/* =========================
   MENÚ
   ========================= */

function doGet(e) {
  return HtmlService.createHtmlOutput("¡Hola, mundo!");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestión de Datos')
    .addItem('Display', 'upgDisplay')
    .addItem('Buscar Nombres', 'upgBuscarNombres')
    .addItem('Buscar Duplicados', 'upgBuscarDuplicidades')
    .addItem('Insertar Números', 'upgInsertarNumeroDesdeColumnaC')
    .addItem('Status Final', 'upgEstatusFinal')
    .addToUi();
}

/* =========================
   DISPLAY
   ========================= */

function upgDisplay() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila === 0) return;

  const valores = hoja.getRange(1, 1, ultimaFila, 1).getValues();
  const resultados = [];

  for (let i = 0; i < valores.length; i++) {
    const valor = valores[i][0];

    if (!valor) {
      resultados.push(['']);
      continue;
    }

    const texto = valor.toString();
    const indice = texto.indexOf('‡');
    let textoHasta = indice !== -1 ? texto.substring(0, indice) : texto;

    if (textoHasta.endsWith('^EI^E')) {
      resultados.push([textoHasta]);
    } else {
      const limpio = textoHasta.replace(/(\^)+$/, '');
      resultados.push([limpio + '^EI^E']);
    }
  }

  hoja.getRange(1, 2, resultados.length, 1).setValues(resultados);
}

/* =========================
   BUSCAR NOMBRES
   ========================= */

function upgBuscarNombres() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila === 0) return;

  const data = hoja.getRange(1, 1, ultimaFila, 1).getValues().flat();

  hoja.getRange(1, 2, hoja.getMaxRows(), 3).clearContent().setBackground(null);

  const resultados = [];
  const palabrasComunes = ['angel', 'joy', 'passenger', 'list', 'doc'];

  for (let i = 0; i < data.length; i++) {
    const line = upgSafeTrim(data[i]);

    if (!line || !line.startsWith("G-")) continue;

    const gLine = line;
    const nombreCompleto = upgExtraerNombreDesdeLineaG(gLine);

    if (!nombreCompleto) {
      resultados.push([gLine, "", "SIN COINCIDENCIA"]);
      continue;
    }

    let mejorCoincidencia = {
      nombre: "",
      numero: "",
      similitud: 0
    };

    let estadoEspecial = "";
    const similitudMinima = nombreCompleto.length < 5 ? 0.9 : 0.8;

    for (let j = i + 1; j < data.length; j++) {
      const nextLine = upgSafeTrim(data[j]);

      if (!nextLine) continue;

      if (nextLine.startsWith("IGD")) {
        break;
      }

      if (nextLine.includes("VERIFY TRAVEL DOCUMENT DATA, THEN ADD DOCV")) {
        estadoEspecial = "VERIFY TRAVEL DOCUMENT DATA, THEN ADD DOCV";
        break;
      }

      if (nextLine.startsWith("‡") || nextLine.includes("TICKETLESS")) {
        mejorCoincidencia = {
          nombre: nombreCompleto,
          numero: "1",
          similitud: 1
        };
        break;
      }

      if (!/^\d/.test(nextLine)) continue;

      const numero = upgExtraerNumero(nextLine);
      const nombreEncontrado = upgExtraerNombreDesdeLineaNumerada(nextLine);

      if (!nombreEncontrado) continue;

      const nombreEncontradoNormalizado = upgNormalizarNombre(nombreEncontrado);

      if (palabrasComunes.includes(nombreEncontradoNormalizado)) {
        continue;
      }

      const similitud = upgCompararSimilitudNombreCompleto(nombreCompleto, nombreEncontrado);

      if (upgNormalizarNombre(nombreCompleto) === nombreEncontradoNormalizado) {
        mejorCoincidencia = {
          nombre: nombreEncontrado,
          numero: numero,
          similitud: 1
        };
        break;
      }

      if (similitud >= similitudMinima && similitud > mejorCoincidencia.similitud) {
        mejorCoincidencia = {
          nombre: nombreEncontrado,
          numero: numero,
          similitud: similitud
        };
      }
    }

    let valorColumnaC = "";
    let estado = "SIN COINCIDENCIA";

    if (estadoEspecial === "VERIFY TRAVEL DOCUMENT DATA, THEN ADD DOCV") {
      valorColumnaC = "";
      estado = "VERIFY TRAVEL DOCUMENT DATA, THEN ADD DOCV";
    } else if (mejorCoincidencia.similitud >= similitudMinima) {
      valorColumnaC = `${mejorCoincidencia.numero} ${mejorCoincidencia.nombre}`.trim();
      estado = upgValidarRelacionNombre(nombreCompleto, mejorCoincidencia.nombre);
    }

    resultados.push([gLine, valorColumnaC, estado]);
  }

  if (resultados.length === 0) return;

  hoja.getRange(1, 2, resultados.length, 3).setValues(resultados);

  for (let i = 0; i < resultados.length; i++) {
    const estado = resultados[i][2];
    let color = "#f4cccc";

    if (estado === "OK") {
      color = "#d9ead3";
    } else if (estado === "REVISAR") {
      color = "#fff2cc";
    } else if (estado === "VERIFY TRAVEL DOCUMENT DATA, THEN ADD DOCV") {
      color = "#cfe2f3";
    }

    hoja.getRange(i + 1, 2, 1, 3).setBackground(color);
  }
}

/* =========================
   BUSCAR DUPLICIDADES
   ========================= */

function upgBuscarDuplicidades() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila === 0) return;

  const data = hoja.getRange(1, 1, ultimaFila, 1).getValues().flat();
  const resultados = [];

  hoja.getRange(1, 5, hoja.getMaxRows(), 3).clearContent().setBackground(null);

  for (let i = 0; i < data.length; i++) {
    const line = upgSafeTrim(data[i]);

    if (!line || !line.startsWith("G-")) continue;

    const gLine = line;
    const nombresPorClave = {};
    const lineasPorClave = {};

    for (let j = i + 1; j < data.length; j++) {
      const nextLine = upgSafeTrim(data[j]);

      if (!nextLine) continue;

      if (nextLine.startsWith("IGD")) {
        break;
      }

      if (!/^\d/.test(nextLine)) continue;

      const nombreEncontrado = upgExtraerNombreDesdeLineaNumerada(nextLine);
      if (!nombreEncontrado) continue;

      const nombreBase = upgConstruirNombreBaseDuplicidad(nombreEncontrado);
      if (!nombreBase) continue;

      if (!nombresPorClave[nombreBase]) {
        nombresPorClave[nombreBase] = 0;
        lineasPorClave[nombreBase] = [];
      }

      nombresPorClave[nombreBase]++;
      lineasPorClave[nombreBase].push(nextLine);
    }

    const clavesDuplicadas = Object.keys(nombresPorClave).filter(clave => nombresPorClave[clave] > 1);

    if (clavesDuplicadas.length > 0) {
      let lineasDuplicadas = [];

      clavesDuplicadas.forEach(clave => {
        lineasDuplicadas = lineasDuplicadas.concat(lineasPorClave[clave]);
      });

      resultados.push([
        gLine,
        lineasDuplicadas.join("\n"),
        "NOMBRE DUPLICADO"
      ]);
    } else {
      resultados.push([
        gLine,
        "No se encontraron duplicidades",
        ""
      ]);
    }
  }

  if (resultados.length === 0) return;

  hoja.getRange(1, 5, resultados.length, 3).setValues(resultados);

  for (let i = 0; i < resultados.length; i++) {
    if (resultados[i][2] === "NOMBRE DUPLICADO") {
      hoja.getRange(i + 1, 5, 1, 3).setBackground("#f4cccc");
    }
  }
}

/* =========================
   INSERTAR NÚMERO DESDE C
   ========================= */

function upgInsertarNumeroDesdeColumnaC() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila === 0) return;

  const dataB = hoja.getRange(1, 2, ultimaFila, 1).getValues();
  const dataC = hoja.getRange(1, 3, ultimaFila, 1).getValues();

  const salidaB = [];

  for (let i = 0; i < ultimaFila; i++) {
    const celdaB = dataB[i][0];
    const celdaC = dataC[i][0];

    if (!celdaB || !celdaC) {
      salidaB.push([celdaB || ""]);
      continue;
    }

    const matchNumero = celdaC.toString().match(/^\d+/);

    if (!matchNumero) {
      salidaB.push([celdaB]);
      continue;
    }

    const numero = matchNumero[0];
    const textoB = celdaB.toString();

    const partes = textoB.split('^');
    const parteAntesCaret = partes[0];
    const parteDespuesCaret = partes.slice(1).join('^');

    const modificadaDespuesCaret = parteDespuesCaret
      .replace(/\bEG-(?!\d)/g, 'EG-' + numero)
      .replace(/\bKG(?!\d|[-])/g, 'KG' + numero);

    const resultadoFinal = parteAntesCaret + '^' + modificadaDespuesCaret;
    salidaB.push([resultadoFinal]);
  }

  hoja.getRange(1, 2, salidaB.length, 1).setValues(salidaB);
}

/* =========================
   STATUS FINAL
   ========================= */

function upgEstatusFinal() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila === 0) return;

  const datosA = hoja.getRange(1, 1, ultimaFila, 1).getValues().flat();
  const datosF = hoja.getRange(2, 6, Math.max(ultimaFila - 1, 1), 1).getValues();
  const datosG = hoja.getRange(2, 7, Math.max(ultimaFila - 1, 1), 1).getValues();

  const erroresF = datosF
    .map(row => row[0])
    .filter(v => v !== "" && v !== null && v !== undefined);

  const reemplazosG = datosG.map(row => row[0]);

  const resultadosB = [];
  const resultadosC = [];
  const patronLinea = /^G-.*(J80|J20|J18)$/;

  for (let i = 0; i < datosA.length; i++) {
    const texto = datosA[i];

    if (!texto || !patronLinea.test(texto.toString())) continue;

    resultadosB.push([texto]);

    let encontrado = false;

    for (let j = i + 1; j < datosA.length && !encontrado; j++) {
      const textoDebajo = datosA[j];

      if (typeof textoDebajo !== 'string' || textoDebajo.trim() === '') {
        continue;
      }

      for (let k = 0; k < erroresF.length; k++) {
        if (textoDebajo.includes(erroresF[k])) {
          resultadosC.push([textoDebajo]);
          encontrado = true;
          break;
        }
      }
    }

    if (!encontrado) {
      resultadosC.push([""]);
    }
  }

  hoja.getRange(1, 2, hoja.getMaxRows(), 2).clearContent();

  if (resultadosB.length > 0) {
    hoja.getRange(1, 2, resultadosB.length, 1).setValues(resultadosB);
  }

  if (resultadosC.length > 0) {
    hoja.getRange(1, 3, resultadosC.length, 1).setValues(resultadosC);
  }

  const datosCActuales = resultadosC.length > 0
    ? hoja.getRange(1, 3, resultadosC.length, 1).getValues()
    : [];

  for (let i = 0; i < datosCActuales.length; i++) {
    const valorC = datosCActuales[i][0];

    for (let j = 0; j < erroresF.length; j++) {
      if (valorC === erroresF[j]) {
        datosCActuales[i][0] = reemplazosG[j];
        break;
      }
    }
  }

  if (datosCActuales.length > 0) {
    hoja.getRange(1, 3, datosCActuales.length, 1).setValues(datosCActuales);
  }
}

/* =========================
   VALIDACIÓN DE NOMBRES
   ========================= */

function upgValidarRelacionNombre(nombreOriginal, nombreEncontrado) {
  const original = upgNormalizarNombre(nombreOriginal).split(" ").filter(Boolean);
  const encontrado = upgNormalizarNombre(nombreEncontrado).split(" ").filter(Boolean);

  if (original.length === 0 || encontrado.length === 0) {
    return "SIN COINCIDENCIA";
  }

  const palabrasCompartidas = original.filter(p => encontrado.includes(p));

  if (original.join(" ") === encontrado.join(" ")) {
    return "OK";
  }

  if (palabrasCompartidas.length >= 2) {
    return "REVISAR";
  }

  if (palabrasCompartidas.length === 1) {
    let mejorSimilitud = 0;

    for (let i = 0; i < original.length; i++) {
      for (let j = 0; j < encontrado.length; j++) {
        const sim = upgJaroWinkler(original[i], encontrado[j]);
        if (sim > mejorSimilitud) {
          mejorSimilitud = sim;
        }
      }
    }

    if (mejorSimilitud >= 0.93) {
      return "REVISAR";
    }

    return "SIN RELACION";
  }

  let coincidenciasFuertes = 0;

  for (let i = 0; i < original.length; i++) {
    for (let j = 0; j < encontrado.length; j++) {
      const sim = upgJaroWinkler(original[i], encontrado[j]);
      if (sim >= 0.93) {
        coincidenciasFuertes++;
        break;
      }
    }
  }

  if (coincidenciasFuertes >= 2) {
    return "REVISAR";
  }

  return "SIN RELACION";
}

/* =========================
   EXTRACCIÓN / LIMPIEZA
   ========================= */

function upgSafeTrim(valor) {
  return typeof valor === 'string' ? valor.trim() : null;
}

function upgExtraerNumero(linea) {
  const match = linea.toString().match(/^\d+/);
  return match ? match[0] : "";
}

function upgExtraerNombreDesdeLineaG(gLine) {
  if (!gLine) return null;

  const partes = gLine.split("-");
  if (partes.length < 3) return null;

  let nombre = partes.slice(2).join("-");

  if (nombre.indexOf("‡") !== -1) {
    nombre = nombre.substring(0, nombre.indexOf("‡"));
  }

  if (nombre.indexOf("^") !== -1) {
    nombre = nombre.substring(0, nombre.indexOf("^"));
  }

  nombre = nombre
    .replace(/[^\p{L}\s]/gu, " ")
    .replace(/\s+/g, " ")
    .trim();

  return nombre || null;
}

function upgExtraerNombreDesdeLineaNumerada(linea) {
  if (!linea) return null;

  const sinNumero = linea.toString().replace(/^\d+\s*/, '');

  const limpio = sinNumero
    .replace(/[^\p{L}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return limpio || null;
}

function upgConstruirNombreBaseDuplicidad(nombreCompleto) {
  const partes = upgNormalizarNombre(nombreCompleto).split(" ").filter(Boolean);
  if (partes.length < 2) return "";

  const tokensLimpios = partes.filter(function(token) {
    return !upgEsTokenNoNombre(token);
  });

  if (tokensLimpios.length < 2) return "";

  return tokensLimpios.join(" ");
}

function upgEsTokenNoNombre(token) {
  const blacklist = [
    'gru', 'cwb', 'poa', 'for', 'ccp', 'anf', 'lsc', 'clo', 'bog', 'smr',
    'nb', 'aci', 'sl', 'ff', 'plt', 'gld', 'blk', 'glp', 'sig', 'prch',
    'docs', 'eti', 'et', 'ae', 'ob', 'bf', 'bg', 'ak', 'af', 'q', 'o',
    'm', 'v', 'c', 'f', 'i'
  ];

  if (!token) return true;
  if (blacklist.includes(token)) return true;
  if (/^\d+[a-z]*$/i.test(token)) return true;
  if (/^[a-z]\d+[a-z]*$/i.test(token)) return true;
  if (/^\d+[a-z]+\*?$/i.test(token)) return true;
  if (/^[a-z]+\d+\*?$/i.test(token)) return true;
  if (token.length === 1) return true;

  return false;
}

function upgNormalizarNombre(nombre) {
  return (nombre || "")
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z\s]/g, "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ");
}

/* =========================
   SIMILITUD
   ========================= */

function upgCompararSimilitudNombreCompleto(nombre1, nombre2) {
  const partes1 = upgNormalizarNombre(nombre1).split(' ').filter(Boolean);
  const partes2 = upgNormalizarNombre(nombre2).split(' ').filter(Boolean);

  if (!partes1.length || !partes2.length) return 0;

  if (nombre1.length < 5 || nombre2.length < 5) {
    return upgJaroWinkler(upgNormalizarNombre(nombre1), upgNormalizarNombre(nombre2));
  }

  let puntajeTotal = 0;
  let comparaciones = 0;

  partes1.forEach(parte1 => {
    let mejorSimilitud = 0;

    partes2.forEach(parte2 => {
      const similitud = upgJaroWinkler(parte1, parte2);
      if (similitud > mejorSimilitud) {
        mejorSimilitud = similitud;
      }
    });

    puntajeTotal += mejorSimilitud;
    comparaciones++;
  });

  return comparaciones > 0 ? puntajeTotal / comparaciones : 0;
}

/* =========================
   JARO-WINKLER
   ========================= */

function upgJaroWinkler(s1, s2) {
  const jaro = upgJaroDistance(s1, s2);
  const prefixLength = upgPrefixLength(s1, s2);
  return jaro + (0.1 * prefixLength * (1 - jaro));
}

function upgJaroDistance(s1, s2) {
  if (!s1 || !s2) return 0;
  if (s1 === s2) return 1;

  const matchWindow = Math.floor(Math.max(s1.length, s2.length) / 2) - 1;
  const matches1 = new Array(s1.length).fill(false);
  const matches2 = new Array(s2.length).fill(false);
  let matches = 0;

  for (let i = 0; i < s1.length; i++) {
    const start = Math.max(0, i - matchWindow);
    const end = Math.min(i + matchWindow + 1, s2.length);

    for (let j = start; j < end; j++) {
      if (matches2[j]) continue;
      if (s1[i] !== s2[j]) continue;

      matches1[i] = true;
      matches2[j] = true;
      matches++;
      break;
    }
  }

  if (matches === 0) return 0;

  let t = 0;
  let k = 0;

  for (let i = 0; i < s1.length; i++) {
    if (!matches1[i]) continue;
    while (!matches2[k]) k++;
    if (s1[i] !== s2[k]) t++;
    k++;
  }

  return (matches / s1.length + matches / s2.length + (matches - t / 2) / matches) / 3;
}

function upgPrefixLength(s1, s2) {
  let prefixLength = 0;
  const maxPrefixLength = 4;

  for (let i = 0; i < Math.min(s1.length, s2.length); i++) {
    if (s1[i] !== s2[i]) break;
    prefixLength++;
  }

  return Math.min(prefixLength, maxPrefixLength);
}