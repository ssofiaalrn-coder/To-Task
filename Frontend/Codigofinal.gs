const SPREADSHEET_ID = "1RcfcLUKPGtqV98h85HlhhkEUkVIPeQ9XhhWKcaFp6yI";
const HOJA_MATERIAS = "Materias";
const HOJA_TAREAS = "Tareas";
const HOJA_CATEGORIAS = "Categorias";
const HOJA_ESTADOS = "Estados";

const ESTADOS_TAREA = {
  PENDIENTE: "PENDIENTE",
  EN_PROCESO: "EN_PROCESO",
  COMPLETADA: "COMPLETADA",
  ATRASADA: "ATRASADA"
};


const GEMINI_MODELOS = ["gemini-3.8-flash", "gemini-3.7-flash", "gemini-3.6-flash"];
// Esperas entre reintentos (backoff) cuando hay saturación: 1s, 2.5s, 5s
const GEMINI_ESPERAS_MS = [1000, 2500, 5000];
const GEMINI_CODIGOS_REINTENTABLES = [429, 500, 502, 503, 504, 529];

function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile("index")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle("To-Task | Agenda y Tareas Universitarias");
  } catch (error) {
    return HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: Arial; text-align: center;">
        <h2 style="color: #dc3545;">Error del Sistema</h2>
        <p>No se pudo cargar la aplicación: ${error.message}</p>
        <button onclick="window.location.reload()">Reintentar</button>
      </div>
    `);
  }
}

function parsearFechaLocal(valorFecha) {
  if (!valorFecha) return null;
  if (valorFecha instanceof Date) {
    return isNaN(valorFecha.getTime()) ? null : valorFecha;
  }
  const texto = valorFecha.toString().trim();
  const partesISO = texto.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (partesISO) {
    return new Date(
      parseInt(partesISO[1], 10),
      parseInt(partesISO[2], 10) - 1,
      parseInt(partesISO[3], 10),
      12, 0, 0
    );
  }
  const fechaGenerica = new Date(texto);
  return isNaN(fechaGenerica.getTime()) ? null : fechaGenerica;
}

function normalizarEstadoEfectivo(estadoRaw, avanceNum, fechaEntregaDate) {
  let estado = (estadoRaw || ESTADOS_TAREA.PENDIENTE).toString().trim().toUpperCase();
  if (!Object.values(ESTADOS_TAREA).includes(estado)) {
    estado = ESTADOS_TAREA.PENDIENTE;
  }
  if (avanceNum >= 100) {
    return ESTADOS_TAREA.COMPLETADA;
  }
  if (estado !== ESTADOS_TAREA.COMPLETADA && fechaEntregaDate) {
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const entregaSinHora = new Date(fechaEntregaDate);
    entregaSinHora.setHours(0, 0, 0, 0);
    if (entregaSinHora < hoy) {
      return ESTADOS_TAREA.ATRASADA;
    }
  }
  return estado;
}

function registrarMateria(materia) {
  try {
    if (!materia || !materia.codigo || !materia.nombre) {
      return "Datos incompletos. El código de asignatura y el título son obligatorios.";
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_MATERIAS);

    if (!sheet) {
      throw new Error(`La hoja '${HOJA_MATERIAS}' no existe. Inicialice el sistema primero.`);
    }

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 6).setValues([[
        "Código", "Título / Materia", "Modalidad", "Categoría", "Prioridad (1-5)", "Fecha Creación"
      ]]);
      sheet.getRange(1, 1, 1, 6).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
    }

    const codigoNormalizado = materia.codigo.toString().trim().toUpperCase();
    const nombre = materia.nombre.toString().trim();
    const modalidad = (materia.modalidad || "Presencial").toString().trim();
    const categoria = (materia.categoria || "Tarea").toString().trim();
    const prioridadRaw = parseInt(materia.prioridad, 10);
    const prioridad = isNaN(prioridadRaw) ? 3 : Math.min(5, Math.max(1, prioridadRaw));

    if (nombre.length < 2) {
      return "El título o nombre debe tener al menos 2 caracteres.";
    }

    const datos = sheet.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][0].toString().trim().toUpperCase() === codigoNormalizado) {
        return "Ya existe un registro con este código.";
      }
    }

    sheet.appendRow([
      codigoNormalizado,
      nombre,
      modalidad,
      categoria,
      prioridad,
      new Date()
    ]);

    return "Asignatura o tarea base registrada correctamente.";
  } catch (error) {
    console.error("Error en registrarMateria:", error);
    return `Error al registrar: ${error.message}`;
  }
}

function registrarTarea(tarea) {
  try {
    if (!tarea || !tarea.codigo || !tarea.fechaEntrega || !tarea.estado) {
      return "Datos incompletos. Código, fecha de entrega y estado son obligatorios.";
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const matSheet = ss.getSheetByName(HOJA_MATERIAS);
    const tarSheet = ss.getSheetByName(HOJA_TAREAS);

    if (!matSheet || !tarSheet) {
      throw new Error("Las hojas del sistema no existen. Inicialice el sistema primero.");
    }

    if (tarSheet.getLastRow() === 0) {
      tarSheet.getRange(1, 1, 1, 8).setValues([[
        "Código", "Fecha Entrega", "Estado", "Avance (%)", "Usuario", "Timestamp", "Descripción / Notas", "Prioridad"
      ]]);
      tarSheet.getRange(1, 1, 1, 8).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
    }

    const codigoNormalizado = tarea.codigo.toString().trim().toUpperCase();
    const fechaEntrega = parsearFechaLocal(tarea.fechaEntrega);
    if (!fechaEntrega) {
      return "La fecha de entrega no tiene un formato válido.";
    }

    const estadoSolicitado = tarea.estado.toString().trim().toUpperCase();
    if (!Object.values(ESTADOS_TAREA).includes(estadoSolicitado)) {
      return `Estado de tarea inválido: ${estadoSolicitado}`;
    }

    const avanceParseado = parseFloat(tarea.avance);
    let avance = isNaN(avanceParseado) ? 0 : Math.min(100, Math.max(0, avanceParseado));
    if (estadoSolicitado === ESTADOS_TAREA.COMPLETADA && avance < 100) {
      avance = 100;
    }

    const materias = matSheet.getDataRange().getValues();
    let existeMateria = false;
    let prioridadBase = 3;

    for (let i = 1; i < materias.length; i++) {
      if (materias[i][0] && materias[i][0].toString().trim().toUpperCase() === codigoNormalizado) {
        existeMateria = true;
        const p = parseInt(materias[i][4], 10);
        prioridadBase = isNaN(p) ? 3 : Math.min(5, Math.max(1, p));
        break;
      }
    }

    if (!existeMateria) {
      return "El código de asignatura/tarea no existe. Regístrelo primero en el catálogo.";
    }

    const estadoFinal = normalizarEstadoEfectivo(estadoSolicitado, avance, fechaEntrega);
    const usuario = Session.getActiveUser().getEmail() || "Estudiante";
    const descripcion = (tarea.observaciones || tarea.descripcion || "").toString().trim();

    tarSheet.appendRow([
      codigoNormalizado,
      fechaEntrega,
      estadoFinal,
      avance,
      usuario,
      new Date(),
      descripcion,
      prioridadBase
    ]);

    return "Entrega / tarea registrada correctamente.";
  } catch (error) {
    console.error("Error en registrarTarea:", error);
    return `Error al registrar tarea: ${error.message}`;
  }
}

function actualizarEstadoTarea(filaHoja, nuevoEstado, nuevoAvance) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tarSheet = ss.getSheetByName(HOJA_TAREAS);
    if (!tarSheet) {
      return "No se encontró la hoja de tareas.";
    }

    const filaNum = parseInt(filaHoja, 10);
    if (isNaN(filaNum) || filaNum < 2 || filaNum > tarSheet.getLastRow()) {
      return "Fila de tarea no válida.";
    }

    const estadoNorm = (nuevoEstado || "").toString().trim().toUpperCase();
    if (!Object.values(ESTADOS_TAREA).includes(estadoNorm)) {
      return "Estado no permitido.";
    }

    let avance = parseFloat(nuevoAvance);
    if (isNaN(avance)) {
      avance = estadoNorm === ESTADOS_TAREA.COMPLETADA ? 100 : 50;
    }
    avance = Math.min(100, Math.max(0, avance));
    if (estadoNorm === ESTADOS_TAREA.COMPLETADA) {
      avance = 100;
    }

    const fechaEntrega = parsearFechaLocal(tarSheet.getRange(filaNum, 2).getValue());
    const estadoEfectivo = normalizarEstadoEfectivo(estadoNorm, avance, fechaEntrega);

    tarSheet.getRange(filaNum, 3).setValue(estadoEfectivo);
    tarSheet.getRange(filaNum, 4).setValue(avance);

    return "Estado de la tarea actualizado correctamente.";
  } catch (error) {
    console.error("Error en actualizarEstadoTarea:", error);
    return `Error al actualizar tarea: ${error.message}`;
  }
}

function buscarPorCodigo(codigo) {
  try {
    if (!codigo || codigo.toString().trim().length < 1) {
      return [];
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(HOJA_MATERIAS);
    if (!sheet) return [];

    const datos = sheet.getDataRange().getValues();
    if (datos.length <= 1) return [];

    const prefijo = codigo.toString().trim().toUpperCase();
    const encontrados = [];

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (fila[0] && fila[0].toString().trim().toUpperCase().startsWith(prefijo)) {
        encontrados.push({
          codigo: fila[0].toString(),
          nombre: fila[1] ? fila[1].toString() : "",
          modalidad: fila[2] ? fila[2].toString() : "Presencial",
          categoria: fila[3] ? fila[3].toString() : "Tarea"
        });
      }
    }

    return encontrados.slice(0, 10);
  } catch (error) {
    console.error("Error en buscarPorCodigo:", error);
    return [];
  }
}

function obtenerCalendarioEntregas(filtros) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const matSheet = ss.getSheetByName(HOJA_MATERIAS);
    const tarSheet = ss.getSheetByName(HOJA_TAREAS);

    if (!matSheet || !tarSheet) {
      return [];
    }

    const materias = matSheet.getDataRange().getValues();
    const tareas = tarSheet.getDataRange().getValues();
    if (tareas.length <= 1) {
      return [];
    }

    const mapaMaterias = {};
    for (let i = 1; i < materias.length; i++) {
      if (materias[i][0]) {
        const cod = materias[i][0].toString().trim().toUpperCase();
        mapaMaterias[cod] = {
          titulo: materias[i][1] ? materias[i][1].toString() : "Sin título",
          categoria: materias[i][3] ? materias[i][3].toString() : "Tarea",
          prioridad: parseInt(materias[i][4], 10) || 3
        };
      }
    }

    const filtroEstado = filtros && filtros.estado ? filtros.estado.toString().trim().toUpperCase() : "TODOS";
    const filtroTexto = filtros && filtros.texto ? filtros.texto.toString().trim().toLowerCase() : "";

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    const items = [];
    for (let i = 1; i < tareas.length; i++) {
      const fila = tareas[i];
      if (!fila[0]) continue;

      const codigo = fila[0].toString().trim().toUpperCase();
      const fechaEntrega = parsearFechaLocal(fila[1]);
      const avanceNum = parseFloat(fila[3]) || 0;
      const estadoEfectivo = normalizarEstadoEfectivo(fila[2], avanceNum, fechaEntrega);
      const infoMat = mapaMaterias[codigo] || { titulo: codigo, categoria: "General", prioridad: 3 };
      const prioridadNum = parseInt(fila[7], 10) || infoMat.prioridad;
      const descripcion = fila[6] ? fila[6].toString() : "";

      let diasRestantes = null;
      let fechaISO = "";
      let fechaTexto = "Sin fecha";
      if (fechaEntrega) {
        const entregaDia = new Date(fechaEntrega);
        entregaDia.setHours(0, 0, 0, 0);
        diasRestantes = Math.round((entregaDia.getTime() - hoy.getTime()) / (1000 * 60 * 60 * 24));
        fechaISO = Utilities.formatDate(fechaEntrega, Session.getScriptTimeZone(), "yyyy-MM-dd");
        fechaTexto = Utilities.formatDate(fechaEntrega, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }

      if (filtroEstado !== "TODOS" && estadoEfectivo !== filtroEstado) {
        continue;
      }

      if (filtroTexto) {
        const coincidencia =
          codigo.toLowerCase().includes(filtroTexto) ||
          infoMat.titulo.toLowerCase().includes(filtroTexto) ||
          infoMat.categoria.toLowerCase().includes(filtroTexto) ||
          descripcion.toLowerCase().includes(filtroTexto);
        if (!coincidencia) continue;
      }

      items.push({
        fila: i + 1,
        codigo: codigo,
        titulo: infoMat.titulo,
        categoria: infoMat.categoria,
        fechaEntrega: fechaTexto,
        fechaISO: fechaISO,
        diasRestantes: diasRestantes,
        estado: estadoEfectivo,
        avance: Math.round(avanceNum * 100) / 100,
        prioridad: prioridadNum,
        descripcion: descripcion
      });
    }

    return items.sort((a, b) => {
      if (a.fechaISO && b.fechaISO) {
        return a.fechaISO.localeCompare(b.fechaISO);
      }
      return b.prioridad - a.prioridad;
    });
  } catch (error) {
    console.error("Error en obtenerCalendarioEntregas:", error);
    return [];
  }
}

function obtenerResumenUniversitario() {
  try {
    const entregas = obtenerCalendarioEntregas({ estado: "TODOS" });
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const matSheet = ss.getSheetByName(HOJA_MATERIAS);
    const totalAsignaturas = matSheet ? Math.max(0, matSheet.getLastRow() - 1) : 0;

    let pendientes = 0;
    let enProceso = 0;
    let completadas = 0;
    let atrasadas = 0;
    let proximas7Dias = 0;
    let sumaAvance = 0;

    entregas.forEach(item => {
      sumaAvance += item.avance;
      if (item.estado === ESTADOS_TAREA.COMPLETADA) {
        completadas++;
      } else if (item.estado === ESTADOS_TAREA.ATRASADA) {
        atrasadas++;
      } else if (item.estado === ESTADOS_TAREA.EN_PROCESO) {
        enProceso++;
      } else {
        pendientes++;
      }

      if (
        item.estado !== ESTADOS_TAREA.COMPLETADA &&
        item.diasRestantes !== null &&
        item.diasRestantes >= 0 &&
        item.diasRestantes <= 7
      ) {
        proximas7Dias++;
      }
    });

    const avancePromedio = entregas.length > 0
      ? Math.round((sumaAvance / entregas.length) * 10) / 10
      : 0;

    return {
      totalTareas: entregas.length,
      totalAsignaturas: totalAsignaturas,
      pendientes: pendientes,
      enProceso: enProceso,
      completadas: completadas,
      atrasadas: atrasadas,
      proximas7Dias: proximas7Dias,
      avancePromedio: avancePromedio
    };
  } catch (error) {
    console.error("Error en obtenerResumenUniversitario:", error);
    return {
      totalTareas: 0,
      totalAsignaturas: 0,
      pendientes: 0,
      enProceso: 0,
      completadas: 0,
      atrasadas: 0,
      proximas7Dias: 0,
      avancePromedio: 0
    };
  }
}

function guardarApiKeyIA(apiKey) {
  const limpia = (apiKey || "").toString().trim();
  if (!limpia) {
    throw new Error("Debes proporcionar una API key válida.");
  }
  PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", limpia);
  return "API key guardada de forma segura en Script Properties.";
}

/**
 * Llama a Gemini con reintentos (backoff) y modelos de respaldo.
 * Devuelve { ok: true, texto, modelo } o { ok: false, error }.
 */
function llamarGemini(prompt, apiKey, generationConfig) {
  const payload = { contents: [{ role: "user", parts: [{ text: prompt }] }] };
  if (generationConfig) payload.generationConfig = generationConfig;

  const errores = [];
  let ultimoError = "Error desconocido";

  for (const modelo of GEMINI_MODELOS) {
    ultimoError = "Error desconocido";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelo}:generateContent`;
    const opciones = {
      method: "post",
      contentType: "application/json",
      headers: { "x-goog-api-key": apiKey }, // la clave va en el header, no en la URL
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    for (let intento = 0; intento <= GEMINI_ESPERAS_MS.length; intento++) {
      const hayMasIntentos = intento < GEMINI_ESPERAS_MS.length;

      let response;
      try {
        response = UrlFetchApp.fetch(url, opciones);
      } catch (e) {
        // Error de red / timeout de UrlFetch
        ultimoError = e.message;
        if (hayMasIntentos) {
          Utilities.sleep(GEMINI_ESPERAS_MS[intento]);
          continue;
        }
        break; // pasa al siguiente modelo
      }

      const codigo = response.getResponseCode();
      let data = null;
      try {
        data = JSON.parse(response.getContentText());
      } catch (e) {
        data = null;
      }

      if (codigo >= 200 && codigo < 300) {
        const bloqueo = data && data.promptFeedback && data.promptFeedback.blockReason;
        if (bloqueo) {
          return { ok: false, error: `La consulta fue bloqueada por la IA (${bloqueo})` };
        }
        const partes = data && data.candidates && data.candidates[0] &&
          data.candidates[0].content && data.candidates[0].content.parts;
        const texto = Array.isArray(partes)
          ? partes.filter(p => !p.thought).map(p => p.text || "").join("").trim()
          : "";
        if (texto) {
          return { ok: true, texto: texto, modelo: modelo };
        }
        ultimoError = "La IA devolvió una respuesta vacía";
        break; // prueba con el siguiente modelo
      }

      ultimoError = (data && data.error && data.error.message) ? data.error.message : `HTTP ${codigo}`;

      if (codigo === 404) {
        break; // el modelo no existe o ya no está disponible: siguiente modelo
      }
      if (GEMINI_CODIGOS_REINTENTABLES.indexOf(codigo) !== -1) {
        if (hayMasIntentos) {
          Utilities.sleep(GEMINI_ESPERAS_MS[intento]);
          continue;
        }
        break; // agotó reintentos con este modelo: siguiente modelo
      }

      // Clave inválida: ningún modelo va a funcionar, se detiene de una vez
      if (codigo === 401 || /api key/i.test(ultimoError)) {
        return { ok: false, error: ultimoError };
      }
      break; // otros errores (400, 403...): se prueba el siguiente modelo
    }
    errores.push(`${modelo}: ${ultimoError}`);
  }

  return { ok: false, error: errores.join(" | ") || ultimoError };
}

function consultarAsistenteEstudio(consulta) {
  const texto = (consulta || "").toString().trim();
  if (!texto) {
    return {
      respuesta: "Escribe cualquier pregunta: sobre tus entregas, prioridades, cómo organizarte o dudas de estudio.",
      tareasRelacionadas: []
    };
  }

  const entregas = obtenerCalendarioEntregas({ estado: "TODOS" });
  const abiertas = entregas.filter(t => t.estado !== ESTADOS_TAREA.COMPLETADA);
  const completadas = entregas.filter(t => t.estado === ESTADOS_TAREA.COMPLETADA);
  const relacionadasLocal = filtrarTareasPorConsulta(entregas, abiertas, texto); // solo respaldo

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    return {
      respuesta: construirRespuestaLocal(texto, abiertas, relacionadasLocal) +
        "\n\n(Nota: aún no has configurado GEMINI_API_KEY en Script Properties; se usó el modo local).",
      tareasRelacionadas: relacionadasLocal
    };
  }

  try {
    const zona = Session.getScriptTimeZone();
    const hoy = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");

    // Contexto: todas las abiertas (hasta 80) + las últimas 30 completadas.
    // Cada tarea lleva un id para que la IA pueda indicar cuáles mostrar en la tabla.
    const base = abiertas.slice(0, 80).concat(completadas.slice(-30));
    const contexto = base.map((t, i) => ({
      id: i,
      codigo: t.codigo,
      titulo: t.titulo,
      categoria: t.categoria,
      fechaEntrega: t.fechaEntrega,
      diasRestantes: t.diasRestantes,
      estado: t.estado,
      avance: t.avance,
      prioridad: t.prioridad,
      descripcion: t.descripcion
    }));

    const prompt = [
      "Eres el asistente académico de la app To-Task para estudiantes universitarios. Responde siempre en español, de forma clara, práctica y concisa.",
      `Fecha de hoy: ${hoy} (formato de fechaEntrega: dd/MM/yyyy).`,
      "Puedes responder CUALQUIER pregunta del estudiante: sobre sus tareas (filtros por prioridad, estado, fecha, materia, avance, comparaciones, planes de estudio, qué hacer primero) y también consejos generales de estudio o dudas académicas.",
      "Cuando hables de las tareas del estudiante, usa únicamente los datos del JSON; no inventes entregas, fechas ni valores.",
      "Prioridad va de 1 a 5 (5 = la más alta). Estados posibles: " + Object.values(ESTADOS_TAREA).join(", ") + ".",
      "Si la pregunta pide un filtro concreto (por ejemplo prioridad 5), incluye SOLO las tareas que cumplan exactamente ese filtro.",
      `Tareas del estudiante: ${JSON.stringify(contexto)}`,
      `Pregunta del estudiante: ${texto}`,
      "Devuelve un JSON con: 'respuesta' (tu respuesta en texto) e 'ids' (lista de los id de las tareas relevantes para tu respuesta, en orden de importancia; [] si la pregunta no trata de tareas concretas)."
    ].join("\n\n");

    const resultado = llamarGemini(prompt, apiKey, {
      responseMimeType: "application/json",
      responseSchema: {
        type: "OBJECT",
        properties: {
          respuesta: { type: "STRING" },
          ids: { type: "ARRAY", items: { type: "INTEGER" } }
        },
        required: ["respuesta", "ids"]
      }
    });

    if (!resultado.ok) {
      return {
        respuesta: `No se pudo consultar la API de IA (${resultado.error}). Resumen local:\n${construirRespuestaLocal(texto, abiertas, relacionadasLocal)}`,
        tareasRelacionadas: relacionadasLocal
      };
    }

    let respuesta = resultado.texto;
    let tareas = [];
    try {
      const json = JSON.parse(resultado.texto);
      respuesta = (json.respuesta || "").toString();
      const vistos = {};
      tareas = (json.ids || [])
        .filter(id => Number.isInteger(id) && !vistos[id] && (vistos[id] = true))
        .map(id => base[id])
        .filter(Boolean)
        .slice(0, 15);
    } catch (e) {
      // Si no llegó JSON válido, se muestra el texto tal cual y sin tabla
    }

    return {
      respuesta: respuesta || construirRespuestaLocal(texto, abiertas, relacionadasLocal),
      tareasRelacionadas: tareas
    };
  } catch (error) {
    console.error("Error en consultarAsistenteEstudio con IA:", error);
    return {
      respuesta: `Error al conectar con la IA (${error.message}). Resumen local:\n${construirRespuestaLocal(texto, abiertas, relacionadasLocal)}`,
      tareasRelacionadas: relacionadasLocal
    };
  }
}

function quitarTildes(s) {
  return (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function filtrarTareasPorConsulta(entregas, abiertas, texto) {
  const t = quitarTildes(texto.toLowerCase());
  const pidePrioridad = t.match(/prioridad\s*(?:de\s*|nivel\s*)?([1-5])/) || t.match(/\b([1-5])\s*\/\s*5\b/);
  const pideCompletadas = /completad|terminad|finalizad|entregad|hech[ao]s?\b/.test(t);

  let lista = pideCompletadas
    ? entregas.filter(x => x.estado === ESTADOS_TAREA.COMPLETADA)
    : abiertas;
  let hayFiltro = pideCompletadas;

  if (pidePrioridad) {
    const nivel = Number(pidePrioridad[1]);
    lista = lista.filter(x => Number(x.prioridad) === nivel);
    hayFiltro = true;
  }

  if (/atras|vencid/.test(t)) {
    lista = lista.filter(x => x.estado === ESTADOS_TAREA.ATRASADA);
    hayFiltro = true;
  } else if (/proxim|semana|hoy|urgent/.test(t)) {
    lista = lista.filter(x => x.diasRestantes !== null && x.diasRestantes <= 7);
    hayFiltro = true;
  } else if (/en proceso|progreso|avanzando/.test(t)) {
    lista = lista.filter(x => x.estado === ESTADOS_TAREA.EN_PROCESO);
    hayFiltro = true;
  } else if (/pendiente/.test(t)) {
    lista = lista.filter(x => x.estado === ESTADOS_TAREA.PENDIENTE);
    hayFiltro = true;
  }

  if (hayFiltro) {
    return lista.slice(0, 10);
  }

  // Sin filtros especiales: búsqueda por palabras clave (código, título, categoría, descripción)
  const ignoradas = ["cual", "cuales", "que", "son", "mis", "las", "los", "del", "con", "para", "una", "uno",
    "hay", "tengo", "dame", "muestrame", "mostrar", "todas", "todos", "entregas", "entrega", "tareas", "tarea", "como", "por"];
  const palabras = t.replace(/[^a-z0-9\s\-]/g, " ").split(/\s+/)
    .filter(p => p.length >= 3 && ignoradas.indexOf(p) === -1);

  if (palabras.length > 0) {
    const coincidencias = entregas.filter(x => {
      const campo = quitarTildes([x.codigo, x.titulo, x.categoria, x.descripcion].join(" ").toLowerCase());
      return palabras.some(p => campo.indexOf(p) !== -1);
    });
    if (coincidencias.length > 0) {
      return coincidencias.slice(0, 10);
    }
  }

  return abiertas.slice(0, 10);
}

function construirRespuestaLocal(texto, abiertas, filtradas) {
  if (!filtradas || filtradas.length === 0) {
    return `Consulta: "${texto}" (${abiertas.length} tarea(s) abierta(s) en total):\nNo encontré tareas que coincidan con tu consulta.`;
  }
  const lineas = filtradas.slice(0, 5).map(item =>
    `• [${item.codigo}] ${item.titulo} — Entrega: ${item.fechaEntrega} (${item.estado}, ${item.avance}%, prioridad ${item.prioridad}/5)`
  );
  return `Consulta: "${texto}" (${abiertas.length} tarea(s) abierta(s) en total):\n${lineas.join("\n")}`;
}

function obtenerListasAcademicas() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let catSheet = ss.getSheetByName(HOJA_CATEGORIAS);
    if (!catSheet) {
      catSheet = ss.insertSheet(HOJA_CATEGORIAS);
      const categoriasBase = [
        ["Categoría"],
        ["Parcial / Examen"],
        ["Tarea / Taller"],
        ["Proyecto"],
        ["Exposición"],
        ["Lectura / Quiz"],
        ["Laboratorio"]
      ];
      catSheet.getRange(1, 1, categoriasBase.length, 1).setValues(categoriasBase);
      catSheet.getRange(1, 1).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
    }

    let estSheet = ss.getSheetByName(HOJA_ESTADOS);
    if (!estSheet) {
      estSheet = ss.insertSheet(HOJA_ESTADOS);
      const estadosBase = [
        ["Estado"],
        [ESTADOS_TAREA.PENDIENTE],
        [ESTADOS_TAREA.EN_PROCESO],
        [ESTADOS_TAREA.COMPLETADA],
        [ESTADOS_TAREA.ATRASADA]
      ];
      estSheet.getRange(1, 1, estadosBase.length, 1).setValues(estadosBase);
      estSheet.getRange(1, 1).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
    }

    const categorias = catSheet.getDataRange().getValues()
      .slice(1).map(r => r[0]).filter(c => c && c.toString().trim());
    const estados = estSheet.getDataRange().getValues()
      .slice(1).map(r => r[0]).filter(e => e && e.toString().trim());

    return {
      categorias: categorias.length ? categorias : ["Tarea / Taller", "Parcial / Examen", "Proyecto"],
      estados: estados.length ? estados : Object.values(ESTADOS_TAREA)
    };
  } catch (error) {
    console.error("Error en obtenerListasAcademicas:", error);
    return {
      categorias: ["Tarea / Taller", "Parcial / Examen", "Proyecto", "Exposición"],
      estados: Object.values(ESTADOS_TAREA)
    };
  }
}

function inicializarSistemaUniversitario() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let matSheet = ss.getSheetByName(HOJA_MATERIAS);
    if (!matSheet) {
      matSheet = ss.insertSheet(HOJA_MATERIAS);
    }
    if (matSheet.getLastRow() === 0) {
      matSheet.getRange(1, 1, 1, 6).setValues([[
        "Código", "Título / Materia", "Modalidad", "Categoría", "Prioridad (1-5)", "Fecha Creación"
      ]]);
      matSheet.getRange(1, 1, 1, 6).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
      matSheet.getRange("A:A").setNumberFormat("@");
      matSheet.getRange("E:E").setNumberFormat("0");
      matSheet.getRange("F:F").setNumberFormat("dd/mm/yyyy hh:mm");
      matSheet.autoResizeColumns(1, 6);
    }

    let tarSheet = ss.getSheetByName(HOJA_TAREAS);
    if (!tarSheet) {
      tarSheet = ss.insertSheet(HOJA_TAREAS);
    }
    if (tarSheet.getLastRow() === 0) {
      tarSheet.getRange(1, 1, 1, 8).setValues([[
        "Código", "Fecha Entrega", "Estado", "Avance (%)", "Usuario", "Timestamp", "Descripción / Notas", "Prioridad"
      ]]);
      tarSheet.getRange(1, 1, 1, 8).setBackground("#2E86AB").setFontColor("white").setFontWeight("bold");
      tarSheet.getRange("A:A").setNumberFormat("@");
      tarSheet.getRange("B:B").setNumberFormat("dd/mm/yyyy");
      tarSheet.getRange("D:D").setNumberFormat("0.##");
      tarSheet.getRange("F:F").setNumberFormat("dd/mm/yyyy hh:mm:ss");
      tarSheet.getRange("H:H").setNumberFormat("0");
      tarSheet.autoResizeColumns(1, 8);
    }

    obtenerListasAcademicas();
    return "Sistema To-Task inicializado correctamente con las hojas Materias, Tareas, Categorias y Estados.";
  } catch (error) {
    console.error("Error en inicializarSistemaUniversitario:", error);
    return `Error al inicializar sistema: ${error.message}`;
  }
}

function validarIntegridadAcademica() {
  const errores = [];
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requeridas = [HOJA_MATERIAS, HOJA_TAREAS, HOJA_CATEGORIAS, HOJA_ESTADOS];
    requeridas.forEach(nombre => {
      if (!ss.getSheetByName(nombre)) {
        errores.push(`Falta la hoja requerida: ${nombre}`);
      }
    });

    if (errores.length > 0) {
      return { errores };
    }

    const matSheet = ss.getSheetByName(HOJA_MATERIAS);
    const tarSheet = ss.getSheetByName(HOJA_TAREAS);
    const materias = matSheet.getDataRange().getValues();
    const tareas = tarSheet.getDataRange().getValues();

    const codigosValidos = new Set();
    for (let i = 1; i < materias.length; i++) {
      if (!materias[i][0]) continue;
      const cod = materias[i][0].toString().trim().toUpperCase();
      if (codigosValidos.has(cod)) {
        errores.push(`Código duplicado en hoja Materias: ${cod} (fila ${i + 1})`);
      }
      codigosValidos.add(cod);

      if (!materias[i][1] || materias[i][1].toString().trim().length < 2) {
        errores.push(`Registro ${cod} sin título válido (fila ${i + 1})`);
      }
    }

    for (let i = 1; i < tareas.length; i++) {
      if (!tareas[i][0]) continue;
      const cod = tareas[i][0].toString().trim().toUpperCase();
      const fecha = parsearFechaLocal(tareas[i][1]);
      const estado = (tareas[i][2] || "").toString().trim().toUpperCase();
      const avance = parseFloat(tareas[i][3]);

      if (!codigosValidos.has(cod)) {
        errores.push(`Tarea con código no registrado en Materias: ${cod} (fila ${i + 1})`);
      }
      if (!fecha) {
        errores.push(`Fecha de entrega inválida en Tareas (fila ${i + 1})`);
      }
      if (!Object.values(ESTADOS_TAREA).includes(estado)) {
        errores.push(`Estado inválido '${estado}' en Tareas (fila ${i + 1})`);
      }
      if (isNaN(avance) || avance < 0 || avance > 100) {
        errores.push(`Porcentaje de avance fuera de rango [0-100] en Tareas (fila ${i + 1})`);
      }
    }

    return { errores };
  } catch (error) {
    errores.push(`Error al validar integridad: ${error.message}`);
    return { errores };
  }
}

function exportarCalendarioCSV() {
  try {
    const items = obtenerCalendarioEntregas({ estado: "TODOS" });
    if (items.length === 0) {
      return null;
    }

    let csv = "\uFEFF";
    csv += "Código,Asignatura / Tarea,Categoría,Fecha Entrega,Días Restantes,Estado,Avance (%),Prioridad,Descripción\n";

    items.forEach(item => {
      const descEscapada = (item.descripcion || "").replace(/"/g, '""');
      const tituloEscapado = (item.titulo || "").replace(/"/g, '""');
      csv += `"${item.codigo}","${tituloEscapado}","${item.categoria}","${item.fechaEntrega}",${item.diasRestantes !== null ? item.diasRestantes : ""},"${item.estado}",${item.avance},${item.prioridad},"${descEscapada}"\n`;
    });

    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
    const blob = Utilities.newBlob(csv, "text/csv; charset=utf-8", `ToTask_Entregas_${stamp}.csv`);

    let carpeta;
    try {
      carpeta = DriveApp.getFoldersByName("To-Task Reportes").next();
    } catch (e) {
      carpeta = DriveApp.getRootFolder();
    }

    const archivo = carpeta.createFile(blob);
    return archivo.getUrl();
  } catch (error) {
    console.error("Error en exportarCalendarioCSV:", error);
    return null;
  }
}
