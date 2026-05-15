// ══════════════════════════════════════════════════════════════════════════════
// GOOGLE APPS SCRIPT — SISTEMA INMU COMPLETO
// Incluye: Asistencia + Notas + MÓDULO DE INFORMES + Asistencia de Alumnos + Configuración
// ══════════════════════════════════════════════════════════════════════════════

// ══════════════════════════════════════════════════════════════════════════════
// HELPERS GENERALES
// ══════════════════════════════════════════════════════════════════════════════

function parseSafeJSON(value) {
  try {
    return (typeof value === 'string' && value) ? JSON.parse(value) : [];
  } catch (e) {
    return [];
  }
}

function toBooleanGas(value) {
  return value === true || value === 'true' || value === 1 || value === '1';
}

function normalizarEscalaMateriaGas(valor) {
  return (valor === '0-5' || valor === 5 || valor === '5') ? '0-5' : '0-10';
}

function safeJSONStringifyGas(value, fallback) {
  try {
    return JSON.stringify(value);
  } catch (e) {
    return JSON.stringify(fallback || []);
  }
}

function normalizarTexto(valor) {
  return (valor || "")
    .toString()
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}


// ══════════════════════════════════════════════════════════════════════════════
// CACHÉ DE DATOS — evita leer la hoja en cada petición (mejora velocidad 10x)
// ══════════════════════════════════════════════════════════════════════════════

var _cacheAlumnos = null; // caché en memoria para esta ejecución

function getAlumnosCache(ss) {
  if (_cacheAlumnos) return _cacheAlumnos; // hit en memoria (misma ejecución)

  var cache = CacheService.getScriptCache();
  var cached = cache.get('alumnos_data');
  if (cached) {
    try {
      _cacheAlumnos = JSON.parse(cached);
      return _cacheAlumnos;
    } catch(e) {}
  }

  // Miss: leer hojas y guardar en caché por 10 minutos
  var alumnos = [];
  var hojas = ['alumnos', 'di_refuerzo'];
  hojas.forEach(function(nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (!hoja) return;
    var rows = hoja.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][2]) { // tiene nombre
        alumnos.push({
          grado: rows[i][0], seccion: rows[i][1], nombre: rows[i][2],
          sexo: rows[i][3] || '', nie: (rows[i][4] || '').toString().trim(),
          telefono: rows[i][5] || '', fuente: nombre
        });
      }
    }
  });

  _cacheAlumnos = alumnos;
  try {
    var json = JSON.stringify(alumnos);
    if (json.length < 100000) { // CacheService tiene límite de 100KB por entrada
      cache.put('alumnos_data', json, 600); // 10 minutos
    }
  } catch(e) {}
  return alumnos;
}

function invalidarCacheAlumnos() {
  _cacheAlumnos = null;
  try { CacheService.getScriptCache().remove('alumnos_data'); } catch(e) {}
}

function getEspecialidadesPredeterminadasGas() {
  return [
    "1° Adm. Contable",
    "1° S. Eléctricos",
    "2° Adm. Contable",
    "2° Sist. Eléctricos",
    "3° Adm. Contable",
    "3° Sist. Eléctricos"
  ];
}

function getCatalogoPredeterminadoGas() {
  var especialidades = getEspecialidadesPredeterminadasGas();
  return [
    { nombre: "Lenguaje y Literatura", categoria: "basica", escala: "0-10", requiere_especialidad: false, especialidades: [] },
    { nombre: "Matemática", categoria: "basica", escala: "0-10", requiere_especialidad: false, especialidades: [] },
    { nombre: "Estudios Sociales y Cívica", categoria: "basica", escala: "0-10", requiere_especialidad: false, especialidades: [] },
    { nombre: "Ciencia, Salud y Medio Ambiente", categoria: "basica", escala: "0-10", requiere_especialidad: false, especialidades: [] },
    { nombre: "Inglés", categoria: "basica", escala: "0-10", requiere_especialidad: false, especialidades: [] },
    { nombre: "Informática", categoria: "basica", escala: "0-10", requiere_especialidad: false, especialidades: [] },
    { nombre: "Educación Física", categoria: "basica", escala: "0-10", requiere_especialidad: false, especialidades: [] },
    { nombre: "Moral, Urbanidad y Cívica", categoria: "basica", escala: "0-10", requiere_especialidad: false, especialidades: [] },
    { nombre: "Seminario", categoria: "basica", escala: "0-10", requiere_especialidad: false, especialidades: [] },
    { nombre: "Módulo 1", categoria: "modulo", escala: "0-5", requiere_especialidad: true, especialidades: especialidades },
    { nombre: "Módulo 2", categoria: "modulo", escala: "0-5", requiere_especialidad: true, especialidades: especialidades },
    { nombre: "Módulo 3", categoria: "modulo", escala: "0-5", requiere_especialidad: true, especialidades: especialidades }
  ];
}

function asegurarHojaDocentes(ss) {
  var sheet = ss.getSheetByName("docentes");
  if (!sheet) {
    sheet = ss.insertSheet("docentes");
    sheet.appendRow(["Nombre", "Grado", "Seccion", "Materia", "Escala", "Admin"]);
    return sheet;
  }

  var lastCol = Math.max(sheet.getLastColumn(), 6);
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var yaMigrada = normalizarTexto(header[3]) === 'materia' &&
                  normalizarTexto(header[4]) === 'escala' &&
                  normalizarTexto(header[5]) === 'admin';
  if (!yaMigrada) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var oldAdmin = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
      sheet.getRange(2, 6, lastRow - 1, 1).setValues(oldAdmin);
      sheet.getRange(2, 4, lastRow - 1, 2).clearContent();
    }
    sheet.getRange(1, 1, 1, 6).setValues([["Nombre", "Grado", "Seccion", "Materia", "Escala", "Admin"]]);
  }
  return sheet;
}

function asegurarHojaAsignacionesDocente(ss) {
  var sheet = ss.getSheetByName("docente_materias");
  var headers = ["Docente", "Grado", "Seccion", "TipoMateria", "Materia", "Escala", "Especialidad", "EsOrientado", "Activo"];
  if (!sheet) {
    sheet = ss.insertSheet("docente_materias");
    sheet.appendRow(headers);
    return sheet;
  }
  if (sheet.getLastColumn() < headers.length) {
    sheet.insertColumnsAfter(sheet.getLastColumn(), headers.length - sheet.getLastColumn());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function asegurarHojaCatalogoMaterias(ss) {
  var sheet = ss.getSheetByName("catalogo_materias");
  var headers = ["Nombre", "Categoria", "Escala", "RequiereEspecialidad", "EspecialidadesJSON", "Activo"];
  if (!sheet) {
    sheet = ss.insertSheet("catalogo_materias");
    sheet.appendRow(headers);
  } else {
    if (sheet.getLastColumn() < headers.length) {
      sheet.insertColumnsAfter(sheet.getLastColumn(), headers.length - sheet.getLastColumn());
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  if (sheet.getLastRow() < 2) {
    var filasBase = getCatalogoPredeterminadoGas().map(function(item) {
      return [
        item.nombre,
        item.categoria,
        normalizarEscalaMateriaGas(item.escala),
        item.requiere_especialidad ? "true" : "false",
        safeJSONStringifyGas(item.especialidades || [], []),
        "true"
      ];
    });
    if (filasBase.length) {
      sheet.getRange(2, 1, filasBase.length, headers.length).setValues(filasBase);
    }
  }
  return sheet;
}

function getCatalogoMateriasGas(ss) {
  var sheet = asegurarHojaCatalogoMaterias(ss);
  var rows = sheet.getDataRange().getValues();
  var data = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0] || (rows[i][5] !== '' && !toBooleanGas(rows[i][5]))) continue;
    data.push({
      nombre: rows[i][0] || '',
      categoria: rows[i][1] || 'basica',
      escala: normalizarEscalaMateriaGas(rows[i][2]),
      requiere_especialidad: toBooleanGas(rows[i][3]),
      especialidades: parseSafeJSON(rows[i][4]).filter(function(v) { return v !== null && v !== ''; })
    });
  }
  return data;
}

function normalizarAsignacionDocenteGas(item, docente, fallbackGrado, fallbackSeccion) {
  if (!item) item = {};
  var escala = normalizarEscalaMateriaGas(item.escala || item.tipo_materia || item.escala_notas);
  var tipo = item.tipo_materia || item.categoria || (escala === '0-5' ? 'modulo' : 'basica');
  return {
    docente: docente || item.docente || '',
    grado: item.grado || fallbackGrado || '',
    seccion: item.seccion || fallbackSeccion || '',
    tipo_materia: tipo === 'modulo' ? 'modulo' : 'basica',
    materia: item.materia || item.nombre || '',
    escala: escala,
    especialidad: item.especialidad || '',
    es_orientado: toBooleanGas(item.es_orientado || item.orientado || item.esOrientado)
  };
}

function getMapaAsignacionesDocenteGas(ss) {
  var map = {};
  var asignSheet = asegurarHojaAsignacionesDocente(ss);
  var asignRows = asignSheet.getDataRange().getValues();

  for (var i = 1; i < asignRows.length; i++) {
    if (!asignRows[i][0]) continue;
    if (asignRows[i][8] !== '' && !toBooleanGas(asignRows[i][8])) continue;
    var docente = asignRows[i][0];
    if (!map[docente]) map[docente] = [];
    map[docente].push({
      docente: docente,
      grado: asignRows[i][1] || '',
      seccion: asignRows[i][2] || '',
      tipo_materia: asignRows[i][3] || (normalizarEscalaMateriaGas(asignRows[i][5]) === '0-5' ? 'modulo' : 'basica'),
      materia: asignRows[i][4] || '',
      escala: normalizarEscalaMateriaGas(asignRows[i][5]),
      especialidad: asignRows[i][6] || '',
      es_orientado: toBooleanGas(asignRows[i][7])
    });
  }

  var docentesSheet = asegurarHojaDocentes(ss);
  var docentesRows = docentesSheet.getDataRange().getValues();
  for (var j = 1; j < docentesRows.length; j++) {
    var nombre = docentesRows[j][0];
    if (!nombre || (map[nombre] && map[nombre].length > 0)) continue;
    if (!docentesRows[j][3]) continue;
    map[nombre] = [normalizarAsignacionDocenteGas({
      materia: docentesRows[j][3] || '',
      escala: docentesRows[j][4] || '0-10',
      es_orientado: true
    }, nombre, docentesRows[j][1] || '', docentesRows[j][2] || '')];
  }

  Object.keys(map).forEach(function(nombreDocente) {
    map[nombreDocente].sort(function(a, b) {
      return (b.es_orientado ? 1 : 0) - (a.es_orientado ? 1 : 0);
    });
  });

  return map;
}

function reemplazarAsignacionesDocenteGas(ss, docente, asignaciones) {
  var sheet = asegurarHojaAsignacionesDocente(ss);
  var rows = sheet.getDataRange().getValues();
  var nuevasFilas = [rows[0]];
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '') !== docente) nuevasFilas.push(rows[i]);
  }

  (asignaciones || []).forEach(function(item, idx) {
    var normalizado = normalizarAsignacionDocenteGas(item, docente);
    if (!normalizado.materia) return;
    nuevasFilas.push([
      docente,
      normalizado.grado,
      normalizado.seccion,
      normalizado.tipo_materia,
      normalizado.materia,
      normalizado.escala,
      normalizado.especialidad || '',
      (idx === 0 ? true : normalizado.es_orientado) ? "true" : "false",
      "true"
    ]);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
}

function upsertDocenteBaseGas(ss, docenteData) {
  var sheet = asegurarHojaDocentes(ss);
  var rows = sheet.getDataRange().getValues();
  var nombre = docenteData.nombre || '';
  var admin = toBooleanGas(docenteData.admin);
  var asignaciones = docenteData.materias_asignadas || [];
  var principal = null;
  for (var i = 0; i < asignaciones.length; i++) {
    if (toBooleanGas(asignaciones[i].es_orientado)) {
      principal = asignaciones[i];
      break;
    }
  }
  if (!principal && asignaciones.length) principal = asignaciones[0];

  var fila = [
    nombre,
    docenteData.grado_orientado || docenteData.grado || (principal ? principal.grado : '') || '',
    docenteData.seccion_orientada || docenteData.seccion || (principal ? principal.seccion : '') || '',
    principal ? (principal.materia || '') : (docenteData.materia || ''),
    principal ? normalizarEscalaMateriaGas(principal.escala) : normalizarEscalaMateriaGas(docenteData.tipo_materia || docenteData.escala || '0-10'),
    admin ? "true" : "false"
  ];

  var filaEncontrada = -1;
  for (var r = 1; r < rows.length; r++) {
    if ((rows[r][0] || '') === nombre) {
      filaEncontrada = r + 1;
      break;
    }
  }

  if (filaEncontrada > 0) {
    sheet.getRange(filaEncontrada, 1, 1, fila.length).setValues([fila]);
  } else {
    sheet.appendRow(fila);
  }
}

function getDocentesConAsignacionesGas(ss) {
  var sheet = asegurarHojaDocentes(ss);
  var rows = sheet.getDataRange().getValues();
  var mapaAsignaciones = getMapaAsignacionesDocenteGas(ss);
  var data = [];

  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var nombre = rows[i][0];
    var asignaciones = (mapaAsignaciones[nombre] || []).slice();
    var orientada = null;
    for (var j = 0; j < asignaciones.length; j++) {
      if (toBooleanGas(asignaciones[j].es_orientado)) {
        orientada = asignaciones[j];
        break;
      }
    }
    if (!orientada && asignaciones.length) orientada = asignaciones[0];
    var principal = orientada || {
      materia: rows[i][3] || '',
      escala: rows[i][4] || '0-10',
      tipo_materia: normalizarEscalaMateriaGas(rows[i][4]) === '0-5' ? 'modulo' : 'basica'
    };

    data.push({
      nombre: nombre,
      grado: rows[i][1] || '',
      seccion: rows[i][2] || '',
      grado_orientado: rows[i][1] || '',
      seccion_orientada: rows[i][2] || '',
      materia: principal.materia || '',
      tipo_materia: principal.escala || "0-10",
      escala: principal.escala || "0-10",
      admin: toBooleanGas(rows[i][5]),
      materias_asignadas: asignaciones
    });
  }

  return data;
}

function obtenerHojaDiRefuerzo(ss) {
  var sheet = ss.getSheetByName("di_refuerzo");
  if (!sheet) {
    sheet = ss.insertSheet("di_refuerzo");
    sheet.appendRow(["Grado", "Seccion", "Nombre", "Sexo", "NIE", "Telefono"]);
  }
  return sheet;
}

function leerAlumnosDesdeHoja(sheet) {
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var data = [];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][2]) {
      data.push({
        grado: rows[i][0], seccion: rows[i][1], nombre: rows[i][2],
        sexo: rows[i][3], nie: rows[i][4], telefono: rows[i][5]
      });
    }
  }
  return data;
}

// ══════════════════════════════════════════════════════════════════════════════
// doGet — ENDPOINTS GET
// ══════════════════════════════════════════════════════════════════════════════

function doGet(e) {
  var props = PropertiesService.getScriptProperties();
  var enMantenimiento = props.getProperty('MANTENIMIENTO') === 'true';
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : null;

  // Helper JSONP: responde con callback(json) si hay ?callback=, si no responde JSON normal
  // Esto resuelve el bloqueo CORS desde GitHub Pages y otras páginas externas
  function jsonResp(obj) {
    var json = JSON.stringify(obj);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  if (e && e.parameter && e.parameter.tipo === "check_mantenimiento") {
    var loginHabilitado = props.getProperty('LOGIN_HABILITADO');
    if (loginHabilitado === null) loginHabilitado = 'true';
    var modoAlumnoActivo = props.getProperty('MODO_ALUMNO_ACTIVO');
    if (modoAlumnoActivo === null) modoAlumnoActivo = 'true';
    return jsonResp({
      mantenimiento: enMantenimiento,
      login_habilitado: (loginHabilitado === 'true'),
      modo_alumno: (modoAlumnoActivo === 'true')
    });
  }

  if (enMantenimiento) {
    return jsonResp({ error: "mantenimiento", mantenimiento: true });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tipo = (e && e.parameter) ? e.parameter.tipo : null;

  // ── reportes ──────────────────────────────────────────────────────────────
  if (tipo === "reportes") {
    var sheet = ss.getSheetByName("reportes");
    if (!sheet) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var rows = sheet.getDataRange().getValues();
    var data = [];
    for (var i = 1; i < rows.length; i++) {
      data.push({
        fecha: rows[i][0], grado: rows[i][1], seccion: rows[i][2], docente: rows[i][3],
        presentes: rows[i][4], ausentes: rows[i][5], permisos: rows[i][6] || 0,
        m: rows[i][7], f: rows[i][8],
        asistentes: parseSafeJSON(rows[i][9]),
        ausentes_lista: parseSafeJSON(rows[i][10]),
        permisos_lista: parseSafeJSON(rows[i][11])
      });
    }
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "lista_docentes" || tipo === "docentes") {
    return ContentService.createTextOutput(JSON.stringify(getDocentesConAsignacionesGas(ss))).setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "catalogo_materias") {
    return ContentService.createTextOutput(JSON.stringify(getCatalogoMateriasGas(ss))).setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "alumnos") {
    var gradoParam = normalizarTexto(e.parameter && e.parameter.grado ? e.parameter.grado : '');
    var seccionParam = normalizarTexto(e.parameter && e.parameter.seccion ? e.parameter.seccion : '');

    var esDiRefuerzo = (gradoParam === "di refuerzo") && (seccionParam === "unica" || seccionParam === "única");
    if (esDiRefuerzo) {
      return ContentService.createTextOutput(JSON.stringify(leerAlumnosDesdeHoja(obtenerHojaDiRefuerzo(ss))))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var data = [];
    function agregarDesdeSheet(sheet) {
      if (!sheet) return;
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var nombre = (rows[i][2] || "").toString().trim();
        if (nombre !== "") {
          var gradoRow = normalizarTexto(rows[i][0]);
          var seccionRow = normalizarTexto(rows[i][1]);
          if ((gradoParam === "" || gradoRow === gradoParam) && (seccionParam === "" || seccionRow === seccionParam)) {
            data.push({ grado: rows[i][0], seccion: rows[i][1], nombre: nombre, sexo: rows[i][3], nie: rows[i][4], telefono: rows[i][5] });
          }
        }
      }
    }
    agregarDesdeSheet(ss.getSheetByName("alumnos"));
    agregarDesdeSheet(ss.getSheetByName("di_refuerzo"));
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "detalles_asistencia") {
    var sheet = ss.getSheetByName("reportes");
    if (!sheet) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var grado = e.parameter.grado || '';
    var seccion = e.parameter.seccion || '';
    var docente = e.parameter.docente || '';
    var fecha = e.parameter.fecha || '';
    var filtroFecha = fecha ? new Date(fecha) : null;
    var rows = sheet.getDataRange().getValues();
    var data = [];
    for (var i = rows.length - 1; i >= 1; i--) {
      var filaFecha = rows[i][0] ? new Date(rows[i][0]) : null;
      var coincidente = rows[i][1] === grado && rows[i][2] === seccion && rows[i][3] === docente;
      if (coincidente && filtroFecha) coincidente = filaFecha && filtroFecha.getTime() === filaFecha.getTime();
      if (coincidente) {
        try {
          data.push({
            fecha: rows[i][0], grado: rows[i][1], seccion: rows[i][2], docente: rows[i][3],
            presentes: rows[i][4], ausentes: rows[i][5], permisos: rows[i][6],
            asistentes: parseSafeJSON(rows[i][9]) || [],
            ausentes_lista: parseSafeJSON(rows[i][10]) || [],
            permisos_lista: parseSafeJSON(rows[i][11]) || []
          });
        } catch (err) {}
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "permisos") {
    var sheet = ss.getSheetByName("permisos");
    if (!sheet) {
      sheet = ss.insertSheet("permisos");
      sheet.appendRow(["Fecha", "Grado", "Seccion", "Docente", "Estudiante", "Sexo", "NIE", "Justificante"]);
    }
    var permisoRows = sheet.getDataRange().getValues();
    var data = [];
    for (var i = 1; i < permisoRows.length; i++) {
      if (permisoRows[i][0]) {
        data.push({
          fecha: permisoRows[i][0], grado: permisoRows[i][1], seccion: permisoRows[i][2],
          docente: permisoRows[i][3], estudiante: permisoRows[i][4],
          sexo: permisoRows[i][5] || '', nie: permisoRows[i][6] || '',
          justificante: permisoRows[i][7] !== undefined ? permisoRows[i][7] : permisoRows[i][5] || ''
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "detalle_alumno") {
    var nombre = e.parameter.nombre || '';
    var nie = e.parameter.nie || '';
    var gradoAlumno = e.parameter.grado || '';
    var seccionAlumno = e.parameter.seccion || '';
    var fecha = e.parameter.fecha || '';
    var permisosSheet = ss.getSheetByName("permisos");
    var justificantes = [];
    var sexo = '';
    if (permisosSheet) {
      var permisoRows = permisosSheet.getDataRange().getValues();
      var usesExtendedColumns = permisoRows[0] && permisoRows[0].length >= 8 && permisoRows[0][5] === 'Sexo' && permisoRows[0][6] === 'NIE';
      for (var i = 1; i < permisoRows.length; i++) {
        if (!permisoRows[i][0]) continue;
        var rowNombre = permisoRows[i][4] || '';
        var rowSexo = usesExtendedColumns ? (permisoRows[i][5] || '') : '';
        var rowNie = usesExtendedColumns ? (permisoRows[i][6] || '') : '';
        var rowJustificante = usesExtendedColumns ? (permisoRows[i][7] || '') : (permisoRows[i][5] || '');
        if (permisoRows[i][1] === gradoAlumno && permisoRows[i][2] === seccionAlumno && (rowNombre === nombre || (rowNie && rowNie === nie))) {
          sexo = sexo || rowSexo;
          justificantes.push({ fecha: permisoRows[i][0], docente: permisoRows[i][3] || '', justificante: rowJustificante, sexo: rowSexo, nie: rowNie });
        }
      }
    }
    var faltasTotales = 0;
    var conteoSheet = ss.getSheetByName("conteo_ausencias");
    if (conteoSheet) {
      var conteoRows = conteoSheet.getDataRange().getValues();
      for (var j = 1; j < conteoRows.length; j++) {
        if (conteoRows[j][0] === nombre && conteoRows[j][2] === gradoAlumno && conteoRows[j][3] === seccionAlumno) {
          faltasTotales = parseInt(conteoRows[j][4], 10) || 0;
          break;
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({
      nombre: nombre, nie: nie, sexo: sexo, grado: gradoAlumno, seccion: seccionAlumno,
      fecha: fecha, faltas_totales: faltasTotales, justificantes: justificantes
    })).setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "estudiantes_peligro") {
    return ContentService.createTextOutput(JSON.stringify(
      getEstudiantesEnPeligro(e.parameter.grado || '', e.parameter.seccion || '')
    )).setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "docentes_status") {
    // FIX: usar jsonResp para soporte JSONP (evita bloqueo CORS)
    return jsonResp(getDocentesStatus());

  } else if (tipo === "obtener_observaciones") {
    var nie = e.parameter.nie;
    if (!nie) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var sheet = ss.getSheetByName("observaciones");
    if (!sheet) return ContentService.createTextOutput("[]").setMimeType(ContentService.MimeType.JSON);
    var rows = sheet.getDataRange().getValues();
    var data = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] == nie) {
        data.push({ nie: rows[i][0], fecha: rows[i][1], docente: rows[i][2], observacion: rows[i][3] });
      }
    }
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "notas") {
    var gradoParam = (e && e.parameter && e.parameter.grado) ? e.parameter.grado : '';
    var seccionParam = (e && e.parameter && e.parameter.seccion) ? e.parameter.seccion : '';
    var escalaParam = (e && e.parameter && e.parameter.escala) ? e.parameter.escala : '0-10';
    var materiaClaveParam = (e && e.parameter && e.parameter.materia_clave) ? e.parameter.materia_clave : '';
    // Usar jsonResp para soportar JSONP cuando se llame con ?callback= (evita bloqueo CORS)
    return jsonResp(obtenerNotasPorGrado(gradoParam, escalaParam, materiaClaveParam, seccionParam));

  } else if (tipo === "historial_informes") {
    return ContentService.createTextOutput(JSON.stringify(obtenerHistorialInformes()))
      .setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "buscar_alumno") {
    var query = (e.parameter.query || '').toString().trim();
    var gradoFiltro = (e.parameter.grado || '').toString().trim();
    var seccionFiltro = (e.parameter.seccion || '').toString().trim();
    return ContentService.createTextOutput(JSON.stringify(buscarAlumno(query, gradoFiltro, seccionFiltro)))
      .setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "obtener_alumno_nie") {
    var nie = (e.parameter.nie || '').toString().trim();
    return ContentService.createTextOutput(JSON.stringify(obtenerAlumnoPorNIE(nie)))
      .setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "expediente_alumno") {
    var nie = (e.parameter.nie || '').toString().trim();
    return ContentService.createTextOutput(JSON.stringify(obtenerExpedienteAlumno(nie)))
      .setMimeType(ContentService.MimeType.JSON);

  } else if (tipo === "horario_asistencia") {
    var props2 = PropertiesService.getScriptProperties();
    var modoActivo = props2.getProperty('MODO_ALUMNO_ACTIVO');
    // acceso_alumnos: separado de "activo" (horario), controla si el portal de alumnos está habilitado
    var accesoAlumnos = (modoActivo !== 'false'); // default true
    return jsonResp({
      inicio:         props2.getProperty('HORARIO_INICIO') || '07:00',
      fin:            props2.getProperty('HORARIO_FIN')    || '08:00',
      activo:         accesoAlumnos,
      acceso_alumnos: accesoAlumnos,
      mantenimiento:  enMantenimiento
    });


  } else if (tipo === "marcar_alumno") {
    // FIX: endpoint GET/JSONP para marcar asistencia desde portal de alumnos
    // El POST original se perdía en el redirect 302 de Google (body vaciado)
    var nieM = (e.parameter.nie          || '').toString().trim();
    var estM = (e.parameter.estado       || 'presente').toString().trim();
    var jusM = (e.parameter.justificante || '').toString().trim();
    if (!nieM) return jsonResp({ exito: false, error: 'NIE requerido' });
    var resM = marcarAsistenciaAlumno({ nie: nieM, estado: estM, justificante: jusM });
    return jsonResp(resM);

  } else if (tipo === "asistencia_diaria_grado") {
    // FIX: sin filtro grado/seccion en servidor — evita fallo por diferencia de
    // caracteres (grado sign \u00b0 vs ordinal \u00ba). El HTML cruza por NIE.
    var asistSheet2 = ss.getSheetByName("asistencia_alumnos");
    if (!asistSheet2) {
      asistSheet2 = ss.insertSheet("asistencia_alumnos");
      asistSheet2.appendRow(["Fecha", "NIE", "Nombre", "Grado", "Seccion", "Estado", "Hora", "Justificante"]);
      var hdrRange = asistSheet2.getRange(1, 1, 1, 8);
      hdrRange.setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
    }
    var mapa = {};
    if (asistSheet2) {
      var hoy2  = new Date().toLocaleDateString('es-SV');
      var rows2 = asistSheet2.getDataRange().getValues();
      for (var r2 = 1; r2 < rows2.length; r2++) {
        var rFecha2 = rows2[r2][0] ? new Date(rows2[r2][0]).toLocaleDateString('es-SV') : '';
        if (rFecha2 !== hoy2) continue;
        var rNie2 = (rows2[r2][1] || '').toString().trim();
        if (!rNie2) continue;
        mapa[rNie2] = { estado: rows2[r2][5] || 'presente', hora: rows2[r2][6] || '', grado: rows2[r2][3] || '', seccion: rows2[r2][4] || '' };
      }
    }
    return jsonResp(mapa);
  } else if (tipo === "validar_alumno_nie") {
    var nie = (e.parameter.nie || '').toString().trim();
    var alumno = obtenerAlumnoPorNIE(nie);
    if (!alumno) {
      return jsonResp({ valido: false });
    }
    var hoy = new Date().toLocaleDateString('es-SV');
    var marcadoHoy = false;
    var asistSheet = ss.getSheetByName("asistencia_alumnos");
    if (asistSheet) {
      var asistRows = asistSheet.getDataRange().getValues();
      for (var q = 1; q < asistRows.length; q++) {
        var rowFecha = asistRows[q][0] ? new Date(asistRows[q][0]).toLocaleDateString('es-SV') : '';
        if (asistRows[q][1] == nie && rowFecha === hoy) { marcadoHoy = true; break; }
      }
    }
    alumno.valido = true;
    alumno.marcado_hoy = marcadoHoy;
    return jsonResp(alumno);

  } else {
    // Devolver todos los alumnos usando caché
    var data = getAlumnosCache(ss);
    return jsonResp(data);
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// doPost — ENDPOINTS POST
// ══════════════════════════════════════════════════════════════════════════════

function doPost(e) {
  var props = PropertiesService.getScriptProperties();
  var payload = (e && e.postData && e.postData.contents) ? e.postData.contents : '';
  var data = {};
  try {
    data = JSON.parse(payload);
  } catch (err) {
    data = {};
  }
  if ((!data || !data.tipo_post) && e && e.parameter) {
    data = Object.assign({}, data, e.parameter);
  }

  if (data.tipo_post === "toggle_mantenimiento") {
    if (data.password === "747-8") {
      var estadoActual = props.getProperty('MANTENIMIENTO') === 'true';
      props.setProperty('MANTENIMIENTO', (!estadoActual).toString());
      return ContentService.createTextOutput(JSON.stringify({ exito: true })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ exito: false })).setMimeType(ContentService.MimeType.JSON);
  }

  if (data.tipo_post === "toggle_login") {
    if (data.password === "747-8") {
      var loginActual = props.getProperty('LOGIN_HABILITADO');
      if (loginActual === null) loginActual = 'true';
      props.setProperty('LOGIN_HABILITADO', loginActual === 'true' ? 'false' : 'true');
      return ContentService.createTextOutput(JSON.stringify({ exito: true })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ exito: false })).setMimeType(ContentService.MimeType.JSON);
  }

  // ── NUEVO: toggle_modo_alumno — controla si alumnos pueden usar index.html ──
  if (data.tipo_post === "toggle_modo_alumno") {
    var nuevoActivo = (data.activo === true || data.activo === 'true');
    props.setProperty('MODO_ALUMNO_ACTIVO', nuevoActivo ? 'true' : 'false');
    return ContentService.createTextOutput(JSON.stringify({ exito: true, activo: nuevoActivo })).setMimeType(ContentService.MimeType.JSON);
  }

  if (data.tipo_post === "update_docente_status") {
    updateDocenteStatus(data.docente, data.status, data.timestamp || null);
    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  }

  if (props.getProperty('MANTENIMIENTO') === 'true') {
    return ContentService.createTextOutput(JSON.stringify({ error: "mantenimiento" })).setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (data.tipo_post === "asistencia") {
    var sheet = ss.getSheetByName("reportes");
    if (!sheet) {
      sheet = ss.insertSheet("reportes");
      sheet.appendRow(["Fecha", "Grado", "Seccion", "Docente", "Presentes", "Ausentes", "Permisos", "M", "F", "Asistentes", "Ausentes_Lista", "Permisos_Lista"]);
    }
    var permisosSheet = ss.getSheetByName("permisos");
    if (!permisosSheet) {
      permisosSheet = ss.insertSheet("permisos");
      permisosSheet.appendRow(["Fecha", "Grado", "Seccion", "Docente", "Estudiante", "Sexo", "NIE", "Justificante"]);
    }
    var conteoSheet = ss.getSheetByName("conteo_ausencias");
    if (!conteoSheet) {
      conteoSheet = ss.insertSheet("conteo_ausencias");
      conteoSheet.appendRow(["Estudiante", "NIE", "Grado", "Seccion", "Conteo"]);
    }
    sheet.appendRow([
      new Date(), data.grado, data.seccion, data.docente,
      data.presentes, data.ausentes, data.permisos || 0, data.m, data.f,
      JSON.stringify(data.asistentes || []),
      JSON.stringify(data.ausentes_lista || []),
      JSON.stringify(data.permisos_lista || [])
    ]);
    if (data.permisos_lista && Array.isArray(data.permisos_lista)) {
      data.permisos_lista.forEach(function(permiso) {
        permisosSheet.appendRow([new Date(), data.grado, data.seccion, data.docente, permiso.nombre || '', permiso.sexo || '', permiso.nie || '', permiso.justificante || '']);
      });
    }
    if (data.ausentes_lista && Array.isArray(data.ausentes_lista)) {
      data.ausentes_lista.forEach(function(ausente) { incrementarConteoAusencia(ausente, data.grado, data.seccion); });
    }
    return ContentService.createTextOutput("Exito").setMimeType(ContentService.MimeType.TEXT);

  } else if (data.tipo_post === "actualizar_asistencia") {
    var sheet = ss.getSheetByName("asistencia_actualizaciones");
    if (!sheet) {
      sheet = ss.insertSheet("asistencia_actualizaciones");
      sheet.appendRow(["Fecha", "Grado", "Seccion", "Docente", "NIE", "Nombre", "Estado", "Asistio", "Justificante", "Observacion"]);
    }
    if (data.actualizaciones && Array.isArray(data.actualizaciones)) {
      data.actualizaciones.forEach(function(act) {
        sheet.appendRow([new Date(), data.grado, data.seccion, data.docente || '', act.nie || '', act.nombre || '', act.estado || '', act.asistio === false ? 'false' : 'true', act.justificante || '', act.observacion || '']);
        if (act.estado === 'ausente') incrementarConteoAusencia(act.nombre, data.grado, data.seccion);
        if (act.estado === 'permiso') {
          var permisosSheet = ss.getSheetByName("permisos");
          if (!permisosSheet) { permisosSheet = ss.insertSheet("permisos"); permisosSheet.appendRow(["Fecha", "Grado", "Seccion", "Docente", "Estudiante", "Sexo", "NIE", "Justificante"]); }
          permisosSheet.appendRow([new Date(), data.grado, data.seccion, data.docente || '', act.nombre || '', act.sexo || '', act.nie || '', act.justificante || '']);
        }
      });
    }
    return ContentService.createTextOutput("Exito").setMimeType(ContentService.MimeType.TEXT);

  } else if (data.tipo_post === "delete_reportes") {
    var sheet = ss.getSheetByName("reportes");
    if (sheet) {
      if (data.rango === "all") {
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      } else if (data.rango === "hoy") {
        var rows = sheet.getDataRange().getValues();
        for (var i = rows.length - 1; i >= 1; i--) {
          if (new Date(rows[i][0]).toLocaleDateString('es-ES') === data.fecha) sheet.deleteRow(i + 1);
        }
      }
    }
    return ContentService.createTextOutput("Exito").setMimeType(ContentService.MimeType.TEXT);

  } else if (data.tipo_post === "alumno") {
    var gradoAlumno = normalizarTexto(data.grado);
    var seccionAlumno = normalizarTexto(data.seccion);
    var esDiRefuerzo = (gradoAlumno === "di refuerzo") && (seccionAlumno === "unica" || seccionAlumno === "única");
    var sheet = esDiRefuerzo ? obtenerHojaDiRefuerzo(ss) : ss.getSheets()[0];
    if (esDiRefuerzo) {
      var dataDi = sheet.getDataRange().getValues();
      var existeDi = false;
      var dataNie = (data.nie || '').toString().trim();
      for (var i = 1; i < dataDi.length; i++) {
        if ((dataNie && (dataDi[i][4] || '').toString().trim() === dataNie) || (data.nombre && normalizarTexto(dataDi[i][2]) === normalizarTexto(data.nombre))) { existeDi = true; break; }
      }
      if (!existeDi) sheet.appendRow([data.grado, data.seccion, data.nombre, data.sexo, data.nie, data.telefono]);
    } else {
      sheet.appendRow([data.grado, data.seccion, data.nombre, data.sexo, data.nie, data.telefono]);
    }
    return ContentService.createTextOutput("Exito").setMimeType(ContentService.MimeType.TEXT);

  } else if (data.tipo_post === "di_agregar") {
    var diSheetAdd = obtenerHojaDiRefuerzo(ss);
    var dataDiAdd = diSheetAdd.getDataRange().getValues();
    var agregar = data.alumno || {};
    var nieAlumnoAgregar = (agregar.nie || '').toString().trim();
    var existeAgregar = false;
    for (var a = 1; a < dataDiAdd.length; a++) {
      if ((nieAlumnoAgregar && (dataDiAdd[a][4] || '').toString().trim() === nieAlumnoAgregar) || (agregar.nombre && normalizarTexto(dataDiAdd[a][2]) === normalizarTexto(agregar.nombre))) { existeAgregar = true; break; }
    }
    if (!existeAgregar && agregar.nombre) diSheetAdd.appendRow(["DI REFUERZO", "Única", agregar.nombre || '', agregar.sexo || '', agregar.nie || '', agregar.telefono || '']);
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "di_eliminar") {
    var diSheetDel = obtenerHojaDiRefuerzo(ss);
    var dataDiDel = diSheetDel.getDataRange().getValues();
    var eliminado = false;
    var buscar = data.alumno || {};
    var nieAlumnoEliminar = (buscar.nie || '').toString().trim();
    for (var d = dataDiDel.length - 1; d >= 1; d--) {
      if ((nieAlumnoEliminar && (dataDiDel[d][4] || '').toString().trim() === nieAlumnoEliminar) || (buscar.nombre && normalizarTexto(dataDiDel[d][2]) === normalizarTexto(buscar.nombre))) {
        diSheetDel.deleteRow(d + 1); eliminado = true; break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ success: eliminado })).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "nuevo_docente") {
    var asignacionesNuevoDocente = Array.isArray(data.materias_asignadas) ? data.materias_asignadas : [];
    if (!asignacionesNuevoDocente.length && data.materia) {
      asignacionesNuevoDocente = [normalizarAsignacionDocenteGas({
        grado: data.grado_orientado || data.grado || '',
        seccion: data.seccion_orientada || data.seccion || '',
        tipo_materia: data.categoria || data.tipo_materia || '',
        materia: data.materia || '',
        escala: data.tipo_materia || data.escala || '0-10',
        especialidad: data.especialidad || '',
        es_orientado: true
      }, data.nombre || '')];
    }
    upsertDocenteBaseGas(ss, {
      nombre: data.nombre || '',
      grado_orientado: data.grado_orientado || data.grado || '',
      seccion_orientada: data.seccion_orientada || data.seccion || '',
      admin: data.admin,
      materias_asignadas: asignacionesNuevoDocente
    });
    reemplazarAsignacionesDocenteGas(ss, data.nombre || '', asignacionesNuevoDocente);
    return ContentService.createTextOutput("Exito").setMimeType(ContentService.MimeType.TEXT);

  } else if (data.tipo_post === "guardar_asignaciones_docente") {
    var asignacionesDocente = Array.isArray(data.materias_asignadas) ? data.materias_asignadas : [];
    upsertDocenteBaseGas(ss, {
      nombre: data.nombre || '',
      grado_orientado: data.grado_orientado || data.grado || '',
      seccion_orientada: data.seccion_orientada || data.seccion || '',
      admin: data.admin,
      materias_asignadas: asignacionesDocente
    });
    reemplazarAsignacionesDocenteGas(ss, data.nombre || '', asignacionesDocente);
    return ContentService.createTextOutput(JSON.stringify({ exito: true })).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "guardar_catalogo_materias") {
    var catalogo = Array.isArray(data.catalogo) ? data.catalogo : [];
    var catalogoSheet = asegurarHojaCatalogoMaterias(ss);
    var filasCatalogo = [["Nombre", "Categoria", "Escala", "RequiereEspecialidad", "EspecialidadesJSON", "Activo"]];
    for (var c = 0; c < catalogo.length; c++) {
      var itemCat = catalogo[c] || {};
      if (!itemCat.nombre) continue;
      filasCatalogo.push([
        itemCat.nombre || '',
        itemCat.categoria === 'modulo' ? 'modulo' : 'basica',
        normalizarEscalaMateriaGas(itemCat.escala),
        toBooleanGas(itemCat.requiere_especialidad) ? "true" : "false",
        safeJSONStringifyGas(itemCat.especialidades || [], []),
        "true"
      ]);
    }
    catalogoSheet.clearContents();
    catalogoSheet.getRange(1, 1, filasCatalogo.length, filasCatalogo[0].length).setValues(filasCatalogo);
    return ContentService.createTextOutput(JSON.stringify({ exito: true })).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "agregar_observacion") {
    var sheet = ss.getSheetByName("observaciones");
    if (!sheet) { sheet = ss.insertSheet("observaciones"); sheet.appendRow(["NIE", "Fecha", "Docente", "Observacion"]); }
    sheet.appendRow([data.nie, data.fecha, data.docente, data.observacion]);
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "guardar_estado_alumno") {
    var sheet = ss.getSheetByName("estado_alumnos");
    if (!sheet) { sheet = ss.insertSheet("estado_alumnos"); sheet.appendRow(["Fecha", "NIE", "Nombre", "Grado", "Seccion", "Estado", "Docente", "Observacion"]); }
    sheet.appendRow([
      data.fecha || new Date().toLocaleDateString('es-ES'),
      data.nie || '', data.nombre || '', data.grado || '',
      data.seccion || '', data.estado || '', data.docente || '', data.observacion || ''
    ]);
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "guardar_notas_alumno") {
    var resultado = guardarNotaAlumno(data);
    return ContentService.createTextOutput(JSON.stringify(resultado)).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "guardar_notas_grupo") {
    var resultado = guardarNotasGrupo(data);
    return ContentService.createTextOutput(JSON.stringify(resultado)).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "guardar_informe") {
    var resultado = guardarInforme(data.informe || data);
    return ContentService.createTextOutput(JSON.stringify(resultado)).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "eliminar_informe") {
    var resultado = eliminarInforme(data.id);
    return ContentService.createTextOutput(JSON.stringify(resultado)).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "exportar_excel_notas") {
    var resultado = exportarExcelNotasInstitucional(data);
    return ContentService.createTextOutput(JSON.stringify(resultado)).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "marcar_asistencia_alumno") {
    var resultado = marcarAsistenciaAlumno(data);
    return ContentService.createTextOutput(JSON.stringify(resultado)).setMimeType(ContentService.MimeType.JSON);

  } else if (data.tipo_post === "configurar_horario") {
    if (data.password === "747-8") {
      var props2 = PropertiesService.getScriptProperties();
      if (data.inicio) props2.setProperty('HORARIO_INICIO', data.inicio);
      if (data.fin)    props2.setProperty('HORARIO_FIN',    data.fin);
      if (data.modo_alumno !== undefined) props2.setProperty('MODO_ALUMNO_ACTIVO', data.modo_alumno ? 'true' : 'false');
      return ContentService.createTextOutput(JSON.stringify({ exito: true })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ exito: false })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ exito: false })).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO DE INFORMES
// ══════════════════════════════════════════════════════════════════════════════

function buscarAlumno(query, grado, seccion) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qNorm = normalizarTexto(query);
  var gradoNorm = normalizarTexto(grado || '');
  var seccionNorm = normalizarTexto(seccion || '');
  var resultado = [];

  function buscarEnHoja(sheet) {
    if (!sheet) return;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var nombre = (rows[i][2] || '').toString().trim();
      if (!nombre) continue;
      var nie = (rows[i][4] || '').toString().trim();
      var g = normalizarTexto(rows[i][0]);
      var s = normalizarTexto(rows[i][1]);
      if (gradoNorm && g !== gradoNorm) continue;
      if (seccionNorm && s !== seccionNorm) continue;
      var nombreNorm = normalizarTexto(nombre);
      if (nombreNorm.includes(qNorm) || nie.includes(query)) {
        resultado.push({ grado: rows[i][0], seccion: rows[i][1], nombre: nombre, sexo: rows[i][3] || '', nie: nie, telefono: rows[i][5] || '' });
      }
    }
  }

  buscarEnHoja(ss.getSheetByName("alumnos"));
  buscarEnHoja(ss.getSheetByName("di_refuerzo"));
  return resultado.slice(0, 20);
}

function obtenerAlumnoPorNIE(nie) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nieStr = (nie || '').toString().trim();
  if (!nieStr) return null;
  // Usar caché en lugar de leer la hoja cada vez
  var alumnos = getAlumnosCache(ss);
  for (var i = 0; i < alumnos.length; i++) {
    if (alumnos[i].nie === nieStr) return alumnos[i];
  }
  return null;
}

function guardarInforme(informe) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("informes");
  if (!sheet) {
    sheet = ss.insertSheet("informes");
    sheet.appendRow([
      "ID", "FechaRegistro", "Tipo", "TipoLabel", "Fecha", "Docente",
      "Grado", "Seccion", "Asunto", "Contenido", "Alumnos",
      "Periodo", "Observaciones", "Testigos"
    ]);
    var headerRange = sheet.getRange(1, 1, 1, 14);
    headerRange.setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  }

  var id = 'INF-' + new Date().getTime();
  var fila = [
    id,
    new Date(),
    informe.tipo || '',
    informe.tipoLabel || informe.tipo || '',
    informe.fecha || '',
    informe.docente || '',
    informe.grado || '',
    informe.seccion || '',
    informe.asunto || '',
    informe.descripcion || '',
    JSON.stringify(informe.alumnos || []),
    informe.periodo || '',
    informe.observaciones || '',
    informe.testigos || ''
  ];
  sheet.appendRow(fila);
  return { exito: true, id: id };
}

function obtenerHistorialInformes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("informes");
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var resultado = [];
  for (var i = rows.length - 1; i >= 1; i--) {
    if (!rows[i][0]) continue;
    resultado.push({
      id:          rows[i][0],
      fechaReg:    rows[i][1],
      tipo:        rows[i][2],
      tipoLabel:   rows[i][3],
      fecha:       rows[i][4],
      docente:     rows[i][5],
      grado:       rows[i][6],
      seccion:     rows[i][7],
      asunto:      rows[i][8],
      descripcion: rows[i][9],
      alumnos:     parseSafeJSON(rows[i][10]),
      periodo:     rows[i][11],
      observaciones: rows[i][12],
      testigos:    rows[i][13]
    });
  }
  return resultado;
}

function eliminarInforme(id) {
  if (!id) return { exito: false };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("informes");
  if (!sheet) return { exito: false };
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { exito: true };
    }
  }
  return { exito: false, error: 'No encontrado' };
}

// ══════════════════════════════════════════════════════════════════════════════
// NOTAS POR PERIODO
// ══════════════════════════════════════════════════════════════════════════════

function getEscalaNotasGas(escala) {
  return (escala === '0-5' || escala === 5 || escala === '5') ? 5 : 10;
}

function getNombreHojaNotas(escala) {
  return getEscalaNotasGas(escala) === 5 ? 'nota-tecnicos' : 'notas';
}

function asegurarEstructuraNotasSheet(sheet) {
  var headers = [
    "Grado", "Seccion", "Nombre", "NIE", "Asignatura", "Especialidad", "ClaveMateria",
    "P1_Cuaderno", "P1_Integradora", "P1_Examen",
    "P2_Cuaderno", "P2_Integradora", "P2_Examen",
    "P3_Cuaderno", "P3_Integradora", "P3_Examen",
    "P4_Cuaderno", "P4_Integradora", "P4_Examen"
  ];
  if (sheet.getLastColumn() < headers.length) {
    sheet.insertColumnsAfter(4, headers.length - sheet.getLastColumn());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function getOrCreateNotasSheet(ss, escala) {
  var sheet = ss.getSheetByName(getNombreHojaNotas(escala));
  if (!sheet) {
    sheet = ss.insertSheet(getNombreHojaNotas(escala));
    sheet.appendRow([
      "Grado", "Seccion", "Nombre", "NIE", "Asignatura", "Especialidad", "ClaveMateria",
      "P1_Cuaderno", "P1_Integradora", "P1_Examen",
      "P2_Cuaderno", "P2_Integradora", "P2_Examen",
      "P3_Cuaderno", "P3_Integradora", "P3_Examen",
      "P4_Cuaderno", "P4_Integradora", "P4_Examen"
    ]);
  } else {
    asegurarEstructuraNotasSheet(sheet);
  }
  return sheet;
}

function obtenerNotasPorGrado(grado, escala, materiaClave, seccion) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gradoNorm = normalizarTexto(grado);
  var seccionNorm = normalizarTexto(seccion || '');
  var escalaMax = getEscalaNotasGas(escala);
  var materiaClaveNorm = normalizarTexto(materiaClave || '');

  var sheet = getOrCreateNotasSheet(ss, escalaMax);

  var rows = sheet.getDataRange().getValues();
  var resultado = [];

  for (var i = 1; i < rows.length; i++) {
    var fila = rows[i];
    if (!fila[2] && !fila[3]) continue;
    if (gradoNorm && normalizarTexto(fila[0]) !== gradoNorm) continue;
    if (seccionNorm && normalizarTexto(fila[1]) !== seccionNorm) continue;
    if (materiaClaveNorm && normalizarTexto(fila[6] || fila[4] || '') !== materiaClaveNorm) continue;

    var est = {
      grado:          fila[0]  || '',
      seccion:        fila[1]  || '',
      nombre:         fila[2]  || '',
      nie:            fila[3]  || '',
      asignatura:     fila[4]  || '',
      especialidad:   fila[5]  || '',
      materia_clave:  fila[6]  || '',
      p1_cuaderno:    parseNotaSegura(fila[7]),
      p1_integradora: parseNotaSegura(fila[8]),
      p1_examen:      parseNotaSegura(fila[9]),
      p2_cuaderno:    parseNotaSegura(fila[10]),
      p2_integradora: parseNotaSegura(fila[11]),
      p2_examen:      parseNotaSegura(fila[12]),
      p3_cuaderno:    parseNotaSegura(fila[13]),
      p3_integradora: parseNotaSegura(fila[14]),
      p3_examen:      parseNotaSegura(fila[15]),
      p4_cuaderno:    parseNotaSegura(fila[16]),
      p4_integradora: parseNotaSegura(fila[17]),
      p4_examen:      parseNotaSegura(fila[18])
    };
    est.periodo1  = calcularNotaPeriodoGas(est.p1_cuaderno, est.p1_integradora, est.p1_examen, escalaMax);
    est.periodo2  = calcularNotaPeriodoGas(est.p2_cuaderno, est.p2_integradora, est.p2_examen, escalaMax);
    est.periodo3  = calcularNotaPeriodoGas(est.p3_cuaderno, est.p3_integradora, est.p3_examen, escalaMax);
    est.periodo4  = calcularNotaPeriodoGas(est.p4_cuaderno, est.p4_integradora, est.p4_examen, escalaMax);
    est.notaFinal = Math.min(escalaMax, (est.periodo1 + est.periodo2 + est.periodo3 + est.periodo4) / 4);
    est.escala = escalaMax;
    resultado.push(est);
  }

  return resultado;
}

function guardarNotaAlumno(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var escalaMax = getEscalaNotasGas(data.escala || data.escala_notas || data.tipo_materia);
  var sheet = getOrCreateNotasSheet(ss, escalaMax);

  var rows = sheet.getDataRange().getValues();
  var gradoNorm  = normalizarTexto(data.grado  || '');
  var seccionNorm = normalizarTexto(data.seccion || '');
  var nombreNorm = normalizarTexto(data.nombre || '');
  var nieStr     = normalizarTexto(data.nie || '');
  var materiaClave = (data.materia_clave || data.asignatura || data.materia || '').toString().trim();
  var materiaClaveNorm = normalizarTexto(materiaClave);

  var notasValores = [
    toNum(data.p1_cuaderno, escalaMax),    toNum(data.p1_integradora, escalaMax), toNum(data.p1_examen, escalaMax),
    toNum(data.p2_cuaderno, escalaMax),    toNum(data.p2_integradora, escalaMax), toNum(data.p2_examen, escalaMax),
    toNum(data.p3_cuaderno, escalaMax),    toNum(data.p3_integradora, escalaMax), toNum(data.p3_examen, escalaMax),
    toNum(data.p4_cuaderno, escalaMax),    toNum(data.p4_integradora, escalaMax), toNum(data.p4_examen, escalaMax)
  ];

  var filaEncontrada = -1;
  for (var i = 1; i < rows.length; i++) {
    var filaGrado  = normalizarTexto(rows[i][0]);
    var filaSeccion = normalizarTexto(rows[i][1]);
    var filaFila   = normalizarTexto(rows[i][2]);
    var filaNie    = normalizarTexto(rows[i][3] || '');
    var filaMateria = normalizarTexto(rows[i][6] || rows[i][4] || '');
    var coincide = filaGrado === gradoNorm &&
      filaSeccion === seccionNorm &&
      filaMateria === materiaClaveNorm &&
      (filaNie !== '' && filaNie === nieStr || filaFila === nombreNorm);
    if (coincide) { filaEncontrada = i + 1; break; }
  }

  if (filaEncontrada > 0) {
    sheet.getRange(filaEncontrada, 5, 1, 3).setValues([[
      data.asignatura || data.materia || '',
      data.especialidad || '',
      materiaClave
    ]]);
    sheet.getRange(filaEncontrada, 8, 1, 12).setValues([notasValores]);
  } else {
    sheet.appendRow([
      data.grado   || '',
      data.seccion || '',
      data.nombre  || '',
      data.nie     || '',
      data.asignatura || data.materia || '',
      data.especialidad || '',
      materiaClave
    ].concat(notasValores));
  }

  return { exito: true };
}

function guardarNotasGrupo(data) {
  var notas = data.notas || [];
  if (!Array.isArray(notas)) return { exito: false, error: 'Notas inválidas' };

  // Optimización: una sola lectura de hoja y búsqueda por índice (evita O(n*m))
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var escalaMax = getEscalaNotasGas(data.escala || data.escala_notas || data.tipo_materia || '0-10');
  var sheet = getOrCreateNotasSheet(ss, escalaMax);
  asegurarEstructuraNotasSheet(sheet);

  var rows = sheet.getDataRange().getValues();
  var idx = {};

  // Índice por (grado|seccion|materia|nie)
  for (var r = 1; r < rows.length; r++) {
    var fila = rows[r];
    var g = normalizarTexto(fila[0]);
    var s = normalizarTexto(fila[1]);
    var m = normalizarTexto(fila[6] || fila[4] || '');
    var nie = normalizarTexto(fila[3] || '');
    if (!g || !s || !m || !nie) continue;
    idx[g + '|' + s + '|' + m + '|' + nie] = r + 1; // 1-based
  }

  function norm(v) { return normalizarTexto((v || '').toString().trim()); }

  var guardadas = 0;
  var nuevos = [];

  for (var i = 0; i < notas.length; i++) {
    var item = notas[i] || {};
    var grado = item.grado || data.grado || '';
    var seccion = item.seccion || data.seccion || '';
    var asignatura = item.asignatura || data.asignatura || data.materia || '';
    var especialidad = item.especialidad || data.especialidad || '';
    var materiaClave = (item.materia_clave || data.materia_clave || asignatura || '').toString().trim();
    var nieStr = normalizarTexto(item.nie || '');
    var nombre = item.nombre || '';
    var key = norm(grado) + '|' + norm(seccion) + '|' + norm(materiaClave) + '|' + nieStr;

    var notasValores = [
      toNum(item.p1_cuaderno, escalaMax),    toNum(item.p1_integradora, escalaMax), toNum(item.p1_examen, escalaMax),
      toNum(item.p2_cuaderno, escalaMax),    toNum(item.p2_integradora, escalaMax), toNum(item.p2_examen, escalaMax),
      toNum(item.p3_cuaderno, escalaMax),    toNum(item.p3_integradora, escalaMax), toNum(item.p3_examen, escalaMax),
      toNum(item.p4_cuaderno, escalaMax),    toNum(item.p4_integradora, escalaMax), toNum(item.p4_examen, escalaMax)
    ];

    var filaEncontrada = idx[key] || -1;
    if (filaEncontrada > 0) {
      // Actualizar columnas Asignatura/Especialidad/ClaveMateria + notas
      sheet.getRange(filaEncontrada, 5, 1, 3).setValues([[asignatura, especialidad, materiaClave]]);
      sheet.getRange(filaEncontrada, 8, 1, 12).setValues([notasValores]);
    } else {
      nuevos.push([grado, seccion, nombre, nieStr, asignatura, especialidad, materiaClave].concat(notasValores));
      // Reservar para evitar duplicados dentro del mismo batch
      idx[key] = rows.length + nuevos.length; // aprox (se corrige al escribir)
    }
    guardadas++;
  }

  if (nuevos.length) {
    var start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, nuevos.length, nuevos[0].length).setValues(nuevos);
  }

  return {
    exito: true,
    guardadas: guardadas,
    hoja: getNombreHojaNotas(escalaMax)
  };
}

function toNum(val, escala) {
  if (val === null || val === undefined || val === '') return '';
  var n = parseFloat(val);
  if (isNaN(n)) return '';
  var max = getEscalaNotasGas(escala);
  return Math.min(max, Math.max(0, n));
}

function calcularNotaPeriodoGas(cuaderno, integradora, examen, escala) {
  var max = getEscalaNotasGas(escala);
  var nota = ((parseFloat(cuaderno)    || 0) * 0.35)
           + ((parseFloat(integradora) || 0) * 0.35)
           + ((parseFloat(examen)      || 0) * 0.30);
  return Math.min(max, Math.max(0, nota));
}

function parseNotaSegura(valor) {
  if (valor === '' || valor === null || valor === undefined) return null;
  var n = parseFloat(valor);
  return isNaN(n) ? null : n;
}

// ══════════════════════════════════════════════════════════════════════════════
// EXPORTACIÓN EXCEL INSTITUCIONAL
// ══════════════════════════════════════════════════════════════════════════════

function exportarExcelNotasInstitucional(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grado    = params.grado    || '';
  var seccion  = params.seccion  || '';
  var asignatura = params.asignatura || '';
  var materia_clave = params.materia_clave || params.clave || '';
  var docente  = params.docente  || '';
  var anio     = params.anio     || new Date().getFullYear();
  var escalaMax = getEscalaNotasGas(params.escala || params.escala_notas || params.tipo_materia);
  var notaAprobado = escalaMax === 5 ? 3 : 6;
  var notaRiesgo = escalaMax === 5 ? 2.5 : 5;

  var notasData = obtenerNotasPorGrado(grado, escalaMax, materia_clave, seccion);
  if (seccion) {
    var secNorm = normalizarTexto(seccion);
    notasData = notasData.filter(function(n){ return normalizarTexto(n.seccion) === secNorm; });
  }

  var nombreHoja = 'Cuadro_' + grado.replace(/\s/g,'') + '_' + seccion + '_' + new Date().getTime();
  var tempSheet = ss.insertSheet(nombreHoja);

  try {
    tempSheet.setColumnWidth(1, 35);
    tempSheet.setColumnWidth(2, 75);
    tempSheet.setColumnWidth(3, 210);
    tempSheet.setColumnWidth(4, 45);
    for (var c = 5; c <= 20; c++) {
      tempSheet.setColumnWidth(c, c % 4 === 0 ? 55 : 48);
    }
    tempSheet.setColumnWidth(21, 62);

    var totalCols = 21;

    tempSheet.getRange(1, 1).setValue("MINISTERIO DE EDUCACIÓN CIENCIA Y TECNOLOGÍA");
    tempSheet.getRange(1, Math.ceil(totalCols / 2) - 2).setValue("INSTITUTO NACIONAL DE MERCEDES UMAÑA");
    tempSheet.getRange(1, totalCols - 2).setValue("DIRECCIÓN NACIONAL DE EDUCACIÓN");

    tempSheet.getRange(1, 1, 1, 4).setFontSize(8).setFontWeight('bold');
    tempSheet.getRange(1, Math.ceil(totalCols / 2) - 2, 1, 5).setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center');
    tempSheet.getRange(1, totalCols - 2, 1, 3).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('right');

    tempSheet.getRange(2, 1).setValue("DIRECCIÓN  NACIONAL DE EDUCACIÓN");
    tempSheet.getRange(2, Math.ceil(totalCols / 2) - 2).setValue("CUADRO DE REGISTRO DE EVALUACIÓN DE LOS  APRENDIZAJES");
    tempSheet.getRange(2, Math.ceil(totalCols / 2) - 2, 1, 6).setHorizontalAlignment('center').setFontSize(10).setFontWeight('bold');

    tempSheet.getRange(3, 1).setValue("EL SALVADOR");
    tempSheet.getRange(3, Math.ceil(totalCols / 2) - 2).setValue("POR ASIGNATURA Y PERIODOS");
    tempSheet.getRange(3, Math.ceil(totalCols / 2) - 2, 1, 4).setHorizontalAlignment('center').setFontSize(10).setFontWeight('bold');

    tempSheet.getRange(4, 1).setValue("EDUCACION MEDIA:");
    tempSheet.getRange(4, 1, 1, 4).setBackground('#000000').setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);

    var datosGrupo = 'Sección:  ' + grado + ' ' + seccion + '  ' + anio +
                     '          Asignatura:     ' + asignatura +
                     '          Profesor:' + docente +
                     '          Escala: 0-' + escalaMax;
    tempSheet.getRange(4, 6).setValue(datosGrupo);
    tempSheet.getRange(4, 6, 1, totalCols - 5).setFontSize(9).setFontWeight('bold');

    var headerCells = [
      [5, 1, 'N°'],
      [5, 2, 'NIE'],
      [5, 3, 'NOMBRE DEL ALUMNO/A'],
      [5, 4, 'GE\nNE\nRO']
    ];
    headerCells.forEach(function(hc) {
      var r = tempSheet.getRange(hc[0], hc[1]);
      r.setValue(hc[2]).setBackground('#000000').setFontColor('#ffffff')
       .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
       .setVerticalAlignment('middle').setWrap(true);
    });

    var periodos = ['Primer Periodo', 'Segundo Período', 'Tercer Período', 'Cuarto Período'];
    var colBase = 5;
    periodos.forEach(function(pLabel) {
      var r = tempSheet.getRange(5, colBase, 1, 4);
      r.merge().setValue(pLabel)
       .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold')
       .setFontSize(8).setHorizontalAlignment('center');
      colBase += 4;
    });

    var notaFinalRange = tempSheet.getRange(5, colBase);
    notaFinalRange.setValue('Nota\nFinal').setBackground('#1a3a5c').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center')
      .setVerticalAlignment('middle').setWrap(true);

    colBase = 5;
    periodos.forEach(function() {
      tempSheet.getRange(6, colBase, 1, 3).merge().setValue('Actividades')
        .setBackground('#2d5a8e').setFontColor('#ffffff').setFontSize(8)
        .setHorizontalAlignment('center');
      tempSheet.getRange(6, colBase + 3).setValue('Pro\nme\ndio')
        .setBackground('#1a3a5c').setFontColor('#ffffff').setFontSize(7)
        .setHorizontalAlignment('center').setWrap(true);
      colBase += 4;
    });

    colBase = 5;
    periodos.forEach(function() {
      var pcts = ['35%', '35%', '30%'];
      pcts.forEach(function(pct, pi) {
        tempSheet.getRange(7, colBase + pi).setValue(pct)
          .setBackground('#3b78c4').setFontColor('#ffffff').setFontSize(7)
          .setHorizontalAlignment('center').setFontWeight('bold');
      });
      tempSheet.getRange(7, colBase + 3).setValue('')
        .setBackground('#1a3a5c');
      colBase += 4;
    });

    [1, 2, 3, 4].forEach(function(col) {
      tempSheet.getRange(6, col, 2, 1).merge()
        .setBackground('#000000').setFontColor('#ffffff');
    });

    var ROW_START = 8;

    notasData.forEach(function(est, idx) {
      var row = ROW_START + idx;
      var bg = idx % 2 === 0 ? '#f8fafc' : '#ffffff';

      tempSheet.getRange(row, 1).setValue(idx + 1);
      tempSheet.getRange(row, 2).setValue(est.nie || '');
      tempSheet.getRange(row, 3).setValue(est.nombre || '');
      tempSheet.getRange(row, 4).setValue(est.sexo ? est.sexo.toString().substring(0,1).toUpperCase() : '');

      var notas = [
        [est.p1_cuaderno, est.p1_integradora, est.p1_examen, est.periodo1],
        [est.p2_cuaderno, est.p2_integradora, est.p2_examen, est.periodo2],
        [est.p3_cuaderno, est.p3_integradora, est.p3_examen, est.periodo3],
        [est.p4_cuaderno, est.p4_integradora, est.p4_examen, est.periodo4]
      ];

      var colN = 5;
      notas.forEach(function(p) {
        var v1 = p[0] !== null && p[0] !== undefined ? p[0] : '';
        tempSheet.getRange(row, colN).setValue(v1);

        var v2 = p[1] !== null && p[1] !== undefined ? p[1] : '';
        tempSheet.getRange(row, colN + 1).setValue(v2);

        var v3 = p[2] !== null && p[2] !== undefined ? p[2] : '';
        tempSheet.getRange(row, colN + 2).setValue(v3);

        var prom = Math.min(escalaMax, p[3] || 0);
        var promCell = tempSheet.getRange(row, colN + 3);
        promCell.setValue(prom > 0 ? Math.round(prom * 100) / 100 : 0);

        var promBg = (prom < notaRiesgo && prom >= 0) ? '#e53e3e' : (prom < notaAprobado ? '#fff2cc' : bg);
        var promFg = (prom < notaRiesgo && prom >= 0) ? '#ffffff' : '#1e293b';
        promCell.setBackground(promBg).setFontColor(promFg).setFontWeight('bold');

        colN += 4;
      });

      var notaFinal = Math.min(escalaMax, est.notaFinal || 0);
      notaFinal = Math.round(notaFinal * 100) / 100;
      var nfBg = notaFinal < notaRiesgo ? '#e53e3e' : '#1a3a5c';
      var nfFg = '#ffffff';
      tempSheet.getRange(row, colN)
        .setValue(notaFinal > 0 ? notaFinal : 0)
        .setBackground(nfBg).setFontColor(nfFg).setFontWeight('bold');

      tempSheet.getRange(row, 1, 1, totalCols).setFontSize(9);
      tempSheet.getRange(row, 1).setHorizontalAlignment('center').setFontWeight('bold').setBackground('#f1f5f9');
      tempSheet.getRange(row, 2).setHorizontalAlignment('center').setFontSize(8);
      tempSheet.getRange(row, 3).setHorizontalAlignment('left');
      tempSheet.getRange(row, 4).setHorizontalAlignment('center');

      tempSheet.getRange(row, 1, 1, totalCols)
        .setBorder(true, true, true, true, true, true, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
    });

    if (notasData.length > 0) {
      var totRow = ROW_START + notasData.length;
      tempSheet.getRange(totRow, 1, 1, 4)
        .merge().setValue('TOTAL ALUMNOS: ' + notasData.length)
        .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
    }

    tempSheet.setRowHeight(5, 42);
    tempSheet.setRowHeight(6, 22);
    tempSheet.setRowHeight(7, 18);
    for (var r = ROW_START; r < ROW_START + notasData.length; r++) {
      tempSheet.setRowHeight(r, 20);
    }

    tempSheet.setFrozenRows(7);
    tempSheet.setFrozenColumns(4);

    SpreadsheetApp.flush();

    var nombreArchivo = 'Cuadro_Notas_' + grado.replace(/\s+/g,'_') + '_' + seccion + '_' + asignatura.replace(/\s+/g,'_') + '_' + anio + '_0-' + escalaMax + '.xlsx';

    var driveFile = DriveApp.getFileById(ss.getId());
    var parents = driveFile.getParents();
    var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

    var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
                    '/export?format=xlsx&gid=' + tempSheet.getSheetId() +
                    '&access_token=' + ScriptApp.getOAuthToken();

    var response = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    var blob = response.getBlob().setName(nombreArchivo);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    ss.deleteSheet(tempSheet);

    return { exito: true, url: file.getDownloadUrl(), nombre: nombreArchivo, fileId: file.getId() };

  } catch (err) {
    try { ss.deleteSheet(tempSheet); } catch(e2) {}
    return { exito: false, error: err.toString() };
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// SISTEMA DE INACTIVIDAD DE DOCENTES
// ══════════════════════════════════════════════════════════════════════════════

function updateDocenteStatus(docente, status, timestamp) {
  if (timestamp === undefined) timestamp = null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("EstadosDocentes");
  if (!sheet) {
    sheet = ss.insertSheet("EstadosDocentes");
    sheet.appendRow(["Docente", "Status", "ultima_actividad"]);
  }
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === docente) {
      sheet.getRange(i + 1, 2).setValue(status);
      sheet.getRange(i + 1, 3).setValue(timestamp || new Date().getTime());
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([docente, status, timestamp || new Date().getTime()]);
}

function getDocentesStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("EstadosDocentes");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var ahora = new Date().getTime();
  var TIEMPO_INACTIVIDAD = 5 * 60 * 1000;
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var docente = data[i][0];
    var status = data[i][1];
    var ultimaActividad = data[i][2] || 0;
    if (status === "online" && (ahora - ultimaActividad) > TIEMPO_INACTIVIDAD) {
      status = "offline";
      sheet.getRange(i + 1, 2).setValue(status);
    }
    result.push({ docente: docente, status: status, ultima_actividad: ultimaActividad });
  }
  return result;
}

// ══════════════════════════════════════════════════════════════════════════════
// CONTEO DE AUSENCIAS
// ══════════════════════════════════════════════════════════════════════════════

function incrementarConteoAusencia(estudianteNombre, grado, seccion) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("conteo_ausencias");
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var encontrado = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === estudianteNombre && data[i][2] === grado && data[i][3] === seccion) {
      sheet.getRange(i + 1, 5).setValue((parseInt(data[i][4]) || 0) + 1);
      encontrado = true;
      break;
    }
  }
  if (!encontrado) {
    var nie = '';
    var alumnosSheet = ss.getSheetByName("alumnos");
    if (alumnosSheet) {
      var alumnosData = alumnosSheet.getDataRange().getValues();
      for (var j = 1; j < alumnosData.length; j++) {
        if (alumnosData[j][2] === estudianteNombre && alumnosData[j][0] === grado && alumnosData[j][1] === seccion) { nie = alumnosData[j][4] || ''; break; }
      }
    }
    if (!nie) {
      var diSheet = ss.getSheetByName("di_refuerzo");
      if (diSheet) {
        var diData = diSheet.getDataRange().getValues();
        for (var k = 1; k < diData.length; k++) {
          if (diData[k][2] === estudianteNombre && diData[k][0] === grado && diData[k][1] === seccion) { nie = diData[k][4] || ''; break; }
        }
      }
    }
    sheet.appendRow([estudianteNombre, nie, grado, seccion, 1]);
  }
}

function getEstudiantesEnPeligro(grado, seccion) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("conteo_ausencias");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === grado && data[i][3] === seccion && parseInt(data[i][4]) >= 30) {
      result.push({ nombre: data[i][0], nie: data[i][1], conteo: parseInt(data[i][4]) });
    }
  }
  return result;
}

// ══════════════════════════════════════════════════════════════════════════════
// ASISTENCIA AUTÓNOMA DE ALUMNOS
// ══════════════════════════════════════════════════════════════════════════════

function marcarAsistenciaAlumno(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("asistencia_alumnos");
  if (!sheet) {
    sheet = ss.insertSheet("asistencia_alumnos");
    sheet.appendRow(["Fecha", "NIE", "Nombre", "Grado", "Seccion", "Estado", "Hora", "Justificante"]);
    var hdr = sheet.getRange(1,1,1,8);
    hdr.setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');
  }

  var nie = (data.nie || '').toString().trim();
  if (!nie) return { exito: false, error: 'NIE requerido' };

  var alumno = obtenerAlumnoPorNIE(nie);
  if (!alumno) return { exito: false, error: 'Alumno no encontrado' };

  var ahora = new Date();
  var hoy   = ahora.toLocaleDateString('es-SV');
  var hora  = ahora.toLocaleTimeString('es-SV', { hour:'2-digit', minute:'2-digit' });

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowFecha = rows[i][0] ? new Date(rows[i][0]).toLocaleDateString('es-SV') : '';
    if (rows[i][1] == nie && rowFecha === hoy) {
      return { exito: false, ya_marcado: true, estado: rows[i][5], hora: rows[i][6] };
    }
  }

  var estado = data.estado || 'presente';
  sheet.appendRow([ahora, nie, alumno.nombre, alumno.grado, alumno.seccion, estado, hora, data.justificante || '']);

  if (estado === 'permiso') {
    var permisosSheet = ss.getSheetByName("permisos");
    if (!permisosSheet) {
      permisosSheet = ss.insertSheet("permisos");
      permisosSheet.appendRow(["Fecha","Grado","Seccion","Docente","Estudiante","Sexo","NIE","Justificante"]);
    }
    permisosSheet.appendRow([ahora, alumno.grado, alumno.seccion, 'ALUMNO-AUTO', alumno.nombre, alumno.sexo||'', nie, data.justificante||'']);
  }

  return { exito: true, nombre: alumno.nombre, grado: alumno.grado, seccion: alumno.seccion, hora: hora, estado: estado };
}

function obtenerExpedienteAlumno(nie) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var nieStr = (nie || '').toString().trim();
  if (!nieStr) return null;

  var alumno = obtenerAlumnoPorNIE(nieStr);
  if (!alumno) return null;

  var faltasTotales = 0;
  var conteoSheet = ss.getSheetByName("conteo_ausencias");
  if (conteoSheet) {
    var conteoRows = conteoSheet.getDataRange().getValues();
    for (var i = 1; i < conteoRows.length; i++) {
      if ((conteoRows[i][1]||'').toString().trim() === nieStr ||
          normalizarTexto(conteoRows[i][0]) === normalizarTexto(alumno.nombre)) {
        faltasTotales = parseInt(conteoRows[i][4]||0) || 0; break;
      }
    }
  }

  var permisos = [];
  var permisosSheet = ss.getSheetByName("permisos");
  if (permisosSheet) {
    var permRows = permisosSheet.getDataRange().getValues();
    var conCols = permRows[0] && permRows[0].length >= 8 && permRows[0][5]==='Sexo';
    for (var j = 1; j < permRows.length; j++) {
      if (!permRows[j][0]) continue;
      var rowNie = conCols ? (permRows[j][6]||'').toString().trim() : '';
      var rowNom = (permRows[j][4]||'').toString().trim();
      if (rowNie === nieStr || normalizarTexto(rowNom) === normalizarTexto(alumno.nombre)) {
        permisos.push({
          fecha: permRows[j][0], grado: permRows[j][1], seccion: permRows[j][2],
          docente: permRows[j][3]||'',
          justificante: conCols ? (permRows[j][7]||'') : (permRows[j][5]||'')
        });
      }
    }
  }

  var asistenciaPropia = [];
  var asistSheet = ss.getSheetByName("asistencia_alumnos");
  if (asistSheet) {
    var asistRows = asistSheet.getDataRange().getValues();
    for (var k = 1; k < asistRows.length; k++) {
      if ((asistRows[k][1]||'').toString().trim() === nieStr) {
        asistenciaPropia.push({
          fecha: asistRows[k][0], estado: asistRows[k][5],
          hora: asistRows[k][6], justificante: asistRows[k][7]||''
        });
      }
    }
  }

  var notas = null;
  var hojasNotas = [
    { sheet: ss.getSheetByName("notas"), escala: 10 },
    { sheet: ss.getSheetByName("nota-tecnicos"), escala: 5 }
  ];
  for (var hn = 0; hn < hojasNotas.length && !notas; hn++) {
    var notasSheet = hojasNotas[hn].sheet;
    if (!notasSheet) continue;
    asegurarEstructuraNotasSheet(notasSheet);
    var escalaNotasAlumno = hojasNotas[hn].escala;
    var notasRows = notasSheet.getDataRange().getValues();
    for (var n = 1; n < notasRows.length; n++) {
      if ((notasRows[n][3]||'').toString().trim() === nieStr ||
          normalizarTexto(notasRows[n][2]) === normalizarTexto(alumno.nombre)) {
        var baseNota = 7;
        notas = {
          escala: escalaNotasAlumno,
          p1: calcularNotaPeriodoGas(parseNotaSegura(notasRows[n][baseNota]), parseNotaSegura(notasRows[n][baseNota + 1]), parseNotaSegura(notasRows[n][baseNota + 2]), escalaNotasAlumno),
          p2: calcularNotaPeriodoGas(parseNotaSegura(notasRows[n][baseNota + 3]), parseNotaSegura(notasRows[n][baseNota + 4]), parseNotaSegura(notasRows[n][baseNota + 5]), escalaNotasAlumno),
          p3: calcularNotaPeriodoGas(parseNotaSegura(notasRows[n][baseNota + 6]), parseNotaSegura(notasRows[n][baseNota + 7]), parseNotaSegura(notasRows[n][baseNota + 8]), escalaNotasAlumno),
          p4: calcularNotaPeriodoGas(parseNotaSegura(notasRows[n][baseNota + 9]), parseNotaSegura(notasRows[n][baseNota + 10]), parseNotaSegura(notasRows[n][baseNota + 11]), escalaNotasAlumno)
        };
        notas.final = Math.min(escalaNotasAlumno, (notas.p1 + notas.p2 + notas.p3 + notas.p4) / 4);
        break;
      }
    }
  }

  return {
    alumno: alumno,
    faltas_totales: faltasTotales,
    permisos: permisos,
    asistencia_propia: asistenciaPropia,
    notas: notas
  };
}
