/*******************************************************************************
 * REPORTE ANUAL DE ASISTENCIA -> EXCEL (.xlsx) - Google Apps Script
 * Variante anual de asistencia.js.
 *
 * QUÉ HACE:
 *  - Lee TODOS los archivos de la carpeta de ENTRADA ANUAL (no solo el más
 *    reciente, como hace asistencia.js):
 *      ReporteAsistenciaEmpresa*.xls*     (todos, uno o muchos)
 *      ReporteInasistenciaEmpresa*.xls*   (todos, uno o muchos)
 *      ReporteRegistroFallido*.xls*       (todos, uno o muchos)
 *      Trabajadores_*.xls*                (lista maestra, se usa el más reciente)
 *      calendario*.xlsx                   (hoja "festivos", se usa el más reciente)
 *    Si el maestro o el calendario no están en la carpeta anual, se buscan en
 *    la carpeta de insumos diaria (ID_CARPETA_INSUMOS_RESPALDO).
 *    Si los reportes se traslapan (el mismo día viene en dos archivos), no se
 *    duplica nada: todo se consolida por RUT + fecha.
 *  - Arma en un Google Sheet TEMPORAL las hojas:
 *      "Reporte anual"      matriz persona x día de todo el año
 *      "Resumen"            totales por persona (asistencias, retardos,
 *                           faltas, vacaciones, etc.) + faltas por mes
 *      "Faltas"             detalle de cada falta sin justificar del año
 *      "Retardos"           detalle de cada retardo del año
 *      "Intentos fallidos"  días que cuentan como asistencia por intento de
 *                           checada fallido (trazabilidad)
 *      "Leyenda"
 *  - Exporta ese Sheet como EXCEL (.xlsx) a la carpeta de SALIDA ANUAL y
 *    borra (manda a la papelera) el Sheet temporal. Si ya existía un .xlsx
 *    con el mismo nombre en la carpeta de salida, el anterior se manda a la
 *    papelera y se deja el nuevo.
 *
 * REGLAS DE PRIORIDAD (las mismas de asistencia.js), por RUT + fecha:
 *   1. Si hay un registro REAL en ReporteAsistenciaEmpresa ese día, en
 *      CUALQUIER recinto -> asistencia (A / A-N). Se ignora cualquier
 *      inasistencia de ese día para esa persona (p.ej. la de otro recinto).
 *      Si checó en varios recintos, se toma la checada con menos retardo.
 *   2. Si NO hay asistencia real pero SÍ hay un intento en
 *      ReporteRegistroFallido ese día -> asistencia (A), con nota en la
 *      celda con el detalle del intento.
 *   3. Si no aplica ninguna de las anteriores, se usa el Motivo de la
 *      inasistencia (V / L / P / I / FJ). Si la persona tiene varias filas
 *      de inasistencia ese día (distintos recintos) y ALGUNA trae
 *      justificación, gana la justificación.
 *   => Solo es FALTA (F) si y solo si: no checó en ningún recinto, no tuvo
 *      intento de checada y ninguna de sus inasistencias de ese día está
 *      justificada.
 *   Los días festivos (hoja "festivos" del calendario) se marcan DF.
 *
 * CONFIGURACIÓN (bloque CONFIG_ANUAL):
 *  1. ID_CARPETA_ENTRADA_ANUAL: carpeta donde subes TODOS los reportes de Buk
 *     del año.
 *  2. ID_CARPETA_SALIDA_ANUAL: carpeta donde se guarda el .xlsx generado
 *     (puede ser la misma que la de entrada; el .xlsx de salida no choca con
 *     los patrones de búsqueda de los reportes).
 *  3. ANIO: año a reportar. Si se deja en null, se toma el año de la fecha
 *     más reciente encontrada en los reportes.
 *  4. (Solo si algún archivo es un .xls/.xlsx binario que no abre directo)
 *     habilitar el servicio avanzado "Drive API" (identificador "Drive"),
 *     igual que en asistencia.js.
 *
 * CÓMO SE EJECUTA:
 *  - `diagnosticarAnual`: no genera nada; lista qué archivos encontró y
 *    cuántas filas/fechas lee de cada uno.
 *  - `generarExcelAnual`: genera el .xlsx.
 *  - `crearTriggerAnualSemanal` (opcional, correr UNA vez): lo ejecuta solo
 *    cada lunes.
 *
 * NOTA: todas las funciones y variables internas viven dentro del objeto
 * `AsistenciaAnual`, así que este archivo puede convivir en el MISMO proyecto
 * de Apps Script que asistencia.js sin chocar nombres (main, CONFIG,
 * parseFecha, onOpen, etc.).
 *
 * NOTA DE RENDIMIENTO: Apps Script corta la ejecución a los 6 minutos. Abrir
 * cada archivo cuesta unos segundos, así que conviene exportar de Buk por
 * mes o por rango (unos pocos archivos grandes) en vez de un archivo por día.
 ******************************************************************************/

// ============================================================================
// CONFIG - AJUSTA ESTOS VALORES
// ============================================================================

var CONFIG_ANUAL = {
  // Carpeta donde subes TODOS los reportes del año (asistencias,
  // inasistencias y registros fallidos de Buk).
  ID_CARPETA_ENTRADA_ANUAL: 'PON_AQUI_EL_ID_DE_LA_CARPETA_DE_REPORTES_ANUALES',
  // Carpeta donde se guardará el Excel generado.
  ID_CARPETA_SALIDA_ANUAL: 'PON_AQUI_EL_ID_DE_LA_CARPETA_DE_SALIDA',
  // Si el maestro de trabajadores o el calendario no están en la carpeta
  // anual, se buscan aquí (carpeta "Informes_de_asistencia" de asistencia.js).
  ID_CARPETA_INSUMOS_RESPALDO: '1J8HKFr8BxSiKZ7xnH2neuYIspeaU1jQ8',
  // Año a reportar (ej. 2026). null = año de la fecha más reciente en los datos.
  ANIO: null,
  // true = columnas del 1-ene al 31-dic; false = del 1-ene a la última fecha
  // con datos.
  HASTA_FIN_DE_ANIO: false,
  TOLERANCIA_RETARDO_MIN: 10,
  // Nombre del archivo; {ANIO} se reemplaza por el año.
  NOMBRE_ARCHIVO: 'Asistencia_anual_{ANIO}',
  // Incluir en el reporte a personas que aparecen en los reportes de Buk
  // pero ya no están en el maestro (bajas durante el año).
  INCLUIR_FUERA_DE_MAESTRO: true,
  ENVIAR_CORREO: false,
  CORREOS: [
    'ccarbajal@abcsc.mx'
  ],
  ASUNTO_CORREO: 'Asistencia anual'
};

// ============================================================================
// PUNTOS DE ENTRADA (visibles en el selector de funciones del editor)
// ============================================================================

function generarExcelAnual() {
  AsistenciaAnual.generar();
}

function diagnosticarAnual() {
  AsistenciaAnual.diagnosticar();
}

function crearTriggerAnualSemanal() {
  ScriptApp.newTrigger('generarExcelAnual').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
}

// ============================================================================
// IMPLEMENTACIÓN
// ============================================================================

var AsistenciaAnual = (function () {

  var RE_ASISTENCIAS = /ReporteAsistenciaEmpresa.*\.xls/i;
  var RE_INASISTENCIAS = /ReporteInasistenciaEmpresa.*\.xls/i;
  var RE_FALLIDOS = /ReporteRegistroFallido.*\.xls/i;
  var RE_MAESTRO = /Trabajadores_.*\.xls/i;
  var RE_CALENDARIO = /calendario.*\.xlsx?/i;

  var HOJAS = {
    REPORTE: 'Reporte anual',
    RESUMEN: 'Resumen',
    FALTAS: 'Faltas',
    RETARDOS: 'Retardos',
    FALLIDOS: 'Intentos fallidos',
    LEYENDA: 'Leyenda'
  };

  var COLOR_HEADER_FONDO = '#1F4E78';
  var COLOR_HEADER_TEXTO = '#FFFFFF';
  var COLOR_MES_FONDO = '#D6E4F0';
  var COLOR_RETARDO = '#FFF2CC';
  var COLOR_FALTA = '#FCE4E4';
  var COLOR_JUSTIFICADA = '#E2EFDA';
  var COLOR_FALLIDO = '#DDEBF7';
  var COLOR_DF = '#D9D9D9';

  var MOTIVO_A_CODIGO = {
    '-': 'F',
    'V': 'V',
    'L': 'L',
    'P': 'P',
    'I': 'I',
    'FJ': 'FJ'
  };

  var LEYENDA = [
    ['A', 'ASISTENCIA'],
    ['A-N', 'ASISTENCIA CON RETARDO (N = minutos tarde, ya restada la tolerancia)'],
    ['F', 'FALTA (SIN JUSTIFICAR, SIN CHECADA EN NINGÚN RECINTO NI INTENTO FALLIDO)'],
    ['V', 'VACACIONES'],
    ['I', 'INCAPACIDAD'],
    ['P', 'PERMISO'],
    ['L', 'LICENCIA (MATERNIDAD / PATERNIDAD / IMSS)'],
    ['FJ', 'FALTA JUSTIFICADA'],
    ['DF', 'DÍA FESTIVO'],
    ['-', 'SIN INCIDENCIA (SIN DATOS ESE DÍA)'],
    ['(celda azul)', 'ASISTENCIA INFERIDA POR INTENTO DE CHECADA FALLIDO (ver nota de la celda)']
  ];

  var COLS_META = ['RUT', 'Nombre completo', 'Área', 'Horario Turno', 'Recinto', 'Localidad', 'Supervisor', 'En maestro'];
  var FILA_MESES = 1;
  var FILA_ENCABEZADO = 2;
  var PRIMERA_FILA_DATOS = 3;

  var MESES_CORTOS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  var MESES_LARGOS = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
    'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

  // ==========================================================================
  // GENERAR
  // ==========================================================================

  function generar() {
    var t0 = new Date();
    var carpetaEntrada = DriveApp.getFolderById(CONFIG_ANUAL.ID_CARPETA_ENTRADA_ANUAL);
    var carpetaSalida = DriveApp.getFolderById(CONFIG_ANUAL.ID_CARPETA_SALIDA_ANUAL);

    var archivos = localizarArchivos(carpetaEntrada, true);

    var tablaAsistencias = cargarVarias(archivos.asistencias);
    var tablaInasistencias = cargarVarias(archivos.inasistencias);
    var tablaFallidos = cargarVarias(archivos.fallidos);
    var roster = cargarMaestro(archivos.maestro);
    var festivos = archivos.calendario ? cargarFestivos(archivos.calendario) : new Set();

    Logger.log('Asistencias: ' + tablaAsistencias.rows.length + ' filas de ' + archivos.asistencias.length + ' archivo(s).');
    Logger.log('Inasistencias: ' + tablaInasistencias.rows.length + ' filas de ' + archivos.inasistencias.length + ' archivo(s).');
    Logger.log('Registros fallidos: ' + tablaFallidos.rows.length + ' filas de ' + archivos.fallidos.length + ' archivo(s).');
    Logger.log('Maestro: ' + roster.length + ' personas. Festivos: ' + festivos.size + ' fechas.');

    // --- Año y rango de fechas ---
    var todasFechas = []
      .concat(tablaAsistencias.rows.map(function (r) { return parseFecha(r['Fecha Entrada']); }))
      .concat(tablaInasistencias.rows.map(function (r) { return parseFecha(r['Día']); }))
      .concat(tablaFallidos.rows.map(function (r) { return parseFecha(r['Fecha intento']); }))
      .filter(Boolean);
    if (todasFechas.length === 0) {
      Logger.log('No se encontraron fechas válidas en los reportes. Nada que hacer.');
      return;
    }
    todasFechas = uniqueSorted(todasFechas);

    var anio = CONFIG_ANUAL.ANIO || +todasFechas[todasFechas.length - 1].substring(0, 4);
    var prefijoAnio = anio + '-';
    var fechasDelAnio = todasFechas.filter(function (f) { return f.indexOf(prefijoAnio) === 0; });
    if (fechasDelAnio.length === 0) {
      Logger.log('No hay datos del año ' + anio + ' en los reportes. Nada que hacer.');
      return;
    }
    var fechaFin = CONFIG_ANUAL.HASTA_FIN_DE_ANIO ? anio + '-12-31' : fechasDelAnio[fechasDelAnio.length - 1];
    var fechas = rangoDeFechas(anio + '-01-01', fechaFin);
    Logger.log('Año ' + anio + ': ' + fechas[0] + ' a ' + fechaFin + ' (' + fechas.length + ' días).');

    // --- Estatus diario ---
    var res = construirDatosPorFecha(tablaAsistencias, tablaInasistencias, tablaFallidos,
      CONFIG_ANUAL.TOLERANCIA_RETARDO_MIN, prefijoAnio);

    // --- Roster final (maestro + personas fuera de maestro si aplica) ---
    var personas = construirPersonas(roster, res);
    Logger.log('Personas en el reporte: ' + personas.length);

    // --- Matriz final ---
    var matriz = construirMatriz(personas, fechas, festivos, res.datos);

    // --- Escribir en Sheet temporal y exportar a Excel ---
    var nombre = CONFIG_ANUAL.NOMBRE_ARCHIVO.replace('{ANIO}', anio);
    var ssTmp = SpreadsheetApp.create('TMP_' + nombre + '_' + new Date().getTime());
    var archivoXlsx;
    try {
      escribirReporteAnual(ssTmp, personas, fechas, matriz);
      escribirResumen(ssTmp, personas, fechas, matriz);
      escribirTabla(ssTmp, HOJAS.FALTAS, columnasFaltas(), listaFaltas(personas, fechas, matriz, res.datos), COLOR_FALTA);
      escribirTabla(ssTmp, HOJAS.RETARDOS, columnasRetardos(), listaRetardos(personas, fechas, matriz, res.datos), null);
      escribirTabla(ssTmp, HOJAS.FALLIDOS, columnasFallidos(), listaFallidos(personas, fechas, matriz, res.datos), COLOR_FALLIDO);
      escribirLeyenda(ssTmp);
      // Quita la "Hoja 1" vacía que crea SpreadsheetApp.create()
      ssTmp.getSheets().forEach(function (sh) {
        if (!esHojaNuestra(sh.getName()) && ssTmp.getSheets().length > 1) ssTmp.deleteSheet(sh);
      });
      ssTmp.setActiveSheet(ssTmp.getSheetByName(HOJAS.REPORTE));
      SpreadsheetApp.flush();

      archivoXlsx = exportarComoExcel(ssTmp.getId(), nombre, carpetaSalida);
    } finally {
      DriveApp.getFileById(ssTmp.getId()).setTrashed(true);
    }

    Logger.log('Listo. Excel generado: ' + archivoXlsx.getName() + ' -> ' + archivoXlsx.getUrl());
    Logger.log('Tiempo total: ' + ((new Date() - t0) / 1000) + 's');

    if (CONFIG_ANUAL.ENVIAR_CORREO) {
      enviarCorreo(anio, fechas, personas, matriz, archivoXlsx);
    }
  }

  function esHojaNuestra(nombre) {
    return Object.keys(HOJAS).some(function (k) { return HOJAS[k] === nombre; });
  }

  // ==========================================================================
  // DIAGNÓSTICO
  // ==========================================================================

  function diagnosticar() {
    var carpeta = DriveApp.getFolderById(CONFIG_ANUAL.ID_CARPETA_ENTRADA_ANUAL);
    Logger.log('=== Archivos en la carpeta de entrada anual ===');
    var it = carpeta.getFiles();
    while (it.hasNext()) {
      var f = it.next();
      Logger.log('  "' + f.getName() + '"  (mimeType=' + f.getMimeType() + ')');
    }

    var archivos = localizarArchivos(carpeta, false);

    function mostrar(etiqueta, file, colFecha) {
      if (!file) { Logger.log(etiqueta + ': NO ENCONTRADO'); return; }
      var t = cargarTabla(file);
      var fs = colFecha ? uniqueSorted(t.rows.map(function (r) { return parseFecha(r[colFecha]); }).filter(Boolean)) : [];
      Logger.log('=== ' + etiqueta + ': "' + file.getName() + '" -> ' + t.rows.length + ' filas' +
        (colFecha ? (', fechas ' + (fs[0] || '?') + ' a ' + (fs[fs.length - 1] || '?') + ' (' + fs.length + ' días)') : ''));
      Logger.log('  Encabezados: ' + JSON.stringify(t.headers));
      if (colFecha && t.rows.length > 0 && fs.length === 0) {
        Logger.log('  AVISO: no se reconoció ninguna fecha en "' + colFecha + '". Valor crudo fila 1: ' + t.rows[0][colFecha]);
      }
    }

    archivos.asistencias.forEach(function (f) { mostrar('Asistencias', f, 'Fecha Entrada'); });
    archivos.inasistencias.forEach(function (f) { mostrar('Inasistencias', f, 'Día'); });
    archivos.fallidos.forEach(function (f) { mostrar('Registro fallido', f, 'Fecha intento'); });
    mostrar('Maestro', archivos.maestro, null);
    Logger.log('Calendario: ' + (archivos.calendario ? archivos.calendario.getName() : 'NO ENCONTRADO (no se marcarán festivos)'));
    Logger.log('=== Fin diagnóstico ===');
  }

  // ==========================================================================
  // LOCALIZACIÓN DE ARCHIVOS
  // ==========================================================================

  function localizarArchivos(carpetaEntrada, obligatorio) {
    var respaldo = CONFIG_ANUAL.ID_CARPETA_INSUMOS_RESPALDO
      ? DriveApp.getFolderById(CONFIG_ANUAL.ID_CARPETA_INSUMOS_RESPALDO) : null;

    var asistencias = encontrarArchivos(carpetaEntrada, RE_ASISTENCIAS);
    var inasistencias = encontrarArchivos(carpetaEntrada, RE_INASISTENCIAS);
    var fallidos = encontrarArchivos(carpetaEntrada, RE_FALLIDOS);

    var maestro = masReciente(encontrarArchivos(carpetaEntrada, RE_MAESTRO));
    if (!maestro && respaldo) maestro = masReciente(encontrarArchivos(respaldo, RE_MAESTRO));
    var calendario = masReciente(encontrarArchivos(carpetaEntrada, RE_CALENDARIO));
    if (!calendario && respaldo) calendario = masReciente(encontrarArchivos(respaldo, RE_CALENDARIO));

    if (obligatorio) {
      if (asistencias.length === 0) throw new Error('No hay archivos ReporteAsistenciaEmpresa* en la carpeta de entrada anual.');
      if (inasistencias.length === 0) throw new Error('No hay archivos ReporteInasistenciaEmpresa* en la carpeta de entrada anual.');
      if (!maestro) throw new Error('No se encontró el maestro Trabajadores_* ni en la carpeta anual ni en la de respaldo.');
      if (!calendario) Logger.log('Aviso: no se encontró calendario*; no se marcarán días festivos.');
    }

    Logger.log('Archivos de asistencia: ' + nombres(asistencias));
    Logger.log('Archivos de inasistencia: ' + nombres(inasistencias));
    Logger.log('Archivos de registro fallido: ' + nombres(fallidos));
    Logger.log('Maestro: ' + (maestro ? maestro.getName() : '-') + ' | Calendario: ' + (calendario ? calendario.getName() : '-'));

    return {
      asistencias: asistencias, inasistencias: inasistencias, fallidos: fallidos,
      maestro: maestro, calendario: calendario
    };
  }

  function nombres(files) {
    return files.length ? files.map(function (f) { return f.getName(); }).join(', ') : '(ninguno)';
  }

  function encontrarArchivos(carpeta, regex) {
    var out = [];
    var it = carpeta.getFiles();
    while (it.hasNext()) {
      var f = it.next();
      if (regex.test(f.getName())) out.push(f);
    }
    out.sort(function (a, b) { return a.getName().localeCompare(b.getName()); });
    return out;
  }

  function masReciente(files) {
    if (files.length === 0) return null;
    files.sort(function (a, b) { return b.getLastUpdated() - a.getLastUpdated(); });
    return files[0];
  }

  // ==========================================================================
  // UTILIDADES DE FECHA / HORA (mismas que asistencia.js)
  // ==========================================================================

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  function parseFecha(valor) {
    if (valor === null || valor === undefined || valor === '') return null;
    if (Object.prototype.toString.call(valor) === '[object Date]') return dateKey(valor);
    if (typeof valor === 'number') return dateKey(serialAFecha(valor));
    var s = String(valor).trim();
    if (s === '' || s === '-' || s.toLowerCase() === 'nan' || s.toLowerCase() === 'nat') return null;
    var m;
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]);
    m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (m) return m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]);
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]);
    return null;
  }

  function serialAFecha(serial) {
    var d = new Date(Math.round((serial - 25569) * 86400 * 1000));
    return {
      getFullYear: function () { return d.getUTCFullYear(); },
      getMonth: function () { return d.getUTCMonth(); },
      getDate: function () { return d.getUTCDate(); }
    };
  }

  function dateKey(d) {
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  }

  function dateKeyToDate(key) {
    var p = key.split('-');
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }

  function rangoDeFechas(desde, hasta) {
    var out = [];
    var cursor = dateKeyToDate(desde);
    var fin = dateKeyToDate(hasta);
    while (cursor <= fin) {
      out.push(dateKey(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }
    return out;
  }

  function formatearFechaCorta(fechaKey) {
    var d = dateKeyToDate(fechaKey);
    return pad2(d.getDate()) + '-' + MESES_CORTOS[d.getMonth()];
  }

  function parseHora(valor) {
    if (valor === null || valor === undefined || valor === '') return null;
    if (Object.prototype.toString.call(valor) === '[object Date]') {
      return valor.getHours() * 60 + valor.getMinutes() + valor.getSeconds() / 60;
    }
    if (typeof valor === 'number') {
      var frac = valor - Math.floor(valor);
      return Math.round(frac * 1440 * 100) / 100;
    }
    var s = String(valor).trim();
    var m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return null;
    return (+m[1]) * 60 + (+m[2]) + (m[3] ? (+m[3]) / 60 : 0);
  }

  function horaMinutosAString(minutos, conSegundos) {
    var h = Math.floor(minutos / 60);
    var m = Math.floor(minutos % 60);
    var s = Math.round((minutos - Math.floor(minutos)) * 60);
    return pad2(h) + ':' + pad2(m) + (conSegundos ? ':' + pad2(s) : '');
  }

  function parseRangoHorario(valor) {
    if (valor === null || valor === undefined) return { inicio: null, fin: null };
    var s = String(valor).trim();
    if (s === '' || s === '-' || s.toLowerCase() === 'nan') return { inicio: null, fin: null };
    var partes = s.split('-');
    if (partes.length !== 2) return { inicio: null, fin: null };
    return { inicio: parseHora(partes[0]), fin: parseHora(partes[1]) };
  }

  function formatearMinutos(minutos) {
    minutos = Math.round(minutos);
    var horas = Math.floor(minutos / 60);
    var resto = minutos % 60;
    if (horas) return horas + 'h ' + pad2(resto) + 'm';
    return resto + 'm';
  }

  function limpiarRut(valor) {
    return valor === null || valor === undefined ? '' : String(valor).trim().toUpperCase();
  }

  function uniqueSorted(arr) {
    var set = {};
    arr.forEach(function (k) { set[k] = true; });
    return Object.keys(set).sort();
  }

  function nombreDeFila(r) {
    return [r['Primer Apellido'] || '', r['Segundo Apellido'] || '', r['Nombre'] || '']
      .join(' ').replace(/\s+/g, ' ').trim();
  }

  function calcularLocalidad(recinto) {
    if (recinto && String(recinto).toLowerCase().indexOf('veracruz') !== -1) return 'veracruz';
    return 'cdmx';
  }

  // ==========================================================================
  // NORMALIZACIÓN DE ENCABEZADOS (misma que asistencia.js)
  // ==========================================================================

  var HEADERS_CONOCIDOS = [
    'Recinto', 'RUT', 'Primer Apellido', 'Segundo Apellido', 'Nombre', 'Especialidad',
    'Área', 'Contrato', 'Supervisor', 'Fecha Entrada', 'Hora Entrada', 'Fecha Salida',
    'Hora Salida', 'Sigla Turno', 'Horario Turno', 'Día', 'Horario', 'Motivo',
    'ID Dispositivo', 'Error al marcar', 'Sentido', 'Fecha intento', 'Hora intento',
    'Empresa', 'Código', 'Ciudad', 'Comuna', 'Turno'
  ];

  function normalizarClave(s) {
    return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toLowerCase().replace(/\s+/g, ' ').trim();
  }

  var HEADERS_NORMALIZADOS = (function () {
    var m = {};
    HEADERS_CONOCIDOS.forEach(function (h) { m[normalizarClave(h)] = h; });
    return m;
  })();

  function encabezadoCanonico(raw) {
    var limpio = String(raw).replace(/[​-‍﻿ ]/g, ' ').trim();
    return HEADERS_NORMALIZADOS[normalizarClave(limpio)] || limpio;
  }

  // ==========================================================================
  // CARGA DE DATOS FUENTE (misma estrategia que asistencia.js)
  // ==========================================================================

  function abrirComoSheet(file) {
    try {
      return { ss: SpreadsheetApp.open(file), tempId: null };
    } catch (eDirecto) {
      Logger.log('  "' + file.getName() + '" no se pudo abrir directo (' + eDirecto.message +
        '); se intentará convertir vía Drive API...');
    }
    if (typeof Drive === 'undefined') {
      throw new Error('El archivo "' + file.getName() + '" no se pudo abrir directamente y el servicio avanzado ' +
        '"Drive API" no está habilitado. Ve al editor -> Servicios (ícono +) -> agrega "Google Drive API" ' +
        '(identificador "Drive") y vuelve a ejecutar.');
    }
    var copiado = Drive.Files.copy({
      name: 'TMP_' + file.getName() + '_' + new Date().getTime(),
      mimeType: MimeType.GOOGLE_SHEETS
    }, file.getId());
    return { ss: SpreadsheetApp.openById(copiado.id), tempId: copiado.id };
  }

  function cerrarSheetTemporal(abierto) {
    if (abierto.tempId) DriveApp.getFileById(abierto.tempId).setTrashed(true);
  }

  function cargarTabla(file) {
    var abierto = abrirComoSheet(file);
    try {
      var sheet = abierto.ss.getSheets()[0];
      var valores = sheet.getDataRange().getValues();
      if (valores.length === 0) return { headers: [], rows: [] };
      var headers = valores[0].map(encabezadoCanonico);
      var rows = [];
      for (var i = 1; i < valores.length; i++) {
        var fila = valores[i];
        if (fila.every(function (v) { return v === '' || v === null; })) continue;
        var obj = {};
        headers.forEach(function (h, j) { obj[h] = fila[j]; });
        rows.push(obj);
      }
      Logger.log('  "' + file.getName() + '": ' + rows.length + ' filas.');
      return { headers: headers, rows: rows };
    } finally {
      cerrarSheetTemporal(abierto);
    }
  }

  function cargarVarias(files) {
    var headers = [];
    var rows = [];
    files.forEach(function (f) {
      var t = cargarTabla(f);
      if (!headers.length) headers = t.headers;
      rows = rows.concat(t.rows);
    });
    return { headers: headers, rows: rows };
  }

  function cargarMaestro(file) {
    var t = cargarTabla(file);
    var vistos = {};
    var out = [];
    t.rows.forEach(function (r) {
      var rut = limpiarRut(r['RUT']);
      if (!rut || vistos[rut]) return; // se conserva la primera fila
      vistos[rut] = true;
      r.RUT = rut;
      out.push(r);
    });
    return out;
  }

  function cargarFestivos(file) {
    var abierto = abrirComoSheet(file);
    try {
      var sheet = abierto.ss.getSheetByName('festivos');
      if (!sheet) {
        Logger.log('Aviso: no se encontró la hoja "festivos" en ' + file.getName());
        return new Set();
      }
      var claves = [];
      sheet.getDataRange().getValues().forEach(function (fila) {
        var f = parseFecha(fila[1]); // columna B
        if (f) claves.push(f);
      });
      return new Set(claves);
    } finally {
      cerrarSheetTemporal(abierto);
    }
  }

  // ==========================================================================
  // CÁLCULO DEL ESTATUS DIARIO
  //
  // datos[fecha][rut] = {
  //   codigo, comentario,
  //   tipo: 'asistencia' | 'fallido' | 'inasistencia',
  //   fila: fila fuente elegida, retardoMin: minutos brutos de diferencia
  // }
  // ==========================================================================

  function describirIntento(row) {
    return ((row['Sentido'] || '') + ' ' + (row['Hora intento'] || '') + ' (' + (row['Error al marcar'] || '') + ')').trim();
  }

  function construirDatosPorFecha(tablaAsistencias, tablaInasistencias, tablaFallidos, toleranciaMin, prefijoAnio) {
    var datos = {};
    var horario = {};      // horario[rut] = { fecha, valor } (el más reciente)
    var infoFuente = {};   // infoFuente[rut] = fila más reciente (para personas fuera de maestro)

    function registrarFuente(rut, f, row) {
      if (!infoFuente[rut] || infoFuente[rut].fecha <= f) infoFuente[rut] = { fecha: f, fila: row };
    }
    function slot(f) { return (datos[f] = datos[f] || {}); }

    // --- 1) Asistencias reales, en cualquier recinto: siempre ganan ---
    tablaAsistencias.rows.forEach(function (row) {
      var f = parseFecha(row['Fecha Entrada']);
      var rut = limpiarRut(row['RUT']);
      if (!f || !rut || f.indexOf(prefijoAnio) !== 0) return;
      registrarFuente(rut, f, row);

      var horarioTurno = String(row['Horario Turno'] || '').trim();
      if (horarioTurno && horarioTurno !== '-' && (!horario[rut] || horario[rut].fecha <= f)) {
        horario[rut] = { fecha: f, valor: horarioTurno };
      }

      var rango = parseRangoHorario(row['Horario Turno']);
      var horaReal = parseHora(row['Hora Entrada']);
      var diferencia = (rango.inicio === null || horaReal === null) ? null : horaReal - rango.inicio;
      var codigo = (diferencia !== null && diferencia > toleranciaMin)
        ? 'A-' + Math.round(diferencia - toleranciaMin) : 'A';
      var retardoEfectivo = (diferencia !== null && diferencia > toleranciaMin) ? diferencia : 0;

      var previo = slot(f)[rut];
      // Si checó en varios recintos el mismo día, se queda la checada con
      // menos retardo.
      if (previo && previo.tipo === 'asistencia' && previo.retardoEfectivo <= retardoEfectivo) return;
      datos[f][rut] = {
        codigo: codigo, comentario: null, tipo: 'asistencia', fila: row,
        retardoMin: diferencia, retardoEfectivo: retardoEfectivo,
        horaProgramada: rango.inicio, horaReal: horaReal
      };
    });

    // --- 2) Intentos de checada fallidos: cuentan como asistencia si no hay
    //        asistencia real ese día ---
    var fallidos = {}; // fallidos[f][rut] = { detalles: {texto:true}, fila }
    tablaFallidos.rows.forEach(function (row) {
      var f = parseFecha(row['Fecha intento']);
      var rut = limpiarRut(row['RUT']);
      if (!f || !rut || f.indexOf(prefijoAnio) !== 0) return;
      registrarFuente(rut, f, row);
      fallidos[f] = fallidos[f] || {};
      fallidos[f][rut] = fallidos[f][rut] || { detalles: {}, fila: row };
      fallidos[f][rut].detalles[describirIntento(row)] = true; // sin duplicados si los archivos se traslapan
    });
    Object.keys(fallidos).forEach(function (f) {
      Object.keys(fallidos[f]).forEach(function (rut) {
        if (slot(f)[rut]) return; // ya tiene asistencia real
        var detalle = Object.keys(fallidos[f][rut].detalles).join('; ');
        datos[f][rut] = {
          codigo: 'A', tipo: 'fallido', fila: fallidos[f][rut].fila, detalle: detalle,
          comentario: 'Asistencia inferida por intento de checada fallido: ' + detalle
        };
      });
    });

    // --- 3) Inasistencias: solo si no hubo asistencia real ni intento.
    //        Si hay varias filas ese día (varios recintos), cualquier
    //        justificación le gana a la falta sin justificar. ---
    tablaInasistencias.rows.forEach(function (row) {
      var f = parseFecha(row['Día']);
      var rut = limpiarRut(row['RUT']);
      if (!f || !rut || f.indexOf(prefijoAnio) !== 0) return;
      registrarFuente(rut, f, row);

      var previo = slot(f)[rut];
      if (previo && previo.tipo !== 'inasistencia') return; // asistió o tuvo intento
      if (previo && previo.codigo !== 'F') return;          // ya hay una justificación

      var motivo = String(row['Motivo'] || '').trim();
      var codigo = MOTIVO_A_CODIGO.hasOwnProperty(motivo) ? MOTIVO_A_CODIGO[motivo] : null;
      if (codigo === null) {
        Logger.log('Aviso: motivo "' + motivo + '" no está en el catálogo (RUT ' + rut + ', ' + f + '); se deja tal cual.');
        codigo = motivo || 'F';
      }
      if (previo && codigo === 'F') return; // no se sobreescribe con otra F
      datos[f][rut] = { codigo: codigo, comentario: null, tipo: 'inasistencia', fila: row };
    });

    return { datos: datos, horario: horario, infoFuente: infoFuente };
  }

  // ==========================================================================
  // PERSONAS Y MATRIZ
  // ==========================================================================

  function construirPersonas(roster, res) {
    var enMaestro = {};
    var personas = roster.map(function (p) {
      enMaestro[p.RUT] = true;
      return {
        RUT: p.RUT,
        'Nombre completo': nombreDeFila(p),
        'Área': p['Área'] || '',
        'Horario Turno': res.horario[p.RUT] ? res.horario[p.RUT].valor : '-',
        'Recinto': p['Recinto'] || '',
        'Localidad': calcularLocalidad(p['Recinto']),
        'Supervisor': p['Supervisor'] || '',
        'En maestro': 'Sí'
      };
    });

    if (CONFIG_ANUAL.INCLUIR_FUERA_DE_MAESTRO) {
      var extra = Object.keys(res.infoFuente).filter(function (rut) { return !enMaestro[rut]; });
      extra.sort();
      extra.forEach(function (rut) {
        var r = res.infoFuente[rut].fila;
        personas.push({
          RUT: rut,
          'Nombre completo': nombreDeFila(r),
          'Área': r['Área'] || '',
          'Horario Turno': res.horario[rut] ? res.horario[rut].valor : '-',
          'Recinto': r['Recinto'] || '',
          'Localidad': calcularLocalidad(r['Recinto']),
          'Supervisor': r['Supervisor'] || '',
          'En maestro': 'No'
        });
      });
      if (extra.length) Logger.log('Personas fuera del maestro incluidas: ' + extra.length);
    }
    return personas;
  }

  // matriz[i][k] = código de la persona i en la fecha k
  function construirMatriz(personas, fechas, festivos, datos) {
    return personas.map(function (p) {
      p._notas = {};     // notas de celda (asistencia inferida por intento fallido)
      p._inferidos = 0;  // días contados como asistencia por intento fallido
      return fechas.map(function (f) {
        if (festivos.has(f)) return 'DF';
        var d = datos[f] && datos[f][p.RUT];
        if (!d) return '-';
        if (d.comentario) p._notas[f] = d.comentario;
        if (d.tipo === 'fallido') p._inferidos++;
        return d.codigo;
      });
    });
  }

  function esRetardo(codigo) { return typeof codigo === 'string' && codigo.indexOf('A-') === 0; }
  function esAsistencia(codigo) { return codigo === 'A' || esRetardo(codigo); }

  // ==========================================================================
  // HOJA "Reporte anual"
  // ==========================================================================

  function escribirReporteAnual(ss, personas, fechas, matriz) {
    var sheet = ss.insertSheet(HOJAS.REPORTE);
    var nMeta = COLS_META.length;
    var primeraColFecha = nMeta + 1;
    var colResumen = primeraColFecha + fechas.length;
    var colsResumen = ['Asistencias', 'Retardos', 'Faltas', 'Fechas de falta'];
    var totalCols = colResumen + colsResumen.length - 1;
    var nFilas = personas.length;

    asegurarTamano(sheet, PRIMERA_FILA_DATOS + nFilas, totalCols);

    // --- Fila de meses (agrupa las columnas de fecha por mes) ---
    var inicioMes = 0;
    for (var k = 1; k <= fechas.length; k++) {
      if (k === fechas.length || fechas[k].substring(0, 7) !== fechas[inicioMes].substring(0, 7)) {
        var mes = +fechas[inicioMes].substring(5, 7) - 1;
        var rMes = sheet.getRange(FILA_MESES, primeraColFecha + inicioMes, 1, k - inicioMes);
        if (k - inicioMes > 1) rMes.merge();
        rMes.getCell(1, 1).setValue(MESES_LARGOS[mes] + ' ' + fechas[inicioMes].substring(0, 4));
        rMes.setBackground(COLOR_MES_FONDO).setFontWeight('bold').setHorizontalAlignment('center');
        inicioMes = k;
      }
    }

    // --- Encabezados ---
    var encabezados = COLS_META
      .concat(fechas.map(dateKeyToDate))
      .concat(colsResumen);
    var rHeader = sheet.getRange(FILA_ENCABEZADO, 1, 1, totalCols);
    rHeader.setValues([encabezados]);
    estiloEncabezado(rHeader);
    sheet.getRange(FILA_ENCABEZADO, primeraColFecha, 1, fechas.length).setNumberFormat('dd-mmm');

    if (nFilas > 0) {
      var valores = [];
      var fondos = [];
      var notas = [];
      personas.forEach(function (p, i) {
        var fila = COLS_META.map(function (c) { return p[c]; });
        var fondoFila = COLS_META.map(function () { return null; });
        var notaFila = COLS_META.map(function () { return ''; });
        var asistencias = 0, retardos = 0, fechasFalta = [];

        fechas.forEach(function (f, k) {
          var codigo = matriz[i][k];
          var nota = p._notas[f] || '';
          fila.push(codigo);
          fondoFila.push(colorDeCodigo(codigo, !!nota));
          notaFila.push(nota);
          if (esAsistencia(codigo)) asistencias++;
          if (esRetardo(codigo)) retardos++;
          if (codigo === 'F') fechasFalta.push(formatearFechaCorta(f));
        });

        fila.push(asistencias, retardos, fechasFalta.length, fechasFalta.join(', '));
        fondoFila.push(null, retardos ? COLOR_RETARDO : null, fechasFalta.length ? COLOR_FALTA : null, fechasFalta.length ? COLOR_FALTA : null);
        notaFila.push('', '', '', '');
        valores.push(fila);
        fondos.push(fondoFila);
        notas.push(notaFila);
      });

      var rDatos = sheet.getRange(PRIMERA_FILA_DATOS, 1, nFilas, totalCols);
      rDatos.setValues(valores);
      rDatos.setBackgrounds(fondos);
      rDatos.setNotes(notas);
      rDatos.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
      sheet.getRange(PRIMERA_FILA_DATOS, primeraColFecha, nFilas, fechas.length + 3).setHorizontalAlignment('center');
      sheet.getRange(PRIMERA_FILA_DATOS, colResumen + 3, nFilas, 1).setFontSize(9).setWrap(true);
    }

    sheet.setFrozenRows(FILA_ENCABEZADO);
    sheet.setFrozenColumns(2); // RUT + Nombre siempre visibles al desplazarse
    sheet.setColumnWidth(1, 95);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(4, 110);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 80);
    sheet.setColumnWidth(7, 190);
    sheet.setColumnWidth(8, 80);
    if (fechas.length) sheet.setColumnWidths(primeraColFecha, fechas.length, 52);
    sheet.setColumnWidths(colResumen, 3, 80);
    sheet.setColumnWidth(colResumen + 3, 320);
  }

  function colorDeCodigo(codigo, esInferido) {
    if (codigo === 'DF') return COLOR_DF;
    if (codigo === 'F') return COLOR_FALTA;
    if (esRetardo(codigo)) return COLOR_RETARDO;
    if (codigo === 'A') return esInferido ? COLOR_FALLIDO : null;
    if (codigo === '-') return null;
    return COLOR_JUSTIFICADA; // V, I, P, L, FJ u otro motivo
  }

  // ==========================================================================
  // HOJA "Resumen"
  // ==========================================================================

  function escribirResumen(ss, personas, fechas, matriz) {
    var sheet = ss.insertSheet(HOJAS.RESUMEN);
    var conteos = ['Asistencias', 'Asist. por intento fallido', 'Retardos', 'Faltas', 'V', 'I', 'P', 'L', 'FJ', 'Otros motivos', 'Sin datos'];
    var columnas = ['RUT', 'Nombre completo', 'Área', 'Recinto', 'Supervisor', 'En maestro']
      .concat(conteos)
      .concat(MESES_CORTOS.map(function (m) { return 'Faltas ' + m; }));

    var filas = personas.map(function (p, i) {
      var c = {};
      conteos.forEach(function (k) { c[k] = 0; });
      var faltasMes = MESES_CORTOS.map(function () { return 0; });
      matriz[i].forEach(function (codigo, k) {
        if (esAsistencia(codigo)) c['Asistencias']++;
        if (esRetardo(codigo)) c['Retardos']++;
        if (codigo === 'F') {
          c['Faltas']++;
          faltasMes[+fechas[k].substring(5, 7) - 1]++;
        } else if (codigo === '-') c['Sin datos']++;
        else if (['V', 'I', 'P', 'L', 'FJ'].indexOf(codigo) !== -1) c[codigo]++;
        else if (!esAsistencia(codigo) && codigo !== 'DF') c['Otros motivos']++;
      });
      c['Asist. por intento fallido'] = p._inferidos || 0;
      return [p.RUT, p['Nombre completo'], p['Área'], p['Recinto'], p['Supervisor'], p['En maestro']]
        .concat(conteos.map(function (k) { return c[k]; }))
        .concat(faltasMes);
    });

    asegurarTamano(sheet, filas.length + 1, columnas.length);
    var rHeader = sheet.getRange(1, 1, 1, columnas.length);
    rHeader.setValues([columnas]);
    estiloEncabezado(rHeader);
    if (filas.length) {
      var rDatos = sheet.getRange(2, 1, filas.length, columnas.length);
      rDatos.setValues(filas).setFontFamily('Arial').setFontSize(10);
      sheet.getRange(2, 7, filas.length, columnas.length - 6).setHorizontalAlignment('center');
      var colFaltas = columnas.indexOf('Faltas') + 1;
      var fondos = filas.map(function (f) { return [f[colFaltas - 1] > 0 ? COLOR_FALTA : null]; });
      sheet.getRange(2, colFaltas, filas.length, 1).setBackgrounds(fondos);
    }
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
    sheet.setColumnWidth(1, 95);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 190);
    sheet.setColumnWidths(6, columnas.length - 5, 75);
  }

  // ==========================================================================
  // HOJAS DE DETALLE: "Faltas", "Retardos", "Intentos fallidos"
  // ==========================================================================

  function columnasFaltas() {
    return ['Fecha', 'RUT', 'Nombre completo', 'Área', 'Recinto', 'Supervisor', 'Contrato', 'Horario', 'Sigla Turno'];
  }

  function listaFaltas(personas, fechas, matriz, datos) {
    var out = [];
    fechas.forEach(function (f, k) {
      personas.forEach(function (p, i) {
        if (matriz[i][k] !== 'F') return;
        var r = (datos[f] && datos[f][p.RUT] && datos[f][p.RUT].fila) || {};
        out.push([dateKeyToDate(f), p.RUT, p['Nombre completo'], r['Área'] || p['Área'], r['Recinto'] || p['Recinto'],
          r['Supervisor'] || p['Supervisor'], r['Contrato'] || '', r['Horario'] || '', r['Sigla Turno'] || '']);
      });
    });
    return out;
  }

  function columnasRetardos() {
    return ['Fecha', 'RUT', 'Nombre completo', 'Área', 'Recinto', 'Supervisor', 'Sigla Turno',
      'Hora Entrada Programada', 'Hora Entrada Real', 'Minutos Tarde', 'Tiempo Tarde', 'Código'];
  }

  function listaRetardos(personas, fechas, matriz, datos) {
    var out = [];
    fechas.forEach(function (f, k) {
      personas.forEach(function (p, i) {
        if (!esRetardo(matriz[i][k])) return;
        var d = datos[f][p.RUT];
        var r = d.fila;
        out.push([dateKeyToDate(f), p.RUT, p['Nombre completo'], r['Área'] || p['Área'], r['Recinto'] || '',
          r['Supervisor'] || p['Supervisor'], r['Sigla Turno'] || '',
          horaMinutosAString(d.horaProgramada, false), horaMinutosAString(d.horaReal, true),
          Math.round(d.retardoMin), formatearMinutos(d.retardoMin), matriz[i][k]]);
      });
    });
    return out;
  }

  function columnasFallidos() {
    return ['Fecha', 'RUT', 'Nombre completo', 'Área', 'Recinto', 'Supervisor', 'Detalle del intento'];
  }

  function listaFallidos(personas, fechas, matriz, datos) {
    var out = [];
    fechas.forEach(function (f, k) {
      personas.forEach(function (p, i) {
        var d = datos[f] && datos[f][p.RUT];
        if (!d || d.tipo !== 'fallido' || matriz[i][k] === 'DF') return;
        var r = d.fila;
        out.push([dateKeyToDate(f), p.RUT, p['Nombre completo'], r['Área'] || p['Área'], r['Recinto'] || '',
          r['Supervisor'] || p['Supervisor'], d.detalle]);
      });
    });
    return out;
  }

  function escribirTabla(ss, nombreHoja, columnas, filas, colorFilas) {
    var sheet = ss.insertSheet(nombreHoja);
    asegurarTamano(sheet, filas.length + 3, columnas.length);
    var rHeader = sheet.getRange(1, 1, 1, columnas.length);
    rHeader.setValues([columnas]);
    estiloEncabezado(rHeader);
    sheet.setFrozenRows(1);
    if (filas.length) {
      var rDatos = sheet.getRange(2, 1, filas.length, columnas.length);
      rDatos.setValues(filas).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
      if (colorFilas) rDatos.setBackground(colorFilas);
      sheet.getRange(2, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    }
    sheet.getRange(filas.length + 3, 1).setValue('Total de registros: ' + filas.length)
      .setFontStyle('italic').setFontSize(9).setFontColor('#808080');
    sheet.autoResizeColumns(1, columnas.length);
  }

  function escribirLeyenda(ss) {
    var sheet = ss.insertSheet(HOJAS.LEYENDA);
    var rHeader = sheet.getRange(1, 1, 1, 2);
    rHeader.setValues([['Código', 'Significado']]);
    estiloEncabezado(rHeader);
    sheet.getRange(2, 1, LEYENDA.length, 2).setValues(LEYENDA).setFontFamily('Arial').setFontSize(10);
    sheet.getRange(2, 1, LEYENDA.length, 1).setBackgrounds(LEYENDA.map(function (par) {
      return [par[0] === '(celda azul)' ? COLOR_FALLIDO : colorDeCodigo(par[0] === 'A-N' ? 'A-1' : par[0], false)];
    })).setFontWeight('bold');
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 560);
  }

  function estiloEncabezado(range) {
    range.setFontFamily('Arial').setFontWeight('bold').setFontColor(COLOR_HEADER_TEXTO)
      .setBackground(COLOR_HEADER_FONDO).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true);
  }

  function asegurarTamano(sheet, filas, columnas) {
    if (sheet.getMaxRows() < filas) sheet.insertRowsAfter(sheet.getMaxRows(), filas - sheet.getMaxRows());
    if (sheet.getMaxColumns() < columnas) sheet.insertColumnsAfter(sheet.getMaxColumns(), columnas - sheet.getMaxColumns());
  }

  // ==========================================================================
  // EXPORTAR A EXCEL
  // ==========================================================================

  function exportarComoExcel(spreadsheetId, nombre, carpetaSalida) {
    var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=xlsx';
    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error('No se pudo exportar a Excel (HTTP ' + resp.getResponseCode() + '): ' +
        resp.getContentText().substring(0, 300));
    }
    var nombreArchivo = nombre + '.xlsx';

    // Reemplaza la versión anterior (queda en la papelera por si se necesita)
    var previos = carpetaSalida.getFilesByName(nombreArchivo);
    while (previos.hasNext()) previos.next().setTrashed(true);

    return carpetaSalida.createFile(resp.getBlob().setName(nombreArchivo));
  }

  // ==========================================================================
  // CORREO
  // ==========================================================================

  function enviarCorreo(anio, fechas, personas, matriz, archivo) {
    if (!CONFIG_ANUAL.CORREOS.length) return;
    var faltas = 0, retardos = 0;
    matriz.forEach(function (fila) {
      fila.forEach(function (c) {
        if (c === 'F') faltas++;
        if (esRetardo(c)) retardos++;
      });
    });
    var html = ''
      + '<div style="font-family:Arial, sans-serif; font-size:14px; color:#222;">'
      + '<p>Estimados,</p>'
      + '<p>Se generó el <strong>reporte anual de asistencia ' + anio + '</strong> '
      + '(del ' + fechas[0] + ' al ' + fechas[fechas.length - 1] + ', ' + personas.length + ' personas).</p>'
      + '<p>Faltas sin justificar: <strong>' + faltas + '</strong> · Retardos: <strong>' + retardos + '</strong></p>'
      + '<p>Archivo: <a href="' + archivo.getUrl() + '">' + archivo.getName() + '</a></p>'
      + '<p>Saludos cordiales.</p>'
      + '</div>';
    MailApp.sendEmail({
      to: CONFIG_ANUAL.CORREOS.join(','),
      subject: CONFIG_ANUAL.ASUNTO_CORREO + ' ' + anio,
      htmlBody: html
    });
    Logger.log('Correo enviado a ' + CONFIG_ANUAL.CORREOS.join(', '));
  }

  return { generar: generar, diagnosticar: diagnosticar };
})();
