/*******************************************************************************
 * REPORTE DE ASISTENCIA / FALTAS / RETARDOS - Google Apps Script
 * Puerto de generar_reporte_faltas.py (faltas.py) a Google Sheets + Drive.
 *
 * QUÉ HACE:
 *  - Busca en la carpeta de ENTRADA (Informes_de_asistencia) los archivos
 *    del día/quincena:
 *      ReporteAsistenciaEmpresa*.xls*
 *      ReporteInasistenciaEmpresa*.xls*
 *      ReporteRegistroFallido*.xls*   (puede haber varios)
 *      Trabajadores_*.xls*            (lista maestra)
 *      calendario*.xlsx               (hoja "festivos")
 *  - Si el archivo ya es un Google Sheet nativo (lo normal al subirlo a
 *    Drive), lo lee directamente. Si llegara a ser un binario de Office
 *    real, lo convierte a un Sheet temporal, lo lee y borra la copia.
 *  - Escribe/actualiza 3 hojas dentro del Sheet YA EXISTENTE (no crea uno
 *    nuevo): "Faltas", "Retardos" y "Reporte de asistencia".
 *  - "Reporte de asistencia" es acumulativa: conserva el historial de fechas
 *    ya escritas y agrega/actualiza las de la corrida actual.
 *  - Al final de las columnas de fecha se agregan dos columnas resumen:
 *      "Retardos": conteo total de días con código A-N (retardo) para esa
 *                  persona, considerando todo el histórico acumulado hasta
 *                  el momento.
 *      "Faltas":   lista (separada por coma) de las fechas en las que esa
 *                  persona tiene código F (falta sin justificar).
 *  - Cuando toca generar el Sheet en blanco de la siguiente quincena, lo
 *    crea en la carpeta de SALIDA (Asistencias_quincena), NO en la carpeta
 *    de entrada.
 *  - Al final, opcionalmente manda un correo de resumen con el link al Sheet.
 *
 * REGLAS DE PRIORIDAD PARA EL ESTATUS DIARIO DE CADA PERSONA (RUT + fecha):
 *   1. Si hay un registro REAL en ReporteAsistenciaEmpresa ese día (en
 *      cualquier recinto) -> gana la asistencia (A / A-N). Se ignora
 *      cualquier inasistencia de ese mismo día para esa persona, sin
 *      importar el Motivo que traiga.
 *   2. Si NO hay asistencia real pero SÍ hay un intento en
 *      ReporteRegistroFallido ese mismo día -> se cuenta como asistencia
 *      (A), con una nota en la celda indicando que viene de un intento de
 *      checada fallido (trazabilidad, no bloquea el conteo).
 *   3. Si no aplica ninguna de las dos anteriores, se usa el Motivo de la
 *      inasistencia (F / V / L / P / I / FJ).
 *   Por esto, "Faltas" excluye automáticamente a quien caiga en el caso 1 o
 *   2 (no se queda como "revisar", se resuelve solo).
 *
 * CONFIGURACIÓN REQUERIDA (ver bloque CONFIG más abajo):
 *  1. ID_CARPETA_INSUMOS: carpeta de ENTRADA ("Informes_de_asistencia") de
 *     donde se leen los archivos de cada corrida + el maestro de
 *     trabajadores + el calendario.
 *  2. ID_CARPETA_REPORTES: carpeta de SALIDA ("Asistencias_quincena") donde
 *     se generan los Sheets en blanco de cada nueva quincena.
 *  3. (Solo si algún archivo llega a subirse como .xls/.xlsx "de verdad",
 *     no como Google Sheet nativo) Habilitar el servicio avanzado
 *     "Drive API": Editor -> Servicios (ícono +) -> "Google Drive API" ->
 *     Agregar (deja el identificador "Drive"). No es necesario si todos tus
 *     archivos quedan como Google Sheets al subirlos, que es el caso actual.
 *  4. Revisar/ajustar CORREOS_ASISTENCIA y ENVIAR_CORREO.
 *
 * CÓMO SE EJECUTA:
 *  - Diagnóstico (recomendado primero): corre la función `diagnosticar` desde
 *    el editor. No escribe nada en el Sheet, solo muestra en el log qué
 *    archivos encontró y qué está leyendo de cada uno (encabezados, primeras
 *    filas, cómo interpretó la fecha). Úsala si algo sale raro.
 *  - Manual: abre el proyecto, selecciona la función `main` y da "Ejecutar"
 *    (la primera vez pedirá autorizar permisos).
 *  - Con menú: al abrir el Sheet aparece un menú "📋 Reporte Asistencia" con
 *    la opción "Actualizar reporte ahora".
 *  - Programado: puedes correr `crearTriggerDiario()` una sola vez desde el
 *    editor para que se ejecute solo cada día (opcional).
 ******************************************************************************/

// ============================================================================
// CONFIG - AJUSTA ESTOS VALORES
// ============================================================================

var CONFIG = {
  ID_SPREADSHEET: '1tpmymeKrgM7uaRroPtYKOuRpV3Q_m98un7P-i1O8xHE',
  // Carpeta de ENTRADA: "Informes_de_asistencia" -> de aquí se LEEN los
  // reportes de cada corrida + el maestro de trabajadores + el calendario.
  ID_CARPETA_INSUMOS: '1J8HKFr8BxSiKZ7xnH2neuYIspeaU1jQ8',
  // Carpeta de SALIDA: "Asistencias_quincena" -> aquí se GENERA el Sheet en
  // blanco de cada nueva quincena (nunca se lee de aquí).
  ID_CARPETA_REPORTES: '1L1hp_PVfmFRTmWRB_d2vyXSySaCsckA7',
  TOLERANCIA_RETARDO_MIN: 10,
  CORREOS_ASISTENCIA: [
    'asalazar@abcsc.mx',
    'kfernandez@abcsc.mx',
    'ogutierrez@abcsc.mx',
    //'atrujillo@abcsc.mx',
    //'imedrano@abcsc.mx',
    'kluna@abcsc.mx',
    'ccarbajal@abcsc.mx',
    'jperez@abcsc.mx',
  ],
  ENVIAR_CORREO: true,
  ASUNTO_CORREO: 'Asistencia'
};

var NOMBRES_HOJA = {
  FALTAS: 'Faltas',
  RETARDOS: 'Retardos',
  REPORTE: 'Reporte de asistencia'
};

var COLOR_HEADER_FONDO = '#1F4E78';
var COLOR_HEADER_TEXTO = '#FFFFFF';
var COLOR_REVISAR = '#FFF2CC';
var COLOR_FALTA = '#FCE4E4';
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
  ['F', 'FALTA (SIN JUSTIFICAR)'],
  ['V', 'VACACIONES'],
  ['I', 'INCAPACIDAD'],
  ['P', 'PERMISO'],
  ['L', 'LICENCIA (MATERNIDAD / PATERNIDAD / IMSS)'],
  ['FJ', 'FALTA JUSTIFICADA'],
  ['DF', 'DÍA FESTIVO'],
  ['-', 'SIN INCIDENCIA (SIN DATOS ESE DÍA)']
];

var COLS_META = ['Nombre completo', 'Área', 'Horario Turno', 'Recinto', 'Localidad', 'Supervisor'];
var PRIMERA_COL_FECHA = 8; // columna H
var FILA_ENCABEZADO = 2;
var PRIMERA_FILA_DATOS = 3;
var COL_RUT_OCULTA = 1;

// ============================================================================
// MENÚ / ENTRADA
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Reporte Asistencia')
    .addItem('Actualizar reporte ahora', 'main')
    .addToUi();
}

function crearTriggerDiario() {
  // Ejecuta `main` todos los días entre 7 y 8am. Córrelo UNA vez manualmente.
  ScriptApp.newTrigger('main').timeBased().everyDays(1).atHour(7).create();
}

// ============================================================================
// MAIN
// ============================================================================

function main() {
  var t0 = new Date();
  Logger.log('Buscando archivos de entrada...');
  var carpetaInsumos = DriveApp.getFolderById(CONFIG.ID_CARPETA_INSUMOS);
  var carpetaReportes = DriveApp.getFolderById(CONFIG.ID_CARPETA_REPORTES);

  var archivoAsistencias = encontrarArchivo(carpetaInsumos, /ReporteAsistenciaEmpresa.*\.xls/i, true);
  var archivoInasistencias = encontrarArchivo(carpetaInsumos, /ReporteInasistenciaEmpresa.*\.xls/i, true);
  var archivosFallidos = encontrarArchivos(carpetaInsumos, /ReporteRegistroFallido.*\.xls/i);
  var archivoMaestro = encontrarArchivo(carpetaInsumos, /Trabajadores_.*\.xls/i, true);
  var archivoCalendario = encontrarArchivo(carpetaInsumos, /calendario.*\.xlsx?/i, true);

  Logger.log('  Asistencias: ' + archivoAsistencias.getName());
  Logger.log('  Inasistencias: ' + archivoInasistencias.getName());
  Logger.log('  Registro(s) fallido(s): ' + archivosFallidos.map(function (f) { return f.getName(); }).join(', '));
  Logger.log('  Maestro: ' + archivoMaestro.getName());
  Logger.log('  Calendario: ' + archivoCalendario.getName());

  var tablaAsistencias = cargarTabla(archivoAsistencias);
  var tablaInasistencias = cargarTabla(archivoInasistencias);
  var tablaFallidos = cargarRegistrosFallidos(archivosFallidos);
  var roster = cargarMaestro(archivoMaestro);
  var festivos = cargarFestivos(archivoCalendario);
  var quincenas = cargarQuincenas(archivoCalendario);

  // --- Diagnóstico: confirma que sí se están leyendo filas y encabezados ---
  Logger.log('[Diag] Asistencias -> ' + tablaAsistencias.rows.length + ' filas. Encabezados: ' + JSON.stringify(tablaAsistencias.headers));
  if (tablaAsistencias.rows[0]) Logger.log('[Diag] Asistencias fila 1: ' + JSON.stringify(tablaAsistencias.rows[0]));
  Logger.log('[Diag] Inasistencias -> ' + tablaInasistencias.rows.length + ' filas. Encabezados: ' + JSON.stringify(tablaInasistencias.headers));
  if (tablaInasistencias.rows[0]) Logger.log('[Diag] Inasistencias fila 1: ' + JSON.stringify(tablaInasistencias.rows[0]));
  Logger.log('[Diag] Fallidos -> ' + tablaFallidos.rows.length + ' filas.');
  Logger.log('[Diag] Roster -> ' + roster.length + ' personas.');
  Logger.log('[Diag] Festivos -> ' + festivos.size + ' fechas.');
  Logger.log('[Diag] Quincenas -> ' + quincenas.length + ' periodos leídos del calendario.');

  // --- Generar (si corresponde) el Sheet en blanco de la siguiente quincena.
  // Se crea en la carpeta de SALIDA (Asistencias_quincena), no en la de
  // entrada. Se revisa cada vez que corre main(): si la fecha de hoy ya
  // alcanzó o pasó la "Semana 2 Fin" de algún periodo y todavía no existe un
  // archivo con el nombre de "quincena_nom" en la carpeta de reportes, se
  // crea uno nuevo en blanco (sin copiar hojas ni datos) ahí mismo.
  generarArchivosQuincenaSiCorresponde(carpetaReportes, quincenas);

  // --- Fecha objetivo: la más reciente de la corrida ---
  var fechasAsistencia = tablaAsistencias.rows.map(function (r) { return parseFecha(r['Fecha Entrada']); }).filter(Boolean);
  var fechasInasistencia = tablaInasistencias.rows.map(function (r) { return parseFecha(r['Día']); }).filter(Boolean);
  var todasFechasNuevas = uniqueSorted(fechasAsistencia.concat(fechasInasistencia));

  if (todasFechasNuevas.length === 0) {
    Logger.log('No se encontraron fechas válidas en los archivos de entrada. Nada que hacer.');
    return;
  }

  var fechaMasReciente = todasFechasNuevas[todasFechasNuevas.length - 1];
  var fechaMinCorrida = todasFechasNuevas[0];
  Logger.log('Rango de fechas en esta corrida: ' + fechaMinCorrida + ' a ' + fechaMasReciente);

  Logger.log('Construyendo hoja de Faltas...');
  var faltas = construirFaltas(tablaAsistencias, tablaInasistencias, tablaFallidos, fechaMasReciente);

  Logger.log('Construyendo hoja de Retardos...');
  var retardos = construirRetardos(tablaAsistencias, fechaMasReciente, CONFIG.TOLERANCIA_RETARDO_MIN);

  Logger.log('Calculando estatus diario por persona...');
  var res = construirDatosPorFecha(tablaAsistencias, tablaInasistencias, tablaFallidos, CONFIG.TOLERANCIA_RETARDO_MIN);
  var datosNuevos = res.datos;
  var horarioReciente = res.horarioMasReciente;

  var ss = SpreadsheetApp.openById(CONFIG.ID_SPREADSHEET);

  // --- Leer historial existente de "Reporte de asistencia" (antes de sobrescribir) ---
  var historialInfo = leerHistorialExistente(ss);
  var historial = historialInfo.historial;
  var metaExistente = historialInfo.meta;
  var fechasExistentes = historialInfo.fechas;

  // --- Rango completo de fechas de esta corrida (incluye festivos y días sin datos) ---
  var rangoCorrida = [];
  var cursor = dateKeyToDate(fechaMinCorrida);
  var fin = dateKeyToDate(fechaMasReciente);
  while (cursor <= fin) {
    rangoCorrida.push(dateKey(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  var rangoCorridaSet = new Set(rangoCorrida);

  var fechasFinales = uniqueSortedKeys(fechasExistentes.concat(rangoCorrida));

  // --- Construir datos_finales[fecha][RUT] = codigo, comentarios[fecha][RUT] ---
  var rutsRoster = roster.map(function (p) { return p.RUT; });
  var datosFinales = {};
  var comentarios = {};

  fechasFinales.forEach(function (f) {
    datosFinales[f] = {};
    comentarios[f] = {};
    var esFestivo = festivos.has(f);
    var esDeEstaCorrida = rangoCorridaSet.has(f);

    rutsRoster.forEach(function (rut) {
      if (esFestivo) {
        datosFinales[f][rut] = 'DF';
        return;
      }
      if (esDeEstaCorrida && datosNuevos[f] && datosNuevos[f][rut]) {
        var par = datosNuevos[f][rut];
        datosFinales[f][rut] = par.codigo;
        if (par.comentario) comentarios[f][rut] = par.comentario;
      } else if (historial[f] && historial[f][rut]) {
        datosFinales[f][rut] = historial[f][rut];
      } else {
        datosFinales[f][rut] = '-';
      }
    });
  });

  // --- Refrescar metadatos del roster ---
  roster.forEach(function (p) {
    p['Localidad'] = calcularLocalidad(p['Recinto']);
    p['Nombre completo'] = [p['Primer Apellido'] || '', p['Segundo Apellido'] || '', p['Nombre'] || '']
      .join(' ').replace(/\s+/g, ' ').trim();
    if (horarioReciente[p.RUT]) {
      p['Horario Turno'] = horarioReciente[p.RUT];
    } else if (metaExistente[p.RUT] && metaExistente[p.RUT]['Horario Turno']) {
      p['Horario Turno'] = metaExistente[p.RUT]['Horario Turno'];
    } else {
      p['Horario Turno'] = '-';
    }
  });

  // --- Escribir hojas en el Sheet existente ---
  escribirHojaTabular(ss, NOMBRES_HOJA.FALTAS, faltas.columnas, faltas.filas, true);
  escribirHojaTabular(ss, NOMBRES_HOJA.RETARDOS, retardos.columnas, retardos.filas, false);
  escribirReporteAsistencia(ss, roster, fechasFinales, datosFinales, comentarios);

  var faltasRevision = faltas.filas.filter(function (r) { return r['Estatus'] === 'Revisar - posible asistencia'; }).length;

  Logger.log('Listo. Reporte actualizado.');
  Logger.log('  Faltas (' + fechaMasReciente + '): ' + faltas.filas.length + ' (' + faltasRevision + ' para revisión)');
  Logger.log('  Retardos (' + fechaMasReciente + '): ' + retardos.filas.length);
  Logger.log('  Reporte de asistencia: ' + fechasFinales.length + ' fechas x ' + roster.length + ' personas');
  Logger.log('Tiempo total: ' + ((new Date() - t0) / 1000) + 's');

  if (CONFIG.ENVIAR_CORREO) {
    enviarCorreoResumen({
      fecha: fechaMasReciente,
      faltas: faltas.filas.length,
      faltasRevision: faltasRevision,
      retardos: retardos.filas.length,
      fechas: fechasFinales.length,
      personas: roster.length,
      url: ss.getUrl()
    });
  }
}

// ============================================================================
// DIAGNÓSTICO (no escribe nada, solo lee y muestra en el log qué está
// encontrando en cada archivo — córrela primero si algo sale raro en `main`)
// ============================================================================

function diagnosticar() {
  var carpeta = DriveApp.getFolderById(CONFIG.ID_CARPETA_INSUMOS);

  Logger.log('=== Archivos en la carpeta de entrada (Informes_de_asistencia) ===');
  var it = carpeta.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    Logger.log('  "' + f.getName() + '"  (mimeType=' + f.getMimeType() + ', modificado=' + f.getLastUpdated() + ')');
  }

  function mostrar(etiqueta, file) {
    if (!file) { Logger.log(etiqueta + ': NO ENCONTRADO'); return; }
    Logger.log('=== ' + etiqueta + ': "' + file.getName() + '" ===');
    var t = cargarTabla(file);
    Logger.log('  Encabezados leídos: ' + JSON.stringify(t.headers));
    for (var i = 0; i < Math.min(3, t.rows.length); i++) {
      Logger.log('  Fila ' + (i + 1) + ': ' + JSON.stringify(t.rows[i], function (k, v) {
        return Object.prototype.toString.call(v) === '[object Date]' ? ('DATE:' + v.toISOString()) : v;
      }));
    }
    if (t.rows.length > 0) {
      var muestraFecha = t.headers.indexOf('Fecha Entrada') !== -1 ? 'Fecha Entrada'
        : (t.headers.indexOf('Día') !== -1 ? 'Día' : (t.headers.indexOf('Fecha intento') !== -1 ? 'Fecha intento' : null));
      if (muestraFecha) {
        var crudo = t.rows[0][muestraFecha];
        var parseada = parseFecha(crudo);
        Logger.log('  Valor crudo de "' + muestraFecha + '": ' + crudo + '  (tipo: ' + typeof crudo +
          ', ¿parseFecha la reconoce?: ' + parseada + ')');
      }
    }
  }

  mostrar('Asistencias', encontrarArchivo(carpeta, /ReporteAsistenciaEmpresa.*\.xls/i, false));
  mostrar('Inasistencias', encontrarArchivo(carpeta, /ReporteInasistenciaEmpresa.*\.xls/i, false));
  var fallidos = encontrarArchivos(carpeta, /ReporteRegistroFallido.*\.xls/i);
  fallidos.forEach(function (f) { mostrar('Registro Fallido', f); });
  mostrar('Maestro', encontrarArchivo(carpeta, /Trabajadores_.*\.xls/i, false));
  mostrar('Calendario', encontrarArchivo(carpeta, /calendario.*\.xlsx?/i, false));

  Logger.log('=== Carpeta de salida (Asistencias_quincena) ===');
  var carpetaReportes = DriveApp.getFolderById(CONFIG.ID_CARPETA_REPORTES);
  var itR = carpetaReportes.getFiles();
  while (itR.hasNext()) {
    var fr = itR.next();
    Logger.log('  "' + fr.getName() + '"  (modificado=' + fr.getLastUpdated() + ')');
  }

  Logger.log('=== Fin diagnóstico ===');
}

// ============================================================================
// LOCALIZACIÓN DE ARCHIVOS EN DRIVE
// ============================================================================

function encontrarArchivo(carpeta, regex, obligatorio) {
  var coincidencias = encontrarArchivos(carpeta, regex);
  if (coincidencias.length === 0) {
    if (obligatorio) {
      throw new Error('No se encontró ningún archivo que coincida con ' + regex + ' en la carpeta.');
    }
    return null;
  }
  coincidencias.sort(function (a, b) { return b.getLastUpdated() - a.getLastUpdated(); });
  if (coincidencias.length > 1) {
    Logger.log('Aviso: se encontró más de un archivo para ' + regex + '; se usa el más reciente: ' +
      coincidencias[0].getName() + ' (se ignoran: ' +
      coincidencias.slice(1).map(function (f) { return f.getName(); }).join(', ') + ')');
  }
  return coincidencias[0];
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

// ============================================================================
// UTILIDADES DE FECHA / HORA
// ============================================================================

function pad2(n) { return (n < 10 ? '0' : '') + n; }

// Devuelve fecha normalizada como string 'yyyy-MM-dd' (clave interna), o null.
function parseFecha(valor) {
  if (valor === null || valor === undefined || valor === '') return null;
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return dateKey(valor);
  }
  if (typeof valor === 'number') {
    // Número de serie de Excel/Sheets (días desde 1899-12-30).
    return dateKey(serialAFecha(valor));
  }
  var s = String(valor).trim();
  if (s === '' || s === '-' || s.toLowerCase() === 'nan' || s.toLowerCase() === 'nat') return null;

  var m;
  // dd/mm/yyyy
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]);
  // dd-mm-yyyy
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) return m[3] + '-' + pad2(+m[2]) + '-' + pad2(+m[1]);
  // yyyy-mm-dd
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]);
  return null;
}

// Convierte un número de serie de Excel/Sheets a un objeto Date (en UTC, sin
// desfase por zona horaria) que representa el mismo día/hora "de pared".
function serialAFecha(serial) {
  var ms = Math.round((serial - 25569) * 86400 * 1000); // 25569 = días entre 1899-12-30 y 1970-01-01
  var d = new Date(ms);
  // Usamos getUTC* porque construimos el Date desde epoch UTC directamente.
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

// Formato corto legible para mostrar una fecha clave 'yyyy-MM-dd' en las
// columnas resumen (ej. '12-ago').
var MESES_CORTOS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
function formatearFechaCorta(fechaKey) {
  var d = dateKeyToDate(fechaKey);
  return pad2(d.getDate()) + '-' + MESES_CORTOS[d.getMonth()];
}

// Devuelve minutos-desde-medianoche (number), o null.
function parseHora(valor) {
  if (valor === null || valor === undefined || valor === '') return null;
  if (Object.prototype.toString.call(valor) === '[object Date]') {
    return valor.getHours() * 60 + valor.getMinutes() + valor.getSeconds() / 60;
  }
  if (typeof valor === 'number') {
    // Fracción de día (0.5 = 12:00). Si viene con parte entera (fecha+hora), se ignora la parte entera.
    var frac = valor - Math.floor(valor);
    var totalMin = frac * 1440;
    return Math.round(totalMin * 100) / 100;
  }
  var s = String(valor).trim();
  if (s === '' || s === '-' || s.toLowerCase() === 'nan' || s.toLowerCase() === 'nat') return null;
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

// Devuelve {programada, real} en minutos, o {null,null}.
function parseRangoHorario(valor) {
  if (valor === null || valor === undefined || valor === '') return { inicio: null, fin: null };
  var s = String(valor).trim();
  if (s === '' || s === '-' || s.toLowerCase() === 'nan') return { inicio: null, fin: null };
  var partes = s.split('-');
  if (partes.length !== 2) return { inicio: null, fin: null };
  return { inicio: parseHora(partes[0]), fin: parseHora(partes[1]) };
}

function minutosDeDiferencia(horaProgramadaMin, horaRealMin) {
  return horaRealMin - horaProgramadaMin;
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
  // arr de claves 'yyyy-MM-dd'
  var set = {};
  arr.forEach(function (k) { set[k] = true; });
  return Object.keys(set).sort();
}
function uniqueSortedKeys(arr) { return uniqueSorted(arr); }

// ============================================================================
// NORMALIZACIÓN DE ENCABEZADOS (por si el .xls llega con acentos rotos o
// caracteres invisibles tras la conversión de Drive)
// ============================================================================

var HEADERS_CONOCIDOS = [
  'Recinto', 'RUT', 'Primer Apellido', 'Segundo Apellido', 'Nombre', 'Especialidad',
  'Área', 'Contrato', 'Supervisor', 'Fecha Entrada', 'Hora Entrada', 'Fecha Salida',
  'Hora Salida', 'Sigla Turno', 'Horario Turno', 'Día', 'Horario', 'Motivo',
  'ID Dispositivo', 'Error al marcar', 'Sentido', 'Fecha intento', 'Hora intento',
  'Empresa', 'Código', 'Ciudad', 'Comuna', 'Turno'
];

function normalizarClave(s) {
  return String(s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // quita acentos
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

var HEADERS_NORMALIZADOS = (function () {
  var m = {};
  HEADERS_CONOCIDOS.forEach(function (h) { m[normalizarClave(h)] = h; });
  return m;
})();

function encabezadoCanonico(raw) {
  var limpio = String(raw).replace(/[\u200B-\u200D\uFEFF\u00A0]/g, ' ').trim();
  var norm = normalizarClave(limpio);
  return HEADERS_NORMALIZADOS[norm] || limpio;
}

// Normalización de encabezados para la hoja de quincenas del calendario
// (acepta "Quincena_nom", "Quincena Nom", "quincena  nom", etc.)
function normalizarClaveQuincena(s) {
  return String(s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[_\s]+/g, ' ').trim();
}

// ============================================================================
// CARGA DE DATOS FUENTE
//
// Si el archivo YA es un Google Sheet nativo (Drive lo convirtió al subirlo,
// que es el caso normal cuando ves URLs docs.google.com/spreadsheets/...),
// se abre directamente con SpreadsheetApp, sin pasos extra.
// Si el archivo sigue siendo un binario de Office (.xls/.xlsx real, mimeType
// distinto), se hace una copia temporal convertida a Sheets y se borra al
// terminar. Ese segundo caso requiere el servicio avanzado "Drive API"
// habilitado (Servicios -> + -> Google Drive API).
// ============================================================================

function abrirComoSheet(file) {
  // Se intenta abrir directo primero: funciona tanto si Drive ya convirtió el
  // archivo a Sheets nativo, como en el caso (común con archivos subidos como
  // .xls) en que Drive lo deja editable directo sin conversión explícita.
  try {
    var ssDirecto = SpreadsheetApp.open(file);
    return { ss: ssDirecto, tempId: null };
  } catch (eDirecto) {
    Logger.log('  "' + file.getName() + '" no se pudo abrir directo (' + eDirecto.message +
      '); se intentará convertir vía Drive API...');
  }

  if (typeof Drive === 'undefined') {
    throw new Error('El archivo "' + file.getName() + '" no se pudo abrir directamente y el servicio avanzado ' +
      '"Drive API" no está habilitado. Ve al editor -> Servicios (ícono +) -> agrega "Google Drive API" ' +
      '(identificador "Drive") y vuelve a ejecutar.');
  }

  var recurso = {
    name: 'TMP_' + file.getName() + '_' + new Date().getTime(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var copiado = Drive.Files.copy(recurso, file.getId());
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
    var headers = valores[0].map(function (h) { return encabezadoCanonico(h); });
    var rows = [];
    for (var i = 1; i < valores.length; i++) {
      var fila = valores[i];
      var vacio = fila.every(function (v) { return v === '' || v === null; });
      if (vacio) continue;
      var obj = {};
      headers.forEach(function (h, j) { obj[h] = fila[j]; });
      rows.push(obj);
    }
    Logger.log('  "' + file.getName() + '" (hoja "' + sheet.getName() + '"): ' +
      headers.length + ' columnas, ' + rows.length + ' filas.');
    return { headers: headers, rows: rows };
  } finally {
    cerrarSheetTemporal(abierto);
  }
}

function cargarRegistrosFallidos(files) {
  if (files.length === 0) return { headers: [], rows: [] };
  var headers = null;
  var rows = [];
  files.forEach(function (f) {
    var t = cargarTabla(f);
    if (!headers) headers = t.headers;
    rows = rows.concat(t.rows);
  });
  return { headers: headers, rows: rows };
}

function cargarMaestro(file) {
  var t = cargarTabla(file);
  var vistos = {};
  var out = [];
  var duplicados = {};
  t.rows.forEach(function (r) {
    var rut = limpiarRut(r['RUT']);
    r.RUT = rut;
    if (vistos[rut]) {
      duplicados[rut] = duplicados[rut] || [];
      duplicados[rut].push(r['Recinto']);
      return; // se conserva la primera fila (drop_duplicates keep='first')
    }
    vistos[rut] = true;
    out.push(r);
  });
  var dupKeys = Object.keys(duplicados);
  if (dupKeys.length) {
    Logger.log('Aviso: RUTs duplicados en la lista maestra (se conserva la primera fila): ' + dupKeys.join(', '));
  }
  return out;
}

function calcularLocalidad(recinto) {
  if (recinto && String(recinto).toLowerCase().indexOf('veracruz') !== -1) return 'veracruz';
  return 'cdmx';
}

function cargarFestivos(file) {
  var abierto = abrirComoSheet(file);
  try {
    var sheet = abierto.ss.getSheetByName('festivos');
    if (!sheet) {
      Logger.log('Aviso: no se encontró la hoja "festivos" en ' + file.getName());
      return new Set();
    }
    var valores = sheet.getDataRange().getValues();
    var claves = [];
    valores.forEach(function (fila) {
      var v = fila[1]; // columna B
      var f = parseFecha(v);
      if (f) claves.push(f);
    });
    return new Set(claves);
  } finally {
    cerrarSheetTemporal(abierto);
  }
}

// ============================================================================
// QUINCENAS (calendario.xlsx) -> generación de un Sheet nuevo por periodo
//
// Busca, dentro del mismo archivo de calendario, una hoja que tenga las
// columnas "Semana 1 Inicio", "Semana 2 Fin" y "Quincena_nom" (el nombre de
// la hoja no importa, se detecta por encabezados). Cada fila representa un
// periodo de quincena.
// ============================================================================

function cargarQuincenas(file) {
  var abierto = abrirComoSheet(file);
  try {
    var hojas = abierto.ss.getSheets();
    for (var h = 0; h < hojas.length; h++) {
      var sheet = hojas[h];
      var valores = sheet.getDataRange().getValues();
      if (valores.length < 2) continue;

      var headersRaw = valores[0];
      var colInicio = -1, colFin = -1, colNombre = -1;
      headersRaw.forEach(function (hVal, idx) {
        var norm = normalizarClaveQuincena(hVal);
        if (norm === 'semana 1 inicio') colInicio = idx;
        else if (norm === 'semana 2 fin') colFin = idx;
        else if (norm === 'quincena nom') colNombre = idx;
      });

      if (colInicio === -1 || colFin === -1 || colNombre === -1) continue; // no es esta hoja

      var out = [];
      for (var i = 1; i < valores.length; i++) {
        var fila = valores[i];
        var inicio = parseFecha(fila[colInicio]);
        var fin = parseFecha(fila[colFin]);
        var nombre = fila[colNombre] ? String(fila[colNombre]).trim() : '';
        if (!fin || !nombre) continue;
        out.push({ inicio: inicio, fin: fin, nombre: nombre });
      }
      Logger.log('Quincenas leídas de la hoja "' + sheet.getName() + '" de ' + file.getName() + ': ' + out.length);
      return out;
    }
    Logger.log('Aviso: no se encontró ninguna hoja en "' + file.getName() +
      '" con las columnas "Semana 1 Inicio", "Semana 2 Fin", "Quincena_nom". No se generarán archivos de quincena.');
    return [];
  } finally {
    cerrarSheetTemporal(abierto);
  }
}

// Revisa cada quincena: si hoy ya alcanzó o pasó su "Semana 2 Fin" y todavía
// no existe un archivo con el nombre de "quincena_nom" en la carpeta de
// REPORTES (salida), crea un Google Sheet nuevo y en blanco (sin copiar
// hojas ni datos) con ese nombre, dentro de esa misma carpeta.
function generarArchivosQuincenaSiCorresponde(carpetaReportes, quincenas) {
  var hoy = dateKey(new Date());
  quincenas.forEach(function (q) {
    if (hoy < q.fin) return; // esta quincena todavía no termina

    var existentes = carpetaReportes.getFilesByName(q.nombre);
    if (existentes.hasNext()) {
      Logger.log('Ya existe un archivo llamado "' + q.nombre + '" en la carpeta de reportes; no se genera de nuevo.');
      return;
    }

    var nuevoSs = SpreadsheetApp.create(q.nombre);
    var nuevoFile = DriveApp.getFileById(nuevoSs.getId());
    moverArchivoACarpeta(nuevoFile, carpetaReportes);
    Logger.log('Se generó el Sheet en blanco "' + q.nombre + '" en la carpeta de reportes para la quincena que terminó el ' + q.fin + '.');
  });
}

// SpreadsheetApp.create() crea el archivo en la raíz de Drive; esta función
// lo mueve (no lo copia) a la carpeta destino indicada.
function moverArchivoACarpeta(file, carpetaDestino) {
  var padres = file.getParents();
  while (padres.hasNext()) {
    var padre = padres.next();
    if (padre.getId() !== carpetaDestino.getId()) {
      padre.removeFile(file);
    }
  }
  carpetaDestino.addFile(file);
}

// ============================================================================
// CÁLCULO DE CÓDIGO DIARIO DE ASISTENCIA
//
// PRIORIDAD (por RUT + fecha):
//   1. Asistencia real (ReporteAsistenciaEmpresa) -> siempre gana, aunque esa
//      misma persona/fecha tenga también una fila de inasistencia (p.ej. en
//      otro recinto). Se ignora la inasistencia por completo en ese caso.
//   2. Si no hay asistencia real, pero sí un intento en ReporteRegistroFallido
//      ese mismo día -> se cuenta como asistencia (A), con nota en la celda.
//   3. Si no aplica ninguna de las dos, se usa el Motivo de la inasistencia.
// ============================================================================

function calcularCodigoAsistencia(horaProgramadaMin, horaRealMin, toleranciaMin) {
  var diferencia = minutosDeDiferencia(horaProgramadaMin, horaRealMin);
  if (diferencia > toleranciaMin) {
    var minutosTarde = Math.round(diferencia - toleranciaMin);
    return 'A-' + minutosTarde;
  }
  return 'A';
}

function indexarFallidosPorFecha(tablaFallidos) {
  var fallidosPorFecha = {};
  tablaFallidos.rows.forEach(function (row) {
    var f = parseFecha(row['Fecha intento']);
    var rut = limpiarRut(row['RUT']);
    if (!f || !rut) return;
    fallidosPorFecha[f] = fallidosPorFecha[f] || {};
    fallidosPorFecha[f][rut] = fallidosPorFecha[f][rut] || [];
    fallidosPorFecha[f][rut].push(
      (row['Sentido'] || '') + ' ' + (row['Hora intento'] || '') + ' (' + (row['Error al marcar'] || '') + ')'
    );
  });
  return fallidosPorFecha;
}

function construirDatosPorFecha(tablaAsistencias, tablaInasistencias, tablaFallidos, toleranciaMin) {
  var datos = {};
  var horarioMasReciente = {};
  var tieneAsistencia = {}; // tieneAsistencia[fecha][rut] = true

  var fallidosPorFecha = indexarFallidosPorFecha(tablaFallidos);

  // --- 1) Asistencias reales: siempre se procesan primero y ganan ---
  tablaAsistencias.rows.forEach(function (row) {
    var f = parseFecha(row['Fecha Entrada']);
    var rut = limpiarRut(row['RUT']);
    if (!f || !rut) return;

    var horarioTurno = row['Horario Turno'];
    var rango = parseRangoHorario(horarioTurno);
    var horaReal = parseHora(row['Hora Entrada']);

    if (horarioTurno && String(horarioTurno).trim() !== '' && String(horarioTurno).trim() !== '-') {
      horarioMasReciente[rut] = String(horarioTurno).trim();
    }

    var codigo;
    if (rango.inicio === null || horaReal === null) {
      codigo = 'A';
    } else {
      codigo = calcularCodigoAsistencia(rango.inicio, horaReal, toleranciaMin);
    }
    datos[f] = datos[f] || {};
    datos[f][rut] = { codigo: codigo, comentario: null };

    tieneAsistencia[f] = tieneAsistencia[f] || {};
    tieneAsistencia[f][rut] = true;
  });

  // --- 2) Inasistencias: solo aplican si esa persona NO tiene asistencia real ese día ---
  tablaInasistencias.rows.forEach(function (row) {
    var f = parseFecha(row['Día']);
    var rut = limpiarRut(row['RUT']);
    if (!f || !rut) return;

    if (tieneAsistencia[f] && tieneAsistencia[f][rut]) {
      // Ya hay asistencia real ese día (posiblemente en otro recinto) -> se
      // ignora esta inasistencia, la persona ya quedó marcada como A / A-N.
      return;
    }

    var intentosFallidos = fallidosPorFecha[f] && fallidosPorFecha[f][rut];
    if (intentosFallidos) {
      // No hay asistencia real, pero sí un intento de checada fallido ese
      // mismo día -> se cuenta como asistencia.
      datos[f] = datos[f] || {};
      datos[f][rut] = {
        codigo: 'A',
        comentario: 'Asistencia inferida por intento de checada fallido: ' + intentosFallidos.join('; ')
      };
      return;
    }

    var motivo = String(row['Motivo'] || '').trim();
    var codigo = MOTIVO_A_CODIGO.hasOwnProperty(motivo) ? MOTIVO_A_CODIGO[motivo] : null;
    if (codigo === null) {
      Logger.log('Aviso: motivo "' + motivo + '" no está en el catálogo (RUT ' + rut + ', ' + f + '); se deja tal cual.');
      codigo = motivo;
    }

    datos[f] = datos[f] || {};
    datos[f][rut] = { codigo: codigo, comentario: null };
  });

  return { datos: datos, horarioMasReciente: horarioMasReciente };
}

// ============================================================================
// CONSTRUCCIÓN DE "FALTAS" Y "RETARDOS" (solo día más reciente de la corrida)
// ============================================================================

function construirFaltas(tablaAsistencias, tablaInasistencias, tablaFallidos, fechaObjetivo) {
  // RUTs con asistencia real ese día (en cualquier recinto)
  var asistenciaHoy = {};
  tablaAsistencias.rows.forEach(function (r) {
    if (parseFecha(r['Fecha Entrada']) === fechaObjetivo) {
      asistenciaHoy[limpiarRut(r['RUT'])] = true;
    }
  });

  // RUTs con intento de checada fallido ese día
  var fallidosHoy = {};
  tablaFallidos.rows.forEach(function (r) {
    if (parseFecha(r['Fecha intento']) === fechaObjetivo) {
      var rut = limpiarRut(r['RUT']);
      fallidosHoy[rut] = fallidosHoy[rut] || [];
      fallidosHoy[rut].push(
        (r['Sentido'] || '') + ' ' + (r['Hora intento'] || '') + ' (' + (r['Error al marcar'] || '') + ')'
      );
    }
  });

  var diaRows = tablaInasistencias.rows.filter(function (r) { return parseFecha(r['Día']) === fechaObjetivo; });
  var candidatas = diaRows.filter(function (r) { return String(r['Motivo'] || '').trim() === '-'; });

  var columnas = ['Recinto', 'Día', 'RUT', 'Nombre completo', 'Área', 'Contrato',
    'Supervisor', 'Horario', 'Sigla Turno', 'Estatus', 'Detalle Registro Fallido'];

  var filas = [];
  candidatas.forEach(function (row) {
    var rut = limpiarRut(row['RUT']);

    // Si en realidad sí asistió (en cualquier recinto) o tiene intento
    // fallido ese día, ya no es una falta -> se excluye de esta hoja.
    if (asistenciaHoy[rut]) return;
    if (fallidosHoy[rut]) return;

    filas.push({
      'Recinto': row['Recinto'] || '',
      'Día': row['Día'] || '',
      'RUT': rut,
      'Nombre completo': ((row['Primer Apellido'] || '') + ' ' + (row['Nombre'] || '')).trim(),
      'Área': row['Área'] || '',
      'Contrato': row['Contrato'] || '',
      'Supervisor': row['Supervisor'] || '',
      'Horario': row['Horario'] || '',
      'Sigla Turno': row['Sigla Turno'] || '',
      'Estatus': 'Falta',
      'Detalle Registro Fallido': ''
    });
  });

  return { columnas: columnas, filas: filas };
}

function construirRetardos(tablaAsistencias, fechaObjetivo, toleranciaMin) {
  var diaRows = tablaAsistencias.rows.filter(function (r) { return parseFecha(r['Fecha Entrada']) === fechaObjetivo; });

  var columnas = ['Recinto', 'Fecha', 'RUT', 'Nombre completo', 'Área', 'Supervisor',
    'Sigla Turno', 'Hora Entrada Programada', 'Hora Entrada Real', 'Minutos Tarde', 'Tiempo Tarde'];

  var filas = [];
  diaRows.forEach(function (row) {
    var rango = parseRangoHorario(row['Horario Turno']);
    var horaReal = parseHora(row['Hora Entrada']);
    if (rango.inicio === null || horaReal === null) return;

    var diferenciaMin = minutosDeDiferencia(rango.inicio, horaReal);
    if (diferenciaMin > toleranciaMin) {
      var nombre = [row['Primer Apellido'], row['Segundo Apellido'], row['Nombre']]
        .filter(function (p) { return p; }).join(' ').trim();
      filas.push({
        'Recinto': row['Recinto'] || '',
        'Fecha': row['Fecha Entrada'] || '',
        'RUT': limpiarRut(row['RUT']),
        'Nombre completo': nombre,
        'Área': row['Área'] || '',
        'Supervisor': row['Supervisor'] || '',
        'Sigla Turno': row['Sigla Turno'] || '',
        'Hora Entrada Programada': horaMinutosAString(rango.inicio, false),
        'Hora Entrada Real': horaMinutosAString(horaReal, true),
        'Minutos Tarde': Math.round(diferenciaMin),
        'Tiempo Tarde': formatearMinutos(diferenciaMin)
      });
    }
  });

  filas.sort(function (a, b) { return b['Minutos Tarde'] - a['Minutos Tarde']; });
  return { columnas: columnas, filas: filas };
}

// ============================================================================
// LECTURA DEL HISTORIAL EXISTENTE DE "Reporte de asistencia"
// ============================================================================

function leerHistorialExistente(ss) {
  var historial = {};
  var meta = {};
  var fechas = [];

  var sheet = ss.getSheetByName(NOMBRES_HOJA.REPORTE);
  if (!sheet) return { historial: historial, meta: meta, fechas: fechas };

  var ultimaCol = sheet.getLastColumn();
  var ultimaFila = sheet.getLastRow();
  if (ultimaFila < PRIMERA_FILA_DATOS || ultimaCol < PRIMERA_COL_FECHA) {
    return { historial: historial, meta: meta, fechas: fechas };
  }

  // Detectar columnas de fecha en la fila de encabezado. Se detiene en la
  // primera columna que ya no parsea como fecha (ahí empiezan las columnas
  // resumen "Retardos" / "Faltas" y, más a la derecha, la leyenda).
  var colFechas = {}; // col (1-indexed) -> 'yyyy-MM-dd'
  for (var col = PRIMERA_COL_FECHA; col <= ultimaCol; col++) {
    var valor = sheet.getRange(FILA_ENCABEZADO, col).getValue();
    var f = parseFecha(valor);
    if (!f) break;
    colFechas[col] = f;
    fechas.push(f);
  }

  var rango = sheet.getRange(PRIMERA_FILA_DATOS, 1, ultimaFila - PRIMERA_FILA_DATOS + 1, ultimaCol).getValues();
  for (var i = 0; i < rango.length; i++) {
    var fila = rango[i];
    var rut = fila[COL_RUT_OCULTA - 1];
    if (rut === null || rut === undefined || String(rut).trim() === '') break;
    rut = limpiarRut(rut);

    meta[rut] = {
      'Nombre completo': fila[1],
      'Área': fila[2],
      'Horario Turno': fila[3],
      'Recinto': fila[4],
      'Localidad': fila[5],
      'Supervisor': fila[6]
    };

    Object.keys(colFechas).forEach(function (colStr) {
      var colIdx = +colStr;
      var v = fila[colIdx - 1];
      if (v !== null && v !== undefined && String(v).trim() !== '') {
        var f = colFechas[colIdx];
        historial[f] = historial[f] || {};
        historial[f][rut] = String(v).trim();
      }
    });
  }

  return { historial: historial, meta: meta, fechas: fechas };
}

// ============================================================================
// ESCRITURA DE HOJAS
// ============================================================================

function obtenerOCrearHojaLimpia(ss, nombre) {
  var sheet = ss.getSheetByName(nombre);
  if (sheet) {
    sheet.clear();
    sheet.clearNotes();
  } else {
    sheet = ss.insertSheet(nombre);
  }
  return sheet;
}

function escribirHojaTabular(ss, nombreHoja, columnas, filas, colorearEstatus) {
  var sheet = obtenerOCrearHojaLimpia(ss, nombreHoja);
  if (columnas.length === 0) return;

  // Encabezado
  var headerRange = sheet.getRange(1, 1, 1, columnas.length);
  headerRange.setValues([columnas]);
  headerRange.setFontFamily('Arial').setFontWeight('bold').setFontColor(COLOR_HEADER_TEXTO)
    .setBackground(COLOR_HEADER_FONDO).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 28);

  if (filas.length > 0) {
    var datos = filas.map(function (fila) { return columnas.map(function (c) { return fila[c]; }); });
    var dataRange = sheet.getRange(2, 1, filas.length, columnas.length);
    dataRange.setValues(datos);
    dataRange.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');

    if (colorearEstatus && columnas.indexOf('Estatus') !== -1) {
      var idxEstatus = columnas.indexOf('Estatus');
      for (var i = 0; i < filas.length; i++) {
        var val = String(filas[i]['Estatus'] || '');
        var color = val.indexOf('Revisar') !== -1 ? COLOR_REVISAR : COLOR_FALTA;
        sheet.getRange(2 + i, 1, 1, columnas.length).setBackground(color);
      }
    }
  }

  // Autoajuste de columnas
  for (var c = 1; c <= columnas.length; c++) {
    sheet.autoResizeColumn(c);
  }

  var filaNota = filas.length + 3;
  sheet.getRange(filaNota, 1).setValue('Total de registros: ' + filas.length)
    .setFontFamily('Arial').setFontStyle('italic').setFontSize(9).setFontColor('#808080');
}

function escribirReporteAsistencia(ss, roster, fechasOrdenadas, datosFinales, comentarios) {
  var sheet = obtenerOCrearHojaLimpia(ss, NOMBRES_HOJA.REPORTE);

  var encabezados = ['RUT'].concat(COLS_META);
  var headerRange = sheet.getRange(FILA_ENCABEZADO, 1, 1, encabezados.length);
  headerRange.setValues([encabezados]);
  headerRange.setFontFamily('Arial').setFontWeight('bold').setFontColor(COLOR_HEADER_TEXTO)
    .setBackground(COLOR_HEADER_FONDO).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

  fechasOrdenadas.forEach(function (f, k) {
    var col = PRIMERA_COL_FECHA + k;
    var celda = sheet.getRange(FILA_ENCABEZADO, col);
    celda.setValue(dateKeyToDate(f));
    celda.setNumberFormat('dd-mmm-yy');
    celda.setFontFamily('Arial').setFontWeight('bold').setFontColor(COLOR_HEADER_TEXTO)
      .setBackground(COLOR_HEADER_FONDO).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  });

  // --- Columnas resumen: "Retardos" (conteo) y "Faltas" (fechas), justo
  //     después de la última columna de fecha ---
  var colRetardos = PRIMERA_COL_FECHA + fechasOrdenadas.length;
  var colFaltas = colRetardos + 1;

  var headerResumen = sheet.getRange(FILA_ENCABEZADO, colRetardos, 1, 2);
  headerResumen.setValues([['Retardos', 'Faltas']]);
  headerResumen.setFontFamily('Arial').setFontWeight('bold').setFontColor(COLOR_HEADER_TEXTO)
    .setBackground(COLOR_HEADER_FONDO).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

  sheet.setRowHeight(FILA_ENCABEZADO, 26);
  sheet.setFrozenRows(PRIMERA_FILA_DATOS - 1);
  sheet.setFrozenColumns(PRIMERA_COL_FECHA - 1);

  var filaBase = PRIMERA_FILA_DATOS;
  roster.forEach(function (persona, idx) {
    var fila = filaBase + idx;
    var rut = persona.RUT;

    sheet.getRange(fila, 1).setValue(rut).setFontFamily('Arial').setFontSize(9).setFontColor('#808080');

    var valoresMeta = [
      persona['Nombre completo'], persona['Área'], persona['Horario Turno'],
      persona['Recinto'], persona['Localidad'], persona['Supervisor']
    ];
    var metaRange = sheet.getRange(fila, 2, 1, valoresMeta.length);
    metaRange.setValues([valoresMeta]);
    metaRange.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');

    var totalRetardos = 0;
    var fechasFalta = [];

    fechasOrdenadas.forEach(function (f, k) {
      var col = PRIMERA_COL_FECHA + k;
      var codigo = (datosFinales[f] && datosFinales[f][rut]) || '-';
      var celda = sheet.getRange(fila, col);
      celda.setValue(codigo);
      celda.setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

      if (codigo === 'DF') {
        celda.setBackground(COLOR_DF);
      } else if (codigo === 'F' || (typeof codigo === 'string' && codigo !== 'A' && codigo !== '-' && codigo !== 'DF' && codigo.indexOf('A-') !== 0)) {
        celda.setBackground(COLOR_FALTA);
      } else if (typeof codigo === 'string' && codigo.indexOf('A-') === 0) {
        celda.setBackground(COLOR_REVISAR);
      }

      var textoComentario = comentarios[f] && comentarios[f][rut];
      if (textoComentario) {
        celda.setNote(textoComentario);
      }

      // --- Acumular para las columnas resumen ---
      if (typeof codigo === 'string' && codigo.indexOf('A-') === 0) {
        totalRetardos++;
      } else if (codigo === 'F') {
        fechasFalta.push(formatearFechaCorta(f));
      }
    });

    var celdaRetardos = sheet.getRange(fila, colRetardos);
    celdaRetardos.setValue(totalRetardos).setFontFamily('Arial').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (totalRetardos > 0) celdaRetardos.setBackground(COLOR_REVISAR);

    var celdaFaltas = sheet.getRange(fila, colFaltas);
    celdaFaltas.setValue(fechasFalta.join(', ')).setFontFamily('Arial').setFontSize(9)
      .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);
    if (fechasFalta.length > 0) celdaFaltas.setBackground(COLOR_FALTA);
  });

  // Ocultar columna RUT (A)
  sheet.hideColumns(1);

  // Anchos de columna
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 170);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 190);
  fechasOrdenadas.forEach(function (f, k) {
    sheet.setColumnWidth(PRIMERA_COL_FECHA + k, 60);
  });
  sheet.setColumnWidth(colRetardos, 75);
  sheet.setColumnWidth(colFaltas, 220);

  // Leyenda a la derecha de las fechas (ahora después de Retardos/Faltas)
  var colLeyenda = colFaltas + 2;
  var headerLeyenda = sheet.getRange(FILA_ENCABEZADO, colLeyenda, 1, 2);
  headerLeyenda.setValues([['Código', 'Significado']]);
  headerLeyenda.setFontFamily('Arial').setFontWeight('bold').setFontColor(COLOR_HEADER_TEXTO).setBackground(COLOR_HEADER_FONDO);

  LEYENDA.forEach(function (par, i) {
    var fila = FILA_ENCABEZADO + 1 + i;
    sheet.getRange(fila, colLeyenda).setValue(par[0]).setFontFamily('Arial').setFontWeight('bold').setFontSize(10);
    sheet.getRange(fila, colLeyenda + 1).setValue(par[1]).setFontFamily('Arial').setFontSize(10);
  });
  sheet.setColumnWidth(colLeyenda, 60);
  sheet.setColumnWidth(colLeyenda + 1, 380);
}

// ============================================================================
// CORREO DE RESUMEN
// ============================================================================

function enviarCorreoResumen(resumen) {
  if (!CONFIG.CORREOS_ASISTENCIA.length) return;

  var html = ''
    + '<div style="font-family:Arial, sans-serif; font-size:14px; color:#222;">'
    + '<p>Estimados,</p>'
    + '<p>Se actualizó el <strong>reporte de asistencias, retardos, faltas y el histórico</strong> '
    + 'correspondiente a la quincena actual. Pueden consultarlo directamente en el Sheet:</p>'
    + '<p><a href="' + resumen.url + '">' + resumen.url + '</a></p>'
    + '<table style="border-collapse:collapse; margin:10px 0;">'
    + '<tr><th style="border:1px solid #999; background:#1F4E78; color:#fff; padding:6px 10px; text-align:left;">Concepto</th>'
    + '<th style="border:1px solid #999; background:#1F4E78; color:#fff; padding:6px 10px; text-align:left;">Detalle</th></tr>'
    + '<tr><td style="border:1px solid #999; padding:6px 10px;">Faltas (' + resumen.fecha + ')</td>'
    + '<td style="border:1px solid #999; padding:6px 10px;">' + resumen.faltas + ' registros (' + resumen.faltasRevision + ' para revisión)</td></tr>'
    + '<tr><td style="border:1px solid #999; padding:6px 10px;">Retardos (' + resumen.fecha + ')</td>'
    + '<td style="border:1px solid #999; padding:6px 10px;">' + resumen.retardos + ' registros</td></tr>'
    + '<tr><td style="border:1px solid #999; padding:6px 10px;">Reporte de asistencia</td>'
    + '<td style="border:1px solid #999; padding:6px 10px;">' + resumen.fechas + ' fechas x ' + resumen.personas + ' personas</td></tr>'
    + '</table>'
    + '<p>En caso de requerir aclaraciones o notificar algún detalle sobre las inasistencias, '
    + 'favor de hacerlo por este medio o al correo <strong>ccarbajal@abcsc.mx</strong>.</p>'
    + '<p>Saludos cordiales,.</p>'
    + '</div>';

  MailApp.sendEmail({
    to: CONFIG.CORREOS_ASISTENCIA.join(','),
    subject: CONFIG.ASUNTO_CORREO,
    htmlBody: html
  });
  Logger.log('Correo de asistencia enviado a ' + CONFIG.CORREOS_ASISTENCIA.join(', '));
}