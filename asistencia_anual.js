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
 *  - `generarExcelAnual`: inicia (o continúa, si hay uno a medias) el proceso
 *    que genera el .xlsx. Se continúa solo por triggers hasta terminar; puedes
 *    seguir el avance en Ejecuciones.
 *  - `crearTriggerAnualSemanal` (opcional, correr UNA vez): lo ejecuta solo
 *    cada lunes.
 *
 * NOTA: todas las funciones y variables internas viven dentro del objeto
 * `AsistenciaAnual`, así que este archivo puede convivir en el MISMO proyecto
 * de Apps Script que asistencia.js sin chocar nombres (main, CONFIG,
 * parseFecha, onOpen, etc.).
 *
 * PROCESAMIENTO POR LOTES (mismo esquema que SN/main.js):
 *  Apps Script corta cada ejecución a los 6 minutos, y con los reportes de
 *  todo el año eso no alcanza. Por eso el proceso avanza por paquetes:
 *   1. La primera ejecución lista los archivos y guarda el estado
 *      ("leyendo", "generando").
 *   2. Lee archivo por archivo, acumulando los datos. Antes de cada archivo
 *      revisa el tiempo; al llegar a LIMITE_SEGUNDOS guarda el avance y
 *      programa un trigger de continuación a 1 minuto (`continuarExcelAnual`).
 *   3. Cuando ya leyó todo, genera el Excel (si ya no queda tiempo en esa
 *      ejecución, lo deja para la siguiente continuación) y limpia el estado
 *      y el trigger de continuación.
 *  Como PropertiesService solo admite ~9 KB por valor, en Properties se
 *  guarda solo el estado y el ID de un archivo JSON de trabajo
 *  ("_estado_asistencia_anual (no borrar).json", en la carpeta de salida)
 *  que contiene la lista de archivos, el índice y los datos acumulados. Ese
 *  JSON se manda a la papelera al terminar.
 *  - Si un archivo no se puede leer, se registra el error, se salta y se
 *    sigue con el siguiente (se lista al final en el log y en el correo).
 *  - Si el estado queda a medias por más de HORAS_EXPIRACION_ESTADO, o está
 *    corrupto, la siguiente ejecución lo descarta y empieza de cero.
 *  - `reiniciarExcelAnual`: función de emergencia que limpia el estado, el
 *    JSON de trabajo y el trigger de continuación.
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
  ASUNTO_CORREO: 'Asistencia anual',
  // --- Procesamiento por lotes ---
  // Segundos de lectura por ejecución antes de guardar y programar la
  // continuación (el tope de Apps Script es 360).
  LIMITE_SEGUNDOS: 270,
  // Si al terminar de leer ya pasaron más de estos segundos, la generación
  // del Excel se deja para la siguiente ejecución (necesita su propio tiempo).
  LIMITE_SEGUNDOS_PARA_GENERAR: 150,
  // Un proceso a medias más viejo que esto se descarta y se empieza de cero.
  HORAS_EXPIRACION_ESTADO: 24
};

// ============================================================================
// PUNTOS DE ENTRADA (visibles en el selector de funciones del editor)
// ============================================================================

function generarExcelAnual() {
  AsistenciaAnual.ejecutar();
}

// Handler del trigger de continuación (no la corras a mano, usa
// generarExcelAnual; hacen lo mismo).
function continuarExcelAnual() {
  AsistenciaAnual.ejecutar();
}

// Función de emergencia: descarta el proceso a medias.
function reiniciarExcelAnual() {
  AsistenciaAnual.reiniciar();
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
  // ORQUESTACIÓN POR LOTES
  // ==========================================================================

  var PROP = {
    ESTADO: 'AA_estado',             // 'leyendo' | 'generando'
    ARCHIVO_ESTADO: 'AA_archivoEstadoId',
    INICIADO: 'AA_iniciado'
  };
  var HANDLER_CONTINUACION = 'continuarExcelAnual';
  var NOMBRE_ARCHIVO_ESTADO = '_estado_asistencia_anual (no borrar).json';

  function ejecutar() {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(1000)) {
      Logger.log('Ya hay otra ejecución del reporte anual en curso; se omite esta.');
      return;
    }
    try {
      ejecutarConLock(new Date().getTime());
    } finally {
      lock.releaseLock();
    }
  }

  function ejecutarConLock(inicio) {
    var props = PropertiesService.getScriptProperties();
    var carpetaSalida = DriveApp.getFolderById(CONFIG_ANUAL.ID_CARPETA_SALIDA_ANUAL);

    var estado = cargarEstado(props);
    if (!estado) estado = inicializarEstado(props, carpetaSalida);
    if (!estado) return;

    if (props.getProperty(PROP.ESTADO) === 'leyendo') {
      var terminoLectura = procesarLote(estado, inicio);
      if (!terminoLectura) {
        guardarEstado(props, carpetaSalida, estado);
        programarContinuacion();
        return;
      }
      props.setProperty(PROP.ESTADO, 'generando');
      var transcurrido = (new Date().getTime() - inicio) / 1000;
      if (transcurrido >= CONFIG_ANUAL.LIMITE_SEGUNDOS_PARA_GENERAR) {
        Logger.log('Lectura terminada, pero quedan pocos segundos en esta ejecución (' +
          Math.round(transcurrido) + 's). La generación del Excel sigue en la próxima continuación.');
        guardarEstado(props, carpetaSalida, estado);
        programarContinuacion();
        return;
      }
    }

    generarReporteFinal(estado, carpetaSalida);
    cancelarTriggersContinuacion();
    limpiarEstado(props);
  }

  // Devuelve el estado guardado, o null si no hay (o si estaba corrupto o
  // expirado, en cuyo caso lo limpia).
  function cargarEstado(props) {
    var fase = props.getProperty(PROP.ESTADO);
    if (!fase) return null;

    var iniciado = props.getProperty(PROP.INICIADO);
    var horas = iniciado ? (new Date() - new Date(iniciado)) / 3600000 : Infinity;
    if (horas > CONFIG_ANUAL.HORAS_EXPIRACION_ESTADO) {
      Logger.log('Aviso: había un proceso a medias iniciado el ' + iniciado + ' (hace más de ' +
        CONFIG_ANUAL.HORAS_EXPIRACION_ESTADO + 'h). Se descarta y se empieza de cero.');
      limpiarEstado(props);
      return null;
    }

    try {
      var archivo = DriveApp.getFileById(props.getProperty(PROP.ARCHIVO_ESTADO));
      var estado = JSON.parse(archivo.getBlob().getDataAsString());
      if (!estado || !Array.isArray(estado.lista) || typeof estado.indice !== 'number' || !estado.acc) {
        throw new Error('estructura inválida');
      }
      Logger.log('Continuando proceso iniciado el ' + iniciado + ' (fase "' + fase + '", archivo ' +
        estado.indice + ' de ' + estado.lista.length + ').');
      return estado;
    } catch (e) {
      Logger.log('Aviso: estado corrupto o archivo de trabajo inaccesible (' + e.message + '). Reiniciando automáticamente...');
      limpiarEstado(props);
      return null;
    }
  }

  function inicializarEstado(props, carpetaSalida) {
    var carpetaEntrada = DriveApp.getFolderById(CONFIG_ANUAL.ID_CARPETA_ENTRADA_ANUAL);
    var archivos = localizarArchivos(carpetaEntrada, true);

    var lista = []
      .concat(archivos.asistencias.map(function (f) { return { id: f.getId(), nombre: f.getName(), tipo: 'asistencia' }; }))
      .concat(archivos.fallidos.map(function (f) { return { id: f.getId(), nombre: f.getName(), tipo: 'fallido' }; }))
      .concat(archivos.inasistencias.map(function (f) { return { id: f.getId(), nombre: f.getName(), tipo: 'inasistencia' }; }));

    var estado = {
      lista: lista,
      indice: 0,
      maestroId: archivos.maestro.getId(),
      calendarioId: archivos.calendario ? archivos.calendario.getId() : null,
      errores: [],
      filas: { asistencia: 0, fallido: 0, inasistencia: 0 },
      acc: nuevoAcumulado()
    };

    props.setProperty(PROP.INICIADO, new Date().toISOString());
    props.setProperty(PROP.ESTADO, 'leyendo');
    guardarEstado(props, carpetaSalida, estado);
    Logger.log('Estado inicializado. Archivos por leer: ' + lista.length);
    return estado;
  }

  // Lee archivos desde estado.indice hasta terminar o hasta llegar al límite
  // de tiempo. Devuelve true si ya leyó todos.
  function procesarLote(estado, inicio) {
    var total = estado.lista.length;
    Logger.log('▶ Leyendo desde el archivo ' + (estado.indice + 1) + ' de ' + total);

    while (estado.indice < total) {
      var transcurrido = (new Date().getTime() - inicio) / 1000;
      if (transcurrido >= CONFIG_ANUAL.LIMITE_SEGUNDOS) {
        Logger.log('⏱ Límite de tiempo alcanzado (' + Math.round(transcurrido) + 's). Se guarda el avance en el archivo ' +
          estado.indice + ' de ' + total + '.');
        return false;
      }

      var item = estado.lista[estado.indice];
      try {
        var tabla = cargarTabla(DriveApp.getFileById(item.id));
        acumularTabla(estado.acc, item.tipo, tabla, CONFIG_ANUAL.TOLERANCIA_RETARDO_MIN);
        estado.filas[item.tipo] += tabla.rows.length;
      } catch (e) {
        Logger.log('⚠️ Error leyendo "' + item.nombre + '": ' + e.message + '. Se omite y se sigue con el siguiente.');
        estado.errores.push(item.nombre + ': ' + e.message);
      }
      estado.indice++;
    }

    Logger.log('✅ Lectura completa. Filas -> asistencias: ' + estado.filas.asistencia +
      ', registros fallidos: ' + estado.filas.fallido + ', inasistencias: ' + estado.filas.inasistencia +
      (estado.errores.length ? ('. Archivos con error: ' + estado.errores.length) : ''));
    return true;
  }

  // Guarda el estado en un JSON nuevo de la carpeta de salida y manda el
  // anterior a la papelera (se crea uno nuevo en vez de sobrescribir para no
  // depender del límite de tamaño de setContent).
  function guardarEstado(props, carpetaSalida, estado) {
    var anteriorId = props.getProperty(PROP.ARCHIVO_ESTADO);
    var nuevo = carpetaSalida.createFile(
      Utilities.newBlob(JSON.stringify(estado), 'application/json', NOMBRE_ARCHIVO_ESTADO));
    props.setProperty(PROP.ARCHIVO_ESTADO, nuevo.getId());
    if (anteriorId) {
      try { DriveApp.getFileById(anteriorId).setTrashed(true); } catch (e) { /* ya no existe */ }
    }
  }

  function limpiarEstado(props) {
    var archivoId = props.getProperty(PROP.ARCHIVO_ESTADO);
    if (archivoId) {
      try { DriveApp.getFileById(archivoId).setTrashed(true); } catch (e) { /* ya no existe */ }
    }
    props.deleteProperty(PROP.ESTADO);
    props.deleteProperty(PROP.ARCHIVO_ESTADO);
    props.deleteProperty(PROP.INICIADO);
    Logger.log('🧹 Estado del reporte anual limpiado.');
  }

  function programarContinuacion() {
    cancelarTriggersContinuacion();
    ScriptApp.newTrigger(HANDLER_CONTINUACION).timeBased().after(60 * 1000).create();
    Logger.log('⏰ Trigger de continuación programado en 1 minuto.');
  }

  // Solo borra los triggers de continuación; el trigger semanal
  // (generarExcelAnual) no se toca.
  function cancelarTriggersContinuacion() {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === HANDLER_CONTINUACION &&
          t.getEventType() === ScriptApp.EventType.CLOCK) {
        ScriptApp.deleteTrigger(t);
      }
    });
  }

  function reiniciar() {
    limpiarEstado(PropertiesService.getScriptProperties());
    cancelarTriggersContinuacion();
    Logger.log('🔄 Listo. Puedes volver a ejecutar generarExcelAnual().');
  }

  // ==========================================================================
  // GENERAR EL EXCEL (con todos los archivos ya leídos)
  // ==========================================================================

  function generarReporteFinal(estado, carpetaSalida) {
    var t0 = new Date();
    var acc = estado.acc;
    var roster = cargarMaestro(DriveApp.getFileById(estado.maestroId));
    var festivos = estado.calendarioId ? cargarFestivos(DriveApp.getFileById(estado.calendarioId)) : new Set();
    Logger.log('Maestro: ' + roster.length + ' personas. Festivos: ' + festivos.size + ' fechas.');

    // --- Año y rango de fechas ---
    var todasFechas = uniqueSorted(Object.keys(acc.asis).concat(Object.keys(acc.fall)).concat(Object.keys(acc.inas)));
    if (todasFechas.length === 0) {
      Logger.log('No se encontraron fechas válidas en los reportes. Nada que hacer.');
      return;
    }
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

    var datos = resolverDatos(acc, prefijoAnio);
    var personas = construirPersonas(roster, acc, datos);
    Logger.log('Personas en el reporte: ' + personas.length);
    var matriz = construirMatriz(personas, fechas, festivos, datos);

    // --- Escribir en Sheet temporal y exportar a Excel ---
    var nombre = CONFIG_ANUAL.NOMBRE_ARCHIVO.replace('{ANIO}', anio);
    var ssTmp = SpreadsheetApp.create('TMP_' + nombre + '_' + new Date().getTime());
    var archivoXlsx;
    try {
      escribirReporteAnual(ssTmp, personas, fechas, matriz);
      escribirResumen(ssTmp, personas, fechas, matriz);
      escribirTabla(ssTmp, HOJAS.FALTAS, columnasFaltas(), listaFaltas(personas, fechas, matriz, datos), COLOR_FALTA);
      escribirTabla(ssTmp, HOJAS.RETARDOS, columnasRetardos(), listaRetardos(personas, fechas, matriz, datos), null);
      escribirTabla(ssTmp, HOJAS.FALLIDOS, columnasFallidos(), listaFallidos(personas, fechas, matriz, datos), COLOR_FALLIDO);
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

    if (estado.errores.length) {
      Logger.log('⚠️ Archivos que no se pudieron leer (' + estado.errores.length + '): ' + estado.errores.join(' | '));
    }
    Logger.log('🎉 Listo. Excel generado: ' + archivoXlsx.getName() + ' -> ' + archivoXlsx.getUrl());
    Logger.log('Tiempo de generación: ' + ((new Date() - t0) / 1000) + 's');

    if (CONFIG_ANUAL.ENVIAR_CORREO) {
      enviarCorreo(anio, fechas, personas, matriz, archivoXlsx, estado.errores);
    }
  }

  function esHojaNuestra(nombre) {
    return Object.keys(HOJAS).some(function (k) { return HOJAS[k] === nombre; });
  }

  // ==========================================================================
  // DIAGNÓSTICO (ligero: lista todo, pero solo abre el primer archivo de cada
  // tipo para no chocar con el límite de 6 minutos)
  // ==========================================================================

  function diagnosticar() {
    var carpeta = DriveApp.getFolderById(CONFIG_ANUAL.ID_CARPETA_ENTRADA_ANUAL);
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

    mostrar('Asistencias (1er archivo)', archivos.asistencias[0], 'Fecha Entrada');
    mostrar('Inasistencias (1er archivo)', archivos.inasistencias[0], 'Día');
    mostrar('Registro fallido (1er archivo)', archivos.fallidos[0], 'Fecha intento');
    mostrar('Maestro', archivos.maestro, null);
    Logger.log('Calendario: ' + (archivos.calendario ? archivos.calendario.getName() : 'NO ENCONTRADO (no se marcarán festivos)'));

    var props = PropertiesService.getScriptProperties();
    Logger.log('Proceso a medias: ' + (props.getProperty(PROP.ESTADO)
      ? ('sí, fase "' + props.getProperty(PROP.ESTADO) + '", iniciado ' + props.getProperty(PROP.INICIADO))
      : 'no'));
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
  // ACUMULADO DE DATOS (se llena archivo por archivo y se guarda en el JSON
  // de trabajo entre ejecuciones, por eso usa arreglos compactos).
  //
  //   asis[f][rut] = [codigo, retardoEfectivo, diferenciaMin, horaProg, horaReal,
  //                   recinto, area, supervisor, sigla]
  //   fall[f][rut] = [{detalle: true}, recinto, area, supervisor]
  //   inas[f][rut] = [codigo, recinto, area, supervisor, contrato, horario, sigla]
  //   horario[rut] = [fecha, horarioTurno]         (el más reciente)
  //   fuente[rut]  = [fecha, nombre, area, recinto, supervisor] (el más reciente)
  //
  // Cada tipo guarda su mejor fila por RUT + fecha, así que el orden en que se
  // leen los archivos no importa. Las reglas de prioridad se aplican al final,
  // en resolverDatos().
  // ==========================================================================

  function nuevoAcumulado() {
    return { asis: {}, fall: {}, inas: {}, horario: {}, fuente: {}, motivosDesconocidos: {} };
  }

  function acumularTabla(acc, tipo, tabla, toleranciaMin) {
    if (tipo === 'asistencia') tabla.rows.forEach(function (r) { acumularAsistencia(acc, r, toleranciaMin); });
    else if (tipo === 'fallido') tabla.rows.forEach(function (r) { acumularFallido(acc, r); });
    else tabla.rows.forEach(function (r) { acumularInasistencia(acc, r); });
  }

  function registrarFuente(acc, rut, f, row) {
    if (!acc.fuente[rut] || acc.fuente[rut][0] <= f) {
      acc.fuente[rut] = [f, nombreDeFila(row), row['Área'] || '', row['Recinto'] || '', row['Supervisor'] || ''];
    }
  }

  function celda(mapa, f) { return (mapa[f] = mapa[f] || {}); }

  function acumularAsistencia(acc, row, toleranciaMin) {
    var f = parseFecha(row['Fecha Entrada']);
    var rut = limpiarRut(row['RUT']);
    if (!f || !rut) return;
    registrarFuente(acc, rut, f, row);

    var horarioTurno = String(row['Horario Turno'] || '').trim();
    if (horarioTurno && horarioTurno !== '-' && (!acc.horario[rut] || acc.horario[rut][0] <= f)) {
      acc.horario[rut] = [f, horarioTurno];
    }

    var rango = parseRangoHorario(row['Horario Turno']);
    var horaReal = parseHora(row['Hora Entrada']);
    var diferencia = (rango.inicio === null || horaReal === null) ? null : horaReal - rango.inicio;
    var tarde = diferencia !== null && diferencia > toleranciaMin;
    var codigo = tarde ? 'A-' + Math.round(diferencia - toleranciaMin) : 'A';
    var retardoEfectivo = tarde ? diferencia : 0;

    // Si checó en varios recintos el mismo día, se queda la checada con
    // menos retardo.
    var previo = celda(acc.asis, f)[rut];
    if (previo && previo[1] <= retardoEfectivo) return;
    acc.asis[f][rut] = [codigo, retardoEfectivo, diferencia, rango.inicio, horaReal,
      row['Recinto'] || '', row['Área'] || '', row['Supervisor'] || '', row['Sigla Turno'] || ''];
  }

  function acumularFallido(acc, row) {
    var f = parseFecha(row['Fecha intento']);
    var rut = limpiarRut(row['RUT']);
    if (!f || !rut) return;
    registrarFuente(acc, rut, f, row);
    var c = celda(acc.fall, f);
    c[rut] = c[rut] || [{}, row['Recinto'] || '', row['Área'] || '', row['Supervisor'] || ''];
    c[rut][0][describirIntento(row)] = true; // sin duplicados si los archivos se traslapan
  }

  function acumularInasistencia(acc, row) {
    var f = parseFecha(row['Día']);
    var rut = limpiarRut(row['RUT']);
    if (!f || !rut) return;
    registrarFuente(acc, rut, f, row);

    var motivo = String(row['Motivo'] || '').trim();
    var codigo = MOTIVO_A_CODIGO.hasOwnProperty(motivo) ? MOTIVO_A_CODIGO[motivo] : null;
    if (codigo === null) {
      acc.motivosDesconocidos[motivo] = (acc.motivosDesconocidos[motivo] || 0) + 1;
      codigo = motivo || 'F';
    }

    // Varias inasistencias el mismo día (distintos recintos): cualquier
    // justificación le gana a la falta sin justificar.
    var previo = celda(acc.inas, f)[rut];
    if (previo && (previo[0] !== 'F' || codigo === 'F')) return;
    acc.inas[f][rut] = [codigo, row['Recinto'] || '', row['Área'] || '', row['Supervisor'] || '',
      row['Contrato'] || '', row['Horario'] || '', row['Sigla Turno'] || ''];
  }

  function describirIntento(row) {
    return ((row['Sentido'] || '') + ' ' + (row['Hora intento'] || '') + ' (' + (row['Error al marcar'] || '') + ')').trim();
  }

  // ==========================================================================
  // REGLAS DE PRIORIDAD -> datos[fecha][rut] = { codigo, comentario, tipo,
  //   fila, retardoMin, horaProgramada, horaReal, detalle }
  //   1. asistencia real (cualquier recinto)  2. intento fallido
  //   3. motivo de inasistencia (F solo si ninguna fila está justificada)
  // ==========================================================================

  function resolverDatos(acc, prefijoAnio) {
    var datos = {};
    var desconocidos = Object.keys(acc.motivosDesconocidos);
    if (desconocidos.length) {
      Logger.log('Aviso: motivos que no están en el catálogo (se dejan tal cual): ' +
        desconocidos.map(function (m) { return '"' + m + '" x' + acc.motivosDesconocidos[m]; }).join(', '));
    }

    function delAnio(mapa, fn) {
      Object.keys(mapa).forEach(function (f) {
        if (f.indexOf(prefijoAnio) !== 0) return;
        Object.keys(mapa[f]).forEach(function (rut) {
          datos[f] = datos[f] || {};
          if (datos[f][rut]) return; // ya lo resolvió una regla de mayor prioridad
          datos[f][rut] = fn(mapa[f][rut]);
        });
      });
    }

    delAnio(acc.asis, function (a) {
      return {
        codigo: a[0], comentario: null, tipo: 'asistencia',
        retardoMin: a[2], horaProgramada: a[3], horaReal: a[4],
        fila: { 'Recinto': a[5], 'Área': a[6], 'Supervisor': a[7], 'Sigla Turno': a[8] }
      };
    });
    delAnio(acc.fall, function (x) {
      var detalle = Object.keys(x[0]).join('; ');
      return {
        codigo: 'A', tipo: 'fallido', detalle: detalle,
        comentario: 'Asistencia inferida por intento de checada fallido: ' + detalle,
        fila: { 'Recinto': x[1], 'Área': x[2], 'Supervisor': x[3] }
      };
    });
    delAnio(acc.inas, function (i) {
      return {
        codigo: i[0], comentario: null, tipo: 'inasistencia',
        fila: { 'Recinto': i[1], 'Área': i[2], 'Supervisor': i[3], 'Contrato': i[4], 'Horario': i[5], 'Sigla Turno': i[6] }
      };
    });
    return datos;
  }

  // ==========================================================================
  // PERSONAS Y MATRIZ
  // ==========================================================================

  function construirPersonas(roster, acc, datos) {
    function horarioDe(rut) { return acc.horario[rut] ? acc.horario[rut][1] : '-'; }
    var enMaestro = {};
    var personas = roster.map(function (p) {
      enMaestro[p.RUT] = true;
      return {
        RUT: p.RUT,
        'Nombre completo': nombreDeFila(p),
        'Área': p['Área'] || '',
        'Horario Turno': horarioDe(p.RUT),
        'Recinto': p['Recinto'] || '',
        'Localidad': calcularLocalidad(p['Recinto']),
        'Supervisor': p['Supervisor'] || '',
        'En maestro': 'Sí'
      };
    });

    if (CONFIG_ANUAL.INCLUIR_FUERA_DE_MAESTRO) {
      // Solo quienes tienen algún registro en el año reportado.
      var conDatos = {};
      Object.keys(datos).forEach(function (f) {
        Object.keys(datos[f]).forEach(function (rut) { conDatos[rut] = true; });
      });
      var extra = Object.keys(conDatos).filter(function (rut) { return !enMaestro[rut] && acc.fuente[rut]; });
      extra.sort();
      extra.forEach(function (rut) {
        var src = acc.fuente[rut]; // [fecha, nombre, area, recinto, supervisor]
        personas.push({
          RUT: rut,
          'Nombre completo': src[1],
          'Área': src[2],
          'Horario Turno': horarioDe(rut),
          'Recinto': src[3],
          'Localidad': calcularLocalidad(src[3]),
          'Supervisor': src[4],
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

  function enviarCorreo(anio, fechas, personas, matriz, archivo, errores) {
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
      + (errores.length ? ('<p style="color:#a00;">Archivos que no se pudieron leer (' + errores.length + '): '
        + errores.join('<br>') + '</p>') : '')
      + '<p>Saludos cordiales.</p>'
      + '</div>';
    MailApp.sendEmail({
      to: CONFIG_ANUAL.CORREOS.join(','),
      subject: CONFIG_ANUAL.ASUNTO_CORREO + ' ' + anio,
      htmlBody: html
    });
    Logger.log('Correo enviado a ' + CONFIG_ANUAL.CORREOS.join(', '));
  }

  return { ejecutar: ejecutar, reiniciar: reiniciar, diagnosticar: diagnosticar };
})();
