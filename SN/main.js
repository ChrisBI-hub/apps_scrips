// ============================================================
//  REPORTE DE METAS ADUANALES — Procesamiento por lotes
// ============================================================

var CONFIG = {
  ID_CARPETA_PRINCIPAL : "1jvDFCzA2jdUJtSR5PrsnVkzSngtM6p4z",
  ID_REPORTE           : "18wWlUrB41-uWs0XhOPq_WcjYl5PbV4vkQVfX4qmYmlA",
  META_DOCUMENTOS      : 19,
  CORREO_NOTIFICACION  : ['ccarbajal@abcsc.mx','jsalazar@abcsc.mx','imedrano@abcsc.mx','jestrada@abcsc.mx'],
  LOTE_TAMANO          : 15,
  LIMITE_SEGUNDOS      : 300
};

// ─── PUNTO DE ENTRADA PRINCIPAL ─────────────────────────────
function iniciarReporteMetasAduana() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("estado")) {
    _inicializarEstado();
  }
  _procesarLote();
}

// ─── INICIALIZAR ESTADO ──────────────────────────────────────
function _inicializarEstado() {
  var props            = PropertiesService.getScriptProperties();
  var carpetaPrincipal = DriveApp.getFolderById(CONFIG.ID_CARPETA_PRINCIPAL);
  var iterador         = carpetaPrincipal.getFolders();
  var listaIds         = [];

  while (iterador.hasNext()) {
    listaIds.push(iterador.next().getId());
  }

  props.setProperties({
    "estado"     : "procesando",
    "indice"     : "0",
    "listaIds"   : JSON.stringify(listaIds),
    "resultados" : JSON.stringify([]),
    "iniciado"   : new Date().toISOString()
  });

  Logger.log("✅ Estado inicializado. Carpetas encontradas: " + listaIds.length);
}

// ─── PROCESAR UN LOTE ────────────────────────────────────────
function _procesarLote() {
  var props      = PropertiesService.getScriptProperties();
  var inicio     = new Date().getTime();
  var listaIds   = JSON.parse(props.getProperty("listaIds")   || "[]");
  var resultados = JSON.parse(props.getProperty("resultados") || "[]");
  var indice     = parseInt(props.getProperty("indice") || "0", 10);
  var total      = listaIds.length;

  if (!Array.isArray(resultados)) resultados = [];

  if (total === 0) {
    Logger.log("⚠️ No hay carpetas para procesar.");
    _limpiarEstado();
    return;
  }

  Logger.log("▶ Procesando desde índice " + indice + " / " + total);

  while (indice < total) {
    var transcurrido = (new Date().getTime() - inicio) / 1000;
    if (transcurrido >= CONFIG.LIMITE_SEGUNDOS) {
      Logger.log("⏱ Límite de tiempo alcanzado. Guardando estado en índice " + indice);
      props.setProperty("indice",     indice.toString());
      props.setProperty("resultados", JSON.stringify(resultados));
      _programarContinuacion();
      return;
    }

    var clienteId    = listaIds[indice];
    var carpeta      = DriveApp.getFolderById(clienteId);
    var datosCliente = {
      id              : clienteId,
      nombre          : carpeta.getName(),
      url             : carpeta.getUrl(),
      conteo          : 0,
      ultimaFecha     : 0,
      ultimoArchivoId : null,
      ultimoEditor    : "—"
    };

    recorrerCarpeta(carpeta, datosCliente);

    if (datosCliente.ultimoArchivoId) {
      try {
        var info = Drive.Files.get(datosCliente.ultimoArchivoId, {
          fields            : "lastModifyingUser/displayName,lastModifyingUser/emailAddress",
          supportsAllDrives : true
        });
        var u = info.lastModifyingUser;
        datosCliente.ultimoEditor = (u && (u.displayName || u.emailAddress)) || "—";
      } catch (e) {
        datosCliente.ultimoEditor = "Sin permiso";
      }
    }

    resultados.push(datosCliente);
    indice++;

    if (indice % CONFIG.LOTE_TAMANO === 0) {
      Logger.log("📦 Lote completado: " + indice + " / " + total);
    }
  }

  Logger.log("✅ Todos los clientes procesados (" + total + "). Generando reporte...");
  props.setProperty("resultados", JSON.stringify(resultados));
  props.setProperty("indice",     indice.toString());

  _generarReporteFinal();
}

// ─── GENERAR REPORTE FINAL ───────────────────────────────────
function _generarReporteFinal() {
  var props         = PropertiesService.getScriptProperties(); // ✅ siempre propio
  var listaClientes = JSON.parse(props.getProperty("resultados") || "[]");

  if (!Array.isArray(listaClientes) || listaClientes.length === 0) {
    Logger.log("❌ No hay datos que reportar. Abortando.");
    _limpiarEstado();
    return;
  }

  var reporte = SpreadsheetApp.openById(CONFIG.ID_REPORTE);
  var hoja    = reporte.getSheets()[0];
  hoja.clearContents();

  hoja.getRange(1, 1, 1, 7).setValues([[
    "Nombre del Cliente", "Estatus", "Docs Contados",
    "% de Avance", "Última Modificación", "Modificado Por", "Link"
  ]]);

  var filasParaInsertar   = [];
  var matrizColores       = [];
  var totalClientes       = listaClientes.length;
  var clientesCompletos   = 0;
  var sumaPorcentajes     = 0;
  var clientesConProgreso = 0;

  for (var i = 0; i < listaClientes.length; i++) {
    var cliente  = listaClientes[i];
    var completo = (cliente.conteo >= CONFIG.META_DOCUMENTOS);
    if (completo) clientesCompletos++;

    var porcentaje = Math.min(cliente.conteo / CONFIG.META_DOCUMENTOS, 1);
    if (porcentaje > 0) {
      sumaPorcentajes += porcentaje;
      clientesConProgreso++;
    }

    var fechaFormateada = "Sin archivos";
    if (cliente.ultimaFecha > 0) {
      fechaFormateada = Utilities.formatDate(
        new Date(cliente.ultimaFecha),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm"
      );
    }

    filasParaInsertar.push([
      cliente.nombre,
      completo ? "COMPLETO ✓" : "INCOMPLETO ✗",
      cliente.conteo,
      porcentaje,
      fechaFormateada,
      cliente.ultimoEditor,
      cliente.url
    ]);

    matrizColores.push(Array(7).fill(completo ? "#d9ead3" : "#f4cccc"));
  }

  var rangoDatos = hoja.getRange(2, 1, filasParaInsertar.length, 7);
  rangoDatos.setValues(filasParaInsertar);
  rangoDatos.setBackgrounds(matrizColores);
  hoja.getRange("D2:D" + hoja.getLastRow()).setNumberFormat("0%");

  var filaResumen = hoja.getLastRow() + 2;
  var efectividad = clientesCompletos / totalClientes;
  var promedio    = clientesConProgreso > 0 ? sumaPorcentajes / clientesConProgreso : 0;

  hoja.getRange(filaResumen, 1, 3, 7).setValues([
    ["", "", "", "", "", "", ""],
    ["RESUMEN DE LA AGENCIA", "Clientes Completos", "Total Clientes",
     "EFECTIVIDAD TOTAL", "PROMEDIO AVANCE (Sin 0%)", "", ""],
    ["", clientesCompletos, totalClientes, efectividad, promedio, "", ""]
  ]);

  hoja.getRange(filaResumen + 1, 1, 1, 5)
      .setBackground("#444444").setFontColor("white").setFontWeight("bold");
  hoja.getRange(filaResumen + 2, 4, 1, 2)
      .setNumberFormat("0%").setFontWeight("bold").setBackground("#cfe2f3");

  hoja.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#f3f3f3");
  hoja.autoResizeColumns(1, 7);

  enviarCorreoNotificacion(totalClientes, clientesCompletos, efectividad, promedio);
  _cancelarTriggersContinuacion();
  _limpiarEstado();

  Logger.log("🎉 REPORTE COMPLETADO: " + totalClientes + " clientes. Efectividad: " +
             (efectividad * 100).toFixed(1) + "%");
}

// ─── TRIGGER ENCADENADO ──────────────────────────────────────
function _programarContinuacion() {
  _cancelarTriggersContinuacion();
  ScriptApp.newTrigger("iniciarReporteMetasAduana")
    .timeBased()
    .after(60 * 1000)
    .create();
  Logger.log("⏰ Trigger de continuación programado en 1 minuto.");
}

function _cancelarTriggersContinuacion() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    if (t.getHandlerFunction() === "iniciarReporteMetasAduana" &&
        t.getEventType()       === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(t);
    }
  }
}

// ─── LIMPIAR ESTADO ──────────────────────────────────────────
function _limpiarEstado() {
  var props = PropertiesService.getScriptProperties(); // ✅ siempre propio
  props.deleteProperty("estado");
  props.deleteProperty("indice");
  props.deleteProperty("listaIds");
  props.deleteProperty("resultados");
  props.deleteProperty("iniciado");
  Logger.log("🧹 Estado de PropertiesService limpiado.");
}

// ─── CORREO DE NOTIFICACIÓN ──────────────────────────────────
function enviarCorreoNotificacion(totalClientes, clientesCompletos, efectividad, promedio) {
  var ahora          = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  var efectividadPct = (efectividad * 100).toFixed(1) + "%";
  var promedioPct    = (promedio    * 100).toFixed(1) + "%";
  var urlReporte     = "https://docs.google.com/spreadsheets/d/" + CONFIG.ID_REPORTE + "/edit";
  var asunto         = "📋 Reporte de Documentación Aduanal — " + ahora;

  var cuerpoHtml =
    "<div style='font-family:Arial,sans-serif;max-width:620px;color:#333;'>" +
    "<div style='background:#2c5f8a;padding:20px 24px;border-radius:6px 6px 0 0;'>" +
      "<h2 style='margin:0;color:#fff;font-size:20px;'>📋 Reporte de Documentación Aduanal</h2>" +
      "<p style='margin:6px 0 0;color:#cce0f5;font-size:13px;'>Generado el " + ahora + "</p>" +
    "</div>" +
    "<div style='background:#f9f9f9;padding:24px;border:1px solid #ddd;border-top:none;border-radius:0 0 6px 6px;'>" +
      "<p style='margin-top:0;'>El reporte de control de documentación ha sido generado. Resumen ejecutivo:</p>" +
      "<table style='border-collapse:collapse;width:100%;margin:16px 0;font-size:14px;'>" +
        "<thead><tr style='background:#444;color:#fff;'>" +
          "<th style='padding:10px 16px;text-align:left;'>Indicador</th>" +
          "<th style='padding:10px 16px;text-align:center;'>Valor</th>" +
        "</tr></thead>" +
        "<tbody>" +
          "<tr style='background:#fff;'><td style='padding:10px 16px;border-bottom:1px solid #e0e0e0;'>Total de clientes</td>" +
            "<td style='padding:10px 16px;text-align:center;border-bottom:1px solid #e0e0e0;'><strong>" + totalClientes + "</strong></td></tr>" +
          "<tr style='background:#f2f2f2;'><td style='padding:10px 16px;border-bottom:1px solid #e0e0e0;'>Clientes completos</td>" +
            "<td style='padding:10px 16px;text-align:center;border-bottom:1px solid #e0e0e0;'><strong style='color:#2e7d32;'>" + clientesCompletos + "</strong></td></tr>" +
          "<tr style='background:#fff;'><td style='padding:10px 16px;border-bottom:1px solid #e0e0e0;'>Efectividad total</td>" +
            "<td style='padding:10px 16px;text-align:center;border-bottom:1px solid #e0e0e0;'><strong style='color:#1565c0;font-size:16px;'>" + efectividadPct + "</strong></td></tr>" +
          "<tr style='background:#f2f2f2;'><td style='padding:10px 16px;'>Promedio de avance (sin 0%)</td>" +
            "<td style='padding:10px 16px;text-align:center;'><strong style='color:#6a1e8c;font-size:16px;'>" + promedioPct + "</strong></td></tr>" +
        "</tbody>" +
      "</table>" +
      "<div style='text-align:center;margin-top:24px;'>" +
        "<a href='" + urlReporte + "' style='background:#2c5f8a;color:#fff;padding:12px 28px;text-decoration:none;border-radius:4px;font-weight:bold;font-size:14px;display:inline-block;'>Ver reporte completo →</a>" +
      "</div>" +
      "<hr style='margin:28px 0 16px;border:none;border-top:1px solid #ddd;'/>" +
      "<p style='font-size:11px;color:#999;margin:0;'>Correo generado automáticamente. No es necesario responder.</p>" +
    "</div></div>";

  try {
    MailApp.sendEmail({ to: CONFIG.CORREO_NOTIFICACION, subject: asunto, htmlBody: cuerpoHtml });
    Logger.log("✅ Correo enviado a: " + CONFIG.CORREO_NOTIFICACION);
  } catch (e) {
    Logger.log("❌ Error al enviar correo: " + e.toString());
  }
}

// ─── UTILIDADES ─────────────────────────────────────────────
function recorrerCarpeta(carpeta, cliente) {
  try {
    var archivos = carpeta.getFiles();
    while (archivos.hasNext()) {
      var archivo = archivos.next();
      cliente.conteo++;
      var fecha = archivo.getLastUpdated().getTime();
      if (fecha > cliente.ultimaFecha) {
        cliente.ultimaFecha     = fecha;
        cliente.ultimoArchivoId = archivo.getId();
      }
    }
    var subcarpetas = carpeta.getFolders();
    while (subcarpetas.hasNext()) {
      recorrerCarpeta(subcarpetas.next(), cliente);
    }
  } catch (e) {
    Logger.log("⚠️ Error en carpeta '" + carpeta.getName() + "': " + e.toString());
  }
}

// ─── FUNCIÓN DE EMERGENCIA ───────────────────────────────────
function reiniciarProcesoLimpio() {
  _limpiarEstado();
  _cancelarTriggersContinuacion();
  Logger.log("🔄 Estado limpiado. Puedes volver a ejecutar iniciarReporteMetasAduana()");
}