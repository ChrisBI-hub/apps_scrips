// ─── PUNTO DE ENTRADA PRINCIPAL ─────────────────────────────
function iniciarReporteMetasAduana() {
  var props    = PropertiesService.getScriptProperties();
  var estado   = props.getProperty("estado");
  var listaIds = props.getProperty("listaIds");

  // ✅ Si hay estado pero no hay lista válida → datos corruptos, reiniciar
  if (estado && (!listaIds || JSON.parse(listaIds).length === 0)) {
    Logger.log("⚠️ Estado corrupto detectado. Reiniciando automáticamente...");
    _limpiarEstado();
  }

  if (!props.getProperty("estado")) {
    _inicializarEstado();
  }

  _procesarLote();
}