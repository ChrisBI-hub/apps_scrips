// ─── RECUPERAR Y FINALIZAR (ejecutar una sola vez) ───────────
function recuperarYFinalizar() {
  var props      = PropertiesService.getScriptProperties();
  var resultados = JSON.parse(props.getProperty("resultados") || "[]");
  Logger.log("🔄 Recuperando " + resultados.length + " clientes para generar reporte...");
  _generarReporteFinal(resultados, props);
}