// =========================================================================
// MENU.gs — MENÚ DE LA BARRA DE GOOGLE SHEETS
// =========================================================================
// Apps Script admite una sola función onOpen() por proyecto, así que los
// menús de proveedores y de contratos se crean aquí.
// =========================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('🚀 Control Proveedores')
    .addItem('📊 Generar Dashboard Directivo', 'generarDashboardProveedores')
    .addItem('🔍 Rastrear Respuestas (Gmail)', 'auditarRespuestasGmail')
    .addItem('🔍 Forzar Rayos X (Manual)', 'rescateRayosX')
    .addItem('📤 Forzar Envío Ahora (Manual)', 'forzarEnvioManual')
    .addItem('🔄 Sincronizar Calificaciones', 'FaltantePorCalificacion')
    .addSeparator()
    .addSubMenu(ui.createMenu('🚦 Escalamiento')
      .addItem('📞 Marcar Seguimiento Telefónico', 'marcarSeguimientoTelefonico')
      .addItem('💬 Marcar Contacto WhatsApp', 'marcarContactoWhatsApp')
      .addItem('📧 Marcar Recordatorio Enviado (Correo)', 'marcarRecordatorioCorreo')
      .addItem('✅ Marcar Resuelto', 'marcarResuelto')
      .addItem('🔁 Reabrir Caso', 'reabrirCaso')
      .addSeparator()
      .addItem('👁 Ver Bitácora', 'mostrarBitacora')
      .addItem('🙈 Ocultar Bitácora', 'ocultarBitacora'))
    .addSubMenu(ui.createMenu('⚙️ Instalación')
      .addItem('1. Crear Tabla Control y Bitácora', 'inicializarEstructuraEscalamiento')
      .addItem('2. Programar envío diario', 'instalarTriggerEnvioDiario')
      .addItem('3. Programar detección de casos sin seguimiento', 'instalarTriggerDeteccionEscalamiento')
      .addItem('Probar calendario de festivos', 'verificarCalendarioFestivos'))
    .addToUi();

  ui.createMenu('📑 Contratos')
    .addItem('🔄 Actualizar Dashboard General', 'generarDashboardContratos')
    .addItem('🚨 Forzar Envío de Alertas de Correo', 'alertarContratosAviso')
    .addToUi();
}
