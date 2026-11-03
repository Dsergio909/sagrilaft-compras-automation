// =========================================================================
// MÓDULO 1: MOTOR PRINCIPAL Y CAPTURA INMEDIATA DE ENLACES (PROVEEDORES)
// =========================================================================
// UBICACIÓN: Bound al Google Sheet de Proveedores.
// PERMISOS: Sheets, Gmail (enviar correo), Drive (leer PDF/GIF adjuntos).
// TRIGGERS: ejecucionAutomaticaDias() se conecta a un trigger de tiempo
// (diario). Decide internamente si hoy es martes o viernes antes de
// llamar a la función que realmente envía correos.
//
// Usa agregarNotaHistorial() (Módulo 5, más abajo en este archivo) para
// que la columna M ("Notas del Robot") acumule un historial completo.
// =========================================================================

function ejecucionAutomaticaDias() {
  var hoy = new Date();
  var diaSemana = hoy.getDay();

  if (diaSemana === 2 || diaSemana === 5) {
    enviarAvisosSagrilaftProduccion();
  } else {
    Logger.log("Hoy no es Martes ni Viernes. El sistema automático descansa.");
  }
}

// Envío manual desde el menú, con confirmación.
function forzarEnvioManual() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert("Enviar solicitudes ahora",
    "Se enviarán las solicitudes pendientes de hoy, aunque no sea martes ni viernes. ¿Continuar?",
    ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) enviarAvisosSagrilaftProduccion();
}

// Crea el disparador diario que llama a ejecucionAutomaticaDias().
function instalarTriggerEnvioDiario() {
  var existentes = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existentes.length; i++) {
    if (existentes[i].getHandlerFunction() === "ejecucionAutomaticaDias") {
      SpreadsheetApp.getUi().alert("Ya existe un trigger para esta función. No se creó uno nuevo.");
      return;
    }
  }
  ScriptApp.newTrigger("ejecucionAutomaticaDias").timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert("✅ Envío diario programado (7:00 a. m.; solo envía martes y viernes).");
}

function esFinDeLista(nombreMayus) {
  return (CONFIG.MARCADORES_FIN_DE_LISTA || []).some(function (m) {
    return m && nombreMayus.includes(String(m).toUpperCase());
  });
}

function copiaPorSede(clasificacion) {
  var mapa = CONFIG.CC_POR_SEDE || {};
  for (var etiqueta in mapa) {
    if (clasificacion.includes(etiqueta.toUpperCase())) return "," + mapa[etiqueta];
  }
  return "";
}

function enviarAvisosSagrilaftProduccion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var zonaHoraria = ss.getSpreadsheetTimeZone();
  var hoy = new Date();

  var hojaDatos = ss.getSheetByName("BASE DE PROVEEDORES COMPRAS");
  var hojaControl = ss.getSheetByName("Tabla Control");

  if (!hojaDatos) return;
  var ultimaFilaMatriz = hojaDatos.getLastRow();
  if (ultimaFilaMatriz < 6) return;

  var datos = hojaDatos.getRange(6, 1, ultimaFilaMatriz - 5, hojaDatos.getLastColumn()).getValues();
  var datosDisplay = hojaDatos.getRange(6, 1, ultimaFilaMatriz - 5, hojaDatos.getLastColumn()).getDisplayValues();
  var datosControl = hojaControl.getRange(2, 1, Math.max(1, hojaControl.getLastRow() - 1), 13).getDisplayValues();

  var diasAviso = 60;

  // ⚠️ VERIFICAR PERIÓDICAMENTE: este ID apunta al PDF que se envía a los
  // proveedores. Confirma que corresponde al formato oficial vigente cada
  // vez que compliance actualice el documento.
  var archivoPdf;
  try { archivoPdf = DriveApp.getFileById(CONFIG.ID_PDF_FORMATO_PROVEEDOR); } catch (e) { archivoPdf = null; }
  var archivoFirma = DriveApp.getFileById(CONFIG.ID_GIF_FIRMA_PROVEEDORES);

  var enviados = [];
  var alertadosSinCorreo = [];

  for (var c = 0; c < datosControl.length; c++) {
    if (datosControl[c][2] !== "") {
      var nitC = String(datosControl[c][1]).replace(/[^0-9]/g, "");
      var fechaC = String(datosControl[c][4]).trim();
      var estadoControl = String(datosControl[c][5]).trim();

      var llave = nitC + "_" + fechaC;

      if (estadoControl.includes("Falta Correo") || estadoControl.includes("Error") || estadoControl.includes("Datos")) {
        alertadosSinCorreo.push(llave);
      } else {
        enviados.push(llave);
      }
    }
  }

  for (var i = 0; i < datos.length; i++) {
    var proveedor = String(datos[i][2]).trim();
    if (proveedor === "") continue;

    var provMayus = proveedor.toUpperCase();
    var compradorRaw = String(datos[i][1]).trim().toUpperCase();
    var nitSucio = String(datosDisplay[i][3]).trim();
    var nitLimpio = nitSucio.replace(/[^0-9]/g, "");
    var clasificacion = String(datos[i][15]).toUpperCase();
    var estadoComprador = String(datosDisplay[i][1]).toUpperCase();

    // 🎯 ESCUDO CRÍTICO 1: fin de la lista de activos
    // Desde el marcador (CONFIG.MARCADORES_FIN_DE_LISTA) empiezan los
    // proveedores inactivos o inhabilitados: no deben recibir solicitudes.
    if (esFinDeLista(provMayus)) {
      Logger.log("Fin de la lista de activos en: " + proveedor);
      break;
    }

    // 🎯 ESCUDO CRÍTICO 2
    if (compradorRaw.includes("INACTIVO") || compradorRaw.includes("INHABILITADO") || estadoComprador.includes("INACTIVO") || estadoComprador.includes("INHABILITADO")) {
      continue;
    }

    var fechaSagrilaft = datos[i][7];
    var fechaVisual = String(datosDisplay[i][7]).trim();

    var contenidoCeldaCorreo = String(datos[i][18]);
    var regexCorreo = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
    var correosEncontrados = contenidoCeldaCorreo.match(regexCorreo);
    var correoDestinoFinal = correosEncontrados ? correosEncontrados.join(",") : "";

    if (Object.prototype.toString.call(fechaSagrilaft) === "[object Date]" && !isNaN(fechaSagrilaft)) {
      var llaveFiltro = nitLimpio + "_" + fechaVisual;

      if (enviados.includes(llaveFiltro)) continue;

      var diasRestantes = Math.ceil((fechaSagrilaft.getTime() - hoy.getTime()) / (1000 * 3600 * 24));

      if (diasRestantes <= diasAviso) {
        var fechaActual = Utilities.formatDate(hoy, zonaHoraria, "dd/MM/yyyy HH:mm");
        var fechaSoloTexto = Utilities.formatDate(hoy, zonaHoraria, "dd/MM/yyyy");
        var fechaSeguimiento = Utilities.formatDate(new Date(hoy.getTime() + (8 * 24 * 60 * 60 * 1000)), zonaHoraria, "dd/MM/yyyy");

        if (!archivoPdf || correoDestinoFinal === "") {
          if (!alertadosSinCorreo.includes(llaveFiltro) || !archivoPdf) {
            var errorMsg = !archivoPdf ? "❌ Error: PDF no encontrado" : "⚠️ Alerta: Falta Correo";
            hojaControl.appendRow([fechaActual, nitSucio, proveedor, clasificacion, fechaVisual, errorMsg, "", "", "", "0", "", "", ""]);
            var filaSinCorreo = hojaControl.getLastRow();
            agregarNotaHistorial(hojaControl, filaSinCorreo, fechaSoloTexto + " - Requiere revisión manual", null);
            if (correoDestinoFinal === "") alertadosSinCorreo.push(llaveFiltro);
          }
          continue;
        }

        var asunto = `Solicitud de documentos para creación o actualización de proveedor – ${proveedor}`;
        var cuerpoTextoPlano = "Su cliente de correo no soporta formato HTML.";
        var mensajeHtml = obtenerPlantillaHtml();

        var ccFinal = CONFIG.CORREOS_EQUIPO_COMPRAS;
        if (clasificacion.includes("ARRENDAMIENTO") || clasificacion.includes("ADMINISTRACION") || clasificacion.includes("ADMINISTRACIÓN")) {
          ccFinal += copiaPorSede(clasificacion);
        }

        GmailApp.sendEmail(correoDestinoFinal, asunto, cuerpoTextoPlano, {
          htmlBody: mensajeHtml,
          attachments: [archivoPdf],
          cc: ccFinal,
          inlineImages: { firmaAnimada: archivoFirma.getBlob() }
        });

        Utilities.sleep(3000);
        var hilosRecientes = GmailApp.search('to:' + correoDestinoFinal + ' subject:"' + asunto + '"', 0, 1);
        var idHiloRobot = "";
        var enlaceRobotUrl = "";
        var notaTextoInicial = fechaSoloTexto + " - No indexado a tiempo (Usar Rayos X)";

        if (hilosRecientes.length > 0) {
          idHiloRobot = hilosRecientes[0].getId();
          enlaceRobotUrl = "https://mail.google.com/mail/u/0/#all/" + idHiloRobot;
          notaTextoInicial = fechaSoloTexto + " - Captura Inmediata";
        }

        hojaControl.appendRow([fechaActual, nitSucio, proveedor, clasificacion, fechaVisual, "Aviso Enviado Automático", fechaSeguimiento, correoDestinoFinal, "", "1", enlaceRobotUrl, idHiloRobot, ""]);
        var filaInsertada = hojaControl.getLastRow();

        if (enlaceRobotUrl !== "") {
          var enlaceLimpio = SpreadsheetApp.newRichTextValue().setText("Abrir Correo").setLinkUrl(enlaceRobotUrl).build();
          hojaControl.getRange(filaInsertada, 11).setRichTextValue(enlaceLimpio);
        }

        agregarNotaHistorial(hojaControl, filaInsertada, notaTextoInicial, enlaceRobotUrl !== "" ? enlaceRobotUrl : null);

        enviados.push(llaveFiltro);
      }
    }
  }
}

// =========================================================================
// MÓDULO 2: PLANTILLA VISUAL DEL CORREO
// =========================================================================
// La lista de documentos sale de CONFIG.DOCUMENTOS: para cambiar lo que
// se pide no hay que tocar el HTML. Revisa esa lista cada vez que cambie
// el formulario oficial de registro de proveedores.
// =========================================================================

function obtenerPlantillaHtml() {
  var anioActual = new Date().getFullYear();
  var anioCerrado = anioActual - 1;  // último año fiscal cerrado
  var lista = function (items) {
    return items.map(function (d) { return "<li>" + d.replace("{anioCerrado}", anioCerrado) + "</li>"; }).join("");
  };

  return `
    <div style="font-family: Arial, sans-serif; font-size: 14px; color: #333; line-height: 1.6;">
      <p><strong>Estimados señores.</strong></p>
      <p>Reciban un cordial saludo de parte del Departamento de Compras de <strong>${CONFIG.NOMBRE_EMPRESA}</strong>.</p>
      <p>Con el objetivo de formalizar su registro en nuestro sistema y dar cumplimiento estricto a la política de <strong>SAGRILAFT</strong>, solicitamos el envío de la siguiente documentación en formato <strong>PDF</strong>.</p>
      <p>Por favor, identifique los requisitos obligatorios según su naturaleza jurídica:</p>

      <div style="background-color: #f9f9f9; padding: 15px; border-left: 4px solid #2e7d32; margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #2e7d32;">I. SI USTED ES PERSONA JURÍDICA (EMPRESA):</h4>
        <ul style="margin-bottom: 0; padding-left: 20px;">${lista(CONFIG.DOCUMENTOS.juridica)}</ul>
      </div>

      <div style="background-color: #f9f9f9; padding: 15px; border-left: 4px solid #1976d2; margin-bottom: 20px;">
        <h4 style="margin-top: 0; color: #1976d2;">II. SI USTED ES PERSONA NATURAL:</h4>
        <ul style="margin-bottom: 0; padding-left: 20px;">${lista(CONFIG.DOCUMENTOS.natural)}</ul>
      </div>

      <div style="border: 1px solid #ffc107; background-color: #fffde7; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
        <p style="margin-top: 0;"><strong>⚠️ NOTA IMPORTANTE:</strong> Agradecemos el envío de la carpeta completa en un solo mensaje para agilizar su habilitación.</p>
        <p style="margin-bottom: 0;">Si por algún motivo de confidencialidad no es posible adjuntar algún documento, es indispensable <strong>informarlo en el cuerpo del correo y adjuntar una carta formal firmada</strong> explicando las razones.</p>
      </div>

      <p style="margin-top: 20px;">Quedamos atentos a su pronta respuesta.</p>
      <p><strong>Cordialmente,</strong></p>
    </div>

    <div style="font-family: Arial, sans-serif; margin-top: 25px;">
      <p style="color: #2e7d32; font-style: italic; margin: 0 0 10px 0; font-size: 14px;">${CONFIG.FIRMA.slogan}</p>
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td valign="middle" style="padding-right: 20px; width: 400px;">
            <img src="cid:firmaAnimada" width="400" style="width: 400px; display: block;" alt="Firma">
          </td>
          <td valign="middle">
            <div style="color: #d32f2f; font-weight: bold; font-size: 16px; margin-bottom: 2px;">${CONFIG.FIRMA.nombre}</div>
            <div style="color: #455a64; font-size: 14px; margin-bottom: 2px;">${CONFIG.FIRMA.cargo}</div>
            <div style="color: #1976d2; font-size: 13px; margin-bottom: 5px;">${CONFIG.FIRMA.correo}</div>
            <div style="color: #000000; font-size: 13px;">${CONFIG.FIRMA.direccion}</div>
            <div style="color: #000000; font-size: 13px;">${CONFIG.FIRMA.telefono}</div>
          </td>
        </tr>
      </table>
    </div>
  `;
}

// =========================================================================
// MÓDULO 3: RAYOS X (BOTÓN DE EMERGENCIA MANUAL)
// =========================================================================
// UBICACIÓN: Bound al mismo Sheet. Ejecución SOLO manual desde el menú.
// PERMISOS: Sheets, Gmail (solo lectura de hilos vía búsqueda).
// LÍMITE CONOCIDO, NO RESUELTO AÚN: la búsqueda trae máximo 500 hilos.
// =========================================================================

function rescateRayosX() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaControl = ss.getSheetByName("Tabla Control");
  var datos = hojaControl.getRange(2, 1, hojaControl.getLastRow() - 1, 13).getValues();
  var hilosEnviados = GmailApp.search('subject:"Solicitud de documentos para creación"', 0, 500);
  var fechaHoyTexto = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");

  var contadorReparados = 0;
  var contadorNuevos = 0;

  var indiceHilos = [];
  for (var j = 0; j < hilosEnviados.length; j++) {
    var mensajes = hilosEnviados[j].getMessages();
    if (mensajes.length === 0) continue;

    var primerMensaje = mensajes[0];
    indiceHilos.push({
      id: hilosEnviados[j].getId(),
      asuntoLimpio: primerMensaje.getSubject().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, ""),
      destinatarios: primerMensaje.getTo().toLowerCase()
    });
  }

  for (var i = 0; i < datos.length; i++) {
    var proveedorOriginal = String(datos[i][2]).trim();
    var correoDestinoExcel = String(datos[i][7]).trim().toLowerCase();
    var idActual = String(datos[i][11]).trim();
    var enlaceActual = String(datos[i][10]).trim();
    var filaReal = i + 2;

    if (proveedorOriginal === "") continue;

    if (idActual !== "" && (enlaceActual === "" || enlaceActual.toUpperCase().includes("ERROR"))) {
      var urlGmailReparado = "https://mail.google.com/mail/u/0/#all/" + idActual;
      var enlaceLimpioReparado = SpreadsheetApp.newRichTextValue().setText("Abrir Correo").setLinkUrl(urlGmailReparado).build();
      hojaControl.getRange(filaReal, 11).setRichTextValue(enlaceLimpioReparado);
      agregarNotaHistorial(hojaControl, filaReal, fechaHoyTexto + " - Enlace reparado (Rayos X)", urlGmailReparado);
      contadorReparados++;
      continue;
    }

    if (idActual === "") {
      var provLimpio = proveedorOriginal.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
      var primerCorreoExcel = correoDestinoExcel.split(",")[0].trim();

      var nombreConfiable = provLimpio.length > 4;

      for (var k = 0; k < indiceHilos.length; k++) {
        var hilo = indiceHilos[k];

        var matchPorAsunto = nombreConfiable && hilo.asuntoLimpio.includes(provLimpio);
        var matchPorCorreo = (primerCorreoExcel !== "") && hilo.destinatarios.includes(primerCorreoExcel);

        if (matchPorAsunto || matchPorCorreo) {
          var urlGmailNuevo = "https://mail.google.com/mail/u/0/#all/" + hilo.id;
          var enlaceLimpioNuevo = SpreadsheetApp.newRichTextValue().setText("Abrir Correo").setLinkUrl(urlGmailNuevo).build();

          hojaControl.getRange(filaReal, 11).setRichTextValue(enlaceLimpioNuevo);
          hojaControl.getRange(filaReal, 12).setValue(hilo.id);
          agregarNotaHistorial(
            hojaControl, filaReal,
            fechaHoyTexto + " - Rescatado (Rayos X) - match por " + (matchPorAsunto ? "asunto" : "correo"),
            urlGmailNuevo
          );
          contadorNuevos++;
          break;
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert("📊 RAYOS X EJECUTADO\n\nEnlaces reparados instantáneos: " + contadorReparados + "\nNuevos correos vinculados: " + contadorNuevos);
}

// =========================================================================
// MÓDULO 4: DASHBOARD CORPORATIVO INTELIGENTE
// =========================================================================
// UBICACIÓN: Bound al mismo Sheet. Ejecución manual desde el menú.
// PERMISOS: Solo Sheets. No usa Gmail ni Drive.
//
// CAMBIO CRÍTICO: antes, 9 columnas se referenciaban por número — si
// alguien insertaba una columna nueva en "BASE DE PROVEEDORES COMPRAS",
// el dashboard seguía corriendo SIN error, pero leyendo datos de la
// columna equivocada, silenciosamente. Ahora se buscan por el NOMBRE
// real del encabezado (fila 5). Si un encabezado obligatorio no se
// encuentra, la función se detiene con un error explícito.
//
// DECISIÓN DE DISEÑO: el índice de cumplimiento se calcula solo con la
// fecha de vencimiento ("Prox Sagrilaf"), no con una columna de estado
// que se llena a mano: esa columna se desactualiza y daría falsos "al día".
// =========================================================================

function obtenerIndicesBaseProveedores(hojaDatos) {
  var headers = hojaDatos.getRange(5, 1, 1, hojaDatos.getLastColumn()).getValues()[0]
    .map(function (h) { return normalizarTextoDashboard(h); });

  function buscar(nombreBuscado, obligatorio) {
    var idx = headers.indexOf(normalizarTextoDashboard(nombreBuscado));
    if (idx === -1 && obligatorio) {
      throw new Error("No se encontró la columna obligatoria '" + nombreBuscado + "' en la fila 5 de 'BASE DE PROVEEDORES COMPRAS'. Revisa si el encabezado cambió de nombre o de posición.");
    }
    if (idx === -1) {
      Logger.log("⚠️ Columna opcional '" + nombreBuscado + "' no encontrada. Esa métrica se omitirá del dashboard.");
    }
    return idx;
  }

  return {
    comprador: buscar("comprador", true),
    nombreProveedor: buscar("nombre proveedor", true),
    proxSagrilaf: buscar("prox sagrilaf", true),
    resultadoEvaluacion: buscar("resultado utima evaluacion", false),
    calidad: buscar("calidad", false),
    sst: buscar("sst", false),
    ambiental: buscar("ambiental", false),
    clasificacion: buscar("clasificacion", true)
  };
}

function normalizarTextoDashboard(texto) {
  return String(texto).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function generarDashboardProveedores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDatos = ss.getSheetByName("BASE DE PROVEEDORES COMPRAS");

  if (!hojaDatos) {
    try { SpreadsheetApp.getUi().alert("Error: No se encuentra 'BASE DE PROVEEDORES COMPRAS'."); } catch(e){}
    return;
  }

  var ultimaFilaMatriz = hojaDatos.getLastRow();
  if (ultimaFilaMatriz < 6) {
    try { SpreadsheetApp.getUi().alert("⚠️ La pestaña 'BASE DE PROVEEDORES COMPRAS' no contiene datos suficientes."); } catch(e){}
    return;
  }

  var idx;
  try {
    idx = obtenerIndicesBaseProveedores(hojaDatos);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ No se pudo generar el dashboard:\n\n" + e.message);
    return;
  }

  var idHojaDatos = hojaDatos.getSheetId();
  var hojaDash = ss.getSheetByName("DASHBOARD SAGRILAFT");

  if (!hojaDash) {
    hojaDash = ss.insertSheet("DASHBOARD SAGRILAFT", 0);
  } else {
    hojaDash.clear();
  }

  var datosProv = hojaDatos.getRange(6, 1, ultimaFilaMatriz - 5, hojaDatos.getLastColumn()).getValues();
  var hoy = new Date();
  hoy.setHours(0,0,0,0);

  var metricas = {
    totalFilas: 0, activos: 0, inactivos: 0, inhabilitados: 0,
    arrendamientosFiducias: 0, administraciones: 0, serviciosGenerales: 0,
    vigenSagrilaft: 0, porVenceSagrilaft: 0, vencidoSagrilaft: 0,
    criticosCalidad: 0, criticosSST: 0, criticosAmbiental: 0
  };

  var rendimientoCompradores = {};
  var planDeAccionPrioritario = [];

  for (var i = 0; i < datosProv.length; i++) {
    var fila = datosProv[i];
    var filaRealExcel = i + 6;

    var nombreProv = String(fila[idx.nombreProveedor]).trim();
    if (nombreProv === "") continue;

    metricas.totalFilas++;
    var compradorRaw = String(fila[idx.comprador]).trim().toUpperCase();

    if (compradorRaw.includes("INACTIVO")) {
      metricas.inactivos++;
      continue;
    } else if (compradorRaw.includes("INHABILITADO")) {
      metricas.inhabilitados++;
      continue;
    }

    metricas.activos++;
    var comprador = compradorRaw || "SIN ASIGNAR";
    var clasificacionRaw = String(fila[idx.clasificacion]).trim().toUpperCase();
    var evaluacion = (idx.resultadoEvaluacion !== -1 ? String(fila[idx.resultadoEvaluacion]).trim() : "") || "N/A";

    if (clasificacionRaw.includes("ARRENDAMIENTO") || nombreProv.includes("FIDUCIARIA") || clasificacionRaw.includes("FIDUCIA")) {
      metricas.arrendamientosFiducias++;
    } else if (clasificacionRaw.includes("ADMINISTRACIÓN") || clasificacionRaw.includes("ADMINISTRACION")) {
      metricas.administraciones++;
    } else {
      metricas.serviciosGenerales++;
    }

    var cal = idx.calidad !== -1 ? String(fila[idx.calidad]).toUpperCase() : "";
    var sst = idx.sst !== -1 ? String(fila[idx.sst]).toUpperCase() : "";
    var amb = idx.ambiental !== -1 ? String(fila[idx.ambiental]).toUpperCase() : "";

    if (cal.includes("CRITICO") && !cal.includes("NO")) metricas.criticosCalidad++;
    if (sst.includes("CRITICO") && !sst.includes("NO")) metricas.criticosSST++;
    if (amb.includes("CRITICO") && !amb.includes("NO")) metricas.criticosAmbiental++;

    var fechaSag = fila[idx.proxSagrilaf];
    var estadoPlazo = "VIGENTE";
    var diasRestantes = 0;
    var nivelGravedad = 0;

    if (!(fechaSag instanceof Date) || isNaN(fechaSag)) {
      estadoPlazo = "SIN FECHA REGISTRADA";
      nivelGravedad = 2;
    } else {
      diasRestantes = Math.ceil((fechaSag.getTime() - hoy.getTime()) / (1000 * 3600 * 24));
      if (diasRestantes < 0) {
        estadoPlazo = "VENCIDO (" + Math.abs(diasRestantes) + " días)";
        nivelGravedad = 2;
      } else if (diasRestantes <= 45) {
        estadoPlazo = "POR VENCER (" + diasRestantes + " días)";
        nivelGravedad = 1;
      }
    }

    if (nivelGravedad === 2) metricas.vencidoSagrilaft++;
    else if (nivelGravedad === 1) metricas.porVenceSagrilaft++;
    else metricas.vigenSagrilaft++;

    if (!rendimientoCompradores[comprador]) {
      rendimientoCompradores[comprador] = { activos: 0, alertas: 0 };
    }
    rendimientoCompradores[comprador].activos++;

    if (nivelGravedad === 2) {
      rendimientoCompradores[comprador].alertas++;
      planDeAccionPrioritario.push({ name: nombreProv, gestor: comprador, eval: evaluacion, obs: estadoPlazo, gravedad: 2, rCal: cal, rSst: sst, rAmb: amb, fila: filaRealExcel });
    } else if (nivelGravedad === 1) {
      planDeAccionPrioritario.push({ name: nombreProv, gestor: comprador, eval: evaluacion, obs: estadoPlazo, gravedad: 1, rCal: cal, rSst: sst, rAmb: amb, fila: filaRealExcel });
    }
  }

  var scoreCumplimiento = metricas.activos > 0 ? Math.round(((metricas.activos - metricas.vencidoSagrilaft) / metricas.activos) * 100) : 0;
  planDeAccionPrioritario.sort(function(a, b) { return b.gravedad - a.gravedad; });

  var listaCompradores = Object.keys(rendimientoCompradores).map(function(key) {
    return { name: key, datos: rendimientoCompradores[key] };
  }).sort(function(a, b) { return b.datos.activos - a.datos.activos; });

  hojaDash.setHiddenGridlines(true);
  var cSlate = "#2c3e50", cGrisSuave = "#f8f9fa", cBorde = "#e2e8f0", cRojo = "#c0392b", cNaranja = "#d35400", cVerde = "#27ae60";

  hojaDash.getRange("B2:M2").merge().setBackground(cSlate).setValue("  📊 TORRE DE CONTROL DE PROVEEDORES: EXPEDIENTES Y CUMPLIMIENTO SAGRILAFT").setFontSize(13).setFontColor("white").setFontWeight("bold").setVerticalAlignment("middle");
  hojaDash.getRange("B3").setValue("Último escaneo del sistema: " + Utilities.formatDate(hoy, "GMT-5", "dd/MM/yyyy HH:mm")).setFontSize(9).setFontColor("#7f8c8d").setFontStyle("italic");

  hojaDash.getRange("B5:C5").merge().setValue("PROVEEDORES ACTIVOS").setFontSize(8).setFontColor("#7f8c8d").setFontWeight("bold").setHorizontalAlignment("center").setBackground(cGrisSuave);
  hojaDash.getRange("B6:C7").merge().setValue(metricas.activos).setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(cSlate).setBorder(true, true, true, true, false, false, cBorde, SpreadsheetApp.BorderStyle.SOLID);

  hojaDash.getRange("E5:F5").merge().setValue("ÍNDICE DE COMPLIANCE").setFontSize(8).setFontColor("#7f8c8d").setFontWeight("bold").setHorizontalAlignment("center").setBackground(cGrisSuave);
  hojaDash.getRange("E6:F7").merge().setValue(scoreCumplimiento + "%").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(scoreCumplimiento > 85 ? cVerde : cRojo).setBorder(true, true, true, true, false, false, cBorde, SpreadsheetApp.BorderStyle.SOLID);

  hojaDash.getRange("H5:I5").merge().setValue("🚨 EXPEDIENTES EN RIESGO").setFontSize(8).setFontColor(cRojo).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#fdedec");
  hojaDash.getRange("H6:I7").merge().setValue(metricas.vencidoSagrilaft).setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(cRojo).setBorder(true, true, true, true, false, false, "#f5b7b1", SpreadsheetApp.BorderStyle.SOLID);

  hojaDash.getRange("K5:L5").merge().setValue("ARCHIVO HISTÓRICO").setFontSize(8).setFontColor("#7f8c8d").setFontWeight("bold").setHorizontalAlignment("center").setBackground(cGrisSuave);
  hojaDash.getRange("K6:L7").merge().setValue("Inactivos: " + metricas.inactivos + " | Bloqueados: " + metricas.inhabilitados).setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor("#7f8c8d").setBorder(true, true, true, true, false, false, cBorde, SpreadsheetApp.BorderStyle.SOLID);

  hojaDash.getRange("B10:L10").merge().setBackground("#34495e").setValue("  DIAGNÓSTICO OPERATIVO DE LA BASE ACTIVA").setFontSize(10).setFontColor("white").setFontWeight("bold");
  hojaDash.getRange("B12:E12").setValues([["Arrendamientos/Fiducias", "Administraciones", "Proveedores Comunes", "Alertas Sistema (Cal/SST/Amb)"]]).setBackground(cGrisSuave).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
  hojaDash.getRange("B13:E13").setValues([[metricas.arrendamientosFiducias, metricas.administraciones, metricas.serviciosGenerales, (metricas.criticosCalidad + " | " + metricas.criticosSST + " | " + metricas.criticosAmbiental)]]).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(12).setFontColor(cSlate);
  hojaDash.getRange("B12:E13").setBorder(true, true, true, true, true, true, cBorde, SpreadsheetApp.BorderStyle.SOLID);

  hojaDash.getRange("G12:K12").setValues([["DOCUMENTOS AL DÍA", "", "POR ACTUALIZAR (<45 D)", "", "REQUERIMIENTO CRÍTICO"]]).setBackground(cGrisSuave).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
  hojaDash.getRange("G13:K13").setValues([[metricas.vigenSagrilaft, "", metricas.porVenceSagrilaft, "", metricas.vencidoSagrilaft]]).setHorizontalAlignment("center").setFontWeight("bold").setFontSize(13);
  hojaDash.getRange("G12:G13").setFontColor(cVerde); hojaDash.getRange("I12:I13").setFontColor(cNaranja); hojaDash.getRange("K12:K13").setFontColor(cRojo);
  hojaDash.getRange("G12:K13").setBorder(true, true, true, true, true, true, cBorde, SpreadsheetApp.BorderStyle.SOLID);

  hojaDash.getRange("B16").setValue("👥 Distribución Operativa del Equipo Interno").setFontSize(10).setFontWeight("bold").setFontColor(cSlate);
  hojaDash.getRange("B17:E17").setValues([["Comprador / Gestor", "Expedientes Activos", "Retrasos Pendientes", "% de Eficiencia"]]).setBackground(cGrisSuave).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");

  var filaEq = 18;
  for (var k = 0; k < listaCompradores.length; k++) {
    var eq = listaCompradores[k];
    var ef = eq.datos.activos > 0 ? Math.round(((eq.datos.activos - eq.datos.alertas) / eq.datos.activos) * 100) : 0;
    hojaDash.getRange("B" + filaEq).setValue(eq.name).setFontSize(9);
    hojaDash.getRange("C" + filaEq).setValue(eq.datos.activos).setHorizontalAlignment("center").setFontSize(9);
    hojaDash.getRange("D" + filaEq).setValue(eq.datos.alertas).setHorizontalAlignment("center").setFontWeight("bold").setFontColor(eq.datos.alertas > 0 ? cRojo : "#7f8c8d").setFontSize(9);
    hojaDash.getRange("E" + filaEq).setValue(ef + "%").setHorizontalAlignment("center").setFontWeight("bold").setFontColor(ef === 100 ? cVerde : cSlate).setFontSize(9);
    filaEq++;
  }
  hojaDash.getRange("B17:E" + (filaEq - 1)).setBorder(true, true, true, true, true, true, cBorde, SpreadsheetApp.BorderStyle.SOLID);

  var filaAcc = Math.max(filaEq + 2, 22);
  hojaDash.getRange("B" + filaAcc).setValue("📋 Plan de Acción (Proveedores que requieren gestión de documentos)").setFontSize(10).setFontWeight("bold").setFontColor(cSlate);
  filaAcc++;

  var filaHeaderAccion = filaAcc;
  hojaDash.getRange("B" + filaHeaderAccion + ":J" + filaHeaderAccion).setValues([["Urgencia", "Proveedor / Tercero", "Gestor", "Estado del Expediente", "Eval.", "Cal.", "SST", "Amb.", "Acceso Directo"]]).setBackground(cGrisSuave).setFontColor(cSlate).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");
  filaAcc++;

  if (planDeAccionPrioritario.length === 0) {
    hojaDash.getRange("B" + filaAcc + ":J" + filaAcc).merge().setValue("Todos los expedientes se encuentran vigentes y completos.").setFontStyle("italic").setFontColor(cVerde).setHorizontalAlignment("center");
    hojaDash.getRange("B" + filaHeaderAccion + ":J" + filaAcc).setBorder(true, true, true, true, true, true, cBorde, SpreadsheetApp.BorderStyle.SOLID);
  } else {
    var totalCasos = planDeAccionPrioritario.length;
    var casosAMostrar = Math.min(40, totalCasos);

    for (var a = 0; a < casosAMostrar; a++) {
      var act = planDeAccionPrioritario[a];
      var colorUrg = act.gravedad === 2 ? cRojo : cNaranja;
      var tagUrg = act.gravedad === 2 ? "🔴 CRÍTICO" : "🟡 ADVERTENCIA";

      hojaDash.getRange("B" + filaAcc).setValue(tagUrg).setFontColor(colorUrg).setFontWeight("bold").setFontSize(8).setHorizontalAlignment("center");
      hojaDash.getRange("C" + filaAcc).setValue(act.name).setFontWeight("bold").setFontSize(9);
      hojaDash.getRange("D" + filaAcc).setValue(act.gestor).setHorizontalAlignment("center").setFontSize(9);
      hojaDash.getRange("E" + filaAcc).setValue(act.obs).setFontColor(colorUrg).setFontSize(9);
      hojaDash.getRange("F" + filaAcc).setValue(act.eval).setHorizontalAlignment("center").setFontSize(9);
      hojaDash.getRange("G" + filaAcc).setValue(act.rCal).setHorizontalAlignment("center").setFontSize(8).setFontColor(act.rCal.includes("NO") ? "#7f8c8d" : cRojo);
      hojaDash.getRange("H" + filaAcc).setValue(act.rSst).setHorizontalAlignment("center").setFontSize(8).setFontColor(act.rSst.includes("NO") ? "#7f8c8d" : cRojo);
      hojaDash.getRange("I" + filaAcc).setValue(act.rAmb).setHorizontalAlignment("center").setFontSize(8).setFontColor(act.rAmb.includes("NO") ? "#7f8c8d" : cRojo);
      hojaDash.getRange("J" + filaAcc).setFormula('=HYPERLINK("#gid=' + idHojaDatos + '&range=A' + act.fila + '"; "✏️ Fila ' + act.fila + ' →")').setFontSize(8).setHorizontalAlignment("center").setFontWeight("bold");
      filaAcc++;
    }

    hojaDash.getRange("B" + filaHeaderAccion + ":J" + (filaAcc - 1)).setBorder(true, true, true, true, true, true, cBorde, SpreadsheetApp.BorderStyle.SOLID);
    for (var f = filaHeaderAccion + 1; f < filaAcc; f++) { if (f % 2 === 0) hojaDash.getRange("B" + f + ":J" + f).setBackground("#fcfdfe"); }

    if (totalCasos > 40) {
      hojaDash.getRange("B" + filaAcc).setValue("⚠️ Mostrando 40 de " + totalCasos + " casos totales — hay " + (totalCasos - 40) + " más sin mostrar aquí.").setFontStyle("italic").setFontColor(cRojo).setFontWeight("bold");
      filaAcc++;
    }
  }

  hojaDash.setColumnWidth(1, 15); hojaDash.setColumnWidth(2, 110); hojaDash.setColumnWidth(3, 250); hojaDash.setColumnWidth(4, 120);
  hojaDash.setColumnWidth(5, 180); hojaDash.setColumnWidth(6, 60); hojaDash.setColumnWidth(7, 65); hojaDash.setColumnWidth(8, 65);
  hojaDash.setColumnWidth(9, 65); hojaDash.setColumnWidth(10, 110);

  try { SpreadsheetApp.getUi().alert("📊 ¡Éxito! Dashboard de Proveedores Saneado y Completo."); } catch(e){}
}

// =========================================================================
// MÓDULO 5: AUDITORÍA DE RESPUESTAS EN GMAIL — VERSIÓN MULTICAPA
// =========================================================================
// UBICACIÓN: Bound al mismo Sheet. Ejecución manual desde el menú.
// PERMISOS: Sheets + Gmail (lectura). Sin permisos nuevos.
// DEPENDENCIA: usa NOMBRE_HOJA_BITACORA e inicializarEstructuraEscalamiento()
// del Módulo 7, en este mismo archivo.
//
// CAPAS DE DETECCIÓN, en orden de ejecución:
//   CAPA 0 — REBOTE: mailer-daemon/postmaster mencionando el correo del
//     proveedor después de la fecha de envío.
//   CAPA 1 — ID de hilo guardado (columna L). Certeza máxima.
//   CAPA 2 — Búsqueda exacta por destinatario + asunto. Certeza alta.
//     Si encuentra algo y la columna L estaba vacía, se AUTO-REPARA.
//   CAPA 3a/3b/3c — heurísticas (remitente exacto / dominio corporativo /
//     NIT en el cuerpo). Certeza media-baja. NUNCA tocan las columnas K/L
//     "oficiales" del caso — dejan un enlace clicable en la columna M
//     para que un humano verifique con un clic.
//
// LÍMITE ESTRUCTURAL: el script solo ve el buzón de la cuenta que lo ejecuta. Si el
// proveedor responde a otro buzón sin copiar esta cuenta, no hay forma de
// detectarlo desde aquí — mitigación de proceso, no de código.
// =========================================================================

var COL_TC_FECHA_GESTION = 1;    // A
var COL_TC_NIT = 2;              // B
var COL_TC_PROVEEDOR = 3;        // C
var COL_TC_ESTADO = 6;           // F
var COL_TC_CORREO_DESTINO = 8;   // H
var COL_TC_ENLACE_CORREO = 11;   // K
var COL_TC_ID_HILO = 12;         // L
var COL_TC_NOTA_ROBOT = 13;      // M

var DOMINIOS_PUBLICOS = [
  "gmail.com", "hotmail.com", "outlook.com", "yahoo.com", "yahoo.es",
  "live.com", "icloud.com", "aol.com", "protonmail.com", "msn.com",
  "hotmail.es", "outlook.es"
];

function registrarSoloEnBitacora(fila, nit, proveedor, tipoAccion, responsable, comentario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaBitacora = ss.getSheetByName(NOMBRE_HOJA_BITACORA);
  if (!hojaBitacora) {
    inicializarEstructuraEscalamiento();
    hojaBitacora = ss.getSheetByName(NOMBRE_HOJA_BITACORA);
  }
  var ahora = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");
  hojaBitacora.appendRow([ahora, nit, proveedor, fila, tipoAccion, responsable, comentario || ""]);
}

function parsearFechaGestionTexto(texto) {
  var fechaParte = String(texto).trim().split(" ")[0];
  var dmy = fechaParte.split("/");
  if (dmy.length !== 3) return null;
  var dia = parseInt(dmy[0], 10), mes = parseInt(dmy[1], 10) - 1, anio = parseInt(dmy[2], 10);
  if (isNaN(dia) || isNaN(mes) || isNaN(anio)) return null;
  return new Date(anio, mes, dia);
}

function variantesNit(nitLimpio) {
  if (nitLimpio.length < 6) return [nitLimpio];
  var conPuntos = nitLimpio.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return [nitLimpio, conPuntos];
}

function urlDeHilo(idHilo) {
  return "https://mail.google.com/mail/u/0/#all/" + idHilo;
}

// -------------------------------------------------------------------
// Agrega una entrada nueva a la columna M SIN borrar lo que ya había —
// las entradas quedan separadas por " - ", cada una con su fecha. Los
// enlaces de entradas anteriores TAMBIÉN se conservan clicables.
// -------------------------------------------------------------------
function agregarNotaHistorial(hojaControl, filaReal, textoNuevo, urlNueva) {
  var celda = hojaControl.getRange(filaReal, COL_TC_NOTA_ROBOT);
  var richActual = celda.getRichTextValue();
  var textoActual = richActual ? richActual.getText() : "";

  var separador = textoActual === "" ? "" : " - ";
  var textoCombinado = textoActual + separador + textoNuevo;

  var builder = SpreadsheetApp.newRichTextValue().setText(textoCombinado);

  if (richActual) {
    var runs = richActual.getRuns();
    for (var r = 0; r < runs.length; r++) {
      var urlExistente = runs[r].getLinkUrl();
      if (urlExistente) {
        builder.setLinkUrl(runs[r].getStartIndex(), runs[r].getEndIndex(), urlExistente);
      }
    }
  }

  if (urlNueva) {
    var inicioNuevo = textoActual.length + separador.length;
    var finNuevo = textoCombinado.length;
    builder.setLinkUrl(inicioNuevo, finNuevo, urlNueva);
  }

  celda.setRichTextValue(builder.build());
}

function auditarRespuestasGmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaControl = ss.getSheetByName("Tabla Control");

  if (!hojaControl) {
    SpreadsheetApp.getUi().alert("❌ Error: No se encuentra la pestaña 'Tabla Control'.");
    return;
  }

  var ultimaFila = hojaControl.getLastRow();
  if (ultimaFila < 2) {
    SpreadsheetApp.getUi().alert("ℹ️ La Tabla Control no tiene registros para auditar.");
    return;
  }

  var datosDisplay = hojaControl.getRange(2, 1, ultimaFila - 1, 13).getDisplayValues();

  var contadorSinRespuesta = 0;
  var contadorRespondidos = 0;
  var contadorPosibles = 0;
  var contadorRebotados = 0;
  var contadorNoEncontrado = 0;
  var contadorAutoReparados = 0;

  var estadosParaAuditar = [
    "Aviso Enviado Automático", "Aviso Enviado Manual", "Pendiente de respuesta",
    "Sin Respuesta (Seguimiento)", "🔍 Posible Respuesta (Verificar Manualmente)"
  ];

  for (var i = 0; i < datosDisplay.length; i++) {
    var filaReal = i + 2;
    var estadoActual = datosDisplay[i][COL_TC_ESTADO - 1].trim();
    var correoDestino = datosDisplay[i][COL_TC_CORREO_DESTINO - 1].trim();
    var nit = datosDisplay[i][COL_TC_NIT - 1].trim();
    var proveedor = datosDisplay[i][COL_TC_PROVEEDOR - 1].trim();
    var idHiloGuardado = datosDisplay[i][COL_TC_ID_HILO - 1].trim();
    var fechaGestionTexto = datosDisplay[i][COL_TC_FECHA_GESTION - 1];

    if (estadosParaAuditar.indexOf(estadoActual) === -1) continue;
    if (correoDestino === "") continue;

    var primerCorreo = correoDestino.split(",")[0].trim().toLowerCase();
    var fechaEnvio = parsearFechaGestionTexto(fechaGestionTexto);
    var fechaBusqueda = fechaEnvio
      ? Utilities.formatDate(fechaEnvio, ss.getSpreadsheetTimeZone(), "yyyy/MM/dd")
      : null;

    // ============ CAPA 0: ¿EL CORREO REBOTÓ? ============
    if (fechaBusqueda) {
      var rebotes = GmailApp.search(
        'from:(mailer-daemon OR postmaster) "' + primerCorreo + '" after:' + fechaBusqueda, 0, 1
      );
      if (rebotes.length > 0) {
        var estadoRebote = "❌ Correo Rebotado (Verificar Dirección)";
        if (estadoActual !== estadoRebote) {
          hojaControl.getRange(filaReal, COL_TC_ESTADO).setValue(estadoRebote);
          var fechaRebote = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
          agregarNotaHistorial(hojaControl, filaReal, fechaRebote + " - Se detectó rebote de " + primerCorreo, null);
          registrarSoloEnBitacora(filaReal, nit, proveedor,
            "❌ Auditoría Gmail: Correo Rebotado",
            "Sistema (Auditoría Gmail)",
            "Mailer-daemon reportó fallo de entrega a " + primerCorreo + ". La dirección probablemente no existe o está llena.");
        }
        contadorRebotados++;
        continue;
      }
    }

    var hilo = null;
    var capaDeteccion = "";

    // ============ CAPA 1: ID de hilo guardado ============
    if (idHiloGuardado !== "") {
      try { hilo = GmailApp.getThreadById(idHiloGuardado); } catch (e) { hilo = null; }
      if (hilo) { capaDeteccion = "ID"; }
    }

    // ============ CAPA 2: Búsqueda exacta por destinatario + asunto ============
    var seAutoReparo = false;
    if (!hilo) {
      var hilosExactos = GmailApp.search("to:" + primerCorreo + " subject:\"Solicitud de documentos para creación\"", 0, 1);
      if (hilosExactos.length > 0) {
        hilo = hilosExactos[0];
        capaDeteccion = "Asunto";

        if (idHiloGuardado === "") {
          var idEncontrado = hilo.getId();
          hojaControl.getRange(filaReal, COL_TC_ID_HILO).setValue(idEncontrado);
          var enlaceReparado = SpreadsheetApp.newRichTextValue().setText("Abrir Correo").setLinkUrl(urlDeHilo(idEncontrado)).build();
          hojaControl.getRange(filaReal, COL_TC_ENLACE_CORREO).setRichTextValue(enlaceReparado);
          seAutoReparo = true;
          contadorAutoReparados++;
        }
      }
    }

    // ============ CAPAS 3a/3b/3c: heurísticas de correo nuevo ============
    var viaHeuristica = false;
    if (!hilo && fechaBusqueda) {
      var porRemitente = GmailApp.search("from:" + primerCorreo + " after:" + fechaBusqueda, 0, 1);
      if (porRemitente.length > 0) {
        hilo = porRemitente[0]; viaHeuristica = true; capaDeteccion = "Remitente exacto";
      }

      if (!hilo) {
        var dominio = primerCorreo.split("@")[1] || "";
        if (dominio !== "" && DOMINIOS_PUBLICOS.indexOf(dominio) === -1) {
          var porDominio = GmailApp.search("from:@" + dominio + " after:" + fechaBusqueda, 0, 1);
          if (porDominio.length > 0) {
            hilo = porDominio[0]; viaHeuristica = true; capaDeteccion = "Dominio corporativo (@" + dominio + ")";
          }
        }
      }

      if (!hilo) {
        var nitLimpio = nit.replace(/[^0-9]/g, "");
        if (nitLimpio.length >= 6) {
          var formatosNit = variantesNit(nitLimpio);
          for (var v = 0; v < formatosNit.length && !hilo; v++) {
            var porNit = GmailApp.search('"' + formatosNit[v] + '" after:' + fechaBusqueda + ' -from:me', 0, 1);
            if (porNit.length > 0) {
              hilo = porNit[0]; viaHeuristica = true; capaDeteccion = "NIT en cuerpo (" + formatosNit[v] + ")";
            }
          }
        }
      }
    }

    if (!hilo) {
      contadorNoEncontrado++;
      continue;
    }

    var nuevoEstado, tipoAccionBitacora, comentarioBitacora;

    if (viaHeuristica) {
      nuevoEstado = "🔍 Posible Respuesta (Verificar Manualmente)";
      tipoAccionBitacora = "🔍 Auditoría Gmail: Posible Respuesta (" + capaDeteccion + ")";
      comentarioBitacora = "Detectado por capa heurística: " + capaDeteccion + ". No coincide por ID ni asunto — requiere verificación humana. Enlace: " + urlDeHilo(hilo.getId());
    } else {
      var mensajes = hilo.getMessages();
      var ultimoMensaje = mensajes[mensajes.length - 1];
      var remitenteUltimo = ultimoMensaje.getFrom().toLowerCase();

      var matches = remitenteUltimo.match(/@([a-z0-9.-]+)/);
      var ultimoEsInterno = false;
      if (matches && matches.length > 1 && CONFIG.DOMINIOS_INTERNOS.indexOf(matches[1].trim()) !== -1) {
        ultimoEsInterno = true;
      }

      nuevoEstado = ultimoEsInterno ? "Sin Respuesta (Seguimiento)" : "Respondido al Hilo";
      tipoAccionBitacora = "🔄 Auditoría Gmail: " + (ultimoEsInterno ? "Sin Respuesta" : "Respondido") + " (vía " + capaDeteccion + (seAutoReparo ? ", auto-reparado" : "") + ")";
      comentarioBitacora = ultimoEsInterno
        ? "El último mensaje del hilo es interno — el proveedor no ha respondido."
        : "Se detectó un mensaje de un dominio externo en el hilo.";
    }

    if (nuevoEstado !== estadoActual) {
      hojaControl.getRange(filaReal, COL_TC_ESTADO).setValue(nuevoEstado);

      if (viaHeuristica) {
        var fechaDeteccion = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        var textoNota = fechaDeteccion + " - Se encuentra posible correo (" + capaDeteccion + ")";
        agregarNotaHistorial(hojaControl, filaReal, textoNota, urlDeHilo(hilo.getId()));
      }

      registrarSoloEnBitacora(filaReal, nit, proveedor, tipoAccionBitacora, "Sistema (Auditoría Gmail)", comentarioBitacora);
    }

    if (viaHeuristica) contadorPosibles++;
    else if (nuevoEstado === "Sin Respuesta (Seguimiento)") contadorSinRespuesta++;
    else contadorRespondidos++;
  }

  var reporte = "📊 AUDITORÍA DE RESPUESTAS TERMINADA 📊\n\n";
  reporte += "🔴 Sin Respuesta (Mudos totales): " + contadorSinRespuesta + "\n";
  reporte += "🟢 Respuestas confirmadas al hilo: " + contadorRespondidos + "\n";
  reporte += "🔍 Posibles respuestas por correo nuevo (verificar, con enlace en columna M): " + contadorPosibles + "\n";
  reporte += "❌ Correos rebotados (dirección mala): " + contadorRebotados + "\n";
  reporte += "🔧 Enlaces auto-reparados (columna L estaba vacía): " + contadorAutoReparados + "\n";
  reporte += "⚠️ Sin rastro alguno: " + contadorNoEncontrado + "\n\n";
  reporte += "Detalles de cada detección quedaron en la Bitácora de Auditoría.";

  SpreadsheetApp.getUi().alert(reporte);
}

// =========================================================================
// MÓDULO 6: MENÚ
// =========================================================================
// El menú de la barra está en Menu.gs, junto con el de contratos: Apps
// Script admite una sola función onOpen() por proyecto.
// =========================================================================

// =========================================================================
// MÓDULO 7: ESCALAMIENTO Y BITÁCORA DE AUDITORÍA (Tabla Control)
// =========================================================================
// UBICACIÓN: Bound al Google Sheet de Proveedores.
// Agrega columnas N-Q a "Tabla Control" y una hoja "Bitácora Auditoría"
// que nunca se edita ni se borra: solo se le añaden filas.
//
// PERMISOS:
// - Session.getActiveUser().getEmail(): scope de identidad.
// - CalendarApp (festivos Colombia): permiso de solo lectura.
// =========================================================================

var COL_ESTADO_ESCALAMIENTO = 14; // N
var COL_FECHA_ULTIMA_ACCION = 15; // O
var COL_RESPONSABLE = 16;         // P
var COL_ALERTA_AUTOMATICA = 17;   // Q
var COL_PROXIMO_SEGUIMIENTO = 7;  // G
var DIAS_PROXIMO_SEGUIMIENTO = 5; // En DÍAS HÁBILES

var NOMBRE_HOJA_BITACORA = "Bitácora Auditoría";

function verificarCalendarioFestivos() {
  var ui = SpreadsheetApp.getUi();
  try {
    var calendario = CalendarApp.getCalendarById(CONFIG.ID_CALENDARIO_FESTIVOS_CO);
    if (!calendario) {
      ui.alert("❌ No se pudo acceder al calendario. Revisa el ID o el permiso de Calendario.");
      return;
    }
    var primeroEnero = new Date(new Date().getFullYear(), 0, 1);
    var eventos = calendario.getEventsForDay(primeroEnero);
    if (eventos.length > 0) {
      ui.alert("✅ Calendario verificado.\n\nEvento el 1 de enero: \"" + eventos[0].getTitle() + "\"");
    } else {
      ui.alert("⚠️ No se encontró evento el 1 de enero. El ID probablemente NO es correcto para Colombia. No confíes en días hábiles todavía.");
    }
  } catch (e) {
    ui.alert("❌ Error al acceder al calendario: " + e.message);
  }
}

function esFinDeSemanaOFestivo(fecha) {
  var diaSemana = fecha.getDay();
  if (diaSemana === 0 || diaSemana === 6) return true;
  try {
    var calendario = CalendarApp.getCalendarById(CONFIG.ID_CALENDARIO_FESTIVOS_CO);
    var eventos = calendario.getEventsForDay(fecha);
    return eventos.length > 0;
  } catch (e) {
    Logger.log("⚠️ Chequeo de festivos caído para " + fecha + ": " + e.message);
    return false;
  }
}

function sumarDiasHabiles(fechaInicio, diasHabiles) {
  var fecha = new Date(fechaInicio);
  var contados = 0;
  while (contados < diasHabiles) {
    fecha.setDate(fecha.getDate() + 1);
    if (!esFinDeSemanaOFestivo(fecha)) contados++;
  }
  return fecha;
}

function diasHabilesEntre(fechaInicio, fechaFin) {
  var fecha = new Date(fechaInicio);
  var contados = 0;
  while (fecha.getTime() < fechaFin.getTime()) {
    fecha.setDate(fecha.getDate() + 1);
    if (!esFinDeSemanaOFestivo(fecha)) contados++;
  }
  return contados;
}

function inicializarEstructuraEscalamiento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaControl = ss.getSheetByName("Tabla Control");
  if (!hojaControl) {
    throw new Error("No se encontró la hoja 'Tabla Control'. Verifica el nombre exacto.");
  }

  var encabezadoN = hojaControl.getRange(1, COL_ESTADO_ESCALAMIENTO).getValue();
  if (encabezadoN === "") {
    hojaControl.getRange(1, COL_ESTADO_ESCALAMIENTO, 1, 4).setValues(
      [["Estado Escalamiento", "Fecha Última Acción", "Responsable", "Alerta Automática"]]
    ).setFontWeight("bold").setBackground("#37474f").setFontColor("white");
  }

  var hojaBitacora = ss.getSheetByName(NOMBRE_HOJA_BITACORA);
  if (!hojaBitacora) {
    hojaBitacora = ss.insertSheet(NOMBRE_HOJA_BITACORA);
    hojaBitacora.appendRow([
      "Fecha y Hora", "NIT", "Proveedor", "Fila en Tabla Control",
      "Tipo de Acción", "Responsable", "Comentario"
    ]);
    hojaBitacora.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#37474f").setFontColor("white");
    hojaBitacora.setFrozenRows(1);

    try {
      var proteccion = hojaBitacora.protect().setDescription("Bitácora auditoría - solo el sistema escribe aquí");
      proteccion.removeEditors(proteccion.getEditors());
      proteccion.setDomainEdit(false);
    } catch (e) {
      Logger.log("No se pudo aplicar protección de dominio a la Bitácora: " + e.message);
    }

    hojaBitacora.hideSheet();
  }
}

function mostrarBitacora() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(NOMBRE_HOJA_BITACORA);
  if (!hoja) { SpreadsheetApp.getUi().alert("La Bitácora aún no existe. Ejecuta primero 'inicializarEstructuraEscalamiento'."); return; }
  hoja.showSheet();
  ss.setActiveSheet(hoja);
}

function ocultarBitacora() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(NOMBRE_HOJA_BITACORA);
  if (!hoja) return;
  var hojaControl = ss.getSheetByName("Tabla Control");
  if (ss.getActiveSheet().getName() === NOMBRE_HOJA_BITACORA && hojaControl) {
    ss.setActiveSheet(hojaControl);
  }
  hoja.hideSheet();
}

function escribirAccionEscalamiento(fila, tipoAccion, responsable, comentario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaControl = ss.getSheetByName("Tabla Control");
  var hojaBitacora = ss.getSheetByName(NOMBRE_HOJA_BITACORA);

  if (!hojaControl || !hojaBitacora) {
    inicializarEstructuraEscalamiento();
    hojaControl = ss.getSheetByName("Tabla Control");
    hojaBitacora = ss.getSheetByName(NOMBRE_HOJA_BITACORA);
  }

  var nit = String(hojaControl.getRange(fila, 2).getValue()).trim();
  var proveedor = String(hojaControl.getRange(fila, 3).getValue()).trim();
  var ahora = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");

  hojaBitacora.appendRow([ahora, nit, proveedor, fila, tipoAccion, responsable, comentario || ""]);

  hojaControl.getRange(fila, COL_ESTADO_ESCALAMIENTO).setValue(tipoAccion);
  hojaControl.getRange(fila, COL_FECHA_ULTIMA_ACCION).setValue(ahora);
  hojaControl.getRange(fila, COL_RESPONSABLE).setValue(responsable);

  if (responsable !== "Sistema (automático)") {
    hojaControl.getRange(fila, COL_ALERTA_AUTOMATICA).setValue("");
  }
}

function registrarAccionEscalamientoDesdeMenu(tipoAccion, actualizaProximoSeguimiento) {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaActiva = ss.getActiveSheet();
  var hojaControl = ss.getSheetByName("Tabla Control");

  if (hojaActiva.getName() !== "Tabla Control") {
    ui.alert("⚠️ Debes tener la pestaña 'Tabla Control' abierta y seleccionar la fila del proveedor antes de usar esta opción.");
    return;
  }

  var fila = ss.getActiveRange().getRow();
  if (fila < 2) {
    ui.alert("⚠️ Selecciona una fila de datos (no el encabezado) antes de continuar.");
    return;
  }

  var proveedor = String(hojaControl.getRange(fila, 3).getValue()).trim();
  var nit = String(hojaControl.getRange(fila, 2).getValue()).trim();

  if (proveedor === "") {
    ui.alert("⚠️ La fila seleccionada no tiene proveedor. Verifica que seleccionaste la fila correcta.");
    return;
  }

  var confirmacion = ui.alert(
    "Confirmar acción de escalamiento",
    "Vas a registrar: \"" + tipoAccion + "\"\n\nProveedor: " + proveedor + "\nNIT: " + nit + "\nFila: " + fila +
    "\n\n¿Es correcto?",
    ui.ButtonSet.YES_NO
  );

  if (confirmacion !== ui.Button.YES) return;

  var respuestaComentario = ui.prompt(
    "Comentario (opcional)",
    "¿Algo que quieras dejar registrado sobre esta acción? Deja vacío si no aplica.",
    ui.ButtonSet.OK_CANCEL
  );

  if (respuestaComentario.getSelectedButton() !== ui.Button.OK) return;
  var comentario = respuestaComentario.getResponseText().trim();

  var responsable = Session.getActiveUser().getEmail();
  if (!responsable) {
    var respNombre = ui.prompt("No se pudo detectar tu usuario automáticamente", "Escribe tu nombre para dejarlo en el registro:", ui.ButtonSet.OK_CANCEL);
    if (respNombre.getSelectedButton() !== ui.Button.OK) return;
    responsable = respNombre.getResponseText().trim() || "Sin identificar";
  }

  escribirAccionEscalamiento(fila, tipoAccion, responsable, comentario);

  var fechaSoloTexto = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  var textoHistorial = fechaSoloTexto + " - " + tipoAccion + (comentario ? " (" + comentario + ")" : "");
  agregarNotaHistorial(hojaControl, fila, textoHistorial, null);

  if (actualizaProximoSeguimiento) {
    var proximaFecha = sumarDiasHabiles(new Date(), DIAS_PROXIMO_SEGUIMIENTO);
    var proximaFechaTexto = Utilities.formatDate(proximaFecha, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    hojaControl.getRange(fila, COL_PROXIMO_SEGUIMIENTO).setValue(proximaFechaTexto);
  }

  ui.alert("✅ Registrado: " + tipoAccion + " para " + proveedor);
}

function marcarSeguimientoTelefonico() { registrarAccionEscalamientoDesdeMenu("📞 Seguimiento Telefónico", false); }
function marcarContactoWhatsApp()      { registrarAccionEscalamientoDesdeMenu("💬 Contacto WhatsApp", false); }
function marcarResuelto()              { registrarAccionEscalamientoDesdeMenu("✅ Resuelto", false); }
function reabrirCaso()                 { registrarAccionEscalamientoDesdeMenu("🔁 Reabierto", false); }
function marcarRecordatorioCorreo()    { registrarAccionEscalamientoDesdeMenu("📧 Recordatorio Enviado (Correo)", true); }

function detectarCasosSinSeguimiento() {
  var DIAS_LIMITE_SIN_ACCION = 5;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaControl = ss.getSheetByName("Tabla Control");
  if (!hojaControl) return;

  var ultimaFila = hojaControl.getLastRow();
  if (ultimaFila < 2) return;

  var datos = hojaControl.getRange(2, 1, ultimaFila - 1, COL_ALERTA_AUTOMATICA).getValues();
  var hoy = new Date();
  var marcados = 0;

  for (var i = 0; i < datos.length; i++) {
    var fila = i + 2;
    var proveedor = String(datos[i][2]).trim();
    if (proveedor === "") continue;

    var estadoEnvio = String(datos[i][5]).trim();
    var estadoEscalamiento = String(datos[i][COL_ESTADO_ESCALAMIENTO - 1]).trim();
    var fechaUltimaAccion = datos[i][COL_FECHA_ULTIMA_ACCION - 1];

    var envioValido = estadoEnvio.includes("Aviso Enviado") || estadoEnvio.includes("Sin Respuesta") || estadoEnvio.includes("Correo Rebotado") || estadoEnvio.includes("Posible Respuesta");
    if (!envioValido) continue;
    if (estadoEscalamiento === "✅ Resuelto") continue;

    var fechaReferencia = (fechaUltimaAccion instanceof Date) ? fechaUltimaAccion : datos[i][0];
    if (!(fechaReferencia instanceof Date)) continue;

    var diasSinAccion = diasHabilesEntre(fechaReferencia, hoy);

    if (diasSinAccion >= DIAS_LIMITE_SIN_ACCION && estadoEscalamiento !== "⚠️ Requiere Escalamiento") {
      hojaControl.getRange(fila, COL_ALERTA_AUTOMATICA).setValue("⚠️ Requiere Escalamiento (" + diasSinAccion + "d)");
      marcados++;
    }
  }

  Logger.log("Casos marcados para escalamiento manual: " + marcados);
}

function instalarTriggerDeteccionEscalamiento() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "detectarCasosSinSeguimiento") {
      SpreadsheetApp.getUi().alert("Ya existe un trigger para esta función. No se creó uno nuevo.");
      return;
    }
  }
  ScriptApp.newTrigger("detectarCasosSinSeguimiento")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  SpreadsheetApp.getUi().alert("✅ Trigger diario instalado: se ejecutará todos los días alrededor de las 7am.");
}

// =========================================================================
// MÓDULO EXTRA: FaltantePorCalificacion (SINCRONIZACIÓN + FUZZY MATCH)
// =========================================================================
// UBICACIÓN: Bound al mismo Sheet. Ejecución manual desde el menú.
// PERMISOS: Sheets únicamente — abre OTRO archivo de Sheets externo.
// =========================================================================

function FaltantePorCalificacion() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaMaster = ss.getSheetByName("BASE DE PROVEEDORES COMPRAS");
  if (!hojaMaster) return SpreadsheetApp.getUi().alert("❌ Error: No se encuentra la base maestra.");

  var ssForm;
  try {
    ssForm = SpreadsheetApp.openByUrl(CONFIG.URL_FORMULARIO_CALIFICACIONES);
  } catch(e) {
    return SpreadsheetApp.getUi().alert("❌ Error Crítico: No puedo abrir el enlace.\n\nAsegúrate de copiar bien la URL completa del archivo de respuestas.");
  }

  var nombreArchivo = ssForm.getName();
  // ⚠️ Ajusta esta validación con el/los identificadores de versión que
  // correspondan a TU formulario vigente, para evitar leer accidentalmente
  // un archivo viejo/duplicado.
  if (nombreArchivo.includes("VERSION_ESPERADA") === false) {
     SpreadsheetApp.getUi().alert(
       "🛑 ¡ALTO AHÍ! ESTÁS LEYENDO EL ARCHIVO EQUIVOCADO 🛑",
       "El robot acaba de abrir este archivo:\n👉 '" + nombreArchivo + "'\n\nRevisa si es tu formulario vigente. Si no lo es, corrige CONFIG.URL_FORMULARIO_CALIFICACIONES.",
       SpreadsheetApp.getUi().ButtonSet.OK
     );
     return;
  }

  var hojaForm = ssForm.getSheets().find(s => s.getName().toLowerCase().includes("respuesta") || s.getName().toLowerCase().includes("formulario"));
  if (!hojaForm) hojaForm = ssForm.getSheets()[0];

  var ui = SpreadsheetApp.getUi();
  var forzarTodo = (ui.alert("🚀 SINCRONIZACIÓN MAESTRA", "¿Escanear todo el historial ignorando duplicados?", ui.ButtonSet.YES_NO) == ui.Button.YES);

  var rangoMaster = hojaMaster.getRange(5, 1, hojaMaster.getLastRow() - 4, hojaMaster.getLastColumn()).getValues();
  var rangoFormValores = hojaForm.getDataRange().getValues();
  var rangoFormTextos = hojaForm.getDataRange().getDisplayValues();

  var headersMaster = rangoMaster[0].map(h => String(h).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
  var idxMasterNombre = headersMaster.indexOf("nombre proveedor");
  var idxMasterNit = headersMaster.indexOf("nit");
  var idxMasterComprador = headersMaster.indexOf("comprador");
  if(idxMasterComprador === -1) idxMasterComprador = 1;
  var idxMasterFechaEval = headersMaster.findIndex(h => h.includes("fecha") && h.includes("evaluaci"));
  var idxMasterResultado = headersMaster.findIndex(h => h.includes("resultado") && (h.includes("evaluaci") || h.includes("utima") || h.includes("util")));

  var headersForm = rangoFormTextos[0].map(h => String(h).trim().toLowerCase());
  var idxFormFecha = headersForm.findIndex(h => h.includes("marca temporal") || h.includes("fecha"));
  var idxFormNit = headersForm.findIndex(h => h === "nit" || h.includes("nit"));
  var idxFormNombre = headersForm.findIndex(h => h.includes("proveedor a calificar") || h.includes("proveedor"));
  var idxFormScore = headersForm.findIndex(h => h.includes("promedio"));
  var idxFormEstado = headersForm.findIndex(h => h.includes("estado robot"));

  if (idxFormEstado === -1) {
    hojaForm.getRange(1, headersForm.length + 1).setValue("Estado Robot");
    idxFormEstado = headersForm.length;
    rangoFormValores = hojaForm.getDataRange().getValues();
    rangoFormTextos = hojaForm.getDataRange().getDisplayValues();
  }

  var mapaMasterPorNit = {};
  var mapaMasterPorNombre = {};

  for (var m = 1; m < rangoMaster.length; m++) {
    var estadoComprador = String(rangoMaster[m][idxMasterComprador]).toUpperCase();
    if (estadoComprador.includes("INACTIVO") || estadoComprador.includes("INHABILITADO")) continue;

    var nitM = String(rangoMaster[m][idxMasterNit]).replace(/[^0-9]/g, "").trim();
    var nombreM = limpiarParaComparar(rangoMaster[m][idxMasterNombre]);
    var filaReal = m + 5;

    if (nitM !== "" && !mapaMasterPorNit[nitM]) mapaMasterPorNit[nitM] = filaReal;
    if (nombreM !== "" && !mapaMasterPorNombre[nombreM]) mapaMasterPorNombre[nombreM] = filaReal;
  }

  var actualizacionesPorFilaMaster = {};
  var filasFormAActualizar = [];

  for (var f = 1; f < rangoFormValores.length; f++) {
    var filaTexto = rangoFormTextos[f];
    var filaValor = rangoFormValores[f];

    if (!filaTexto || filaTexto[idxFormNombre] === "") continue;
    if (!forzarTodo && filaTexto[idxFormEstado] === "PROCESADO") continue;

    var nitForm = String(filaTexto[idxFormNit]).replace(/[^0-9]/g, "").trim();
    var nombreForm = limpiarParaComparar(filaTexto[idxFormNombre]);
    var score = filaValor[idxFormScore];

    var strFecha = String(filaTexto[idxFormFecha]).split(" ")[0].trim();
    var ts = 0;
    if (filaValor[idxFormFecha] instanceof Date) {
       ts = filaValor[idxFormFecha].getTime();
    } else {
       var match = strFecha.match(/(\d{1,4})[\/\-](\d{1,2})[\/\-](\d{1,4})/);
       if (match) {
          var p1 = parseInt(match[1], 10), p2 = parseInt(match[2], 10), p3 = parseInt(match[3], 10);
          if (p1 > 1000) ts = new Date(p1, p2 - 1, p3).getTime();
          else if (p3 > 1000) ts = new Date(p3, p2 - 1, p1).getTime();
       }
       if (isNaN(ts) || ts === 0) ts = new Date(strFecha).getTime();
    }

    var filaDestinoMaster = null;
    var tipoMatch = "";

    if (nitForm !== "" && mapaMasterPorNit[nitForm]) {
      filaDestinoMaster = mapaMasterPorNit[nitForm];
      tipoMatch = "NIT exacto";
    } else if (nombreForm !== "" && mapaMasterPorNombre[nombreForm]) {
      filaDestinoMaster = mapaMasterPorNombre[nombreForm];
      tipoMatch = "Nombre exacto";
    } else {
       var mejorSimilitud = 0;
       var filaCandidata = null;
       var nombreCandidato = "";
       for (var claveNombre in mapaMasterPorNombre) {
          if (nombreForm.length > 4 && (claveNombre.includes(nombreForm) || nombreForm.includes(claveNombre))) {
             filaDestinoMaster = mapaMasterPorNombre[claveNombre];
             tipoMatch = "Substring (\"" + nombreForm + "\" ~ \"" + claveNombre + "\")";
             break;
          }
          var similitudActual = calcularPorcentajeSimilitud(nombreForm, claveNombre);
          if (similitudActual > mejorSimilitud) {
            mejorSimilitud = similitudActual;
            filaCandidata = mapaMasterPorNombre[claveNombre];
            nombreCandidato = claveNombre;
          }
       }
       if (!filaDestinoMaster && mejorSimilitud >= 85) {
         filaDestinoMaster = filaCandidata;
         tipoMatch = "Fuzzy " + mejorSimilitud + "% (\"" + nombreForm + "\" ~ \"" + nombreCandidato + "\")";
       }
    }

    if (filaDestinoMaster) {
       if (!actualizacionesPorFilaMaster[filaDestinoMaster] || ts > actualizacionesPorFilaMaster[filaDestinoMaster].ts) {
          actualizacionesPorFilaMaster[filaDestinoMaster] = { ts: ts, fechaTexto: strFecha, score: score, tipoMatch: tipoMatch, nombreForm: nombreForm };
       }
       filasFormAActualizar.push(f + 1);
    }
  }

  var actualizados = 0;
  var actualizadosConDudaBaja = 0;

  for (var fila in actualizacionesPorFilaMaster) {
     var data = actualizacionesPorFilaMaster[fila];
     hojaMaster.getRange(parseInt(fila), idxMasterFechaEval + 1).setValue("'" + data.fechaTexto);
     var celdaScore = hojaMaster.getRange(parseInt(fila), idxMasterResultado + 1);
     celdaScore.setValue(parseFloat(String(data.score).replace(",", ".")));
     celdaScore.setNumberFormat("0.0");
     actualizados++;

     if (data.tipoMatch.indexOf("Substring") === 0 || data.tipoMatch.indexOf("Fuzzy") === 0) {
       var nitFilaDestino = String(hojaMaster.getRange(parseInt(fila), idxMasterNit + 1).getValue()).trim();
       var nombreFilaDestino = String(hojaMaster.getRange(parseInt(fila), idxMasterNombre + 1).getValue()).trim();
       registrarSoloEnBitacora(
         parseInt(fila), nitFilaDestino, nombreFilaDestino,
         "🔄 Sincronización Calificaciones: match no exacto",
         "Sistema (FaltantePorCalificacion)",
         "Tipo de match: " + data.tipoMatch + ". Verificar que corresponda al proveedor correcto."
       );
       actualizadosConDudaBaja++;
     }
  }

  if (filasFormAActualizar.length > 0) {
    for (var idx = 0; idx < filasFormAActualizar.length; idx++) {
       hojaForm.getRange(filasFormAActualizar[idx], idxFormEstado + 1).setValue("PROCESADO");
    }
  }

  ui.alert("✅ ÉXITO TOTAL\n\n▶️ Proveedores actualizados: " + actualizados +
    "\n⚠️ De ellos, con match NO exacto (revisar en Bitácora): " + actualizadosConDudaBaja +
    "\n\nSe leyó el formulario correcto y las fechas se forzaron.");
}

function limpiarParaComparar(texto) {
  if (!texto) return "";
  var txt = texto.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  // Ejemplo de caso especial de normalización de nombre — reemplaza o
  // elimina esta línea según tus propios casos reales.
  // if (txt.includes("nombre_especial")) return "nombrenormalizado";
  return txt.replace(/\bs\.a\.s\.?\b|\bsas\b|\bltda\.?\b|\blimitada\b|\bs\.a\.?\b|\bsa\b|\be\.u\.?\b|\beu\b|\bbic\b|\bips\b|\bcorp\b/g, "").replace(/[^a-z0-9]/g, "").trim();
}

function calcularPorcentajeSimilitud(s1, s2) {
  var longer = s1; var shorter = s2;
  if (s1.length < s2.length) { longer = s2; shorter = s1; }
  var longerLength = longer.length;
  if (longerLength === 0) return 100;
  var distance = calcularDistanciaEdicion(longer, shorter);
  return Math.round(((longerLength - distance) / longerLength) * 100);
}

function calcularDistanciaEdicion(s1, s2) {
  s1 = s1.toLowerCase(); s2 = s2.toLowerCase();
  var costs = new Array();
  for (var i = 0; i <= s1.length; i++) {
    var lastValue = i;
    for (var j = 0; j <= s2.length; j++) {
      if (i == 0) costs[j] = j;
      else {
        if (j > 0) {
          var newValue = costs[j - 1];
          if (s1.charAt(i - 1) != s2.charAt(j - 1)) newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
          costs[j - 1] = lastValue; lastValue = newValue;
        }
      }
    }
    if (i > 0) costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}
