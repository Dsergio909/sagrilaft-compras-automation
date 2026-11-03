// =========================================================================
// MÓDULO 1: MOTOR DE ALERTAS CONTRACTUALES
// =========================================================================
// Avisa al equipo de compras de los contratos vencidos y de los que están
// por vencer. Cada aviso queda registrado en "Control Alertas Contratos"
// para no repetirlo. Correos, IDs y firma vienen de Config.gs.
// =========================================================================

function alertarContratosAviso() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Contratos");

  if (!hoja) {
    try { SpreadsheetApp.getUi().alert("❌ No existe la hoja 'Contratos'."); } catch(e) { Logger.log("❌ No existe la hoja 'Contratos'."); }
    return;
  }

  // Pestaña histórica de candados para evitar el SPAM
  var hojaControl = ss.getSheetByName("Control Alertas Contratos");
  if (!hojaControl) {
    hojaControl = ss.insertSheet("Control Alertas Contratos");
    hojaControl.appendRow(["Fecha Notificación", "Proveedor", "NIT", "Fecha Vencimiento", "Responsable", "Llave Control"]);
  }

  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  var datos = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getValues();
  var datosDisplay = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn()).getDisplayValues();
  var hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  var registrosControl = hojaControl.getDataRange().getDisplayValues();
  var llavesEnviadas = [];
  for (var c = 1; c < registrosControl.length; c++) {
    var ll = String(registrosControl[c][5]).trim();
    if (ll !== "") llavesEnviadas.push(ll);
  }

  var contratosVencidos = [];
  var contratosProximos = [];
  var nuevasLlavesParaRegistrar = [];

  var archivoFirma;
  try { archivoFirma = DriveApp.getFileById(CONFIG.ID_GIF_FIRMA_CONTRATOS); } catch (e) { archivoFirma = null; }

  for (var i = 0; i < datos.length; i++) {
    var proveedor = String(datos[i][1]).trim();
    var nit = String(datosDisplay[i][2]).trim();
    var objeto = String(datos[i][4]).trim();
    var fechaVenc = datos[i][9];
    var fechaVencVisual = String(datosDisplay[i][9]).trim();
    var fechaAviso = datos[i][12];
    var estado = String(datos[i][14]).trim();
    var responsable = String(datos[i][17]).trim();

    if (!proveedor || !fechaAviso || !fechaVenc) continue;
    if (estado !== "Activo") continue;
    if (!(fechaAviso instanceof Date) || !(fechaVenc instanceof Date)) continue;

    var aviso = new Date(fechaAviso);
    aviso.setHours(0, 0, 0, 0);

    // SÚPER CANDADO COMPUESTO (NIT + Objeto + Vencimiento)
    var nitLimpio = nit.replace(/[^0-9]/g, "") || proveedor.replace(/[^a-zA-Z0-9]/g, "");
    var objetoFiltro = objeto.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "").substring(0, 15);
    var llaveContrato = nitLimpio + "_" + objetoFiltro + "_" + fechaVencVisual;

    if (llavesEnviadas.includes(llaveContrato)) continue;

    if (hoy.getTime() >= aviso.getTime()) {
      var diasRestantes = Math.ceil((fechaVenc.getTime() - hoy.getTime()) / (1000 * 3600 * 24));

      var infoContrato = {
        proveedor: proveedor,
        objeto: objeto,
        fechaVenc: fechaVencVisual,
        dias: diasRestantes,
        responsable: responsable || "Sin Asignar"
      };

      if (diasRestantes < 0) {
        contratosVencidos.push(infoContrato);
      } else {
        contratosProximos.push(infoContrato);
      }

      nuevasLlavesParaRegistrar.push([
        Utilities.formatDate(hoy, "GMT-5", "dd/MM/yyyy HH:mm"),
        proveedor,
        nit,
        fechaVencVisual,
        responsable || "Sin Asignar",
        llaveContrato
      ]);
    }
  }

  if (contratosVencidos.length === 0 && contratosProximos.length === 0) {
    Logger.log("Bandeja limpia: No se detectaron nuevos vencimientos contractuales.");
    try { SpreadsheetApp.getUi().alert("ℹ️ No hay novedades pendientes de envío hoy (ya fueron registradas en el historial)."); } catch(e){}
    return;
  }

  // --- ESTRUCTURA HTML SOBRIA Y FORMAL ---
  var html = `
    <div style="font-family: Arial, sans-serif; font-size: 13px; color: #333333; line-height: 1.6;">
      <h2 style="color: #2c3e50; border-bottom: 1px solid #cfd8dc; padding-bottom: 8px; font-weight: normal; font-size: 18px;">REPORTE EJECUTIVO: Control de Vencimientos Contractuales</h2>
      <p>Estimado equipo del Departamento de Compras,</p>
      <p>A continuación se detalla el estado consolidado de los acuerdos comerciales vigentes que registran plazos próximos a vencimiento o requieren regularización legal:</p>
  `;

  if (contratosVencidos.length > 0) {
    html += `
      <h3 style="color: #333333; margin-top: 25px; text-transform: uppercase; font-size: 11px; letter-spacing: 1px; border-left: 3px solid #c62828; padding-left: 8px;">Contratos Expirados (Requiere Regularización)</h3>
      <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 12px; margin-top: 10px;">
        <thead>
          <tr style="background-color: #f5f7f8; color: #333333; text-align: left; font-weight: bold; border-bottom: 2px solid #cfd8dc;">
            <th style="padding: 10px; border-bottom: 1px solid #e0e0e0;">Proveedor / Objeto del Contrato</th>
            <th style="padding: 10px; border-bottom: 1px solid #e0e0e0; text-align: center;">Fecha Vencimiento</th>
            <th style="padding: 10px; border-bottom: 1px solid #e0e0e0; text-align: center;">Estado Actual</th>
            <th style="padding: 10px; border-bottom: 1px solid #e0e0e0;">Responsable</th>
          </tr>
        </thead>
        <tbody>
    `;
    contratosVencidos.forEach(function (a) {
      html += `
          <tr style="border-bottom: 1px solid #eee;">
            <td style="padding: 10px;"><span style="font-weight: bold; color: #111111;">${a.proveedor}</span><br><span style="font-size: 11px; color: #666666;">${a.objeto}</span></td>
            <td style="padding: 10px; text-align: center; white-space: nowrap; color: #555555;">${a.fechaVenc}</td>
            <td style="padding: 10px; text-align: center; color: #c62828; font-weight: bold; white-space: nowrap;">Vencido (-${Math.abs(a.dias)} días)</td>
            <td style="padding: 10px; color: #555555;">${a.responsable}</td>
          </tr>
      `;
    });
    html += `</tbody></table>`;
  }

  if (contratosProximos.length > 0) {
    html += `
      <h3 style="color: #333333; margin-top: 25px; text-transform: uppercase; font-size: 11px; letter-spacing: 1px; border-left: 3px solid #b0bec5; padding-left: 8px;">Contratos Próximos a Vencer (Periodo de Preaviso)</h3>
      <table style="width: 100%; border-collapse: collapse; margin-bottom: 25px; font-size: 12px; margin-top: 10px;">
        <thead>
          <tr style="background-color: #f5f7f8; color: #333333; text-align: left; font-weight: bold; border-bottom: 2px solid #cfd8dc;">
            <th style="padding: 10px; border-bottom: 1px solid #e0e0e0;">Proveedor / Objeto del Contrato</th>
            <th style="padding: 10px; border-bottom: 1px solid #e0e0e0; text-align: center;">Fecha Vencimiento</th>
            <th style="padding: 10px; border-bottom: 1px solid #e0e0e0; text-align: center;">Término Restante</th>
            <th style="padding: 10px; border-bottom: 1px solid #e0e0e0;">Responsable</th>
          </tr>
        </thead>
        <tbody>
    `;
    contratosProximos.forEach(function (a) {
      var colorDias = (a.dias <= 15) ? "color: #c62828; font-weight: bold;" : "color: #e65100; font-weight: 500;";
      html += `
          <tr style="border-bottom: 1px solid #eee;">
            <td style="padding: 10px;"><span style="font-weight: bold; color: #111111;">${a.proveedor}</span><br><span style="font-size: 11px; color: #666666;">${a.objeto}</span></td>
            <td style="padding: 10px; text-align: center; white-space: nowrap; color: #555555;">${a.fechaVenc}</td>
            <td style="padding: 10px; text-align: center; ${colorDias} white-space: nowrap;">${a.dias} días</td>
            <td style="padding: 10px; color: #555555;">${a.responsable}</td>
          </tr>
      `;
    });
    html += `</tbody></table>`;
  }

  html += `
      <p style="color: #546e7a; font-style: italic; margin-top: 25px; font-size: 12px;">Nota: Esta notificación automática consolida únicamente las modificaciones detectadas en el periodo actual. El histórico global interactivo está disponible en la pestaña 'DASHBOARD GENERAL'.</p>
    </div>

    <div style="font-family: Arial, sans-serif; margin-top: 35px; border-top: 1px solid #e0e0e0; padding-top: 15px;">
      <p style="color: #2e7d32; font-style: italic; margin: 0 0 10px 0; font-size: 13px;">${CONFIG.FIRMA.slogan}</p>
      <table cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td valign="middle" style="padding-right: 20px; width: 400px;">
            <img src="cid:firmaAnimada" width="400" style="width: 400px; display: block;" alt="Firma">
          </td>
          <td valign="middle">
            <div style="color: #d32f2f; font-weight: bold; font-size: 15px; margin-bottom: 2px;">${CONFIG.FIRMA.nombre}</div>
            <div style="color: #455a64; font-size: 13px; margin-bottom: 2px;">${CONFIG.FIRMA.cargo}</div>
            <div style="color: #1976d2; font-size: 12px; margin-bottom: 5px;">${CONFIG.FIRMA.correo}</div>
            <div style="color: #333333; font-size: 12px;">${CONFIG.FIRMA.direccion}</div>
            <div style="color: #333333; font-size: 12px;">${CONFIG.FIRMA.telefono}</div>
          </td>
        </tr>
      </table>
    </div>
  `;

  // --- CONFIGURACIÓN DE ENVÍO MASIVO AL EQUIPO DE COMPRAS ---
  GmailApp.sendEmail(CONFIG.CORREOS_EQUIPO_COMPRAS, "REPORTE EJECUTIVO – Control de Vencimientos de Contratos", "Contenido HTML.", {
    htmlBody: html,
    inlineImages: { firmaAnimada: archivoFirma ? archivoFirma.getBlob() : null }
  });

  nuevasLlavesParaRegistrar.forEach(function(filaControl) {
    hojaControl.appendRow(filaControl);
  });

  Logger.log("✅ Reporte formal masivo enviado con éxito al equipo de compras.");
  try { SpreadsheetApp.getUi().alert("✅ Correo formal sobrio enviado con éxito a todo el equipo de Compras."); } catch(e){}
}

// =========================================================================
// MÓDULO 2: DASHBOARD GENERAL
// =========================================================================

function generarDashboardContratos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaDatos = ss.getSheetByName("Contratos");

  if (!hojaDatos) {
    try { SpreadsheetApp.getUi().alert("Error: No existe la hoja 'Contratos'"); } catch(e){}
    return;
  }

  var hojaDash = ss.getSheetByName("DASHBOARD GENERAL");
  if (!hojaDash) {
    hojaDash = ss.insertSheet("DASHBOARD GENERAL", 0);
  } else {
    hojaDash.getRange("A1:N50").clear();
  }

  var datos = hojaDatos.getRange(2, 1, hojaDatos.getLastRow() - 1, hojaDatos.getLastColumn()).getValues();

  var totalActivos = 0;
  var totalVencidos = 0;
  var totalProximos30 = 0;
  var totalProximos60 = 0;

  var responsables = {};
  var listaPlanAccion = [];
  var hoy = new Date();
  hoy.setHours(0,0,0,0);

  datos.forEach(function (fila) {
    var proveedor = fila[1];
    if (!proveedor) return;

    var fechaVenc = fila[9];
    var estado = String(fila[14]).trim();
    var responsable = String(fila[17]).trim() || "Sin Asignar";
    var objeto = String(fila[4]).trim();

    if (estado === "Activo") {
      totalActivos++;

      if (!responsables[responsable]) responsables[responsable] = 0;
      responsables[responsable]++;

      if (fechaVenc instanceof Date) {
        var diasRestantes = Math.ceil((fechaVenc.getTime() - hoy.getTime()) / (1000 * 3600 * 24));
        var alerta = "";
        var accionRecomendada = "";

        if (diasRestantes < 0) {
          totalVencidos++;
          alerta = "Expirado";
          accionRecomendada = "Revisión urgente / Comité Legal";
        } else if (diasRestantes <= 30) {
          totalProximos30++;
          alerta = "Crítico (0-30 d)";
          accionRecomendada = "Firmar Prórroga de Inmediato";
        } else if (diasRestantes <= 60) {
          totalProximos60++;
          alerta = "Preventivo (31-60 d)";
          accionRecomendada = "Contactar para renegociación";
        }

        if (diasRestantes <= 60) {
          listaPlanAccion.push({
            proveedor: proveedor,
            vencimiento: Utilities.formatDate(fechaVenc, "GMT-5", "dd/MM/yyyy"),
            dias: diasRestantes,
            alerta: alerta,
            accion: accionRecomendada,
            responsable: responsable
          });
        }
      }
    }
  });

  listaPlanAccion.sort(function(a, b) { return a.dias - b.dias; });
  hojaDash.setHiddenGridlines(true);

  hojaDash.getRange("B2:M2").merge().setBackground("#2c3e50").setValue("  TORRE DE CONTROL DE CONTRATOS - MATRIZ DE COMPRAS").setFontSize(13).setFontColor("white").setFontWeight("bold").setVerticalAlignment("middle");
  hojaDash.getRange("B3").setValue("Sincronización del Sistema: " + Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy HH:mm")).setFontSize(9).setFontColor("#78909c").setFontStyle("italic");

  var bordesGris = "#cfd8dc";

  hojaDash.getRange("B5:C5").merge().setValue("CONTRATOS ACTIVOS").setFontSize(8).setFontColor("#546e7a").setFontWeight("bold").setBackground("#f8f9fa").setHorizontalAlignment("center");
  hojaDash.getRange("B6:C7").merge().setValue(totalActivos).setFontSize(22).setFontWeight("bold").setBackground("#ffffff").setFontColor("#2c3e50").setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setBorder(true, true, true, true, false, false, bordesGris, SpreadsheetApp.BorderStyle.SOLID);

  hojaDash.getRange("E5:F5").merge().setValue("YA VENCIDOS").setFontSize(8).setFontColor("#546e7a").setFontWeight("bold").setBackground("#f8f9fa").setHorizontalAlignment("center");
  hojaDash.getRange("E6:F7").merge().setValue(totalVencidos).setFontSize(22).setFontWeight("bold").setBackground("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setBorder(true, true, true, true, false, false, bordesGris, SpreadsheetApp.BorderStyle.SOLID);
  if (totalVencidos > 0) { hojaDash.getRange("E6:F7").setFontColor("#c62828"); } else { hojaDash.getRange("E6:F7").setFontColor("#37474f"); }

  hojaDash.getRange("H5:I5").merge().setValue("CRÍTICOS (0-30 D)").setFontSize(8).setFontColor("#546e7a").setFontWeight("bold").setBackground("#f8f9fa").setHorizontalAlignment("center");
  hojaDash.getRange("H6:I7").merge().setValue(totalProximos30).setFontSize(22).setFontWeight("bold").setBackground("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setBorder(true, true, true, true, false, false, bordesGris, SpreadsheetApp.BorderStyle.SOLID);
  if (totalProximos30 > 0) { hojaDash.getRange("H6:I7").setFontColor("#e65100"); } else { hojaDash.getRange("H6:I7").setFontColor("#37474f"); }

  hojaDash.getRange("K5:L5").merge().setValue("PREVENTIVOS (31-60 D)").setFontSize(8).setFontColor("#546e7a").setFontWeight("bold").setBackground("#f8f9fa").setHorizontalAlignment("center");
  hojaDash.getRange("K6:L7").merge().setValue(totalProximos60).setFontSize(22).setFontWeight("bold").setBackground("#ffffff").setFontColor("#37474f").setHorizontalAlignment("center").setVerticalAlignment("middle")
          .setBorder(true, true, true, true, false, false, bordesGris, SpreadsheetApp.BorderStyle.SOLID);

  hojaDash.getRange("B10").setValue("Control de Distribución (Activos)").setFontWeight("bold").setFontColor("#2c3e50").setFontSize(10);
  hojaDash.getRange("B11:C11").setValues([["Responsable Compras", "Volumen"]]).setBackground("#37474f").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");

  var filaR = 12;
  for (var res in responsables) {
    hojaDash.getRange("B" + filaR).setValue(res).setFontColor("#333");
    hojaDash.getRange("C" + filaR).setValue(responsables[res]).setHorizontalAlignment("center").setFontColor("#333");
    filaR++;
  }
  if(filaR > 12) hojaDash.getRange("B11:C" + (filaR - 1)).setBorder(true, true, true, true, true, true, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);

  hojaDash.getRange("E10").setValue("Plan de Acción Contractual (Alertas Vigentes)").setFontWeight("bold").setFontColor("#2c3e50").setFontSize(10);
  hojaDash.getRange("E11:I11").setValues([["Proveedor", "Vencimiento", "Término", "Clasificación Alerta", "Acción Recomendada Operativa"]]).setBackground("#2c3e50").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");

  var filaT = 12;
  for (var k = 0; k < listaPlanAccion.length; k++) {
    var p = listaPlanAccion[k];
    hojaDash.getRange("E" + filaT).setValue(p.proveedor).setFontWeight("bold").setFontColor("#333333");
    hojaDash.getRange("F" + filaT).setValue(p.vencimiento).setHorizontalAlignment("center").setFontColor("#555");

    if (p.dias < 0) {
      hojaDash.getRange("G" + filaT).setValue("Vencido").setFontColor("#c62828").setFontWeight("bold").setHorizontalAlignment("center");
      hojaDash.getRange("H" + filaT).setValue(p.alerta).setFontColor("#c62828").setFontWeight("bold").setHorizontalAlignment("center");
    } else {
      hojaDash.getRange("G" + filaT).setValue(p.dias + " días").setFontColor("#333").setHorizontalAlignment("center");
      hojaDash.getRange("H" + filaT).setValue(p.alerta).setHorizontalAlignment("center");
      if (p.dias <= 30) {
        hojaDash.getRange("G" + filaT).setFontWeight("bold").setFontColor("#e65100");
        hojaDash.getRange("H" + filaT).setFontColor("#e65100").setFontWeight("bold");
      } else {
        hojaDash.getRange("H" + filaT).setFontColor("#546e7a");
      }
    }

    hojaDash.getRange("I" + filaT).setValue(p.accion).setFontColor("#455a64");
    filaT++;
  }
  if(filaT > 12) {
    hojaDash.getRange("E11:I" + (filaT - 1)).setBorder(true, true, true, true, true, true, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
    for (var f = 12; f < filaT; f++) {
      if (f % 2 === 0) hojaDash.getRange("E" + f + ":I" + f).setBackground("#fcfdfe");
    }
  }

  hojaDash.setColumnWidth(1, 15);   hojaDash.setColumnWidth(2, 180);  hojaDash.setColumnWidth(3, 80);
  hojaDash.setColumnWidth(4, 25);   hojaDash.setColumnWidth(5, 220);  hojaDash.setColumnWidth(6, 100);
  hojaDash.setColumnWidth(7, 95);   hojaDash.setColumnWidth(8, 140);  hojaDash.setColumnWidth(9, 250);

  try { SpreadsheetApp.getUi().alert("📊 Panel de Control actualizado con éxito bajo el formato corporativo serio."); } catch(e){}
}

// =========================================================================
// UTILIDAD: MIGRACIÓN MASIVA DE ÁREAS SOLICITANTES
// =========================================================================
// Estandariza los nombres de las áreas solicitantes (columna P) con un
// diccionario: nombre tal como aparece → nombre estandarizado.

function migrarAreasContratosMasivo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Contratos");

  if (!hoja) return SpreadsheetApp.getUi().alert("❌ Error: No se encontró la pestaña 'Contratos'.");

  var datos = hoja.getDataRange().getValues();

  // Coordenada exacta (Columna P)
  var idxArea = 15;

  // Diccionario de ejemplo: agrega una entrada por cada variante
  var diccionario = {
    "Nombre de área tal como aparece hoy": "Código y nombre estandarizado",
    "Otra variante del mismo área": "Código y nombre estandarizado"
    // ... agrega aquí tantas entradas como áreas distintas tengas
  };

  var cambiosRealizados = 0;

  for (var i = 1; i < datos.length; i++) {
    var areaAntigua = String(datos[i][idxArea]).trim();
    if (areaAntigua === "") continue;

    var nuevaArea = areaAntigua;

    if (diccionario[areaAntigua]) {
      nuevaArea = diccionario[areaAntigua];
    }

    if (nuevaArea !== areaAntigua) {
      datos[i][idxArea] = nuevaArea;
      cambiosRealizados++;
    }
  }

  var valoresNuevos = [];
  for (var j = 1; j < datos.length; j++) {
    valoresNuevos.push([datos[j][idxArea]]);
  }

  var rangoColumnaArea = hoja.getRange(2, 16, valoresNuevos.length, 1);
  rangoColumnaArea.clearDataValidations();
  rangoColumnaArea.setValues(valoresNuevos);

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert("✅ ACTUALIZACIÓN COMPLETADA EN CONTRATOS\n\nSe estandarizaron " + cambiosRealizados + " áreas solicitantes.");
}
