// ============================================================================
// Code.gs COMPLETO Y LIMPIO -- reemplaza TODO el contenido de tu Code.gs
// en el editor de Apps Script por el contenido de este archivo (selecciona
// todo con Ctrl+A dentro del editor, borra, y pega esto). No hay funciones
// duplicadas ni anidadas por accidente, asi que no deberia volver a pasar
// el problema de "include is not defined".
//
// Despues de pegar: Implementar > Gestionar implementaciones > editar tu
// implementacion > Version: Nueva version > Implementar.
// ============================================================================

var SHEET_URL = "https://docs.google.com/spreadsheets/d/10NRfnS9iwSbLp-UXSMimfmCXAyiIxQ9eud78owsbxZE/edit";
var CARPETA_DIBUJOS = "teamcenter_bot";
var PLANTILLA_RFQ_XLSX_BASE64 = "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";

// ID del archivo PNG del logo en Drive.
var LOGO_FILE_ID = '1Yim1270CD2TWLaKy_JR28GrL5L7CZFnR';

// Version vieja: embebe la imagen como base64 dentro del HTML. Se dejo de
// usar para el logo del Dashboard porque el archivo pesa ~2MB y, como
// texto base64 duplicado en las dos vistas (escritorio + app), estaba
// inflando la pagina varios MB y haciendo que fallara la carga (sobre
// todo en celular). Se deja disponible por si se necesita en otro lado
// (ej. un correo, donde no se puede usar una URL externa).
function getLogoDataUri_() {
  try {
    var logoFile = DriveApp.getFileById(LOGO_FILE_ID);
    var blob = logoFile.getBlob();
    var contentType = blob.getContentType() || 'image/png';
    var base64 = Utilities.base64Encode(blob.getBytes());
    return 'data:' + contentType + ';base64,' + base64;
  } catch (e) {
    Logger.log('getLogoDataUri_ error: ' + e.message);
    return '';
  }
}

// Logo del Dashboard: en vez de embeber la imagen, se sirve como una URL
// normal de Drive (el navegador la descarga aparte, no infla el HTML).
// Se asegura que el archivo sea visible por link para que cargue sin pedir
// autenticacion.
function getLogoUrl_() {
  try {
    var logoFile = DriveApp.getFileById(LOGO_FILE_ID);
    if (logoFile.getSharingAccess() === DriveApp.Access.PRIVATE) {
      logoFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    return "https://drive.google.com/thumbnail?id=" + LOGO_FILE_ID + "&sz=w400";
  } catch (e) {
    Logger.log('getLogoUrl_ error: ' + e.message);
    return '';
  }
}

var MAPA_SOURCING = {
  "ggarcia@criotec.com.mx": "Graciela",
  "pbersoza@criotec.com.mx": "Paola",
  "sbautista@criotec.com.mx": "Santiago",
  "cruiz@criotec.com.mx": "Carolina"
};
var CORREO_JEFE = "hpaez@criotec.com.mx";
var NOMBRE_JEFE = "Hugo"; // nombre para mostrarlo en listas (ej. asignables de tickets)
var CORREO_RESPONSABLE_NA = "pbersoza@criotec.com.mx";
var CORREO_ALERTAS = "pbersoza@criotec.com.mx";

function ingles(texto) {
  return '<i style="color:#1a56db;">' + texto + '</i>';
}
function lineaBilingue(esp, eng) {
  return esp + " / " + ingles(eng);
}
function obtenerCorreoPorNombreSourcing(nombre) {
  var nombreBuscado = String(nombre || "").trim().toLowerCase();
  for (var correo in MAPA_SOURCING) {
    if (MAPA_SOURCING[correo].trim().toLowerCase() === nombreBuscado) return correo;
  }
  return null;
}

function testLogoActual() {
  var uri = getLogoDataUri_();
  Logger.log("Longitud del data URI: " + uri.length);
  Logger.log("Primeros 80 caracteres: " + uri.substring(0, 80));
}

function notificarAsignacionNueva(sheet, filaNueva) {
  var sourcingAsignado = sheet.getRange(filaNueva, 13).getValue(); // col M
  if (!sourcingAsignado) return;

  var correo = obtenerCorreoPorNombreSourcing(sourcingAsignado);
  if (!correo) return; // nombre no reconocido, no truena, solo no notifica

  var codigo = sheet.getRange(filaNueva, 5).getValue();      // col E
  var revision = sheet.getRange(filaNueva, 6).getValue();    // col F
  var descripcion = sheet.getRange(filaNueva, 8).getValue(); // col H
  var area = sheet.getRange(filaNueva, 10).getValue();       // col J
  var commodity = sheet.getRange(filaNueva, 11).getValue();  // col K

  var asunto = "Nueva tarea asignada: " + codigo + " Rev." + revision;
  var cuerpo =
    "Te llego una tarea nueva de Team Center ya asignada a ti en Seguimiento de Tareas:\n\n" +
    "Codigo: " + codigo + "\n" +
    "Revision: " + revision + "\n" +
    "Descripcion: " + (descripcion || "-") + "\n" +
    "Area: " + (area || "-") + "\n" +
    "Commodity: " + (commodity || "-") + "\n\n" +
    "Puedes verla en el dashboard de Team Center.";

  try {
    GmailApp.sendEmail(correo, asunto, cuerpo);
  } catch (err) {
    Logger.log("No se pudo notificar la asignacion nueva: " + err.message);
    notificarError("notificarAsignacionNueva", err.message + "\n\n" + (err.stack || ""));
  }
}

// Lee la hoja "Proveedores" y arma un mapa
// { PREFIJO: { responsable: "...", commodity: "..." } } a partir de la
// columna A (Nomenclatura), columna D (Commodity) y columna E
// (Responsable). Como una misma nomenclatura puede repetirse en varias
// filas de proveedores distintos (cada uno con su propio Commodity/
// Responsable capturado), cada dato se asigna al que aparezca MAS VECES
// para esa nomenclatura (mayoria), no al de la ultima fila leida.
function obtenerMapaPrefijos_() {
  var conteoResponsable = {}; // { prefijo: { responsable: cantidadDeFilas } }
  var conteoCommodity = {};   // { prefijo: { commodity: cantidadDeFilas } }
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var hoja = ss.getSheetByName("Proveedores");
  if (!hoja) return {};

  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var prefijo = String(datos[i][0] || "").trim().toUpperCase(); // col A: Nomenclatura
    var commodity = String(datos[i][3] || "").trim();             // col D: Commodity
    var responsable = String(datos[i][4] || "").trim();           // col E: Responsable
    if (!prefijo) continue;

    if (responsable) {
      if (!conteoResponsable[prefijo]) conteoResponsable[prefijo] = {};
      conteoResponsable[prefijo][responsable] = (conteoResponsable[prefijo][responsable] || 0) + 1;
    }
    if (commodity) {
      if (!conteoCommodity[prefijo]) conteoCommodity[prefijo] = {};
      conteoCommodity[prefijo][commodity] = (conteoCommodity[prefijo][commodity] || 0) + 1;
    }
  }

  function masFrecuentePorPrefijo(conteoPorPrefijo) {
    var mapa = {};
    Object.keys(conteoPorPrefijo).forEach(function(prefijo) {
      var conteos = conteoPorPrefijo[prefijo];
      var mejorValor = null;
      var mejorConteo = 0;
      Object.keys(conteos).forEach(function(valor) {
        if (conteos[valor] > mejorConteo) {
          mejorConteo = conteos[valor];
          mejorValor = valor;
        }
      });
      mapa[prefijo] = mejorValor;
    });
    return mapa;
  }

  var mapaResponsables = masFrecuentePorPrefijo(conteoResponsable);
  var mapaCommodities = masFrecuentePorPrefijo(conteoCommodity);

  var resultado = {};
  Object.keys(mapaResponsables).forEach(function(prefijo) {
    if (!resultado[prefijo]) resultado[prefijo] = {};
    resultado[prefijo].responsable = mapaResponsables[prefijo];
  });
  Object.keys(mapaCommodities).forEach(function(prefijo) {
    if (!resultado[prefijo]) resultado[prefijo] = {};
    resultado[prefijo].commodity = mapaCommodities[prefijo];
  });
  return resultado;
}

// Al llegar una tarea nueva, toma los primeros 3 caracteres del codigo de
// pieza (columna E de la hoja principal) y, segun la hoja "Proveedores",
// escribe el Commodity (columna K) y el Sourcing (columna M) que le
// corresponden a ese prefijo. Siempre sobreescribe esas dos celdas (aunque
// ya tengan una formula o un valor previo) porque la fuente de verdad para
// esos dos campos es la pestana Proveedores, no lo que haya en la hoja
// principal. Si el prefijo no esta dado de alta en Proveedores, no toca
// nada (se deja lo que ya hubiera ahi).
function asignarSourcingPorPrefijo(sheet, filaNueva) {
  var codigo = sheet.getRange(filaNueva, 5).getValue(); // col E
  if (!codigo) return;

  var prefijo = String(codigo).trim().toUpperCase().substring(0, 3);
  if (!prefijo) return;

  var info = obtenerMapaPrefijos_()[prefijo];
  if (!info) {
    Logger.log("asignarSourcingPorPrefijo: prefijo '" + prefijo + "' (codigo " + codigo + ") no esta dado de alta en Proveedores.");
    return;
  }

  if (info.commodity) {
    sheet.getRange(filaNueva, 11).setValue(info.commodity); // col K: Commodity
  }
  if (info.responsable) {
    sheet.getRange(filaNueva, 13).setValue(info.responsable); // col M: Sourcing
  }
  Logger.log(
    "Fila " + filaNueva + " (codigo " + codigo + ", prefijo " + prefijo + ") actualizada -> " +
    "Commodity: " + (info.commodity || "(sin dato)") + ", Sourcing: " + (info.responsable || "(sin dato)")
  );
}

function notificarError(origen, detalle) {
  try {
    var asunto = "⚠️ Error en automatizacion RFQ: " + origen;
    var cuerpo = "Se detecto un error en el proceso automatizado de RFQ's.\n\n" +
                 "Origen: " + origen + "\n" +
                 "Fecha/hora: " + new Date().toLocaleString("es-MX") + "\n\n" +
                 "Detalle:\n" + detalle;
    GmailApp.sendEmail(CORREO_ALERTAS, asunto, cuerpo);
  } catch (errAlerta) {
    Logger.log("No se pudo enviar la alerta de error: " + errAlerta.message);
  }
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];
    var data = JSON.parse(e.postData.contents);

    var colD = sheet.getRange("D1:D").getValues();
    var ultimaFila = 0;
    for (var i = 0; i < colD.length; i++) {
      if (colD[i][0] !== "") ultimaFila = i + 1;
    }
    var filaNueva = ultimaFila + 1;

    sheet.getRange(filaNueva, 3).setValue(data.fecha);
    sheet.getRange(filaNueva, 4).setValue(data.nombreArchivo);
    sheet.getRange(filaNueva, 9).setValue(data.iniciador);

    SpreadsheetApp.flush();

    try {
      cancelarRevisionesAnteriores(sheet, filaNueva);
    } catch (err) {
      Logger.log("Error al intentar cancelar revisiones anteriores: " + err.message);
      notificarError("doPost - cancelarRevisionesAnteriores", err.message + "\n\n" + (err.stack || ""));
    }

    try {
      asignarSourcingPorPrefijo(sheet, filaNueva);
      notificarAsignacionNueva(sheet, filaNueva);
    } catch (err) {
      Logger.log("Error al asignar/notificar sourcing por prefijo: " + err.message);
      notificarError("doPost - asignarSourcingPorPrefijo", err.message + "\n\n" + (err.stack || ""));
    }

    return ContentService.createTextOutput(JSON.stringify({status: 'ok'})).setMimeType(ContentService.MimeType.JSON);
  } catch (errGeneral) {
    Logger.log("Error critico en doPost: " + errGeneral.message);
    notificarError(
      "doPost",
      errGeneral.message + "\n\n" + (errGeneral.stack || "") +
      "\n\nPayload recibido: " + (e && e.postData ? e.postData.contents : "(sin datos)")
    );
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: errGeneral.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function tieneRFQConProveedor(codigo, revision) {
  var hojaLog = obtenerHojaLogRFQ();
  var datos = hojaLog.getDataRange().getValues();
  var key = String(codigo).trim().toUpperCase() + "|" + String(revision).trim().toUpperCase();
  for (var i = 1; i < datos.length; i++) {
    var keyFila = String(datos[i][2]).trim().toUpperCase() + "|" + String(datos[i][3]).trim().toUpperCase();
    if (keyFila === key) return true; // ya se le mando RFQ a algun proveedor para esta pieza+revision
  }
  return false;
}

function cancelarRevisionesAnteriores(sheet, filaNueva) {
  var codigoNuevo = sheet.getRange(filaNueva, 5).getValue();
  var revisionNueva = sheet.getRange(filaNueva, 6).getValue();
  if (!codigoNuevo) return;
  if (filaNueva <= 2) return;

  var data = sheet.getRange(2, 1, filaNueva - 2, 15).getValues();

  for (var i = 0; i < data.length; i++) {
    var filaReal = i + 2;
    var codigo = data[i][4];
    var revision = data[i][5];
    var estado = data[i][14];

    if (!codigo) continue;
    if (String(codigo).trim().toUpperCase() !== String(codigoNuevo).trim().toUpperCase()) continue;
    if (String(revision).trim().toUpperCase() === String(revisionNueva).trim().toUpperCase()) continue;
    if (!tareaSigueAbierta(estado)) continue;

    if (tieneRFQConProveedor(codigo, revision)) {
      Logger.log("Fila " + filaReal + " (" + codigo + " Rev." + revision + ") ya tiene RFQ con proveedor: se deja para cancelacion manual.");
      continue;
    }

    sheet.getRange(filaReal, 15).setValue("Cancelada (nueva revision)");
    cancelarRFQsDeLaPieza(codigo, revision);
    Logger.log("Fila " + filaReal + " (" + codigo + " Rev." + revision + ") cancelada automaticamente por llegar Rev." + revisionNueva);
  }
}

function tareaSigueAbierta(estado) {
  if (!estado) return true;
  var e = String(estado).trim().toLowerCase();
  return (e !== "cerrada" && e !== "cerrado" && e !== "completada" && e !== "completado" &&
          e.indexOf("cancelada") === -1);
}

function cancelarRFQsDeLaPieza(codigo, revision) {
  var hojaLog = obtenerHojaLogRFQ();
  var datos = hojaLog.getDataRange().getValues();
  var key = String(codigo).trim().toUpperCase() + "|" + String(revision).trim().toUpperCase();
  for (var i = 1; i < datos.length; i++) {
    var keyFila = String(datos[i][2]).trim().toUpperCase() + "|" + String(datos[i][3]).trim().toUpperCase();
    if (keyFila === key && datos[i][9] === "Enviado") {
      hojaLog.getRange(i + 1, 10).setValue("Cancelado");
    }
  }
}

function marcarCerrada(filaIndex) {
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];
  sheet.getRange(filaIndex, 15).setValue("Cerrada");
  return {status: 'ok'};
}

function guardarMuestras(filaIndex, valor) {
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];
  sheet.getRange(filaIndex, 14).setValue(valor);
  return {status: 'ok'};
}

function guardarMotivoRevision(filaIndex, valor) {
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];
  sheet.getRange(filaIndex, 16).setValue(valor);
  return {status: 'ok'};
}

function buscarArchivoPorNombre(carpeta, nombreBuscado) {
  var archivos = carpeta.getFilesByName(nombreBuscado);
  if (archivos.hasNext()) return archivos.next();

  var subcarpetas = carpeta.getFolders();
  while (subcarpetas.hasNext()) {
    var sub = subcarpetas.next();
    var archivosSub = sub.getFilesByName(nombreBuscado);
    if (archivosSub.hasNext()) return archivosSub.next();
  }
  return null;
}

function generarExcelCotizacion(codigo, revision, descripcion, muestras) {
  var plantillaBlob = Utilities.newBlob(
    Utilities.base64Decode(PLANTILLA_RFQ_XLSX_BASE64),
    MimeType.MICROSOFT_EXCEL,
    "plantilla.xlsx"
  );

  var archivoTemp = Drive.Files.create(
    { name: "RFQ_" + codigo + "_" + revision + "_" + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS },
    plantillaBlob
  );

  var ss = SpreadsheetApp.openById(archivoTemp.id);
  var sheet = ss.getSheets()[0];

  sheet.getRange("D3").setValue("RFQ-" + codigo + "-" + revision);
  sheet.getRange("C6").setValue(codigo);
  sheet.getRange("C7").setValue(revision);
  sheet.getRange("C8").setValue(descripcion);
  sheet.getRange("C9").setValue(muestras || "");
  sheet.getRange("C10").setValue(new Date());

  SpreadsheetApp.flush();

  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  });
  var blob = response.getBlob().setName("RFQ_" + codigo + "_" + revision + ".xlsx");

  DriveApp.getFileById(ss.getId()).setTrashed(true);

  return blob;
}

function obtenerHojaLogRFQ() {
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var hoja = ss.getSheetByName("Log RFQ");
  if (!hoja) {
    hoja = ss.insertSheet("Log RFQ");
    hoja.appendRow([
      "RFQ ID", "Fecha de Envio", "Codigo", "Revision", "Descripcion",
      "Proveedor", "Contacto", "Correo", "Muestras", "Estado",
      "Fecha de Respuesta", "Precio Unitario", "Moneda", "MOQ",
      "Tiempo de Entrega", "Comentarios",
      "Ultimo Recordatorio", "Recordatorios Enviados", "EAU", "Enviado Por"
    ]);
    hoja.getRange(1, 1, 1, 20).setFontWeight("bold");
    return hoja;
  }

  if (hoja.getLastColumn() < 18) {
    hoja.getRange(1, 17, 1, 2).setValues([["Ultimo Recordatorio", "Recordatorios Enviados"]]);
    hoja.getRange(1, 17, 1, 2).setFontWeight("bold");
  }
  if (hoja.getRange(1, 16).getValue() === "Vigencia de Cotizacion") {
    hoja.getRange(1, 16).setValue("Comentarios");
  }
  if (hoja.getLastColumn() < 19) {
    hoja.getRange(1, 19).setValue("EAU");
    hoja.getRange(1, 19).setFontWeight("bold");
  }
  if (hoja.getLastColumn() < 20) {
    hoja.getRange(1, 20).setValue("Enviado Por");
    hoja.getRange(1, 20).setFontWeight("bold");
  }
  return hoja;
}

function obtenerCotizacionesPorPieza() {
  var hojaLog = obtenerHojaLogRFQ();
  var datos = hojaLog.getDataRange().getValues();
  var mapa = {};

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var codigo = fila[2];
    var revision = fila[3];
    if (!codigo) continue;

    var key = String(codigo).trim().toUpperCase() + "|" + String(revision).trim().toUpperCase();
    if (!mapa[key]) mapa[key] = [];

    var fechaResp = fila[10];
    var comentarios = fila[15];

    mapa[key].push({
      rfqId: fila[0],
      proveedor: fila[5],
      contacto: fila[6],
      correo: fila[7],
      estado: fila[9],
      fechaRespuesta: (fechaResp instanceof Date) ? fechaResp.toISOString() : fechaResp,
      precio: fila[11],
      moneda: fila[12],
      moq: fila[13],
      entrega: fila[14],
      comentarios: comentarios,
      eau: fila[18]
    });
  }
  return mapa;
}

function cancelarRFQ(rfqId, motivo) {
  if (!motivo || !String(motivo).trim()) {
    throw new Error("Debes indicar el motivo de la cancelacion.");
  }
  var hojaLog = obtenerHojaLogRFQ();
  var datos = hojaLog.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === rfqId) {
      var filaReal = i + 1;
      var comentarioExistente = String(datos[i][15] || "").trim(); // col P: Comentarios
      var notaCancelacion = "Cancelado: " + motivo;
      var comentarioFinal = comentarioExistente ? (notaCancelacion + " | " + comentarioExistente) : notaCancelacion;
      hojaLog.getRange(filaReal, 10).setValue("Cancelado");
      hojaLog.getRange(filaReal, 16).setValue(comentarioFinal);
      return {status: 'ok', comentarios: comentarioFinal};
    }
  }
  return {status: 'error', message: 'No se encontro el RFQ ' + rfqId + ' en el Log RFQ.'};
}

// Reintenta encontrar la respuesta de un RFQ puntual, para cuando el
// escaneo automatico (revisarRespuestasRFQ) no lo detecto. A diferencia
// del escaneo automatico, esta busqueda es mas flexible: no exige que el
// remitente coincida exacto con el correo del proveedor (por si respondio
// desde otra cuenta), no se restringe a hilos sin la etiqueta
// "RFQ-Procesado", y busca hasta 90 dias atras en vez de 30. Solo exige
// que el asunto tenga el "[ID]" exacto y que el mensaje traiga un .xlsx.
function reintentarDeteccionRFQ(rfqId) {
  var hojaLog = obtenerHojaLogRFQ();
  var datos = hojaLog.getDataRange().getValues();
  var filaLog = null;
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][0] === rfqId) { filaLog = i + 1; break; }
  }
  if (!filaLog) {
    return { status: 'error', message: 'No se encontro el RFQ ' + rfqId + ' en el Log RFQ.' };
  }

  var NOMBRE_ETIQUETA = "RFQ-Procesado";
  var etiqueta = GmailApp.getUserLabelByName(NOMBRE_ETIQUETA);
  if (!etiqueta) etiqueta = GmailApp.createLabel(NOMBRE_ETIQUETA);

  var hilos = GmailApp.search('has:attachment newer_than:90d "[' + rfqId + ']"');
  Logger.log("reintentarDeteccionRFQ: " + hilos.length + " hilo(s) encontrados para RFQ " + rfqId);

  for (var h = 0; h < hilos.length; h++) {
    var mensajes = hilos[h].getMessages();
    for (var m = 0; m < mensajes.length; m++) {
      var msg = mensajes[m];
      if (msg.getSubject().indexOf("[" + rfqId + "]") === -1) continue;

      var adjuntos = msg.getAttachments();
      var adjuntoExcel = null;
      for (var a = 0; a < adjuntos.length; a++) {
        if (/\.xlsx$/i.test(adjuntos[a].getName())) { adjuntoExcel = adjuntos[a]; break; }
      }
      if (!adjuntoExcel) continue;

      try {
        var datosExcel = leerDatosDeExcelAdjunto(adjuntoExcel);
        var estadoFinal = datosExcel.precio ? "Respondido" : "Respondido incompleto";
        hojaLog.getRange(filaLog, 10).setValue(estadoFinal);
        hojaLog.getRange(filaLog, 11).setValue(new Date());
        hojaLog.getRange(filaLog, 12).setValue(datosExcel.precio);
        hojaLog.getRange(filaLog, 13).setValue(datosExcel.moneda);
        hojaLog.getRange(filaLog, 14).setValue(datosExcel.moq);
        hojaLog.getRange(filaLog, 15).setValue(datosExcel.entrega);
        hojaLog.getRange(filaLog, 16).setValue(datosExcel.comentarios);
        hilos[h].addLabel(etiqueta);
        Logger.log("reintentarDeteccionRFQ: RFQ " + rfqId + " actualizado a " + estadoFinal + ".");
        return {
          status: 'ok', estado: estadoFinal, precio: datosExcel.precio, moneda: datosExcel.moneda,
          moq: datosExcel.moq, entrega: datosExcel.entrega, comentarios: datosExcel.comentarios,
          fechaRespuesta: new Date().toISOString()
        };
      } catch (err) {
        Logger.log("reintentarDeteccionRFQ: error leyendo el excel: " + err.message);
        return { status: 'error', message: 'Se encontro un correo con el ID pero no se pudo leer el Excel: ' + err.message };
      }
    }
  }

  return { status: 'error', message: 'No se encontro ningun correo con "[' + rfqId + ']" en el asunto y un adjunto .xlsx en los ultimos 90 dias.' };
}

function reasignarTarea(filaIndex, nuevoResponsable) {
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];
  sheet.getRange(filaIndex, 13).setValue(nuevoResponsable);
  return {status: 'ok'};
}

function enviarCotizacion(filaIndex, codigo, revision, descripcion, muestras, proveedoresSeleccionados, eau) {
  var enviadoPor = Session.getActiveUser().getEmail() || "";
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];
  var nombreArchivo = sheet.getRange(filaIndex, 4).getValue();

  var carpetas = DriveApp.getFoldersByName(CARPETA_DIBUJOS);
  var archivoDibujo = null;
  if (carpetas.hasNext()) {
    var carpetaRaiz = carpetas.next();
    archivoDibujo = buscarArchivoPorNombre(carpetaRaiz, nombreArchivo);
  }

  var excelBlob = generarExcelCotizacion(codigo, revision, descripcion, muestras);
  var hojaLog = obtenerHojaLogRFQ();

  var lineaMuestrasPlano = muestras ? ("Muestras Requeridas / Samples Required: " + muestras + "\n") : "";
  var lineaEAUPlano = eau ? ("Volumen Anual Estimado / Estimated Annual Usage (EAU): " + eau + "\n") : "";

  var lineaMuestrasHtml = muestras
    ? (lineaBilingue("Muestras Requeridas", "Samples Required") + ": " + muestras + "<br>")
    : "";
  var lineaEAUHtml = eau
    ? (lineaBilingue("Volumen Anual Estimado", "Estimated Annual Usage (EAU)") + ": " + eau + "<br>")
    : "";

  var enviadosA = [];

  for (var i = 0; i < proveedoresSeleccionados.length; i++) {
    var prov = proveedoresSeleccionados[i];
    var rfqId = codigo + "-" + revision + "-" + Utilities.getUuid().substring(0, 6);

    var asunto = "RFQ - " + codigo + " Rev. " + revision + " [" + rfqId + "]";

    var cuerpo = "Buen dia / Good day,\n\n" +
                 "Favor de cotizar la siguiente pieza. Adjunto encontraras el dibujo tecnico y un formato de cotizacion; favor de llenarlo y regresarlo junto con tu respuesta (respondiendo a este mismo correo).\n" +
                 "Please quote the following part. Attached you will find the technical drawing and a quotation form; please fill it out and return it along with your response (by replying to this same email).\n\n" +
                 "Numero de Parte / Part Number: " + codigo + "\n" +
                 "Revision / Revision: " + revision + "\n" +
                 "Descripcion / Description: " + descripcion + "\n" +
                 lineaMuestrasPlano +
                 lineaEAUPlano + "\n" +
                 "Saludos / Best regards.";

    var cuerpoHtml =
      lineaBilingue("Buen dia", "Good day") + ",<br><br>" +
      "Favor de cotizar la siguiente pieza. Adjunto encontraras el dibujo tecnico y un formato de cotizacion; favor de llenarlo y regresarlo junto con tu respuesta (respondiendo a este mismo correo).<br>" +
      ingles("Please quote the following part. Attached you will find the technical drawing and a quotation form; please fill it out and return it along with your response (by replying to this same email).") + "<br><br>" +
      lineaBilingue("Numero de Parte", "Part Number") + ": " + codigo + "<br>" +
      lineaBilingue("Revision", "Revision") + ": " + revision + "<br>" +
      lineaBilingue("Descripcion", "Description") + ": " + descripcion + "<br>" +
      lineaMuestrasHtml +
      lineaEAUHtml + "<br>" +
      lineaBilingue("Saludos", "Best regards") + ".";

    var opciones = { attachments: [excelBlob], htmlBody: cuerpoHtml };
    if (archivoDibujo) {
      opciones.attachments.push(archivoDibujo.getBlob());
    }

    GmailApp.sendEmail(prov.correo, asunto, cuerpo, opciones);
    enviadosA.push(prov.correo);

    hojaLog.appendRow([
      rfqId, new Date(), codigo, revision, descripcion,
      prov.proveedor || "", prov.contacto || "", prov.correo, muestras || "", "Enviado",
      "", "", "", "", "", "", "", "", eau || "", enviadoPor
    ]);
  }

  // Marca la tarea como "Enviado" en la hoja principal, para no perder el
  // tracking de cuales tareas ya se mandaron a cotizar.
  sheet.getRange(filaIndex, 15).setValue("Enviado"); // columna O: Estado

  return {status: 'ok', encontrado: !!archivoDibujo, nombreBuscado: nombreArchivo, enviadoA: enviadosA};
}

function obtenerCarpetaDibujos_() {
  var valor = String(CARPETA_DIBUJOS || "").trim();
  var pareceId = /^[-\w]{25,}$/.test(valor);

  if (pareceId) {
    return DriveApp.getFolderById(valor);
  }

  var carpetas = DriveApp.getFoldersByName(valor);
  if (carpetas.hasNext()) {
    return carpetas.next();
  }

  throw new Error("No se encontro ninguna carpeta llamada '" + valor + "' en el Drive de esta cuenta.");
}

function buscarArchivoPorCodigo_(folderRaiz, codigoNorm, revisionNorm) {
  var candidatoGenerico = null;
  var revisionGenerico = null;
  var prefijoGenerico = codigoNorm + "_";
  var prefijoExacto = revisionNorm ? (codigoNorm + "_" + revisionNorm) : null;

  var pila = [folderRaiz];
  while (pila.length > 0) {
    var folderActual = pila.pop();

    var archivos = folderActual.getFiles();
    while (archivos.hasNext()) {
      var archivo = archivos.next();
      var nombre = archivo.getName().trim().toUpperCase();

      if (prefijoExacto && nombre.indexOf(prefijoExacto) === 0) {
        return { archivo: archivo, exacto: true, revisionEncontrada: revisionNorm };
      }
      if (!candidatoGenerico && nombre.indexOf(prefijoGenerico) === 0) {
        candidatoGenerico = archivo;
        var match = nombre.substring(prefijoGenerico.length).match(/^[A-Z0-9]+/);
        revisionGenerico = match ? match[0] : null;
      }
    }

    var subcarpetas = folderActual.getFolders();
    while (subcarpetas.hasNext()) {
      pila.push(subcarpetas.next());
    }
  }

  if (candidatoGenerico) return { archivo: candidatoGenerico, exacto: false, revisionEncontrada: revisionGenerico };
  return null;
}

function obtenerCarpetaPorFecha_(folderRaiz, fecha) {
  if (!fecha) return null;
  var d = new Date(fecha);
  if (isNaN(d.getTime())) return null;

  var anio = d.getFullYear();
  var mes = ("0" + (d.getMonth() + 1)).slice(-2);
  var dia = ("0" + d.getDate()).slice(-2);
  var nombreCarpeta = "dibujos_" + anio + "_" + mes + "_" + dia;

  var subcarpetas = folderRaiz.getFoldersByName(nombreCarpeta);
  if (subcarpetas.hasNext()) return subcarpetas.next();

  Logger.log("obtenerCarpetaPorFecha_: no existe la carpeta '" + nombreCarpeta + "'.");
  return null;
}

// NOTA: ya no se usa desde el boton del "ojo" en el Dashboard (ahora ese
// boton abre directo el link guardado en la columna Q de la hoja principal).
// Se deja disponible por si se necesita la busqueda por Drive en otro lugar.
function obtenerLinkDibujo(codigo, revision, fecha) {
  var folderRaiz = obtenerCarpetaDibujos_();
  var codigoNorm = String(codigo || "").trim().toUpperCase();
  var revisionNorm = String(revision || "").trim().toUpperCase();

  var encontrado = null;

  var carpetaFecha = obtenerCarpetaPorFecha_(folderRaiz, fecha);
  if (carpetaFecha) {
    encontrado = buscarArchivoPorCodigo_(carpetaFecha, codigoNorm, revisionNorm);
  }

  // Respaldo: si no esta en la carpeta de esa fecha (o no se pudo determinar
  // la carpeta), se hace la busqueda recursiva completa como antes.
  if (!encontrado) {
    Logger.log("obtenerLinkDibujo: no se encontro en la carpeta de fecha, se hace busqueda completa como respaldo.");
    encontrado = buscarArchivoPorCodigo_(folderRaiz, codigoNorm, revisionNorm);
  }

  if (!encontrado) {
    Logger.log("obtenerLinkDibujo: no se encontro archivo para codigo " + codigoNorm);
    return null;
  }

  return {
    url: "https://drive.google.com/file/d/" + encontrado.archivo.getId() + "/view",
    coincideRevision: encontrado.exacto,
    revisionEncontrada: encontrado.revisionEncontrada
  };
}

function obtenerMotivoDesdeDibujo(codigo, revision) {
  var folder = obtenerCarpetaDibujos_();
  var codigoNorm = String(codigo || "").trim().toUpperCase();
  var revisionNorm = String(revision || "").trim().toUpperCase();

  var encontrado = buscarArchivoPorCodigo_(folder, codigoNorm, revisionNorm);
  if (!encontrado || !encontrado.exacto) {
    Logger.log("obtenerMotivoDesdeDibujo: no hay dibujo exacto para " + codigoNorm + " rev " + revisionNorm);
    return null;
  }
  var archivoDibujo = encontrado.archivo;

  var docTemporalId = null;
  try {
    var recurso = {
      title: "OCR_TEMP_" + codigoNorm,
      mimeType: MimeType.GOOGLE_DOCS
    };
    var archivoConvertido = Drive.Files.copy(recurso, archivoDibujo.getId(), { ocr: true, ocrLanguage: "es" });
    docTemporalId = archivoConvertido.id;

    var doc = DocumentApp.openById(docTemporalId);
    var cuerpo = doc.getBody();
    var textoCompleto = cuerpo.getText();
    Logger.log("Texto OCR completo: " + textoCompleto);

    var motivoEncontrado = null;
    var letra = revisionNorm.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    var patron = new RegExp("(?:^|\\s)" + letra + "\\.\\s*([^.]+?)(?=\\s+[A-Z]\\.\\s|\\s+\\d{3,}[A-Z]{2}\\d|$)", "i");
    var match = textoCompleto.match(patron);
    if (match) {
      motivoEncontrado = match[1].trim();
    }

    Logger.log("obtenerMotivoDesdeDibujo: motivo encontrado por texto plano = " + motivoEncontrado);

    return motivoEncontrado;
  } catch (err) {
    Logger.log("obtenerMotivoDesdeDibujo error: " + err);
    return null;
  } finally {
    if (docTemporalId) DriveApp.getFileById(docTemporalId).setTrashed(true);
  }
}

function leerDatosDeExcelAdjunto(adjunto) {
  var archivoTemp = Drive.Files.create(
    { name: "temp_rfq_respuesta_" + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS },
    adjunto
  );
  var ssTemp = SpreadsheetApp.openById(archivoTemp.id);
  var hoja = ssTemp.getSheets()[0];
  var todasLasFilas = hoja.getDataRange().getValues();

  function buscarValorPorEtiqueta(textoEtiqueta) {
    for (var i = 0; i < todasLasFilas.length; i++) {
      var celdaA = String(todasLasFilas[i][0] || "");
      if (celdaA.indexOf(textoEtiqueta) !== -1) {
        return todasLasFilas[i][2];
      }
    }
    return "";
  }

  var datos = {
    precio: buscarValorPorEtiqueta("Precio Unitario"),
    moneda: buscarValorPorEtiqueta("Moneda"),
    moq: buscarValorPorEtiqueta("Cantidad Minima"),
    entrega: buscarValorPorEtiqueta("Tiempo de Entrega"),
    comentarios: buscarValorPorEtiqueta("Comentarios")
  };

  DriveApp.getFileById(archivoTemp.id).setTrashed(true);
  return datos;
}

function extraerCorreoDeRemitente(remitente) {
  var match = String(remitente || "").match(/<([^>]+)>/);
  return (match ? match[1] : (remitente || "")).trim().toLowerCase();
}

function revisarRespuestasRFQ() {
  var erroresDetectados = [];
  try {
    revisarRespuestasRFQInterno(erroresDetectados);
  } catch (errGeneral) {
    Logger.log("Error critico en revisarRespuestasRFQ: " + errGeneral.message);
    erroresDetectados.push("Error general deteniendo la revision completa: " + errGeneral.message);
  }
  if (erroresDetectados.length > 0) {
    notificarError("revisarRespuestasRFQ", erroresDetectados.join("\n\n---\n\n"));
  }
}

function revisarRespuestasRFQInterno(erroresDetectados) {
  var hojaLog = obtenerHojaLogRFQ();
  var datosLog = hojaLog.getDataRange().getValues();

  var quienEjecuta = (Session.getEffectiveUser().getEmail() || "").toLowerCase();
  Logger.log("revisarRespuestasRFQ corriendo como: " + (quienEjecuta || "(no se pudo determinar)"));
  var yoSoyElResponsableDeHuerfanos = quienEjecuta === CORREO_RESPONSABLE_NA.toLowerCase();

  var pendientesPorId = {};
  for (var i = 1; i < datosLog.length; i++) {
    if (datosLog[i][9] !== "Enviado") continue;

    var enviadoPor = (datosLog[i][19] || "").toLowerCase();
    var esMio = enviadoPor && enviadoPor === quienEjecuta;
    var esHuerfano = !enviadoPor;
    if (!esMio && !(esHuerfano && yoSoyElResponsableDeHuerfanos)) continue;

    pendientesPorId[datosLog[i][0]] = {
      fila: i + 1,
      correo: datosLog[i][7]
    };
  }

  var idsPendientes = Object.keys(pendientesPorId);
  Logger.log("RFQ IDs pendientes de respuesta: " + JSON.stringify(idsPendientes));
  if (idsPendientes.length === 0) {
    Logger.log("No hay ningun RFQ con Estado = 'Enviado'. Nada que revisar.");
    return;
  }

  var NOMBRE_ETIQUETA = "RFQ-Procesado";
  var etiqueta = GmailApp.getUserLabelByName(NOMBRE_ETIQUETA);
  if (!etiqueta) {
    etiqueta = GmailApp.createLabel(NOMBRE_ETIQUETA);
  }

  var hilos = GmailApp.search('has:attachment newer_than:30d -label:"' + NOMBRE_ETIQUETA + '"');
  Logger.log("Hilos encontrados con adjunto (sin la etiqueta) en los ultimos 30 dias: " + hilos.length);

  hilos.forEach(function(hilo) {
    var yaEtiquetadoEsteHilo = false;
    var mensajes = hilo.getMessages();
    mensajes.forEach(function(msg) {
      var asunto = msg.getSubject();
      Logger.log("Revisando correo con asunto: " + asunto);

      var match = asunto.match(/\[([A-Za-z0-9\-]+)\]/);
      if (!match) { Logger.log("  -> No se encontro un ID entre corchetes [ ] en el asunto. Se ignora."); return; }
      var rfqId = match[1];
      Logger.log("  -> ID encontrado: " + rfqId);

      var infoPendiente = pendientesPorId[rfqId];
      if (!infoPendiente) { Logger.log("  -> Ese ID no esta en la lista de pendientes. Se ignora."); return; }

      var remitente = extraerCorreoDeRemitente(msg.getFrom());
      var correoEsperado = String(infoPendiente.correo || "").trim().toLowerCase();
      if (remitente !== correoEsperado) {
        Logger.log("  -> El remitente (" + remitente + ") no es el proveedor esperado (" + correoEsperado + "). Se ignora.");
        return;
      }

      var adjuntos = msg.getAttachments();
      Logger.log("  -> Adjuntos en este correo: " + adjuntos.map(function(a){ return a.getName(); }).join(", "));

      var adjuntoExcel = null;
      for (var a = 0; a < adjuntos.length; a++) {
        if (/\.xlsx$/i.test(adjuntos[a].getName())) { adjuntoExcel = adjuntos[a]; break; }
      }
      var filaLog = infoPendiente.fila;

      // El proveedor respondio (coincide ID + remitente) pero no mando
      // ningun .xlsx: se marca como incompleto en vez de dejarlo revisando
      // el mismo correo para siempre.
      if (!adjuntoExcel) {
        Logger.log("  -> Ninguno de los adjuntos termina en .xlsx. Se marca como 'Respondido incompleto'.");
        hojaLog.getRange(filaLog, 10).setValue("Respondido incompleto");
        hojaLog.getRange(filaLog, 11).setValue(new Date());
        yaEtiquetadoEsteHilo = true;
        return;
      }

      try {
        var datos = leerDatosDeExcelAdjunto(adjuntoExcel);
        Logger.log("  -> Datos leidos del Excel: " + JSON.stringify(datos));
        // Si el excel llego pero sin precio, tambien es una respuesta
        // incompleta (falta la informacion que se necesita).
        var estadoFinal = datos.precio ? "Respondido" : "Respondido incompleto";
        hojaLog.getRange(filaLog, 10).setValue(estadoFinal);
        hojaLog.getRange(filaLog, 11).setValue(new Date());
        hojaLog.getRange(filaLog, 12).setValue(datos.precio);
        hojaLog.getRange(filaLog, 13).setValue(datos.moneda);
        hojaLog.getRange(filaLog, 14).setValue(datos.moq);
        hojaLog.getRange(filaLog, 15).setValue(datos.entrega);
        hojaLog.getRange(filaLog, 16).setValue(datos.comentarios);
        yaEtiquetadoEsteHilo = true;
        Logger.log("  -> Fila " + filaLog + " del Log RFQ actualizada con exito (" + estadoFinal + ").");
      } catch (err) {
        Logger.log("  -> ERROR al leer o guardar el adjunto: " + err.message);
        erroresDetectados.push("RFQ " + rfqId + " (correo " + correoEsperado + "): " + err.message);
      }
    });
    if (yaEtiquetadoEsteHilo) { hilo.addLabel(etiqueta); }
  });
}

// Corrige filas viejas que quedaron en Estado = "Respondido" sin precio
// (de antes de que existiera el estado "Respondido incompleto"). Las pasa
// a "Respondido incompleto" en vez de revertirlas a "Enviado", porque el
// proveedor ya contesto (aunque incompleto) y el correo ya quedo
// etiquetado como procesado en Gmail -- revertir a "Enviado" solo
// causaria que se le sigan mandando recordatorios de mas.
function revisarRespondidosSospechosos(soloSimular) {
  if (soloSimular === undefined) soloSimular = true;
  var hojaLog = obtenerHojaLogRFQ();
  var datos = hojaLog.getDataRange().getValues();
  var sospechosos = [];

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var estado = fila[9];
    if (estado !== "Respondido") continue;

    var precio = fila[11];
    if (precio) continue; // ya tiene precio, no es sospechoso

    sospechosos.push({ fila: i + 1, rfqId: fila[0], codigo: fila[2], revision: fila[3], correo: fila[7] });
  }

  Logger.log("Encontrados " + sospechosos.length + " RFQ's marcados 'Respondido' sin precio:");
  sospechosos.forEach(function(s) {
    Logger.log("  Fila " + s.fila + " | " + s.rfqId + " | " + s.codigo + " Rev." + s.revision + " | " + s.correo);
  });

  if (!soloSimular) {
    sospechosos.forEach(function(s) {
      hojaLog.getRange(s.fila, 10).setValue("Respondido incompleto");
    });
    Logger.log(sospechosos.length + " filas corregidas a 'Respondido incompleto'.");
  } else {
    Logger.log("Modo simulacion: no se modifico nada.");
  }

  return sospechosos;
}

function diagnosticoRecordatorios() {
  var hojaLog = obtenerHojaLogRFQ();
  var datos = hojaLog.getDataRange().getValues();
  var ahora = new Date();
  var CUATRO_DIAS_MS = 4 * 24 * 60 * 60 * 1000;
  var quienEjecuta = (Session.getEffectiveUser().getEmail() || "").toLowerCase();
  Logger.log("Diagnostico corriendo como: " + (quienEjecuta || "(no se pudo determinar)"));

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var rfqId = fila[0];
    if (!rfqId) continue;

    var estadoCrudo = fila[9];
    var fechaEnvio = fila[1];
    var ultimoRecordatorio = fila[16];
    var enviadoPor = fila[19] || "(vacio - RFQ viejo, sin dueño registrado)";

    var esEnviadoExacto = estadoCrudo === "Enviado";
    var fechaBase = (ultimoRecordatorio instanceof Date) ? ultimoRecordatorio : fechaEnvio;
    var fechaBaseEsDate = fechaBase instanceof Date;

    var linea = "Fila " + (i + 1) + " | RFQ " + rfqId + " | Enviado por: " + enviadoPor +
                " | Estado crudo: " + JSON.stringify(estadoCrudo) + " (coincide exacto con 'Enviado': " + esEnviadoExacto + ")";

    if (!esEnviadoExacto) { Logger.log(linea + " => SE SALTA."); continue; }
    if (!fechaBaseEsDate) { Logger.log(linea + " => SE SALTA: sin fecha base valida."); continue; }

    var transcurridoMs = ahora.getTime() - fechaBase.getTime();
    var dias = (transcurridoMs / (24 * 60 * 60 * 1000)).toFixed(2);
    var mandaria = transcurridoMs >= CUATRO_DIAS_MS;
    Logger.log(linea + " | Dias transcurridos: " + dias + " => " + (mandaria ? "SI MANDARIA." : "NO todavia."));
  }
  Logger.log("--- Fin del diagnostico ---");
}

function diagnosticoRespondidosSospechosos() { revisarRespondidosSospechosos(true); }
function corregirRespondidosSospechosos() { revisarRespondidosSospechosos(false); }

function enviarRecordatoriosRFQ() {
  var erroresDetectados = [];
  try {
    enviarRecordatoriosRFQInterno(erroresDetectados);
  } catch (errGeneral) {
    Logger.log("Error critico en enviarRecordatoriosRFQ: " + errGeneral.message);
    erroresDetectados.push("Error general deteniendo el envio de recordatorios: " + errGeneral.message);
  }
  if (erroresDetectados.length > 0) {
    notificarError("enviarRecordatoriosRFQ", erroresDetectados.join("\n\n---\n\n"));
  }
}

function enviarRecordatoriosRFQInterno(erroresDetectados) {
  var hojaLog = obtenerHojaLogRFQ();
  var datos = hojaLog.getDataRange().getValues();
  var ahora = new Date();
  var CUATRO_DIAS_MS = 4 * 24 * 60 * 60 * 1000;
  var enviados = 0;

  var quienEjecuta = (Session.getEffectiveUser().getEmail() || "").toLowerCase();
  Logger.log("enviarRecordatoriosRFQ corriendo como: " + (quienEjecuta || "(no se pudo determinar)"));

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var estado = fila[9];
    if (estado !== "Enviado") continue;

    var enviadoPor = (fila[19] || "").toLowerCase();
    var esMio = enviadoPor && enviadoPor === quienEjecuta;
    var esHuerfano = !enviadoPor;
    var yoSoyElResponsableDeHuerfanos = quienEjecuta === CORREO_RESPONSABLE_NA.toLowerCase();
    if (!esMio && !(esHuerfano && yoSoyElResponsableDeHuerfanos)) continue;

    var fechaEnvio = fila[1];
    var ultimoRecordatorio = fila[16];
    var fechaBase = (ultimoRecordatorio instanceof Date) ? ultimoRecordatorio : fechaEnvio;
    if (!(fechaBase instanceof Date)) continue;

    var transcurrido = ahora.getTime() - fechaBase.getTime();
    if (transcurrido < CUATRO_DIAS_MS) continue;

    var rfqId = fila[0];
    var codigo = fila[2];
    var revision = fila[3];
    var descripcion = fila[4];
    var correo = fila[7];
    var filaReal = i + 1;

    var asunto = "Recordatorio RFQ - " + codigo + " Rev. " + revision + " [" + rfqId + "]";

    var cuerpo = "Buen dia / Good day,\n\n" +
                 "Este es un recordatorio de la solicitud de cotizacion enviada anteriormente; aun no hemos recibido tu respuesta.\n" +
                 "This is a reminder of the previously sent quotation request; we have not yet received your response.\n\n" +
                 "Numero de Parte / Part Number: " + codigo + "\n" +
                 "Revision / Revision: " + revision + "\n" +
                 "Descripcion / Description: " + descripcion + "\n\n" +
                 "Agradecemos tu respuesta a la brevedad posible. / We would appreciate your prompt response.\n\n" +
                 "Saludos / Best regards.";

    var cuerpoHtml =
      lineaBilingue("Buen dia", "Good day") + ",<br><br>" +
      "Este es un recordatorio de la solicitud de cotizacion enviada anteriormente; aun no hemos recibido tu respuesta.<br>" +
      ingles("This is a reminder of the previously sent quotation request; we have not yet received your response.") + "<br><br>" +
      lineaBilingue("Numero de Parte", "Part Number") + ": " + codigo + "<br>" +
      lineaBilingue("Revision", "Revision") + ": " + revision + "<br>" +
      lineaBilingue("Descripcion", "Description") + ": " + descripcion + "<br><br>" +
      "Agradecemos tu respuesta a la brevedad posible. / " + ingles("We would appreciate your prompt response.") + "<br><br>" +
      lineaBilingue("Saludos", "Best regards") + ".";

    try {
      GmailApp.sendEmail(correo, asunto, cuerpo, { htmlBody: cuerpoHtml });
      hojaLog.getRange(filaReal, 17).setValue(ahora);
      var conteoPrevio = fila[17];
      var nuevoConteo = (typeof conteoPrevio === "number" ? conteoPrevio : 0) + 1;
      hojaLog.getRange(filaReal, 18).setValue(nuevoConteo);
      enviados++;
      Logger.log("Recordatorio enviado para " + rfqId + " a " + correo);
    } catch (err) {
      Logger.log("Error al mandar recordatorio para " + rfqId + ": " + err.message);
      erroresDetectados.push("Recordatorio para " + rfqId + " (" + correo + "): " + err.message);
    }
  }
  Logger.log("Total de recordatorios enviados: " + enviados);
}

function instalarRecordatorios() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "enviarRecordatoriosRFQ") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("enviarRecordatoriosRFQ").timeBased().everyDays(1).atHour(9).create();
  return {status: 'ok'};
}

function instalarRevisionAutomatica() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "revisarRespuestasRFQ") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("revisarRespuestasRFQ").timeBased().everyMinutes(30).create();
  return {status: 'ok'};
}

function tieneAutomatizacionesInstaladas() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var tieneRecordatorios = triggers.some(function(t) { return t.getHandlerFunction() === "enviarRecordatoriosRFQ"; });
    var tieneRevision = triggers.some(function(t) { return t.getHandlerFunction() === "revisarRespuestasRFQ"; });
    return { recordatorios: tieneRecordatorios, revision: tieneRevision, completo: tieneRecordatorios && tieneRevision };
  } catch (e) {
    return { recordatorios: false, revision: false, completo: false };
  }
}

function activarAutomatizacionesPropias() {
  instalarRecordatorios();
  instalarRevisionAutomatica();
  return tieneAutomatizacionesInstaladas();
}

// Pega aqui la URL /exec de tu SEGUNDA implementacion (la que le compartes a
// los jefes de departamento e ingenieros).
var URL_VISTA_JEFES = "https://script.google.com/a/macros/criotec.com.mx/s/AKfycbwo-4Fzn1PZY1sCRRi-7aqlKuI1VEeSMQszAciiX0g8IjaqoqRNVHgEw6gN2nhOCx10Ww/exec";

function registrarNP(fecha, nombreArchivo, iniciador) {
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];

  var colD = sheet.getRange("D1:D").getValues();
  var ultimaFila = 0;
  for (var i = 0; i < colD.length; i++) {
    if (colD[i][0] !== "") ultimaFila = i + 1;
  }
  var filaNueva = ultimaFila + 1;

  sheet.getRange(filaNueva, 3).setValue(fecha);
  sheet.getRange(filaNueva, 4).setValue(nombreArchivo);
  sheet.getRange(filaNueva, 9).setValue(iniciador);

  SpreadsheetApp.flush();

  try {
    cancelarRevisionesAnteriores(sheet, filaNueva);
  } catch (err) {
    Logger.log("Error al intentar cancelar revisiones anteriores: " + err.message);
    notificarError("registrarNP - cancelarRevisionesAnteriores", err.message + "\n\n" + (err.stack || ""));
  }

  try {
    asignarSourcingPorPrefijo(sheet, filaNueva);
  } catch (err) {
    Logger.log("Error al asignar sourcing por prefijo: " + err.message);
    notificarError("registrarNP - asignarSourcingPorPrefijo", err.message + "\n\n" + (err.stack || ""));
  }

  try {
    notificarAsignacionNueva(sheet, filaNueva);
  } catch (err) {
    Logger.log("Error al notificar asignacion nueva: " + err.message);
    notificarError("registrarNP - notificarAsignacionNueva", err.message + "\n\n" + (err.stack || ""));
  }

  return { status: 'ok', fila: filaNueva };
}

function manejarRegistrarNPporGet(e) {
  try {
    var fecha = e.parameter.fecha || "";
    var nombreArchivo = e.parameter.nombreArchivo || "";
    var iniciador = e.parameter.iniciador || "";

    var resultado = registrarNP(fecha, nombreArchivo, iniciador);

    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (errGeneral) {
    Logger.log("Error critico en manejarRegistrarNPporGet: " + errGeneral.message);
    notificarError(
      "registrarNP (via GET)",
      errGeneral.message + "\n\n" + (errGeneral.stack || "") +
      "\n\nParametros recibidos: " + JSON.stringify(e && e.parameter ? e.parameter : {})
    );
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: errGeneral.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function extraerIdDeploymentDeUrl(url) {
  if (!url) return "";
  var match = String(url).match(/\/s\/([^\/]+)\//);
  return match ? match[1] : "";
}

function manejarNotificarTareaNuevaPorGet(e) {
  try {
    var jobName = e.parameter.jobName || "(sin nombre)";
    var iniciador = e.parameter.iniciador || "";
    var asunto = "Nueva tarea en tu bandeja de Teamcenter";
    var cuerpo = "Te llego una tarea nueva en Teamcenter:\n\n" +
                 "Tarea: " + jobName + "\n" +
                 (iniciador ? ("Solicitada por: " + iniciador + "\n") : "") +
                 "Fecha/hora detectada: " + new Date().toLocaleString("es-MX");
    GmailApp.sendEmail(CORREO_ALERTAS, asunto, cuerpo);
    return ContentService.createTextOutput(JSON.stringify({status: 'ok'})).setMimeType(ContentService.MimeType.JSON);
  } catch (errGeneral) {
    Logger.log("Error en manejarNotificarTareaNuevaPorGet: " + errGeneral.message);
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: errGeneral.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.accion === "registrarNP") {
    return manejarRegistrarNPporGet(e);
  }
  if (e && e.parameter && e.parameter.accion === "notificarTareaNueva") {
    return manejarNotificarTareaNuevaPorGet(e);
  }

  var correoActual = Session.getActiveUser().getEmail() || "";
  var idDeploymentActual = extraerIdDeploymentDeUrl(ScriptApp.getService().getUrl());
  var idDeploymentJefes = extraerIdDeploymentDeUrl(URL_VISTA_JEFES);

  if (idDeploymentActual === idDeploymentJefes) {
    return renderVistaSolicitante(correoActual);
  }

  return renderDashboardSourcing();
}

function renderDashboardSourcing() {
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];

  var colD = sheet.getRange("D2:D" + sheet.getMaxRows()).getValues();
  var ultimaFila = 1;
  for (var i = 0; i < colD.length; i++) {
    if (colD[i][0] !== "") ultimaFila = i + 2;
  }

  // 17 columnas para incluir la columna Q (link del dibujo).
  var data = sheet.getRange(1, 1, ultimaFila, 17).getValues();
  var filas = [];
  for (var i = 1; i < data.length; i++) {
    var fila = data[i];
    var codigoCompleto = fila[3];
    if (!codigoCompleto) continue;
    var fecha = fila[2];
    filas.push({
      fila: i + 1,
      fecha: (fecha instanceof Date) ? fecha.toISOString() : fecha,
      codigo: fila[4], revision: fila[5], tipo: fila[6], descripcion: fila[7],
      ingeniero: fila[8], area: fila[9], commodity: fila[10], comprador: fila[11],
      sourcing: fila[12], muestras: fila[13], estado: fila[14],
      motivo: fila[15] || "",
      enlaceDibujo: fila[16] || ""
    });
  }

  var mapaCotizaciones = obtenerCotizacionesPorPieza();
  filas.forEach(function(f) {
    var key = String(f.codigo).trim().toUpperCase() + "|" + String(f.revision).trim().toUpperCase();
    f.cotizaciones = mapaCotizaciones[key] || [];
  });

  var template = HtmlService.createTemplateFromFile('Dashboard');
  template.datosJson = JSON.stringify(filas);

  var hojaProveedores = SpreadsheetApp.openByUrl(SHEET_URL).getSheetByName("Proveedores");
  var proveedores = [];
  if (hojaProveedores) {
    var datosProveedores = hojaProveedores.getDataRange().getValues();
    for (var p = 1; p < datosProveedores.length; p++) {
      var filaProv = datosProveedores[p];
      if (!filaProv[2]) continue; // Name vacio = fila vacia
      proveedores.push({
        nomenclatura: filaProv[0],
        supplierId: filaProv[1],
        proveedor: filaProv[2],
        commodity: filaProv[3],
        responsable: filaProv[4],
        compras: filaProv[5],
        contacto: filaProv[6],
        correo: filaProv[7]
      });
    }
  }
  template.proveedoresJson = JSON.stringify(proveedores);
  template.logoUrl = getLogoDataUri_();

  var correoActual = Session.getActiveUser().getEmail() || "";
  var nombreActual = MAPA_SOURCING[correoActual.toLowerCase()] || "";
  var esJefe = correoActual.toLowerCase() === CORREO_JEFE.toLowerCase();
  template.correoActualJson = JSON.stringify(correoActual);
  template.nombreActualJson = JSON.stringify(nombreActual);
  template.esJefeJson = JSON.stringify(esJefe);
  template.correoResponsableNAJson = JSON.stringify(CORREO_RESPONSABLE_NA);

  var estadoAutomatizaciones = { recordatorios: false, revision: false, completo: false };
  try { estadoAutomatizaciones = tieneAutomatizacionesInstaladas(); } catch (e) {}
  template.estadoAutomatizacionesJson = JSON.stringify(estadoAutomatizaciones);

  template.ticketsJson = JSON.stringify(obtenerTickets());
  template.adaptacionesJson = JSON.stringify(obtenerAdaptaciones());
  template.proveedoresCalcasJson = JSON.stringify(obtenerProveedoresCalcas());

  return template.evaluate()
    .setTitle('Seguimiento de tareas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderVistaSolicitante(correoActual) {
  var correoLower = correoActual.toLowerCase();
  var mapaJefes = obtenerMapaJefesDepartamento();
  var mapaIngenieria = obtenerMapaIngenieria();
  var esJefeDepartamento = mapaJefes.hasOwnProperty(correoLower);
  var esIngeniero = mapaIngenieria.hasOwnProperty(correoLower);

  var template = HtmlService.createTemplateFromFile('Solicitudes');
  template.correoActualJson = JSON.stringify(correoActual);
  template.logoUrl = getLogoDataUri_();

  template.esJefeDepartamentoJson = JSON.stringify(esJefeDepartamento);
  template.departamentoJson = JSON.stringify(esJefeDepartamento ? mapaJefes[correoLower] : "");
  template.ticketsJson = JSON.stringify(esJefeDepartamento ? obtenerTickets(correoActual) : []);
  template.sourcingDisponibleJson = JSON.stringify(obtenerListaSourcing());

  template.esIngenieroJson = JSON.stringify(esIngeniero);
  template.nombreIngenieroJson = JSON.stringify(esIngeniero ? mapaIngenieria[correoLower] : "");
  template.adaptacionesJson = JSON.stringify(esIngeniero ? obtenerAdaptaciones(correoActual) : []);

  return template.evaluate()
    .setTitle('Solicitudes a Sourcing')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Nombres asignables en el selector "Asignar a" de tickets: todo Sourcing +
// el jefe (Hugo), aunque el jefe no tenga tareas de Teamcenter directamente.
function obtenerListaSourcing() {
  var nombres = Object.keys(MAPA_SOURCING).map(function(correo) { return MAPA_SOURCING[correo]; });
  nombres.push(NOMBRE_JEFE);
  nombres = Array.from(new Set(nombres));
  nombres.sort(function(a, b) { return String(a).localeCompare(String(b), "es"); });
  return nombres;
}

function obtenerHojaJefesDepartamento() {
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var hoja = ss.getSheetByName("Jefes Departamento");
  if (!hoja) {
    hoja = ss.insertSheet("Jefes Departamento");
    hoja.appendRow(["Correo", "Departamento"]);
    hoja.getRange(1, 1, 1, 2).setFontWeight("bold");
  }
  return hoja;
}

function obtenerMapaJefesDepartamento() {
  var hoja = obtenerHojaJefesDepartamento();
  var datos = hoja.getDataRange().getValues();
  var mapa = {};
  for (var i = 1; i < datos.length; i++) {
    var correo = datos[i][0];
    var departamento = datos[i][1];
    if (!correo) continue;
    mapa[String(correo).trim().toLowerCase()] = departamento;
  }
  return mapa;
}

function obtenerHojaTickets() {
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var hoja = ss.getSheetByName("Tickets Sourcing");
  if (!hoja) {
    hoja = ss.insertSheet("Tickets Sourcing");
    hoja.appendRow([
      "ID", "Fecha Creacion", "Departamento", "Correo Solicitante",
      "Descripcion", "Fecha Requerida", "Prioridad", "Estado",
      "Comentarios Sourcing", "Sourcing Asignado", "Adjunto"
    ]);
    hoja.getRange(1, 1, 1, 11).setFontWeight("bold");
    return hoja;
  }
  if (hoja.getLastColumn() < 9) {
    hoja.getRange(1, 9).setValue("Comentarios Sourcing");
    hoja.getRange(1, 9).setFontWeight("bold");
  }
  if (hoja.getLastColumn() < 10) {
    hoja.getRange(1, 10).setValue("Sourcing Asignado");
    hoja.getRange(1, 10).setFontWeight("bold");
  }
  if (hoja.getLastColumn() < 11) {
    hoja.getRange(1, 11).setValue("Adjunto");
    hoja.getRange(1, 11).setFontWeight("bold");
  }
  return hoja;
}

function quitarAcentos(texto) {
  var codigoInicio = 0x0300;
  var codigoFin = 0x036f;
  var rango = "\\u" + codigoInicio.toString(16) + "-\\u" + codigoFin.toString(16);
  var regexDiacriticos = new RegExp("[" + rango + "]", "g");
  return String(texto || "").normalize("NFD").replace(regexDiacriticos, "");
}

function generarIdTicket(departamento) {
  var prefijo = quitarAcentos(departamento).toUpperCase().replace(/[^A-Z]/g, "").substring(0, 3);
  while (prefijo.length < 3) prefijo += "X";

  var hoja = obtenerHojaTickets();
  var ultimaFila = hoja.getLastRow();
  var maxNumero = 0;

  if (ultimaFila > 1) {
    var idsExistentes = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
    var patron = new RegExp("^" + prefijo + "-(\\d+)$");
    idsExistentes.forEach(function(fila) {
      var id = fila[0];
      if (!id) return;
      var match = String(id).match(patron);
      if (match) {
        var numero = parseInt(match[1], 10);
        if (numero > maxNumero) maxNumero = numero;
      }
    });
  }

  var siguiente = maxNumero + 1;
  var numeroTexto = String(siguiente);
  while (numeroTexto.length < 3) numeroTexto = "0" + numeroTexto;

  return prefijo + "-" + numeroTexto;
}

function obtenerTickets(filtroCorreo) {
  var hoja = obtenerHojaTickets();
  var datos = hoja.getDataRange().getValues();
  var tickets = [];
  var filtroLower = filtroCorreo ? String(filtroCorreo).trim().toLowerCase() : null;

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var id = fila[0];
    if (!id) continue;

    var correo = fila[3];
    if (filtroLower && String(correo).trim().toLowerCase() !== filtroLower) continue;

    var fechaCreacion = fila[1];
    var fechaRequerida = fila[5];

    tickets.push({
      fila: i + 1,
      id: id,
      fechaCreacion: (fechaCreacion instanceof Date) ? fechaCreacion.toISOString() : fechaCreacion,
      departamento: fila[2],
      correo: correo,
      descripcion: fila[4],
      fechaRequerida: (fechaRequerida instanceof Date) ? fechaRequerida.toISOString() : fechaRequerida,
      prioridad: fila[6],
      estado: fila[7],
      comentarios: fila[8],
      sourcingAsignado: fila[9],
      adjunto: fila[10]
    });
  }
  return tickets;
}

// Carpeta compartida de Drive donde se guardan los documentos que adjuntan
// los departamentos al crear un ticket.
function obtenerCarpetaAdjuntosTickets() {
  var nombreCarpeta = "Tickets Sourcing - Adjuntos";
  var carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(nombreCarpeta);
}

// archivoNombre/archivoMimeType/archivoBase64 son opcionales: si el
// solicitante no adjunta nada, se llaman con "" o se omiten.
function crearTicket(descripcion, fechaRequerida, prioridad, sourcingAsignado, archivoNombre, archivoMimeType, archivoBase64) {
  var correoActual = Session.getActiveUser().getEmail() || "";
  var correoLower = correoActual.toLowerCase();
  var mapaJefes = obtenerMapaJefesDepartamento();
  var departamento = mapaJefes[correoLower];

  if (!departamento) {
    throw new Error(
      "Tu correo (" + correoActual + ") no esta registrado como jefe de departamento. " +
      "Pide a Sourcing que te agreguen en la hoja 'Jefes Departamento'."
    );
  }
  if (!descripcion || !String(descripcion).trim()) {
    throw new Error("La descripcion no puede estar vacia.");
  }

  var urlAdjunto = "";
  if (archivoBase64 && archivoNombre) {
    try {
      var carpeta = obtenerCarpetaAdjuntosTickets();
      var blob = Utilities.newBlob(
        Utilities.base64Decode(archivoBase64),
        archivoMimeType || "application/octet-stream",
        archivoNombre
      );
      var archivo = carpeta.createFile(blob);
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urlAdjunto = archivo.getUrl();
    } catch (errAdjunto) {
      Logger.log("No se pudo guardar el adjunto del ticket: " + errAdjunto.message);
      notificarError("crearTicket - adjunto", errAdjunto.message + "\n\n" + (errAdjunto.stack || ""));
    }
  }

  var hoja = obtenerHojaTickets();
  var id = generarIdTicket(departamento);
  var fechaCreacion = new Date();
  var fechaRequeridaDate = fechaRequerida ? new Date(fechaRequerida) : "";
  var prioridadFinal = prioridad || "Media";
  var sourcingAsignadoFinal = (sourcingAsignado && String(sourcingAsignado).trim()) ? sourcingAsignado : "Sin preferencia";

  hoja.appendRow([
    id, fechaCreacion, departamento, correoActual, descripcion,
    fechaRequeridaDate, prioridadFinal, "Abierto", "", sourcingAsignadoFinal, urlAdjunto
  ]);

  return {
    fila: hoja.getLastRow(),
    id: id,
    fechaCreacion: fechaCreacion.toISOString(),
    departamento: departamento,
    correo: correoActual,
    descripcion: descripcion,
    fechaRequerida: fechaRequeridaDate ? fechaRequeridaDate.toISOString() : "",
    prioridad: prioridadFinal,
    estado: "Abierto",
    comentarios: "",
    sourcingAsignado: sourcingAsignadoFinal,
    adjunto: urlAdjunto
  };
}

function esUsuarioDeSourcing(correo) {
  var correoLower = String(correo || "").toLowerCase();
  return MAPA_SOURCING.hasOwnProperty(correoLower) || correoLower === CORREO_JEFE.toLowerCase();
}

function cambiarEstadoTicket(filaIndex, nuevoEstado) {
  if (!esUsuarioDeSourcing(Session.getActiveUser().getEmail())) {
    throw new Error("Solo el equipo de Sourcing puede cambiar el estado de un ticket.");
  }
  var hoja = obtenerHojaTickets();
  hoja.getRange(filaIndex, 8).setValue(nuevoEstado);
  return {status: 'ok'};
}

function enviarCorreoActualizacionTicket(correoDestino, asunto, cuerpo) {
  try {
    GmailApp.sendEmail(correoDestino, asunto, cuerpo);
  } catch (err) {
    Logger.log("No se pudo enviar el correo de actualizacion de ticket: " + err.message);
    notificarError("Tickets Sourcing - notificacion", err.message + "\n\nDestinatario: " + correoDestino + "\nAsunto: " + asunto);
  }
}

function marcarTrabajandoEnTicket(filaIndex) {
  if (!esUsuarioDeSourcing(Session.getActiveUser().getEmail())) {
    throw new Error("Solo el equipo de Sourcing puede hacer esto.");
  }

  var hoja = obtenerHojaTickets();
  var fila = hoja.getRange(filaIndex, 1, 1, 8).getValues()[0];
  var id = fila[0];
  var departamento = fila[2];
  var correoCreador = fila[3];
  var descripcion = fila[4];

  hoja.getRange(filaIndex, 8).setValue("En Proceso");

  var asunto = "Tu ticket a Sourcing (" + id + ") ya esta en proceso";
  var cuerpo = "Hola,\n\n" +
    "Te avisamos que el equipo de Sourcing ya empezo a trabajar en tu solicitud:\n\n" +
    "Ticket: " + id + "\n" +
    "Departamento: " + departamento + "\n" +
    "Descripcion: " + descripcion + "\n\n" +
    "Te seguiremos informando conforme haya avances.\n\n" +
    "Saludos,\nSourcing Criotec";

  enviarCorreoActualizacionTicket(correoCreador, asunto, cuerpo);

  return {status: 'ok', estado: 'En Proceso'};
}

function guardarComentarioTicket(filaIndex, comentario) {
  if (!esUsuarioDeSourcing(Session.getActiveUser().getEmail())) {
    throw new Error("Solo el equipo de Sourcing puede hacer esto.");
  }

  var hoja = obtenerHojaTickets();
  var fila = hoja.getRange(filaIndex, 1, 1, 8).getValues()[0];
  var id = fila[0];
  var departamento = fila[2];
  var correoCreador = fila[3];
  var descripcion = fila[4];

  hoja.getRange(filaIndex, 9).setValue(comentario);

  if (comentario && String(comentario).trim()) {
    var asunto = "Actualizacion de tu ticket a Sourcing (" + id + ")";
    var cuerpo = "Hola,\n\n" +
      "Sourcing agrego una actualizacion a tu solicitud:\n\n" +
      "Ticket: " + id + "\n" +
      "Departamento: " + departamento + "\n" +
      "Descripcion: " + descripcion + "\n\n" +
      "Comentario de Sourcing:\n" + comentario + "\n\n" +
      "Saludos,\nSourcing Criotec";

    enviarCorreoActualizacionTicket(correoCreador, asunto, cuerpo);
  }

  return {status: 'ok'};
}

// Cancela un ticket. Solo Sourcing puede hacerlo, y es obligatorio dar un
// motivo (queda guardado en "Comentarios Sourcing" y se le avisa por correo
// a quien lo creo).
function cancelarTicket(filaIndex, motivo) {
  if (!esUsuarioDeSourcing(Session.getActiveUser().getEmail())) {
    throw new Error("Solo el equipo de Sourcing puede cancelar un ticket.");
  }
  if (!motivo || !String(motivo).trim()) {
    throw new Error("Debes indicar el motivo de la cancelacion.");
  }

  var hoja = obtenerHojaTickets();
  var fila = hoja.getRange(filaIndex, 1, 1, 8).getValues()[0];
  var id = fila[0];
  var departamento = fila[2];
  var correoCreador = fila[3];
  var descripcion = fila[4];

  var comentarioFinal = "Cancelado: " + motivo;
  hoja.getRange(filaIndex, 8).setValue("Cancelado");
  hoja.getRange(filaIndex, 9).setValue(comentarioFinal);

  var asunto = "Tu ticket a Sourcing (" + id + ") fue cancelado";
  var cuerpo = "Hola,\n\n" +
    "Te informamos que tu solicitud a Sourcing fue cancelada:\n\n" +
    "Ticket: " + id + "\n" +
    "Departamento: " + departamento + "\n" +
    "Descripcion: " + descripcion + "\n\n" +
    "Motivo de la cancelacion:\n" + motivo + "\n\n" +
    "Si tienes dudas, contacta directamente a Sourcing.\n\n" +
    "Saludos,\nSourcing Criotec";

  enviarCorreoActualizacionTicket(correoCreador, asunto, cuerpo);

  return {status: 'ok', estado: 'Cancelado', comentarios: comentarioFinal};
}

// ============================================================================
// MODULO "ADAPTACIONES"
// ============================================================================
var COMMODITY_CALCAS = "Calcas";

function obtenerHojaIngenieria() {
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var hoja = ss.getSheetByName("Ingenieria");
  if (!hoja) {
    hoja = ss.insertSheet("Ingenieria");
    hoja.appendRow(["Correo", "Nombre"]);
    hoja.getRange(1, 1, 1, 2).setFontWeight("bold");
  }
  return hoja;
}

function obtenerMapaIngenieria() {
  var hoja = obtenerHojaIngenieria();
  var datos = hoja.getDataRange().getValues();
  var mapa = {};
  for (var i = 1; i < datos.length; i++) {
    var correo = datos[i][0];
    var nombre = datos[i][1];
    if (!correo) continue;
    mapa[String(correo).trim().toLowerCase()] = nombre;
  }
  return mapa;
}

function obtenerHojaAdaptaciones() {
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var hoja = ss.getSheetByName("Adaptaciones");
  if (!hoja) {
    hoja = ss.insertSheet("Adaptaciones");
    hoja.appendRow([
      "ID", "Fecha Creacion", "Correo Ingeniero", "Nombre Ingeniero",
      "Carpeta Drive", "Codigos", "Comentario", "Estado",
      "Proveedor Asignado", "Correo Proveedor", "Fecha Envio"
    ]);
    hoja.getRange(1, 1, 1, 11).setFontWeight("bold");
  }
  return hoja;
}

function generarIdAdaptacion() {
  var hoja = obtenerHojaAdaptaciones();
  var ultimaFila = hoja.getLastRow();
  var maxNumero = 0;

  if (ultimaFila > 1) {
    var idsExistentes = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues();
    var patron = /^ADP-(\d+)$/;
    idsExistentes.forEach(function(fila) {
      var id = fila[0];
      if (!id) return;
      var match = String(id).match(patron);
      if (match) {
        var numero = parseInt(match[1], 10);
        if (numero > maxNumero) maxNumero = numero;
      }
    });
  }

  var siguiente = maxNumero + 1;
  var numeroTexto = String(siguiente);
  while (numeroTexto.length < 3) numeroTexto = "0" + numeroTexto;

  return "ADP-" + numeroTexto;
}

function obtenerAdaptaciones(filtroCorreo) {
  var hoja = obtenerHojaAdaptaciones();
  var datos = hoja.getDataRange().getValues();
  var adaptaciones = [];
  var filtroLower = filtroCorreo ? String(filtroCorreo).trim().toLowerCase() : null;

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var id = fila[0];
    if (!id) continue;

    var correoIngeniero = fila[2];
    if (filtroLower && String(correoIngeniero).trim().toLowerCase() !== filtroLower) continue;

    var fechaCreacion = fila[1];
    var fechaEnvio = fila[10];

    adaptaciones.push({
      fila: i + 1,
      id: id,
      fechaCreacion: (fechaCreacion instanceof Date) ? fechaCreacion.toISOString() : fechaCreacion,
      correoIngeniero: correoIngeniero,
      nombreIngeniero: fila[3],
      carpetaDrive: fila[4],
      codigos: fila[5],
      comentario: fila[6],
      estado: fila[7],
      proveedorAsignado: fila[8],
      correoProveedor: fila[9],
      fechaEnvio: (fechaEnvio instanceof Date) ? fechaEnvio.toISOString() : fechaEnvio
    });
  }
  return adaptaciones;
}

function crearAdaptacion(carpetaDrive, codigos, comentario) {
  var correoActual = Session.getActiveUser().getEmail() || "";
  var correoLower = correoActual.toLowerCase();
  var mapaIngenieria = obtenerMapaIngenieria();
  var nombreIngeniero = mapaIngenieria[correoLower];

  if (!nombreIngeniero) {
    throw new Error(
      "Tu correo (" + correoActual + ") no esta registrado como ingeniero. " +
      "Pide a Sourcing que te agreguen en la hoja 'Ingenieria'."
    );
  }
  if (!carpetaDrive || !String(carpetaDrive).trim()) {
    throw new Error("Falta el link de la carpeta de Drive.");
  }
  if (!codigos || !String(codigos).trim()) {
    throw new Error("Falta indicar a que codigos corresponde.");
  }

  var hoja = obtenerHojaAdaptaciones();
  var id = generarIdAdaptacion();
  var fechaCreacion = new Date();

  hoja.appendRow([
    id, fechaCreacion, correoActual, nombreIngeniero,
    carpetaDrive, codigos, comentario || "", "Pendiente", "", "", ""
  ]);

  return {
    fila: hoja.getLastRow(),
    id: id,
    fechaCreacion: fechaCreacion.toISOString(),
    correoIngeniero: correoActual,
    nombreIngeniero: nombreIngeniero,
    carpetaDrive: carpetaDrive,
    codigos: codigos,
    comentario: comentario || "",
    estado: "Pendiente",
    proveedorAsignado: "",
    correoProveedor: "",
    fechaEnvio: ""
  };
}

function enviarAdaptacionAProveedor(filaIndex, proveedoresSeleccionados) {
  if (!esUsuarioDeSourcing(Session.getActiveUser().getEmail())) {
    throw new Error("Solo el equipo de Sourcing puede hacer esto.");
  }
  if (!proveedoresSeleccionados || proveedoresSeleccionados.length === 0) {
    throw new Error("Selecciona al menos un proveedor (o escribe un correo) antes de enviar.");
  }

  var hoja = obtenerHojaAdaptaciones();
  var fila = hoja.getRange(filaIndex, 1, 1, 7).getValues()[0];
  var id = fila[0];
  var nombreIngeniero = fila[3];
  var correoIngeniero = fila[2];
  var carpetaDrive = fila[4];
  var codigos = fila[5];
  var comentario = fila[6];

  var enviadosA = [];
  var nombresProveedores = [];

  proveedoresSeleccionados.forEach(function(p) {
    var asunto = "Adaptacion de calcamonia - " + id + " (Codigos: " + codigos + ")";
    var cuerpo = "Buen dia,\n\n" +
      "Se solicita adaptar la(s) imagen(es) de la siguiente carpeta de Drive para los codigos indicados:\n\n" +
      "Codigos: " + codigos + "\n" +
      "Carpeta de Drive: " + carpetaDrive + "\n" +
      (comentario ? ("Comentario: " + comentario + "\n") : "") +
      "\nSolicitado por: " + nombreIngeniero + " (Ingenieria Criotec)\n\n" +
      "Quedamos al pendiente de su confirmacion.\n\n" +
      "Saludos,\nSourcing Criotec";

    GmailApp.sendEmail(p.correo, asunto, cuerpo, { cc: correoIngeniero });
    enviadosA.push(p.correo);
    nombresProveedores.push(p.nombre || p.correo);
  });

  var fechaEnvio = new Date();
  hoja.getRange(filaIndex, 8).setValue("Enviada");
  hoja.getRange(filaIndex, 9).setValue(nombresProveedores.join(", "));
  hoja.getRange(filaIndex, 10).setValue(enviadosA.join(", "));
  hoja.getRange(filaIndex, 11).setValue(fechaEnvio);

  return {
    status: 'ok',
    estado: 'Enviada',
    proveedorAsignado: nombresProveedores.join(", "),
    correoProveedor: enviadosA.join(", "),
    fechaEnvio: fechaEnvio.toISOString(),
    enviadoA: enviadosA
  };
}

// ============================================================================
// HTML separado en 3 archivos: Dashboard.html (shell), Dashboard-Escritorio.html
// (vista web) y Dashboard-Movil.html (vista app).
// ============================================================================
function include(nombreArchivo, templatePadre) {
  var t = HtmlService.createTemplateFromFile(nombreArchivo);
  if (templatePadre) {
    for (var key in templatePadre) {
      t[key] = templatePadre[key];
    }
  }
  return t.evaluate().getContent();
}

// ============================================================================
// Pestaña "Proveedores": 8 columnas, en este orden:
// Nomenclatura | Supplier ID | Name | Commodity | Responsable | Compras | Contacto | Correo contacto
// ============================================================================

function cancelarTarea(filaIndex, motivo) {
  if (!motivo || !String(motivo).trim()) {
    throw new Error("Debes indicar el motivo de la cancelacion.");
  }

  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];
  var codigo = sheet.getRange(filaIndex, 5).getValue();   // col E
  var revision = sheet.getRange(filaIndex, 6).getValue(); // col F

  var estadoFinal = "Cancelada";
  sheet.getRange(filaIndex, 15).setValue(estadoFinal);       // col O: Estado
  sheet.getRange(filaIndex, 16).setValue(String(motivo));    // col P: Motivo

  try {
    if (codigo) cancelarRFQsDeLaPieza(codigo, revision);
  } catch (err) {
    Logger.log("cancelarTarea: no se pudieron cancelar los RFQ asociados: " + err.message);
    notificarError("cancelarTarea", err.message + "\n\n" + (err.stack || ""));
  }

  return { status: 'ok', estado: estadoFinal };
}

// Cancela una revision especifica que ya dejo de ser la vigente para su
// codigo (p. ej. llego la Rev.B pero la Rev.A seguia abierta sin RFQ con
// proveedor, y alguien la cancela a mano desde el detalle de revisiones).
// A diferencia de cancelarTarea, no pide motivo: la UI ya confirma la accion
// con un dialogo antes de llamar a esta funcion.
function cancelarRevisionPasada(filaIndex, codigo, revision) {
  var sheet = SpreadsheetApp.openByUrl(SHEET_URL).getSheets()[0];
  var estadoFinal = "Cancelada";
  sheet.getRange(filaIndex, 15).setValue(estadoFinal); // col O: Estado

  try {
    if (codigo) cancelarRFQsDeLaPieza(codigo, revision);
  } catch (err) {
    Logger.log("cancelarRevisionPasada: no se pudieron cancelar los RFQ asociados: " + err.message);
    notificarError("cancelarRevisionPasada", err.message + "\n\n" + (err.stack || ""));
  }

  return { status: 'ok', estado: estadoFinal };
}

var ENCABEZADOS_PROVEEDORES = ["Nomenclatura", "Supplier ID", "Name", "Commodity", "Responsable", "Compras", "Contacto", "Correo contacto"];

function agregarProveedor(nuevoProveedor) {
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var hoja = ss.getSheetByName("Proveedores");
  if (!hoja) {
    hoja = ss.insertSheet("Proveedores");
    hoja.appendRow(ENCABEZADOS_PROVEEDORES);
    hoja.getRange(1, 1, 1, ENCABEZADOS_PROVEEDORES.length).setFontWeight("bold");
    hoja.setFrozenRows(1);
  }

  var nomenclatura = String((nuevoProveedor && nuevoProveedor.nomenclatura) || "").trim();
  var supplierId = String((nuevoProveedor && nuevoProveedor.supplierId) || "").trim();
  var proveedor = String((nuevoProveedor && nuevoProveedor.proveedor) || "").trim();
  var commodity = String((nuevoProveedor && nuevoProveedor.commodity) || "").trim();
  var responsable = String((nuevoProveedor && nuevoProveedor.responsable) || "").trim();
  var compras = String((nuevoProveedor && nuevoProveedor.compras) || "").trim();
  var contacto = String((nuevoProveedor && nuevoProveedor.contacto) || "").trim();
  var correo = String((nuevoProveedor && nuevoProveedor.correo) || "").trim();

  if (!proveedor) throw new Error("Falta el nombre (Name) del proveedor.");
  if (!commodity) throw new Error("Falta el commodity del proveedor.");
  if (!correo) throw new Error("Falta el correo de contacto del proveedor.");

  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    var correoFila = String(datos[i][7] || "").trim().toLowerCase();
    var commodityFila = String(datos[i][3] || "").trim().toLowerCase();
    if (correoFila === correo.toLowerCase() && commodityFila === commodity.toLowerCase()) {
      throw new Error("Ese proveedor ya esta en la lista para este commodity.");
    }
  }

  hoja.appendRow([nomenclatura, supplierId, proveedor, commodity, responsable, compras, contacto, correo]);
  return { status: 'ok' };
}

function obtenerProveedoresCalcas() {
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var hoja = ss.getSheetByName("Proveedores");
  var proveedores = [];
  if (!hoja) return proveedores;

  var datos = hoja.getDataRange().getValues();
  var commodityBuscado = COMMODITY_CALCAS.trim().toLowerCase();

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    if (!fila[2]) continue; // Name vacio = fila vacia
    var commodityFila = String(fila[3] || "").trim().toLowerCase();
    if (commodityFila !== commodityBuscado) continue;

    proveedores.push({ proveedor: fila[2], contacto: fila[6], correo: fila[7] });
  }
  return proveedores;
}

// Correr UNA SOLA VEZ desde el editor de Apps Script si tenias datos en el
// formato viejo (4 columnas) sin migrar. Si ya reacomodaste la hoja a mano,
// no necesitas correr esta funcion.
function migrarHojaProveedoresANuevoFormato() {
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var hoja = ss.getSheetByName("Proveedores");
  if (!hoja) return "No existe la hoja Proveedores, no hay nada que migrar.";

  var datos = hoja.getDataRange().getValues();
  if (datos.length === 0) {
    hoja.getRange(1, 1, 1, 8).setValues([ENCABEZADOS_PROVEEDORES]);
    hoja.getRange(1, 1, 1, 8).setFontWeight("bold");
    hoja.setFrozenRows(1);
    return "Hoja vacia: se dejaron solo los encabezados nuevos.";
  }

  var encabezado = datos[0].map(function(v) { return String(v).trim(); });
  var yaMigrada = encabezado[0] === "Nomenclatura" && encabezado[2] === "Name";
  if (yaMigrada) {
    hoja.getRange(1, 1, 1, 8).setFontWeight("bold");
    hoja.setFrozenRows(1);
    return "La hoja Proveedores ya tiene el formato nuevo.";
  }

  var filasNuevas = [ENCABEZADOS_PROVEEDORES];
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var proveedor = fila[0]; // antes: Proveedor
    var contacto = fila[1];  // antes: Contacto
    var correo = fila[2];    // antes: Correo
    var commodity = fila[3]; // antes: Commodity
    if (!proveedor && !correo) continue;
    filasNuevas.push(["", "", proveedor, commodity, "", "", contacto, correo]);
  }

  hoja.clearContents();
  hoja.getRange(1, 1, filasNuevas.length, 8).setValues(filasNuevas);
  hoja.getRange(1, 1, 1, 8).setFontWeight("bold");
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 8);

  return "Migradas " + (filasNuevas.length - 1) + " filas al nuevo formato.";
}
