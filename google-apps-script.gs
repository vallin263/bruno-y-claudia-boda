function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var p = e.parameter;
  // El panel privado (panel.html) pide los datos con accion=panel.
  if (p.accion === 'panel') {
    return datosPanel(p, ss, sheet);
  }
  sheet.appendRow([
    new Date(),
    p.Asistencia || '',
    p.Nombre || '',
    p.Acompanante || '',
    p.Nombre_Acompanante || '',
    p.Autobus || '',
    p.AutobusSentido || '',
    p.Parada || '',
    p.ParadaHora || '',
    p.Alergenos || '',
    p.Cancion || '',
    p.Mensaje || ''
  ]);
  // Aviso por correo: si falla, la confirmacion ya esta guardada en la hoja.
  try {
    avisarPorCorreo(p, ss.getUrl());
  } catch (err) {
    console.error('No se pudo enviar el aviso por correo: ' + err);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Devuelve las confirmaciones solo si la contrasena coincide con la
// propiedad CLAVE_PANEL (Configuracion del proyecto > Propiedades de la
// secuencia de comandos). Sin esa propiedad, el panel queda cerrado.
function datosPanel(p, ss, sheet) {
  var clave = PropertiesService.getScriptProperties().getProperty('CLAVE_PANEL');
  if (!clave || p.clave !== clave) {
    Utilities.sleep(1500);
    return responder({ ok: false });
  }
  var valores = sheet.getDataRange().getValues();
  var filas = [];
  for (var i = 0; i < valores.length; i++) {
    var f = valores[i];
    if (!(f[0] instanceof Date)) continue; // cabecera o fila vacia
    f[0] = f[0].toISOString();
    filas.push(f);
  }
  return responder({ ok: true, filas: filas, hoja: ss.getUrl() });
}

var CORREO_AVISOS ='brunoyclauboda@gmail.com';

function limpiar(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function avisarPorCorreo(p, urlHoja) {
  var asiste = String(p.Asistencia || '').toLowerCase() === 'si';
  var nombre = p.Nombre || 'Sin nombre';
  var datos = [
    ['Asistencia', asiste ? 'Sí, asiste' : 'No asiste'],
    ['Nombre', p.Nombre],
    ['Acompañante', p.Nombre_Acompanante || p.Acompanante],
    ['Autobús', p.Autobus],
    ['Parada', p.Parada],
    ['Alergias', p.Alergenos],
    ['Canción', p.Cancion],
    ['Mensaje', p.Mensaje]
  ].filter(function (d) { return d[1]; });

  var filas = datos.map(function (d) {
    return '<tr>' +
      '<td style="padding:8px 16px 8px 0;color:#8a7f72;font-size:13px;vertical-align:top;white-space:nowrap;">' + limpiar(d[0]) + '</td>' +
      '<td style="padding:8px 0;color:#2f2a25;font-size:15px;">' + limpiar(d[1]) + '</td>' +
      '</tr>';
  }).join('');

  var html =
    '<div style="background:#f6f1e9;padding:28px 12px;font-family:Georgia,serif;">' +
    '<div style="max-width:520px;margin:0 auto;background:#ffffff;border-radius:10px;padding:28px;">' +
    '<p style="margin:0 0 6px;color:#b5654a;font-size:12px;letter-spacing:2px;text-transform:uppercase;">Nueva confirmación</p>' +
    '<h1 style="margin:0 0 4px;color:#2f2a25;font-size:24px;font-weight:normal;">' + limpiar(nombre) + '</h1>' +
    '<p style="margin:0 0 20px;color:' + (asiste ? '#4f7a4a' : '#b5654a') + ';font-size:16px;">' + (asiste ? 'Sí, asiste' : 'No asiste') + '</p>' +
    '<table style="border-collapse:collapse;width:100%;border-top:1px solid #ece4d8;">' + filas + '</table>' +
    '<p style="margin:26px 0 0;"><a href="' + urlHoja + '" style="display:inline-block;background:#b5654a;color:#ffffff;text-decoration:none;padding:12px 22px;border-radius:6px;font-family:Arial,sans-serif;font-size:14px;">Abrir todas las confirmaciones</a></p>' +
    '</div></div>';

  var texto = datos.map(function (d) { return d[0] + ': ' + d[1]; }).join(String.fromCharCode(10)) +
    String.fromCharCode(10, 10) + 'Todas las confirmaciones: ' + urlHoja;

  MailApp.sendEmail({
    to: CORREO_AVISOS,
    subject: 'Boda · ' + nombre + ' · ' + (asiste ? 'Sí, asiste' : 'No asiste'),
    body: texto,
    htmlBody: html,
    name: 'Web de la boda'
  });
}
