function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var p = e.parameter;
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
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
