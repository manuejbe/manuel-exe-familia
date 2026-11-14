/**
 * Confirmados Familia — MANUEL.EXE (19/12)
 * Pegar en: Google Sheet > Extensiones > Apps Script
 *
 * 1) Ejecutar una vez la función `setup` (te pide permisos: aceptalos).
 * 2) Implementar > Nueva implementación > Tipo: Aplicación web
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 * 3) Copiar la URL que termina en /exec y pegarla en CONFIG.SCRIPT_URL del index.html
 */

const SHEET_NAME = 'Respuestas';
const HEADERS = ['Fecha y hora', 'Nombre', '¿Viene?', 'Cantidad de personas', 'Restricción alimentaria', 'Mensaje', 'Dispositivo'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const sh = getSheet_();
    sh.appendRow([
      new Date(),
      clean_(p.nombre),
      clean_(p.viene),
      Math.max(0, parseInt(p.cantidad, 10) || 0),
      clean_(p.restriccion),
      clean_(p.mensaje),
      clean_(p.ua),
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Para probar que la URL anda: abrila en el navegador.
function doGet() {
  return json_({ ok: true, msg: 'MANUEL.EXE escuchando. Usá POST.' });
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('America/Argentina/Buenos_Aires');
  const sh = getSheet_();
  // Encabezados siempre al día (por si la hoja se creó con una versión anterior)
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#000080').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 150);
  sh.setColumnWidths(2, 2, 170);
  sh.setColumnWidth(4, 150);
  sh.setColumnWidths(5, 2, 300);
  sh.setColumnWidth(7, 200);
  sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');

  // Pestaña Resumen con fórmulas (se actualiza sola)
  let r = ss.getSheetByName('Resumen');
  if (!r) r = ss.insertSheet('Resumen', 0);
  r.clear();
  const R = SHEET_NAME;
  r.getRange('A1:B6').setValues([
    ['Resumen', ''],
    ['Total de personas que vienen', `=SUMIF(${R}!C2:C,"Sí",${R}!D2:D)`],
    ['Respuestas que confirmaron', `=COUNTIF(${R}!C2:C,"Sí")`],
    ['No vienen', `=COUNTIF(${R}!C2:C,"No")`],
    ['Total de respuestas', `=COUNTA(${R}!B2:B)`],
    ['Grupos con restricción alimentaria', `=COUNTIFS(${R}!C2:C,"Sí",${R}!E2:E,"<>Ninguna",${R}!E2:E,"<>")`],
  ]);
  r.getRange('A8').setValue('Restricciones de los que vienen');
  r.getRange('A9').setFormula(
    `=IFERROR(FILTER({${R}!B2:B,${R}!E2:E},${R}!C2:C="Sí",${R}!E2:E<>"Ninguna",${R}!E2:E<>""),"Sin restricciones por ahora")`
  );
  r.getRange('A1').setFontSize(14).setFontWeight('bold');
  r.getRange('A2:A6').setFontWeight('bold');
  r.getRange('A8').setFontWeight('bold');
  r.setColumnWidth(1, 260);
  r.setColumnWidth(2, 300);

  // Borrar la hoja vacía por defecto si existe
  ['Hoja 1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#000080').setFontColor('#ffffff');
  }
  return sh;
}

// Limita largo y evita que alguien escriba una fórmula (=, +, -, @) en la planilla
function clean_(v) {
  v = String(v == null ? '' : v).trim().slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
