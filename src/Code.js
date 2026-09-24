/**
 * Tablero de ventas DEFINIT — web app de Google Apps Script.
 *
 * Corre con la cuenta de quien lo despliega y lee "Facturación 2026" en vivo
 * cada vez que se abre o se toca "Actualizar": no hay copias de los datos.
 */

// ID de "Facturación 2026" (docs.google.com/spreadsheets/d/<ID>/edit).
var SPREADSHEET_ID = '1vYjYEeJFU6IJlHv0Tnrj2UZsOC0FrUlpNtpxq6PDLMk';
var PESTANA_RESUMEN = 'Ventas Mensuales';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('DEFINIT: Tablero Mensual de Ventas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

/**
 * Lee la planilla y devuelve el modelo como JSON (google.script.run no admite
 * objetos Date). Solo se leen "Ventas Mensuales" y las pestañas mensuales; el
 * resto (por ejemplo la de clientes con DNI) nunca se pide.
 *
 * Usa la API de Sheets (servicio avanzado "Sheets") en lugar de SpreadsheetApp
 * porque SpreadsheetApp.openById exige el permiso de edición; la API funciona
 * con spreadsheets.readonly, así el tablero no puede modificar la planilla.
 */
function getData() {
  var info = Sheets.Spreadsheets.get(SPREADSHEET_ID, { fields: 'properties/title,sheets/properties/title' });
  var nombres = info.sheets.map(function (s) { return s.properties.title; });
  if (nombres.indexOf(PESTANA_RESUMEN) < 0) throw new Error('No se encontró la pestaña "' + PESTANA_RESUMEN + '".');
  var aLeer = nombres.filter(function (n) { return n === PESTANA_RESUMEN || esPestanaMensual(n); });

  // Solo valores (sin formatos): es una lectura liviana. Las fechas llegan como
  // número de serie y marcarFechas() las reconoce por su posición en la hoja.
  var lectura = Sheets.Spreadsheets.Values.batchGet(SPREADSHEET_ID, {
    ranges: aLeer.map(function (n) { return "'" + n.replace(/'/g, "''") + "'"; }),
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER'
  });

  var hojas = {};
  lectura.valueRanges.forEach(function (vr, i) {
    var nombre = aLeer[i];
    hojas[nombre] = marcarFechas(vr.values || [], nombre === PESTANA_RESUMEN);
  });

  var resumen = hojas[PESTANA_RESUMEN];
  delete hojas[PESTANA_RESUMEN];
  var modelo = construirModelo(resumen, hojas);
  modelo.planilla = info.properties.title;
  return JSON.stringify(modelo);
}
