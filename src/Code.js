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
 * resto (por ejemplo la de clientes con DNI) nunca se abre.
 */
function getData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hojaResumen = ss.getSheetByName(PESTANA_RESUMEN);
  if (!hojaResumen) throw new Error('No se encontró la pestaña "' + PESTANA_RESUMEN + '".');

  var pestanasMes = {};
  ss.getSheets().forEach(function (hoja) {
    var nombre = hoja.getName();
    if (esPestanaMensual(nombre)) pestanasMes[nombre] = hoja.getDataRange().getValues();
  });

  var modelo = construirModelo(hojaResumen.getDataRange().getValues(), pestanasMes);
  modelo.planilla = ss.getName();
  return JSON.stringify(modelo);
}
