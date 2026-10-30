/**
 * Vista LEADS: lee la pestaña "leads_vista" de "DEFINIT – Base comercial" (la arma el procesador,
 * procesador/Analisis.js). Una fila por mes del lead, local, origen, calificación y tipo de lead, con
 * cantidades: leads, compradores, facturación y días hasta la compra. Ningún nombre, mail ni teléfono.
 */
function getLeads() {
  var lectura;
  try {
    lectura = Sheets.Spreadsheets.Values.batchGet(BASE_COMERCIAL_ID, { ranges: ['leads_vista'], valueRenderOption: 'UNFORMATTED_VALUE' });
  } catch (e) {
    var msg = String(e && e.message || e);
    if (SIN_AUTORIZAR.test(msg)) throw new Error(MSG_AUTORIZAR + ' (' + msg + ')');
    if (/parse range|leads_vista/i.test(msg)) return JSON.stringify({ sinDatos: true });
    if (SIN_ACCESO.test(msg)) return JSON.stringify({ sinAcceso: true, detalle: msg });
    throw e;
  }
  var filas = lectura.valueRanges[0].values || [];
  return JSON.stringify({ vista: filas.map(function (f) {
    // Sheets puede haber guardado el mes como fecha: llega como número de serie.
    if (typeof f[1] === 'number') f[1] = serialAFecha(f[1]).slice(0, 7);
    return f;
  }) });
}
