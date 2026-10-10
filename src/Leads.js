/**
 * Vista LEADS: lee la pestaña "leads" de "DEFINIT – Base comercial", que arma el procesador
 * (procesador/Procesador.js) a partir del Excel de leads. Ahí solo hay cantidades por día, local,
 * origen y calificación: ningún nombre, mail ni teléfono.
 */
function getLeads() {
  var lectura;
  try {
    lectura = Sheets.Spreadsheets.Values.batchGet(BASE_COMERCIAL_ID, { ranges: ['leads', 'info'], valueRenderOption: 'UNFORMATTED_VALUE' });
  } catch (e) {
    var msg = String(e && e.message || e);
    if (/parse range|leads/i.test(msg)) return JSON.stringify({ sinDatos: true });
    if (/permission|permiso|403|404|not found|no se encontr/i.test(msg)) return JSON.stringify({ sinAcceso: true });
    throw e;
  }
  var leads = lectura.valueRanges[0].values || [];
  var info = lectura.valueRanges[1].values || [];
  var filas = leads.slice(1).map(function (f) {
    // Sheets puede haber guardado el día como fecha: llega como número de serie.
    if (typeof f[0] === 'number') f[0] = serialAFecha(f[0]);
    return [String(f[0]).slice(0, 10), String(f[1] || ''), String(f[2] || ''), String(f[3] || ''), Number(f[4]) || 0];
  }).filter(function (f) { return /^\d{4}-\d{2}-\d{2}$/.test(f[0]); });
  var archivo = info.filter(function (f) { return f[0] === 'leads'; })[0] || [];
  if (typeof archivo[2] === 'number') {
    archivo[2] = Utilities.formatDate(new Date(Date.UTC(1899, 11, 30) + Math.round(archivo[2] * 86400000)), 'UTC', 'yyyy-MM-dd HH:mm');
  }
  return JSON.stringify({ filas: filas, archivo: archivo[1] || 'Leads', actualizadoArchivo: archivo[2] || '' });
}
