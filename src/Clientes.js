/**
 * Vista CLIENTES: lee las pestañas clientes / paquetes / items de "DEFINIT – Base comercial",
 * que arma el procesador (procesador/Procesador.js) a partir de Venta_Items_Sesiones.
 * Ahí solo hay cantidades por mes y local: ningún nombre de cliente.
 *
 * Se lee con el permiso de solo lectura de hojas de cálculo, con la cuenta de quien abre el
 * tablero: ve los datos quien tenga acceso a esa planilla.
 */

// "DEFINIT – Base comercial" (la crea el procesador en el Drive de quien lo configuró).
var BASE_COMERCIAL_ID = '12loXy63CpNzdfJPyACAqFrvR9zWsz-LFr2UuvlAOkQw';

function getClientes() {
  var lectura;
  try {
    lectura = Sheets.Spreadsheets.Values.batchGet(BASE_COMERCIAL_ID, {
      ranges: ['clientes', 'paquetes', 'items', 'info'], valueRenderOption: 'UNFORMATTED_VALUE'
    });
  } catch (e) {
    var msg = String(e && e.message || e);
    if (/permission|permiso|403|404|not found|no se encontr/i.test(msg)) return JSON.stringify({ sinAcceso: true });
    throw e;
  }
  var tablas = {};
  lectura.valueRanges.forEach(function (vr) {
    var nombre = vr.range.split('!')[0].replace(/^'|'$/g, '');
    // Sheets puede haber guardado "2026-01" como fecha: llega como número de serie.
    tablas[nombre] = (vr.values || []).map(function (f, i) {
      if (i > 0 && typeof f[0] === 'number') f[0] = serialAFecha(f[0]).slice(0, 7);
      return f;
    });
  });

  var equivalencias = {};
  ['clientes', 'paquetes', 'items'].forEach(function (t) {
    (tablas[t] || []).slice(1).forEach(function (f) {
      var local = String(f[1] || '');
      if (local && !(local in equivalencias)) equivalencias[local] = local === 'TOTAL' ? null : claveFacturacion(local);
    });
  });
  var meses = {};
  (tablas.clientes || []).slice(1).forEach(function (f) { if (f[0]) meses[f[0]] = true; });

  var archivo = (tablas.info || []).filter(function (f) { return f[0] === 'sesiones'; })[0] || [];
  return JSON.stringify({
    meses: Object.keys(meses).sort(),
    equivalencias: equivalencias,
    clientes: tablas.clientes || [],
    paquetes: tablas.paquetes || [],
    items: tablas.items || [],
    archivo: archivo[1] || 'Venta_Items_Sesiones',
    actualizadoArchivo: archivo[2] || ''
  });
}
