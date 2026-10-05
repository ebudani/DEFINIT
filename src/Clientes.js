/**
 * Vista CLIENTES: lee las pestañas clientes / clientes_periodo / paquetes / items de
 * "DEFINIT – Base comercial", que arma el procesador (procesador/Procesador.js) a partir de
 * Venta_Items_Sesiones. Ahí solo hay cantidades por mes (o período) y local: ningún nombre.
 *
 * Se lee con el permiso de solo lectura de hojas de cálculo, con la cuenta de quien abre el
 * tablero: ve los datos quien tenga acceso a esa planilla.
 */

// "DEFINIT – Base comercial" (la crea el procesador en el Drive de quien lo configuró).
var BASE_COMERCIAL_ID = '12loXy63CpNzdfJPyACAqFrvR9zWsz-LFr2UuvlAOkQw';

function getClientes() {
  var pestanas = ['clientes', 'clientes_periodo', 'paquetes', 'items', 'info'];
  var lectura;
  try {
    lectura = leerPestanas(pestanas);
  } catch (e) {
    var msg = String(e && e.message || e);
    if (/permission|permiso|403|404|not found|no se encontr/i.test(msg)) return JSON.stringify({ sinAcceso: true });
    // Mientras el procesador no corrió con la versión nueva, la pestaña por período no existe.
    if (/parse range|clientes_periodo/i.test(msg)) lectura = leerPestanas(pestanas.filter(function (p) { return p !== 'clientes_periodo'; }));
    else throw e;
  }
  var tablas = {};
  lectura.valueRanges.forEach(function (vr) {
    var nombre = vr.range.split('!')[0].replace(/^'|'$/g, '');
    // Sheets puede haber guardado "2026-01" como fecha: llega como número de serie.
    var columnasMes = nombre === 'clientes_periodo' ? 2 : 1;
    tablas[nombre] = (vr.values || []).map(function (f, i) {
      for (var c = 0; i > 0 && c < columnasMes; c++) if (typeof f[c] === 'number') f[c] = serialAFecha(f[c]).slice(0, 7);
      return f;
    });
  });

  var equivalencias = {};
  ['clientes', 'clientes_periodo', 'paquetes', 'items'].forEach(function (t) {
    var col = t === 'clientes_periodo' ? 2 : 1;
    (tablas[t] || []).slice(1).forEach(function (f) {
      var local = String(f[col] || '');
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
    periodo: tablas.clientes_periodo || null,
    paquetes: tablas.paquetes || [],
    items: tablas.items || [],
    archivo: archivo[1] || 'Venta_Items_Sesiones',
    actualizadoArchivo: archivo[2] || ''
  });
}

function leerPestanas(nombres) {
  return Sheets.Spreadsheets.Values.batchGet(BASE_COMERCIAL_ID, { ranges: nombres, valueRenderOption: 'UNFORMATTED_VALUE' });
}
