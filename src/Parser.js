/**
 * Convierte las grillas crudas de la planilla "Facturación 2026" en un modelo
 * de datos limpio para el tablero.
 *
 * Funciona igual en Apps Script (V8) y en Node (para pruebas locales):
 * recibe matrices de valores tal como las devuelve Range.getValues().
 *
 * Estructura esperada de la planilla:
 *  - "Ventas Mensuales": fila 1 con la fecha de actualización en A1 y la fecha
 *    de cada mes en bloques de 9 columnas (E1, N1, W1, ...). Cada bloque arranca
 *    3 columnas antes de la fecha: Local | Venta | Meta | % | Súper | % | ...
 *    Primero los locales propios (encabezado "Local"), luego "FRANQUICIAS".
 *  - Una pestaña por mes ("Septiembre26", "Agosto26", ...): por cada local una
 *    fila con el nombre y las fechas, y debajo la fila "Ventas" con los importes.
 *  - Cualquier otra pestaña (por ejemplo "JANET", con datos personales) se ignora.
 */

var MESES_ES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
  'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

var PESTANA_MES_RE = /^(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\s*(\d{2}|\d{4})$/i;

function esPestanaMensual(nombre) {
  return PESTANA_MES_RE.test(String(nombre).trim());
}

/** Devuelve 'YYYY-MM-DD' para un Date o un string ISO; null si no es fecha. */
function aFechaISO(v) {
  if (v instanceof Date && !isNaN(v)) {
    return v.getFullYear() + '-' + pad2(v.getMonth() + 1) + '-' + pad2(v.getDate());
  }
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
  return null;
}

function pad2(n) { return (n < 10 ? '0' : '') + n; }

function sumarDias(fechaISO, n) {
  var p = fechaISO.split('-');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
}

function aNumero(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (typeof v !== 'string') return null;
  var s = v.replace(/[$\s]/g, '').replace(/\./g, '').replace(',', '.');
  if (s === '' || s === '-') return null;
  var n = Number(s);
  return isFinite(n) ? n : null;
}

function texto(v) { return v == null ? '' : String(v).trim(); }

/** Clave estable para cruzar locales entre pestañas ("Lanus" == "Lanús"). */
function claveLocal(nombre) {
  return texto(nombre).toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ');
}

/**
 * Convierte la GridData de la API de Sheets en una matriz como la de
 * Range.getValues(): números, textos y fechas (las celdas con formato de fecha
 * llegan como número de serie y se devuelven como 'YYYY-MM-DD').
 */
function grillaAValores(grilla) {
  if (!grilla || !grilla.rowData) return [];
  var desdeFila = grilla.startRow || 0, desdeCol = grilla.startColumn || 0;
  var valores = [];
  for (var i = 0; i < desdeFila; i++) valores.push([]);
  grilla.rowData.forEach(function (fila) {
    var salida = [];
    for (var j = 0; j < desdeCol; j++) salida.push('');
    (fila.values || []).forEach(function (celda) {
      var ev = celda.effectiveValue || {};
      var tipo = celda.effectiveFormat && celda.effectiveFormat.numberFormat && celda.effectiveFormat.numberFormat.type;
      if (ev.numberValue != null) {
        salida.push(tipo === 'DATE' || tipo === 'DATE_TIME' ? serialAFecha(ev.numberValue) : ev.numberValue);
      } else if (ev.stringValue != null) salida.push(ev.stringValue);
      else if (ev.boolValue != null) salida.push(ev.boolValue);
      else salida.push('');
    });
    valores.push(salida);
  });
  return valores;
}

/** Número de serie de Sheets (días desde 30/12/1899) -> 'YYYY-MM-DD'. */
function serialAFecha(serial) {
  var d = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000);
  return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
}

/** Lee la pestaña "Ventas Mensuales": un bloque por mes con locales, metas y súper. */
function parsearResumen(valores) {
  var fila0 = valores[0] || [];
  var actualizado = aFechaISO(fila0[0]);
  var meses = [];

  for (var c = 1; c < fila0.length; c++) {
    var fechaMes = aFechaISO(fila0[c]);
    if (!fechaMes) continue;
    var ini = c - 3; // columna "Local" del bloque
    if (ini < 0) continue;

    var mes = { mes: fechaMes.slice(0, 7), diasRemanentes: null, locales: [] };
    var tipo = null;

    for (var r = 1; r < valores.length; r++) {
      var fila = valores[r] || [];
      var etiqueta = texto(fila[ini]);
      var filaTexto = fila.slice(ini, ini + 9).map(texto).join('|').toUpperCase();

      if (filaTexto.indexOf('FRANQUICIAS') >= 0) { tipo = 'franquicia'; continue; }
      if (etiqueta === 'Local') { if (!tipo) tipo = 'propio'; continue; }
      if (/^d[ií]as remanentes$/i.test(etiqueta)) {
        mes.diasRemanentes = aNumero(fila[ini + 1]);
        continue;
      }
      if (!tipo || !etiqueta || /^total/i.test(etiqueta)) continue;

      var venta = aNumero(fila[ini + 1]);
      var meta = aNumero(fila[ini + 2]);
      var superMeta = aNumero(fila[ini + 4]);
      if (venta == null && meta == null) continue;

      mes.locales.push({
        nombre: etiqueta,
        clave: claveLocal(etiqueta),
        tipo: claveLocal(etiqueta) === 'ecom' ? 'ecom' : tipo,
        venta: venta || 0,
        meta: meta || 0,
        super: superMeta || 0
      });
    }
    if (mes.locales.length) meses.push(mes);
  }

  meses.sort(function (a, b) { return a.mes < b.mes ? -1 : 1; });
  return { actualizado: actualizado, meses: meses };
}

/** Lee una pestaña mensual: ventas diarias por local. Devuelve { mes, locales: {clave: {nombre, dias}} }. */
function parsearMes(valores) {
  var locales = {};
  var mes = null;

  for (var r = 0; r < valores.length - 1; r++) {
    var fila = valores[r] || [];
    var siguiente = valores[r + 1] || [];
    var nombre = texto(fila[0]);
    if (!nombre || !/^ventas/i.test(texto(siguiente[0]))) continue;

    // Hay celdas de encabezado escritas a mano ("miercoles 1-04") o pisadas con
    // un importe: la fecha de cada columna se deduce de la primera fecha válida
    // de la fila (una columna = un día).
    var ancla = -1, fechaAncla = null;
    for (var a = 1; a < fila.length && ancla < 0; a++) {
      var f = aFechaISO(fila[a]);
      if (f) { ancla = a; fechaAncla = f; }
    }
    if (ancla < 0) continue;

    var dias = {};
    for (var c = 1; c < fila.length; c++) {
      if (/^totales?$/i.test(texto(fila[c]))) break;
      var fecha = sumarDias(fechaAncla, c - ancla);
      if (!mes) mes = fechaAncla.slice(0, 7);
      if (fecha.slice(0, 7) !== mes) continue;
      dias[fecha] = aNumero(siguiente[c]) || 0;
    }
    locales[claveLocal(nombre)] = { nombre: nombre, dias: dias };
  }
  return { mes: mes, locales: locales };
}

/**
 * Arma el modelo completo.
 * @param {Array[]} resumen  valores de "Ventas Mensuales"
 * @param {Object<string, Array[]>} pestanasMes  nombre de pestaña -> valores
 */
function construirModelo(resumen, pestanasMes) {
  var modelo = parsearResumen(resumen);
  var diario = {};

  Object.keys(pestanasMes).forEach(function (nombre) {
    if (!esPestanaMensual(nombre)) return;
    var m = parsearMes(pestanasMes[nombre]);
    if (!m.mes) return;
    var porLocal = {};
    Object.keys(m.locales).forEach(function (k) {
      var dias = m.locales[k].dias;
      porLocal[k] = Object.keys(dias).sort().map(function (f) { return [f, dias[f]]; });
    });
    diario[m.mes] = porLocal;
  });

  modelo.diario = diario;
  modelo.generado = new Date().toISOString();
  return modelo;
}

if (typeof module !== 'undefined') {
  module.exports = {
    construirModelo: construirModelo, parsearResumen: parsearResumen,
    parsearMes: parsearMes, esPestanaMensual: esPestanaMensual, claveLocal: claveLocal,
    grillaAValores: grillaAValores
  };
}
