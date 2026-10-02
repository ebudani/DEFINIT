/**
 * DEFINIT – Procesador comercial (proyecto de Apps Script aparte del tablero).
 *
 * Cada hora revisa la carpeta compartida de los Excel comerciales. Si cambió alguno de
 * los archivos conocidos, lo convierte a una hoja de cálculo temporal en los servidores
 * de Google, calcula los totales por local y mes, los escribe en la planilla
 * "DEFINIT – Base comercial" y borra la copia temporal.
 *
 * El tablero solo lee esa planilla (permiso de solo lectura de hojas de cálculo), así que
 * nunca abre los Excel ni necesita permiso de Drive. Quién ve los indicadores comerciales
 * se decide compartiendo esa planilla.
 *
 * Una sola vez: abrir este proyecto en script.google.com y ejecutar configurar().
 */

var CARPETA_ID = '1rqvtYLmOmLqtuvovHk4jIZJIZVhg3TpB';
var NOMBRE_BASE = 'DEFINIT – Base comercial';

// Qué archivo de la carpeta alimenta cada pestaña de la base.
var FUENTES = [
  { id: 'cobros', patron: /^descuentos por locales.*\.xlsx$/i, procesar: procesarCobros },
  { id: 'sesiones', patron: /^venta_items_sesiones.*\.xlsx$/i, procesar: procesarSesiones },
  { id: 'agenda', patron: /^minutos x mes x sucursal.*\.xlsx$/i, procesar: procesarAgenda }
];

/** Instala la revisión automática (cada hora) y procesa todo por primera vez. */
function configurar() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'procesar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('procesar').timeBased().everyHours(1).create();
  PropertiesService.getScriptProperties().deleteProperty('procesados');
  procesar();
  Logger.log('Listo. Base comercial: ' + base().getUrl());
}

/** Revisa la carpeta y procesa los archivos que cambiaron desde la última vez. */
function procesar() {
  var props = PropertiesService.getScriptProperties();
  var hechos = JSON.parse(props.getProperty('procesados') || '{}');
  var archivos = listarCarpeta();
  var libro = base();
  var info = [];

  FUENTES.forEach(function (fuente) {
    var archivo = masReciente(archivos, fuente.patron);
    if (!archivo) { info.push([fuente.id, '(no hay archivo en la carpeta)', '', '']); return; }
    var firma = archivo.getId() + ':' + archivo.getLastUpdated().getTime();
    if (hechos[fuente.id] === firma) {
      info.push([fuente.id, archivo.getName(), fecha(archivo.getLastUpdated()), 'sin cambios']);
      return;
    }
    var inicio = Date.now();
    var temporal = convertir(archivo);
    try {
      var hojas = SpreadsheetApp.openById(temporal).getSheets();
      diagnosticar(libro, fuente.id, archivo, hojas);
      // Cada fuente devuelve una tabla, o varias como { pestaña: filas }.
      var tablas = fuente.procesar(hojas);
      if (Array.isArray(tablas)) { var una = {}; una[fuente.id] = tablas; tablas = una; }
      var resumen = Object.keys(tablas).map(function (nombre) {
        escribir(libro, nombre, tablas[nombre]);
        return nombre + ' ' + (tablas[nombre].length - 1);
      });
      hechos[fuente.id] = firma;
      props.setProperty('procesados', JSON.stringify(hechos));
      info.push([fuente.id, archivo.getName(), fecha(archivo.getLastUpdated()),
        'procesado ' + fecha(new Date()) + ' · filas: ' + resumen.join(', ') + ' · ' + Math.round((Date.now() - inicio) / 1000) + ' s']);
    } finally {
      Drive.Files.remove(temporal);
    }
  });

  escribir(libro, 'info', [['fuente', 'archivo', 'actualizado', 'estado']].concat(info));
}

// ---------- Drive ----------

function listarCarpeta() {
  var res = [], it = DriveApp.getFolderById(CARPETA_ID).getFiles();
  while (it.hasNext()) res.push(it.next());
  return res;
}

function masReciente(archivos, patron) {
  return archivos.filter(function (f) { return patron.test(f.getName()); })
    .sort(function (a, b) { return b.getLastUpdated() - a.getLastUpdated(); })[0] || null;
}

/** Copia el .xlsx como hoja de cálculo de Google (la conversión la hace Drive). Devuelve el ID. */
function convertir(archivo) {
  var copia = Drive.Files.copy({ name: 'tmp – ' + archivo.getName(), mimeType: MimeType.GOOGLE_SHEETS }, archivo.getId());
  return copia.id;
}

/** Planilla de resultados: se crea la primera vez en el Drive de quien ejecuta el proceso. */
function base() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('baseId');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* se borró: se crea de nuevo */ } }
  var libro = SpreadsheetApp.create(NOMBRE_BASE);
  props.setProperty('baseId', libro.getId());
  return libro;
}

function escribir(libro, nombre, filas) {
  var hoja = libro.getSheetByName(nombre) || libro.insertSheet(nombre);
  hoja.clearContents();
  if (!filas.length) return;
  var ancho = filas.reduce(function (m, f) { return Math.max(m, f.length); }, 0);
  var datos = filas.map(function (f) { while (f.length < ancho) f.push(''); return f; });
  // Primera columna como texto: si no, Sheets convierte "2026-01" en una fecha.
  hoja.getRange(1, 1, datos.length, 1).setNumberFormat('@');
  hoja.getRange(1, 1, datos.length, ancho).setValues(datos);
  var sobra = libro.getSheetByName('Hoja 1') || libro.getSheetByName('Sheet1');
  if (sobra && libro.getSheets().length > 1) libro.deleteSheet(sobra);
}

/** Encabezados y primeras filas de cada pestaña del Excel, para revisar su formato. */
function diagnosticar(libro, fuenteId, archivo, hojas) {
  var filas = [['fuente', 'archivo', 'pestaña', 'filas', 'columnas', 'fila', 'valores']];
  var previas = (libro.getSheetByName('diagnostico') || { getDataRange: function () { return { getValues: function () { return []; } }; } })
    .getDataRange().getValues().slice(1).filter(function (f) { return f[0] && f[0] !== fuenteId; });
  hojas.forEach(function (h) {
    var n = Math.min(4, h.getLastRow());
    var vals = n ? h.getRange(1, 1, n, Math.min(h.getLastColumn(), 20)).getValues() : [];
    vals.forEach(function (v, i) {
      filas.push([fuenteId, archivo.getName(), h.getName(), h.getLastRow(), h.getLastColumn(), i + 1,
        v.map(function (x) { return x instanceof Date ? fecha(x) : x; }).join(' | ')]);
    });
  });
  escribir(libro, 'diagnostico', [filas[0]].concat(previas, filas.slice(1)));
}

function fecha(d) { return Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm'); }

// ---------- Lectura genérica ----------

/**
 * Busca la pestaña cuyo encabezado (en las primeras 10 filas) tiene las columnas pedidas
 * y lee solo esas columnas. Las claves que terminan en "?" son opcionales.
 * Devuelve { n, col: { clave: [valores] } }.
 */
function tablaCon(hojas, columnas) {
  for (var i = 0; i < hojas.length; i++) {
    var h = hojas[i], ultima = h.getLastRow(), ancho = h.getLastColumn();
    if (!ultima || !ancho) continue;
    var cabecera = h.getRange(1, 1, Math.min(10, ultima), ancho).getValues();
    for (var r = 0; r < cabecera.length; r++) {
      var idx = indices(cabecera[r], columnas);
      if (!idx) continue;
      var n = ultima - (r + 1), col = {};
      Object.keys(idx).forEach(function (k) {
        col[k] = n > 0 ? h.getRange(r + 2, idx[k] + 1, n, 1).getValues().map(function (v) { return v[0]; }) : [];
      });
      return { n: n, col: col };
    }
  }
  throw new Error('No encontré una pestaña con las columnas: ' + Object.keys(columnas).join(', '));
}

function indices(encabezado, columnas) {
  var idx = {};
  Object.keys(columnas).forEach(function (k) {
    for (var i = 0; i < encabezado.length; i++) if (columnas[k].test(String(encabezado[i] || '').trim())) { idx[k] = i; break; }
  });
  return Object.keys(columnas).every(function (k) { return idx[k] != null || /\?$/.test(k); }) ? idx : null;
}

function mesDe(v) {
  if (v instanceof Date) return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2);
  var m = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2);
  m = String(v || '').match(/^(\d{4})-(\d{2})/);
  return m ? m[1] + '-' + m[2] : null;
}

function num(v) {
  if (typeof v === 'number') return v;
  var s = String(v || '').trim();
  if (!s || s === '-') return 0;
  s = s.replace(/[$\s]/g, '');
  // "139,748" (miles con coma) o "1.234,56" (formato argentino)
  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
  else s = s.replace(/\./g, '').replace(',', '.');
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

function nombreLocal(s) {
  return String(s || '').replace(/^AR\s*-\s*DEFINIT\s*-\s*/i, '').trim().toLowerCase()
    .replace(/(^|\s)\S/g, function (x) { return x.toUpperCase(); });
}

function acumular(mapa, clave, campos) {
  var x = mapa[clave] = mapa[clave] || {};
  Object.keys(campos).forEach(function (k) { x[k] = (x[k] || 0) + campos[k]; });
}

function aFilas(mapa, encabezado) {
  var filas = [encabezado];
  Object.keys(mapa).sort().forEach(function (k) {
    var partes = k.split('|'), x = mapa[k];
    filas.push(partes.concat(encabezado.slice(partes.length).map(function (c) { return Math.round((x[c] || 0) * 100) / 100; })));
  });
  return filas;
}

// ---------- Cobros: "Descuentos por locales" ----------

var CUOTAS = ['1 pago', '3 cuotas', '6 cuotas', '12 cuotas', 'Otras'];
var FORMAS = ['Crédito', 'Débito', 'Web', 'Transferencia', 'Efectivo', 'Otras'];

function cuotaDe(condicion, forma) {
  var m = String(condicion || '').match(/^(\d+)\s+PAGO/i);
  if (m) return { 1: '1 pago', 3: '3 cuotas', 6: '6 cuotas', 12: '12 cuotas' }[+m[1]] || 'Otras';
  return /efectivo|transfer|d[eé]bito/i.test(forma || '') ? '1 pago' : 'Otras';
}

function formaDe(forma) {
  var f = String(forma || '');
  if (f.indexOf(';') >= 0) return 'Otras';
  if (/cr[eé]dito/i.test(f)) return 'Crédito';
  if (/d[eé]bito/i.test(f)) return 'Débito';
  if (/web/i.test(f)) return 'Web';
  if (/transfer/i.test(f)) return 'Transferencia';
  if (/efectivo/i.test(f)) return 'Efectivo';
  return 'Otras';
}

/**
 *  - ops / neto / cuotas / formas: cobros con venta neta > 0 (ticket = neto ÷ 1,21 ÷ ops)
 *  - bruto / desc: todos los cobros, igual que la tabla dinámica del archivo
 */
function procesarCobros(hojas) {
  var t = tablaCon(hojas, { fecha: /^fecha de pago$/i, local: /^establecimiento$/i, bruto: /^v\.?\s*bruto$/i,
    desc: /^v\.?\s*descuento$/i, neto: /^v\.?\s*neto$/i, 'forma?': /^forma de pago$/i, 'condicion?': /^condici[oó]n de pago$/i });
  return agregarCobros(t);
}

function agregarCobros(t) {
  var mapa = {}, c = t.col;
  for (var i = 0; i < t.n; i++) {
    var mes = mesDe(c.fecha[i]), local = nombreLocal(c.local[i]);
    if (!mes || !local) continue;
    var neto = num(c.neto[i]), campos = { bruto: num(c.bruto[i]), desc: num(c.desc[i]) };
    if (neto > 0) {
      var forma = c['forma?'] ? c['forma?'][i] : '';
      campos.ops = 1; campos.neto = neto;
      campos['c:' + cuotaDe(c['condicion?'] ? c['condicion?'][i] : '', forma)] = neto;
      campos['f:' + formaDe(forma)] = neto;
    }
    acumular(mapa, mes + '|' + local, campos);
  }
  return aFilas(mapa, ['mes', 'local', 'ops', 'neto', 'bruto', 'desc']
    .concat(CUOTAS.map(function (c) { return 'c:' + c; }), FORMAS.map(function (c) { return 'f:' + c; })));
}

// ---------- Sesiones: "Venta_Items_Sesiones" ----------

/**
 * Venta de ítems y sesiones: totales por mes y local, y el análisis de clientes.
 * Los nombres de clientes solo se usan acá para contar personas distintas y saber si ya habían
 * comprado antes; a la base comercial (y al tablero) llegan únicamente cantidades.
 */
function procesarSesiones(hojas) {
  var t = tablaCon(hojas, { fecha: /^fecha de pago$/i, local: /^establecimiento$/i, item: /^[íi]tem$/i,
    sesiones: /^n[uú]mero de sesiones$/i, cortesia: /cortes[íi]a/i, bruto: /^v\.?\s*bruto$/i,
    desc: /^v\.?\s*descuento$/i, neto: /^v\.?\s*neto$/i, minutos: /^minutos totales$/i,
    'cliente?': /^cliente$/i, 'paquete?': /^paquete$/i });
  var tablas = { sesiones: agregarSesiones(t) };
  if (t.col['cliente?']) {
    var c = agregarClientes(t);
    tablas.clientes = c.clientes;
    tablas.paquetes = c.paquetes;
    tablas.items = c.items;
  }
  return tablas;
}

function agregarSesiones(t) {
  var mapa = {}, c = t.col;
  for (var i = 0; i < t.n; i++) {
    var mes = mesDe(c.fecha[i]), local = nombreLocal(c.local[i]);
    if (!mes || !local) continue;
    var cortesia = c.cortesia[i] === true || /^(true|verdadero|s[ií]|1)$/i.test(String(c.cortesia[i]).trim());
    var ses = num(c.sesiones[i]), min = num(c.minutos[i]);
    acumular(mapa, mes + '|' + local, cortesia
      ? { items: 1, sesiones_cortesia: ses, minutos_cortesia: min }
      : { items: 1, sesiones: ses, minutos: min, bruto: num(c.bruto[i]), desc: num(c.desc[i]), neto: num(c.neto[i]) });
  }
  return aFilas(mapa, ['mes', 'local', 'items', 'sesiones', 'minutos', 'sesiones_cortesia', 'minutos_cortesia', 'bruto', 'desc', 'neto']);
}

/** Tipo de paquete a partir del nombre ("AXILAS MUJER - CORTESÍA 3 SESIONES", "BOZO MUJER MEDIO PAQUETE"…). */
function tipoPaquete(nombre, cortesia) {
  var p = String(nombre || '').toUpperCase();
  if (cortesia || /CORTES[IÍ]A/.test(p)) return 'Cortesía';
  if (/MEDIO\s+PAQUETE/.test(p)) return 'Medio paquete';
  if (/PAQUETE/.test(p)) return 'Paquete completo';
  if (/SESI[OÓ]N|SUELTA|INDIVIDUAL/.test(p)) return 'Sesión suelta';
  if (/PROMO|COMBO|PACK/.test(p)) return 'Promo / combo';
  return 'Otros';
}

/**
 * Por mes y local (y "TOTAL" para la empresa): clientes distintos, nuevos (primera compra en el
 * archivo) y recurrentes, ítems, sesiones, venta y clientes con cortesía. Además, el mix de paquetes
 * y los ítems (zonas) por mes y local.
 */
function agregarClientes(t) {
  var c = t.col, n = t.n;
  var claveCliente = function (i) { return String(c['cliente?'][i] || '').trim().toUpperCase().replace(/\s+/g, ' '); };
  var primerMes = {};
  for (var i = 0; i < n; i++) {
    var mes = mesDe(c.fecha[i]), cli = claveCliente(i);
    if (!mes || !cli) continue;
    if (!primerMes[cli] || mes < primerMes[cli]) primerMes[cli] = mes;
  }
  var grupos = {}, paquetes = {}, items = {};
  for (i = 0; i < n; i++) {
    mes = mesDe(c.fecha[i]); cli = claveCliente(i);
    var local = nombreLocal(c.local[i]);
    if (!mes || !local || !cli) continue;
    var cortesia = c.cortesia[i] === true || /^(true|verdadero|s[ií]|1)$/i.test(String(c.cortesia[i]).trim());
    var neto = cortesia ? 0 : num(c.neto[i]), ses = num(c.sesiones[i]);
    [local, 'TOTAL'].forEach(function (l) {
      var g = grupos[mes + '|' + l] = grupos[mes + '|' + l] || { clientes: {}, cortesia: {}, items: 0, sesiones: 0, neto: 0 };
      g.clientes[cli] = true; g.items++; g.sesiones += ses; g.neto += neto;
      if (cortesia) g.cortesia[cli] = true;
    });
    var tipo = tipoPaquete(c['paquete?'] ? c['paquete?'][i] : '', cortesia);
    acumular(paquetes, mes + '|' + local + '|' + tipo, { items: 1, sesiones: ses, neto: neto });
    var item = String(c.item[i] || '').trim().toUpperCase() || 'SIN ÍTEM';
    acumular(items, mes + '|' + local + '|' + item, { items: 1, sesiones: ses, neto: neto });
  }
  var filas = [['mes', 'local', 'clientes', 'nuevos', 'recurrentes', 'clientes_cortesia', 'items', 'sesiones', 'neto']];
  Object.keys(grupos).sort().forEach(function (k) {
    var g = grupos[k], partes = k.split('|'), lista = Object.keys(g.clientes);
    var nuevos = lista.filter(function (x) { return primerMes[x] === partes[0]; }).length;
    filas.push([partes[0], partes[1], lista.length, nuevos, lista.length - nuevos, Object.keys(g.cortesia).length,
      g.items, Math.round(g.sesiones), Math.round(g.neto)]);
  });
  return {
    clientes: filas,
    paquetes: aFilas(paquetes, ['mes', 'local', 'tipo', 'items', 'sesiones', 'neto']),
    items: aFilas(items, ['mes', 'local', 'item', 'items', 'sesiones', 'neto'])
  };
}

// ---------- Agenda: "Minutos x mes x sucursal" ----------

/** Por mes y local: turnos y minutos de consultorio según el estado (finalizado, cancelado, ausente…). */
function procesarAgenda(hojas) {
  var t = tablaCon(hojas, { local: /^unidade$/i, tipo: /^localidade$/i, estado: /^status$/i, fecha: /^data$/i,
    minutos: /^minutos de consultorio$/i });
  return agregarAgenda(t);
}

function agregarAgenda(t) {
  var mapa = {}, c = t.col;
  for (var i = 0; i < t.n; i++) {
    var mes = mesDe(c.fecha[i]), local = nombreLocal(c.local[i]);
    if (!mes || !local) continue;
    var estado = String(c.estado[i] || '').trim().toLowerCase() || 'sin estado';
    var min = num(c.minutos[i]);
    var campos = { turnos: 1, minutos: min };
    campos['t:' + estado] = 1;
    campos['m:' + estado] = min;
    acumular(mapa, mes + '|' + local, campos);
  }
  var estados = {};
  Object.keys(mapa).forEach(function (k) { Object.keys(mapa[k]).forEach(function (c) { if (/^[tm]:/.test(c)) estados[c] = true; }); });
  return aFilas(mapa, ['mes', 'local', 'turnos', 'minutos'].concat(Object.keys(estados).sort()));
}

if (typeof module !== 'undefined') {
  module.exports = { agregarCobros: agregarCobros, agregarSesiones: agregarSesiones, agregarAgenda: agregarAgenda,
    agregarClientes: agregarClientes, tipoPaquete: tipoPaquete,
    indices: indices, num: num, mesDe: mesDe };
}
