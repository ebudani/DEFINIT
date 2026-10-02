/**
 * Indicadores comerciales: KPIs por local y mes a partir del último Excel
 * "Descuentos por locales" que haya en la carpeta compartida.
 *
 * El Excel se lee tal cual (sin convertirlo ni copiarlo): se descomprime y se
 * recorre la hoja de detalle (una fila por cobro). El resultado, que es chico,
 * queda en caché 6 h por archivo; cada persona igual tiene que poder ver la
 * carpeta en Drive para recibirlo.
 *
 * Las funciones puras (leerHojaXlsx, agregarCobros) también corren en Node
 * para probarlas con los archivos reales.
 */

// Carpeta de Drive con los Excel comerciales.
var CARPETA_COMERCIAL_ID = '1rqvtYLmOmLqtuvovHk4jIZJIZVhg3TpB';
var PATRON_DESCUENTOS = /^descuentos por locales.*\.xlsx$/i;
var PATRON_EERR = /^reporte eerr.*\.xlsx$/i;

function getComercial() {
  // Google permite autorizar los permisos por separado: si quien abre el tablero no
  // dio el de Drive, se devuelve el link para darlo en vez de fallar.
  try {
    var info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL, ['https://www.googleapis.com/auth/drive.readonly']);
    if (info.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
      return JSON.stringify({ necesitaAutorizar: true, url: info.getAuthorizationUrl() });
    }
  } catch (e) { /* si no se puede consultar, se sigue y el error real aparece más abajo */ }

  var carpeta;
  try {
    carpeta = DriveApp.getFolderById(CARPETA_COMERCIAL_ID);
    carpeta.getName();
  } catch (e) {
    var msg = String(e && e.message || e);
    // Falta autorizar el permiso de Drive: no es falta de acceso a la carpeta.
    if (/permission to call|required permissions|autoriz|authoriz/i.test(msg)) {
      throw new Error('Falta autorizar el permiso para ver archivos de Drive. Abrí el tablero en una ventana de incógnito y aceptá los permisos. (' + msg + ')');
    }
    return JSON.stringify({ sinAcceso: true, detalle: msg });
  }

  var archivos = [];
  var it = carpeta.getFiles();
  while (it.hasNext()) archivos.push(it.next());
  var masReciente = function (patron) {
    return archivos.filter(function (f) { return patron.test(f.getName()); })
      .sort(function (a, b) { return b.getLastUpdated() - a.getLastUpdated(); })[0] || null;
  };
  var archivo = masReciente(PATRON_DESCUENTOS);
  if (!archivo) throw new Error('No encontré ningún Excel "Descuentos por locales" en la carpeta.');
  var eerr = masReciente(PATRON_EERR);

  var firma = function (f) { return f ? f.getId() + ':' + f.getLastUpdated().getTime() : '-'; };
  var clave = 'comercial3:' + firma(archivo) + ':' + firma(eerr);
  var cache = CacheService.getScriptCache();
  var guardado = cache.get(clave);
  if (guardado) return guardado;

  var libro = descomprimir(archivo);
  var compartidos = leerSharedStrings(libro.texto('xl/sharedStrings.xml'));
  var filas = null;
  libro.hojas.slice().sort(function (a, b) { return b.tamanio - a.tamanio; }).some(function (h) {
    var f = leerHojaXlsx(libro.texto(h.ruta), compartidos);
    if (f.length && encontrarColumnas(f[0])) { filas = f; return true; }
    return false;
  });
  if (!filas) throw new Error('El Excel "' + archivo.getName() + '" no tiene la hoja de detalle esperada (Fecha de Pago, Establecimiento, V. Bruto…).');

  var resultado = agregarCobros(filas);
  resultado.archivo = archivo.getName();
  resultado.actualizadoArchivo = Utilities.formatDate(archivo.getLastUpdated(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');

  // Ventas del Reporte EERR (sin IVA): con ellas se estiman los tickets como en el panel de Martín
  // (ventas ÷ ticket promedio). Si falta el archivo o no se puede leer, quedan los cobros contados.
  if (eerr) {
    try {
      var libroEerr = descomprimir(eerr);
      var textos = leerSharedStrings(libroEerr.texto('xl/sharedStrings.xml'));
      var hojas = {};
      libroEerr.hojas.forEach(function (h) { hojas[h.nombre] = leerHojaXlsx(libroEerr.texto(h.ruta), textos); });
      resultado.eerr = armarEerr(hojas);
      resultado.eerr.archivo = eerr.getName();
    } catch (e) {
      resultado.eerr = { error: String(e && e.message || e) };
    }
  }

  var json = JSON.stringify(resultado);
  try { cache.put(clave, json, 21600); } catch (e) { /* más de 100 KB: se recalcula la próxima vez */ }
  return json;
}

/** Descomprime un .xlsx de Drive. Devuelve { texto(ruta), hojas: [{ nombre, ruta, tamanio }] }. */
function descomprimir(archivo) {
  var partes = {};
  Utilities.unzip(archivo.getBlob().setContentType('application/zip')).forEach(function (b) { partes[b.getName()] = b; });
  var texto = function (ruta) { return partes[ruta] ? partes[ruta].getDataAsString('UTF-8') : ''; };
  var hojas = hojasDelLibro(texto('xl/workbook.xml'), texto('xl/_rels/workbook.xml.rels'));
  hojas.forEach(function (h) { h.tamanio = partes[h.ruta] ? partes[h.ruta].getBytes().length : 0; });
  return { texto: texto, hojas: hojas.filter(function (h) { return partes[h.ruta]; }) };
}

/** workbook.xml + sus relaciones -> [{ nombre de la pestaña, ruta del .xml }]. */
function hojasDelLibro(workbookXml, relsXml) {
  var destino = {}, m;
  var reRel = /<Relationship\s[^>]*>/g;
  while ((m = reRel.exec(relsXml))) {
    var id = (m[0].match(/\sId="([^"]+)"/) || [])[1], target = (m[0].match(/\sTarget="([^"]+)"/) || [])[1];
    if (id && target) destino[id] = 'xl/' + target.replace(/^\/?xl\//, '').replace(/^\//, '');
  }
  var hojas = [], reHoja = /<sheet\s[^>]*>/g;
  while ((m = reHoja.exec(workbookXml))) {
    var nombre = (m[0].match(/\sname="([^"]+)"/) || [])[1], rid = (m[0].match(/\sr:id="([^"]+)"/) || [])[1];
    if (nombre && destino[rid]) hojas.push({ nombre: desescapar(nombre), ruta: destino[rid] });
  }
  return hojas;
}

/**
 * Pestañas del Reporte EERR -> { total: { mes: ventas }, porLocal: { clave: { mes: ventas } } }.
 * "EERR" es el consolidado; las pestañas con nombre de local (Juramento, Caballito, Ecom…) se cruzan
 * con Facturación. En cada una se toma la primera fila "Ventas" debajo de la fila de meses.
 */
function armarEerr(hojas) {
  var res = { total: {}, porLocal: {} };
  Object.keys(hojas).forEach(function (nombre) {
    var ventas = ventasDeHojaEerr(hojas[nombre]);
    if (!ventas) return;
    if (/^eerr$/i.test(nombre.trim())) { res.total = ventas; return; }
    var clave = claveFacturacion(nombre);
    if (clave) res.porLocal[clave] = ventas;
  });
  return res;
}

function ventasDeHojaEerr(filas) {
  var meses = null;
  for (var r = 0; r < Math.min(filas.length, 30); r++) {
    var fila = filas[r] || [];
    if (!meses) {
      var m = {}, n = 0;
      for (var c = 2; c < fila.length; c++) {
        if (typeof fila[c] === 'number' && fila[c] > 40000 && fila[c] < 60000) { m[c] = mesDeFecha(fila[c]); n++; }
      }
      if (n >= 3) meses = m;
      continue;
    }
    if (/^ventas$/i.test(String(fila[1] || '').trim())) {
      var ventas = {}, alguno = false;
      Object.keys(meses).forEach(function (c) {
        if (typeof fila[c] === 'number' && fila[c] !== 0) { ventas[meses[c]] = Math.round(fila[c]); alguno = true; }
      });
      if (alguno) return ventas;
    }
  }
  return null;
}

/** sharedStrings.xml -> array de textos (con <r> enriquecidos unidos). */
function leerSharedStrings(xml) {
  var res = [];
  var re = /<si>([\s\S]*?)<\/si>/g, m;
  while ((m = re.exec(xml))) {
    var t = '', r = /<t[^>]*>([\s\S]*?)<\/t>/g, mt;
    while ((mt = r.exec(m[1]))) t += mt[1];
    res.push(desescapar(t));
  }
  return res;
}

function desescapar(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

/** Hoja .xml -> matriz de valores (números, textos). Respeta las columnas vacías por la referencia A1. */
function leerHojaXlsx(xml, compartidos) {
  var filas = [];
  var reFila = /<row[^>]*>([\s\S]*?)<\/row>/g, mf;
  var reCelda = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  while ((mf = reFila.exec(xml))) {
    var fila = [], mc;
    reCelda.lastIndex = 0;
    while ((mc = reCelda.exec(mf[1]))) {
      var col = 0, letras = mc[1];
      for (var i = 0; i < letras.length; i++) col = col * 26 + (letras.charCodeAt(i) - 64);
      var atributos = mc[2], cuerpo = mc[3] || '';
      var tipo = (atributos.match(/t="([^"]+)"/) || [])[1];
      var v = (cuerpo.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
      var valor = '';
      if (tipo === 's') valor = compartidos[+v] || '';
      else if (tipo === 'inlineStr') valor = desescapar((cuerpo.match(/<t[^>]*>([\s\S]*?)<\/t>/) || [])[1] || '');
      else if (tipo === 'str') valor = desescapar(v || '');
      else if (v != null && v !== '') valor = Number(v);
      fila[col - 1] = valor;
    }
    filas.push(fila);
  }
  return filas;
}

function encontrarColumnas(encabezado) {
  var idx = {};
  var buscar = { fecha: /fecha de pago/i, local: /establecimiento/i, bruto: /v\.?\s*bruto/i, desc: /v\.?\s*descuento/i,
    neto: /v\.?\s*neto/i, forma: /forma de pago/i, condicion: /condici[oó]n de pago/i };
  Object.keys(buscar).forEach(function (k) {
    for (var i = 0; i < encabezado.length; i++) if (buscar[k].test(String(encabezado[i] || ''))) { idx[k] = i; break; }
  });
  return ['fecha', 'local', 'bruto', 'desc', 'neto'].every(function (k) { return idx[k] != null; }) ? idx : null;
}

var CUOTAS_COMERCIAL = ['1 pago', '3 cuotas', '6 cuotas', '12 cuotas', 'Otras'];
var FORMAS_COMERCIAL = ['Crédito', 'Débito', 'Web', 'Transferencia', 'Efectivo', 'Otras'];

function mesDeFecha(v) {
  if (typeof v === 'number') {
    var d = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
    return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2);
  }
  var m = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2);
  m = String(v || '').match(/^(\d{4})-(\d{2})/);
  return m ? m[1] + '-' + m[2] : null;
}

function nombreLocalComercial(s) {
  return String(s || '').replace(/^AR\s*-\s*DEFINIT\s*-\s*/i, '').trim().toLowerCase()
    .replace(/(^|\s)\S/g, function (x) { return x.toUpperCase(); });
}

function cuotaDe(condicion, forma) {
  var m = String(condicion || '').match(/^(\d+)\s+PAGO/i);
  if (m) return { 1: '1 pago', 3: '3 cuotas', 6: '6 cuotas', 12: '12 cuotas' }[+m[1]] || 'Otras';
  return /efectivo|transfer|d[eé]bito/i.test(forma || '') ? '1 pago' : 'Otras';
}

function formaDe(forma) {
  var f = String(forma || '');
  if (f.indexOf(';') >= 0) return 'Otras'; // pagos combinados
  if (/cr[eé]dito/i.test(f)) return 'Crédito';
  if (/d[eé]bito/i.test(f)) return 'Débito';
  if (/web/i.test(f)) return 'Web';
  if (/transfer/i.test(f)) return 'Transferencia';
  if (/efectivo/i.test(f)) return 'Efectivo';
  return 'Otras';
}

/**
 * Filas del detalle de cobros -> { meses, locales, datos: { local: { mes: {...} } } }.
 *  - ops / neto / cuotas / formas: solo cobros con venta neta > 0 (ticket = neto ÷ 1,21 ÷ ops)
 *  - bruto / desc: todos los cobros, igual que la tabla dinámica del archivo (% desc = desc ÷ bruto)
 */
function agregarCobros(filas) {
  var idx = encontrarColumnas(filas[0]);
  var datos = {}, meses = {};
  for (var i = 1; i < filas.length; i++) {
    var f = filas[i];
    var mes = mesDeFecha(f[idx.fecha]);
    var local = nombreLocalComercial(f[idx.local]);
    if (!mes || !local) continue;
    meses[mes] = true;
    var a = (datos[local] = datos[local] || {});
    var x = (a[mes] = a[mes] || { ops: 0, neto: 0, bruto: 0, desc: 0, cuotas: {}, formas: {} });
    var bruto = +f[idx.bruto] || 0, desc = +f[idx.desc] || 0, neto = +f[idx.neto] || 0;
    x.bruto += bruto; x.desc += desc;
    if (neto > 0) {
      var forma = idx.forma != null ? f[idx.forma] : '';
      var c = cuotaDe(idx.condicion != null ? f[idx.condicion] : '', forma), fp = formaDe(forma);
      x.ops++; x.neto += neto;
      x.cuotas[c] = (x.cuotas[c] || 0) + neto;
      x.formas[fp] = (x.formas[fp] || 0) + neto;
    }
  }
  Object.keys(datos).forEach(function (l) {
    Object.keys(datos[l]).forEach(function (m) {
      var x = datos[l][m];
      ['neto', 'bruto', 'desc'].forEach(function (k) { x[k] = Math.round(x[k]); });
      [x.cuotas, x.formas].forEach(function (o) { Object.keys(o).forEach(function (k) { o[k] = Math.round(o[k]); }); });
    });
  });
  var locales = Object.keys(datos).sort(function (a, b) { return a.localeCompare(b, 'es'); });
  var equivalencias = {};
  locales.forEach(function (l) { equivalencias[l] = claveFacturacion(l); });
  return {
    meses: Object.keys(meses).sort(),
    locales: locales, equivalencias: equivalencias,
    cuotas: CUOTAS_COMERCIAL, formas: FORMAS_COMERCIAL, datos: datos
  };
}

// Nombre en el sistema de cobros -> nombre en Facturación 2026 (cuando no es el mismo).
// Caballito = Little Horse. Tom no está en Facturación: queda sin equivalencia.
var EQUIVALENCIAS_LOCALES = {
  'alcorta shopping': 'alcorta', 'coronel diaz': 'coronel', 'dot baires shopping': 'dot',
  'las palmas del pilar': 'palmas', 'lomas de san isidro': 'san isidro', 'portal palermo': 'portal',
  'ramos mejia': 'ramos', 'solar shopping': 'solar', 'ecommerce': 'ecom', 'caballito': 'little horse',
  'tom': null
};

/** Clave del local en Facturación 2026 (misma normalización que claveLocal de Parser.js), o null. */
function claveFacturacion(nombre) {
  var k = String(nombre || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim();
  return k in EQUIVALENCIAS_LOCALES ? EQUIVALENCIAS_LOCALES[k] : k;
}

if (typeof module !== 'undefined') {
  module.exports = { leerSharedStrings: leerSharedStrings, leerHojaXlsx: leerHojaXlsx, agregarCobros: agregarCobros,
    hojasDelLibro: hojasDelLibro, armarEerr: armarEerr };
}
