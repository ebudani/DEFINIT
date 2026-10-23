/**
 * Facturación neta de los leads: después del cruce con la cartera (analizarLeadsCartera), suma la venta
 * neta real de Venta_Items_Sesiones (carpeta de Martín) de lo que cada lead compró dentro de los 90 días.
 *
 * Cada lead se busca en las ventas primero por nombre y apellido completos y, si no aparece, por nombre y
 * último apellido. Cuando el nombre se repite entre clientes distintos (homónimos), se valida con el mail o el
 * teléfono de la pestaña "base Clientes"; si tampoco hay nombre, se busca solo por mail / teléfono.
 * Venta_Items_Sesiones tiene solo locales propios: en los locales que no figuran ahí (franquicias) se deja
 * el valor del contrato de la cartera, que es lo que ya trae la vista. Solo se escriben cantidades y montos.
 */
function agregarFacturacionLeads() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'agregarFacturacionLeads') ScriptApp.deleteTrigger(t); });
  var inicio = Date.now(), libro = libroAnalisis(), temporales = [];
  try {
    var hojaClaves = libro.getSheetByName('leads_claves');
    if (!hojaClaves) throw new Error('Falta la pestaña leads_claves: primero tiene que correr analizarLeadsCartera.');
    var claves = hojaClaves.getDataRange().getValues().slice(1);
    var V = leerVentasItems(temporales);
    var segVentas = Math.round((Date.now() - inicio) / 1000);

    var leads = leerLeadsBase(temporales);
    if (leads.length !== claves.length) throw new Error('La base de leads cambió desde el último cruce: volver a correr analizarLeadsCartera.');
    var fact = {}, pagaron = {}, conMonto = 0, niveles = {};
    leads.forEach(function (l, k) {
      var cr = comprasDeLead(l, V);
      niveles[cr.nivel + 1] = (niveles[cr.nivel + 1] || 0) + 1;
      var r = medirCompras(cr.listas, l);
      if (!r.neto && !r.pago) return;
      if (r.pago) conMonto++;
      claves[k].slice(1).forEach(function (clave) {
        fact[clave] = (fact[clave] || 0) + r.neto;
        if (r.pago) pagaron[clave] = (pagaron[clave] || 0) + 1;
      });
    });

    // Reemplaza "facturacion" por la venta neta y agrega "pagaron" (compradores con venta neta > 0).
    var vista = libro.getSheetByName('vista').getDataRange().getValues();
    var col = vista[0].indexOf('facturacion'), colC = vista[0].indexOf('compraron'), colP = vista[0].indexOf('pagaron');
    if (colP < 0) { colP = vista[0].length; vista.forEach(function (f, r) { f[colP] = r === 0 ? 'pagaron' : ''; }); }
    vista.forEach(function (f, r) {
      if (r === 0 || f[0] === 'info') return;
      var clave = f.slice(0, 4).join('|');
      // Locales con ventas en el archivo (propios): venta neta. El resto (franquicias): queda el valor del contrato.
      if (V.locales[f[2]]) { f[col] = Math.round(fact[clave] || 0); f[colP] = pagaron[clave] || 0; }
      else f[colP] = f[colC];
    });
    escribir(libro, 'vista', vista.map(function (f) { return f.slice(); }));
    libro.setActiveSheet(libro.getSheetByName('vista')); libro.moveActiveSheet(1);
    escribir(base(), 'leads_vista', vista);
    var info = libro.getSheetByName('info').getDataRange().getValues().filter(function (f) { return !/^facturación/.test(f[0]); });
    info.push(['facturación', 'Venta neta de ' + V.archivo + ' (' + V.primera + ' a ' + V.ultima + ', compras dentro de los ' + DIAS_CONVERSION +
      ' días). Leads que pagaron: ' + conMonto + ' · leads por nivel de cruce: ' + Object.keys(niveles).sort().map(function (n) { return n + ') ' + niveles[n]; }).join(', ') +
      ' · segundos ventas / total: ' + segVentas + ' / ' + Math.round((Date.now() - inicio) / 1000)]);
    escribir(libro, 'info', info);
  } finally {
    temporales.forEach(function (id) { try { Drive.Files.remove(id); } catch (e) { /* ya no está */ } });
  }
}

/** Ventas de Venta_Items_Sesiones indexadas por nombre completo y parcial, más los contactos de "base Clientes". */
function leerVentasItems(temporales) {
  var archivo = masReciente(listarCarpeta(), /^venta_items_sesiones.*\.xlsx$/i);
  var tv = convertir(archivo); temporales.push(tv);
  var hojas = SpreadsheetApp.openById(tv).getSheets();
  var v = tablaCon(hojas, { fecha: /^fecha de pago$/i, local: /^establecimiento$/i, cliente: /^cliente$/i, cortesia: /cortes[íi]a/i, neto: /^v\.?\s*neto$/i });
  var V = { archivo: archivo.getName(), porCompleto: {}, porParcial: {}, contCompleto: {}, contParcial: {}, porTel: {}, porMail: {},
    locales: {}, primera: null, ultima: null, dirFilas: 0, dirTel: 0, dirMail: 0, telTodos: {}, mailTodos: {}, formas: {} };
  for (var i = 0; i < v.n; i++) {
    V.locales[claveLead(nombreLocal(v.col.local[i]))] = true;
    var dia = diaDe(v.col.fecha[i]);
    if (!dia) continue;
    var tk = tokensNombre(v.col.cliente[i]);
    if (tk.length < 2) continue;
    var cort = v.col.cortesia[i] === true || /^(true|verdadero|s[ií]|1)$/i.test(String(v.col.cortesia[i]).trim());
    var c = { ms: Date.UTC(+dia.slice(0, 4), +dia.slice(5, 7) - 1, +dia.slice(8, 10)), neto: cort ? 0 : num(v.col.neto[i]), cort: cort };
    var kc = claveCompleta(tk), kp = tk[0] + '|' + tk[tk.length - 1];
    (V.porCompleto[kc] = V.porCompleto[kc] || []).push(c);
    (V.porParcial[kp] = V.porParcial[kp] || []).push(c);
    if (!V.primera || dia < V.primera) V.primera = dia;
    if (!V.ultima || dia > V.ultima) V.ultima = dia;
  }
  v = null;
  // Contactos de cada nombre (completo y parcial) y nombre de cada contacto, según "base Clientes".
  var d = tablaCon(hojas, { cliente: /^cliente$/i, mail: /^mail$/i, tel: /^tel[eé]fono$/i });
  function anotar(mapa, k, t, m) {
    var x = mapa[k] = mapa[k] || { tel: {}, mail: {} };
    if (t) x.tel[t] = true;
    if (m) x.mail[m] = true;
  }
  // Forma del dato sin el dato (9 = dígito, a = letras), para ver si las columnas vienen bien.
  function forma(x) {
    return String(x == null ? '' : x).slice(0, 40).replace(/[0-9]/g, '9').replace(/[A-Za-zÀ-ÿ]/g, 'a').replace(/a+/g, 'a').replace(/9{4,}/g, '9999') + ' (' + typeof x + ')';
  }
  V.dirFilas = d.n;
  for (i = 0; i < d.n; i++) {
    var tkc = tokensNombre(d.col.cliente[i]);
    if (tkc.length < 2) continue;
    var k1 = claveCompleta(tkc), k2 = tkc[0] + '|' + tkc[tkc.length - 1];
    var t = telefono10(d.col.tel[i]), m = mailLimpio(d.col.mail[i]);
    if (t) { V.dirTel++; V.telTodos[t] = true; }
    if (m) { V.dirMail++; V.mailTodos[m] = true; }
    if (i % 5000 === 0) { var fz = 'tel ' + forma(d.col.tel[i]) + ' · mail ' + forma(d.col.mail[i]); V.formas[fz] = (V.formas[fz] || 0) + 1; }
    if (V.porCompleto[k1]) {
      anotar(V.contCompleto, k1, t, m);
      if (t) (V.porTel[t] = V.porTel[t] || {})[k1] = true;
      if (m) (V.porMail[m] = V.porMail[m] || {})[k1] = true;
    }
    if (V.porParcial[k2]) anotar(V.contParcial, k2, t, m);
  }
  return V;
}

var NIVELES_CRUCE = ['1. nombre completo + mail/tel coinciden', '2. nombre completo, sin homónimos', '3. nombre completo repetido y mail/tel no coinciden (descartado)',
  '4. nombre + último apellido + mail/tel coinciden', '5. nombre + último apellido, sin homónimos', '6. solo mail/tel (nombre escrito distinto)', '7. sin cruce'];

/** Compras de un lead: { nivel: índice en NIVELES_CRUCE, listas: listas de compras de V }. */
function comprasDeLead(l, V) {
  function validado(cont) { return !!cont && !!((l.tel && cont.tel[l.tel]) || (l.mail && cont.mail[l.mail])); }
  function homonimo(cont) {
    if (!cont) return false;
    var nt = Object.keys(cont.tel).length, nm = Object.keys(cont.mail).length;
    return (nt >= 2 && nm !== 1) || (nm >= 2 && nt !== 1);
  }
  var completo = !!(l.nom && V.porCompleto[l.nom]);
  if (completo) {
    var cc = V.contCompleto[l.nom];
    if (validado(cc)) return { nivel: 0, listas: [V.porCompleto[l.nom]] };
    if (!homonimo(cc)) return { nivel: 1, listas: [V.porCompleto[l.nom]] };
  } else if (l.parcial && V.porParcial[l.parcial]) {
    var cp = V.contParcial[l.parcial];
    if (validado(cp)) return { nivel: 3, listas: [V.porParcial[l.parcial]] };
    if (!homonimo(cp)) return { nivel: 4, listas: [V.porParcial[l.parcial]] };
  }
  var porContacto = Object.keys(Object.assign({}, V.porTel[l.tel] || {}, V.porMail[l.mail] || {}));
  if (porContacto.length) return { nivel: 5, listas: porContacto.map(function (n) { return V.porCompleto[n]; }) };
  return { nivel: completo ? 2 : 6, listas: [] };
}

/** Venta neta de las compras dentro de los 90 días desde el día del lead; pago = alguna compra con venta neta > 0. */
function medirCompras(listas, l) {
  var DIA = 864e5, d0 = Math.floor(l.ms / DIA) * DIA, fin = d0 + DIAS_CONVERSION * DIA, neto = 0, pago = false, cortesia = false;
  listas.forEach(function (lista) {
    lista.forEach(function (c) {
      if (c.ms < d0 || c.ms > fin) return;
      if (c.cort) cortesia = true; else { neto += c.neto; if (c.neto > 0) pago = true; }
    });
  });
  return { neto: neto, pago: pago, soloCortesia: cortesia && !pago };
}

/**
 * Diagnóstico (no cambia la vista): leads de locales propios de ene–sep 2026 por nivel de cruce con las ventas,
 * y calidad de los mails / teléfonos de "base Clientes". Pestaña "comparacion_cruce" de la planilla privada.
 */
function compararCruceVentas() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'compararCruceVentas') ScriptApp.deleteTrigger(t); });
  var inicio = Date.now(), libro = libroAnalisis(), temporales = [];
  try {
    var V = leerVentasItems(temporales);
    var res = NIVELES_CRUCE.map(function (n) { return { nivel: n, leads: 0, pago: 0, soloCortesia: 0, neto: 0 }; });
    var desde = Date.UTC(2026, 0, 1), hasta = Date.UTC(2026, 9, 1), total = 0, conContacto = 0, telEnDir = 0, mailEnDir = 0;
    leerLeadsBase(temporales).forEach(function (l) {
      if (l.ms < desde || l.ms >= hasta || !V.locales[l.local]) return;
      total++;
      if (l.tel || l.mail) conContacto++;
      if (l.tel && V.telTodos[l.tel]) telEnDir++;
      if (l.mail && V.mailTodos[l.mail]) mailEnDir++;
      var cr = comprasDeLead(l, V), r = medirCompras(cr.listas, l), x = res[cr.nivel];
      x.leads++; x.neto += r.neto; if (r.pago) x.pago++; if (r.soloCortesia) x.soloCortesia++;
    });
    var filas = [['nivel de cruce', 'leads', 'compraron pagando', 'solo cortesía', 'venta neta', 'venta neta por comprador que pagó']];
    var tot = { pago: 0, soloCortesia: 0, neto: 0 };
    res.forEach(function (r) {
      filas.push([r.nivel, r.leads, r.pago, r.soloCortesia, Math.round(r.neto), r.pago ? Math.round(r.neto / r.pago) : '']);
      tot.pago += r.pago; tot.soloCortesia += r.soloCortesia; tot.neto += r.neto;
    });
    filas.push(['TOTAL', total, tot.pago, tot.soloCortesia, Math.round(tot.neto), tot.pago ? Math.round(tot.neto / tot.pago) : '']);
    filas.push(['', '', '', '', '', '']);
    filas.push(['de esos leads, con mail o teléfono', conContacto, '', '', '', '']);
    filas.push(['ventas del archivo', V.primera + ' a ' + V.ultima, '', '', '', '']);
    filas.push(['base Clientes: filas / con teléfono válido / con mail válido', V.dirFilas + ' / ' + V.dirTel + ' / ' + V.dirMail, '', '', '', '']);
    filas.push(['leads cuyo teléfono / mail aparece en base Clientes', telEnDir + ' / ' + mailEnDir, '', '', '', '']);
    Object.keys(V.formas).sort(function (a, b) { return V.formas[b] - V.formas[a]; }).slice(0, 8).forEach(function (fz) {
      filas.push(['forma de los datos (9 = dígito, a = letras)', fz + ' × ' + V.formas[fz], '', '', '', '']);
    });
    filas.push(['calculado', fecha_(new Date()) + ' · ' + Math.round((Date.now() - inicio) / 1000) + ' s', '', '', '', '']);
    escribir(libro, 'comparacion_cruce', filas);
  } finally {
    temporales.forEach(function (id) { try { Drive.Files.remove(id); } catch (e) { /* ya no está */ } });
  }
}

/** Lo llama la ejecución de cada hora: cuando cambia la versión del cálculo, recalcula la facturación una vez (aparte). */
var VERSION_FACTURACION = 'v3';
function programarFacturacion() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('facturacionVersion') === VERSION_FACTURACION) return;
  ScriptApp.newTrigger('agregarFacturacionLeads').timeBased().after(60 * 1000).create();
  props.setProperty('facturacionVersion', VERSION_FACTURACION);
}
