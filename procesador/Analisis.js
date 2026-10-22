/**
 * Análisis de leads (PRIVADO): cruza la base de leads de Natalia ("Base desde 1 de enero 2025.xlsx",
 * carpeta "Leads Euge") con las ventas (Venta_Items_Sesiones.xlsx) para saber cuántos leads compraron,
 * cuánto facturaron, cuántos ya eran clientes o leads, cuántos no concretaron y cuánto tardan en comprar.
 *
 * Los archivos solo se leen (no se modifican). Nombres, mails y teléfonos se usan únicamente en memoria
 * para cruzar: a la planilla de resultados ("DEFINIT – Análisis de leads (privado)", que no se comparte
 * con nadie) llegan solo cantidades, montos y días. No toca la Base comercial ni el tablero.
 */

var LEADS_BASE_ID = '19XNbeIznMXqr8zYC6PpSMU1OkVBWM94s';
var NOMBRE_ANALISIS = 'DEFINIT – Análisis de leads (privado)';
var DIAS_CONVERSION = 60; // una compra cuenta para el lead si ocurre dentro de estos días

/** La ejecución automática lo programa una sola vez; corre en su propia ejecución (tarda unos minutos). */
function programarAnalisisLeads() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('analisisLeads') || props.getProperty('analisisLeadsProgramado')) return;
  ScriptApp.newTrigger('analizarLeads').timeBased().after(60 * 1000).create();
  props.setProperty('analisisLeadsProgramado', new Date().toISOString());
}

function analizarLeads() {
  var props = PropertiesService.getScriptProperties();
  // Este disparador es de una sola vez: se borra al empezar.
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'analizarLeads') ScriptApp.deleteTrigger(t); });
  var libro = libroAnalisis();
  try {
    var r = calcularAnalisisLeads();
    escribirAnalisis(libro, r);
    props.setProperty('analisisLeads', new Date().toISOString());
  } catch (e) {
    escribir(libro, 'info', [['dato', 'valor'], ['error', String(e && e.stack || e)]]);
    props.deleteProperty('analisisLeadsProgramado');
    throw e;
  }
}

function libroAnalisis() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('analisisId');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* se borró: se crea de nuevo */ } }
  var libro = SpreadsheetApp.create(NOMBRE_ANALISIS);
  props.setProperty('analisisId', libro.getId());
  return libro;
}

// ---------- Claves para cruzar (solo en memoria) ----------

function tokensNombre(s) {
  return reparar(String(s || '')).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase()
    .replace(/[^A-Z]+/g, ' ').trim().split(' ').filter(function (t) { return t.length > 1; });
}
function claveCompleta(tokens) { return tokens.slice().sort().join(' '); }
function telefono10(v) { var d = String(v || '').replace(/\D/g, ''); return d.length >= 8 ? d.slice(-10) : ''; }

// ---------- Cálculo ----------

function calcularAnalisisLeads() {
  var temporales = [];
  try {
    // Ventas: compras con venta neta > 0, por cliente (clave de nombre).
    var ventasArchivo = masReciente(listarCarpeta(), /^venta_items_sesiones.*\.xlsx$/i);
    var tv = convertir(ventasArchivo); temporales.push(tv);
    var hojasV = SpreadsheetApp.openById(tv).getSheets();
    var v = tablaCon(hojasV, { fecha: /^fecha de pago$/i, local: /^establecimiento$/i, cliente: /^cliente$/i,
      cortesia: /cortes[íi]a/i, neto: /^v\.?\s*neto$/i });
    var compras = {}, porPar = {}, minVenta = null, maxVenta = null, nombreCliente = {};
    for (var i = 0; i < v.n; i++) {
      var dia = diaDe(v.col.fecha[i]), neto = num(v.col.neto[i]);
      if (!dia) continue;
      if (!minVenta || dia < minVenta) minVenta = dia;
      if (!maxVenta || dia > maxVenta) maxVenta = dia;
      var cort = v.col.cortesia[i] === true || /^(true|verdadero|s[ií]|1)$/i.test(String(v.col.cortesia[i]).trim());
      if (cort || neto <= 0) continue;
      var tk = tokensNombre(v.col.cliente[i]);
      if (tk.length < 2) continue;
      var k = claveCompleta(tk);
      (compras[k] = compras[k] || []).push({ dia: dia, neto: neto, local: claveLead(nombreLocal(v.col.local[i])) });
      if (!nombreCliente[k]) nombreCliente[k] = String(v.col.cliente[i]).trim();
      // Clave corta (primer y último nombre) para los que en un sistema figuran con todos los nombres y en otro no.
      var par = claveCompleta([tk[0], tk[tk.length - 1]]);
      (porPar[par] = porPar[par] || {})[k] = true;
    }
    Object.keys(compras).forEach(function (k) { compras[k].sort(function (a, b) { return a.dia < b.dia ? -1 : 1; }); });

    // Leads.
    var tl = convertir(DriveApp.getFileById(LEADS_BASE_ID)); temporales.push(tl);
    var t = tablaCon(SpreadsheetApp.openById(tl).getSheets(), { fecha: /^fecha y hora/i, origen: /^source$/i,
      calificacion: /^rating$/i, nombre: /^name$/i, apellido: /^surname$/i, email: /^email$/i,
      tel: /^phone$/i, cel: /^mobile_phone$/i, local: /^store$/i });
    var leads = [], vistos = {}, crudosPorMes = {};
    var cruce = { completo: 0, corto: 0, ambiguo: 0, sin: 0 };
    for (i = 0; i < t.n; i++) {
      var d = diaDe(t.col.fecha[i]);
      if (!d) continue;
      crudosPorMes[d.slice(0, 7)] = (crudosPorMes[d.slice(0, 7)] || 0) + 1;
      var tkL = tokensNombre(t.col.nombre[i]).concat(tokensNombre(t.col.apellido[i]));
      var persona = telefono10(t.col.cel[i]) || telefono10(t.col.tel[i]) ||
        String(t.col.email[i] || '').trim().toLowerCase().replace(/^-$/, '') || claveCompleta(tkL);
      if (!persona) continue;
      if (vistos[persona + '|' + d]) continue; // mismo lead cargado dos veces el mismo día
      vistos[persona + '|' + d] = true;
      // Cliente de ventas que corresponde a este lead.
      var cliente = null, tipo = 'sin';
      if (tkL.length >= 2) {
        var kc = claveCompleta(tkL);
        if (compras[kc]) { cliente = kc; tipo = 'completo'; }
        else {
          var p = porPar[claveCompleta([tkL[0], tkL[tkL.length - 1]])];
          var cands = p ? Object.keys(p) : [];
          if (cands.length === 1) { cliente = cands[0]; tipo = 'corto'; }
          else if (cands.length > 1) tipo = 'ambiguo';
        }
      }
      leads.push({ dia: d, persona: persona, cliente: cliente, tipoCruce: tipo,
        nombreLead: (String(t.col.nombre[i] || '') + ' ' + String(t.col.apellido[i] || '')).trim(),
        local: claveLead(t.col.local[i]), origen: reparar(String(t.col.origen[i] || '-')).trim() || '-',
        calif: reparar(String(t.col.calificacion[i] || '-')).trim() || '-' });
    }
    leads.sort(function (a, b) { return a.dia < b.dia ? -1 : a.dia > b.dia ? 1 : 0; });

    // Leads anteriores de la misma persona.
    var ultimoLead = {};
    leads.forEach(function (l) { l.yaEraLead = !!ultimoLead[l.persona]; ultimoLead[l.persona] = l.dia; });

    // Cada compra se atribuye al lead más reciente de esa persona dentro de los 60 días previos.
    var porCliente = {};
    leads.forEach(function (l) { if (l.cliente) (porCliente[l.cliente] = porCliente[l.cliente] || []).push(l); });
    Object.keys(porCliente).forEach(function (k) {
      var ls = porCliente[k];
      compras[k].forEach(function (c) {
        var elegido = null;
        ls.forEach(function (l) { if (l.dia <= c.dia && diasEntre(l.dia, c.dia) <= DIAS_CONVERSION) elegido = l; });
        if (!elegido) return;
        if (elegido.dias == null) { elegido.dias = diasEntre(elegido.dia, c.dia); elegido.localCompra = c.local; }
        elegido.facturacion = (elegido.facturacion || 0) + c.neto;
      });
      ls.forEach(function (l) { l.yaEraCliente = compras[k].some(function (c) { return c.dia < l.dia; }); });
    });

    return { leads: leads, minVenta: minVenta, maxVenta: maxVenta, cruce: contarCruce(leads), compras: compras,
      nombreCliente: nombreCliente, crudosPorMes: crudosPorMes,
      archivos: [ventasArchivo.getName(), DriveApp.getFileById(LEADS_BASE_ID).getName()] };
  } finally {
    temporales.forEach(function (id) { try { Drive.Files.remove(id); } catch (e) { /* ya no está */ } });
  }
}

function diasEntre(a, b) {
  var ua = Date.UTC(+a.slice(0, 4), +a.slice(5, 7) - 1, +a.slice(8, 10)), ub = Date.UTC(+b.slice(0, 4), +b.slice(5, 7) - 1, +b.slice(8, 10));
  return Math.round((ub - ua) / 864e5);
}
function sumarDiasIso(iso, n) { return Utilities.formatDate(new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) + n * 864e5), 'UTC', 'yyyy-MM-dd'); }
function contarCruce(leads) { var c = {}; leads.forEach(function (l) { c[l.tipoCruce] = (c[l.tipoCruce] || 0) + 1; }); return c; }

// ---------- Resultados ----------

var COLUMNAS_ANALISIS = ['leads', 'personas', 'calif_presupuesto', 'calif_compro', 'sin_calificar', 'compraron', 'pct_compraron',
  'facturacion_neta', 'ticket_por_comprador', 'ya_eran_clientes', 'ya_eran_leads', 'nuevos', 'nuevos_compraron',
  'presupuesto_y_compraron', 'no_concretaron', 'en_curso', 'dias_promedio', 'dias_mediana', 'pct_en_7_dias', 'pct_en_30_dias', 'compraron_mismo_local'];

function agrupar(leads, clave, corte) {
  var g = {};
  leads.forEach(function (l) {
    var k = clave(l);
    var x = g[k] = g[k] || { leads: 0, personas: {}, presu: 0, compro: 0, sin: 0, compraron: 0, fact: 0, yaCli: 0, yaLead: 0,
      nuevos: 0, nuevosComp: 0, presuComp: 0, noConc: 0, enCurso: 0, dias: [], mismo: 0 };
    x.leads++; x.personas[l.persona] = true;
    if (/presupuesto/i.test(l.calif)) x.presu++;
    if (/compr/i.test(l.calif)) x.compro++;
    if (l.calif === '-') x.sin++;
    if (l.yaEraCliente) x.yaCli++;
    if (l.yaEraLead) x.yaLead++;
    var nuevo = !l.yaEraCliente && !l.yaEraLead;
    if (nuevo) x.nuevos++;
    if (l.dias != null) {
      x.compraron++; x.fact += l.facturacion || 0; x.dias.push(l.dias);
      if (nuevo) x.nuevosComp++;
      if (/presupuesto/i.test(l.calif)) x.presuComp++;
      if (l.localCompra === l.local) x.mismo++;
    } else if (l.dia <= corte) x.noConc++;
    else x.enCurso++;
  });
  return Object.keys(g).sort().map(function (k) {
    var x = g[k], dias = x.dias.slice().sort(function (a, b) { return a - b; });
    var prom = dias.length ? dias.reduce(function (s, d) { return s + d; }, 0) / dias.length : '';
    return [k, x.leads, Object.keys(x.personas).length, x.presu, x.compro, x.sin, x.compraron,
      x.leads ? Math.round(x.compraron / x.leads * 1000) / 10 : '', Math.round(x.fact), x.compraron ? Math.round(x.fact / x.compraron) : '',
      x.yaCli, x.yaLead, x.nuevos, x.nuevosComp, x.presuComp, x.noConc, x.enCurso,
      prom === '' ? '' : Math.round(prom * 10) / 10, dias.length ? dias[Math.floor(dias.length / 2)] : '',
      dias.length ? Math.round(dias.filter(function (d) { return d <= 7; }).length / dias.length * 1000) / 10 : '',
      dias.length ? Math.round(dias.filter(function (d) { return d <= 30; }).length / dias.length * 1000) / 10 : '', x.mismo];
  });
}

function escribirAnalisis(libro, r) {
  // Solo leads desde que hay ventas (para saber si ya eran clientes y si compraron).
  var leads = r.leads.filter(function (l) { return l.dia >= r.minVenta && l.dia <= r.maxVenta; });
  var corte = sumarDiasIso(r.maxVenta, -DIAS_CONVERSION); // antes de esta fecha ya se puede decir "no concretó"
  var enc = function (primera) { return [[primera].concat(COLUMNAS_ANALISIS)]; };
  var total = agrupar(leads, function () { return 'TOTAL'; }, corte);
  escribir(libro, 'por_mes', enc('mes').concat(agrupar(leads, function (l) { return l.dia.slice(0, 7); }, corte), total));
  escribir(libro, 'por_local', enc('local').concat(agrupar(leads, function (l) { return l.local; }, corte), total));
  var origenes = agrupar(leads, function (l) { return l.origen; }, corte).sort(function (a, b) { return b[1] - a[1]; });
  escribir(libro, 'por_origen', enc('origen').concat(origenes, total));
  escribir(libro, 'por_calificacion', enc('calificacion').concat(agrupar(leads, function (l) { return l.calif; }, corte), total));
  // Días hasta la primera compra.
  var tramos = [[0, 0, 'el mismo día'], [1, 3, '1 a 3 días'], [4, 7, '4 a 7 días'], [8, 15, '8 a 15 días'], [16, 30, '16 a 30 días'], [31, 60, '31 a 60 días']];
  var conv = leads.filter(function (l) { return l.dias != null; });
  escribir(libro, 'dias_hasta_compra', [['tramo', 'compradores', 'pct']].concat(tramos.map(function (tr) {
    var n = conv.filter(function (l) { return l.dias >= tr[0] && l.dias <= tr[1]; }).length;
    return [tr[2], n, conv.length ? Math.round(n / conv.length * 1000) / 10 : ''];
  })));
  escribirValidacion(libro, r, leads);
  escribir(libro, 'info', [['dato', 'valor'],
    ['archivos', r.archivos.join(' + ')],
    ['calculado', fecha(new Date())],
    ['ventas disponibles', r.minVenta + ' a ' + r.maxVenta],
    ['leads analizados (desde que hay ventas)', leads.length],
    ['leads en la base (todos, sin duplicados del mismo día)', r.leads.length],
    ['cruce por nombre completo', r.cruce.completo || 0],
    ['cruce por primer nombre + apellido', r.cruce.corto || 0],
    ['nombre ambiguo (no se cruzó)', r.cruce.ambiguo || 0],
    ['sin compras con ese nombre', r.cruce.sin || 0],
    ['regla', 'Compró = compra con venta neta > 0 dentro de los ' + DIAS_CONVERSION + ' días del lead. "No concretó" solo para leads hasta el ' + corte + '; los posteriores quedan "en curso".'],
    ['regla', 'Ya era cliente = compró antes del lead (desde ' + r.minVenta + '). Ya era lead = mismo teléfono / mail con un lead anterior desde 2025.'],
    ['privacidad', 'Todas las pestañas tienen solo cantidades, salvo "validacion_muestra", que tiene nombres para revisar el cruce a mano: borrarla después de validar.']]);
  var sobra = libro.getSheetByName('Hoja 1') || libro.getSheetByName('Sheet1');
  if (sobra && libro.getSheets().length > 1) libro.deleteSheet(sobra);
}

/** Pestañas para validar el cruce: totales con y sin limpieza, "Compró" sin compra y una muestra para revisar a mano. */
function escribirValidacion(libro, r, leads) {
  // 1. Leads por mes: filas del archivo vs. sin repetidos (misma persona el mismo día).
  var limpios = {};
  r.leads.forEach(function (l) { limpios[l.dia.slice(0, 7)] = (limpios[l.dia.slice(0, 7)] || 0) + 1; });
  var filas = [['mes', 'filas_en_archivo', 'sin_repetidos_mismo_dia', 'diferencia']];
  Object.keys(r.crudosPorMes).sort().forEach(function (m) {
    filas.push([m, r.crudosPorMes[m], limpios[m] || 0, r.crudosPorMes[m] - (limpios[m] || 0)]);
  });
  escribir(libro, 'validacion_totales', filas);

  // 2. Leads calificados "Compró" que no tienen compra atribuida: por qué.
  var motivos = { 'nombre no encontrado en ventas': 0, 'nombre ambiguo (varios clientes)': 0, 'compró antes del lead': 0,
    'compró después de los 60 días': 0, 'cliente encontrado sin compras con venta > 0 cerca': 0 };
  leads.filter(function (l) { return /compr/i.test(l.calif) && l.dias == null; }).forEach(function (l) {
    if (l.tipoCruce === 'ambiguo') { motivos['nombre ambiguo (varios clientes)']++; return; }
    if (!l.cliente) { motivos['nombre no encontrado en ventas']++; return; }
    var cs = r.compras[l.cliente] || [];
    if (cs.some(function (c) { return c.dia < l.dia; }) && !cs.some(function (c) { return c.dia >= l.dia; })) motivos['compró antes del lead']++;
    else if (cs.some(function (c) { return c.dia >= l.dia && diasEntre(l.dia, c.dia) > DIAS_CONVERSION; })) motivos['compró después de los 60 días']++;
    else motivos['cliente encontrado sin compras con venta > 0 cerca']++;
  });
  escribir(libro, 'validacion_compro', [['motivo', 'leads']].concat(Object.keys(motivos).map(function (k) { return [k, motivos[k]]; })));

  // 3. Muestra para revisar a mano (CON NOMBRES: pestaña privada, borrar después de validar).
  var elegir = function (lista, n) {
    var copia = lista.slice(), out = [];
    while (copia.length && out.length < n) out.push(copia.splice(Math.floor(Math.random() * copia.length), 1)[0]);
    return out;
  };
  var muestra = [['grupo', 'fecha_lead', 'local_lead', 'origen', 'calificacion', 'nombre_en_lead', 'nombre_en_ventas', 'tipo_cruce',
    'dias_hasta_compra', 'facturacion_atribuida', 'local_compra']];
  var agregar = function (grupo, ls) {
    ls.forEach(function (l) {
      muestra.push([grupo, l.dia, l.local, l.origen, l.calif, l.nombreLead, l.cliente ? r.nombreCliente[l.cliente] || '' : '',
        l.tipoCruce, l.dias == null ? '' : l.dias, l.facturacion ? Math.round(l.facturacion) : '', l.localCompra || '']);
    });
  };
  agregar('A. cruce por nombre completo y compró', elegir(leads.filter(function (l) { return l.tipoCruce === 'completo' && l.dias != null; }), 15));
  agregar('B. cruce por primer nombre + apellido y compró', elegir(leads.filter(function (l) { return l.tipoCruce === 'corto' && l.dias != null; }), 15));
  agregar('C. calificado "Compró" sin compra encontrada', elegir(leads.filter(function (l) { return /compr/i.test(l.calif) && l.dias == null; }), 15));
  escribir(libro, 'validacion_muestra', muestra);
}
