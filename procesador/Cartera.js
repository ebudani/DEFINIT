/**
 * Cruce de leads con la CARTERA DE CLIENTES (ventas con mail y teléfono), en la carpeta "Leads Euge".
 * Cada lead se busca en la cartera por teléfono y mail (y, si no aparece, por nombre). Así se sabe si
 * compró dentro de los 90 días, cuánto tardó y si ya era cliente, también en franquicias.
 *
 * Los archivos solo se leen. Nombres, mails y teléfonos se usan en memoria para cruzar; a la planilla
 * privada de resultados llegan solo cantidades (pestaña "vista", la que usa la sección Leads).
 */

var CARPETA_LEADS_ID = '1O_UcRDzs2fPujZzlLXVx2xxPdDs0te75';
var PATRON_CARTERA = /^cartera de clientes.*\.xlsx$/i;
var ESTADOS_VENTA = { 'Finalizado': true, 'Pago': true, 'Parcialmente Cancelado': true };

function analizarLeadsCartera() {
  var inicio = Date.now(), libro = libroAnalisis(), temporales = [];
  try {
    var archivo = masRecienteEn(CARPETA_LEADS_ID, PATRON_CARTERA);
    if (!archivo) throw new Error('No hay archivo "CARTERA DE CLIENTES…" en la carpeta de leads.');
    var tc = convertir(archivo); temporales.push(tc);
    var c = tablaCon(SpreadsheetApp.openById(tc).getSheets(), { nombre: /^nombre$/i, email: /^e-?mail$/i,
      tel: /^(telefone|tel[eé]fono)$/i, fecha: /^fecha de venta$/i, suc: /^sucursal$/i, valor: /^valor$/i, estado: /^status$/i });
    var n = c.n;
    // Ventas en arreglos (más livianos que un objeto por fila).
    var fecha = new Float64Array(n), valor = new Float64Array(n), esVenta = new Uint8Array(n), local = new Array(n);
    var porTel = {}, porMail = {}, porNombre = {}, maxVenta = 0;
    var agregar = function (mapa, k, i) { if (k) (mapa[k] = mapa[k] || []).push(i); };
    for (var i = 0; i < n; i++) {
      var f = fechaHoraMs(c.col.fecha[i]);
      fecha[i] = f == null ? NaN : f;
      valor[i] = num(c.col.valor[i]);
      esVenta[i] = ESTADOS_VENTA[String(c.col.estado[i]).trim()] ? 1 : 0;
      local[i] = claveLead(nombreLocal(c.col.suc[i]));
      if (esVenta[i] && f > maxVenta) maxVenta = f;
      agregar(porTel, telefono10(c.col.tel[i]), i);
      agregar(porMail, mailLimpio(c.col.email[i]), i);
      var tk = tokensNombre(c.col.nombre[i]);
      if (tk.length >= 2) agregar(porNombre, claveCompleta(tk), i);
    }
    c = null;
    var segCartera = Math.round((Date.now() - inicio) / 1000);

    // Leads (sin repetidos dentro de las 24 h).
    var tl = convertir(DriveApp.getFileById(LEADS_BASE_ID)); temporales.push(tl);
    var t = tablaCon(SpreadsheetApp.openById(tl).getSheets(), { fecha: /^fecha y hora/i, origen: /^source$/i,
      calificacion: /^rating$/i, nombre: /^name$/i, apellido: /^surname$/i, email: /^email$/i,
      tel: /^phone$/i, cel: /^mobile_phone$/i, local: /^store$/i });
    var filas = [];
    for (i = 0; i < t.n; i++) {
      var ms = fechaHoraMs(t.col.fecha[i]);
      if (ms == null) continue;
      var tkL = tokensNombre(t.col.nombre[i]).concat(tokensNombre(t.col.apellido[i]));
      var tel = telefono10(t.col.cel[i]) || telefono10(t.col.tel[i]), mail = mailLimpio(t.col.email[i]);
      var persona = tel || mail || claveCompleta(tkL);
      if (!persona) continue;
      filas.push({ ms: ms, tel: tel, mail: mail, nom: tkL.length >= 2 ? claveCompleta(tkL) : '', persona: persona,
        local: claveLead(t.col.local[i]), origen: origenLimpio(reparar(String(t.col.origen[i] || '-')).trim() || '-'),
        calif: reparar(String(t.col.calificacion[i] || '-')).trim() || '-' });
    }
    t = null;
    filas.sort(function (a, b) { return a.ms - b.ms; });

    var DIA = 864e5, PLAZO = DIAS_CONVERSION * DIA, corte = maxVenta - PLAZO;
    var ultimo = {}, vistos = {}, leads = [];
    filas.forEach(function (l) {
      var p = ultimo[l.persona];
      if (p != null && l.ms - p < DIA) return;
      ultimo[l.persona] = l.ms;
      l.yaEraLead = !!vistos[l.persona]; vistos[l.persona] = true;
      var idx = {};
      (porTel[l.tel] || []).concat(porMail[l.mail] || []).forEach(function (k) { idx[k] = true; });
      var ks = Object.keys(idx);
      if (!ks.length && l.nom) ks = (porNombre[l.nom] || []).map(String);
      var d0 = Math.floor(l.ms / DIA) * DIA, primera = null, previa = false, val = 0;
      ks.forEach(function (s) {
        var k = +s;
        if (!esVenta[k] || isNaN(fecha[k])) return;
        if (fecha[k] < d0) { previa = true; return; }
        if (fecha[k] <= d0 + PLAZO) { val += valor[k]; if (primera == null || fecha[k] < fecha[primera]) primera = k; }
      });
      l.yaEraCliente = previa;
      l.dias = primera == null ? null : Math.floor((fecha[primera] - d0) / DIA);
      l.valor = val;
      l.mismo = primera != null && local[primera] === l.local;
      l.mes = Utilities.formatDate(new Date(l.ms), 'UTC', 'yyyy-MM');
      l.completo = d0 <= corte;
      leads.push(l);
    });

    escribirVistaCartera(libro, leads, maxVenta, corte);
    escribir(libro, 'info', [['dato', 'valor'],
      ['archivos', archivo.getName() + ' + ' + DriveApp.getFileById(LEADS_BASE_ID).getName()],
      ['calculado', fecha_(new Date())],
      ['método', 'Cruce de cada lead con la cartera por teléfono y mail (y por nombre si no aparece). Compró = venta finalizada / paga dentro de los ' + DIAS_CONVERSION + ' días.'],
      ['ventas de la cartera hasta', Utilities.formatDate(new Date(maxVenta), 'UTC', 'yyyy-MM-dd')],
      ['leads (sin repetidos 24 h)', leads.length],
      ['segundos: cartera / total', segCartera + ' / ' + Math.round((Date.now() - inicio) / 1000)],
      ['privacidad', 'Solo cantidades: ningún nombre, mail ni teléfono.']]);
  } finally {
    temporales.forEach(function (id) { try { Drive.Files.remove(id); } catch (e) { /* ya no está */ } });
  }
}

function fecha_(d) { return Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm'); }
function mailLimpio(v) { var m = String(v || '').trim().toLowerCase(); return /@.+\./.test(m) ? m : ''; }

function masRecienteEn(carpetaId, patron) {
  var res = [], it = DriveApp.getFolderById(carpetaId).getFiles();
  while (it.hasNext()) { var f = it.next(); if (patron.test(f.getName())) res.push(f); }
  return res.sort(function (a, b) { return b.getLastUpdated() - a.getLastUpdated(); })[0] || null;
}

/** Misma pestaña "vista" (bloques A, B y C) que usa la sección Leads, con el cruce de la cartera. */
function escribirVistaCartera(libro, leads, maxVenta, corte) {
  var cuenta = {};
  leads.forEach(function (l) { cuenta[l.origen] = (cuenta[l.origen] || 0) + 1; });
  var top = {};
  Object.keys(cuenta).sort(function (a, b) { return cuenta[b] - cuenta[a]; }).slice(0, 25).forEach(function (o) { top[o] = true; });
  var A = {}, B = {}, C = {};
  var sumar = function (mapa, k, l) {
    var x = mapa[k] = mapa[k] || [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    x[0]++;
    if (l.dias != null) {
      x[1]++; x[2] += l.valor;
      var d = l.dias, i = d === 0 ? 3 : d <= 3 ? 4 : d <= 7 ? 5 : d <= 15 ? 6 : d <= 30 ? 7 : d <= 60 ? 8 : 9;
      x[i]++; x[10] += d; if (l.mismo) x[11]++;
    } else if (l.completo) x[12]++;
    else x[13]++;
  };
  leads.forEach(function (l) {
    var tipo = l.yaEraCliente ? 'ya_cliente' : l.yaEraLead ? 'ya_lead' : 'nuevo';
    sumar(A, ['A', l.mes, l.local, tipo].join('|'), l);
    sumar(B, ['B', l.mes, l.local, l.calif].join('|'), l);
    sumar(C, ['C', l.mes, l.local, top[l.origen] ? l.origen : 'otros'].join('|'), l);
  });
  var filas = [['bloque', 'mes', 'local', 'clave', 'leads', 'compraron', 'facturacion', 'd0', 'd1_3', 'd4_7', 'd8_15', 'd16_30', 'd31_60', 'd61_90', 'dias_suma', 'mismo_local', 'no_concretaron', 'en_curso']];
  [A, B, C].forEach(function (mapa) {
    Object.keys(mapa).sort().forEach(function (k) {
      var x = mapa[k];
      filas.push(k.split('|').concat([x[0], x[1], Math.round(x[2])], x.slice(3, 10), [x[10], x[11], x[12], x[13]]));
    });
  });
  filas.push(['info', 'ventas 2018-01-01 a ' + Utilities.formatDate(new Date(maxVenta), 'UTC', 'yyyy-MM-dd'), 'plazo ' + DIAS_CONVERSION + ' días',
    'no concretó hasta ' + Utilities.formatDate(new Date(corte), 'UTC', 'yyyy-MM-dd'), '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  escribir(libro, 'vista', filas);
  libro.setActiveSheet(libro.getSheetByName('vista'));
  libro.moveActiveSheet(1);
}
