/**
 * Facturación neta de los leads: después del cruce con la cartera (analizarLeadsCartera), suma la venta
 * neta real de Venta_Items_Sesiones (carpeta de Martín) de lo que cada lead compró dentro de los 90 días.
 *
 * Venta_Items_Sesiones tiene las ventas con el nombre del cliente y, en la pestaña "base Clientes", el mail
 * y el teléfono de cada cliente. Cada lead se busca ahí por teléfono / mail (y si no, por nombre) para llegar
 * a su nombre de cliente, y con ese nombre a sus compras. Solo se escriben cantidades y montos.
 * Venta_Items_Sesiones tiene solo locales propios: en los locales que no figuran ahí (franquicias) se deja
 * el valor del contrato de la cartera, que es lo que ya trae la vista.
 */
function agregarFacturacionLeads() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'agregarFacturacionLeads') ScriptApp.deleteTrigger(t); });
  var inicio = Date.now(), libro = libroAnalisis(), temporales = [];
  try {
    var hojaClaves = libro.getSheetByName('leads_claves');
    if (!hojaClaves) throw new Error('Falta la pestaña leads_claves: primero tiene que correr analizarLeadsCartera.');
    var claves = hojaClaves.getDataRange().getValues().slice(1);

    var archivo = masReciente(listarCarpeta(), /^venta_items_sesiones.*\.xlsx$/i);
    var tv = convertir(archivo); temporales.push(tv);
    var hojas = SpreadsheetApp.openById(tv).getSheets();
    var v = tablaCon(hojas, { fecha: /^fecha de pago$/i, local: /^establecimiento$/i, cliente: /^cliente$/i, cortesia: /cortes[íi]a/i, neto: /^v\.?\s*neto$/i });
    var compras = {}, localesConVentas = {};
    for (var i = 0; i < v.n; i++) {
      localesConVentas[claveLead(nombreLocal(v.col.local[i]))] = true;
      var dia = diaDe(v.col.fecha[i]), neto = num(v.col.neto[i]);
      var cort = v.col.cortesia[i] === true || /^(true|verdadero|s[ií]|1)$/i.test(String(v.col.cortesia[i]).trim());
      if (!dia || cort || !neto) continue;
      var tk = tokensNombre(v.col.cliente[i]);
      if (tk.length < 2) continue;
      var k = claveCompleta(tk);
      (compras[k] = compras[k] || []).push({ ms: Date.UTC(+dia.slice(0, 4), +dia.slice(5, 7) - 1, +dia.slice(8, 10)), neto: neto });
    }
    v = null;
    // Directorio de clientes: teléfono / mail -> nombre de cliente.
    var d = tablaCon(hojas, { cliente: /^cliente$/i, mail: /^mail$/i, tel: /^tel[eé]fono$/i });
    var porTel = {}, porMail = {};
    for (i = 0; i < d.n; i++) {
      var tkc = tokensNombre(d.col.cliente[i]);
      if (tkc.length < 2) continue;
      var kc = claveCompleta(tkc);
      if (!compras[kc]) continue; // solo interesan clientes con compras
      var t = telefono10(d.col.tel[i]), m = mailLimpio(d.col.mail[i]);
      if (t) (porTel[t] = porTel[t] || {})[kc] = true;
      if (m) (porMail[m] = porMail[m] || {})[kc] = true;
    }
    d = null;
    var segVentas = Math.round((Date.now() - inicio) / 1000);

    var leads = leerLeadsBase(temporales);
    if (leads.length !== claves.length) throw new Error('La base de leads cambió desde el último cruce: volver a correr analizarLeadsCartera.');
    var DIA = 864e5, PLAZO = DIAS_CONVERSION * DIA;
    var fact = {}, conMonto = 0;
    leads.forEach(function (l, k) {
      var nombres = {};
      Object.keys(porTel[l.tel] || {}).concat(Object.keys(porMail[l.mail] || {})).forEach(function (n) { nombres[n] = true; });
      if (!Object.keys(nombres).length && l.nom && compras[l.nom]) nombres[l.nom] = true;
      var d0 = Math.floor(l.ms / DIA) * DIA, total = 0;
      Object.keys(nombres).forEach(function (n) {
        compras[n].forEach(function (c) { if (c.ms >= d0 && c.ms <= d0 + PLAZO) total += c.neto; });
      });
      if (!total) return;
      conMonto++;
      claves[k].slice(1).forEach(function (clave) { fact[clave] = (fact[clave] || 0) + total; });
    });

    // Reemplaza la columna "facturacion" de la vista por la venta neta (en la planilla privada y en la Base comercial).
    var vista = libro.getSheetByName('vista').getDataRange().getValues();
    var col = vista[0].indexOf('facturacion');
    vista.forEach(function (f, r) {
      if (r === 0 || f[0] === 'info') return;
      // Locales con ventas en el archivo (propios): venta neta. El resto (franquicias): queda el valor del contrato.
      if (localesConVentas[f[2]]) f[col] = Math.round(fact[f.slice(0, 4).join('|')] || 0);
    });
    escribir(libro, 'vista', vista.map(function (f) { return f.slice(); }));
    libro.setActiveSheet(libro.getSheetByName('vista')); libro.moveActiveSheet(1);
    escribir(base(), 'leads_vista', vista);
    var info = libro.getSheetByName('info').getDataRange().getValues().filter(function (f) { return !/^facturación/.test(f[0]); });
    info.push(['facturación', 'Venta neta de ' + archivo.getName() + ' (compras dentro de los ' + DIAS_CONVERSION + ' días). Leads con monto: ' + conMonto +
      ' · segundos ventas / total: ' + segVentas + ' / ' + Math.round((Date.now() - inicio) / 1000)]);
    escribir(libro, 'info', info);
  } finally {
    temporales.forEach(function (id) { try { Drive.Files.remove(id); } catch (e) { /* ya no está */ } });
  }
}

/**
 * Diagnóstico (no cambia la vista): cruza los leads de locales propios de ene–sep 2026 con las ventas de
 * Venta_Items_Sesiones primero por nombre y apellido, y valida con mail / teléfono de "base Clientes" cuando
 * el nombre se repite (homónimos). Escribe en la planilla privada la pestaña "comparacion_cruce", solo cantidades.
 */
function compararCruceVentas() {
  var inicio = Date.now(), libro = libroAnalisis(), temporales = [];
  try {
    var archivo = masReciente(listarCarpeta(), /^venta_items_sesiones.*\.xlsx$/i);
    var tv = convertir(archivo); temporales.push(tv);
    var hojas = SpreadsheetApp.openById(tv).getSheets();
    var v = tablaCon(hojas, { fecha: /^fecha de pago$/i, local: /^establecimiento$/i, cliente: /^cliente$/i, cortesia: /cortes[íi]a/i, neto: /^v\.?\s*neto$/i });
    var porCompleto = {}, porParcial = {}, locales = {}, primera = null, ultima = null;
    for (var i = 0; i < v.n; i++) {
      locales[claveLead(nombreLocal(v.col.local[i]))] = true;
      var dia = diaDe(v.col.fecha[i]);
      if (!dia) continue;
      var tk = tokensNombre(v.col.cliente[i]);
      if (tk.length < 2) continue;
      var cort = v.col.cortesia[i] === true || /^(true|verdadero|s[ií]|1)$/i.test(String(v.col.cortesia[i]).trim());
      var c = { ms: Date.UTC(+dia.slice(0, 4), +dia.slice(5, 7) - 1, +dia.slice(8, 10)), neto: cort ? 0 : num(v.col.neto[i]), cort: cort };
      (porCompleto[claveCompleta(tk)] = porCompleto[claveCompleta(tk)] || []).push(c);
      var kp = tk[0] + '|' + tk[tk.length - 1];
      (porParcial[kp] = porParcial[kp] || []).push(c);
      if (!primera || dia < primera) primera = dia;
      if (!ultima || dia > ultima) ultima = dia;
    }
    v = null;
    // Contactos de cada nombre (completo y parcial) y nombre de cada contacto, según "base Clientes".
    var d = tablaCon(hojas, { cliente: /^cliente$/i, mail: /^mail$/i, tel: /^tel[eé]fono$/i });
    var contCompleto = {}, contParcial = {}, porTel = {}, porMail = {};
    function anotar(mapa, k, t, m) {
      var x = mapa[k] = mapa[k] || { tel: {}, mail: {} };
      if (t) x.tel[t] = true;
      if (m) x.mail[m] = true;
    }
    for (i = 0; i < d.n; i++) {
      var tkc = tokensNombre(d.col.cliente[i]);
      if (tkc.length < 2) continue;
      var kc = claveCompleta(tkc), kpc = tkc[0] + '|' + tkc[tkc.length - 1];
      var t = telefono10(d.col.tel[i]), m = mailLimpio(d.col.mail[i]);
      if (porCompleto[kc]) anotar(contCompleto, kc, t, m);
      if (porParcial[kpc]) anotar(contParcial, kpc, t, m);
      if (porCompleto[kc]) {
        if (t) (porTel[t] = porTel[t] || {})[kc] = true;
        if (m) (porMail[m] = porMail[m] || {})[kc] = true;
      }
    }
    d = null;

    function validado(cont, l) { return !!cont && ((l.tel && cont.tel[l.tel]) || (l.mail && cont.mail[l.mail])); }
    function homonimo(cont) {
      if (!cont) return false;
      var nt = Object.keys(cont.tel).length, nm = Object.keys(cont.mail).length;
      return (nt >= 2 && nm !== 1) || (nm >= 2 && nt !== 1);
    }
    var NIVELES = ['1. nombre completo + mail/tel coinciden', '2. nombre completo, sin homónimos', '3. nombre completo repetido y mail/tel no coinciden (descartado)',
      '4. nombre + último apellido + mail/tel coinciden', '5. nombre + último apellido, sin homónimos', '6. solo mail/tel (nombre escrito distinto)', '7. sin cruce'];
    var res = NIVELES.map(function (n) { return { nivel: n, leads: 0, pago: 0, soloCortesia: 0, neto: 0 }; });
    var actual = { leads: 0, pago: 0, soloCortesia: 0, neto: 0 };
    var DIA = 864e5, PLAZO = DIAS_CONVERSION * DIA, desde = Date.UTC(2026, 0, 1), hasta = Date.UTC(2026, 9, 1);
    var conContacto = 0, total = 0;

    function medir(listas, l) {
      var d0 = Math.floor(l.ms / DIA) * DIA, neto = 0, pago = false, cortesia = false;
      listas.forEach(function (lista) {
        lista.forEach(function (c) {
          if (c.ms < d0 || c.ms > d0 + PLAZO) return;
          if (c.cort) cortesia = true; else { neto += c.neto; if (c.neto > 0) pago = true; }
        });
      });
      return { neto: neto, pago: pago, soloCortesia: cortesia && !pago };
    }
    function sumar(x, r) { x.leads++; x.neto += r.neto; if (r.pago) x.pago++; if (r.soloCortesia) x.soloCortesia++; }

    leerLeadsBase(temporales).forEach(function (l) {
      if (l.ms < desde || l.ms >= hasta || !locales[l.local]) return;
      total++;
      if (l.tel || l.mail) conContacto++;
      var porContacto = Object.keys(Object.assign({}, porTel[l.tel] || {}, porMail[l.mail] || {}));
      var nivel, listas = [];
      var kp = l.parcial;
      if (l.nom && porCompleto[l.nom]) {
        var cc = contCompleto[l.nom];
        if (validado(cc, l)) { nivel = 0; listas = [porCompleto[l.nom]]; }
        else if (!homonimo(cc)) { nivel = 1; listas = [porCompleto[l.nom]]; }
        else nivel = 2;
      }
      if (nivel == null || nivel === 2) {
        if (kp && porParcial[kp] && nivel == null) {
          var cp = contParcial[kp];
          if (validado(cp, l)) { nivel = 3; listas = [porParcial[kp]]; }
          else if (!homonimo(cp)) { nivel = 4; listas = [porParcial[kp]]; }
        }
        if ((nivel == null || nivel === 2) && porContacto.length) {
          nivel = 5; listas = porContacto.map(function (n) { return porCompleto[n]; });
        }
      }
      if (nivel == null) nivel = 6;
      sumar(res[nivel], medir(listas, l));
      // Método actual (Facturacion): contacto primero, si no nombre completo.
      var act = porContacto.length ? porContacto.map(function (n) { return porCompleto[n]; }) : (l.nom && porCompleto[l.nom] ? [porCompleto[l.nom]] : []);
      sumar(actual, medir(act, l));
    });

    var filas = [['nivel de cruce', 'leads', 'compraron pagando', 'solo cortesía', 'venta neta', 'venta neta por comprador que pagó']];
    var tot = { leads: 0, pago: 0, soloCortesia: 0, neto: 0 };
    res.forEach(function (r, k) {
      filas.push([r.nivel, r.leads, r.pago, r.soloCortesia, Math.round(r.neto), r.pago ? Math.round(r.neto / r.pago) : '']);
      if (k !== 2 && k !== 6) { tot.pago += r.pago; tot.soloCortesia += r.soloCortesia; tot.neto += r.neto; }
      tot.leads += r.leads;
    });
    filas.push(['TOTAL nuevo método (niveles 1, 2, 4, 5 y 6)', tot.leads, tot.pago, tot.soloCortesia, Math.round(tot.neto), tot.pago ? Math.round(tot.neto / tot.pago) : '']);
    filas.push(['método actual (mail/tel primero, si no nombre)', actual.leads, actual.pago, actual.soloCortesia, Math.round(actual.neto), actual.pago ? Math.round(actual.neto / actual.pago) : '']);
    filas.push(['', '', '', '', '', '']);
    filas.push(['leads de locales propios ene–sep 2026', total, '', '', '', '']);
    filas.push(['de esos, con mail o teléfono', conContacto, '', '', '', '']);
    filas.push(['ventas del archivo', primera + ' a ' + ultima, '', '', '', '']);
    filas.push(['calculado', fecha_(new Date()) + ' · ' + Math.round((Date.now() - inicio) / 1000) + ' s', '', '', '', '']);
    escribir(libro, 'comparacion_cruce', filas);
  } finally {
    temporales.forEach(function (id) { try { Drive.Files.remove(id); } catch (e) { /* ya no está */ } });
  }
}
