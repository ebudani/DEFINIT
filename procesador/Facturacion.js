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
