/**
 * Conexión con la API de EVUP (https://definit.api.evup.dev) para leer leads.
 *
 * El usuario y la contraseña NO van en el código: se cargan una sola vez en
 * Configuración del proyecto → Propiedades de la secuencia de comandos, con los nombres
 * EVUP_USUARIO y EVUP_CLAVE. Solo se usan endpoints de lectura (Login, Lead/Count, Lead/List).
 *
 * La API devuelve datos personales (nombre, mail, teléfono, documento). Acá solo se cuentan
 * y se leen campos sin datos personales (fecha, estado, local, origen): nunca se escriben en la base.
 */

var EVUP_URL = 'https://definit.api.evup.dev';

function evupToken() {
  var props = PropertiesService.getScriptProperties();
  var usuario = props.getProperty('EVUP_USUARIO'), clave = props.getProperty('EVUP_CLAVE');
  if (!usuario || !clave) throw new Error('Faltan EVUP_USUARIO / EVUP_CLAVE en las propiedades del proyecto.');
  var r = UrlFetchApp.fetch(EVUP_URL + '/Auth/Login', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ Login: usuario, Password: clave })
  });
  if (r.getResponseCode() !== 200) throw new Error('EVUP rechazó el login (HTTP ' + r.getResponseCode() + ').');
  // La respuesta puede ser el token solo (texto) o un JSON que lo contiene.
  var texto = r.getContentText().trim(), datos;
  try { datos = JSON.parse(texto); } catch (e) { datos = texto.replace(/^"|"$/g, ''); }
  var token = buscarToken(datos);
  if (!token) throw new Error('EVUP respondió al login pero no encontré el token en la respuesta.');
  return token;
}

/** El nombre del campo del token no está documentado: se busca uno que se llame "token" o parezca un JWT. */
function buscarToken(o) {
  if (typeof o === 'string') return /^[\w-]+\.[\w-]+\.[\w-]+$/.test(o) ? o : null;
  if (!o || typeof o !== 'object') return null;
  var claves = Object.keys(o).sort(function (a, b) { return (/token/i.test(b) ? 1 : 0) - (/token/i.test(a) ? 1 : 0); });
  for (var i = 0; i < claves.length; i++) { var t = buscarToken(o[claves[i]]); if (t) return t; }
  return null;
}

function evup(token, ruta, cuerpo) {
  var r = UrlFetchApp.fetch(EVUP_URL + ruta, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token }, payload: JSON.stringify(cuerpo || {})
  });
  if (r.getResponseCode() !== 200) throw new Error('EVUP ' + ruta + ': HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
  return JSON.parse(r.getContentText());
}

/**
 * Prueba de conexión: ejecutarla a mano una vez. Escribe en la pestaña "evup_prueba" de la base
 * solo la estructura de la respuesta (nombres de campos) y, de los primeros leads, los campos sin
 * datos personales. Sirve para saber cómo armar el conteo de leads.
 */
function probarEvup() {
  var token = evupToken();
  var filas = [['dato', 'valor']];
  var cuenta = evup(token, '/Lead/Count', { PageNumber: 1, Filters: [], Sort: [] });
  filas.push(['Lead/Count', JSON.stringify(cuenta).slice(0, 300)]);
  var lista = evup(token, '/Lead/List', { PageNumber: 1, Filters: [], Sort: [{ Field: 'CreationDate', Operation: 'desc' }] });
  var items = Array.isArray(lista) ? lista : (lista.items || lista.Items || lista.data || lista.Data || []);
  filas.push(['Lead/List: campos de la respuesta', Array.isArray(lista) ? '(lista)' : Object.keys(lista).join(', ')]);
  if (!Array.isArray(lista)) Object.keys(lista).forEach(function (k) {
    if (typeof lista[k] !== 'object') filas.push(['Lead/List: ' + k, String(lista[k])]);
  });
  filas.push(['Lead/List: leads en la página', items.length]);
  if (items.length) filas.push(['Lead: campos de cada lead', Object.keys(items[0]).join(', ')]);
  // Solo campos sin datos personales.
  items.slice(0, 20).forEach(function (l, i) {
    var org = l.OrganizationalStructure || l.organizationalStructure || {};
    var media = l.Media || l.media || {};
    filas.push(['lead ' + (i + 1), [l.CreationDate || l.creationDate, l.Status || l.status, org.Description || org.description,
      media.Name || media.name, l.UTM_Source || l.utM_Source || '', (l.ClientId || l.clientId) ? 'con cliente' : 'sin cliente'].join(' | ')]);
  });
  escribir(base(), 'evup_prueba', filas);
  Logger.log('Listo: ver la pestaña evup_prueba de la Base comercial.');
}

/**
 * Segunda prueba: qué locales / países hay en EVUP y cuántos leads tiene cada uno.
 * Escribe en la pestaña "evup_estructura" solo nombres de locales y cantidades (sin datos personales).
 */
function probarEstructura() {
  var token = evupToken();
  var estructuras = [], pagina = 1;
  while (pagina <= 40) {
    var r = evup(token, '/OrganizationalStructure/List', { PageNumber: pagina, Filters: [], Sort: [] });
    var items = Array.isArray(r) ? r : (r.items || r.Items || []);
    if (!items.length) break;
    estructuras = estructuras.concat(items);
    if (items.length < 50) break;
    pagina++;
  }
  var filas = [['id', 'local', 'depende de', 'tipo', 'leads totales', 'leads desde 2025']];
  var desde2025 = [{ Field: 'CreationDate', FieldType: 'date', Operation: 'greater_eq', Value: '2025-01-01' }];
  estructuras.forEach(function (e) {
    var id = e.Id || e.id;
    var total = '', reciente = '';
    try { total = cantidad(evup(token, '/Lead/Count', { FilterOrgStructure: id, PageNumber: 1, Filters: [], Sort: [] })); } catch (x) { total = 'error'; }
    try { reciente = cantidad(evup(token, '/Lead/Count', { FilterOrgStructure: id, PageNumber: 1, Filters: desde2025, Sort: [] })); } catch (x) { reciente = 'error'; }
    filas.push([id, e.Description || e.description, e.FatherDescription || e.fatherDescription || '', e.ClassifierName || e.classifierName || '', total, reciente]);
  });
  try { filas.push(['', 'TODOS desde 2025', '', '', '', cantidad(evup(token, '/Lead/Count', { PageNumber: 1, Filters: desde2025, Sort: [] }))]); } catch (x) { filas.push(['', 'TODOS desde 2025', '', '', '', 'error: ' + x.message]); }
  escribir(base(), 'evup_estructura', filas);
  Logger.log('Listo: ' + estructuras.length + ' estructuras. Ver la pestaña evup_estructura de la Base comercial.');
}

function cantidad(r) { return typeof r === 'number' ? r : (r.count != null ? r.count : r.Count != null ? r.Count : JSON.stringify(r)); }
