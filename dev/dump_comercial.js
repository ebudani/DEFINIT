// Genera dev/comercial.json a partir de un Excel "Descuentos por locales" local,
// con el mismo código que usa Apps Script (src/Comercial.js). Solo para desarrollo.
//
// Uso: descomprimí el .xlsx (es un zip) y pasá la tabla de textos y la hoja de detalle:
//   node dev/dump_comercial.js xl/sharedStrings.xml xl/worksheets/sheet2.xml
const fs = require('fs');
const path = require('path');
const C = require('../src/Comercial.js');

const [ss, hoja] = process.argv.slice(2);
const r = C.agregarCobros(C.leerHojaXlsx(fs.readFileSync(hoja, 'utf8'), C.leerSharedStrings(fs.readFileSync(ss, 'utf8'))));
r.archivo = 'Descuentos por locales (prueba local)';
r.actualizadoArchivo = new Date().toISOString().slice(0, 16).replace('T', ' ');
fs.writeFileSync(path.join(__dirname, 'comercial.json'), JSON.stringify(r));
console.log('dev/comercial.json listo:', r.meses.length, 'meses,', r.locales.length, 'locales');
