// Arma dev/index.html: el mismo tablero que sirve Apps Script, pero con los
// datos de dev/raw.json y un google.script.run simulado. Solo para desarrollo.
//
// Uso: node dev/build-preview.js
const fs = require('fs');
const path = require('path');
const { construirModelo } = require('../src/Parser.js');

const raiz = path.join(__dirname, '..');
const raw = JSON.parse(fs.readFileSync(path.join(__dirname, 'raw.json'), 'utf8'));
const { 'Ventas Mensuales': resumen, ...pestanas } = raw;
const modelo = construirModelo(resumen, pestanas);

const incluir = (nombre) => fs.readFileSync(path.join(raiz, 'src', nombre + '.html'), 'utf8');
let html = incluir('Index').replace(/<\?!=\s*include\('(\w+)'\);?\s*\?>/g, (_, n) => incluir(n));

const mock = `<script>
window.google = { script: { run: {
  _ok: null, _err: null,
  withSuccessHandler(f) { this._ok = f; return this; },
  withFailureHandler(f) { this._err = f; return this; },
  getData() { const ok = this._ok; setTimeout(() => ok(${JSON.stringify(JSON.stringify(modelo))}), 150); }
} } };
</script>`;
html = html.replace('</head>', mock + '\n</head>');

fs.writeFileSync(path.join(__dirname, 'index.html'), html);
console.log('dev/index.html listo:', modelo.meses.length, 'meses,', Object.keys(modelo.diario).length, 'pestañas diarias, actualizado', modelo.actualizado);
