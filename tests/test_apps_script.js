/**
 * Verifica que el clasificador de Apps Script devuelva lo mismo que el de
 * Python, sobre los mismos casos (tests/casos.json).
 *
 * Uso: node tests/test_apps_script.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const raiz = path.resolve(__dirname, '..');
const fuente = fs.readFileSync(path.join(raiz, 'apps_script', 'Codigo.gs'), 'utf8');
const casos = JSON.parse(fs.readFileSync(path.join(raiz, 'tests', 'casos.json'), 'utf8'));

// Apps Script expone SpreadsheetApp; aca solo necesitamos clasificar().
const contexto = { SpreadsheetApp: undefined, console };
vm.createContext(contexto);
vm.runInContext(fuente, contexto);

let fallos = 0;
for (const caso of casos) {
  const obtenido = contexto.clasificar(caso.detalle);
  if (obtenido !== caso.categoria) {
    fallos++;
    console.error(
      `FALLA detalle=${JSON.stringify(caso.detalle)}\n  esperado: ${caso.categoria}\n  obtenido: ${obtenido}`
    );
  }
}

if (fallos) {
  console.error(`\n${fallos} de ${casos.length} casos fallaron.`);
  process.exit(1);
}
console.log(`OK: ${casos.length} casos, Apps Script coincide con Python.`);
