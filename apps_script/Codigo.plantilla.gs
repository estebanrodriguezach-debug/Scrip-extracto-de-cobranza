/**
 * Clasificador de cobranzas dentro de la planilla.
 *
 * NO EDITAR Codigo.gs a mano: se genera desde Codigo.plantilla.gs + reglas.json
 * con `python3 scripts/generar_apps_script.py`.
 *
 * Que hace:
 *  - Menu "Cobranzas" para clasificar todo o solo lo que falta.
 *  - onEdit: al escribir una nota en Detalle (col E), completa Categoria (col F).
 *  - Recrea el desplegable de categorias desde la lista de reglas.
 */

/*__REGLAS__*/

var HOJA_DATOS = 'Cobranzas';
var HOJA_LISTAS = 'Listas';
var COL_DETALLE = 5; // E
var COL_CATEGORIA = 6; // F
var PRIMERA_FILA = 2;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cobranzas')
    .addItem('Clasificar filas sin categoría', 'clasificarPendientes')
    .addItem('Reclasificar TODO (pisa manuales)', 'clasificarTodo')
    .addSeparator()
    .addItem('Recrear desplegable de categorías', 'recrearValidacion')
    .addToUi();
}

/** Normaliza igual que clasificador/reglas.py: minúsculas, sin acentos, espacios colapsados. */
function normalizar(texto) {
  if (texto === null || texto === undefined) return '';
  return String(texto)
    .normalize('NFKD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** Devuelve la categoría para una nota de Detalle. */
function clasificar(detalle) {
  var n = normalizar(detalle);
  if (!n) return REGLAS.categoria_sin_dato;
  for (var i = 0; i < REGLAS.reglas.length; i++) {
    var regla = REGLAS.reglas[i];
    for (var j = 0; j < regla.contiene.length; j++) {
      if (n.indexOf(normalizar(regla.contiene[j])) !== -1) return regla.categoria;
    }
  }
  return REGLAS.categoria_por_defecto;
}

function _hojaDatos() {
  var hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_DATOS);
  if (!hoja) throw new Error('No existe la hoja "' + HOJA_DATOS + '".');
  return hoja;
}

function _clasificarHoja(soloVacias) {
  var hoja = _hojaDatos();
  var ultima = hoja.getLastRow();
  if (ultima < PRIMERA_FILA) return 0;

  var filas = ultima - PRIMERA_FILA + 1;
  var detalles = hoja.getRange(PRIMERA_FILA, COL_DETALLE, filas, 1).getValues();
  var rango = hoja.getRange(PRIMERA_FILA, COL_CATEGORIA, filas, 1);
  var categorias = rango.getValues();

  var cambios = 0;
  for (var i = 0; i < filas; i++) {
    var actual = String(categorias[i][0] || '').trim();
    if (soloVacias && actual) continue;
    var nueva = clasificar(detalles[i][0]);
    if (nueva !== actual) {
      categorias[i][0] = nueva;
      cambios++;
    }
  }
  if (cambios) rango.setValues(categorias);
  return cambios;
}

function clasificarPendientes() {
  var n = _clasificarHoja(true);
  SpreadsheetApp.getActive().toast(n + ' fila(s) clasificada(s).', 'Cobranzas', 5);
}

function clasificarTodo() {
  var n = _clasificarHoja(false);
  SpreadsheetApp.getActive().toast(n + ' fila(s) actualizada(s).', 'Cobranzas', 5);
}

/** Completa la categoría al cargar o editar una nota en Detalle. */
function onEdit(e) {
  if (!e || !e.range) return;
  var hoja = e.range.getSheet();
  if (hoja.getName() !== HOJA_DATOS) return;
  if (e.range.getColumn() !== COL_DETALLE) return;
  if (e.range.getRow() < PRIMERA_FILA) return;

  var filas = e.range.getNumRows();
  var detalles = e.range.getValues();
  var destino = hoja.getRange(e.range.getRow(), COL_CATEGORIA, filas, 1);
  var categorias = destino.getValues();
  for (var i = 0; i < filas; i++) {
    categorias[i][0] = clasificar(detalles[i][0]);
  }
  destino.setValues(categorias);
}

function recrearValidacion() {
  var ss = SpreadsheetApp.getActive();
  var listas = ss.getSheetByName(HOJA_LISTAS) || ss.insertSheet(HOJA_LISTAS);
  listas.clear();
  var cats = REGLAS.categorias_validas.map(function (c) {
    return [c];
  });
  listas.getRange(1, 1, cats.length, 1).setValues(cats);

  var hoja = _hojaDatos();
  var filas = Math.max(hoja.getMaxRows() - PRIMERA_FILA + 1, 1);
  var regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(REGLAS.categorias_validas, true)
    .setAllowInvalid(true)
    .build();
  hoja.getRange(PRIMERA_FILA, COL_CATEGORIA, filas, 1).setDataValidation(regla);
  hoja.setFrozenRows(1);
  ss.toast('Desplegable actualizado.', 'Cobranzas', 5);
}
