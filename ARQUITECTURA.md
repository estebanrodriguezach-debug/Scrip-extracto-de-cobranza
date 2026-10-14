# Arquitectura de la automatización de cobranzas

Entregable 1 del pedido: la elección técnica, justificada, y lo que se midió
sobre los archivos reales para tomar cada decisión.

> Todos los números de este documento salen de correr los validadores sobre
> `Ingresos Bancarios Clientes 2025-2026.xlsx`, `Cobranza General 2026 (1).xlsx`
> y `Listado banco-2026.xlsx`. Se pueden reproducir con los comandos de la
> sección [Cómo verificar](#cómo-verificar).

---

## 1. La decisión técnica: Python + openpyxl, con un escritor propio

### Por qué no Google Apps Script

El circuito arranca en un libro de **105 hojas** y termina en otro de **40.505
filas** (29 MB de XML en una sola hoja). Apps Script corre con un límite de 6
minutos por ejecución y una cuota de memoria pensada para planillas de Google,
no para `.xlsx` de este tamaño; además los tres archivos son `.xlsx` nativos, y
convertirlos a formato Google rompería los comentarios y los formatos. Sí es
buen lugar para lo que ya resuelve en este repo: clasificar una fila cuando
alguien escribe una nota en la planilla (`apps_script/Codigo.gs`).

### Por qué no VBA ni Power Query

VBA ataría el proceso a una PC con Excel y a un archivo habilitado para macros,
que es justo lo que complica compartir el libro en SharePoint; y el código
quedaría adentro del mismo archivo que los datos, sin historial ni tests. Power
Query es excelente para **leer** y transformar, pero no sabe escribir de vuelta
en una celda de otro libro: el Módulo 2 necesita completar `Fecha de pago`,
`Monto de Pago` y `Medio de Pago` en filas existentes.

### Por qué no pandas

`pandas` lee cómodo, pero **solo sabe escribir archivos nuevos desde cero**. Y
acá el archivo de destino no es descartable:

| Archivo | Qué tiene adentro que no se puede perder |
|---|---|
| `Cobranza General` | **1.119 comentarios de celda** con anotaciones del equipo, más el VML que los dibuja y los metadatos de comentarios con hilo |
| `Listado banco` | el logo de la empresa (`xl/media/image1.png`), filtros guardados y vistas personalizadas |

### Por qué openpyxl para leer, pero **no** para escribir

Esto se midió, no se supuso. Un `load_workbook()` + `save()` de openpyxl sobre
`Cobranza General` descarta estas partes del archivo:

```
xl/comments1..11.xml + sus _rels     los 1.119 comentarios, reescritos en otro
                                     formato que Excel no lee igual
xl/drawings/vmlDrawing1..11.vml      los globos de los comentarios
xl/commentsmeta0..9                  metadatos de comentarios con hilo
xl/calcChain.xml, xl/metadata        cadena de cálculo y metadatos
xl/sharedStrings.xml                 reescrito
```

Son datos que nadie puede reconstruir. Así que la lectura va con openpyxl (es
sólida y ya estaba en el repo) y la escritura va por un módulo propio,
[`clasificador/xlsx_parche.py`](clasificador/xlsx_parche.py), que trata el
`.xlsx` por lo que es —un zip de XMLs— y **reescribe solo el XML de las hojas
que cambian, tocando solo las celdas pedidas**. Todo lo demás se copia byte a
byte.

Resultado medido al conciliar un mes completo sobre el archivo real:

```
celdas escritas:   1.518
partes del zip perdidas:  NINGUNA
comentarios conservados:  1.119
hojas modificadas: solo la del mes (las otras 20 quedan byte a byte iguales)
```

El XML se edita como texto y no con un parser DOM a propósito:
`ElementTree` renombra los prefijos de namespace al serializar, y el atributo
`mc:Ignorable` los nombra como texto, así que un round-trip "prolijo" rompe el
archivo. El XML de una hoja lo genera Excel y es regular, así que la cirugía por
posición es segura y, sobre todo, no toca nada más.

Tres detalles que el módulo resuelve y que se descubrieron probando:

- **Formatos heredados.** Una celda nueva hereda el estilo de su columna, y si
  el valor es una fecha se busca un estilo que *sea* de fecha (leyendo
  `styles.xml`). Sin eso, las filas anexadas al Listado banco mostraban `46309`
  en lugar de `14/10/2026`.
- **Recálculo.** Las fórmulas se escriben sin valor cacheado y se marca
  `fullCalcOnLoad` en `workbook.xml`, así Excel recalcula al abrir y no hace
  falta tocar `calcChain.xml`.
- **Orden de atributos.** Excel escribe `Id` antes de `Target` en los rels y
  openpyxl al revés. La primera versión asumía un orden y funcionaba con los
  archivos reales pero fallaba con cualquier libro generado por openpyxl; lo
  encontró un test.

### Lo que este diseño no hace

**No agrega hojas nuevas a los archivos existentes.** Crear una hoja obliga a
tocar `workbook.xml`, los rels y `[Content_Types].xml`, que es justo el tipo de
cirugía estructural que no conviene hacerle a un archivo con 1.119 comentarios
adentro. Por eso el tablero de mora, el reporte de conciliación y la
exportación son **archivos aparte**, que además se regeneran en cada corrida y
se comparten e imprimen sin tocar la base.

---

## 2. Qué resultó ser cada archivo

Lo que se encontró al abrirlos, y que cambió varias decisiones del diseño.

### `Cobranza General 2026`

- 21 hojas: 10 mensuales (`Enero`..`Octubre`), 10 de pagarés y `Adelantos`.
- **El encabezado está en la fila 2**, y las columnas A, B y C **no tienen
  título**. Por eso las columnas se declaran por letra en `conciliacion.json` y
  el encabezado solo se usa para avisar si algo se corrió de lugar.
- **Abajo de la tabla hay un bloque de estadísticas** en las mismas columnas
  (`Cobrado`, `Pendiente`, `Baja` con totales y porcentajes). No son clientes:
  se reconocen porque no tienen nombre en la columna B. Sin filtrarlas, en
  Enero entran 885 filas donde hay 863 clientes.
- `Cuota actual` (J) **es un importe**, no un número de cuota. Eso hace que la
  regla `Monto de Pago >= Cuota actual` del pedido tenga sentido.
- `Medio de Pago` (I) usa 'rio', 'siro', 'galicia', 'roela', 'macro',
  'patagonia'.
- **`Estado` tiene un tercer valor que no estaba en la especificación: 3**,
  cargado a mano en 620 filas, con la palabra 'baja' escrita en la columna de
  fecha de pago. Son contratos dados de baja. Ni la conciliación, ni la regla de
  estado, ni el tablero de mora los tocan.
- La columna `Telefono` tiene celdas con formato de fecha y un número
  imposible; openpyxl las reporta como error. Se leen como texto.

### `Listado banco-2026`, hoja `H1`

27.434 filas con datos, desde 2020. Medido sobre las **5.040 filas de 2026**,
que son las que muestran la práctica actual:

| Columna | Qué lleva de verdad |
|---|---|
| `ss` | **vacía** (vacía en 5.040 de 5.040) |
| `Cuenta` | **vacía** en todo 2026 (`005989/5` aparece solo en filas viejas) |
| `Cliente` | código de 9 dígitos, ver abajo |
| `Importe` | el monto cobrado |
| `Fecha de Proceso` | el día en que se armó el lote |
| `Fecha de Pago` | la fecha real del pago |
| `Fecha de Acred.` | igual a la de proceso en el 62% de los casos |
| `Canal` | `TR <BANCO>` para transferencias, o el código de la red |
| `SubTotal` | **vacía**, salvo cuando una transferencia pagó varias cuotas: ahí lleva el total del grupo |

### El código `Cliente`: la regla, medida

El pedido decía que el contrato "se transforma según reglas del sistema destino
(ej. a formatos como `2161.0`, `500008202.0` o relleno de ceros)". La regla
resultó ser una sola:

```
Cliente = prefijo (1 dígito) + contrato con ceros hasta 7 dígitos + sufijo (1 dígito)
```

Verificado por tres caminos independientes:

1. En `H1`, de los 27.020 códigos de 9 dígitos, **el tramo central (dígitos 2 a
   8) es un contrato conocido de Cobranza General en el 96,5% de las filas de
   2026**. Probando otras ventanas del código, el acierto cae a 1-2%.
2. El bloque SIRO del libro de Ingresos Bancarios trae el código **ya armado**
   en su columna C, y en el **99,3%** de esas filas el tramo central coincide
   con el contrato de la misma fila.
3. Ejemplos completos: contrato 1313 → `100013131`; contrato 1012 → `000010125`.

Sobre el `2161` del pedido: en `H1` existe `000002161`, pero ese código
corresponde al contrato **216**, no al 2161 (el contrato 2161 daría
`?00021611`). Ni 216 ni 2161 están en Cobranza General 2026 —esa fila de `H1`
es de 2021—, así que el ejemplo del pedido no era un par contrato/código real.
El `.0` es como Excel muestra un número, no parte del código.

### El hallazgo que cierra el circuito: `Detalle` son los códigos de canal

El README de este repo trataba las notas `PF`, `TQR`, `RP`, `FSF` de la columna
`Detalle` como "comentarios sueltos" que caían en la categoría por defecto.
**No son comentarios: son los códigos de `Canal` del Listado banco.** La cuenta
cierra:

```
Cobranza General 2026, Medio de Pago = 'siro'           1.194 filas
Listado banco 2026, canales de la red SIRO              1.144 filas
  PF 731 + TQR 174 + RP 107 + FSF 83 + PP 23 + RSF 17 + CE 3 + CEF 2 + PC 4
```

Y los bancos coinciden uno a uno: `rio` 2.656 ≈ `TR RIO` 2.686, `galicia` 926 ≈
`TR GALICIA` 964, `macro` 81 ≈ `TR MACRO` 88, `patagonia` 38 ≈ `TR PATAGONIA`
40. Es decir: SIRO es la red que cobra a través de Banco Roela, y el código de
la columna `Detalle` dice por qué canal entró cada cobro. Sin ese dato, los
1.194 cobros de SIRO del año no se pueden exportar con su canal correcto.

### El bloque de Roela no son cheques

`origen.py` llamaba "cheques" al bloque izquierdo de las hojas de Roela porque
la columna C parecía un número de cheque. El encabezado real de la hoja dice
**`CUPON PAGO FACIL-RAPIPAGO`**: es el bloque de cobranzas de SIRO, y sus
columnas valen oro para los Módulos 2 y 4:

```
A  fecha de carga        E  fecha de acreditación
B  importe               F  código de canal (PF, TQR, RP, FSF, PP...)
C  código de cliente     G  número de contrato
D  fecha de pago         H  nombre del cliente
```

Eso explica, además, por qué el README había medido que la fecha de la columna D
acertaba el 92% y la A sólo el 14%: la D es la fecha de pago y la A la de carga.

---

## 3. Los cuatro módulos

```
Ingresos Bancarios          Cobranza General              Listado banco
  105 hojas, 4 layouts        21 hojas, 1.119 comentarios   40.505 filas, 1 logo
        │                            │                            │
        ▼                            │                            │
  ┌───────────────┐                  │                            │
  │ 1 extracto.py │ ─ 11.162 movimientos normalizados             │
  └───────┬───────┘                  │                            │
          │                          ▼                            │
          └──────────▶ ┌──────────────────────┐                   │
                       │ 2 conciliacion.py    │ ──▶ parche de celdas G/H/I
                       └──────────┬───────────┘     + reporte de no identificados
                                  │                            │
                       ┌──────────▼───────────┐                │
                       │ 3 estado.py / mora.py│ ──▶ parche de la columna L
                       └──────────┬───────────┘     + tablero de Mora Activa
                                  │                            │
                       ┌──────────▼───────────┐                ▼
                       │ 4 listado.py         │ ──▶ filas nuevas de H1
                       └──────────────────────┘     (archivo aparte o anexadas)
```

### Módulo 1 — Estandarizador

Reusa la detección de layouts de `origen.py`, que ya estaba contrastada contra
el extracto real (11.162 movimientos, 98,1% de las filas ya clasificadas a mano
reencontradas), y le agrega lo que ese código descartaba. Sobre el archivo real:

```
Movimientos: 11.162   (Rio 5.691 · Roela 3.162 · Galicia 1.991 · Macro 249 · Patagonia 64 · Frances 5)
  con contrato:                        11.111  (99,5%)
  con CUIT/CVU:                         6.022  (54,0%)
  con código de canal:                  3.083  (27,6%)
  con código de cliente del destino:    3.075  (27,5%)
```

El CUIT y el CVU se validan **por dígito verificador**, no por forma: un
contrato de 4 dígitos pegado a un importe puede parecer un CUIT de 11, y
validar el verificador descarta esos falsos positivos (la probabilidad de que 11
dígitos al azar pasen el chequeo es ~1/11). El contrato sacado del texto libre
**no acepta números sueltos en medio de una frase**: el concepto de un banco
está lleno de números que no son contratos, y un match equivocado mete una
cobranza en la ficha de otro cliente, que es el peor error posible de este
circuito.

### Módulo 2 — Conciliación

Cuatro niveles, de más a menos certeza:

| Nivel | Criterio | Aplicados sobre el año real |
|---|---|---|
| 1 | contrato exacto + monto = cuota | 3.835 |
| 1c | contrato exacto + el monto cubre **varias** cuotas pendientes | 257 |
| 1b | contrato exacto, monto distinto (parcial o de más) | 809 |
| 2 | nombre parecido, con un único candidato claro | 1 |

El nivel **1c** salió de medir: la primera versión del motor llenaba 845 filas
con un monto distinto al que había puesto el equipo, y casi todas eran lo mismo
—una transferencia de 90.000 que paga tres cuotas de 30.000—. Ahora se reparte,
y cada fila recibe su cuota con la misma fecha y medio. Es también el origen de
la columna `SubTotal` del Listado banco.

Decisiones que vale la pena conocer:

- **Uno a uno.** Cada movimiento se aplica a una sola fila y cada fila recibe un
  solo movimiento. Sin eso, un cliente con tres cuotas iguales en el mes se
  comería tres veces el mismo depósito.
- **Filtro por año, imprescindible.** El libro de origen cubre 2025 y 2026. Un
  cobro de enero de 2025 suele tener el mismo monto que la cuota de enero de
  2026, así que sin filtrar el cruce se equivoca de fila sin que se note: la
  primera versión reproducía el 24% del trabajo manual, y con el filtro pasó al
  75%. El año de cada hoja se deduce de las fechas que tiene adentro, no del
  nombre.
- **La hoja destino sale del nombre de la hoja de origen**, no de la fecha del
  pago: un cupón pagado el 30/12 que figura en "Roela Enero" paga la cuota de
  enero. Si no aparece nada ahí, se prueban los meses vecinos.
- **Greedy, no óptimo.** Se podría plantear como asignación de costo mínimo,
  pero el resultado sería más difícil de auditar. Para quien revisa, "se aplicó
  por contrato y monto exacto" vale más que un apareo global imposible de
  reconstruir a mano.
- **Orden determinista**: dos corridas sobre el mismo archivo dan el mismo
  resultado.

### Módulo 3 — Estado y mora

`Estado` pasa de `=COUNT(G3)` —que mide si *alguien cargó el dato*— a la regla
pedida, escrita como **fórmula** para que la planilla siga viva:

```excel
=IF(N(H3)<=0,0,IF(N(H3)>=N(J3),1,"Parcial"))
```

Sobre el archivo real: **4.828 pagados, 205 parciales, 2.482 pendientes, 671
bajas**. Los 205 parciales eran invisibles con la regla vieja: tenían fecha de
pago, así que contaban como pagados.

El tablero de mora cruza **todos los meses**, porque "Cuotas Adeudadas" no tiene
sentido dentro de una sola hoja: el cliente que no pagó marzo ni abril tiene dos
cuotas en dos pestañas distintas. Al 14/10/2026: **433 clientes, 2.401 cuotas,
$294.727.600 vencidos**.

Si falta `Vto Cuota` (está cargada en ~40% de las filas) se usa el fin del mes
de la hoja, que es lo que el equipo asume al mirar la planilla. Se puede apagar
en la config.

### Módulo 4 — Exportador

Fidelidad contra las 5.040 filas de 2026 ya cargadas a mano:

```
código exacto del pago                      75,9%
código que ese cliente usa de verdad        11,3%   -> 87,3% válido para el cliente
prefijo correcto, sufijo nunca visto        11,8%
prefijo distinto                             0,9%
Canal correcto                              99,7%
```

**El sufijo es el único dato que no se puede deducir de los tres archivos.** Se
probaron las hipótesis razonables contra las 23.584 parejas de pagos
consecutivos del historial:

| Hipótesis | Acierto |
|---|---|
| el sufijo cicla 1..6 y vuelve a 1 con cada cuota | 28,5% (17,4% en 2026) |
| el sufijo es el orden del pago del contrato | 19,8% |
| el sufijo es fijo por contrato | vale para el 56,7% de los contratos |

Así que el módulo **no lo adivina**. Hace lo único defendible: usa el código que
el extracto trae armado cuando existe (cobros SIRO, exacto); si no, el
(prefijo, sufijo) que ese contrato más usó en el historial del Listado banco,
que es la fuente de verdad del sistema destino; y si el contrato nunca apareció,
los valores por defecto de la config **marcando la fila para revisión**. En una
corrida típica de un mes: 141 filas, de las cuales 8 marcadas.

👉 **Esto es lo único que requiere una respuesta de quien administra el sistema
destino.** Cuando confirmen qué significa el sufijo, se carga una vez en
`conciliacion.json` (`mapa_cliente` o los valores por defecto) y el módulo deja
de marcar filas.

Una desviación deliberada del pedido: el pedido decía exportar los "registros
conciliados (`Estado = 1`)", y el exportador incluye **también los parciales**,
porque un pago parcial es plata que entró y el sistema destino la tiene que ver
igual —el propio historial de `H1` tiene importes más chicos que la cuota—. Se
puede volver al criterio literal con `--solo-pagados`.

---

## 4. Cómo verificar

Los dos validadores miden el circuito contra el trabajo manual ya hecho, que es
la única referencia confiable que hay. Es el mismo criterio que ya usaba
`scripts/validar_contra_excel.py` en este repo.

```bash
# Modulo 2: toma las filas que el equipo ya concilio a mano, las vuelve a
# dejar pendientes, y compara lo que el motor habria escrito.
python3 scripts/validar_conciliacion.py INGRESOS.xlsx COBRANZA.xlsx

# Modulo 4: compara los codigos y canales generados contra las filas de H1.
python3 scripts/validar_listado.py INGRESOS.xlsx COBRANZA.xlsx LISTADO.xlsx

# Tests (no necesitan los archivos reales: usan libros sinteticos)
python3 -m unittest discover -s tests
```

Resultado del Módulo 2 sobre las 5.046 filas conciliadas a mano:

```
reproducida                3.808  (75,5%)   misma fila, mismo monto, misma fecha
mismo monto, otra fecha      471  ( 9,3%)   el cliente hizo dos pagos iguales en el mes
                                          -> monto correcto: 4.279 (84,8%)
otro monto, múltiplo         134  ( 2,7%)   el equipo repartió el pago entre meses
otro monto                   185  ( 3,7%)
sin aplicar                  448  ( 8,9%)   queda para revisión humana
```

El 15,7% que no se reproduce no es "error" del motor en todos los casos: la
causa dominante son pagos que el equipo repartió entre meses distintos o
agregó de varios movimientos en una fila. El motor los deja marcados en lugar
de adivinar.

---

## 5. Dónde se cambia cada cosa

| Qué | Dónde |
|---|---|
| Letras de columna, fila del encabezado, meses | `conciliacion.json` → `cobranza_general` |
| Tolerancias de monto, umbral de nombre, meses vecinos | `conciliacion.json` → `conciliacion` |
| Días de gracia de la mora, usar fin de mes | `conciliacion.json` → `mora` |
| Códigos de canal, prefijo/sufijo por defecto, equivalencias | `conciliacion.json` → `listado` |
| Reglas de categoría del extracto | `reglas.json` (como antes) |

Nada de eso está en el código: si el Excel cambia de forma, se edita el JSON.
