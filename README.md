# Extracto de cobranzas → Google Sheets

Vuelca la base `Extractos_cobranzas.xlsx` a una planilla de Google y clasifica
cada movimiento **con las mismas reglas que se venían aplicando a mano**.

Las reglas viven en un único archivo, [`reglas.json`](reglas.json), que usan
tanto el script de Python como el Apps Script que corre dentro de la planilla.
Si cambiás una regla, cambia en los dos lados.

---

## Cómo clasifica (lo que ya hacía la base)

La categoría **no** se deduce del texto del banco (`Extracto`). Sale de la nota
manual de la columna `Detalle`. Sobre esa nota, normalizada (minúsculas, sin
acentos, espacios colapsados), se aplican estas reglas **en orden** y gana la
primera que coincide:

| # | Si `Detalle` contiene | Categoría |
|---|---|---|
| 1 | `pagare` (también `pagaré`, `Pagare`) | Cobranza financiamiento clientes (pagaré) |
| 2 | `cuota 1` | Cobranza cartera planes (cuota 1) |
| 3 | `directa` | Ventas Directas |
| 4 | `adjudicacion` | Adjudicaciones |
| 5 | `presupuesto` o `modificacion` | Cambio estructuras u otras modificaciones |
| 6 | `multa` | Saldo cancelatorio |
| 7 | `adelanto` | Adelantos clientes planes |
| — | **cualquier otra cosa, o `Detalle` vacío** | **Cobranza cartera planes** |

El orden importa y replica lo que estaba hecho: `saldo directa` → *Ventas
Directas*, `cuota 1 / base clientes` → *cuota 1*, `directa / BASE CLIENTE` →
*Ventas Directas*.

Las notas que son comentarios sueltos (`PF`, `TQR`, `RP`, `FSF`, `base
clientes`, `corresponde a 3660`, `modificar en ficha`, …) caen en el default
*Cobranza cartera planes*, igual que en la base: 1658 de 1760 filas.

### Fidelidad contra la base original

```
$ python3 scripts/validar_contra_excel.py Extractos_cobranzas.xlsx
Movimientos: 1760
Coinciden:   1757 (99.83%)
Difieren:    3
```

Las 3 diferencias son **contradicciones de la base misma**, no fallas de la
regla — el mismo texto quedó clasificado de dos formas distintas:

- `SALDO` (filas 1107 y 1108) → *Saldo cancelatorio*, pero otras 9 filas con
  `saldo` / `Saldo` / `SALDO` quedaron como *Cobranza cartera planes*. Se
  respetó la mayoría; por eso `saldo` **no** es palabra clave. Si querés que
  `saldo` mande a *Saldo cancelatorio*, agregalo a la regla
  `saldo_cancelatorio` en `reglas.json`.
- `976125` (fila 770) → *Cobranza financiamiento clientes (pagaré)*, pero la
  misma nota en la fila 141 quedó como *Cobranza cartera planes*. No hay regla
  posible: es una corrección puntual a mano.

Esas filas se pueden corregir a mano en la planilla; el script no las pisa
(ver `--pisar-manuales` más abajo).

---

## Instalación

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Credenciales de Google

Proyecto de Google Cloud: **ACH FACTURACION** (`tranquil-wave-503300-d6`).

### Opción A — Cloud Shell (la más rápida, sin claves)

Cloud Shell ya está autenticada con tu usuario, así que **no hace falta crear
una cuenta de servicio ni descargar ningún JSON**. Desde la Cloud Shell del
proyecto:

```bash
git clone https://github.com/estebanrodriguezach-debug/Scrip-extracto-de-cobranza.git
cd Scrip-extracto-de-cobranza
git checkout claude/cool-hawking-m31d63
bash scripts/setup_cloudshell.sh
```

Eso fija el proyecto, habilita **Google Sheets API** y **Google Drive API**, y
arma el entorno de Python. Después, el único paso manual (te va a dar un link
para autorizar):

```bash
gcloud auth application-default login \
  --scopes=https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/drive.file,https://www.googleapis.com/auth/cloud-platform
```

Subí el `.xlsx` a Cloud Shell con **⋮ → Subir** (menú de la terminal). El
archivo cae en tu home, así que movelo al proyecto:

```bash
mv ~/Extractos_cobranzas.xlsx ~/Scrip-extracto-de-cobranza/
```

> **Aviso de Google:** al correr `application-default login` vas a ver un
> warning de que el scope `spreadsheets` *"will be blocked soon for the default
> client ID"*. Hoy funciona igual, pero cuando eso se corte hay que pasar a la
> Opción B (cuenta de servicio), que no depende del client ID de `gcloud`.

### Opción B — Cuenta de servicio (para dejarlo corriendo solo)

En el proyecto **ya existe** `lector-facturas@tranquil-wave-503300-d6.iam.gserviceaccount.com`,
con permiso de **Lector** sobre el archivo de ingresos bancarios. Se puede
reutilizar para el origen; sólo hace falta una clave:

```bash
gcloud iam service-accounts keys create ~/credenciales-cobranzas.json \
  --iam-account=lector-facturas@tranquil-wave-503300-d6.iam.gserviceaccount.com
export GOOGLE_APPLICATION_CREDENTIALS=~/credenciales-cobranzas.json
```

Con eso **lee** el origen. Para **escribir** en la planilla destino, esa cuenta
necesita además permiso de **Editor** sobre ella: o le compartís la planilla a
ese mail, o la creás con `crear --compartir-con`.

Si preferís una cuenta aparte, sólo para esto:

```bash
gcloud config set project tranquil-wave-503300-d6
gcloud services enable sheets.googleapis.com drive.googleapis.com

gcloud iam service-accounts create cobranzas-sheets \
  --display-name="Volcado de cobranzas a Sheets"

gcloud iam service-accounts keys create ~/credenciales-cobranzas.json \
  --iam-account=cobranzas-sheets@tranquil-wave-503300-d6.iam.gserviceaccount.com
```

Después **compartí la planilla como Editor** con
`cobranzas-sheets@tranquil-wave-503300-d6.iam.gserviceaccount.com` (una cuenta
de servicio no ve nada de tu Drive hasta que se lo compartís) y exportá:

```bash
export GOOGLE_APPLICATION_CREDENTIALS=~/credenciales-cobranzas.json
```

> El JSON de la clave es una credencial: guardalo fuera del repo. El
> `.gitignore` ya bloquea `credenciales*.json` y `service-account*.json`.

> **Ojo con `crear` y cuenta de servicio:** una planilla creada por la cuenta
> de servicio nace en *su* Drive, que no tiene interfaz web — no la podrías
> abrir. Usá `--compartir-con tu-mail@gmail.com` para que te la comparta como
> editor, o creá la planilla a mano en tu Drive y pasale el ID a `volcar`.

El ID de la planilla es el tramo de la URL entre `/d/` y `/edit`:
`https://docs.google.com/spreadsheets/d/`**`1AbC...xyz`**`/edit`.

## Uso

**Crear la planilla ya preparada** (hojas, encabezado y desplegable):

```bash
python3 -m clasificador.cli crear --titulo "Extractos de cobranzas"

# con cuenta de servicio, para poder abrirla despues:
python3 -m clasificador.cli crear --titulo "Extractos de cobranzas" \
  --compartir-con tu-mail@gmail.com
```

Te devuelve la URL y el ID. Si ya tenés una planilla, saltealo y usá su ID.

**Volcar el Excel a la planilla y clasificar:**

```bash
python3 -m clasificador.cli volcar Extractos_cobranzas.xlsx \
  --planilla 1AbC...xyz
```

**Leyendo el origen directo de Drive**, sin bajarlo a mano:

```bash
python3 -m clasificador.cli volcar --drive-id 1IY5waCAZUoD-UEzArucVbQ2q_Vsi13op \
  --planilla 1AbC...xyz
```

Esto sirve para el caso real: el archivo fuente (*Ingresos Bancarios Clientes
2025-2026.xlsx*) vive en Drive como **.xlsx, no como planilla nativa**. La
Sheets API no puede leer un .xlsx, pero la Drive API lo descarga tal cual y
`openpyxl` lo abre igual. Así **no hay que convertir el archivo original** ni
pedirle permisos de edición a su dueño: alcanza con que esté compartido como
**Lector** con quien corre el script.

## El libro de origen: 105 hojas y cuatro layouts

`Ingresos Bancarios Clientes` tiene **una hoja por banco y por mes** (`Rio
Enero`, `Macro Feb 26`, `Galicia oct`, `Patag Dic`...), y el formato **no es
uniforme**:

| Banco | Layout |
|---|---|
| Rio, Galicia, Macro, Frances | `A` extracto, `B` fecha, `C` importe, `D` contrato, `E` nombre, `F` detalle |
| Patagonia | **las columnas se corren entre hojas** (en *Agosto 26* la fecha está en `D`, en *Octubre 26* en `C`) y el extracto ocupa dos columnas |
| Roela | **dos tablas lado a lado**: cheques (`A` carga, `B` importe, `D` fecha del cheque, `F` tipo, `G` contrato, `H` nombre) y transferencias (`M` extracto, `N` fecha, `O` importe, `P` contrato, `Q` nombre, `R` detalle) |

Por eso las columnas **se detectan por hoja** en vez de fijarse. Lo único
constante en todo el libro es que *importe, contrato, nombre y detalle* van
pegados a la derecha de la fecha; el bloque de cheques de Roela es la excepción
y está declarada aparte en `clasificador/origen.py`.

Dos detalles que se descubrieron contrastando contra el extracto ya clasificado:

- En los cheques de Roela la fecha que vale es la de la **columna D** (fecha del
  cheque), no la `A` en que se carga: `D` acierta el 92 % y `A` sólo el 14 %.
- En Roela el extracto no tiene texto del banco, así que se usa el **nombre del
  cliente**, igual que en la base original.

### Fidelidad del circuito completo

Leyendo el libro de origen, clasificando y comparando contra las 1760 filas ya
clasificadas a mano del extracto:

```
Movimientos leídos del origen:  11.162  (105 hojas)
Filas del extracto reencontradas: 1726/1760  (98.1%)
   Rio 100% · Galicia 100% · Macro 100% · Patagonia 100% · Roela 92% · Frances 83%
Categoría coincidente:          1723/1726  (99.83%)
```

Las 3 diferencias de categoría son las mismas contradicciones de la base
original ya documentadas (`SALDO` y `976125`). El 2 % de filas no reencontradas
son cheques de Roela cargados con una fecha distinta a la del cheque.

Crea las hojas `Cobranzas` y `Listas` si no existen, pone el encabezado, arma
el desplegable de categorías, congela la fila 1 y agrega los movimientos.

Es **idempotente**: cada movimiento se identifica por fecha + banco + importe +
texto del extracto, así que volver a correrlo con un Excel actualizado agrega
solo lo nuevo. Los cobros repetidos legítimos (misma fecha, banco, importe y
texto — hay 48 en la base) se numeran internamente y **no** se descartan.

Por defecto respeta la categoría que ya venga cargada en el Excel. Para
recalcular todo:

```bash
python3 -m clasificador.cli volcar Extractos_cobranzas.xlsx \
  --planilla 1AbC...xyz --reclasificar-todo
```

**Reclasificar lo que ya está en la planilla** (solo las filas sin categoría,
para no pisar correcciones manuales):

```bash
python3 -m clasificador.cli reclasificar --planilla 1AbC...xyz
```

Agregá `--pisar-manuales` para recalcular también las que ya tienen categoría.

**Probar una regla sin tocar nada:**

```bash
python3 -m clasificador.cli probar "pagare" "saldo directa" "cuota 1"
```

---

## Clasificación dentro de la planilla (Apps Script)

Para que la planilla clasifique sola cuando alguien escribe una nota, sin
depender de correr Python:

1. En la planilla: **Extensiones → Apps Script**.
2. Pegá el contenido de [`apps_script/Codigo.gs`](apps_script/Codigo.gs) en
   `Código.gs` y guardá.
3. Recargá la planilla: aparece el menú **Cobranzas** con
   *Clasificar filas sin categoría*, *Reclasificar TODO* y
   *Recrear desplegable de categorías*.

Además queda un `onEdit`: al escribir o cambiar una nota en la columna
`Detalle` (E), se completa sola la `Categoría` (F) de esa fila.

> `apps_script/Codigo.gs` **se genera**, no se edita a mano. Después de tocar
> `reglas.json`, corré `python3 scripts/generar_apps_script.py` y volvé a pegar
> el archivo en la planilla.

---

## Cambiar las reglas

Todo se edita en [`reglas.json`](reglas.json). Para agregar una palabra clave,
sumala al array `contiene` de la regla correspondiente; para una categoría
nueva, agregala a `categorias_validas` y creá su regla. Recordá que el orden
del array `reglas` define la prioridad.

Después:

```bash
python3 scripts/generar_apps_script.py   # sincroniza el Apps Script
python3 -m unittest discover -s tests    # tests de Python
node tests/test_apps_script.js           # verifica que el .gs dé lo mismo
```

## Tests

```bash
python3 -m unittest discover -s tests -v
node tests/test_apps_script.js
python3 scripts/validar_contra_excel.py Extractos_cobranzas.xlsx  # opcional
```

`tests/casos.json` es la lista de casos compartida: la usan el test de Python y
el de Apps Script, así los dos clasificadores no se pueden separar.

## Estructura

```
reglas.json                      # única fuente de verdad de la clasificación
clasificador/reglas.py           # motor de reglas
clasificador/excel.py            # lectura del .xlsx
clasificador/drive.py            # baja el .xlsx de origen desde Drive
clasificador/origen.py           # lee el libro de Ingresos Bancarios (105 hojas)
clasificador/sheets.py           # Google Sheets API (volcar / reclasificar)
clasificador/cli.py              # línea de comandos
apps_script/Codigo.plantilla.gs  # plantilla editable del Apps Script
apps_script/Codigo.gs            # generado — pegar en la planilla
scripts/generar_apps_script.py   # plantilla + reglas.json -> Codigo.gs
scripts/setup_cloudshell.sh      # deja Cloud Shell listo para correrlo
scripts/inspeccionar_origen.py   # muestra la estructura de un .xlsx de origen
scripts/validar_contra_excel.py  # compara reglas vs. el Excel real
tests/                           # casos compartidos Python / Apps Script
```

> El `.xlsx` con los datos reales **no** se versiona: tiene nombres de clientes
> y CUITs. El `.gitignore` bloquea `*.xlsx` y los JSON de credenciales.
