# Automatización del proceso de cobranzas

Dos circuitos sobre los mismos datos:

**A. Conciliación de cobranzas (Módulos 1 a 4).** Lee los extractos de todos los
bancos, los cruza con `Cobranza General`, completa los cobros, recalcula el
estado de cada cuota, arma el tablero de mora y genera las filas del
`Listado banco`. Reemplaza el trabajo manual de 400 a 600 transacciones por mes.

- 👉 **Para usarlo:** [GUIA_OPERATIVA.md](GUIA_OPERATIVA.md)
- 👉 **Cómo está hecho y por qué:** [ARQUITECTURA.md](ARQUITECTURA.md)

```bash
python3 -m clasificador.cli conciliar INGRESOS.xlsx --cobranza COBRANZA.xlsx --mes Octubre --con-estado
python3 -m clasificador.cli mora --cobranza COBRANZA.xlsx
python3 -m clasificador.cli exportar-listado --cobranza COBRANZA.xlsx --listado LISTADO.xlsx --ingresos INGRESOS.xlsx
```

Medido contra el trabajo manual ya hecho sobre los archivos reales:

| | |
|---|---|
| Movimientos leídos del extracto | 11.162 (105 hojas, 6 bancos) |
| Filas con contrato detectado | 99,5% |
| Conciliación: monto correcto | **84,8%** de las 5.046 filas conciliadas a mano |
| Listado banco: código de cliente | **87,3%** válido para ese cliente (75,9% exacto) |
| Listado banco: canal | **99,7%** |
| Comentarios de celda conservados al escribir | 1.119 de 1.119 |

**B. Extracto de cobranzas → Google Sheets / SharePoint.** Vuelca la base
`Extractos_cobranzas.xlsx` a una planilla de Google (o a una carpeta de
SharePoint) y clasifica cada movimiento **con las mismas reglas que se venían
aplicando a mano**. Es lo que ya hacía este repo, y sigue funcionando igual.

Las reglas de clasificación viven en un único archivo,
[`reglas.json`](reglas.json), que usan tanto el script de Python como el Apps
Script que corre dentro de la planilla. Si cambiás una regla, cambia en los dos
lados.

---

## Circuito B: extracto a Google Sheets / SharePoint

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

## Volcar a SharePoint

El mismo extracto, con el mismo formato, pero dentro de una carpeta de
SharePoint en lugar de (o además de) Google Sheets:

```bash
export SP_TENANT_ID=c9519bc6-8190-4466-a0d7-d30d923582b7
export SP_CLIENT_ID=...
export SP_CLIENT_SECRET=...

python3 -m clasificador.cli volcar-sharepoint --drive-id 1IY5waCA...i13op \
  --carpeta-url "https://TUEMPRESA.sharepoint.com/:f:/s/TUSITIO/IgBxxxxxxxxxxxxxxxxxxxxx?e=XXXXXX"
```

El archivo que queda en la carpeta es un `.xlsx` normal: hoja **Cobranzas** con
las 6 columnas, hoja **Listas** con las categorías, **desplegable** sobre la
columna `Categoría` y fila 1 congelada. Igual que la planilla de Google — de
hecho es más fiel, porque el archivo se sube tal cual y SharePoint no lo toca.

### El link caduca: vinculá una vez y usá la ruta del sitio

El link compartido lleva un `?e=...` que se puede revocar o vencer. Para dejarlo
en un cron conviene no depender de él:

```bash
python3 -m clasificador.cli vincular-sharepoint "https://...:f:/s/..."
# Carpeta: Extractos
#   drive-id: b!xxxx
#   item-id : 01ABC...

python3 -m clasificador.cli volcar-sharepoint Extractos_cobranzas.xlsx \
  --sitio TUEMPRESA.sharepoint.com:/sites/TUSITIO \
  --subcarpeta "Documentos compartidos/Extractos"
```

La ruta del sitio no caduca. Además es el **único** camino que funciona con el
permiso `Sites.Selected`: el endpoint `/shares`, que resuelve el link, no lo
acepta.

### Setup de Entra ID, paso a paso

Todo esto se hace **una sola vez**. Los pasos 4 y 5 los tiene que hacer un
administrador del tenant: no hay forma de evitarlo, el consentimiento de
administrador es obligatorio para cualquier app que escriba con su propia
identidad.

1. Entrá a [entra.microsoft.com](https://entra.microsoft.com) → **Entra ID** →
   **App registrations** → **New registration**. (Ya no se hace desde el portal
   clásico de Azure.) Hace falta ser al menos *Application Developer*.
2. Nombre: `extractos-cobranza`. En *Supported account types*, **Single tenant**.
   El redirect URI se deja vacío.
3. En **Overview** anotá *Application (client) ID* y *Directory (tenant) ID* →
   son `SP_CLIENT_ID` y `SP_TENANT_ID`.
4. **Certificates & secrets → New client secret**. El valor se muestra una sola
   vez: es `SP_CLIENT_SECRET`. **Caduca** (6 o 24 meses según la política de la
   empresa); cuando vence el script deja de andar sin aviso previo, así que
   anotá la fecha. Si el admin lo permite, un certificado no caduca tan seguido.
5. **API permissions → Add a permission → Microsoft Graph → Application
   permissions → `Sites.Selected`**. Después, *Grant admin consent for \<tenant\>*.
6. El grant por sitio, que es un segundo pedido al admin y el paso donde más
   implementaciones quedan a mitad de camino. `Sites.Selected` por sí solo **no
   da acceso a nada**: hay que habilitar la app sobre ese sitio puntual. Requiere
   rol *SharePoint Administrator*:

   ```powershell
   Grant-PnPAzureADAppSitePermission -AppId <client-id> `
     -DisplayName "extractos-cobranza" -Permissions Write `
     -Site https://TUEMPRESA.sharepoint.com/sites/TUSITIO
   ```

   o el equivalente por Graph:

   ```http
   POST https://graph.microsoft.com/v1.0/sites/{siteId}/permissions
   {"roles":["write"],
    "grantedToIdentities":[{"application":{"id":"<client-id>","displayName":"extractos-cobranza"}}]}
   ```

**Por qué `Sites.Selected` y no `Sites.ReadWrite.All` o `Files.ReadWrite.All`:**
esos dos le dan al script escritura sobre **todos** los sitios y archivos de la
empresa. `Sites.Selected` arranca sin acceso a nada y el admin lo habilita sólo
sobre Administración y Finanzas, revocable con un `DELETE`. Es el argumento que
hace que seguridad diga que sí.

### Sin secreto: device code

Si no te dejan crear un client secret, dejá `SP_CLIENT_SECRET` sin definir. El
script imprime un código, lo pegás en
[microsoft.com/devicelogin](https://microsoft.com/devicelogin) y listo; el token
queda cacheado en `~/.cache/cobranzas/token_sharepoint.json`. Requiere
*Allow public client flows = Yes* en la pestaña **Authentication** de la app.

**No sirve para un cron:** el refresh token caduca (ventana deslizante de ~90
días) y se rompe con cualquier cambio de contraseña o política de acceso
condicional. Y desde el cambio de defaults de Microsoft de 2025 (MC1097272), en
la mayoría de los tenants este camino **también** pide consentimiento de
administrador, así que no evita el trámite.

### Si la organización bloquea el registro de apps

Es muy común (`Users can register applications = No`). Opciones, de mejor a peor:

1. **Que un admin te asigne el rol *Application Developer*.** Es el remedio que
   documenta Microsoft y no concede nada más que eso.
2. **Que el admin cree el registro y te agregue como *owner*.** Podés manejar
   secretos y permisos sin poder crear apps nuevas.
3. **Activar el *admin consent workflow*,** para pedir la aprobación desde el
   propio flujo en vez de por mail.
4. **Sin ninguna app: sincronizar la biblioteca con el cliente de OneDrive** y
   escribir el `.xlsx` en la carpeta local que queda sincronizada. Cero API,
   cero permisos, cero consentimiento, y el extracto aparece en la misma carpeta
   con el mismo formato. Es menos elegante, pero si el tenant está cerrado es la
   salida más rápida, y conviene tenerla sobre la mesa antes de pasar dos
   semanas trabado en Entra.

### Diferencias con el destino Google Sheets

| | Google Sheets | SharePoint |
|---|---|---|
| Qué se escribe | filas nuevas, in situ | el archivo entero, de nuevo |
| Categorías corregidas a mano | se respetan | se respetan (merge por clave) |
| Columnas o hojas agregadas a mano | se conservan | **se pierden** |
| Clasificación automática al escribir | sí (Apps Script) | no hay equivalente |

Como el `.xlsx` se **regenera entero** en cada volcado, la hoja `Cobranzas` es
territorio del script: lo único que sobrevive de las ediciones manuales es la
columna `Categoría` (y el `Detalle`), porque antes de subir se baja el archivo
que está en SharePoint y se mezcla por la misma clave que usa el volcado a
Sheets — fecha + banco + importe + texto del extracto. Si alguien agrega una
columna `G` con notas, una hoja aparte o un gráfico, el próximo volcado se lo
lleva puesto.

Dos detalles operativos:

- Si alguien tiene el `.xlsx` abierto en **Excel de escritorio**, el archivo
  queda bloqueado y la subida falla con `423`. El mensaje de error te lo dice:
  hay que cerrarlo y reintentar.
- Si el archivo cambió en SharePoint mientras corría el script, la subida falla
  con `412` en lugar de pisar el cambio. Volvé a correr el comando y se mergea.

---

## Dejar el extracto en SharePoint vía OneDrive (lo más simple)

Es el camino recomendado: **no necesita ningún permiso ni administrador**.

1. En el navegador, abrí la biblioteca de SharePoint donde va el extracto y
   tocá **Sincronizar**. La carpeta aparece en el Explorador, dentro de
   `OneDrive - <tu empresa>`.
2. Copiá esa ruta del Explorador.
3. Corré:

```bash
python3 -m clasificador.cli volcar-carpeta \
  --drive-id 1IY5waCA...i13op \
  --carpeta "C:/Users/TU_USUARIO/OneDrive - TU EMPRESA/Administracion y Finanzas/Extractos"
```

El script lee Ingresos Bancarios, clasifica, escribe el
`Extractos_cobranzas.xlsx` en esa carpeta y **OneDrive lo sube solo**.

Detalles que importan:

- **Es idempotente**: volver a correrlo con el mismo origen no duplica nada.
- **No pisa las correcciones a mano**: si alguien cambia una categoría en la
  planilla, se conserva. Funciona porque `Movimiento.clave` no depende de
  `Detalle` ni de `Categoría`, así que editarlas no genera una fila nueva.
  Para recalcular igual, usá `--reclasificar-todo`.
- Se escribe primero a un `.tmp` en la misma carpeta y recién después se
  renombra, para que OneDrive no empiece a subir un archivo a medio escribir.

> Como corre en la PC que tiene OneDrive, ahí necesitás Python y las
> credenciales de Google para leer el origen. Si preferís no instalar nada en
> esa máquina, podés correrlo en Cloud Shell contra un `.xlsx` local y después
> copiar el archivo a la carpeta sincronizada.

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
python3 -m unittest discover -s tests          # 164 tests, no necesitan los .xlsx reales
node tests/test_apps_script.js

# Validadores: miden el circuito contra los archivos reales (opcionales)
python3 scripts/validar_contra_excel.py Extractos_cobranzas.xlsx
python3 scripts/validar_conciliacion.py INGRESOS.xlsx COBRANZA.xlsx
python3 scripts/validar_listado.py INGRESOS.xlsx COBRANZA.xlsx LISTADO.xlsx
```

`tests/fixtures.py` arma libros `.xlsx` sintéticos con la misma forma que los
reales (encabezado en la fila 2, el bloque de estadísticas, las fórmulas, la
hoja H1), así los tests no dependen de archivos con datos de clientes.

`tests/casos.json` es la lista de casos compartida: la usan el test de Python y
el de Apps Script, así los dos clasificadores no se pueden separar.

## Estructura

```
reglas.json                      # reglas de clasificación (circuito B)
conciliacion.json                # parámetros del circuito de conciliación (A)

# Circuito A: conciliación de cobranzas
clasificador/config.py           # carga de conciliacion.json
clasificador/patrones.py         # regex de CUIT/CVU/contrato y comparación de nombres
clasificador/extracto.py         # MÓDULO 1: estandariza los extractos bancarios
clasificador/cobranza.py         # lectura de Cobranza General (encabezado en fila 2)
clasificador/conciliacion.py     # MÓDULO 2: motor de conciliación
clasificador/estado.py           # MÓDULO 3: regla de Estado (1 / Parcial / 0)
clasificador/mora.py             # MÓDULO 3: tablero de Mora Activa
clasificador/listado.py          # MÓDULO 4: exportador a la hoja H1
clasificador/xlsx_parche.py      # escritura quirúrgica: no pierde comentarios ni formatos
scripts/validar_conciliacion.py  # mide el Módulo 2 contra el trabajo manual
scripts/validar_listado.py       # mide el Módulo 4 contra el Listado banco real

# Circuito B: extracto a Sheets / SharePoint
clasificador/reglas.py           # motor de reglas
clasificador/excel.py            # lectura del .xlsx
clasificador/drive.py            # baja el .xlsx de origen desde Drive
clasificador/origen.py           # lee el libro de Ingresos Bancarios (105 hojas)
clasificador/sheets.py           # Google Sheets API (volcar / reclasificar)
clasificador/libro.py            # arma el .xlsx del extracto con openpyxl
clasificador/sharepoint.py       # Microsoft Graph (volcar a SharePoint)
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
