# Guía de uso para el equipo de cobranzas

Entregable 3 del pedido. Esto es lo único que hace falta leer para usar la
herramienta todos los días. No hay que saber programar.

---

## Antes de empezar (una sola vez)

```bash
python3 -m venv .venv
source .venv/bin/activate        # en Windows:  .venv\Scripts\activate
pip install -r requirements.txt
```

Después, cada vez que abrís una terminal nueva, solo:

```bash
source .venv/bin/activate
```

Y poné los tres archivos en una carpeta a mano. En los ejemplos se usan estos
nombres:

- `Ingresos Bancarios Clientes 2025-2026.xlsx`
- `Cobranza General 2026.xlsx`
- `Listado banco-2026.xlsx`

> **Importante:** cerrá los archivos en Excel antes de correr cualquier comando.
> Si están abiertos, Windows los bloquea y el script no puede escribir.

---

## La rutina de todos los días

Son tres comandos. El primero es el que hace el trabajo.

### 1. Conciliar los pagos del día

```bash
python3 -m clasificador.cli conciliar "Ingresos Bancarios Clientes 2025-2026.xlsx" \
  --cobranza "Cobranza General 2026.xlsx" \
  --mes Octubre \
  --con-estado
```

Esto lee los extractos de todos los bancos, los cruza con la hoja del mes y
completa `Fecha de pago`, `Monto de Pago` y `Medio de Pago` en las filas que
reconoce. Con `--con-estado` además actualiza la columna `Estado`.

**No toca el archivo original.** Escribe una copia al lado:
`Cobranza General 2026 - conciliado.xlsx`. Revisá esa copia y, si está bien,
reemplazá el original.

Lo que te va a mostrar:

```
Filas completadas: 249 (con 245 movimientos)
   1 contrato+monto           195
   1b contrato                45
   1c contrato+varias cuotas  8
   2 nombre                   1
Ambiguos (revisar):   0
No identificados:     538
Pagares (carga manual): 125
```

Y además deja `Conciliacion.xlsx`, **que es tu cola de trabajo**. Tiene cinco
hojas:

| Hoja | Para qué |
|---|---|
| `Resumen` | los totales de la corrida |
| `Aplicados` | qué se cargó y por qué, con la hoja y fila del extracto de origen |
| `Ambiguos` | dos clientes con nombre parecido: elegí vos |
| **`No identificados`** | pagos que nadie reclamó: **esto es lo que hay que mirar** |
| `Pagares` | pagos de pagarés, que van a mano a la hoja `Pagare <Mes>` |

#### Si querés ver qué va a hacer antes de que lo haga

```bash
python3 -m clasificador.cli conciliar "Ingresos Bancarios Clientes 2025-2026.xlsx" \
  --cobranza "Cobranza General 2026.xlsx" --mes Octubre --simular
```

Muestra todo y escribe solo el reporte. No modifica ningún archivo.

### 2. Sacar la lista de morosos para llamar

```bash
python3 -m clasificador.cli mora --cobranza "Cobranza General 2026.xlsx"
```

Deja `Mora_Activa.xlsx`: un cliente por fila, con teléfono, cuántas cuotas debe,
cuánto debe en total, desde cuándo y quién es su supervisor y vendedor.
Ordenado de mayor a menor deuda, con filtros puestos.

Para ver la foto a una fecha distinta de hoy: `--corte 2026-10-31`.

### 3. Armar el archivo para el banco

```bash
python3 -m clasificador.cli exportar-listado \
  --cobranza "Cobranza General 2026.xlsx" \
  --listado "Listado banco-2026.xlsx" \
  --ingresos "Ingresos Bancarios Clientes 2025-2026.xlsx" \
  --desde 2026-10-01
```

Deja `Listado_banco_nuevo.xlsx` con las filas nuevas, en el mismo formato que la
hoja `H1`. Después seguís el circuito de siempre, el que está escrito en la hoja
`INSTRUCCIONES` del propio Listado banco: copiás las filas **sin el
encabezado** y las pegás al final de `H1` con **Pegado especial → Solo valores**
(`Ctrl+Shift+V`).

Pasale siempre `--ingresos`: de ahí salen los códigos de canal (`PF`, `RP`,
`TQR`, `FSF`…) de los cobros por SIRO. Sin eso, esas filas quedan sin `Canal`.

Mirá las tres hojas del archivo:

- `H1`: las filas para copiar.
- **`Revisar`**: las filas con algo para confirmar. Normalmente son pocas.
- `Procedencia`: de dónde salió el código de cada fila, por si querés auditar.

**No repite pagos**: antes de armar el archivo mira qué hay en el Listado banco
y descarta lo que ya está cargado.

#### Si preferís que las pegue solo

```bash
# ... los mismos parametros, mas:
  --anexar
```

Escribe `Listado banco-2026 - con nuevas filas.xlsx` con las filas ya puestas al
final de `H1`, conservando el logo, los filtros y todo lo demás. Tarda alrededor
de un minuto porque el archivo es grande.

---

## Cómo leer el reporte de "No identificados"

Cada fila es un pago que entró al banco y no se pudo asignar. La columna
`Motivo` te dice por qué:

| Motivo | Qué significa y qué hacer |
|---|---|
| `ningun cliente coincide por contrato ni por nombre` | el pago no trae contrato o el contrato no está en la hoja del mes. Buscalo por el titular o el CUIT, que están en el reporte |
| `el contrato existe pero sus filas ya tienen cobro cargado` | el cliente ya tiene todas sus cuotas del mes cobradas. Puede ser un adelanto de la cuota del mes que viene |
| `mas de un cliente coincide por nombre` (hoja `Ambiguos`) | hay dos clientes con nombre parecido. El reporte lista los candidatos con su fila |

El reporte trae, para cada pago, la **hoja y la fila del libro de Ingresos
Bancarios** de donde salió, así lo podés ir a ver sin adivinar.

Cuando resuelvas uno a mano, cargalo en la planilla como siempre. La próxima
corrida no lo va a volver a tocar.

---

## Preguntas que van a aparecer

**¿Puede escribir sobre el archivo original en lugar de hacer una copia?**
Sí, agregando `--en-sitio`. Conviene usarlo recién cuando ya le tengas
confianza, y teniendo un respaldo.

**¿Se pierden los comentarios de las celdas?**
No. Se conservan los 1.119 comentarios, las fórmulas, los formatos y los
colores. El script reescribe únicamente las celdas que completa.

**¿Qué pasa si lo corro dos veces?**
Nada malo. No pisa una fila que ya tiene cobro cargado, así que la segunda
corrida solo agrega lo nuevo.

**¿Toca las filas dadas de baja?**
No. Las filas con `Estado` = 3 o con "baja" escrita en la fecha de pago quedan
intactas, y tampoco aparecen en la lista de morosos.

**¿Y si alguien corrigió algo a mano?**
Se respeta. Las correcciones manuales en `Estado` (los textos tipo "pago pero se
dio de baja") no se pisan.

**Me cambió el Excel de forma: agregaron una columna.**
Se arregla sin tocar código: en `conciliacion.json`, sección
`cobranza_general.columnas`, están las letras de cada columna. Cambiá la que
corresponda. Si una columna quedó corrida, el script te avisa antes de escribir
nada.

**El Estado quedó en 0 pero el cliente pagó.**
Mirá que `Monto de Pago` tenga un número y no texto. La regla es: sin monto o
monto 0 → `0`; monto menor a la cuota → `Parcial`; monto mayor o igual → `1`.

**¿Por qué algunas filas del Listado banco quedan marcadas en `Revisar`?**
Porque el código de cliente de ese contrato no está en el historial, o es el
segundo pago del mismo contrato en el mismo lote. El sufijo del código (el
último dígito) lo define el sistema destino y no se puede deducir de los
archivos; cuando lo confirmen, se carga una vez en `conciliacion.json` y dejan
de aparecer. Ver el detalle en [ARQUITECTURA.md](ARQUITECTURA.md#módulo-4--exportador).

---

## Si algo sale mal

| Mensaje | Qué pasó |
|---|---|
| `Permission denied` o `[Errno 13]` | el archivo está abierto en Excel. Cerralo |
| `El libro no tiene la hoja 'Noviembre'` | esa hoja todavía no existe en Cobranza General |
| `No encontre ninguna hoja mensual` | el archivo que pasaste no es Cobranza General |
| `la columna de fecha_pago deberia ser G pero ahi dice ...` | se corrió una columna: revisá `conciliacion.json` |
| `El libro usa el sistema de fechas 1904` | alguien cambió esa opción en Excel; hay que volverla a 1900 |

Nada de lo que corre acá borra datos: ante la duda, trabajá sobre la copia (sin
`--en-sitio`) y compará.

---

## Comandos sueltos que pueden servir

```bash
# Ver el extracto de todos los bancos unificado en una sola tabla
python3 -m clasificador.cli estandarizar "Ingresos Bancarios Clientes 2025-2026.xlsx"

# Actualizar solo la columna Estado, sin conciliar nada
python3 -m clasificador.cli estado --cobranza "Cobranza General 2026.xlsx"

# Probar como se clasifica una nota del extracto
python3 -m clasificador.cli probar "pagare" "cuota 1" "PF"

# Ver todos los comandos y sus opciones
python3 -m clasificador.cli --help
python3 -m clasificador.cli conciliar --help
```
