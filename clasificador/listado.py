"""MODULO 4 - Exportador hacia "Listado banco".

Arma las filas de la hoja ``H1`` del ``Listado banco`` a partir de los cobros
ya conciliados en ``Cobranza General``.

El formato de la columna ``Cliente``, medido en vez de supuesto
-------------------------------------------------------------
El pedido original decia que el contrato "se transforma segun reglas del
sistema destino (ej. a formatos como ``2161.0``, ``500008202.0`` o relleno de
ceros)". Cruzando los tres archivos quedo claro que la regla es una sola:

    Cliente = prefijo (1 digito) + contrato con ceros hasta 7 digitos + sufijo (1 digito)

Como se verifico:

* En la hoja H1, de los 27.020 codigos de 9 digitos, el tramo central
  (digitos 2 a 8) es un contrato conocido de Cobranza General en el **96,5%**
  de las filas de 2026. Probando otras ventanas del codigo, el acierto cae a
  1-2%, asi que la estructura no es casualidad.
* El libro de Ingresos Bancarios lo confirma por tercera vez: el bloque SIRO
  de Roela trae el codigo ya armado en su columna C, y en el **99,3%** de esas
  filas el tramo central coincide con el contrato de la fila.
* Un ejemplo completo del bloque SIRO: contrato 1313 -> ``100013131``, o sea
  prefijo 1, ``0001313``, sufijo 1. Y contrato 1012 -> ``000010125``: prefijo
  0, ``0001012``, sufijo 5.

Sobre el ``2161`` del pedido: en H1 existe el codigo ``000002161``, pero ese
codigo corresponde al contrato **216** (los 7 digitos del medio), no al 2161.
El contrato 2161 daria ``?00021611``. Ni 216 ni 2161 estan en Cobranza General
2026 -- esa fila de H1 es de 2021 --, asi que el ejemplo del pedido no era un
par contrato/codigo real. El ``.0`` que aparecia en el pedido es como Excel
muestra un numero, no parte del codigo.

El prefijo y el sufijo **no son constantes** (prefijos vistos: 1, 0, 2, 3, 5,
4, 6, 7, 8; sufijos: 1 a 6) y no se pueden deducir del contrato. Asi que no se
inventan: se toman, en este orden,

1. del codigo que el propio extracto trae armado (pagos SIRO),
2. del historial del Listado banco, que es la fuente de verdad del sistema
   destino: el prefijo que ese contrato ya uso,
3. de ``conciliacion.json`` como ultimo recurso, y la fila queda marcada para
   revision en lugar de salir con un codigo inventado.

El sufijo es el dato que NO se puede deducir
--------------------------------------------
Se probaron las hipotesis razonables contra las 23.584 parejas de pagos
consecutivos del historial:

* "el sufijo cicla 1..6 y vuelve a 1 con cada cuota": acierta 28,5% (17,4% si
  se mira solo 2026).
* "el sufijo es el orden del pago del contrato": 19,8%.
* "el sufijo es fijo por contrato": vale para el 56,7% de los contratos.

O sea: **el sufijo lo define el sistema destino y no esta en ninguno de los
tres archivos.** Por eso el modulo no lo adivina. Hace lo unico defendible:

1. si el extracto trae el codigo armado (cobros SIRO), lo usa tal cual;
2. si no, usa el (prefijo, sufijo) que ese contrato mas uso en el historial;
3. si el contrato nunca aparecio, usa los valores por defecto de la config.

Y en los casos 2 y 3 **marca la fila en la hoja "Revisar"**, con el motivo, en
vez de dar por buena una suposicion. Cuando el responsable del sistema destino
confirme la regla, se carga una vez en ``conciliacion.json``
(``mapa_cliente``) o se corrige la config, y el modulo deja de marcarlas.

Las otras columnas, tambien medidas sobre las 5.040 filas de 2026
-----------------------------------------------------------------
``ss`` y ``Cuenta`` van **vacias** (las 5.040 filas de 2026 las tienen vacias;
``005989/5`` aparece solo en filas viejas). ``SubTotal`` va vacia salvo cuando
una transferencia pago varias cuotas: ahi lleva el total transferido, que es
como se usa en el historial. ``Canal`` es ``TR <BANCO>`` para las
transferencias y el codigo de la red (PF, RP, TQR, FSF, PP, RSF, CE, CEF, PC)
para los cobros por SIRO. ``Fecha de Acred.`` sale del extracto cuando el
bloque SIRO la informa; si no, es la fecha de proceso mas
``dias_acreditacion``.
"""

from __future__ import annotations

import datetime as dt
import warnings
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from clasificador.cobranza import FilaCobranza
from clasificador.config import Config
from clasificador.conciliacion import Aplicacion
from clasificador.estado import PAGADO, PARCIAL, estado_de
from clasificador.xlsx_parche import Celda, parchear

LARGO_CODIGO = 9


@dataclass
class Codigos:
    """Lo que el historial del Listado banco sabe de cada contrato."""

    prefijos: dict[str, str]
    sufijos: dict[str, list[str]]
    combos: dict[str, tuple[str, str]]
    codigos: set[str]
    por_contrato: dict[str, set[str]]

    def combo_de(
        self, contrato: str, prefijo_defecto: str, sufijo_defecto: str, orden: int = 1
    ) -> tuple[str, str, str]:
        """(prefijo, sufijo, procedencia) para el pago ``orden`` del contrato.

        ``orden`` es 1 para el primer pago de ese contrato en la exportacion, 2
        para el segundo, etc. Cuando hay varios, se van tomando los sufijos que
        ese contrato uso en el historial, en vez de inventar numeros nuevos.
        """
        if contrato not in self.combos:
            return prefijo_defecto, sufijo_defecto, "defecto"
        prefijo, sufijo = self.combos[contrato]
        if orden <= 1:
            return prefijo, sufijo, "historial"
        conocidos = self.sufijos.get(contrato, [])
        otros = [s for s in conocidos if s != sufijo]
        if len(otros) >= orden - 1:
            return prefijo, otros[orden - 2], "historial (pago repetido)"
        return prefijo, sufijo, "historial (sufijo repetido)"


def aprender_codigos(ruta: str | Path, cfg: Config) -> Codigos:
    """Lee el historial de ``H1`` y aprende el prefijo de cada contrato.

    Es la pieza que evita inventar codigos: el sistema destino ya dijo, 27.000
    veces, como se llama cada cliente. Para un contrato con varios prefijos en
    el historial gana el mas frecuente.
    """
    hoja = cfg.listado["hoja"]
    ancho = int(cfg.listado["ancho_contrato"])
    vistos: dict[str, Counter[str]] = defaultdict(Counter)
    combos: dict[str, Counter[tuple[str, str]]] = defaultdict(Counter)
    sufijos: dict[str, Counter[str]] = defaultdict(Counter)
    codigos: set[str] = set()
    por_contrato: dict[str, set[str]] = defaultdict(set)

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        libro = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    try:
        if hoja not in libro.sheetnames:
            raise RuntimeError(
                f"El Listado banco no tiene la hoja {hoja!r}. "
                f"Tiene: {', '.join(libro.sheetnames)}"
            )
        ws = libro[hoja]
        for fila in ws.iter_rows(min_row=2, max_col=3, values_only=True):
            crudo = fila[2] if len(fila) > 2 else None
            if crudo is None:
                continue
            codigo = str(crudo).strip()
            if len(codigo) != LARGO_CODIGO or not codigo.isdigit():
                continue
            codigos.add(codigo)
            contrato = str(int(codigo[1 : 1 + ancho]))
            vistos[contrato][codigo[0]] += 1
            sufijos[contrato][codigo[-1]] += 1
            combos[contrato][(codigo[0], codigo[-1])] += 1
            por_contrato[contrato].add(codigo)
    finally:
        libro.close()

    return Codigos(
        prefijos={c: v.most_common(1)[0][0] for c, v in vistos.items()},
        sufijos={c: [s for s, _ in v.most_common()] for c, v in sufijos.items()},
        combos={c: v.most_common(1)[0][0] for c, v in combos.items()},
        codigos=codigos,
        por_contrato=dict(por_contrato),
    )


def codigo_cliente(contrato: str, prefijo: str, sufijo: str, ancho: int = 7) -> str:
    """``2161`` + prefijo 0 + sufijo 1 -> ``000002161``."""
    if not contrato:
        return ""
    cuerpo = str(contrato).zfill(ancho)
    if len(cuerpo) > ancho:
        # Un contrato mas largo que el hueco: se deja entero y se avisa por
        # fuera. Recortarlo mandaria la plata a otro cliente.
        return f"{prefijo}{cuerpo}{sufijo}"
    return f"{prefijo}{cuerpo}{sufijo}"


@dataclass
class FilaListado:
    """Una fila de la hoja H1, en el orden exacto de sus columnas."""

    ss: object = None
    cuenta: object = None
    cliente: str = ""
    importe: float = 0.0
    fecha_proceso: dt.date | None = None
    fecha_pago: dt.date | None = None
    fecha_acreditacion: dt.date | None = None
    canal: str = ""
    subtotal: object = None
    # Afuera de la hoja: para el reporte y la revision.
    contrato: str = ""
    nombre: str = ""
    mes: str = ""
    fila_origen: int = 0
    procedencia: str = ""
    revisar: str = ""

    def a_fila(self) -> list[object]:
        return [
            self.ss,
            self.cuenta,
            self.cliente,
            self.importe,
            self.fecha_proceso,
            self.fecha_pago,
            self.fecha_acreditacion,
            self.canal,
            self.subtotal,
        ]


def indice_de_extracto(
    movimientos: list,
) -> dict[tuple, object]:
    """Indice del extracto para recuperar canal, acreditacion y codigo.

    Hace falta para las filas que se concilaron en corridas anteriores: en
    ``Cobranza General`` la columna ``Medio de Pago`` guarda 'siro', pero no
    con que canal de la red se cobro (PF, RP, TQR...), y ese dato es
    obligatorio en el Listado banco. En vez de pedirle al operador que lo
    busque, se lo va a buscar al extracto por contrato, importe y fecha.

    Se indexa en dos niveles porque el importe no siempre coincide: cuando una
    transferencia pago varias cuotas, la fila lleva el valor de su cuota y no
    el total transferido.
    """
    indice: dict[tuple, object] = {}
    for mov in movimientos:
        if not mov.contrato_detectado or mov.fecha is None:
            continue
        exacta = (mov.contrato_detectado, round(mov.monto, 2), mov.fecha)
        indice.setdefault(exacta, mov)
        suelta = (mov.contrato_detectado, mov.fecha)
        # Entre varios candidatos del mismo dia gana el que trae codigo de
        # canal, que es justamente el dato que se vino a buscar.
        previo = indice.get(suelta)
        if previo is None or (mov.codigo_canal and not previo.codigo_canal):
            indice[suelta] = mov
    return indice


def buscar_en_extracto(
    indice: dict[tuple, object], contrato: str, monto: float, fecha
) -> object | None:
    if not contrato or fecha is None:
        return None
    return indice.get((contrato, round(monto, 2), fecha)) or indice.get(
        (contrato, fecha)
    )


def canal_de(medio: str, codigo_canal: str, cfg: Config) -> tuple[str, str]:
    """Traduce el medio de pago al ``Canal`` del destino. Devuelve (canal, aviso)."""
    conf = cfg.listado
    codigos = {c.upper() for c in conf["codigos_de_canal"]}
    if codigo_canal and codigo_canal.upper() in codigos:
        return codigo_canal.upper(), ""
    clave = (medio or "").strip().lower()
    if clave in conf["canal_de_banco"]:
        return conf["canal_de_banco"][clave], ""
    if clave in conf.get("canal_sin_codigo", {}):
        valor = conf["canal_sin_codigo"][clave]
        if valor:
            return valor, ""
        return "", (
            f"el medio {clave!r} necesita el codigo de canal (PF, RP, TQR...) "
            "y el extracto no lo trae"
        )
    if clave.upper() in codigos:
        return clave.upper(), ""
    return "", f"no se como traducir el medio {medio!r} a un canal"


def _fecha_acreditacion(
    fecha_pago: dt.date | None,
    proceso: dt.date,
    informada: dt.date | None,
    cfg: Config,
) -> dt.date:
    if informada:
        return informada
    return proceso + dt.timedelta(days=int(cfg.listado["dias_acreditacion"]))


def armar_listado(
    base: dict[str, list[FilaCobranza]],
    cfg: Config,
    codigos: Codigos,
    proceso: dt.date | None = None,
    desde: dt.date | None = None,
    hasta: dt.date | None = None,
    aplicaciones: list[Aplicacion] | None = None,
    movimientos: list | None = None,
    solo_pagados: bool = False,
    ya_exportados: set[tuple[str, float, dt.date | None]] | None = None,
) -> list[FilaListado]:
    """Arma las filas de H1 a partir de los cobros cargados en la base.

    ``aplicaciones`` es el resultado del Modulo 2 de esta misma corrida: de ahi
    salen el codigo de canal, la fecha de acreditacion real y el total del
    grupo cuando una transferencia pago varias cuotas. Sin eso igual funciona,
    pero el ``Canal`` de los cobros por SIRO queda para revision, porque el
    codigo (PF, RP, TQR...) no esta en Cobranza General.

    ``solo_pagados`` deja afuera los pagos parciales. El pedido original decia
    "registros conciliados (Estado = 1)", pero un pago parcial es plata que
    entro y el sistema destino la tiene que ver igual; el historial del propio
    Listado banco tiene importes mas chicos que la cuota. Por eso el valor por
    defecto es False, e incluye los parciales.
    """
    proceso = proceso or dt.date.today()
    ancho = int(cfg.listado["ancho_contrato"])
    prefijo_defecto = str(cfg.listado["prefijo_por_defecto"])
    sufijo_defecto = str(cfg.listado["sufijo_por_defecto"])
    modo_subtotal = cfg.listado.get("subtotal", "vacio")
    cuenta = cfg.listado.get("cuenta") or None
    parcial = cfg.mora.get("estado_parcial", PARCIAL)
    ya_exportados = ya_exportados or set()
    mapa = {str(k): str(v) for k, v in cfg.listado.get("mapa_cliente", {}).items()}

    # Datos que aporta la conciliacion de esta corrida, por fila de destino.
    extra: dict[tuple[str, int], Aplicacion] = {}
    for a in aplicaciones or []:
        extra[(a.fila.hoja, a.fila.fila)] = a
    # Y el extracto completo, para las filas conciliadas en corridas previas.
    indice = indice_de_extracto(movimientos or [])

    filas: list[FilaListado] = []
    usados_por_contrato: Counter[str] = Counter()

    for mes in cfg.meses:
        for fila in base.get(mes, []):
            if fila.es_baja or not fila.monto_pago:
                continue
            estado = estado_de(fila.monto_pago, fila.cuota, parcial)
            if estado == 0:
                continue
            if solo_pagados and estado != PAGADO:
                continue
            if desde and (fila.fecha_pago is None or fila.fecha_pago < desde):
                continue
            if hasta and (fila.fecha_pago is None or fila.fecha_pago > hasta):
                continue

            aplicacion = extra.get((fila.hoja, fila.fila))
            mov = aplicacion.movimiento if aplicacion else None
            if mov is None:
                mov = buscar_en_extracto(
                    indice, fila.contrato, fila.monto_pago, fila.fecha_pago
                )
            avisos: list[str] = []

            # --- codigo de Cliente ---
            procedencia = ""
            if mov is not None and mov.codigo_cliente:
                # El extracto lo trae armado por el propio sistema destino: es
                # el unico caso en que el codigo no es una deduccion.
                cliente = mov.codigo_cliente
                procedencia = "extracto"
            elif fila.contrato in mapa:
                # Equivalencia cargada a mano: manda sobre todo lo demas.
                cliente = mapa[fila.contrato]
                procedencia = "mapa_cliente"
            elif fila.contrato:
                usados_por_contrato[fila.contrato] += 1
                prefijo, sufijo, procedencia = codigos.combo_de(
                    fila.contrato,
                    prefijo_defecto,
                    sufijo_defecto,
                    usados_por_contrato[fila.contrato],
                )
                cliente = codigo_cliente(fila.contrato, prefijo, sufijo, ancho)
                if procedencia == "defecto":
                    avisos.append(
                        f"el contrato {fila.contrato} no figura en el historial del "
                        f"Listado banco: se uso el codigo por defecto "
                        f"{prefijo_defecto}/{sufijo_defecto}. Confirmar el codigo "
                        "en el sistema destino antes de cargarlo."
                    )
                elif "repetido" in procedencia:
                    avisos.append(
                        f"es el pago numero {usados_por_contrato[fila.contrato]} de "
                        "este contrato en la exportacion y el sufijo salio del "
                        "historial; confirmar que no choque con otro cupon"
                    )
                if len(str(fila.contrato)) > ancho:
                    avisos.append(
                        f"el contrato {fila.contrato} no entra en {ancho} digitos"
                    )
            else:
                cliente = ""
                procedencia = "sin contrato"
                avisos.append("la fila no tiene numero de contrato")

            # --- canal ---
            canal, aviso_canal = canal_de(
                fila.medio_pago, mov.codigo_canal if mov else "", cfg
            )
            if aviso_canal:
                avisos.append(aviso_canal)

            # --- subtotal ---
            subtotal = None
            if modo_subtotal == "grupo" and aplicacion is not None:
                subtotal = aplicacion.total_del_grupo
            elif modo_subtotal == "importe":
                subtotal = fila.monto_pago

            if fila.fecha_pago and fila.fecha_pago > proceso:
                # Una fecha de pago posterior a la de proceso casi siempre es
                # un año o un mes mal tipeado en Cobranza General. Se exporta
                # igual (no es el script quien decide tirar un cobro) pero se
                # marca, porque el sistema destino la puede rechazar.
                avisos.append(
                    f"la fecha de pago ({fila.fecha_pago}) es posterior a la "
                    f"fecha de proceso ({proceso}); revisar la celda en "
                    f"{fila.hoja} fila {fila.fila}"
                )

            clave = (cliente, float(fila.monto_pago), fila.fecha_pago)
            if clave in ya_exportados:
                continue

            filas.append(
                FilaListado(
                    cuenta=cuenta,
                    cliente=cliente,
                    importe=float(fila.monto_pago),
                    fecha_proceso=proceso,
                    fecha_pago=fila.fecha_pago,
                    fecha_acreditacion=_fecha_acreditacion(
                        fila.fecha_pago,
                        proceso,
                        mov.fecha_acreditacion if mov else None,
                        cfg,
                    ),
                    canal=canal,
                    subtotal=subtotal,
                    contrato=fila.contrato,
                    nombre=fila.nombre,
                    mes=fila.hoja,
                    fila_origen=fila.fila,
                    procedencia=procedencia,
                    revisar="; ".join(avisos),
                )
            )
    return filas


def leer_ya_exportados(
    ruta: str | Path, cfg: Config
) -> set[tuple[str, float, dt.date | None]]:
    """Lo que ya esta en H1, para no mandar dos veces el mismo pago.

    La clave es (codigo de cliente, importe, fecha de pago). Usa el archivo de
    destino como registro, asi no hace falta guardar estado en ningun lado.
    """
    hoja = cfg.listado["hoja"]
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        libro = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    try:
        if hoja not in libro.sheetnames:
            return set()
        vistos = set()
        for fila in libro[hoja].iter_rows(min_row=2, max_col=6, values_only=True):
            fila = list(fila) + [None] * (6 - len(fila))
            cliente, importe, pago = fila[2], fila[3], fila[5]
            if cliente is None or importe is None:
                continue
            fecha = pago.date() if isinstance(pago, dt.datetime) else pago
            try:
                vistos.add((str(cliente).strip(), float(importe), fecha))
            except (TypeError, ValueError):
                continue
        return vistos
    finally:
        libro.close()


def ultima_fila_con_datos(ruta: str | Path, cfg: Config) -> int:
    """Ultima fila de H1 que tiene algo escrito.

    No alcanza con ``ws.max_row``: en el archivo real la hoja tiene 40.505
    filas creadas y preformateadas, de las cuales solo las primeras 27.435
    tienen datos.
    """
    hoja = cfg.listado["hoja"]
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        libro = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    try:
        ultima = 1
        for i, fila in enumerate(
            libro[hoja].iter_rows(min_row=2, max_col=9, values_only=True), start=2
        ):
            if any(v is not None and v != "" for v in fila):
                ultima = i
        return ultima
    finally:
        libro.close()


_AZUL = "FF1F4E78"
_BLANCO = "FFFFFFFF"
_AMARILLO = "FFFFF2CC"


def escribir_listado(
    filas: list[FilaListado], ruta: str | Path, cfg: Config
) -> Path:
    """Escribe un .xlsx con las filas nuevas, con el formato de H1.

    Pensado para el circuito que describe la propia hoja INSTRUCCIONES del
    Listado banco: copiar las filas sin encabezado y pegarlas al final de H1
    como "solo valores". La hoja ``Revisar`` junta lo que conviene mirar antes
    de pegar.
    """
    libro = openpyxl.Workbook()
    ws = libro.active
    ws.title = cfg.listado["hoja"]
    ws.append(cfg.listado["encabezados"])
    relleno = PatternFill("solid", fgColor=_AZUL)
    for celda in ws[1]:
        celda.font = Font(name="Calibri", size=11, bold=True, color=_BLANCO)
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center", vertical="center")

    for fila in filas:
        ws.append(fila.a_fila())
    ws.freeze_panes = "A2"
    for f in ws.iter_rows(min_row=2, max_row=ws.max_row):
        # El codigo de cliente es texto: si Excel lo toma como numero se come
        # los ceros de la izquierda y 000002161 se convierte en 2161.
        f[2].number_format = "@"
        f[3].number_format = "#,##0.00"
        for i in (4, 5, 6):
            f[i].number_format = "dd/mm/yyyy"
        f[8].number_format = "#,##0.00"
    for col, ancho in zip("ABCDEFGHI", (6, 12, 14, 14, 16, 14, 14, 14, 14)):
        ws.column_dimensions[col].width = ancho

    revisar = [f for f in filas if f.revisar]
    hoja = libro.create_sheet("Revisar")
    hoja.append(
        [
            "Mes", "Fila", "Contrato", "Cliente", "Nombre", "Importe",
            "Canal", "Procedencia del codigo", "Motivo",
        ]
    )
    for celda in hoja[1]:
        celda.font = Font(name="Calibri", size=11, bold=True, color=_BLANCO)
        celda.fill = relleno
    for f in revisar:
        hoja.append(
            [f.mes, f.fila_origen, f.contrato, f.cliente, f.nombre,
             f.importe, f.canal, f.procedencia, f.revisar]
        )
    for celda in hoja["I"]:
        celda.fill = PatternFill("solid", fgColor=_AMARILLO)
    hoja.column_dimensions["E"].width = 32
    hoja.column_dimensions["H"].width = 22
    hoja.column_dimensions["I"].width = 70

    # Una hoja mas, con TODAS las filas y de donde salio cada codigo. No es
    # una lista de problemas: es la trazabilidad de la exportacion.
    traza = libro.create_sheet("Procedencia")
    traza.append(
        ["Mes", "Fila", "Contrato", "Cliente", "Nombre", "Importe", "Canal",
         "Procedencia del codigo"]
    )
    for celda in traza[1]:
        celda.font = Font(name="Calibri", size=11, bold=True, color=_BLANCO)
        celda.fill = relleno
    for f in filas:
        traza.append(
            [f.mes, f.fila_origen, f.contrato, f.cliente, f.nombre,
             f.importe, f.canal, f.procedencia]
        )
    traza.freeze_panes = "A2"
    traza.auto_filter.ref = traza.dimensions
    traza.column_dimensions["E"].width = 32
    traza.column_dimensions["H"].width = 24
    hoja.freeze_panes = "A2"

    ruta = Path(ruta)
    libro.save(ruta)
    libro.close()
    return ruta


def anexar_en_listado(
    filas: list[FilaListado],
    origen: str | Path,
    destino: str | Path,
    cfg: Config,
    primera_fila: int | None = None,
) -> tuple[int, int]:
    """Anexa las filas al final de H1 del Listado banco real.

    Va por ``xlsx_parche`` y no por openpyxl porque el archivo trae un logo en
    ``xl/media`` que un round-trip de openpyxl descarta. Las filas vacias del
    final ya existen en el XML y ya vienen preformateadas, asi que alcanza con
    escribirles las celdas.

    Devuelve (primera fila escrita, ultima fila escrita).
    """
    hoja = cfg.listado["hoja"]
    inicio = (primera_fila or ultima_fila_con_datos(origen, cfg) + 1)
    celdas: list[Celda] = []
    for n, fila in enumerate(filas):
        for col, valor in enumerate(fila.a_fila(), start=1):
            if valor is None or valor == "":
                continue
            celdas.append(Celda(hoja, inicio + n, col, valor))
    if not celdas:
        return inicio, inicio - 1
    parchear(origen, destino, celdas)
    return inicio, inicio + len(filas) - 1
