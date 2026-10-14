"""MODULO 1 - Unificador y estandarizador de extractos bancarios.

Toma el libro "Ingresos Bancarios Clientes" (105 hojas, una por banco y mes,
con cuatro layouts distintos) y devuelve una sola tabla con las columnas
pedidas:

    Fecha | Monto | Banco/Medio | Titular_Origen | CUIT_CVU |
    Contrato_Detectado | Concepto

mas unas columnas de trazabilidad (``Codigo_Canal``, ``Detalle``, ``Hoja``,
``Fila``, ``Clave``) que no estaban en el pedido pero sin las cuales no se
puede auditar un pago ni armar el ``Canal`` del Listado banco.

Por que no reescribe la lectura del libro: ``origen.py`` ya resuelve la parte
dificil (detectar por hoja donde esta cada dato, con el bloque extra de Roela
declarado aparte) y esta contrastada contra el extracto real: 11.162
movimientos leidos, 98,1% de las filas ya clasificadas a mano reencontradas.
Ese codigo no se toca; aca se lo reusa y se le agrega lo que descartaba
(contrato, nombre del titular, el medio de pago y el bloque SIRO).

El bloque izquierdo de las hojas de Roela
-----------------------------------------
``origen.py`` lo llama "cheques" porque la columna C parecia un numero de
cheque. Leyendo el encabezado real de la hoja resulta que no: dice
``CUPON PAGO FACIL-RAPIPAGO``. Es el bloque de cobranzas de la red SIRO, y sus
columnas valen oro para los Modulos 2 y 4:

    A  fecha de carga        E  fecha de acreditacion
    B  importe               F  codigo de canal (PF, TQR, RP, FSF, PP...)
    C  codigo de cliente     G  numero de contrato
    D  fecha de pago         H  nombre del cliente

La columna C trae el codigo del sistema destino ya armado: para el contrato
1313 dice 100013131, y para el 1012 dice 10125 (o sea 000010125). Es la misma
estructura ``prefijo + contrato en 7 digitos + sufijo`` que se midio en el
Listado banco, confirmada desde un tercer archivo. Cuando ese codigo esta, el
Modulo 4 lo usa tal cual en vez de deducirlo.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from clasificador.origen import (
    FILAS_MUESTRA,
    ROELA_CHEQUES,
    Bloque,
    _celda,
    _es_fecha,
    _texto,
    banco_de_hoja,
    detectar_bloques,
)
from clasificador.patrones import extraer_contrato, extraer_cuit_cvu, extraer_titular
from clasificador.reglas import normalizar

ENCABEZADOS_EXTRACTO = [
    "Fecha",
    "Monto",
    "Banco/Medio",
    "Titular_Origen",
    "CUIT_CVU",
    "Contrato_Detectado",
    "Concepto",
    "Medio",
    "Codigo_Canal",
    "Codigo_Cliente",
    "Fecha_Acred",
    "Detalle",
    "Hoja",
    "Fila",
    "Clave",
]

HOJA_EXTRACTO = "Extracto"

# Medios de pago normalizados.
TRANSFERENCIA = "Transferencia"
CUPON = "Cupon SIRO"
CHEQUE = "Cheque"
DEPOSITO = "Deposito"
EFECTIVO = "Efectivo"

# Como se llama ese medio en la columna "Medio de Pago" de Cobranza General.
# Son los valores que el equipo ya viene usando: 'rio', 'galicia', 'siro'...
MEDIO_SIRO = "siro"

# Posiciones (0-based) de las columnas extra del bloque SIRO de Roela, las que
# origen.py no mira porque para clasificar no las necesita.
SIRO_CODIGO_CLIENTE = 2
SIRO_FECHA_ACREDITACION = 4

_PISTAS_MEDIO = (
    # "credin" NO entra aca: CREDIN es Credito Inmediato, una
    # transferencia. El bloque de cheques de Roela se reconoce por su
    # posicion en la hoja, no por el texto, asi que no hace falta.
    (CHEQUE, ("cheque", "ech ", "e/ch")),
    (DEPOSITO, ("deposito", "depos", "dep.")),
    (EFECTIVO, ("efectivo",)),
)


def _medio_de_pago(concepto: str, detalle: str, es_siro: bool) -> str:
    """Transferencia salvo que algo diga lo contrario.

    El bloque SIRO de Roela se reconoce por su posicion en la hoja, no por el
    texto: ahi no hay concepto del banco. En el resto, el medio se deduce del
    concepto; en la practica casi todo es transferencia.
    """
    if es_siro:
        return CUPON
    texto = normalizar(f"{concepto} {detalle}")
    for medio, pistas in _PISTAS_MEDIO:
        if any(p in texto for p in pistas):
            return medio
    return TRANSFERENCIA


def _codigo_canal(detalle: str, codigos: frozenset[str]) -> str:
    """Rescata el codigo de canal (PF, RP, TQR, FSF...) de la nota manual.

    Hallazgo de contrastar los tres archivos: esas notas de ``Detalle`` que
    parecian comentarios sueltos son los codigos de ``Canal`` del Listado
    banco. En el libro real: PF 1894, TQR 470, RP 347, FSF 202, PP 65, RSF 61,
    PC 15, CE 14, CEF 14. Sin esto, los pagos de la red SIRO (1194 en el año)
    no se pueden exportar con su canal correcto.
    """
    if not detalle:
        return ""
    for palabra in re.split(r"[^A-Za-z]+", detalle.upper()):
        if palabra in codigos:
            return palabra
    return ""


# Meses como los escribe el equipo en los nombres de hoja ("Macro Feb 26",
# "Galicia oct", "Patag Dic"). Se compara por las primeras tres letras, que es
# lo unico estable entre "Feb", "Febrero" y "feb 26".
MESES_DEL_ANIO = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
_ABREVIATURAS = {"set": "Septiembre", "sept": "Septiembre"}


def mes_de_hoja(titulo: str) -> str:
    """Saca el mes del nombre de la hoja ('Macro Feb 26' -> 'Febrero').

    Importa mas de lo que parece: un cupon pagado el 30 de diciembre entra en
    la hoja "Roela Enero" porque paga la cuota de enero. Si la conciliacion
    usara la fecha del pago para elegir la hoja destino, ese cobro iria a
    diciembre y no se encontraria nunca. El mes que vale es el que el equipo ya
    eligio al archivar el movimiento.
    """
    limpio = normalizar(titulo)
    banco = banco_de_hoja(titulo)
    if banco:
        # El nombre de la hoja arranca con el banco; se lo saca para que
        # "Mar" de "Macro Marzo" no se confunda con el banco mismo.
        for alias in (banco, "patag", "pata"):
            if limpio.startswith(normalizar(alias)):
                limpio = limpio[len(normalizar(alias)):].strip()
                break
    for palabra in limpio.split():
        if palabra in _ABREVIATURAS:
            return _ABREVIATURAS[palabra]
        for mes in MESES_DEL_ANIO:
            if len(palabra) >= 3 and normalizar(mes).startswith(palabra[:4].rstrip()):
                if normalizar(mes).startswith(palabra[:3]):
                    return mes
    return ""


@dataclass
class MovimientoBancario:
    """Un cobro del extracto, ya normalizado."""

    fecha: dt.date | None
    monto: float
    banco: str
    titular_origen: str
    cuit_cvu: str
    contrato_detectado: str
    concepto: str
    medio: str = TRANSFERENCIA
    codigo_canal: str = ""
    codigo_cliente: str = ""
    fecha_acreditacion: dt.date | None = None
    detalle: str = ""
    origen_contrato: str = ""
    hoja: str = ""
    mes_origen: str = ""
    anio_origen: int = 0
    fila: int = 0
    clave: str = field(default="", init=False)

    def __post_init__(self) -> None:
        self.clave = self.calcular_clave()

    def calcular_clave(self) -> str:
        """Huella del movimiento segun lo que informa el banco.

        No entran ni ``detalle`` ni ``codigo_canal``: son anotaciones nuestras,
        y si alguien corrige la nota el movimiento tiene que seguir siendo el
        mismo para que la conciliacion no lo tome por uno nuevo. Mismo criterio
        que ``excel.Movimiento.calcular_clave``.
        """
        partes = [
            self.fecha.isoformat() if self.fecha else "",
            self.banco.strip().lower(),
            f"{self.monto:.2f}",
            " ".join(self.concepto.split()).lower(),
            self.contrato_detectado,
        ]
        return hashlib.sha1("|".join(partes).encode("utf-8")).hexdigest()[:16]

    @property
    def medio_cobranza(self) -> str:
        """Como se escribe este medio en la columna I de Cobranza General.

        Un cobro por cupon es 'siro' para el equipo, no 'Roela': la red SIRO
        cobra a traves de Banco Roela, pero en la base se anota la red. Los
        numeros lo confirman: en 2026 Cobranza General tiene 1194 'siro' y el
        Listado banco 1144 filas con codigos de canal (PF, TQR, RP, FSF...).
        """
        if self.medio == CUPON or self.codigo_canal:
            return MEDIO_SIRO
        return self.banco.lower()

    def a_fila(self) -> list[object]:
        return [
            self.fecha,
            self.monto,
            self.banco,
            self.titular_origen,
            self.cuit_cvu,
            self.contrato_detectado,
            self.concepto,
            self.medio,
            self.codigo_canal,
            self.codigo_cliente,
            self.fecha_acreditacion,
            self.detalle,
            self.hoja,
            self.fila,
            self.clave,
        ]


def _es_bloque_siro(bloque: Bloque, banco: str) -> bool:
    """El bloque de cupones de Roela, el unico que no sigue la adyacencia."""
    return (
        banco == "Roela"
        and bloque.fecha == ROELA_CHEQUES["fecha"]
        and bloque.importe == ROELA_CHEQUES["importe"]
    )


def _contrato_de_columna(valor: object, maximo: int) -> str:
    """Normaliza el contrato de su columna: 3499.0, '3499' y 3499 son lo mismo."""
    if valor is None:
        return ""
    if isinstance(valor, bool):
        return ""
    if isinstance(valor, (int, float)):
        if float(valor) != int(valor) or not 0 < int(valor) <= maximo:
            return ""
        return str(int(valor))
    texto = str(valor).strip()
    if texto.isdigit() and 0 < int(texto) <= maximo:
        return str(int(texto))
    return ""


def _completar_anio(movimientos: list[MovimientoBancario]) -> None:
    """Le pone año a cada hoja del libro, mirando las fechas que tiene adentro.

    Hace falta porque el libro de origen cubre dos años ("Ingresos Bancarios
    Clientes 2025-2026") y los nombres de hoja no son parejos: "Rio Enero" es
    enero de 2025 y "Rio Ene 26" es enero de 2026. Sin el año, la conciliacion
    cruza un cobro de enero de 2025 contra la cuota de enero de 2026 -- el
    monto suele coincidir y el error pasa inadvertido.

    En vez de interpretar el nombre, se usa el año que mas repiten las fechas
    de la propia hoja, contando solo las del mes que anuncia el titulo. Asi
    "Roela Enero", que arranca con cupones pagados el 30/12, igual queda en
    2025: la mayoria de sus fechas de enero son de 2025.
    """
    por_hoja: dict[str, Counter[int]] = {}
    respaldo: dict[str, Counter[int]] = {}
    for mov in movimientos:
        if not mov.fecha:
            continue
        respaldo.setdefault(mov.hoja, Counter())[mov.fecha.year] += 1
        if mov.mes_origen and MESES_DEL_ANIO[mov.fecha.month - 1] == mov.mes_origen:
            por_hoja.setdefault(mov.hoja, Counter())[mov.fecha.year] += 1
    for mov in movimientos:
        conteo = por_hoja.get(mov.hoja) or respaldo.get(mov.hoja)
        if conteo:
            mov.anio_origen = conteo.most_common(1)[0][0]
        elif mov.fecha:
            mov.anio_origen = mov.fecha.year


def leer_extracto(
    ruta: str | Path,
    codigos_de_canal: list[str] | tuple[str, ...] = (),
    contrato_maximo: int = 999999,
) -> tuple[list[MovimientoBancario], list[str]]:
    """Lee el libro de Ingresos Bancarios y devuelve (movimientos, avisos)."""
    codigos = frozenset(c.upper() for c in codigos_de_canal)
    libro = openpyxl.load_workbook(ruta, data_only=True)
    movimientos: list[MovimientoBancario] = []
    avisos: list[str] = []
    try:
        for ws in libro.worksheets:
            banco = banco_de_hoja(ws.title)
            if banco is None:
                avisos.append(f"Hoja {ws.title!r}: no reconozco el banco, la salteo.")
                continue

            filas = [
                (i, f)
                for i, f in enumerate(ws.iter_rows(values_only=True), start=1)
                if any(v is not None for v in f)
            ]
            if not filas:
                continue
            ancho = max(len(f) for _, f in filas)
            bloques = detectar_bloques(
                [f for _, f in filas[:FILAS_MUESTRA]], ancho, banco
            )
            if not bloques:
                avisos.append(
                    f"Hoja {ws.title!r}: no encontre columna de fecha, la salteo."
                )
                continue

            for nro_fila, fila in filas:
                for bloque in bloques:
                    fecha = _celda(fila, bloque.fecha)
                    monto = _celda(fila, bloque.importe)
                    if not _es_fecha(fecha) or not isinstance(monto, (int, float)):
                        continue
                    if isinstance(monto, bool):
                        continue

                    partes = [_texto(_celda(fila, c)) for c in bloque.extracto]
                    concepto = " ".join(p for p in partes if p and not p.isdigit())
                    nombre = _texto(_celda(fila, bloque.nombre))
                    detalle = _texto(_celda(fila, bloque.detalle))
                    if not concepto:
                        # Los cheques de Roela no traen texto del banco; el
                        # extracto original usa el nombre del cliente.
                        concepto = nombre

                    es_siro = _es_bloque_siro(bloque, banco)
                    codigo_cliente = ""
                    fecha_acred = None
                    if es_siro:
                        # El codigo del sistema destino viene ya armado, pero
                        # como numero: 10125 es en realidad '000010125'.
                        crudo = _celda(fila, SIRO_CODIGO_CLIENTE)
                        if isinstance(crudo, (int, float)) and not isinstance(crudo, bool):
                            if float(crudo) == int(crudo) and 0 < int(crudo) < 10**9:
                                codigo_cliente = str(int(crudo)).zfill(9)
                        elif isinstance(crudo, str) and crudo.strip().isdigit():
                            codigo_cliente = crudo.strip().zfill(9)
                        acred = _celda(fila, SIRO_FECHA_ACREDITACION)
                        if _es_fecha(acred):
                            fecha_acred = acred.date()

                    contrato = _contrato_de_columna(
                        _celda(fila, bloque.contrato), contrato_maximo
                    )
                    procedencia = "columna" if contrato else ""
                    if not contrato:
                        contrato, procedencia = extraer_contrato(
                            f"{detalle} {concepto}", contrato_maximo
                        )

                    movimientos.append(
                        MovimientoBancario(
                            fecha=fecha.date(),
                            monto=float(monto),
                            banco=banco,
                            # El nombre de la columna es el dato bueno; el
                            # titular sacado del concepto es el respaldo.
                            titular_origen=nombre or extraer_titular(concepto),
                            cuit_cvu=extraer_cuit_cvu(concepto),
                            contrato_detectado=contrato,
                            concepto=concepto,
                            medio=_medio_de_pago(concepto, detalle, es_siro),
                            codigo_canal=_codigo_canal(detalle, codigos),
                            codigo_cliente=codigo_cliente,
                            fecha_acreditacion=fecha_acred,
                            detalle=detalle,
                            origen_contrato=procedencia,
                            hoja=ws.title,
                            mes_origen=mes_de_hoja(ws.title),
                            fila=nro_fila,
                        )
                    )
        _completar_anio(movimientos)
        return movimientos, avisos
    finally:
        libro.close()


def claves_con_ocurrencia(movimientos: list[MovimientoBancario]) -> list[str]:
    """Numera los movimientos identicos para poder consumirlos de a uno.

    Dos cobros del mismo dia, banco, monto y concepto son legitimos (pasa
    seguido con dos cuotas iguales del mismo cliente). Numerarlos permite
    aplicar el primero y dejar el segundo disponible, en vez de descartarlo.
    """
    vistos: Counter[str] = Counter()
    claves = []
    for mov in movimientos:
        vistos[mov.clave] += 1
        claves.append(f"{mov.clave}#{vistos[mov.clave]}")
    return claves


def resumen(movimientos: list[MovimientoBancario]) -> str:
    """Una linea por banco, para ver de un vistazo que se leyo."""
    por_banco: Counter[str] = Counter(m.banco for m in movimientos)
    con_contrato = sum(1 for m in movimientos if m.contrato_detectado)
    con_cuit = sum(1 for m in movimientos if m.cuit_cvu)
    con_canal = sum(1 for m in movimientos if m.codigo_canal)
    con_codigo = sum(1 for m in movimientos if m.codigo_cliente)
    total = len(movimientos) or 1
    lineas = [f"Movimientos: {len(movimientos)}"]
    lineas += [f"  {banco}: {n}" for banco, n in por_banco.most_common()]
    lineas.append(f"  con contrato: {con_contrato} ({100 * con_contrato / total:.1f}%)")
    lineas.append(f"  con CUIT/CVU: {con_cuit} ({100 * con_cuit / total:.1f}%)")
    lineas.append(f"  con codigo de canal: {con_canal} ({100 * con_canal / total:.1f}%)")
    lineas.append(
        f"  con codigo de cliente del destino: {con_codigo} "
        f"({100 * con_codigo / total:.1f}%)"
    )
    return "\n".join(lineas)


def escribir_extracto(movimientos: list[MovimientoBancario], ruta: str | Path) -> Path:
    """Guarda el extracto estandarizado en un .xlsx nuevo.

    Es un archivo de salida, no una base que alguien edite a mano, asi que se
    escribe entero con openpyxl sin cuidar nada previo.
    """
    libro = openpyxl.Workbook()
    ws = libro.active
    ws.title = HOJA_EXTRACTO
    ws.append(ENCABEZADOS_EXTRACTO)
    for mov in movimientos:
        ws.append(mov.a_fila())
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:O{ws.max_row}"
    for fila in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fila[0].number_format = "dd/mm/yyyy"
        fila[1].number_format = "#,##0.00"
        fila[10].number_format = "dd/mm/yyyy"
    anchos = {"A": 11, "B": 14, "C": 12, "D": 30, "E": 14, "F": 10, "G": 60,
              "H": 14, "I": 12, "J": 14, "K": 12, "L": 14, "M": 16, "N": 7, "O": 18}
    for col, ancho in anchos.items():
        ws.column_dimensions[col].width = ancho
    ruta = Path(ruta)
    libro.save(ruta)
    libro.close()
    return ruta
