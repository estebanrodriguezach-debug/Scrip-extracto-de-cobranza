"""MODULO 2 - Motor de conciliacion automatica y semi-automatica.

Cruza el extracto estandarizado (Modulo 1) contra la hoja del mes de
``Cobranza General`` y completa ``Fecha de pago``, ``Monto de Pago`` y
``Medio de Pago``. Lo que no puede resolver con certeza no lo inventa: va a un
reporte para que una persona lo mire.

Niveles de coincidencia
-----------------------
1  contrato exacto + monto igual a la cuota (dentro de la tolerancia).
1c contrato exacto + el monto cubre **varias** cuotas pendientes de ese mismo
   contrato. Una sola transferencia de 90.000 paga tres cuotas de 30.000: se
   reparte, y cada fila recibe su propia cuota con la misma fecha y medio. Es
   lo que el equipo ya hace a mano -- de las 845 filas que la primera version
   del motor llenaba distinto, casi todas eran esto -- y es tambien el origen
   de la columna ``SubTotal`` del Listado banco, que lleva el total del grupo.
1b contrato exacto, monto distinto y no divisible en cuotas. Es un pago parcial
   o de mas, y se aplica igual porque el Modulo 3 ya sabe representarlo
   (``Estado`` = "Parcial"). Queda marcado en el reporte.
2  nombre parecido, cuando el extracto no trae contrato. Pide puntaje alto y
   **un solo** candidato claro: si dos clientes empatan, no se aplica nada.

Decisiones que vale la pena conocer
-----------------------------------
* **Uno a uno.** Cada movimiento del banco se aplica a una sola fila y cada
  fila recibe un solo movimiento. Sin esto, un cliente con tres cuotas iguales
  en el mes se comeria tres veces el mismo deposito.
* **La hoja destino sale del nombre de la hoja de origen**, no de la fecha del
  pago: un cupon pagado el 30/12 que figura en "Roela Enero" paga la cuota de
  enero. Si no aparece nada ahi, se prueban los meses vecinos.
* **Orden determinista.** Los movimientos se procesan ordenados por fecha,
  monto y clave, asi dos corridas sobre el mismo archivo dan exactamente el
  mismo resultado.
* **Greedy, no optimo.** Se podria plantear como asignacion de costo minimo,
  pero el resultado seria mas dificil de explicar y de auditar. Para quien
  revisa, "se aplico por contrato y monto exacto" vale mas que un apareo
  global imposible de reconstruir a mano.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from clasificador.cobranza import FilaCobranza, celdas_de_cobro
from clasificador.config import Config
from clasificador.extracto import MESES_DEL_ANIO, MovimientoBancario, claves_con_ocurrencia
from clasificador.patrones import parecido_nombres
from clasificador.reglas import JuegoDeReglas
from clasificador.xlsx_parche import Celda

NIVEL_1 = "1 contrato+monto"
NIVEL_1C = "1c contrato+varias cuotas"
NIVEL_1B = "1b contrato"
NIVEL_2 = "2 nombre"

# Motivos por los que un movimiento queda sin aplicar.
SIN_MES = "el mes de la hoja de origen no esta en Cobranza General"
FUERA_DE_PERIODO = "es de otro año que el de este libro"
SIN_CANDIDATO = "ningun cliente coincide por contrato ni por nombre"
TODAS_OCUPADAS = "el contrato existe pero sus filas ya tienen cobro cargado"
AMBIGUO = "mas de un cliente coincide por nombre con puntaje parecido"
PAGARE = "es un pagare: va a la hoja Pagare, que se carga a mano"


@dataclass
class Aplicacion:
    """Un movimiento del banco llevado a una fila concreta de la base."""

    movimiento: MovimientoBancario
    fila: FilaCobranza
    nivel: str
    puntaje: float = 1.0
    monto_aplicado: float | None = None
    total_del_grupo: float | None = None

    @property
    def monto(self) -> float:
        """Lo que se escribe en esta fila.

        Coincide con el monto del movimiento salvo cuando una transferencia se
        reparte entre varias cuotas: ahi cada fila recibe su propia cuota y el
        total queda en ``total_del_grupo``.
        """
        return self.movimiento.monto if self.monto_aplicado is None else self.monto_aplicado

    @property
    def diferencia(self) -> float:
        """Monto cobrado menos cuota esperada. Negativo = pago parcial."""
        if self.fila.cuota is None:
            return 0.0
        return self.monto - self.fila.cuota


@dataclass
class Pendiente:
    """Un movimiento que no se pudo aplicar, con el por que."""

    movimiento: MovimientoBancario
    motivo: str
    candidatos: list[tuple[FilaCobranza, float]] = field(default_factory=list)


@dataclass
class Resultado:
    aplicaciones: list[Aplicacion] = field(default_factory=list)
    ambiguos: list[Pendiente] = field(default_factory=list)
    no_identificados: list[Pendiente] = field(default_factory=list)
    pagares: list[Pendiente] = field(default_factory=list)
    fuera_de_periodo: int = 0
    fuera_de_alcance: int = 0

    @property
    def total_aplicado(self) -> float:
        return sum(a.monto for a in self.aplicaciones)

    @property
    def movimientos_aplicados(self) -> int:
        """Movimientos distintos aplicados; uno repartido cuenta una vez."""
        return len({id(a.movimiento) for a in self.aplicaciones})

    def por_nivel(self) -> dict[str, int]:
        conteo: dict[str, int] = {}
        for a in self.aplicaciones:
            conteo[a.nivel] = conteo.get(a.nivel, 0) + 1
        return conteo

    def celdas(self, cfg: Config) -> list[Celda]:
        """Traduce las aplicaciones a celdas para ``xlsx_parche.parchear``."""
        celdas: list[Celda] = []
        for a in self.aplicaciones:
            celdas.extend(
                celdas_de_cobro(
                    a.fila,
                    cfg,
                    fecha=a.movimiento.fecha,
                    monto=a.monto,
                    medio=a.movimiento.medio_cobranza,
                )
            )
        return celdas


def _meses_a_probar(mes: str, vecinos: int, disponibles: list[str]) -> list[str]:
    """El mes propio primero y despues los vecinos, de mas cerca a mas lejos."""
    if mes not in MESES_DEL_ANIO:
        return [m for m in disponibles if m == mes]
    i = MESES_DEL_ANIO.index(mes)
    orden = [mes]
    for paso in range(1, vecinos + 1):
        for j in (i + paso, i - paso):
            if 0 <= j < len(MESES_DEL_ANIO):
                orden.append(MESES_DEL_ANIO[j])
    return [m for m in orden if m in disponibles]


def _coincide_monto(monto: float, fila: FilaCobranza, tolerancia: float) -> bool:
    if fila.cuota is None:
        return False
    return abs(monto - fila.cuota) <= tolerancia


def _reparto_en_cuotas(
    monto: float, filas: list[FilaCobranza], tolerancia: float
) -> list[FilaCobranza]:
    """Las primeras N cuotas pendientes cuya suma da el monto transferido.

    Se avanza en el orden en que estan las filas en la planilla, que es el
    orden en que el equipo las paga, y se exige que la suma cierre exacta: un
    reparto aproximado desparramaria un error por varias filas a la vez.
    """
    acumulado = 0.0
    tomadas: list[FilaCobranza] = []
    for fila in filas:
        if fila.cuota is None or fila.cuota <= 0:
            return []
        acumulado += fila.cuota
        tomadas.append(fila)
        if abs(acumulado - monto) <= tolerancia:
            return tomadas if len(tomadas) > 1 else []
        if acumulado > monto + tolerancia:
            return []
    return []


def _mejor_por_nombre(
    mov: MovimientoBancario,
    candidatas: list[FilaCobranza],
    cfg: Config,
) -> tuple[FilaCobranza | None, list[tuple[FilaCobranza, float]]]:
    """Busca el cliente por nombre. Devuelve (elegida o None, ranking)."""
    umbral = float(cfg.conciliacion["umbral_nombre"])
    margen = float(cfg.conciliacion["margen_nombre"])
    tolerancia = float(cfg.conciliacion["tolerancia_monto"])
    exige_monto = bool(cfg.conciliacion["exigir_monto_en_nivel2"])

    nombre = mov.titular_origen
    if not nombre:
        return None, []
    puntajes = []
    for fila in candidatas:
        if exige_monto and not _coincide_monto(mov.monto, fila, tolerancia):
            continue
        p = parecido_nombres(nombre, fila.nombre)
        if p >= umbral:
            puntajes.append((fila, p))
    if not puntajes:
        return None, []
    puntajes.sort(key=lambda t: (-t[1], t[0].fila))
    if len(puntajes) > 1 and puntajes[0][1] - puntajes[1][1] < margen:
        return None, puntajes[:5]  # empate: que lo resuelva una persona
    return puntajes[0][0], puntajes[:5]


def conciliar(
    movimientos: list[MovimientoBancario],
    base: dict[str, list[FilaCobranza]],
    cfg: Config,
    juego: JuegoDeReglas | None = None,
    anio: int | None = None,
) -> Resultado:
    """Cruza el extracto contra la base y devuelve el plan de aplicacion.

    No escribe nada: devuelve un ``Resultado`` del que se sacan las celdas. Asi
    se puede mostrar al operador lo que va a pasar antes de tocar el archivo
    (``--simular``).

    ``anio`` descarta los movimientos de otros años. Es imprescindible: el
    libro de origen trae 2025 y 2026, y un cobro de enero de 2025 suele tener
    el mismo monto que la cuota de enero de 2026, asi que sin filtrar por año
    el cruce se equivoca de fila sin que se note.
    """
    tolerancia = float(cfg.conciliacion["tolerancia_monto"])
    vecinos = int(cfg.conciliacion.get("meses_vecinos", 0))
    permitir_1b = bool(cfg.conciliacion["aplicar_contrato_sin_monto"])
    disponibles = [m for m in cfg.meses if m in base]

    # Indices por mes: contrato -> filas pendientes, en orden de aparicion.
    por_contrato: dict[str, dict[str, list[FilaCobranza]]] = {}
    pendientes_mes: dict[str, list[FilaCobranza]] = {}
    for mes, filas in base.items():
        indice: dict[str, list[FilaCobranza]] = {}
        libres = []
        for fila in filas:
            if not fila.pendiente:
                continue
            libres.append(fila)
            if fila.contrato:
                indice.setdefault(fila.contrato, []).append(fila)
        por_contrato[mes] = indice
        pendientes_mes[mes] = libres

    ocupadas: set[tuple[str, int]] = set()
    resultado = Resultado()

    def libre(fila: FilaCobranza) -> bool:
        return (fila.hoja, fila.fila) not in ocupadas

    def tomar(fila: FilaCobranza) -> None:
        ocupadas.add((fila.hoja, fila.fila))

    ordenados = sorted(
        zip(movimientos, claves_con_ocurrencia(movimientos)),
        key=lambda t: (t[0].fecha or dt.date.min, t[0].monto, t[1]),
    )

    for mov, _clave in ordenados:
        if anio and mov.anio_origen and mov.anio_origen != anio:
            resultado.fuera_de_periodo += 1
            continue

        # Los pagares no viven en las hojas mensuales: tienen su propia hoja,
        # con otro layout y sin numero de contrato. Se apartan para que no
        # ensucien el reporte de no identificados.
        if juego is not None and juego.clasificar(mov.detalle).regla_id == "pagare":
            resultado.pagares.append(Pendiente(mov, PAGARE))
            continue

        mes = mov.mes_origen or (
            MESES_DEL_ANIO[mov.fecha.month - 1] if mov.fecha else ""
        )
        candidatos_mes = _meses_a_probar(mes, vecinos, disponibles)
        if not candidatos_mes:
            # El mes del movimiento no esta entre las hojas cargadas. Pasa
            # siempre que se corre con --mes, y no es un pago sin identificar:
            # simplemente no corresponde a esta pasada.
            resultado.fuera_de_alcance += 1
            continue

        aplicado = False
        habia_contrato = False
        ranking: list[tuple[FilaCobranza, float]] = []

        for intento in candidatos_mes:
            filas_contrato = [
                f
                for f in por_contrato[intento].get(mov.contrato_detectado, [])
                if libre(f)
            ]
            if mov.contrato_detectado and por_contrato[intento].get(
                mov.contrato_detectado
            ):
                habia_contrato = True

            # Nivel 1: contrato y monto.
            elegida = next(
                (f for f in filas_contrato if _coincide_monto(mov.monto, f, tolerancia)),
                None,
            )
            nivel = NIVEL_1
            # Nivel 1c: el monto cubre varias cuotas del mismo contrato.
            if elegida is None and len(filas_contrato) > 1:
                grupo = _reparto_en_cuotas(mov.monto, filas_contrato, tolerancia)
                if grupo:
                    for fila_grupo in grupo:
                        tomar(fila_grupo)
                        resultado.aplicaciones.append(
                            Aplicacion(
                                mov,
                                fila_grupo,
                                NIVEL_1C,
                                monto_aplicado=fila_grupo.cuota,
                                total_del_grupo=mov.monto,
                            )
                        )
                    aplicado = True
                    break
            # Nivel 1b: contrato solo.
            if elegida is None and permitir_1b and filas_contrato:
                elegida, nivel = filas_contrato[0], NIVEL_1B
            # Nivel 2: nombre.
            puntaje = 1.0
            if elegida is None and not mov.contrato_detectado:
                libres = [f for f in pendientes_mes[intento] if libre(f)]
                elegida, ranking = _mejor_por_nombre(mov, libres, cfg)
                if elegida is not None:
                    nivel = NIVEL_2
                    puntaje = next(p for f, p in ranking if f is elegida)

            if elegida is not None:
                tomar(elegida)
                resultado.aplicaciones.append(
                    Aplicacion(mov, elegida, nivel, puntaje)
                )
                aplicado = True
                break

        if aplicado:
            continue
        if ranking:
            resultado.ambiguos.append(Pendiente(mov, AMBIGUO, ranking))
        elif habia_contrato:
            resultado.no_identificados.append(Pendiente(mov, TODAS_OCUPADAS))
        else:
            resultado.no_identificados.append(Pendiente(mov, SIN_CANDIDATO))

    return resultado


# ---------------------------------------------------------------------------
# Reporte
# ---------------------------------------------------------------------------

_AZUL = "FF1F4E78"
_BLANCO = "FFFFFFFF"


def _encabezar(ws, titulos: list[str]) -> None:
    ws.append(titulos)
    relleno = PatternFill("solid", fgColor=_AZUL)
    for celda in ws[1]:
        celda.font = Font(name="Calibri", size=11, bold=True, color=_BLANCO)
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"


def escribir_reporte(resultado: Resultado, ruta: str | Path) -> Path:
    """Reporte de la corrida: lo aplicado, lo ambiguo y lo que nadie reclamo.

    La hoja que de verdad importa es "No identificados": es la cola de trabajo
    del equipo. Trae todo lo que se sabe del movimiento (hoja y fila del libro
    de origen incluidas) para poder ir a buscarlo sin adivinar.
    """
    libro = openpyxl.Workbook()

    ws = libro.active
    ws.title = "Resumen"
    _encabezar(ws, ["Concepto", "Movimientos", "Importe"])
    ws.append(["Filas completadas", len(resultado.aplicaciones), resultado.total_aplicado])
    ws.append(["Movimientos aplicados", resultado.movimientos_aplicados, None])
    for nivel, n in sorted(resultado.por_nivel().items()):
        ws.append([f"   nivel {nivel}", n, None])
    ws.append(
        [
            "Ambiguos (revisar)",
            len(resultado.ambiguos),
            sum(p.movimiento.monto for p in resultado.ambiguos),
        ]
    )
    ws.append(
        [
            "No identificados",
            len(resultado.no_identificados),
            sum(p.movimiento.monto for p in resultado.no_identificados),
        ]
    )
    ws.append(
        [
            "Pagares (carga manual)",
            len(resultado.pagares),
            sum(p.movimiento.monto for p in resultado.pagares),
        ]
    )
    ws.append(["De otro año (ignorados)", resultado.fuera_de_periodo, None])
    ws.append(["De un mes fuera de esta pasada", resultado.fuera_de_alcance, None])
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16

    hoja = libro.create_sheet("Aplicados")
    _encabezar(
        hoja,
        [
            "Nivel", "Puntaje", "Mes", "Fila", "Contrato", "Cliente",
            "Fecha pago", "Monto", "Cuota", "Diferencia",
            "Total transferido", "Medio", "Canal", "Titular del pago",
            "Banco", "Hoja origen", "Fila origen",
        ],
    )
    for a in resultado.aplicaciones:
        hoja.append(
            [
                a.nivel, round(a.puntaje, 3), a.fila.hoja, a.fila.fila,
                a.fila.contrato, a.fila.nombre, a.movimiento.fecha,
                a.monto, a.fila.cuota, a.diferencia, a.total_del_grupo,
                a.movimiento.medio_cobranza, a.movimiento.codigo_canal,
                a.movimiento.titular_origen, a.movimiento.banco,
                a.movimiento.hoja, a.movimiento.fila,
            ]
        )

    for titulo, items in (
        ("Ambiguos", resultado.ambiguos),
        ("No identificados", resultado.no_identificados),
        ("Pagares", resultado.pagares),
    ):
        hoja = libro.create_sheet(titulo)
        _encabezar(
            hoja,
            [
                "Motivo", "Fecha", "Monto", "Banco", "Medio", "Canal",
                "Contrato detectado", "Titular del pago", "CUIT/CVU",
                "Concepto", "Detalle", "Hoja origen", "Fila origen",
                "Candidatos",
            ],
        )
        for p in items:
            m = p.movimiento
            hoja.append(
                [
                    p.motivo, m.fecha, m.monto, m.banco, m.medio_cobranza,
                    m.codigo_canal, m.contrato_detectado, m.titular_origen,
                    m.cuit_cvu, m.concepto, m.detalle, m.hoja, m.fila,
                    "; ".join(
                        f"{f.nombre} (fila {f.fila}, {p2:.2f})"
                        for f, p2 in p.candidatos
                    ),
                ]
            )

    for hoja in libro.worksheets:
        if hoja.title == "Resumen":
            continue
        hoja.auto_filter.ref = hoja.dimensions
        for col, ancho in (("A", 18), ("F", 30), ("H", 26), ("J", 30)):
            hoja.column_dimensions[col].width = ancho

    ruta = Path(ruta)
    libro.save(ruta)
    libro.close()
    return ruta
