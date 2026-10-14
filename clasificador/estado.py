"""MODULO 3 (parte 1) - Logica de Estado.

Lo que habia
------------
``Estado`` era ``=COUNT(G3)``: 1 si la celda de fecha de pago tiene algo, 0 si
esta vacia. Mide si **alguien cargo el dato**, no si el cliente pago lo que
debia. Un cliente con cuota de 100.000 que deposito 10.000 queda igual que uno
que pago todo.

Lo que queda
-----------
::

    Estado = 0          si Monto de Pago esta vacio o es 0   (pendiente)
    Estado = "Parcial"  si 0 < Monto de Pago < Cuota actual
    Estado = 1          si Monto de Pago >= Cuota actual     (pagado)

Se escribe como **formula**, no como valor, por dos razones: la planilla sigue
viva (si alguien corrige un monto a mano, el estado se actualiza solo, que es
como viene funcionando hasta hoy) y no hace falta volver a correr nada.

El tercer estado que ya existia y no estaba en la especificacion
----------------------------------------------------------------
En el archivo real hay 620 filas con ``Estado`` = **3** cargado a mano, con la
palabra 'baja' escrita en la columna de fecha de pago: son contratos dados de
baja. Esas filas no se tocan -- ni aca ni en la conciliacion ni en el tablero
de mora. Pisarlas con una formula borraria una decision de negocio que nadie
puede reconstruir desde los datos.
"""

from __future__ import annotations

from openpyxl.utils import get_column_letter

from clasificador.cobranza import FilaCobranza
from clasificador.config import Config
from clasificador.xlsx_parche import Celda

PENDIENTE = 0
PAGADO = 1
PARCIAL = "Parcial"


def estado_de(
    monto_pago: float | None, cuota: float | None, parcial: str = PARCIAL
) -> object:
    """La misma regla que la formula, para los reportes y los tests.

    Tiene que dar exactamente lo mismo que ``formula_estado``: si las dos se
    separan, el tablero de mora diria una cosa y la planilla otra.
    """
    monto = monto_pago or 0.0
    if monto <= 0:
        return PENDIENTE
    if cuota is None or cuota <= 0:
        # Sin cuota de referencia no hay con que comparar: si pago, pago.
        return PAGADO
    return PAGADO if monto >= cuota else parcial


def formula_estado(fila: int, cfg: Config) -> str:
    """La formula de Estado para una fila, con las columnas de la config.

    Se usa ``N()`` para que una celda vacia o con texto cuente como 0 en lugar
    de romper la formula con #VALUE!: en el archivo real la columna de monto
    tiene celdas con texto suelto.
    """
    col_monto = get_column_letter(cfg.indice("monto_pago"))
    col_cuota = get_column_letter(cfg.indice("cuota"))
    parcial = cfg.mora.get("estado_parcial", PARCIAL)
    monto = f"N({col_monto}{fila})"
    cuota = f"N({col_cuota}{fila})"
    return f'IF({monto}<=0,0,IF({monto}>={cuota},1,"{parcial}"))'


def _hay_que_tocar(fila: FilaCobranza) -> bool:
    """No se pisa una baja ni una anotacion manual en la columna Estado."""
    if fila.es_baja:
        return False
    if isinstance(fila.estado, str) and fila.estado.strip():
        # Textos como 'pago pero se dio de baja' o 'se saco porque adjudico'
        # son decisiones de una persona, no un estado calculable.
        try:
            float(fila.estado.replace(",", "."))
        except ValueError:
            return False
    return True


def celdas_de_estado(
    base: dict[str, list[FilaCobranza]], cfg: Config
) -> tuple[list[Celda], dict[str, int]]:
    """Las celdas de ``Estado`` a reescribir. Devuelve (celdas, conteo)."""
    celdas: list[Celda] = []
    conteo = {"escritas": 0, "bajas": 0, "manuales": 0}
    columna = cfg.indice("estado")
    for filas in base.values():
        for fila in filas:
            if fila.es_baja:
                conteo["bajas"] += 1
                continue
            if not _hay_que_tocar(fila):
                conteo["manuales"] += 1
                continue
            celdas.append(
                Celda(
                    fila.hoja,
                    fila.fila,
                    columna,
                    formula=formula_estado(fila.fila, cfg),
                )
            )
            conteo["escritas"] += 1
    return celdas, conteo


def resumen_estados(
    base: dict[str, list[FilaCobranza]], cfg: Config
) -> dict[str, int]:
    """Como quedaria el reparto de estados con la regla nueva."""
    parcial = cfg.mora.get("estado_parcial", PARCIAL)
    conteo = {"pagado": 0, "parcial": 0, "pendiente": 0, "baja": 0}
    for filas in base.values():
        for fila in filas:
            if fila.es_baja:
                conteo["baja"] += 1
                continue
            valor = estado_de(fila.monto_pago, fila.cuota, parcial)
            if valor == PAGADO:
                conteo["pagado"] += 1
            elif valor == parcial:
                conteo["parcial"] += 1
            else:
                conteo["pendiente"] += 1
    return conteo
