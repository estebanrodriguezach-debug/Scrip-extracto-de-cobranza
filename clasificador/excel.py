"""Lectura de la base Extractos_cobranzas.xlsx."""

from __future__ import annotations

import datetime as dt
import hashlib
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import IO

import openpyxl

HOJA_DATOS = "Cobranzas"
ENCABEZADOS = ["Fecha", "Banco", "Importe", "Extracto", "Detalle", "Categoría"]


@dataclass
class Movimiento:
    fecha: dt.date | None
    banco: str
    importe: float | None
    extracto: str
    detalle: str
    categoria: str = ""
    clave: str = field(default="", init=False)

    def __post_init__(self) -> None:
        self.clave = self.calcular_clave()

    def calcular_clave(self) -> str:
        """Huella del movimiento segun lo que informa el banco.

        No incluye ``detalle`` ni ``categoria`` porque son anotaciones nuestras:
        si alguien edita la nota en la planilla, el movimiento tiene que seguir
        siendo el mismo y no duplicarse en el proximo volcado.
        """
        partes = [
            self.fecha.isoformat() if self.fecha else "",
            self.banco.strip().lower(),
            f"{self.importe:.2f}" if self.importe is not None else "",
            " ".join(self.extracto.split()).lower(),
        ]
        return hashlib.sha1("|".join(partes).encode("utf-8")).hexdigest()[:16]

    def a_fila(self) -> list[object]:
        return [
            self.fecha.isoformat() if self.fecha else "",
            self.banco,
            self.importe if self.importe is not None else "",
            self.extracto,
            self.detalle,
            self.categoria,
        ]


def _a_fecha(valor: object) -> dt.date | None:
    if valor is None or valor == "":
        return None
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    return dt.date.fromisoformat(str(valor)[:10])


def _a_texto(valor: object) -> str:
    return "" if valor is None else str(valor).strip()


def _a_importe(valor: object) -> float | None:
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    limpio = str(valor).replace(".", "").replace(",", ".").replace("$", "").strip()
    return float(limpio) if limpio else None


def claves_con_ocurrencia(movimientos: list[Movimiento]) -> list[str]:
    """Desambigua movimientos identicos (misma fecha, banco, importe y texto).

    En la base real hay 48 pares asi: son cobros repetidos legitimos, no
    duplicados a descartar. Se numeran para que el volcado sea idempotente sin
    perderlos. Lo usan los dos destinos (Google Sheets y SharePoint), por eso
    vive aca y no en uno de ellos.
    """
    vistos: Counter[str] = Counter()
    claves = []
    for mov in movimientos:
        vistos[mov.clave] += 1
        claves.append(f"{mov.clave}#{vistos[mov.clave]}")
    return claves


def leer_movimientos(
    ruta: str | Path | IO[bytes], hoja: str = HOJA_DATOS
) -> list[Movimiento]:
    """Lee el extracto. ``ruta`` puede ser un archivo o un buffer en memoria.

    Lo segundo es lo que usa SharePoint: el .xlsx se descarga a un BytesIO y se
    lee sin tocar el disco.
    """
    libro = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    try:
        ws = libro[hoja]
        movimientos = []
        for fila in ws.iter_rows(min_row=2, max_col=6, values_only=True):
            fila = list(fila) + [None] * (6 - len(fila))
            if all(v is None or v == "" for v in fila):
                continue
            movimientos.append(
                Movimiento(
                    fecha=_a_fecha(fila[0]),
                    banco=_a_texto(fila[1]),
                    importe=_a_importe(fila[2]),
                    extracto=_a_texto(fila[3]),
                    detalle=_a_texto(fila[4]),
                    categoria=_a_texto(fila[5]),
                )
            )
        return movimientos
    finally:
        libro.close()
