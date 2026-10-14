"""Carga de ``conciliacion.json``, los parametros de los Modulos 1 a 4.

Mismo criterio que ``reglas.py`` con ``reglas.json``: lo que depende de como
esta armado el Excel (en que fila esta el encabezado, que letra es cada
columna) o de una regla del negocio (tolerancias, codigos de canal) vive en un
archivo de configuracion, no en el codigo. Asi el equipo puede corregir un
cambio de layout sin tocar Python.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from functools import cached_property
from pathlib import Path

from openpyxl.utils import column_index_from_string

RUTA_CONFIG_POR_DEFECTO = Path(__file__).resolve().parent.parent / "conciliacion.json"


@dataclass(frozen=True)
class Config:
    """Vista de solo lectura sobre el JSON, con los atajos que mas se usan."""

    datos: dict

    @property
    def cobranza(self) -> dict:
        return self.datos["cobranza_general"]

    @property
    def conciliacion(self) -> dict:
        return self.datos["conciliacion"]

    @property
    def mora(self) -> dict:
        return self.datos["mora"]

    @property
    def listado(self) -> dict:
        return self.datos["listado"]

    @cached_property
    def indices(self) -> dict[str, int]:
        """Nombre logico de columna -> indice 1-based, listo para openpyxl.

        El JSON guarda letras ("G") porque es lo que el equipo ve en Excel;
        el codigo necesita numeros. La traduccion se hace una sola vez.
        """
        return {
            nombre: column_index_from_string(letra)
            for nombre, letra in self.cobranza["columnas"].items()
        }

    def indice(self, campo: str) -> int:
        try:
            return self.indices[campo]
        except KeyError:
            raise KeyError(
                f"En conciliacion.json falta la columna {campo!r} dentro de "
                f"cobranza_general.columnas. Estan: {sorted(self.indices)}"
            ) from None

    @property
    def meses(self) -> list[str]:
        return list(self.cobranza["meses"])


def cargar_config(ruta: str | Path | None = None) -> Config:
    ruta = Path(ruta) if ruta else RUTA_CONFIG_POR_DEFECTO
    datos = json.loads(ruta.read_text(encoding="utf-8"))
    for seccion in ("cobranza_general", "conciliacion", "mora", "listado"):
        if seccion not in datos:
            raise ValueError(f"{ruta.name} no tiene la seccion {seccion!r}.")
    cfg = Config(datos)
    cfg.indices  # valida las letras de columna ahora, no en medio de un volcado
    return cfg
