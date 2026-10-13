"""Motor de reglas de clasificacion.

Unica fuente de verdad: ``reglas.json`` en la raiz del repo. El mismo archivo se
inyecta en el Apps Script (ver ``scripts/generar_apps_script.py``) para que la
planilla de Google clasifique exactamente igual que este codigo.
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

RUTA_REGLAS_POR_DEFECTO = Path(__file__).resolve().parent.parent / "reglas.json"

_ESPACIOS = re.compile(r"\s+")


def normalizar(texto: object) -> str:
    """Pasa a minusculas, saca acentos y colapsa espacios.

    Las notas de ``Detalle`` las escriben personas distintas: aparecen
    ``pagare``/``Pagare``, ``presupuesto``/``PRESUPUESTO PINTURA``. Normalizar
    antes de comparar evita una regla por cada variante.
    """
    if texto is None:
        return ""
    s = str(texto)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return _ESPACIOS.sub(" ", s).strip().lower()


@dataclass(frozen=True)
class Regla:
    id: str
    categoria: str
    contiene: tuple[str, ...]

    def coincide(self, detalle_normalizado: str) -> str | None:
        for palabra in self.contiene:
            if palabra in detalle_normalizado:
                return palabra
        return None


@dataclass(frozen=True)
class Resultado:
    categoria: str
    regla_id: str
    palabra: str | None


@dataclass(frozen=True)
class JuegoDeReglas:
    version: int
    categoria_por_defecto: str
    categoria_sin_dato: str
    categorias_validas: tuple[str, ...]
    reglas: tuple[Regla, ...]

    def clasificar(self, detalle: object) -> Resultado:
        normalizado = normalizar(detalle)
        if not normalizado:
            return Resultado(self.categoria_sin_dato, "sin_detalle", None)
        for regla in self.reglas:
            palabra = regla.coincide(normalizado)
            if palabra is not None:
                return Resultado(regla.categoria, regla.id, palabra)
        return Resultado(self.categoria_por_defecto, "por_defecto", None)


def cargar_reglas(ruta: str | Path | None = None) -> JuegoDeReglas:
    ruta = Path(ruta) if ruta else RUTA_REGLAS_POR_DEFECTO
    datos = json.loads(ruta.read_text(encoding="utf-8"))

    validas = tuple(datos["categorias_validas"])
    reglas = tuple(
        Regla(
            id=r["id"],
            categoria=r["categoria"],
            contiene=tuple(normalizar(p) for p in r["contiene"]),
        )
        for r in datos["reglas"]
    )

    for regla in reglas:
        if regla.categoria not in validas:
            raise ValueError(
                f"La regla '{regla.id}' apunta a una categoria que no esta en "
                f"categorias_validas: {regla.categoria!r}"
            )

    juego = JuegoDeReglas(
        version=datos["version"],
        categoria_por_defecto=datos["categoria_por_defecto"],
        categoria_sin_dato=datos["categoria_sin_dato"],
        categorias_validas=validas,
        reglas=reglas,
    )
    for nombre in (juego.categoria_por_defecto, juego.categoria_sin_dato):
        if nombre not in validas:
            raise ValueError(f"Categoria por defecto desconocida: {nombre!r}")
    return juego


def clasificar(detalle: object, juego: JuegoDeReglas | None = None) -> str:
    """Atajo: devuelve solo el nombre de la categoria."""
    return (juego or cargar_reglas()).clasificar(detalle).categoria
