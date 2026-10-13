"""Clasificador de extractos de cobranza y sincronizacion con Google Sheets."""

from clasificador.reglas import (
    JuegoDeReglas,
    cargar_reglas,
    clasificar,
    normalizar,
)

__all__ = ["JuegoDeReglas", "cargar_reglas", "clasificar", "normalizar"]
