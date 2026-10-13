"""Casos de clasificacion, tomados de las notas reales de la base."""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from clasificador.reglas import cargar_reglas, normalizar  # noqa: E402

CASOS = json.loads((RAIZ / "tests" / "casos.json").read_text(encoding="utf-8"))


class TestNormalizar(unittest.TestCase):
    def test_saca_acentos_y_colapsa_espacios(self):
        self.assertEqual(normalizar("  Pagaré   X "), "pagare x")

    def test_none_es_vacio(self):
        self.assertEqual(normalizar(None), "")


class TestClasificar(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.juego = cargar_reglas()

    def test_casos(self):
        for caso in CASOS:
            with self.subTest(detalle=caso["detalle"]):
                self.assertEqual(
                    self.juego.clasificar(caso["detalle"]).categoria,
                    caso["categoria"],
                )

    def test_toda_regla_apunta_a_categoria_valida(self):
        for regla in self.juego.reglas:
            self.assertIn(regla.categoria, self.juego.categorias_validas)

    def test_prioridad_pagare_sobre_cuota_1(self):
        self.assertEqual(
            self.juego.clasificar("pagare / cuota 1").categoria,
            "Cobranza financiamiento clientes (pagaré)",
        )


class TestAppsScriptSincronizado(unittest.TestCase):
    """Codigo.gs se genera desde reglas.json: no debe quedar desfasado."""

    def test_codigo_gs_al_dia(self):
        import subprocess

        destino = RAIZ / "apps_script" / "Codigo.gs"
        antes = destino.read_text(encoding="utf-8")
        subprocess.run(
            [sys.executable, str(RAIZ / "scripts" / "generar_apps_script.py")],
            check=True,
            capture_output=True,
        )
        self.assertEqual(
            antes,
            destino.read_text(encoding="utf-8"),
            "apps_script/Codigo.gs esta desactualizado: corre "
            "`python3 scripts/generar_apps_script.py` y commitea el resultado.",
        )


if __name__ == "__main__":
    unittest.main()
