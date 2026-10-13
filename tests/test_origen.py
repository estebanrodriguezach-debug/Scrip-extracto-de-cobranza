"""Deteccion de banco y de columnas en el libro de Ingresos Bancarios."""

from __future__ import annotations

import datetime as dt
import sys
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from clasificador.origen import Bloque, banco_de_hoja, detectar_bloques  # noqa: E402


class TestBancoDeHoja(unittest.TestCase):
    def test_nombres_reales_del_libro(self):
        casos = {
            "Rio Enero": "Rio",
            "Macro Feb 26": "Macro",
            "Galicia oct": "Galicia",
            "Roela  Julio 26": "Roela",
            "Patag Dic": "Patagonia",
            "Pata Ene 26": "Patagonia",
            "Patagonia sep 26": "Patagonia",
            "Frances Sep 26": "Frances",
        }
        for titulo, esperado in casos.items():
            with self.subTest(titulo=titulo):
                self.assertEqual(banco_de_hoja(titulo), esperado)

    def test_hoja_desconocida(self):
        self.assertIsNone(banco_de_hoja("Resumen anual"))


def _fila(celdas: dict[int, object]) -> tuple:
    fila = [None] * (max(celdas) + 1)
    for i, v in celdas.items():
        fila[i] = v
    return tuple(fila)


class TestDetectarBloques(unittest.TestCase):
    def test_layout_generico(self):
        """Rio/Galicia/Macro: extracto A, fecha B, importe C."""
        filas = [
            _fila({0: f"Transferencia Recibida numero {i}", 1: dt.datetime(2026, 7, i + 1),
                     2: 150000.0, 3: 3499.0, 4: "Cattaneo Pamela", 5: ""})
            for i in range(10)
        ]
        bloques = detectar_bloques(filas, 6, "Galicia")
        self.assertEqual(len(bloques), 1)
        self.assertEqual(
            bloques[0],
            Bloque(fecha=1, importe=2, contrato=3, nombre=4, detalle=5, extracto=(0,)),
        )

    def test_patagonia_corrida_una_columna(self):
        """La misma deteccion tiene que servir aunque las columnas se corran."""
        filas = [
            _fila({0: "TRANSFERENCIA ENTRE CUENTAS", 1: f"VAR 2037364539{i} JUAN GABRIEL",
                     2: dt.datetime(2026, 10, i + 1), 3: 30000.0, 4: 2370.0,
                     5: "Currecoy Juan Gabriel", 6: ""})
            for i in range(10)
        ]
        bloques = detectar_bloques(filas, 7, "Patagonia")
        self.assertEqual(len(bloques), 1)
        self.assertEqual(bloques[0].fecha, 2)
        self.assertEqual(bloques[0].importe, 3)
        self.assertEqual(bloques[0].extracto, (0, 1))

    def test_roela_dos_bloques(self):
        """Roela trae cheques a la izquierda y transferencias a la derecha."""
        filas = [
            _fila({0: dt.datetime(2026, 7, i + 1), 1: 400000.01, 2: 32082.0,
                     3: dt.datetime(2026, 6, 30), 4: dt.datetime(2026, 6, 30),
                     5: "TQR", 6: 3208.0, 7: "Pared Victor Hugo",
                     11: 290.0, 12: "TRANSF.INTER.DISTINTO TIT.-20354",
                     13: dt.datetime(2026, 7, i + 1), 14: 40000.0, 15: 1012.0,
                     16: "Viveros Linares Bruno", 17: ""})
            for i in range(10)
        ]
        bloques = detectar_bloques(filas, 18, "Roela")
        self.assertEqual(len(bloques), 2)
        cheques, transferencias = bloques
        # La fecha del cheque es la D, no la A en que se carga.
        self.assertEqual(cheques.fecha, 3)
        self.assertEqual(cheques.importe, 1)
        self.assertEqual(cheques.detalle, 5)
        self.assertEqual(cheques.nombre, 7)
        self.assertEqual(transferencias.fecha, 13)
        self.assertEqual(transferencias.importe, 14)

    def test_hoja_sin_fechas(self):
        filas = [_fila({0: "texto suelto", 1: "otra cosa"}) for _ in range(5)]
        self.assertEqual(detectar_bloques(filas, 2, "Rio"), [])


if __name__ == "__main__":
    unittest.main()
