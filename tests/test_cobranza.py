"""Lectura de Cobranza General: el layout real y sus trampas."""

from __future__ import annotations

import argparse
import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(RAIZ / "tests"))

import fixtures  # noqa: E402

from clasificador.cobranza import anio_de_base, leer_cobranza  # noqa: E402
from clasificador.config import cargar_config  # noqa: E402


class TestLeerCobranza(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cfg = cargar_config()
        cls.tmp = tempfile.TemporaryDirectory()
        cls.ruta = fixtures.libro_cobranza(
            Path(cls.tmp.name) / "cobranza.xlsx",
            {
                "Enero": [
                    {
                        "nombre": "Aban Albino Daniela",
                        "contrato": 2134,
                        "fecha_pago": dt.datetime(2026, 1, 14),
                        "monto_pago": 50000,
                        "medio_pago": "roela",
                        "cuota": 15000,
                        "vto_cuota": dt.datetime(2026, 1, 10),
                    },
                    {"nombre": "Pendiente Pedro", "contrato": 2760, "cuota": 30000},
                    {
                        "nombre": "Parcial Paula",
                        "contrato": 886,
                        "fecha_pago": dt.datetime(2026, 1, 20),
                        "monto_pago": 5000,
                        "medio_pago": "rio",
                        "cuota": 10000,
                    },
                    {
                        # Baja: Estado 3 a mano y la palabra en la fecha de pago.
                        "nombre": "Baja Bruno",
                        "contrato": 1318,
                        "fecha_pago": "baja",
                        "cuota": 15000,
                        "estado": 3,
                    },
                    {"nombre": "Sin contrato Sofia", "contrato": None, "cuota": 20000},
                ],
                "Febrero": [
                    {"nombre": "Pendiente Pedro", "contrato": 2760, "cuota": 30000}
                ],
            },
        )
        cls.base, cls.avisos = leer_cobranza(cls.ruta, cls.cfg)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def _fila(self, nombre, mes="Enero"):
        return next(f for f in self.base[mes] if f.nombre.startswith(nombre))

    def test_no_hay_avisos_de_encabezado(self):
        self.assertEqual(self.avisos, [])

    def test_descarta_el_bloque_de_estadisticas(self):
        """Las filas 'Cobrado'/'Pendiente' de abajo no son clientes."""
        self.assertEqual(len(self.base["Enero"]), 5)
        self.assertNotIn("Cobrado", [f.nombre for f in self.base["Enero"]])

    def test_lee_los_campos_del_cobro(self):
        fila = self._fila("Aban")
        self.assertEqual(fila.contrato, "2134")
        self.assertEqual(fila.fecha_pago, dt.date(2026, 1, 14))
        self.assertEqual(fila.monto_pago, 50000.0)
        self.assertEqual(fila.medio_pago, "roela")
        self.assertEqual(fila.cuota, 15000.0)
        self.assertEqual(fila.vto_cuota, dt.date(2026, 1, 10))

    def test_pendiente_y_tiene_cobro(self):
        self.assertTrue(self._fila("Pendiente").pendiente)
        self.assertFalse(self._fila("Aban").pendiente)
        self.assertTrue(self._fila("Aban").tiene_cobro)

    def test_una_baja_no_es_pendiente(self):
        baja = self._fila("Baja")
        self.assertTrue(baja.es_baja)
        self.assertFalse(baja.pendiente)
        self.assertEqual(baja.texto_fecha_pago, "baja")
        self.assertIsNone(baja.fecha_pago)

    def test_saldo(self):
        self.assertEqual(self._fila("Parcial").saldo, 5000.0)
        self.assertEqual(self._fila("Aban").saldo, 0.0)
        self.assertEqual(self._fila("Pendiente").saldo, 30000.0)

    def test_fila_sin_contrato_se_conserva(self):
        """Sirve para el nivel 2, que cruza por nombre."""
        fila = self._fila("Sin contrato")
        self.assertEqual(fila.contrato, "")
        self.assertTrue(fila.pendiente)

    def test_anio_de_base(self):
        self.assertEqual(anio_de_base(self.base), 2026)

    def test_filtrar_meses(self):
        base, _ = leer_cobranza(self.ruta, self.cfg, ["Febrero"])
        self.assertEqual(list(base), ["Febrero"])

    def test_mes_inexistente_avisa(self):
        base, avisos = leer_cobranza(self.ruta, self.cfg, ["Febrero", "Marzo"])
        self.assertTrue(any("Marzo" in a for a in avisos))

    def test_libro_sin_hojas_mensuales(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = fixtures.libro_cobranza(Path(tmp) / "otro.xlsx", {"Resumen": []})
            with self.assertRaises(RuntimeError) as e:
                leer_cobranza(ruta, self.cfg)
            self.assertIn("hoja mensual", str(e.exception))


if __name__ == "__main__":
    unittest.main()


class TestAnioObligatorio(unittest.TestCase):
    """El CLI no corre sin año: apagar ese filtro en silencio costaba caro."""

    def setUp(self):
        self.cfg = cargar_config()
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def _base(self, filas):
        ruta = fixtures.libro_cobranza(Path(self.tmp.name) / "c.xlsx", {"Enero": filas})
        base, _ = leer_cobranza(ruta, self.cfg)
        return base

    def test_hoja_recien_armada_no_tiene_anio(self):
        """Clientes y cuotas cargados, ningun cobro todavia."""
        base = self._base([{"nombre": "Perez Juan", "contrato": 2161, "cuota": 15000}])
        self.assertEqual(anio_de_base(base), 0)

    def test_el_cli_corta_y_pide_el_anio(self):
        from clasificador.cli import _anio_de

        base = self._base([{"nombre": "Perez Juan", "contrato": 2161, "cuota": 15000}])
        args = argparse.Namespace(anio=None)
        with self.assertRaises(SystemExit) as e:
            _anio_de(args, base)
        self.assertIn("--anio", str(e.exception))

    def test_el_anio_a_mano_manda(self):
        from clasificador.cli import _anio_de

        base = self._base([{"nombre": "Perez Juan", "contrato": 2161, "cuota": 15000}])
        self.assertEqual(_anio_de(argparse.Namespace(anio=2026), base), 2026)

    def test_con_fechas_lo_deduce_solo(self):
        import datetime as dt

        base = self._base(
            [{"nombre": "Perez Juan", "contrato": 2161, "cuota": 15000,
              "fecha_pago": dt.datetime(2026, 1, 5), "monto_pago": 15000}]
        )
        self.assertEqual(_anio_de_cli(base), 2026)


def _anio_de_cli(base):
    from clasificador.cli import _anio_de

    return _anio_de(argparse.Namespace(anio=None), base)
