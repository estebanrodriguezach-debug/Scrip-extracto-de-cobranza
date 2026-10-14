"""El motor de conciliacion: los niveles, el uno a uno y lo que no toca."""

from __future__ import annotations

import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(RAIZ / "tests"))

import fixtures  # noqa: E402

from clasificador.cobranza import leer_cobranza  # noqa: E402
from clasificador.conciliacion import (  # noqa: E402
    NIVEL_1,
    NIVEL_1B,
    NIVEL_1C,
    NIVEL_2,
    SIN_CANDIDATO,
    TODAS_OCUPADAS,
    conciliar,
    escribir_reporte,
)
from clasificador.config import cargar_config  # noqa: E402
from clasificador.extracto import MovimientoBancario, leer_extracto  # noqa: E402
from clasificador.reglas import cargar_reglas  # noqa: E402


def _mov(**kw) -> MovimientoBancario:
    datos = dict(
        fecha=dt.date(2026, 1, 10),
        monto=15000.0,
        banco="Rio",
        titular_origen="",
        cuit_cvu="",
        contrato_detectado="",
        concepto="Transferencia Recibida",
        mes_origen="Enero",
        anio_origen=2026,
    )
    datos.update(kw)
    return MovimientoBancario(**datos)


class TestNiveles(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cfg = cargar_config()
        cls.juego = cargar_reglas()
        cls.tmp = tempfile.TemporaryDirectory()
        cls.ruta = fixtures.libro_cobranza(
            Path(cls.tmp.name) / "c.xlsx",
            {
                "Enero": [
                    {"nombre": "Perez Juan", "contrato": 2161, "cuota": 15000},
                    {"nombre": "Perez Juan", "contrato": 2161, "cuota": 15000},
                    {"nombre": "Gomez Maria Griselda", "contrato": 3607, "cuota": 40000},
                    {"nombre": "Lopez Ana", "contrato": 1000, "cuota": 25000},
                    {
                        "nombre": "Ya Pagado Pablo",
                        "contrato": 999,
                        "cuota": 10000,
                        "fecha_pago": dt.datetime(2026, 1, 2),
                        "monto_pago": 10000,
                        "medio_pago": "rio",
                    },
                    {
                        "nombre": "Baja Bruno",
                        "contrato": 777,
                        "cuota": 10000,
                        "estado": 3,
                        "fecha_pago": "baja",
                    },
                ]
            },
        )

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def _base(self):
        base, _ = leer_cobranza(self.ruta, self.cfg)
        return base

    def _conciliar(self, movimientos):
        return conciliar(movimientos, self._base(), self.cfg, self.juego, anio=2026)

    def test_nivel_1_contrato_y_monto(self):
        r = self._conciliar([_mov(contrato_detectado="2161", monto=15000.0)])
        self.assertEqual(len(r.aplicaciones), 1)
        a = r.aplicaciones[0]
        self.assertEqual(a.nivel, NIVEL_1)
        self.assertEqual(a.fila.contrato, "2161")
        self.assertEqual(a.monto, 15000.0)

    def test_nivel_1b_contrato_con_otro_monto(self):
        """Pago parcial: se aplica igual y el Modulo 3 lo marca Parcial."""
        r = self._conciliar([_mov(contrato_detectado="3607", monto=10000.0)])
        self.assertEqual(r.aplicaciones[0].nivel, NIVEL_1B)
        self.assertEqual(r.aplicaciones[0].diferencia, -30000.0)

    def test_nivel_1c_una_transferencia_paga_dos_cuotas(self):
        """30.000 cubre las dos cuotas de 15.000 del contrato 2161."""
        r = self._conciliar([_mov(contrato_detectado="2161", monto=30000.0)])
        self.assertEqual(len(r.aplicaciones), 2)
        self.assertEqual({a.nivel for a in r.aplicaciones}, {NIVEL_1C})
        # Cada fila recibe su cuota, no el total.
        self.assertEqual([a.monto for a in r.aplicaciones], [15000.0, 15000.0])
        self.assertEqual({a.total_del_grupo for a in r.aplicaciones}, {30000.0})
        self.assertEqual(r.movimientos_aplicados, 1)

    def test_nivel_2_por_nombre(self):
        r = self._conciliar(
            [_mov(titular_origen="MARIA GRISELDA GOMEZ", monto=40000.0)]
        )
        self.assertEqual(r.aplicaciones[0].nivel, NIVEL_2)
        self.assertEqual(r.aplicaciones[0].fila.contrato, "3607")

    def test_nivel_2_no_aplica_si_el_monto_no_coincide(self):
        r = self._conciliar(
            [_mov(titular_origen="MARIA GRISELDA GOMEZ", monto=12345.0)]
        )
        self.assertEqual(r.aplicaciones, [])

    def test_nivel_2_con_empate_no_aplica_nada(self):
        """Dos clientes parecidos: lo resuelve una persona, no el script."""
        with tempfile.TemporaryDirectory() as tmp:
            ruta = fixtures.libro_cobranza(
                Path(tmp) / "x.xlsx",
                {
                    "Enero": [
                        {"nombre": "Gomez Maria", "contrato": 11, "cuota": 5000},
                        {"nombre": "Gomez Maria", "contrato": 22, "cuota": 5000},
                    ]
                },
            )
            base, _ = leer_cobranza(ruta, self.cfg)
            r = conciliar(
                [_mov(titular_origen="Maria Gomez", monto=5000.0)],
                base, self.cfg, self.juego, anio=2026,
            )
        self.assertEqual(r.aplicaciones, [])
        self.assertEqual(len(r.ambiguos), 1)
        self.assertEqual(len(r.ambiguos[0].candidatos), 2)

    def test_uno_a_uno(self):
        """Dos cobros iguales van a las dos filas, no dos veces a la misma."""
        r = self._conciliar(
            [
                _mov(contrato_detectado="2161", monto=15000.0),
                _mov(contrato_detectado="2161", monto=15000.0),
            ]
        )
        self.assertEqual(len(r.aplicaciones), 2)
        self.assertNotEqual(
            r.aplicaciones[0].fila.fila, r.aplicaciones[1].fila.fila
        )

    def test_no_toca_una_fila_ya_cobrada(self):
        r = self._conciliar([_mov(contrato_detectado="999", monto=10000.0)])
        self.assertEqual(r.aplicaciones, [])
        self.assertEqual(len(r.no_identificados), 1)
        self.assertIn("no tiene ninguna fila libre", r.no_identificados[0].motivo)

    def test_no_toca_una_baja(self):
        r = self._conciliar([_mov(contrato_detectado="777", monto=10000.0)])
        self.assertEqual(r.aplicaciones, [])

    def test_descarta_otro_anio(self):
        """Un cobro de enero de 2025 no paga la cuota de enero de 2026."""
        r = self._conciliar(
            [_mov(contrato_detectado="2161", monto=15000.0, anio_origen=2025)]
        )
        self.assertEqual(r.aplicaciones, [])
        self.assertEqual(r.fuera_de_periodo, 1)

    def test_mes_fuera_de_la_pasada(self):
        r = self._conciliar(
            [_mov(contrato_detectado="2161", monto=15000.0, mes_origen="Agosto")]
        )
        self.assertEqual(r.aplicaciones, [])
        self.assertEqual(r.fuera_de_alcance, 1)
        self.assertEqual(r.no_identificados, [])

    def test_mes_vecino(self):
        """Un cobro archivado en febrero que paga una cuota de enero.

        El vecino solo entra si la hoja del mes propio TAMBIEN esta cargada y
        no tiene donde aplicarlo. Ver test_no_usa_vecinos_si_falta_el_propio.
        """
        with tempfile.TemporaryDirectory() as tmp:
            ruta = fixtures.libro_cobranza(
                Path(tmp) / "x.xlsx",
                {
                    "Enero": [{"nombre": "Perez Juan", "contrato": 2161, "cuota": 15000}],
                    "Febrero": [{"nombre": "Otro Cliente", "contrato": 999, "cuota": 1}],
                },
            )
            base, _ = leer_cobranza(ruta, self.cfg)
            r = conciliar(
                [_mov(contrato_detectado="2161", monto=15000.0, mes_origen="Febrero",
                      fecha=dt.date(2026, 2, 1))],
                base, self.cfg, self.juego, anio=2026,
            )
        self.assertEqual(len(r.aplicaciones), 1)
        self.assertEqual(r.aplicaciones[0].fila.hoja, "Enero")

    def test_no_usa_vecinos_si_falta_el_propio(self):
        """Regresion: con --mes, un cobro de otro mes NO entra por la ventana.

        El bug: _meses_a_probar filtraba contra las hojas cargadas, asi que
        corriendo con --mes Octubre el mes propio del movimiento desaparecia
        y la unica hoja cargada sobrevivia como "vecina". Como la cuota es la
        misma todos los meses, el monto coincidia y la fila salia con el
        nivel mas confiable. Sobre el archivo real: 243 de 249 filas de una
        corrida de octubre traian plata de septiembre.
        """
        r = self._conciliar(
            [
                _mov(
                    contrato_detectado="2161",
                    monto=15000.0,
                    mes_origen="Febrero",
                    fecha=dt.date(2026, 2, 1),
                )
            ]
        )
        self.assertEqual(r.aplicaciones, [])
        self.assertEqual(r.fuera_de_alcance, 1)

    def test_los_vecinos_lejanos_tampoco(self):
        for mes in ("Marzo", "Noviembre", "Diciembre"):
            with self.subTest(mes=mes):
                r = self._conciliar(
                    [_mov(contrato_detectado="2161", monto=15000.0, mes_origen=mes)]
                )
                self.assertEqual(r.aplicaciones, [])

    def test_un_movimiento_sin_anio_no_entra(self):
        """Regresion: anio_origen = 0 pasaba el filtro y se aplicaba igual."""
        r = self._conciliar(
            [_mov(contrato_detectado="2161", monto=15000.0, anio_origen=0)]
        )
        self.assertEqual(r.aplicaciones, [])
        self.assertEqual(r.fuera_de_periodo, 1)

    def test_motivo_cuando_el_contrato_existe_pero_esta_ocupado(self):
        """Regresion: el motivo se medía sobre el indice de pendientes.

        Un contrato con todas sus filas ya cobradas se reportaba como "ningun
        cliente coincide", que manda al operador a buscar un cliente que si
        existe. Sobre el archivo real, 478 de 512 no identificados salian con
        el motivo equivocado.
        """
        r = self._conciliar([_mov(contrato_detectado="999", monto=10000.0)])
        self.assertEqual(len(r.no_identificados), 1)
        self.assertEqual(r.no_identificados[0].motivo, TODAS_OCUPADAS)

    def test_motivo_cuando_el_contrato_no_existe(self):
        r = self._conciliar([_mov(contrato_detectado="123456", monto=10000.0)])
        self.assertEqual(r.no_identificados[0].motivo, SIN_CANDIDATO)

    def test_los_pagares_van_aparte(self):
        r = self._conciliar([_mov(contrato_detectado="2161", detalle="PAGARE")])
        self.assertEqual(r.aplicaciones, [])
        self.assertEqual(len(r.pagares), 1)
        self.assertEqual(r.no_identificados, [])

    def test_celdas_a_escribir(self):
        r = self._conciliar([_mov(contrato_detectado="2161", monto=15000.0)])
        celdas = r.celdas(self.cfg)
        self.assertEqual([c.columna for c in celdas], [7, 8, 9])
        self.assertEqual(celdas[0].valor, dt.date(2026, 1, 10))
        self.assertEqual(celdas[1].valor, 15000.0)
        self.assertEqual(celdas[2].valor, "rio")

    def test_medio_siro_cuando_hay_codigo_de_canal(self):
        r = self._conciliar(
            [_mov(contrato_detectado="2161", monto=15000.0, codigo_canal="PF")]
        )
        self.assertEqual(r.celdas(self.cfg)[2].valor, "siro")

    def test_determinista(self):
        movimientos = [
            _mov(contrato_detectado="2161", monto=15000.0),
            _mov(contrato_detectado="3607", monto=40000.0),
            _mov(contrato_detectado="1000", monto=25000.0),
        ]
        una = self._conciliar(movimientos)
        otra = self._conciliar(list(reversed(movimientos)))
        self.assertEqual(
            [(a.fila.fila, a.monto) for a in una.aplicaciones],
            [(a.fila.fila, a.monto) for a in otra.aplicaciones],
        )

    def test_reporte_se_escribe(self):
        r = self._conciliar(
            [
                _mov(contrato_detectado="2161", monto=15000.0),
                _mov(contrato_detectado="0000", monto=1.0),
            ]
        )
        with tempfile.TemporaryDirectory() as tmp:
            ruta = escribir_reporte(r, Path(tmp) / "rep.xlsx")
            self.assertTrue(ruta.exists())
            import openpyxl

            libro = openpyxl.load_workbook(ruta)
            self.assertEqual(
                libro.sheetnames,
                ["Resumen", "Aplicados", "Ambiguos", "No identificados", "Pagares"],
            )
            libro.close()


class TestExtremoAExtremo(unittest.TestCase):
    """Del libro de ingresos a las celdas, pasando por los dos modulos."""

    def test_circuito(self):
        cfg = cargar_config()
        juego = cargar_reglas()
        with tempfile.TemporaryDirectory() as tmp:
            # Varias filas a proposito: la deteccion de columnas de
            # origen.py pide al menos dos fechas en la columna para
            # reconocerla, igual que en una hoja de banco de verdad.
            ingresos = fixtures.libro_ingresos(
                Path(tmp) / "i.xlsx",
                {
                    "Rio Enero": [
                        {
                            "fecha": dt.datetime(2026, 1, 5),
                            "monto": 15000.0,
                            "contrato": 2161,
                            "nombre": "Perez Juan",
                            "concepto": "Transferencia Recibida - De Juan Perez / - Var / 20161815250",
                        },
                        {
                            "fecha": dt.datetime(2026, 1, 8),
                            "monto": 40000.0,
                            "contrato": 3607,
                            "nombre": "Gomez Maria Griselda",
                            "concepto": "Transferencia De Terceros MARIA GRISELDA GOMEZ",
                        },
                        {
                            "fecha": dt.datetime(2026, 1, 9),
                            "monto": 25000.0,
                            "contrato": 1000,
                            "nombre": "Lopez Ana",
                            "concepto": "Transferencia Recibida - De Ana Lopez / - Var /",
                        },
                    ]
                },
            )
            cobranza = fixtures.libro_cobranza(
                Path(tmp) / "c.xlsx",
                {
                    "Enero": [
                        {"nombre": "Perez Juan", "contrato": 2161, "cuota": 15000},
                        {"nombre": "Gomez Maria Griselda", "contrato": 3607, "cuota": 40000},
                        {"nombre": "Lopez Ana", "contrato": 1000, "cuota": 25000},
                    ]
                },
            )
            movimientos, _ = leer_extracto(
                ingresos, cfg.listado["codigos_de_canal"], cfg.cobranza["contrato_maximo"]
            )
            self.assertEqual(len(movimientos), 3)
            primero = next(m for m in movimientos if m.contrato_detectado == "2161")
            self.assertEqual(primero.cuit_cvu, "20161815250")
            self.assertEqual(primero.mes_origen, "Enero")
            self.assertEqual(primero.anio_origen, 2026)
            self.assertEqual(primero.medio_cobranza, "rio")

            base, _ = leer_cobranza(cobranza, cfg)
            r = conciliar(movimientos, base, cfg, juego, anio=2026)
            self.assertEqual(len(r.aplicaciones), 3)
            self.assertEqual({a.nivel for a in r.aplicaciones}, {NIVEL_1})


if __name__ == "__main__":
    unittest.main()
