"""CUIT/CVU, contrato y titular sacados del texto del banco."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from clasificador.patrones import (  # noqa: E402
    cbu_valido,
    clave_nombre,
    cuit_valido,
    extraer_contrato,
    extraer_cuit_cvu,
    extraer_titular,
    parecido_nombres,
)


class TestCuit(unittest.TestCase):
    def test_cuits_reales_del_libro(self):
        for cuit in ("20161815250", "27380919724", "23286244424", "20234451546"):
            with self.subTest(cuit=cuit):
                self.assertTrue(cuit_valido(cuit))

    def test_digito_verificador_mal(self):
        self.assertFalse(cuit_valido("20161815251"))

    def test_prefijo_invalido(self):
        # 11 digitos con verificador que cierra, pero ningun CUIT empieza en 99.
        self.assertFalse(cuit_valido("99161815250"))

    def test_largo_distinto(self):
        self.assertFalse(cuit_valido("2016181525"))
        self.assertFalse(cuit_valido(""))


class TestExtraerCuit(unittest.TestCase):
    def test_conceptos_de_cada_banco(self):
        casos = {
            "Transferencia Recibida - De Carlos Alberto Sanchez / - Var / 20161815250": "20161815250",
            "TRANSF 27380919724 VAR TRANSF": "27380919724",
            "TRANSFERENCIA ENTRE CUENTAS VAR 27163187251 MARIA JULIA ESPINOSA": "27163187251",
            "Transferencia De Terceros MARIA GISELA RODRIGU": "",
        }
        for texto, esperado in casos.items():
            with self.subTest(texto=texto[:30]):
                self.assertEqual(extraer_cuit_cvu(texto), esperado)

    def test_no_confunde_un_importe_con_un_cuit(self):
        self.assertEqual(extraer_cuit_cvu("contrato 2161 importe 150000000"), "")

    def test_cvu_gana_sobre_cuit(self):
        # Un CBU/CVU valido contiene adentro tiras de 11 digitos; tiene que
        # ganar el dato mas especifico.
        cbu = "2850590940090418135201"
        self.assertEqual(extraer_cuit_cvu(f"Transf CVU {cbu} de Juan"), cbu)


class TestCbu(unittest.TestCase):
    """Los dos digitos de control de un CBU/CVU.

    Los casos de referencia son los de python-stdnum (stdnum.ar.cbu), que toma
    el primero como valido y el segundo como checksum fallado.
    """

    def test_cbu_valido(self):
        self.assertTrue(cbu_valido("2850590940090418135201"))

    def test_cbu_con_checksum_fallado(self):
        self.assertFalse(cbu_valido("2810590940090418135201"))

    def test_largo_distinto(self):
        self.assertFalse(cbu_valido("285059094009041813520"))
        self.assertFalse(cbu_valido("no son digitos 123456"))


class TestContrato(unittest.TestCase):
    def test_anunciado_con_palabra(self):
        self.assertEqual(extraer_contrato("contrato 2161"), ("2161", "palabra"))
        self.assertEqual(extraer_contrato("CTO N° 886 pago"), ("886", "palabra"))

    def test_texto_que_es_solo_el_numero(self):
        self.assertEqual(extraer_contrato("3499"), ("3499", "suelto"))

    def test_numero_suelto_en_una_frase_no_cuenta(self):
        # Clave: un match equivocado manda la cobranza a otro cliente.
        self.assertEqual(extraer_contrato("TRANSF 27380919724 VAR"), ("", ""))
        self.assertEqual(extraer_contrato("transf 2161 var"), ("", ""))

    def test_respeta_el_maximo(self):
        self.assertEqual(extraer_contrato("contrato 9999999999"), ("", ""))


class TestTitular(unittest.TestCase):
    def test_saca_el_nombre_de_cada_formato(self):
        casos = {
            "Transferencia Recibida - De Carlos Alberto Sanchez / - Var / 20161815250": "carlos alberto sanchez",
            "Transferencia De Terceros MARIA GISELA RODRIGU": "maria gisela rodrigu",
            "TRANSFERENCIA ENTRE CUENTAS VAR 27163187251 MARIA JULIA ESPINOSA": "maria julia espinosa",
            "Credito Transferencia Coelsa Lucas Santiago Perez": "lucas santiago perez",
        }
        for texto, esperado in casos.items():
            with self.subTest(texto=texto[:30]):
                self.assertEqual(extraer_titular(texto), esperado)

    def test_texto_vacio(self):
        self.assertEqual(extraer_titular(""), "")


class TestParecido(unittest.TestCase):
    def test_mismo_cliente_escrito_distinto(self):
        """Casos reales del libro: orden invertido y tipeo."""
        self.assertGreaterEqual(parecido_nombres("Rosio Orellana", "Orellana,rosio De Los"), 0.88)
        self.assertGreaterEqual(parecido_nombres("Sale Dania Luz", "Sale,diana Luz"), 0.88)
        self.assertGreaterEqual(parecido_nombres("Perez Lucas Santiago", "Lucas Santiago Perez"), 0.88)

    def test_clientes_distintos(self):
        self.assertLess(parecido_nombres("Juan Perez", "Maria Gomez"), 0.5)
        self.assertLess(parecido_nombres("Gomez Maria Griselda", "Gomez Maria Belen"), 0.88)

    def test_conectores_no_cuentan(self):
        self.assertEqual(clave_nombre("Fuentes de la Nieves"), frozenset({"fuentes", "nieves"}))

    def test_nombre_vacio(self):
        self.assertEqual(parecido_nombres("", "Juan Perez"), 0.0)


if __name__ == "__main__":
    unittest.main()
