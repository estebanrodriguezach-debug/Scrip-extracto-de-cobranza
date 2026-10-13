"""Volcado y clasificacion de movimientos en una planilla de Google Sheets.

Autenticacion, en este orden:

1. Ruta explicita a un JSON de cuenta de servicio (``--credenciales``).
2. ``GOOGLE_APPLICATION_CREDENTIALS``, si apunta a un JSON de cuenta de servicio.
3. Credenciales por defecto del entorno (ADC). Es lo que usa Cloud Shell: ya
   estas logueado con tu usuario, no hace falta ninguna clave.

Ver README.
"""

from __future__ import annotations

import os
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

import google.auth
from google.oauth2 import service_account
from googleapiclient.discovery import build

from clasificador.excel import ENCABEZADOS, Movimiento
from clasificador.reglas import JuegoDeReglas

# ``drive.file`` solo da acceso a los archivos creados por esta app: alcanza
# para crear la planilla destino y no pide permiso sobre todo el Drive.
ALCANCES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

# Leer un .xlsx que vive en Drive y que NO creo esta app (el que le comparten a
# la cuenta de servicio) necesita este alcance aparte. Se pide solo cuando hace
# falta: para una cuenta de servicio alcanza unicamente los archivos que le
# compartieron, pero con tu usuario cubriria todo tu Drive.
ALCANCE_LECTURA_DRIVE = "https://www.googleapis.com/auth/drive.readonly"

HOJA_DATOS = "Cobranzas"
HOJA_LISTAS = "Listas"
COL_CATEGORIA = 6  # F, 1-based
COL_DETALLE = 5  # E, 1-based


def _credenciales(
    ruta_cuenta_servicio: str | Path | None = None,
    leer_drive: bool = False,
):
    alcances = list(ALCANCES) + ([ALCANCE_LECTURA_DRIVE] if leer_drive else [])
    ruta = ruta_cuenta_servicio or os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if ruta:
        return service_account.Credentials.from_service_account_file(
            str(ruta), scopes=alcances
        )
    try:
        credenciales, _ = google.auth.default(scopes=alcances)
    except google.auth.exceptions.DefaultCredentialsError as error:
        raise RuntimeError(
            "No encontre credenciales. Opciones:\n"
            "  - En Cloud Shell o en tu maquina: "
            "gcloud auth application-default login --scopes="
            + ",".join(alcances)
            + ",https://www.googleapis.com/auth/cloud-platform\n"
            "  - O pasa --credenciales con el JSON de una cuenta de servicio."
        ) from error
    return credenciales


def crear_planilla(
    titulo: str,
    credenciales_json: str | Path | None = None,
    compartir_con: str | None = None,
) -> tuple[str, str]:
    """Crea una planilla vacia y devuelve (id, url).

    ``compartir_con``: mail al que darle permiso de edicion. Es obligatorio en
    la practica cuando se usa una cuenta de servicio: la planilla nace en el
    Drive de la cuenta de servicio, que no tiene interfaz web, asi que sin esto
    no la podrias abrir desde el navegador.
    """
    credenciales = _credenciales(credenciales_json)
    api = build("sheets", "v4", credentials=credenciales).spreadsheets()
    resp = api.create(
        body={"properties": {"title": titulo}},
        fields="spreadsheetId,spreadsheetUrl",
    ).execute()
    id_planilla, url = resp["spreadsheetId"], resp["spreadsheetUrl"]

    if compartir_con:
        build("drive", "v3", credentials=credenciales).permissions().create(
            fileId=id_planilla,
            body={"type": "user", "role": "writer", "emailAddress": compartir_con},
            sendNotificationEmail=False,
        ).execute()

    return id_planilla, url


def _clave_con_ocurrencia(movimientos: list[Movimiento]) -> list[str]:
    """Desambigua movimientos identicos (misma fecha, banco, importe y texto).

    En la base real hay 48 pares asi: son cobros repetidos legitimos, no
    duplicados a descartar. Se numeran para que el volcado sea idempotente sin
    perderlos.
    """
    vistos: Counter[str] = Counter()
    claves = []
    for mov in movimientos:
        vistos[mov.clave] += 1
        claves.append(f"{mov.clave}#{vistos[mov.clave]}")
    return claves


@dataclass
class ResumenVolcado:
    agregados: int
    ya_existentes: int
    reclasificados: int


class PlanillaCobranzas:
    def __init__(
        self,
        id_planilla: str,
        juego: JuegoDeReglas,
        credenciales_json: str | Path | None = None,
        hoja: str = HOJA_DATOS,
    ) -> None:
        self.id_planilla = id_planilla
        self.juego = juego
        self.hoja = hoja
        self._api = build(
            "sheets", "v4", credentials=_credenciales(credenciales_json)
        ).spreadsheets()

    # --- lectura -----------------------------------------------------------
    def _valores(self, rango: str) -> list[list[object]]:
        resp = (
            self._api.values()
            .get(spreadsheetId=self.id_planilla, range=rango)
            .execute()
        )
        return resp.get("values", [])

    def _hojas_existentes(self) -> dict[str, int]:
        meta = self._api.get(spreadsheetId=self.id_planilla).execute()
        return {h["properties"]["title"]: h["properties"]["sheetId"] for h in meta["sheets"]}

    # --- preparacion -------------------------------------------------------
    def preparar(self) -> None:
        """Crea las hojas, el encabezado y el desplegable de categorias."""
        existentes = self._hojas_existentes()
        pedidos = [
            {"addSheet": {"properties": {"title": t}}}
            for t in (self.hoja, HOJA_LISTAS)
            if t not in existentes
        ]
        if pedidos:
            self._api.batchUpdate(
                spreadsheetId=self.id_planilla, body={"requests": pedidos}
            ).execute()
            existentes = self._hojas_existentes()

        self._api.values().update(
            spreadsheetId=self.id_planilla,
            range=f"{self.hoja}!A1:F1",
            valueInputOption="RAW",
            body={"values": [ENCABEZADOS]},
        ).execute()
        self._api.values().update(
            spreadsheetId=self.id_planilla,
            range=f"{HOJA_LISTAS}!A1",
            valueInputOption="RAW",
            body={"values": [[c] for c in self.juego.categorias_validas]},
        ).execute()

        id_hoja = existentes[self.hoja]
        self._api.batchUpdate(
            spreadsheetId=self.id_planilla,
            body={
                "requests": [
                    {
                        "setDataValidation": {
                            "range": {
                                "sheetId": id_hoja,
                                "startRowIndex": 1,
                                "startColumnIndex": COL_CATEGORIA - 1,
                                "endColumnIndex": COL_CATEGORIA,
                            },
                            "rule": {
                                "condition": {
                                    "type": "ONE_OF_LIST",
                                    "values": [
                                        {"userEnteredValue": c}
                                        for c in self.juego.categorias_validas
                                    ],
                                },
                                "showCustomUi": True,
                                "strict": False,
                            },
                        }
                    },
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": id_hoja,
                                "gridProperties": {"frozenRowCount": 1},
                            },
                            "fields": "gridProperties.frozenRowCount",
                        }
                    },
                ]
            },
        ).execute()

    # --- volcado -----------------------------------------------------------
    def volcar(
        self, movimientos: list[Movimiento], respetar_categoria_existente: bool = True
    ) -> ResumenVolcado:
        """Agrega a la planilla los movimientos que todavia no estan.

        ``respetar_categoria_existente``: si el Excel ya trae una categoria
        cargada a mano, se conserva en lugar de recalcularla.
        """
        actuales = self._valores(f"{self.hoja}!A2:F")
        presentes = set(_clave_con_ocurrencia(_desde_filas(actuales)))

        nuevas_filas = []
        ya = 0
        for mov, clave in zip(movimientos, _clave_con_ocurrencia(movimientos)):
            if clave in presentes:
                ya += 1
                continue
            if not (respetar_categoria_existente and mov.categoria):
                mov.categoria = self.juego.clasificar(mov.detalle).categoria
            nuevas_filas.append(mov.a_fila())

        if nuevas_filas:
            self._api.values().append(
                spreadsheetId=self.id_planilla,
                range=f"{self.hoja}!A1",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={"values": nuevas_filas},
            ).execute()

        return ResumenVolcado(len(nuevas_filas), ya, 0)

    # --- reclasificacion ---------------------------------------------------
    def reclasificar(self, solo_vacias: bool = True) -> ResumenVolcado:
        """Recalcula la columna Categoria sobre lo que ya esta en la planilla.

        ``solo_vacias=True`` toca unicamente las filas sin categoria, para no
        pisar correcciones manuales.
        """
        filas = self._valores(f"{self.hoja}!A2:F")
        actualizaciones = []
        for i, fila in enumerate(filas):
            fila = list(fila) + [""] * (6 - len(fila))
            if all(v in ("", None) for v in fila):
                continue
            actual = str(fila[COL_CATEGORIA - 1] or "").strip()
            if solo_vacias and actual:
                continue
            nueva = self.juego.clasificar(fila[COL_DETALLE - 1]).categoria
            if nueva != actual:
                actualizaciones.append(
                    {"range": f"{self.hoja}!F{i + 2}", "values": [[nueva]]}
                )

        if actualizaciones:
            self._api.values().batchUpdate(
                spreadsheetId=self.id_planilla,
                body={"valueInputOption": "RAW", "data": actualizaciones},
            ).execute()
        return ResumenVolcado(0, len(filas) - len(actualizaciones), len(actualizaciones))


def _desde_filas(filas: list[list[object]]) -> list[Movimiento]:
    from clasificador.excel import _a_fecha, _a_importe, _a_texto

    movimientos = []
    for fila in filas:
        fila = list(fila) + [""] * (6 - len(fila))
        if all(v in ("", None) for v in fila):
            continue
        movimientos.append(
            Movimiento(
                fecha=_a_fecha(fila[0] or None),
                banco=_a_texto(fila[1]),
                importe=_a_importe(fila[2] or None),
                extracto=_a_texto(fila[3]),
                detalle=_a_texto(fila[4]),
                categoria=_a_texto(fila[5]),
            )
        )
    return movimientos
