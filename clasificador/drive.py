"""Lectura del .xlsx de extractos directamente desde Google Drive.

Por que esto existe: el archivo fuente ("Ingresos Bancarios Clientes") esta en
Drive como .xlsx, no como planilla nativa. La Sheets API no puede leerlo, pero
la Drive API lo descarga tal cual y openpyxl lo abre sin problema. Asi no hay
que convertir el archivo ni tocar el original, que suele ser de otra cuenta.
"""

from __future__ import annotations

import io
from pathlib import Path

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

from clasificador.sheets import _credenciales

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_PLANILLA_NATIVA = "application/vnd.google-apps.spreadsheet"


def descargar_xlsx(
    id_archivo: str,
    destino: str | Path,
    credenciales_json: str | Path | None = None,
) -> Path:
    """Baja el .xlsx de Drive a ``destino`` y devuelve la ruta.

    Si el archivo resulta ser una planilla nativa, la exporta como .xlsx, asi
    el resto del codigo no necesita saber cual de las dos cosas era.
    """
    credenciales = _credenciales(credenciales_json, leer_drive=True)
    api = build("drive", "v3", credentials=credenciales).files()

    try:
        meta = api.get(fileId=id_archivo, fields="name,mimeType").execute()
    except HttpError as error:
        if error.resp.status == 404:
            raise RuntimeError(
                f"Drive no encuentra el archivo {id_archivo}. Suele ser un "
                "problema de permisos: si usas una cuenta de servicio, el "
                "archivo tiene que estar compartido con su mail "
                "(alcanza con permiso de Lector)."
            ) from error
        raise

    mime = meta["mimeType"]
    if mime == MIME_PLANILLA_NATIVA:
        pedido = api.export_media(fileId=id_archivo, mimeType=MIME_XLSX)
    elif mime == MIME_XLSX:
        pedido = api.get_media(fileId=id_archivo)
    else:
        raise RuntimeError(
            f"El archivo {meta['name']!r} no es un .xlsx ni una planilla de "
            f"Google (mimeType: {mime})."
        )

    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    buffer = io.BytesIO()
    descarga = MediaIoBaseDownload(buffer, pedido)
    terminado = False
    while not terminado:
        _, terminado = descarga.next_chunk()
    destino.write_bytes(buffer.getvalue())
    return destino
