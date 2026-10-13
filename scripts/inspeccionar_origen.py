"""Muestra la estructura real de un .xlsx de origen, sin volcar nada.

Sirve para descubrir como esta armado el archivo antes de configurar el
mapeo de columnas: nombres de hojas, encabezados y primeras filas.

    python3 scripts/inspeccionar_origen.py --drive-id 1IY5...
    python3 scripts/inspeccionar_origen.py archivo.xlsx

No imprime datos de clientes completos: recorta los textos largos, porque la
salida se suele pegar en un chat.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl  # noqa: E402

LARGO_MAX = 45
FILAS = 6


def _recortar(valor: object) -> str:
    if valor is None:
        return "-"
    texto = str(valor).replace("\n", " ").strip()
    return texto[:LARGO_MAX] + "..." if len(texto) > LARGO_MAX else texto


def inspeccionar(ruta: str | Path) -> None:
    libro = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    try:
        print(f"HOJAS ({len(libro.sheetnames)}): {libro.sheetnames}\n")
        for ws in libro.worksheets:
            print("=" * 72)
            print(f"HOJA: {ws.title}   filas={ws.max_row}  columnas={ws.max_column}")
            columnas = min(ws.max_column or 1, 12)
            letras = [
                openpyxl.utils.get_column_letter(i) for i in range(1, columnas + 1)
            ]
            print("      " + " | ".join(f"{l:^20}" for l in letras))
            for n, fila in enumerate(
                ws.iter_rows(
                    min_row=1, max_row=FILAS, max_col=columnas, values_only=True
                ),
                1,
            ):
                celdas = [f"{_recortar(v):<20.20}" for v in fila]
                print(f"  {n:3d} " + " | ".join(celdas))
            print()
    finally:
        libro.close()


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    grupo = p.add_mutually_exclusive_group(required=True)
    grupo.add_argument("excel", nargs="?", help="Ruta a un .xlsx local.")
    grupo.add_argument("--drive-id", help="ID del .xlsx en Google Drive.")
    p.add_argument("--credenciales", help="JSON de la cuenta de servicio.")
    args = p.parse_args()

    if args.drive_id:
        from clasificador.drive import descargar_xlsx

        with TemporaryDirectory() as tmp:
            print(f"Descargando {args.drive_id} de Drive...\n")
            ruta = descargar_xlsx(
                args.drive_id, Path(tmp) / "origen.xlsx", args.credenciales
            )
            inspeccionar(ruta)
    else:
        inspeccionar(args.excel)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
