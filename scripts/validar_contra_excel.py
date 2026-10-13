"""Compara las reglas contra las categorias cargadas a mano en el Excel.

Uso: python3 scripts/validar_contra_excel.py ruta/al/Extractos_cobranzas.xlsx
"""

from __future__ import annotations

import collections
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from clasificador.excel import leer_movimientos  # noqa: E402
from clasificador.reglas import cargar_reglas  # noqa: E402


def main(ruta: str) -> int:
    juego = cargar_reglas()
    movimientos = leer_movimientos(ruta)
    difs = []
    for i, mov in enumerate(movimientos, start=2):
        esperado = mov.categoria
        obtenido = juego.clasificar(mov.detalle).categoria
        if esperado and esperado != obtenido:
            difs.append((i, mov.detalle, esperado, obtenido))

    total = len(movimientos)
    print(f"Movimientos: {total}")
    print(f"Coinciden:   {total - len(difs)} ({(total - len(difs)) / total:.2%})")
    print(f"Difieren:    {len(difs)}\n")
    for fila, detalle, esperado, obtenido in difs:
        print(f"  fila {fila:5d}  detalle={detalle!r}")
        print(f"      Excel  : {esperado}")
        print(f"      Reglas : {obtenido}")
    if difs:
        print("\nResumen por detalle:")
        for (d, e, o), n in collections.Counter(
            (d, e, o) for _, d, e, o in difs
        ).most_common():
            print(f"  {n}x {d!r}: {e} -> {o}")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        raise SystemExit(2)
    raise SystemExit(main(sys.argv[1]))
