"""CLI: volcar el Excel a Google Sheets y clasificar."""

from __future__ import annotations

import argparse
import sys

from clasificador.excel import leer_movimientos
from clasificador.reglas import cargar_reglas


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="cobranzas",
        description="Vuelca el extracto de cobranzas a Google Sheets y lo clasifica.",
    )
    p.add_argument("--reglas", help="Ruta a reglas.json (por defecto, la del repo).")
    sub = p.add_subparsers(dest="comando", required=True)

    pv = sub.add_parser("volcar", help="Sube el Excel a la planilla y clasifica.")
    pv.add_argument("excel", help="Ruta al .xlsx de extractos.")
    pv.add_argument("--planilla", required=True, help="ID de la planilla de Google.")
    pv.add_argument("--credenciales", help="JSON de la cuenta de servicio.")
    pv.add_argument("--hoja", default="Cobranzas")
    pv.add_argument(
        "--reclasificar-todo",
        action="store_true",
        help="Ignora la categoria que ya trae el Excel y la recalcula.",
    )
    pv.add_argument(
        "--sin-preparar",
        action="store_true",
        help="No toca encabezados ni el desplegable de categorias.",
    )

    pr = sub.add_parser(
        "reclasificar", help="Recalcula la categoria de lo que ya esta en la planilla."
    )
    pr.add_argument("--planilla", required=True)
    pr.add_argument("--credenciales")
    pr.add_argument("--hoja", default="Cobranzas")
    pr.add_argument(
        "--pisar-manuales",
        action="store_true",
        help="Tambien recalcula filas que ya tienen categoria cargada.",
    )

    pp = sub.add_parser(
        "probar", help="Clasifica un texto de Detalle sin tocar nada (para probar reglas)."
    )
    pp.add_argument("detalle", nargs="+")
    return p


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    juego = cargar_reglas(args.reglas)

    if args.comando == "probar":
        for texto in args.detalle:
            r = juego.clasificar(texto)
            print(f"{texto!r:50} -> {r.categoria}   [regla: {r.regla_id}]")
        return 0

    from clasificador.sheets import PlanillaCobranzas  # import perezoso: necesita red

    planilla = PlanillaCobranzas(
        args.planilla, juego, args.credenciales, hoja=args.hoja
    )

    if args.comando == "volcar":
        movimientos = leer_movimientos(args.excel)
        print(f"Leidos {len(movimientos)} movimientos de {args.excel}")
        if not args.sin_preparar:
            planilla.preparar()
        resumen = planilla.volcar(
            movimientos, respetar_categoria_existente=not args.reclasificar_todo
        )
        print(f"Agregados: {resumen.agregados} | Ya estaban: {resumen.ya_existentes}")
        return 0

    resumen = planilla.reclasificar(solo_vacias=not args.pisar_manuales)
    print(f"Reclasificadas: {resumen.reclasificados}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
