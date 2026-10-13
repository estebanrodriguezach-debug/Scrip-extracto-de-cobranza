"""CLI: volcar el Excel a Google Sheets y clasificar."""

from __future__ import annotations

import argparse
import sys

from pathlib import Path

from clasificador.reglas import cargar_reglas


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="cobranzas",
        description="Vuelca el extracto de cobranzas a Google Sheets y lo clasifica.",
    )
    p.add_argument("--reglas", help="Ruta a reglas.json (por defecto, la del repo).")
    sub = p.add_subparsers(dest="comando", required=True)

    pv = sub.add_parser("volcar", help="Sube el Excel a la planilla y clasifica.")
    origen = pv.add_mutually_exclusive_group(required=True)
    origen.add_argument("excel", nargs="?", help="Ruta a un .xlsx local.")
    origen.add_argument(
        "--drive-id",
        metavar="ID",
        help=(
            "ID del .xlsx en Google Drive, en lugar de un archivo local. Es el "
            "tramo de la URL entre /d/ y /edit (o /view). El archivo tiene que "
            "estar compartido con quien corre esto (alcanza con Lector)."
        ),
    )
    pv.add_argument("--planilla", required=True, help="ID de la planilla de Google.")
    pv.add_argument("--credenciales", help="JSON de la cuenta de servicio.")
    pv.add_argument(
        "--hoja", default="Cobranzas", help="Hoja destino en la planilla de Google."
    )
    pv.add_argument(
        "--formato",
        choices=["auto", "ingresos", "cobranzas"],
        default="auto",
        help=(
            "Formato del .xlsx de origen. 'ingresos' es el libro de Ingresos "
            "Bancarios (una hoja por banco y mes); 'cobranzas' es el extracto "
            "ya armado. Por defecto se detecta solo."
        ),
    )
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

    pc = sub.add_parser("crear", help="Crea una planilla nueva y la deja preparada.")
    pc.add_argument(
        "--titulo", default="Extractos de cobranzas", help="Nombre de la planilla."
    )
    pc.add_argument("--credenciales")
    pc.add_argument(
        "--compartir-con",
        metavar="MAIL",
        help=(
            "Mail al que darle permiso de edicion. Obligatorio en la practica "
            "si usas una cuenta de servicio: si no, la planilla queda en el "
            "Drive de la cuenta de servicio y no la podes abrir."
        ),
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

    # imports perezosos: necesitan las libs de Google y red
    from clasificador.sheets import PlanillaCobranzas, crear_planilla

    if args.comando == "crear":
        id_planilla, url = crear_planilla(
            args.titulo, args.credenciales, args.compartir_con
        )
        PlanillaCobranzas(id_planilla, juego, args.credenciales).preparar()
        print(f"Planilla creada: {url}")
        print(f"ID: {id_planilla}")
        return 0

    planilla = PlanillaCobranzas(
        args.planilla, juego, args.credenciales, hoja=args.hoja
    )

    if args.comando == "volcar":
        from clasificador.origen import leer

        if args.drive_id:
            from tempfile import TemporaryDirectory

            from clasificador.drive import descargar_xlsx

            with TemporaryDirectory() as tmp:
                print(f"Descargando {args.drive_id} de Drive...")
                ruta = descargar_xlsx(
                    args.drive_id, Path(tmp) / "origen.xlsx", args.credenciales
                )
                movimientos, avisos = leer(ruta, args.formato)
            origen = f"Drive:{args.drive_id}"
        else:
            movimientos, avisos = leer(args.excel, args.formato)
            origen = args.excel
        for aviso in avisos:
            print(f"  aviso: {aviso}")
        print(f"Leidos {len(movimientos)} movimientos de {origen}")
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
