"""CLI: volcar el Excel a Google Sheets o a SharePoint, y clasificar."""

from __future__ import annotations

import argparse
import sys

from pathlib import Path

from clasificador.reglas import cargar_reglas

# El nombre del archivo vive en sharepoint.py, pero no se importa el modulo aca:
# traeria msal y requests en cada invocacion del CLI, incluso para 'probar'.
NOMBRE_POR_DEFECTO = "Extractos_cobranzas.xlsx"


def _parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="cobranzas",
        description=(
            "Vuelca el extracto de cobranzas a Google Sheets o a SharePoint y "
            "lo clasifica."
        ),
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

    ps = sub.add_parser(
        "volcar-sharepoint",
        help="Igual que 'volcar', pero el destino es una carpeta de SharePoint.",
    )
    origen_sp = ps.add_mutually_exclusive_group(required=True)
    origen_sp.add_argument("excel", nargs="?", help="Ruta a un .xlsx local.")
    origen_sp.add_argument(
        "--drive-id",
        metavar="ID",
        help="ID del .xlsx en Google Drive, igual que en 'volcar'.",
    )
    destino_sp = ps.add_mutually_exclusive_group(required=True)
    destino_sp.add_argument(
        "--carpeta-url",
        metavar="URL",
        help=(
            "Link de la carpeta compartida de SharePoint (el que tiene ':f:'). "
            "Comodo, pero el link puede caducar: para dejarlo en un cron, corre "
            "una vez 'vincular-sharepoint' y despues usa --sitio."
        ),
    )
    destino_sp.add_argument(
        "--sitio",
        metavar="HOST:/sites/NOMBRE",
        help=(
            "Sitio por ruta, estable en el tiempo. Ej: "
            "contoso.sharepoint.com:/sites/Finanzas"
        ),
    )
    ps.add_argument(
        "--subcarpeta",
        default="",
        help="Carpeta dentro de la biblioteca, con --sitio. Ej: 'Extractos'.",
    )
    ps.add_argument(
        "--biblioteca",
        help=(
            "Biblioteca de documentos, con --sitio. Por defecto, la principal "
            "(se llama 'Documents' o 'Documentos' segun el idioma del tenant)."
        ),
    )
    ps.add_argument(
        "--nombre",
        default=NOMBRE_POR_DEFECTO,
        help="Nombre del .xlsx dentro de la carpeta.",
    )
    ps.add_argument(
        "--formato",
        choices=["auto", "ingresos", "cobranzas"],
        default="auto",
        help="Formato del .xlsx de origen. Por defecto se detecta solo.",
    )
    ps.add_argument("--credenciales", help="JSON de la cuenta de servicio de Google.")
    ps.add_argument(
        "--reclasificar-todo",
        action="store_true",
        help="Ignora la categoria que ya trae el Excel y la recalcula.",
    )

    pvs = sub.add_parser(
        "vincular-sharepoint",
        help="Resuelve un link de carpeta de SharePoint y muestra sus ids.",
    )
    pvs.add_argument("carpeta_url", help="Link de la carpeta compartida (':f:').")

    pp = sub.add_parser(
        "probar", help="Clasifica un texto de Detalle sin tocar nada (para probar reglas)."
    )
    pp.add_argument("detalle", nargs="+")
    return p


def _leer_origen(args) -> tuple[list, str]:
    """Lee los movimientos del origen, sea un .xlsx local o uno de Drive.

    Lo comparten 'volcar' y 'volcar-sharepoint': cambia el destino, no el origen.
    """
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
    return movimientos, origen


def _partir_sitio(sitio: str) -> tuple[str, str]:
    """'host:/sites/Nombre' -> ('host', '/sites/Nombre'). Valida antes de la red."""
    hostname, _, ruta_sitio = sitio.partition(":")
    if not hostname or not ruta_sitio:
        raise SystemExit(
            "--sitio va con el formato HOST:/sites/NOMBRE, por ejemplo\n"
            "  contoso.sharepoint.com:/sites/Finanzas"
        )
    return hostname, ruta_sitio


def _carpeta_sharepoint(args, cliente):
    """Resuelve el destino, por link compartido o por ruta de sitio."""
    from clasificador.sharepoint import resolver_carpeta, resolver_sitio

    if args.carpeta_url:
        return resolver_carpeta(cliente, args.carpeta_url)
    hostname, ruta_sitio = _partir_sitio(args.sitio)
    return resolver_sitio(
        cliente, hostname, ruta_sitio, args.subcarpeta, args.biblioteca
    )


def _ejecutar(args) -> int:
    juego = cargar_reglas(args.reglas)

    if args.comando == "probar":
        for texto in args.detalle:
            r = juego.clasificar(texto)
            print(f"{texto!r:50} -> {r.categoria}   [regla: {r.regla_id}]")
        return 0

    # imports perezosos: necesitan libs externas y red
    if args.comando == "vincular-sharepoint":
        from clasificador.sharepoint import ClienteGraph, resolver_carpeta

        carpeta = resolver_carpeta(ClienteGraph.desde_entorno(), args.carpeta_url)
        print(f"Carpeta: {carpeta.nombre}")
        print(f"  drive-id: {carpeta.drive_id}")
        print(f"  item-id : {carpeta.item_id}")
        print(f"  url     : {carpeta.url_web}")
        return 0

    if args.comando == "volcar-sharepoint":
        from clasificador.sharepoint import ClienteGraph, listar
        from clasificador.sharepoint import volcar as volcar_sharepoint

        # Primero el destino: si faltan credenciales o permisos, es mejor
        # enterarse antes de leer las 105 hojas del libro de origen.
        if args.sitio:
            _partir_sitio(args.sitio)
        cliente = ClienteGraph.desde_entorno()
        carpeta = _carpeta_sharepoint(args, cliente)
        print(
            f"Destino: {carpeta.nombre or 'raiz'} "
            f"({len(listar(cliente, carpeta))} archivos)"
        )
        movimientos, _ = _leer_origen(args)
        resumen = volcar_sharepoint(
            cliente,
            carpeta,
            movimientos,
            juego,
            nombre=args.nombre,
            respetar_categoria_existente=not args.reclasificar_todo,
        )
        print(f"Agregados: {resumen.agregados} | Ya estaban: {resumen.ya_existentes}")
        print(resumen.url or carpeta.url_web)
        return 0

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
        movimientos, _ = _leer_origen(args)
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


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        return _ejecutar(args)
    except (RuntimeError, OSError) as e:
        # Los mensajes de sharepoint.py ya vienen escritos para el que corre
        # esto, con el permiso que falta y quien lo concede. Un traceback de 14
        # lineas arriba los esconde y parece que el programa se rompio.
        print(e, file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
