"""CLI: volcar el Excel a Google Sheets o a SharePoint, y clasificar."""

from __future__ import annotations

import argparse
import sys
import warnings

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

    pl = sub.add_parser(
        "volcar-carpeta",
        help="Escribe el extracto en una carpeta del disco (OneDrive sincronizado).",
    )
    origen_local = pl.add_mutually_exclusive_group(required=True)
    origen_local.add_argument("excel", nargs="?", help="Ruta a un .xlsx local.")
    origen_local.add_argument("--drive-id", metavar="ID", help="ID del .xlsx en Drive.")
    pl.add_argument(
        "--carpeta",
        required=True,
        help="Carpeta destino, normalmente la de SharePoint sincronizada con OneDrive.",
    )
    pl.add_argument("--nombre", default="Extractos_cobranzas.xlsx")
    pl.add_argument("--credenciales", help="JSON de la cuenta de servicio de Google.")
    pl.add_argument(
        "--formato", choices=["auto", "ingresos", "cobranzas"], default="auto"
    )
    pl.add_argument(
        "--reclasificar-todo",
        action="store_true",
        help="Ignora la categoria que ya trae el origen y la recalcula.",
    )
    pl.add_argument(
        "--desde",
        metavar="AAAA-MM-DD",
        help="Ignora los movimientos anteriores a esta fecha.",
    )
    pl.add_argument(
        "--hasta",
        metavar="AAAA-MM-DD",
        help="Ignora los movimientos posteriores a esta fecha.",
    )

    pp = sub.add_parser(
        "probar", help="Clasifica un texto de Detalle sin tocar nada (para probar reglas)."
    )
    pp.add_argument("detalle", nargs="+")

    # ------------------------------------------------------------------
    # Circuito de conciliacion (Modulos 1 a 4)
    # ------------------------------------------------------------------
    def agregar_config(parser):
        parser.add_argument(
            "--config",
            help="Ruta a conciliacion.json (por defecto, el del repo).",
        )
        return parser

    pe = agregar_config(
        sub.add_parser(
            "estandarizar",
            help="MODULO 1: unifica los extractos de todos los bancos en una tabla.",
        )
    )
    origen_e = pe.add_mutually_exclusive_group(required=True)
    origen_e.add_argument("ingresos", nargs="?", help="Libro de Ingresos Bancarios.")
    origen_e.add_argument("--drive-id", metavar="ID", help="El mismo libro, en Drive.")
    pe.add_argument("--credenciales", help="JSON de la cuenta de servicio de Google.")
    pe.add_argument(
        "--salida",
        default="Extracto_estandarizado.xlsx",
        help="Archivo .xlsx a escribir.",
    )
    pe.add_argument("--desde", metavar="AAAA-MM-DD")
    pe.add_argument("--hasta", metavar="AAAA-MM-DD")

    pco = agregar_config(
        sub.add_parser(
            "conciliar",
            help="MODULO 2: cruza el extracto con Cobranza General y completa los cobros.",
        )
    )
    origen_c = pco.add_mutually_exclusive_group(required=True)
    origen_c.add_argument("ingresos", nargs="?", help="Libro de Ingresos Bancarios.")
    origen_c.add_argument("--drive-id", metavar="ID")
    pco.add_argument("--cobranza", required=True, help="Cobranza General .xlsx")
    pco.add_argument("--credenciales")
    pco.add_argument(
        "--mes", action="append", help="Limitar a estos meses (repetible)."
    )
    pco.add_argument(
        "--salida",
        help=(
            "Archivo a escribir. Por defecto, una copia nueva al lado del "
            "original: la base no se toca si no lo pedis con --en-sitio."
        ),
    )
    pco.add_argument(
        "--en-sitio",
        action="store_true",
        help="Escribe sobre el mismo archivo de Cobranza General.",
    )
    pco.add_argument(
        "--reporte",
        default="Conciliacion.xlsx",
        help="Reporte de la corrida, con los pagos no identificados.",
    )
    pco.add_argument(
        "--simular",
        action="store_true",
        help="Muestra lo que haria y escribe solo el reporte, sin tocar la base.",
    )
    pco.add_argument(
        "--con-estado",
        action="store_true",
        help="Aprovecha la pasada para reescribir la columna Estado (Modulo 3).",
    )
    pco.add_argument(
        "--anio",
        type=int,
        metavar="AAAA",
        help=(
            "Año de las cuotas del libro. Por defecto se deduce de las fechas "
            "ya cargadas; hace falta pasarlo cuando la hoja esta recien "
            "armada y todavia no tiene ninguna."
        ),
    )

    pes = agregar_config(
        sub.add_parser(
            "estado",
            help="MODULO 3: reescribe la columna Estado con la regla 1/Parcial/0.",
        )
    )
    pes.add_argument("--cobranza", required=True)
    pes.add_argument("--salida")
    pes.add_argument("--en-sitio", action="store_true")
    pes.add_argument("--simular", action="store_true")

    pm = agregar_config(
        sub.add_parser("mora", help="MODULO 3: tablero de Mora Activa por cliente.")
    )
    pm.add_argument("--cobranza", required=True)
    pm.add_argument(
        "--corte", metavar="AAAA-MM-DD", help="Dia de corte. Por defecto, hoy."
    )
    pm.add_argument("--anio", type=int, metavar="AAAA", help="Año de las cuotas.")
    pm.add_argument("--salida", default="Mora_Activa.xlsx")

    pel = agregar_config(
        sub.add_parser(
            "exportar-listado",
            help="MODULO 4: arma las filas de la hoja H1 del Listado banco.",
        )
    )
    pel.add_argument("--cobranza", required=True)
    pel.add_argument(
        "--listado",
        required=True,
        help="Listado banco .xlsx, de donde se aprenden los codigos de cliente.",
    )
    pel.add_argument(
        "--ingresos",
        help=(
            "Libro de Ingresos Bancarios. Opcional, pero sin el los cobros por "
            "SIRO quedan sin codigo de canal (PF, RP, TQR...)."
        ),
    )
    pel.add_argument("--anio", type=int, metavar="AAAA", help="Año de las cuotas.")
    pel.add_argument("--desde", metavar="AAAA-MM-DD", help="Filtra por fecha de pago.")
    pel.add_argument("--hasta", metavar="AAAA-MM-DD")
    pel.add_argument(
        "--proceso", metavar="AAAA-MM-DD", help="Fecha de Proceso. Por defecto, hoy."
    )
    pel.add_argument("--salida", default="Listado_banco_nuevo.xlsx")
    pel.add_argument(
        "--anexar",
        action="store_true",
        help=(
            "Ademas de escribir --salida, pega las filas al final de la hoja H1 "
            "del --listado, conservando el archivo original intacto."
        ),
    )
    pel.add_argument(
        "--destino-anexado",
        help="Con --anexar: archivo a escribir. Por defecto, una copia nueva.",
    )
    pel.add_argument(
        "--solo-pagados",
        action="store_true",
        help="Deja afuera los pagos parciales (Estado = Parcial).",
    )
    pel.add_argument(
        "--sin-control-de-duplicados",
        action="store_true",
        help="No descarta los pagos que ya figuran en el Listado banco.",
    )
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




# ---------------------------------------------------------------------------
# Circuito de conciliacion (Modulos 1 a 4)
# ---------------------------------------------------------------------------


def _fecha_opcional(texto: str | None, flag: str):
    import datetime as dt

    if not texto:
        return None
    try:
        return dt.date.fromisoformat(texto)
    except ValueError:
        raise SystemExit(f"{flag} tiene que ser AAAA-MM-DD, no {texto!r}.")


def _salida_al_lado(original: str, sufijo: str) -> Path:
    """``Cobranza General.xlsx`` -> ``Cobranza General - conciliado.xlsx``.

    El original no se toca por defecto. Es una base con 1.119 comentarios y el
    trabajo de meses del equipo: que haya que pedir --en-sitio a proposito.
    """
    ruta = Path(original)
    return ruta.with_name(f"{ruta.stem} - {sufijo}{ruta.suffix}")


def _anio_de(args, base) -> int:
    """El año de las cuotas del libro, o corta con un mensaje claro.

    Nunca devuelve 0. Un año en 0 apagaria el filtro que separa los cobros de
    2025 de las cuotas de 2026, y lo haria en silencio: medido sobre el
    archivo real, 511 de 626 filas se llenaban con plata del año anterior. Es
    preferible frenar y pedir el dato.
    """
    from clasificador.cobranza import anio_de_base

    anio = getattr(args, "anio", None) or anio_de_base(base)
    if not anio:
        raise SystemExit(
            "No puedo deducir a que año corresponden las cuotas de este libro: "
            "ninguna hoja cargada tiene fechas de pago ni de vencimiento.\n"
            "Pasalo a mano para seguir, por ejemplo:  --anio 2026\n"
            "(sin el año, un cobro del año pasado se aplicaria a la cuota de "
            "este, porque el monto suele coincidir)."
        )
    return anio


def _leer_extracto_de(args, cfg):
    """Modulo 1 sobre un .xlsx local o uno de Drive."""
    from clasificador.extracto import leer_extracto, resumen

    ruta = getattr(args, "ingresos", None)
    if getattr(args, "drive_id", None):
        from tempfile import TemporaryDirectory

        from clasificador.drive import descargar_xlsx

        with TemporaryDirectory() as tmp:
            print(f"Descargando {args.drive_id} de Drive...")
            local = descargar_xlsx(
                args.drive_id, Path(tmp) / "ingresos.xlsx", args.credenciales
            )
            movimientos, avisos = leer_extracto(
                local,
                cfg.listado["codigos_de_canal"],
                cfg.cobranza["contrato_maximo"],
            )
    else:
        movimientos, avisos = leer_extracto(
            ruta, cfg.listado["codigos_de_canal"], cfg.cobranza["contrato_maximo"]
        )
    for aviso in avisos:
        print(f"  aviso: {aviso}")
    print(resumen(movimientos))
    return movimientos


def _ejecutar_estandarizar(args, cfg) -> int:
    from clasificador.extracto import escribir_extracto

    movimientos = _leer_extracto_de(args, cfg)
    desde = _fecha_opcional(args.desde, "--desde")
    hasta = _fecha_opcional(args.hasta, "--hasta")
    if desde:
        movimientos = [m for m in movimientos if m.fecha and m.fecha >= desde]
    if hasta:
        movimientos = [m for m in movimientos if m.fecha and m.fecha <= hasta]
    if desde or hasta:
        print(f"Filtrados por fecha: quedan {len(movimientos)}")
    ruta = escribir_extracto(movimientos, args.salida)
    print(f"\nEscrito: {ruta}")
    return 0


def _ejecutar_conciliar(args, cfg, juego) -> int:
    from clasificador.cobranza import leer_cobranza
    from clasificador.conciliacion import conciliar, escribir_reporte
    from clasificador.estado import celdas_de_estado
    from clasificador.xlsx_parche import parchear

    movimientos = _leer_extracto_de(args, cfg)
    base, avisos = leer_cobranza(args.cobranza, cfg, args.mes)
    for aviso in avisos:
        print(f"  aviso: {aviso}")
    anio = _anio_de(args, base)
    print(
        f"Base: {sum(len(v) for v in base.values())} filas de cliente en "
        f"{len(base)} hojas, año {anio}"
    )

    resultado = conciliar(movimientos, base, cfg, juego, anio=anio)
    print(
        f"\nFilas completadas: {len(resultado.aplicaciones)} "
        f"(con {resultado.movimientos_aplicados} movimientos)"
    )
    for nivel, n in sorted(resultado.por_nivel().items()):
        print(f"   {nivel:26} {n}")
    print(f"Ambiguos (revisar):   {len(resultado.ambiguos)}")
    print(f"No identificados:     {len(resultado.no_identificados)}")
    print(f"Pagares (carga manual): {len(resultado.pagares)}")
    if resultado.fuera_de_periodo:
        print(f"De otro año (ignorados): {resultado.fuera_de_periodo}")
    if resultado.fuera_de_alcance:
        print(f"De un mes fuera de esta pasada: {resultado.fuera_de_alcance}")

    ruta_reporte = escribir_reporte(resultado, args.reporte)
    print(f"\nReporte: {ruta_reporte}")

    celdas = resultado.celdas(cfg)
    if args.con_estado:
        celdas_estado, conteo = celdas_de_estado(base, cfg)
        celdas += celdas_estado
        print(
            f"Estado: {conteo['escritas']} formulas, "
            f"{conteo['bajas']} bajas respetadas"
        )

    if args.simular:
        print(f"\nSimulacion: se escribirian {len(celdas)} celdas. Nada se modifico.")
        return 0
    if not celdas:
        print("\nNo hay nada para escribir.")
        return 0

    destino = (
        Path(args.cobranza)
        if args.en_sitio
        else Path(args.salida or _salida_al_lado(args.cobranza, "conciliado"))
    )
    escritas = parchear(args.cobranza, destino, celdas)
    print(f"\nEscritas {sum(escritas.values())} celdas en {destino}")
    print("  (se conservan formulas, formatos y los comentarios de celda)")
    return 0


def _ejecutar_estado(args, cfg) -> int:
    from clasificador.cobranza import leer_cobranza
    from clasificador.estado import celdas_de_estado, resumen_estados
    from clasificador.xlsx_parche import parchear

    base, avisos = leer_cobranza(args.cobranza, cfg)
    for aviso in avisos:
        print(f"  aviso: {aviso}")
    celdas, conteo = celdas_de_estado(base, cfg)
    reparto = resumen_estados(base, cfg)
    print(f"Formulas a escribir: {conteo['escritas']}")
    print(f"  bajas respetadas (Estado = 3): {conteo['bajas']}")
    print(f"  anotaciones manuales respetadas: {conteo['manuales']}")
    print(
        f"Como queda el reparto: {reparto['pagado']} pagados, "
        f"{reparto['parcial']} parciales, {reparto['pendiente']} pendientes, "
        f"{reparto['baja']} bajas"
    )
    if args.simular:
        print("\nSimulacion: nada se modifico.")
        return 0
    destino = (
        Path(args.cobranza)
        if args.en_sitio
        else Path(args.salida or _salida_al_lado(args.cobranza, "estado"))
    )
    escritas = parchear(args.cobranza, destino, celdas)
    print(f"\nEscritas {sum(escritas.values())} celdas en {destino}")
    return 0


def _ejecutar_mora(args, cfg) -> int:
    from clasificador.cobranza import leer_cobranza
    from clasificador.mora import calcular_mora, escribir_tablero

    base, _ = leer_cobranza(args.cobranza, cfg)
    anio = _anio_de(args, base)
    corte = _fecha_opcional(args.corte, "--corte")
    clientes = calcular_mora(base, cfg, anio, corte)
    total = sum(c.total_vencido for c in clientes)
    cuotas = sum(c.cuotas for c in clientes)
    print(f"Clientes en mora: {len(clientes)}")
    print(f"Cuotas adeudadas: {cuotas}")
    print(f"Monto total vencido: {total:,.2f}")
    ruta = escribir_tablero(clientes, args.salida, corte)
    print(f"\nTablero: {ruta}")
    return 0


def _ejecutar_exportar_listado(args, cfg, juego) -> int:
    from clasificador.cobranza import leer_cobranza
    from clasificador.conciliacion import conciliar
    from clasificador.listado import (
        anexar_en_listado,
        aprender_codigos,
        armar_listado,
        escribir_listado,
        leer_ya_exportados,
    )

    base, _ = leer_cobranza(args.cobranza, cfg)
    anio = _anio_de(args, base)
    print(f"Base: {sum(len(v) for v in base.values())} filas, año {anio}")

    aplicaciones = None
    movimientos = None
    if args.ingresos:
        from clasificador.extracto import leer_extracto

        movimientos, _ = leer_extracto(
            args.ingresos,
            cfg.listado["codigos_de_canal"],
            cfg.cobranza["contrato_maximo"],
        )
        aplicaciones = conciliar(movimientos, base, cfg, juego, anio=anio).aplicaciones
        print(
            f"Extracto: {len(movimientos)} movimientos; {len(aplicaciones)} filas "
            "con canal y fecha de acreditacion del banco"
        )
    else:
        print(
            "Sin --ingresos: los cobros por SIRO van a quedar sin codigo de "
            "canal y marcados para revision."
        )

    print("Aprendiendo los codigos de cliente del Listado banco...")
    codigos = aprender_codigos(args.listado, cfg)
    print(f"  {len(codigos.prefijos)} contratos conocidos")

    ya = None
    if not args.sin_control_de_duplicados:
        ya = leer_ya_exportados(args.listado, cfg)
        print(f"  {len(ya)} pagos ya exportados, no se repiten")

    filas = armar_listado(
        base,
        cfg,
        codigos,
        proceso=_fecha_opcional(args.proceso, "--proceso"),
        desde=_fecha_opcional(args.desde, "--desde"),
        hasta=_fecha_opcional(args.hasta, "--hasta"),
        aplicaciones=aplicaciones,
        movimientos=movimientos,
        solo_pagados=args.solo_pagados,
        ya_exportados=ya,
    )
    from collections import Counter

    a_revisar = sum(1 for f in filas if f.revisar)
    print(f"\nFilas para exportar: {len(filas)}")
    print(f"  importe total: {sum(f.importe for f in filas):,.2f}")
    print("  de donde salio el codigo de cliente:")
    for procedencia, n in Counter(f.procedencia for f in filas).most_common():
        print(f"     {procedencia or 'sin dato':28} {n}")
    print(f"  marcadas para revisar: {a_revisar}")
    if not filas:
        print("No hay nada nuevo para exportar.")
        return 0

    ruta = escribir_listado(filas, args.salida, cfg)
    print(f"\nArchivo para copiar y pegar en H1: {ruta}")

    if args.anexar:
        destino = Path(
            args.destino_anexado or _salida_al_lado(args.listado, "con nuevas filas")
        )
        print("Anexando al final de la hoja H1 (puede tardar, el archivo es grande)...")
        desde_fila, hasta_fila = anexar_en_listado(
            filas, args.listado, destino, cfg
        )
        print(f"  filas {desde_fila} a {hasta_fila} escritas en {destino}")
    return 0


def _ejecutar(args) -> int:
    juego = cargar_reglas(args.reglas)

    if args.comando in (
        "estandarizar", "conciliar", "estado", "mora", "exportar-listado",
    ):
        from clasificador.config import cargar_config

        cfg = cargar_config(args.config)
        if args.comando == "estandarizar":
            return _ejecutar_estandarizar(args, cfg)
        if args.comando == "conciliar":
            return _ejecutar_conciliar(args, cfg, juego)
        if args.comando == "estado":
            return _ejecutar_estado(args, cfg)
        if args.comando == "mora":
            return _ejecutar_mora(args, cfg)
        return _ejecutar_exportar_listado(args, cfg, juego)

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

    if args.comando == "volcar-carpeta":
        from clasificador.libro import volcar_en_carpeta

        movimientos, _ = _leer_origen(args)
        movimientos = _filtrar_por_fecha(movimientos, args.desde, args.hasta)
        destino, agregados, previos = volcar_en_carpeta(
            args.carpeta,
            movimientos,
            juego,
            nombre=args.nombre,
            respetar_categoria_existente=not args.reclasificar_todo,
        )
        print(f"Agregados: {agregados} | Ya estaban: {previos}")
        print(f"Escrito en {destino}")
        if agregados:
            print("OneDrive lo sube solo en unos segundos.")
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


def _filtrar_por_fecha(movimientos, desde: str | None, hasta: str | None):
    """Acota el rango del origen.

    El libro de origen arranca en enero de 2025, mucho antes que el extracto
    que se viene usando. Sin esto, la primera corrida le mete al archivo todos
    los movimientos viejos de golpe.
    """
    import datetime as dt

    def _fecha(texto: str, flag: str) -> dt.date:
        try:
            return dt.date.fromisoformat(texto)
        except ValueError:
            raise SystemExit(f"{flag} tiene que ser AAAA-MM-DD, no {texto!r}.")

    if desde:
        limite = _fecha(desde, "--desde")
        movimientos = [m for m in movimientos if m.fecha and m.fecha >= limite]
    if hasta:
        limite = _fecha(hasta, "--hasta")
        movimientos = [m for m in movimientos if m.fecha and m.fecha <= limite]
    if desde or hasta:
        print(f"Filtrados por fecha: quedan {len(movimientos)} movimientos")
    return movimientos


def main(argv: list[str] | None = None) -> int:
    # Cobranza General tiene celdas de telefono con formato de fecha y un
    # numero imposible; openpyxl avisa una vez por celda y llena la pantalla de
    # ruido. El dato se lee igual (ver cobranza.py), asi que el aviso no aporta.
    warnings.filterwarnings("ignore", module="openpyxl")
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
