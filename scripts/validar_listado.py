#!/usr/bin/env python3
"""Mide el exportador (Modulo 4) contra el Listado banco real.

La hoja H1 ya tiene 5.040 filas de 2026 cargadas a mano: son la verdad de
referencia. El validador arma la exportacion desde Cobranza General y compara,
para los pagos que aparecen en los dos lados (mismo importe y misma fecha de
pago), si el codigo ``Cliente`` y el ``Canal`` que genera el script son los
que el equipo cargo.

    python3 scripts/validar_listado.py INGRESOS.xlsx COBRANZA.xlsx LISTADO.xlsx

Importante: el historial de prefijos se aprende del **mismo** archivo contra
el que se mide. No es trampa -- es como va a funcionar en produccion, porque
el prefijo de un contrato lo define el sistema destino y no hay de donde
sacarlo mas que de ahi. Lo que se mide es si la regla de armado (prefijo +
contrato en 7 digitos + sufijo) reproduce el codigo, y cuanto aporta cada
fuente. La fila "contrato sin historial" muestra el caso de un contrato nuevo,
que es donde la regla tiene que adivinar el prefijo.
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
import warnings
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
warnings.simplefilter("ignore")

import openpyxl

from clasificador.cobranza import anio_de_base, leer_cobranza
from clasificador.conciliacion import conciliar
from clasificador.config import cargar_config
from clasificador.extracto import leer_extracto
from clasificador.listado import aprender_codigos, armar_listado
from clasificador.reglas import cargar_reglas


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("ingresos")
    p.add_argument("cobranza")
    p.add_argument("listado")
    args = p.parse_args(argv)

    cfg = cargar_config()
    juego = cargar_reglas()

    movimientos, _ = leer_extracto(
        args.ingresos, cfg.listado["codigos_de_canal"], cfg.cobranza["contrato_maximo"]
    )
    base, _ = leer_cobranza(args.cobranza, cfg)
    anio = anio_de_base(base)
    print(f"Extracto: {len(movimientos)} movimientos | base de {anio}")

    # El Modulo 2 aporta canal y fecha de acreditacion.
    resultado = conciliar(movimientos, base, cfg, juego, anio=anio)
    print(
        f"Conciliacion de apoyo: {len(resultado.aplicaciones)} filas con datos "
        "del extracto"
    )

    codigos = aprender_codigos(args.listado, cfg)
    print(f"Historial aprendido: {len(codigos.prefijos)} contratos, "
          f"{len(codigos.codigos)} codigos distintos")

    filas = armar_listado(
        base, cfg, codigos, proceso=dt.date(anio, 12, 31),
        aplicaciones=resultado.aplicaciones,
        movimientos=movimientos,
    )
    print(f"Exportacion armada: {len(filas)} filas\n")

    # Verdad de referencia: las filas de H1 del año de la base.
    libro = openpyxl.load_workbook(args.listado, data_only=True, read_only=True)
    reales: dict[tuple[float, dt.date], set[tuple[str, str]]] = defaultdict(set)
    por_contrato: dict[str, set[str]] = defaultdict(set)
    ancho = int(cfg.listado["ancho_contrato"])
    n_reales = 0
    for f in libro[cfg.listado["hoja"]].iter_rows(min_row=2, max_col=9, values_only=True):
        f = list(f) + [None] * (9 - len(f))
        cliente, importe, pago, canal = f[2], f[3], f[5], f[7]
        if cliente is None or importe is None or not isinstance(pago, dt.datetime):
            continue
        if pago.year != anio and not (pago.year == anio - 1 and pago.month == 12):
            continue
        n_reales += 1
        try:
            codigo = str(cliente).strip()
            reales[(round(float(importe), 2), pago.date())].add(
                (codigo, str(canal or "").strip().upper())
            )
            if len(codigo) == 9 and codigo.isdigit():
                por_contrato[str(int(codigo[1 : 1 + ancho]))].add(codigo)
        except (TypeError, ValueError):
            continue
    libro.close()
    print(f"Filas reales de H1 tomadas como referencia: {n_reales}")

    estado: Counter[str] = Counter()
    canal_ok = canal_mal = 0
    ejemplos = []
    sin_historial = 0
    sin_historial_ok = 0
    for fila in filas:
        if not fila.cliente or fila.fecha_pago is None:
            estado["sin codigo o sin fecha"] += 1
            continue
        candidatas = reales.get((round(fila.importe, 2), fila.fecha_pago))
        if not candidatas:
            estado["el pago no esta en H1"] += 1
            continue
        clientes = {c for c, _ in candidatas}
        if fila.cliente in clientes:
            estado["codigo correcto"] += 1
            canales = {k for c, k in candidatas if c == fila.cliente}
            if fila.canal and canales and fila.canal.upper() in canales:
                canal_ok += 1
            elif fila.canal:
                canal_mal += 1
                if len(ejemplos) < 6:
                    ejemplos.append(
                        f"canal: {fila.cliente} -> genero {fila.canal!r}, "
                        f"en H1 {sorted(canales)}"
                    )
        elif fila.cliente in por_contrato.get(fila.contrato, set()):
            # No es el codigo de ESTE pago, pero si uno que el cliente usa de
            # verdad: el prefijo esta bien y el sufijo es otro cupon suyo.
            estado["codigo que el cliente usa"] += 1
        else:
            reales_cliente = por_contrato.get(fila.contrato, set())
            if any(c[0] == fila.cliente[0] for c in reales_cliente):
                estado["prefijo correcto, sufijo nunca visto"] += 1
            elif reales_cliente:
                estado["prefijo distinto"] += 1
            else:
                estado["el contrato no esta en H1 2026"] += 1
            if len(ejemplos) < 6:
                ejemplos.append(
                    f"codigo: contrato {fila.contrato} -> genero "
                    f"{fila.cliente}, el cliente usa {sorted(reales_cliente)[:4]}"
                )
        if fila.contrato and fila.contrato not in codigos.prefijos:
            sin_historial += 1
            if fila.cliente in clientes:
                sin_historial_ok += 1

    claves = (
        "codigo correcto",
        "codigo que el cliente usa",
        "prefijo correcto, sufijo nunca visto",
        "prefijo distinto",
        "el contrato no esta en H1 2026",
    )
    comparables = sum(estado[k] for k in claves) or 1
    print("\nSobre las filas que se pudieron comparar contra H1:")
    for k in claves:
        print(f"  {k:38} {estado[k]:6}  ({100 * estado[k] / comparables:5.1f}%)")
    bien = estado["codigo correcto"] + estado["codigo que el cliente usa"]
    print(f"  {'-- codigo valido para ese cliente':38} {bien:6}  "
          f"({100 * bien / comparables:5.1f}%)")
    print(f"\n  {'(no comparables) el pago no esta en H1':38} "
          f"{estado['el pago no esta en H1']:6}")
    print(f"  {'(no comparables) sin codigo o sin fecha':38} "
          f"{estado['sin codigo o sin fecha']:6}")

    total_canal = canal_ok + canal_mal or 1
    print(f"\nCanal, sobre las filas con codigo correcto:")
    print(f"  coincide      {canal_ok:6}  ({100 * canal_ok / total_canal:5.1f}%)")
    print(f"  no coincide   {canal_mal:6}  ({100 * canal_mal / total_canal:5.1f}%)")

    if sin_historial:
        print(
            f"\nContratos sin historial en H1 (el caso dificil): {sin_historial}, "
            f"de los cuales el prefijo por defecto acerto {sin_historial_ok} "
            f"({100 * sin_historial_ok / sin_historial:.1f}%)"
        )

    if ejemplos:
        print("\nEjemplos de diferencias:")
        for e in ejemplos:
            print("  ", e)
    return 0


if __name__ == "__main__":
    sys.exit(main())
