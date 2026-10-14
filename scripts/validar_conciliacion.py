#!/usr/bin/env python3
"""Mide el motor de conciliacion contra el trabajo manual ya hecho.

Idea: las filas de ``Cobranza General`` que ya tienen fecha y monto cargados
son la verdad de referencia -- alguien las concilio a mano. El validador las
vuelve a poner "pendientes", corre el Modulo 2 sobre el extracto completo y
compara lo que el motor habria escrito contra lo que el equipo escribio.

    python3 scripts/validar_conciliacion.py INGRESOS.xlsx COBRANZA.xlsx [--mes Enero]

Tres resultados posibles por fila ya cargada:

  reproducida  el motor la llena con el mismo monto (y fecha compatible)
  distinta     el motor la llena con otro movimiento  <- lo que hay que mirar
  sin aplicar  el motor no encuentra con que llenarla (queda para revision)
"""

from __future__ import annotations

import argparse
import dataclasses
import sys
import warnings
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
warnings.simplefilter("ignore")

from clasificador.cobranza import anio_de_base, leer_cobranza
from clasificador.conciliacion import conciliar
from clasificador.config import cargar_config
from clasificador.extracto import leer_extracto
from clasificador.reglas import cargar_reglas


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("ingresos", help="Libro de Ingresos Bancarios.")
    p.add_argument("cobranza", help="Cobranza General.")
    p.add_argument("--mes", action="append", help="Limitar a estos meses.")
    p.add_argument("--tolerancia-dias", type=int, default=5)
    args = p.parse_args(argv)

    cfg = cargar_config()
    juego = cargar_reglas()

    movimientos, _ = leer_extracto(
        args.ingresos,
        cfg.listado["codigos_de_canal"],
        cfg.cobranza["contrato_maximo"],
    )
    base, _ = leer_cobranza(args.cobranza, cfg, args.mes)
    print(f"Extracto: {len(movimientos)} movimientos")
    print(f"Base:     {sum(len(v) for v in base.values())} filas de cliente")

    # Verdad de referencia y base "en blanco"
    verdad = {}
    vacia: dict[str, list] = {}
    for mes, filas in base.items():
        vacia[mes] = []
        for f in filas:
            if f.tiene_cobro and not f.es_baja:
                verdad[(f.hoja, f.fila)] = (f.fecha_pago, f.monto_pago, f.medio_pago)
                f = dataclasses.replace(
                    f, fecha_pago=None, texto_fecha_pago="", monto_pago=None,
                    medio_pago="",
                )
            vacia[mes].append(f)
    print(f"Filas ya conciliadas a mano (referencia): {len(verdad)}")
    anio = anio_de_base(base)
    print(f"Año del libro de cobranzas: {anio}\n")

    resultado = conciliar(movimientos, vacia, cfg, juego, anio=anio)

    aplicado_en = {(a.fila.hoja, a.fila.fila): a for a in resultado.aplicaciones}
    tolerancia = float(cfg.conciliacion["tolerancia_monto"])
    estado: Counter[str] = Counter()
    niveles: Counter[str] = Counter()
    ejemplos_distintos = []
    for clave, (fecha, monto, medio) in verdad.items():
        a = aplicado_en.get(clave)
        if a is None:
            estado["sin aplicar"] += 1
            continue
        mismo_monto = monto is not None and abs(a.monto - monto) <= tolerancia
        dias = (
            abs((a.movimiento.fecha - fecha).days)
            if (fecha and a.movimiento.fecha)
            else 0
        )
        if mismo_monto and dias <= args.tolerancia_dias:
            estado["reproducida"] += 1
            niveles[a.nivel] += 1
        elif mismo_monto:
            # Mismo importe, otra fecha: el cliente hizo dos pagos iguales en
            # el mes y el motor tomo el otro. El monto que se escribe es el
            # correcto, asi que no es un error de asignacion de plata.
            estado["mismo monto, otra fecha"] += 1
            niveles[a.nivel] += 1
        else:
            multiplo = (
                monto
                and a.monto
                and abs(a.monto / monto - round(a.monto / monto)) < 0.01
                and round(a.monto / monto) > 1
            )
            estado["otro monto, multiplo" if multiplo else "otro monto"] += 1
            if len(ejemplos_distintos) < 8:
                ejemplos_distintos.append(
                    (clave, monto, fecha, a.monto, a.movimiento.fecha, a.nivel)
                )

    total = sum(estado.values()) or 1
    print("Contra las filas conciliadas a mano:")
    for k in (
        "reproducida",
        "mismo monto, otra fecha",
        "otro monto, multiplo",
        "otro monto",
        "sin aplicar",
    ):
        print(f"  {k:24} {estado[k]:6}  ({100 * estado[k] / total:5.1f}%)")
    acertado = estado["reproducida"] + estado["mismo monto, otra fecha"]
    print(f"  {'-- monto correcto':24} {acertado:6}  ({100 * acertado / total:5.1f}%)")
    print("\nNivel con el que se reprodujeron:")
    for nivel, n in niveles.most_common():
        print(f"  {nivel:20} {n}")

    print("\nLa corrida completa, incluyendo filas que estaban impagas:")
    print(f"  filas completadas: {len(resultado.aplicaciones)}")
    print(f"  movimientos usados:{resultado.movimientos_aplicados}")
    for nivel, n in sorted(resultado.por_nivel().items()):
        print(f"     {nivel:20} {n}")
    print(f"  ambiguos:          {len(resultado.ambiguos)}")
    print(f"  no identificados:  {len(resultado.no_identificados)}")
    print(f"  pagares aparte:    {len(resultado.pagares)}")
    print(f"  de otro año:       {resultado.fuera_de_periodo}")
    motivos = Counter(p.motivo for p in resultado.no_identificados)
    for motivo, n in motivos.most_common():
        print(f"     {n:6}  {motivo}")

    if ejemplos_distintos:
        print("\nPrimeras diferencias (fila: manual vs motor):")
        for clave, m1, f1, m2, f2, nivel in ejemplos_distintos:
            print(
                f"  {clave[0]} fila {clave[1]}: manual {m1} del {f1} | "
                f"motor {m2} del {f2} [{nivel}]"
            )
    return 0


if __name__ == "__main__":
    sys.exit(main())
