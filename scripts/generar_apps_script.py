"""Regenera apps_script/Codigo.gs inyectando reglas.json.

Asi la planilla clasifica con exactamente las mismas reglas que el codigo
Python. Despues de editar reglas.json, correr:

    python3 scripts/generar_apps_script.py
"""

from __future__ import annotations

import json
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
PLANTILLA = RAIZ / "apps_script" / "Codigo.plantilla.gs"
SALIDA = RAIZ / "apps_script" / "Codigo.gs"
MARCA = "/*__REGLAS__*/"


def main() -> int:
    reglas = json.loads((RAIZ / "reglas.json").read_text(encoding="utf-8"))
    bloque = "var REGLAS = " + json.dumps(reglas, ensure_ascii=False, indent=2) + ";"
    texto = PLANTILLA.read_text(encoding="utf-8")
    if MARCA not in texto:
        raise SystemExit(f"No encontre la marca {MARCA} en {PLANTILLA}")
    SALIDA.write_text(texto.replace(MARCA, bloque), encoding="utf-8")
    print(f"Generado {SALIDA.relative_to(RAIZ)} (reglas v{reglas['version']})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
