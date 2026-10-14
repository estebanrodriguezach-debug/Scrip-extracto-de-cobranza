"""Extraccion de datos del texto libre del banco y comparacion de nombres.

El concepto que informa cada banco es texto sin estructura, y cada banco
escribe distinto. Medido sobre el libro real (11.162 movimientos):

* Rio:       'Transferencia Recibida - De Carlos Alberto Sanchez / - Var / 20161815250'
* Galicia:   'Transferencia De Terceros MARIA GISELA RODRIGU'
* Macro:     'TRANSF 27380919724 VAR TRANSF'
* Patagonia: 'TRANSFERENCIA ENTRE CUENTAS VAR 27163187251 MARIA JULIA ESPINOSA'
* Roela:     'Milla Sandro Javier TRANSF.INTER.DISTINTO TIT.-23260942514-DANIELA ANAHI ABAN'
* Frances:   'TRANSFERENCIA 27363765403 MELINA NOELIA ROMERO'

De ahi salen dos necesidades: sacar el CUIT/CVU y el titular del texto, y
comparar nombres escritos de formas distintas ('Rosio Orellana' contra
'Orellana,rosio De Los').
"""

from __future__ import annotations

import difflib
import re

from clasificador.reglas import normalizar

# ---------------------------------------------------------------------------
# CUIT / CVU: se validan por digito verificador, no por forma
# ---------------------------------------------------------------------------
# Un contrato de 4 digitos pegado a un importe puede parecer un CUIT de 11.
# Validar el verificador descarta esos falsos positivos casi por completo:
# la probabilidad de que 11 digitos al azar pasen el chequeo es ~1/11.

_CANDIDATO_CUIT = re.compile(r"(?<!\d)(\d{2})[-\s.]?(\d{8})[-\s.]?(\d)(?!\d)")
_CANDIDATO_CBU = re.compile(r"(?<!\d)(\d{22})(?!\d)")

# Prefijos de CUIT/CUIL que usa AFIP. Sin esto, cualquier numero de 11 digitos
# con verificador valido pasaria (por ejemplo un telefono).
PREFIJOS_CUIT = {"20", "23", "24", "25", "26", "27", "30", "33", "34"}

_PESOS_CUIT = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)


def cuit_valido(digitos: str) -> bool:
    """Verifica el digito 11 de un CUIT/CUIL de 11 digitos."""
    if len(digitos) != 11 or not digitos.isdigit():
        return False
    if digitos[:2] not in PREFIJOS_CUIT:
        return False
    suma = sum(int(d) * p for d, p in zip(digitos[:10], _PESOS_CUIT))
    resto = suma % 11
    verificador = {0: 0, 1: 9}.get(resto, 11 - resto)
    return verificador == int(digitos[10])


def _digito_cbu(digitos: str, pesos: tuple[int, ...]) -> int:
    suma = sum(int(d) * p for d, p in zip(digitos, pesos))
    return (10 - suma % 10) % 10


def cbu_valido(digitos: str) -> bool:
    """Verifica los dos digitos de control de un CBU/CVU de 22 digitos.

    Sirve igual para un CVU (billetera virtual): la estructura y los
    verificadores son los mismos, solo cambia la entidad del primer bloque.
    """
    if len(digitos) != 22 or not digitos.isdigit():
        return False
    bloque1, bloque2 = digitos[:8], digitos[8:]
    if _digito_cbu(bloque1[:7], (7, 1, 3, 9, 7, 1, 3)) != int(bloque1[7]):
        return False
    return _digito_cbu(bloque2[:13], (3, 9, 7, 1, 3, 9, 7, 1, 3, 9, 7, 1, 3)) == int(
        bloque2[13]
    )


def extraer_cuit_cvu(texto: str) -> str:
    """Devuelve el primer CUIT o CVU valido del texto, o cadena vacia.

    Se prueba el CVU primero: un CVU de 22 digitos contiene adentro tiras de 11
    que podrian pasar por CUIT, y el dato mas especifico es el que conviene.
    """
    if not texto:
        return ""
    for m in _CANDIDATO_CBU.finditer(texto):
        if cbu_valido(m.group(1)):
            return m.group(1)
    for m in _CANDIDATO_CUIT.finditer(texto):
        candidato = m.group(1) + m.group(2) + m.group(3)
        if cuit_valido(candidato):
            return candidato
    return ""


# ---------------------------------------------------------------------------
# Numero de contrato
# ---------------------------------------------------------------------------
# En el libro real el 99,6% de los movimientos ya trae el contrato en su propia
# columna, asi que esto es el camino de excepcion: los 50 que no lo traen y los
# casos en que alguien lo anoto dentro del concepto.

_CONTRATO_CON_PALABRA = re.compile(
    r"(?:contrato|contr?\.?|cto|plan|legajo)\s*(?:n[roº°]?\.?|#|:)?\s*(\d{2,7})",
    re.IGNORECASE,
)
_SOLO_NUMERO = re.compile(r"^\s*(\d{2,7})\s*$")


def extraer_contrato(texto: str, maximo: int = 999999) -> tuple[str, str]:
    """Busca un numero de contrato en texto libre.

    Devuelve ``(contrato, procedencia)``, donde procedencia es ``palabra``
    (venia anunciado: "contrato 2161"), ``suelto`` (el texto era solo el
    numero) o ``""`` si no se encontro nada.

    No se aceptan numeros sueltos en medio de una frase a proposito: el
    concepto de un banco esta lleno de numeros que no son contratos (CUIT,
    numero de operacion, CBU). Un match equivocado aca mete una cobranza en la
    ficha de otro cliente, que es el peor error posible de este circuito.
    """
    if not texto:
        return "", ""
    m = _CONTRATO_CON_PALABRA.search(texto)
    if m and int(m.group(1)) <= maximo:
        return str(int(m.group(1))), "palabra"
    m = _SOLO_NUMERO.match(texto)
    if m and int(m.group(1)) <= maximo:
        return str(int(m.group(1))), "suelto"
    return "", ""


# ---------------------------------------------------------------------------
# Titular que origina el pago
# ---------------------------------------------------------------------------

# Texto que pone el banco y no es parte del nombre de nadie. Se saca antes de
# buscar el titular. El orden importa: primero las frases largas.
RUIDO_BANCARIO = [
    "transf recibida cvu dif titular",
    "transferencia entre cuentas",
    "transferencia inmediata coe",
    "credito transferencia coelsa",
    "transf.inter.distinto tit.",
    "credin - credito inmediato distinto titular",
    "transferencia de terceros",
    "transferencia recibida",
    "transferencia haberes",
    "transf. cliente cta.",
    "transferencia d/c",
    "transferencia pull",
    "cr trf inm coelsa",
    "mercado pago",
    "transferencia",
    "transf",
    "coelsa",
    "credin",
    "varios",
    "var",
    "hab",
    "cuo",
    "ninguna",
    "de",
]

_PALABRA = re.compile(r"[a-zñ]+", re.IGNORECASE)


def extraer_titular(texto: str) -> str:
    """Saca el nombre de la persona que transfirio, lo mejor que se pueda.

    Estrategia: normalizar, borrar el texto del banco y los numeros, y quedarse
    con la tira de palabras mas larga que sobreviva. Es deliberadamente simple
    porque el resultado no se usa para decidir solo: el nombre que vale es el
    de la columna del libro, y esto es el respaldo para cuando no esta.
    """
    if not texto:
        return ""
    limpio = normalizar(texto).replace("/", " ").replace("-", " ").replace(",", " ")
    limpio = re.sub(r"\d+", " ", limpio)
    for ruido in RUIDO_BANCARIO:
        limpio = re.sub(rf"(?<!\w){re.escape(ruido)}(?!\w)", " ", limpio)
    # El resto puede quedar partido en varios tramos; gana el mas largo.
    tramos = [t.strip() for t in re.split(r"\s{2,}", limpio) if t.strip()]
    if not tramos:
        palabras = _PALABRA.findall(limpio)
        return " ".join(palabras).strip()
    mejor = max(tramos, key=lambda t: (len(_PALABRA.findall(t)), len(t)))
    palabras = [p for p in mejor.split() if len(p) > 1]
    return " ".join(palabras)


# ---------------------------------------------------------------------------
# Comparacion de nombres
# ---------------------------------------------------------------------------


def clave_nombre(nombre: str) -> frozenset[str]:
    """Conjunto de palabras significativas, para comparar sin importar el orden.

    'Orellana,rosio De Los' y 'Rosio Orellana' tienen que dar parecido alto: en
    el libro real el mismo cliente aparece escrito de las dos formas. Se tiran
    las palabras de una letra y los conectores, que no distinguen a nadie.
    """
    conectores = {"de", "del", "la", "las", "los", "y", "da", "do", "san", "mc"}
    palabras = {
        p for p in _PALABRA.findall(normalizar(nombre)) if len(p) > 1 and p not in conectores
    }
    return frozenset(palabras)


def parecido_nombres(uno: str, otro: str) -> float:
    """Puntaje 0..1 entre dos nombres de persona.

    Combina dos miradas porque cada una falla distinto:

    * Jaccard sobre palabras: aguanta el orden invertido y los nombres de mas
      ('Perez Lucas Santiago' vs 'Lucas Santiago Perez'), pero no perdona un
      error de tipeo.
    * ``SequenceMatcher`` sobre el texto ordenado: aguanta el tipeo
      ('Sale Dania Luz' vs 'Sale,diana Luz'), pero se confunde con el orden.

    Se toma el mayor de los dos: basta que una de las dos miradas reconozca al
    cliente. El umbral de aceptacion lo pone ``conciliacion.json``.
    """
    a, b = clave_nombre(uno), clave_nombre(otro)
    if not a or not b:
        return 0.0
    jaccard = len(a & b) / len(a | b)
    secuencia = difflib.SequenceMatcher(
        None, " ".join(sorted(a)), " ".join(sorted(b))
    ).ratio()
    return max(jaccard, secuencia)
