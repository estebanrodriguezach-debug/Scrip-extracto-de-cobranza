"""Escritura quirurgica de celdas dentro de un .xlsx ya existente.

Por que existe este modulo en vez de usar openpyxl para escribir
--------------------------------------------------------------------
``Cobranza General`` no es una planilla descartable: es la base operativa, y
trae 1.119 comentarios de celda con anotaciones del equipo, mas el VML que los
dibuja. Abrirla con openpyxl y guardarla pierde partes del archivo. Medido
sobre el archivo real, un ``load_workbook`` + ``save`` descarta:

    xl/comments1..11.xml + sus _rels   (los 1.119 comentarios, reescritos en
                                        otro formato que Excel no lee igual)
    xl/drawings/vmlDrawing1..11.vml    (los globos de los comentarios)
    xl/commentsmeta0..9                (metadatos de comentarios con hilo)
    xl/metadata, xl/calcChain.xml

Perder eso es inaceptable: son datos que nadie puede reconstruir. Pandas es
peor todavia, porque solo sabe escribir archivos nuevos desde cero.

La salida es tratar el .xlsx por lo que es -- un zip de XMLs -- y reescribir
unicamente el XML de las hojas que cambian, tocando solo las celdas pedidas.
Todo lo demas se copia byte a byte. Se edita el XML como texto, no con un
parser: ``ElementTree`` renombra los prefijos de namespace al serializar, y
el atributo ``mc:Ignorable`` los nombra como texto, asi que un round-trip
"prolijo" rompe el archivo. El XML de una hoja lo genera Excel y es regular,
asi que la cirugia por posicion es segura y, sobre todo, no toca nada mas.

Lo que este modulo NO hace: crear hojas nuevas. Para eso estan los reportes,
que son archivos aparte (ver ``mora.py`` y ``listado.py``). Agregar una hoja
obliga a tocar ``workbook.xml``, los rels y ``[Content_Types].xml``, que es
justo el tipo de cirugia estructural que conviene no hacerle a la base.
"""

from __future__ import annotations

import datetime as dt
import re
import shutil
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from openpyxl.utils import get_column_letter

# Excel cuenta los dias desde el 1899-12-30 en el sistema 1900 (el desfasaje de
# un dia es el famoso bug del año 1900, que Excel conserva a proposito).
EPOCA_1900 = dt.date(1899, 12, 30)


@dataclass(frozen=True)
class Celda:
    """Una celda a escribir. ``formula`` manda sobre ``valor`` si viene."""

    hoja: str
    fila: int
    columna: int
    valor: object = None
    formula: str | None = None

    @property
    def ref(self) -> str:
        return f"{get_column_letter(self.columna)}{self.fila}"


def _escapar(texto: str) -> str:
    return (
        texto.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _serial_fecha(valor: dt.date) -> float:
    if isinstance(valor, dt.datetime):
        dias = (valor.date() - EPOCA_1900).days
        fraccion = (
            valor.hour * 3600 + valor.minute * 60 + valor.second
        ) / 86400
        return dias + fraccion
    return float((valor - EPOCA_1900).days)


def _contenido(celda: Celda, estilo: str) -> str:
    """Arma el elemento <c> completo para una celda."""
    s = f' s="{estilo}"' if estilo else ""
    if celda.formula is not None:
        # Sin <v>: el valor cacheado se deja afuera a proposito y Excel lo
        # calcula al abrir (ver _forzar_recalculo).
        return f'<c r="{celda.ref}"{s}><f>{_escapar(celda.formula)}</f></c>'
    valor = celda.valor
    if valor is None or valor == "":
        return f'<c r="{celda.ref}"{s}/>'
    if isinstance(valor, bool):
        return f'<c r="{celda.ref}"{s} t="b"><v>{int(valor)}</v></c>'
    if isinstance(valor, (dt.datetime, dt.date)):
        return f'<c r="{celda.ref}"{s}><v>{_serial_fecha(valor):g}</v></c>'
    if isinstance(valor, (int, float)):
        return f'<c r="{celda.ref}"{s}><v>{valor!r}</v></c>'
    texto = _escapar(str(valor))
    return f'<c r="{celda.ref}"{s} t="inlineStr"><is><t>{texto}</t></is></c>'


# numFmtId de formatos de fecha que Excel trae de fabrica. Los personalizados
# se detectan por su codigo (ver _estilos_de_fecha).
NUMFMT_FECHA = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))

_RE_COL = re.compile(r"([A-Z]+)")


def _es_formato_de_fecha(codigo: str) -> bool:
    """True si el codigo de formato dibuja una fecha u hora.

    Se sacan los literales entre comillas y lo que va entre corchetes (como
    ``[$-409]`` o ``[Red]``) antes de mirar si quedan letras de fecha: un
    formato de moneda puede traer la letra 'd' adentro de un literal.
    """
    limpio = re.sub(r'"[^"]*"', "", codigo)
    limpio = re.sub(r"\[[^\]]*\]", "", limpio)
    return any(c in limpio.lower() for c in "dmyhs")


def _estilos_de_fecha(styles_xml: str) -> set[str]:
    """Indices de estilo (el atributo ``s`` de una celda) con formato de fecha.

    Hace falta para no escribir una fecha en una celda cuyo formato la muestra
    como 46309. Paso con las filas vacias del final de la hoja H1 del Listado
    banco: ya vienen preformateadas, pero con formato general, asi que heredar
    el estilo de la celda dejaba el numero de serie a la vista.
    """
    personalizados = {
        int(m.group(1)): m.group(2)
        for m in re.finditer(
            r'<numFmt[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"', styles_xml
        )
    }
    bloque = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles_xml, re.DOTALL)
    if not bloque:
        return set()
    estilos = set()
    for i, xf in enumerate(re.findall(r"<xf\b[^>]*/?>", bloque.group(1))):
        m = re.search(r'numFmtId="(\d+)"', xf)
        if not m:
            continue
        numero = int(m.group(1))
        if numero in NUMFMT_FECHA:
            estilos.add(str(i))
        elif numero in personalizados and _es_formato_de_fecha(personalizados[numero]):
            estilos.add(str(i))
    return estilos


def _columna_de_ref(ref: str) -> int:
    letras = _RE_COL.match(ref)
    if not letras:
        return 0
    n = 0
    for ch in letras.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n


def _rango_fila(xml: str, fila: int) -> tuple[int, int, int] | None:
    """Ubica la fila: (inicio del <row>, fin del tag de apertura, fin del </row>).

    Si la fila viene auto-cerrada (``<row r="5"/>``, una fila sin celdas pero
    con formato) se la expande primero para poder meterle celdas adentro.
    """
    m = re.search(rf'<row[^>]*\sr="{fila}"[^>]*?(/?)>', xml)
    if not m:
        return None
    if m.group(1) == "/":
        return None  # el llamador la expande
    fin = xml.find("</row>", m.end())
    if fin == -1:
        return None
    return m.start(), m.end(), fin


def _expandir_fila_vacia(xml: str, fila: int) -> str:
    """``<row r="5" .../>`` -> ``<row r="5" ...></row>``."""
    return re.sub(
        rf'(<row[^>]*\sr="{fila}"[^>]*?)/>', r"\1></row>", xml, count=1
    )


def _crear_fila(xml: str, fila: int) -> str:
    """Agrega un ``<row>`` vacio en su lugar dentro de ``<sheetData>``.

    Excel exige que las filas vengan ordenadas por numero, asi que se inserta
    antes de la primera fila mayor. Hace falta para anexar al final de una
    hoja: la H1 del Listado banco real trae 40.505 filas ya creadas (vacias
    pero con formato), pero un libro armado de cero no tiene ninguna.
    """
    if re.search(r"<sheetData\s*/>", xml):
        xml = re.sub(r"<sheetData\s*/>", "<sheetData></sheetData>", xml, count=1)
    posicion = None
    for m in re.finditer(r'<row[^>]*\sr="(\d+)"', xml):
        if int(m.group(1)) > fila:
            posicion = m.start()
            break
    if posicion is None:
        posicion = xml.find("</sheetData>")
        if posicion == -1:
            raise ValueError("La hoja no tiene <sheetData>; no puedo agregar filas.")
    return xml[:posicion] + f'<row r="{fila}">' + "</row>" + xml[posicion:]


def _estilos_por_columna(xml: str) -> dict[int, str]:
    """Estilo mas usado en cada columna de la hoja.

    Cuando hay que crear una celda que no existia, hereda el estilo de su
    columna. Es lo que hace que una fecha escrita en una celda antes vacia se
    vea como fecha y no como 46.023, sin tener que tocar ``styles.xml``.

    Si la celda ya existe con su propio estilo, manda el de la celda (ver
    ``_escribir_en_hoja``): en la hoja H1 del Listado banco las filas vacias
    del final ya vienen preformateadas una por una, y respetar ese formato es
    lo que hace que las filas anexadas se vean igual que las de arriba.
    """
    conteo: dict[int, Counter[str]] = {}
    for ref, estilo in re.findall(r'<c r="([A-Z]+\d+)"[^>]*?\ss="(\d+)"', xml):
        conteo.setdefault(_columna_de_ref(ref), Counter())[estilo] += 1
    return {col: c.most_common(1)[0][0] for col, c in conteo.items()}


def _estilos_fecha_por_columna(
    xml: str, estilos_fecha: set[str]
) -> dict[int, str]:
    """Estilo de fecha mas usado en cada columna, si la columna tiene alguno."""
    conteo: dict[int, Counter[str]] = {}
    for ref, estilo in re.findall(r'<c r="([A-Z]+\d+)"[^>]*?\ss="(\d+)"', xml):
        if estilo in estilos_fecha:
            conteo.setdefault(_columna_de_ref(ref), Counter())[estilo] += 1
    return {col: c.most_common(1)[0][0] for col, c in conteo.items()}


def _escribir_en_hoja(
    xml: str, celdas: list[Celda], estilos_fecha: set[str] | None = None
) -> tuple[str, int]:
    """Aplica las celdas sobre el XML de una hoja. Devuelve (xml, escritas)."""
    estilos_fecha = estilos_fecha or set()
    estilos = _estilos_por_columna(xml)
    de_fecha = _estilos_fecha_por_columna(xml, estilos_fecha)
    escritas = 0

    def elegir(celda: Celda, propio: str | None) -> str:
        """Estilo para la celda: el propio, salvo que no sirva para una fecha."""
        candidato = propio or estilos.get(celda.columna, "")
        if isinstance(celda.valor, (dt.datetime, dt.date)) and not isinstance(
            celda.valor, bool
        ):
            if candidato not in estilos_fecha:
                return de_fecha.get(celda.columna, candidato)
        return candidato

    # De abajo hacia arriba y de derecha a izquierda: cada insercion corre las
    # posiciones de lo que sigue, y asi lo que sigue ya fue procesado.
    for celda in sorted(celdas, key=lambda c: (-c.fila, -c.columna)):
        rango = _rango_fila(xml, celda.fila)
        if rango is None:
            # O la fila viene auto-cerrada (existe pero sin celdas), o no
            # existe todavia y hay que crearla.
            expandido = _expandir_fila_vacia(xml, celda.fila)
            xml = expandido if expandido != xml else _crear_fila(xml, celda.fila)
            rango = _rango_fila(xml, celda.fila)
            if rango is None:
                raise ValueError(f"No pude abrir la fila {celda.fila}.")

        inicio_fila, fin_apertura, fin_fila = rango
        cuerpo = xml[fin_apertura:fin_fila]

        m = re.search(rf'<c r="{celda.ref}"(?=[\s/>])', cuerpo)
        if m:
            # La celda ya esta: se reemplaza entera (asi se van tambien el
            # <f> de una formula vieja y el <v> cacheado), pero se conserva su
            # propio estilo, que es mas especifico que el de la columna.
            resto = cuerpo[m.start() :]
            auto = re.match(r"<c[^>]*?/>", resto)
            largo = auto.end() if auto else resto.find("</c>") + len("</c>")
            viejo = resto[:largo]
            propio = re.search(r'\ss="(\d+)"', viejo)
            nuevo = _contenido(
                celda, elegir(celda, propio.group(1) if propio else None)
            )
            cuerpo = cuerpo[: m.start()] + nuevo + cuerpo[m.start() + largo :]
        else:
            # No esta: se inserta respetando el orden de columnas, que Excel
            # exige dentro de cada <row>.
            nuevo = _contenido(celda, elegir(celda, None))
            posicion = len(cuerpo)
            for m2 in re.finditer(r'<c r="([A-Z]+\d+)"', cuerpo):
                if _columna_de_ref(m2.group(1)) > celda.columna:
                    posicion = m2.start()
                    break
            cuerpo = cuerpo[:posicion] + nuevo + cuerpo[posicion:]

        xml = xml[:fin_apertura] + cuerpo + xml[fin_fila:]
        escritas += 1
    return xml, escritas


def _forzar_recalculo(workbook_xml: str) -> str:
    """Deja marcado que Excel recalcule todo al abrir.

    Las formulas que escribimos van sin valor cacheado, y las que ya estaban y
    dependen de una celda que cambio quedaron con el valor viejo. Con
    ``fullCalcOnLoad`` Excel recalcula la planilla entera cuando la abren, que
    es exactamente lo que hace falta y no obliga a tocar ``calcChain.xml``.
    """
    if "<calcPr" in workbook_xml:
        def reemplazo(m: re.Match) -> str:
            tag = m.group(0)
            if "fullCalcOnLoad" in tag:
                return re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', tag)
            return tag[:-2].rstrip() + ' fullCalcOnLoad="1"/>' if tag.endswith("/>") \
                else tag[:-1] + ' fullCalcOnLoad="1">'
        return re.sub(r"<calcPr[^>]*/?>", reemplazo, workbook_xml, count=1)
    return workbook_xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')


def _hojas_del_libro(zf: zipfile.ZipFile) -> dict[str, str]:
    """Nombre de hoja -> ruta de su XML dentro del zip."""
    workbook = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    # Los atributos se buscan de a uno sobre cada tag: el orden no esta
    # garantizado. Excel escribe Id antes de Target y openpyxl al revés, asi
    # que una sola expresion con los dos juntos funciona con unos archivos y
    # falla con otros.
    destino = {}
    for tag in re.findall(r"<Relationship\b[^>]*/?>", rels):
        rid = re.search(r'Id="([^"]+)"', tag)
        target = re.search(r'Target="([^"]+)"', tag)
        if rid and target:
            destino[rid.group(1)] = target.group(1)

    hojas = {}
    for tag in re.findall(r"<sheet\b[^>]*/?>", workbook):
        nombre = re.search(r'name="([^"]*)"', tag)
        rid = re.search(r'r:id="([^"]+)"', tag)
        if not nombre or not rid:
            continue
        ruta = destino.get(rid.group(1), "")
        if not ruta:
            continue
        if ruta.startswith("/"):
            ruta = ruta[1:]
        elif not ruta.startswith("xl/"):
            ruta = "xl/" + ruta
        hojas[_desescapar(nombre.group(1))] = ruta
    return hojas


def _desescapar(texto: str) -> str:
    return (
        texto.replace("&quot;", '"')
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def usa_sistema_1904(ruta: str | Path) -> bool:
    with zipfile.ZipFile(ruta) as zf:
        workbook = zf.read("xl/workbook.xml").decode("utf-8")
    return bool(re.search(r'date1904="(1|true)"', workbook))


def parchear(
    origen: str | Path, destino: str | Path, celdas: list[Celda]
) -> dict[str, int]:
    """Copia ``origen`` en ``destino`` escribiendo solo ``celdas``.

    Devuelve cuantas celdas se escribieron por hoja. Si ``origen`` y
    ``destino`` son el mismo archivo, se escribe primero a un temporal al lado
    y despues se renombra: si algo falla, la base queda intacta.
    """
    origen, destino = Path(origen), Path(destino)
    if not celdas:
        if origen != destino:
            shutil.copy2(origen, destino)
        return {}
    if usa_sistema_1904(origen):
        raise RuntimeError(
            "El libro usa el sistema de fechas 1904; este parcheador asume "
            "1900. Revisar Opciones > Avanzadas en Excel antes de seguir."
        )

    en_sitio = origen.resolve() == destino.resolve()
    salida = destino.with_suffix(destino.suffix + ".tmp") if en_sitio else destino

    por_hoja: dict[str, list[Celda]] = {}
    for celda in celdas:
        por_hoja.setdefault(celda.hoja, []).append(celda)

    escritas: dict[str, int] = {}
    with zipfile.ZipFile(origen) as zf:
        try:
            estilos_fecha = _estilos_de_fecha(
                zf.read("xl/styles.xml").decode("utf-8")
            )
        except KeyError:
            estilos_fecha = set()
        rutas = _hojas_del_libro(zf)
        faltan = [h for h in por_hoja if h not in rutas]
        if faltan:
            raise ValueError(
                f"El libro no tiene la hoja {faltan[0]!r}. Tiene: {sorted(rutas)}"
            )
        a_reescribir = {rutas[h]: h for h in por_hoja}

        salida.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(
            salida, "w", compression=zipfile.ZIP_DEFLATED
        ) as nuevo:
            for info in zf.infolist():
                crudo = zf.read(info.filename)
                if info.filename in a_reescribir:
                    hoja = a_reescribir[info.filename]
                    xml, n = _escribir_en_hoja(
                        crudo.decode("utf-8"), por_hoja[hoja], estilos_fecha
                    )
                    crudo = xml.encode("utf-8")
                    escritas[hoja] = n
                elif info.filename == "xl/workbook.xml":
                    crudo = _forzar_recalculo(crudo.decode("utf-8")).encode("utf-8")
                # Se conserva la compresion original de cada parte para no
                # inflar el archivo ni reordenar el zip.
                nuevo.writestr(info, crudo)

    if en_sitio:
        salida.replace(destino)
    return escritas
