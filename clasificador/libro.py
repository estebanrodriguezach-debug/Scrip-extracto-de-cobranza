"""Armado del .xlsx del extracto, con el mismo formato que la planilla de Google.

Para Google Sheets el formato se arma con la API (ver ``sheets.preparar``). Para
SharePoint no hay API equivalente: la Excel REST API de Graph no admite permisos
de aplicacion y encima no expone validacion de datos, asi que no podria crear el
desplegable de Categoria. La salida, entonces, es armar el libro entero aca con
openpyxl y subirlo como archivo; Graph lo guarda tal cual, con hojas,
desplegable y formatos intactos.

Este modulo es el espejo local de ``PlanillaCobranzas.preparar``: si cambia el
formato de una, tiene que cambiar el de la otra.
"""

from __future__ import annotations

import io

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from clasificador.excel import ENCABEZADOS, HOJA_DATOS, Movimiento

HOJA_LISTAS = "Listas"

# El desplegable se aplica a un rango fijo, no solo a las filas escritas: si no,
# las filas que alguien agregue a mano quedarian sin validacion.
FILAS_MINIMAS_DESPLEGABLE = 5000

# Tomados del Extractos_cobranzas.xlsx que se venia usando, para que el
# archivo generado se vea igual al de antes y no apenas parecido. La B no
# estaba fijada en el original, asi que se deja en el ancho por defecto.
ANCHOS = {"A": 12.0, "C": 15.0, "D": 60.0, "E": 30.0, "F": 38.0}

# Encabezado del original: Calibri 11 en negrita, blanco sobre azul oscuro.
COLOR_ENCABEZADO = "FF1F4E78"
COLOR_TEXTO_ENCABEZADO = "FFFFFFFF"

# El original guarda las fechas en mm-dd-yy. Es ambiguo de leer, pero cambiarlo
# haria que el archivo se vea distinto al que el equipo viene mirando.
FORMATO_FECHA = "mm-dd-yy"
FORMATO_IMPORTE = "#,##0.00"


def _fila(mov: Movimiento) -> list[object]:
    """Como ``Movimiento.a_fila`` pero con fecha e importe nativos.

    ``a_fila`` serializa todo a texto porque Google Sheets lo interpreta al
    escribirlo (``USER_ENTERED``). Excel no: una fecha escrita como texto queda
    como texto, no se ordena ni se filtra por mes. Como el archivo se arma aca,
    se escriben los tipos de verdad; ``leer_movimientos`` los lee igual.
    """
    return [
        mov.fecha,
        mov.banco,
        mov.importe,
        mov.extracto,
        mov.detalle,
        mov.categoria,
    ]


def armar_libro(
    movimientos: list[Movimiento], categorias: tuple[str, ...] | list[str]
) -> bytes:
    """Devuelve los bytes de un .xlsx con el formato del extracto.

    Hoja ``Cobranzas`` con las 6 columnas, hoja ``Listas`` con las categorias,
    desplegable sobre la columna F (Categoria) y fila 1 congelada.
    """
    categorias = list(categorias)
    libro = openpyxl.Workbook()

    ws = libro.active
    ws.title = HOJA_DATOS
    ws.append(ENCABEZADOS)
    for mov in movimientos:
        ws.append(_fila(mov))

    listas = libro.create_sheet(HOJA_LISTAS)
    for categoria in categorias:
        listas.append([categoria])

    validacion = DataValidation(
        type="list",
        # Sin el '=' adelante: en OOXML el contenido de <formula1> es la formula
        # pelada y openpyxl la escribe verbatim. Con el '=' Excel de escritorio
        # y Excel Online ofrecen reparar el archivo y descartan la validacion,
        # que es justamente el desplegable que hay que conservar.
        formula1=f"{HOJA_LISTAS}!$A$1:$A${len(categorias)}",
        allow_blank=True,
        # Trampa de OOXML: el atributo se llama showDropDown pero esta invertido,
        # en True ESCONDE la flechita. False es lo que muestra el desplegable.
        showDropDown=False,
    )
    validacion.errorTitle = "Categoría"
    validacion.error = "Elegí una categoría de la lista."
    ws.add_data_validation(validacion)
    ultima = max(ws.max_row, FILAS_MINIMAS_DESPLEGABLE)
    validacion.add(f"F2:F{ultima}")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{ws.max_row}"
    relleno = PatternFill("solid", fgColor=COLOR_ENCABEZADO)
    centrado = Alignment(horizontal="center", vertical="center")
    for celda in ws[1]:
        celda.font = Font(
            name="Calibri", size=11, bold=True, color=COLOR_TEXTO_ENCABEZADO
        )
        celda.fill = relleno
        celda.alignment = centrado
    for columna, ancho in ANCHOS.items():
        ws.column_dimensions[columna].width = ancho
    for fila in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fila[0].number_format = FORMATO_FECHA
        fila[2].number_format = FORMATO_IMPORTE

    buffer = io.BytesIO()
    libro.save(buffer)
    return buffer.getvalue()
