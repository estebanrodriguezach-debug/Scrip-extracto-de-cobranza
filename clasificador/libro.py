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
from openpyxl.styles import Font
from openpyxl.worksheet.datavalidation import DataValidation

from clasificador.excel import ENCABEZADOS, HOJA_DATOS, Movimiento

HOJA_LISTAS = "Listas"

# El desplegable se aplica a un rango fijo, no solo a las filas escritas: si no,
# las filas que alguien agregue a mano quedarian sin validacion.
FILAS_MINIMAS_DESPLEGABLE = 5000

ANCHOS = {"A": 12, "B": 16, "C": 14, "D": 60, "E": 40, "F": 24}


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
        formula1=f"={HOJA_LISTAS}!$A$1:$A${len(categorias)}",
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
    for celda in ws[1]:
        celda.font = Font(bold=True)
    for columna, ancho in ANCHOS.items():
        ws.column_dimensions[columna].width = ancho
    for fila in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fila[0].number_format = "yyyy-mm-dd"
        fila[2].number_format = "#,##0.00"

    buffer = io.BytesIO()
    libro.save(buffer)
    return buffer.getvalue()
