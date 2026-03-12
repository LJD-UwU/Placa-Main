import os
import openpyxl
import pandas as pd
import re

from glob import glob
from backend.config.sap_config import EXTRAER_ARCHIVO

from openpyxl.styles import PatternFill, Font, Alignment


def contiene_chino(texto):
    """Detecta si un texto contiene caracteres chinos."""
    return any('\u4e00' <= c <= '\u9fff' for c in str(texto))


def extraer_codigo_pcb(texto, siguiente_celda=None):
    """Extrae código PCB si el texto contiene letras y números."""
    if isinstance(texto, str) and re.search(r'[A-Za-z].*\d|\d.*[A-Za-z]', texto):
        if siguiente_celda:
            match = re.search(r"\.(\d+)(\\|$)", str(siguiente_celda))
            if match:
                return match.group(1)
    return None


# ================== LÓGICA DE LAS X ==================
def aplicar_logica_x(ws):

    col_indices = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}

    col_item = col_indices.get("ITEM")
    col_level = col_indices.get("LEVEL")

    if not col_item or not col_level:
        return

    filas_protegidas = {2, 3, 4}

    bold_font = Font(bold=True)

    # ITEM vacío → X
    for row in range(2, ws.max_row + 1):

        if row in filas_protegidas:
            continue

        val = ws.cell(row=row, column=col_item).value

        if val is None or str(val).strip() == "":
            ws.cell(row=row, column=col_item).value = "X"

            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).font = bold_font

    # Numeración por bloques
    contador = 10

    for row in range(2, ws.max_row + 1):

        if row in filas_protegidas:
            continue

        val = ws.cell(row=row, column=col_item).value

        if val == "X":
            contador = 10
            continue

        ws.cell(row=row, column=col_item).value = str(contador)

        contador += 10

    # LEVEL jerárquico
    nivel_actual = 1

    for row in range(2, ws.max_row + 1):

        if row in filas_protegidas:
            continue

        val = ws.cell(row=row, column=col_item).value

        if val == "X":
            nivel_actual += 1
        else:
            ws.cell(row=row, column=col_level).value = nivel_actual + 1

    # Rellenar LEVEL vacío
    for row in range(3, ws.max_row + 1):

        if ws.cell(row=row, column=col_level).value in (None, ""):
            ws.cell(row=row, column=col_level).value = ws.cell(row=row - 1, column=col_level).value


# ================== PROCESO PRINCIPAL ==================
def procesar_archivo_principal_mainboard_2(
        ruta_excel_principal: str,
        ruta_salida_principal: str,
        internal_model: str = "",
        plantas: str = "",
        df_no_procesadas: pd.DataFrame = None
):

    wb = openpyxl.load_workbook(ruta_excel_principal)

    ws = wb.active

    # limpiar columnas
    ws.delete_cols(1)
    ws.delete_cols(9, 26)
    ws.delete_rows(1, 9)

    headers = [
        "LEVEL", "ITEM", "MATERIAL",
        "DESCRIPTION IN CHINESE", "DESCRIPTION IN ENGLISH",
        "QTY", "UN", "LINE 1", "LINE 2", "SORTSTRNG"
    ]

    ws.insert_cols(1)

    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    nombre_archivo = os.path.splitext(os.path.basename(ruta_excel_principal))[0]

    ws.insert_rows(2, 2)

    ws["B2"] = "X"
    ws["C2"] = "3TE"

    ws["A3"] = "1"
    ws["B3"] = " "
    ws["C3"] = nombre_archivo

    ws["A4"] = "1"
    ws["B4"] = "X"
    ws["C4"] = nombre_archivo

    aplicar_logica_x(ws)

    # detectar chino
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    col_indices = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}

    for row in range(2, ws.max_row + 1):

        fila_colorear = False

        for col_name in ["LINE 1", "LINE 2"]:

            col = col_indices.get(col_name)

            if col:

                val = ws.cell(row=row, column=col).value

                if val and contiene_chino(val):
                    ws.cell(row=row, column=col).value = None
                    fila_colorear = True

        if fila_colorear:

            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = amarillo

    # obtener número de mainboard
    nombre_archivo = os.path.basename(ruta_excel_principal)

    mainboard_num = os.path.splitext(nombre_archivo)[0]

    mainboard_num = mainboard_num.replace("MB-BMM-", "").replace(".0", "").strip()

    texto_modelo = ""

    # buscar INTERNAL MODEL
    if df_no_procesadas is not None and not df_no_procesadas.empty:

        df_temp = df_no_procesadas.copy()

        df_temp["MAINBOARD PART NUMBER"] = (
            df_temp["MAINBOARD PART NUMBER"]
            .astype(str)
            .str.replace(".0", "", regex=False)
            .str.strip()
        )

        mainboard_num_clean = mainboard_num.replace(".0", "").strip()

        fila_match = df_temp[
            df_temp["MAINBOARD PART NUMBER"].str.contains(mainboard_num_clean, na=False)
        ]

        if not fila_match.empty:
            texto_modelo = str(fila_match.iloc[0]["INTERNAL MODEL"]).strip()

    # escribir valores
    ws["E3"] = f"MAIN BOARD\\{texto_modelo}\\ROH"
    ws["E4"] = f"MAIN BOARD\\{texto_modelo}\\ROH"

    ws["D3"] = plantas.strip() if plantas else ""

    ws.title = "BOMlist"

    ws["A2"] = "0"
    ws["F3"] = "1000"
    ws["J3"] = "HIMEX"
    ws["G3"] = "PC"

    # columnas numéricas
    columnas_numericas = ["LEVEL", "ITEM", "QTY"]

    mapa_columnas = {
        str(ws.cell(row=1, column=c).value).strip().upper():
            openpyxl.utils.get_column_letter(c)
        for c in range(1, ws.max_column + 1)
        if str(ws.cell(row=1, column=c).value).strip().upper() in columnas_numericas
    }

    for letra in mapa_columnas.values():

        for cell in ws[letra][1:]:

            if cell.value is not None:

                valor = str(cell.value).strip()

                if valor != "":
                    try:
                        cell.value = float(valor.replace(",", ""))
                    except:
                        pass

    # alinear texto
    for fila in ws.iter_rows():
        for celda in fila:
            celda.alignment = Alignment(horizontal="left")

    wb.save(ruta_salida_principal)

    print(f"[OK] Proceso completo {ruta_salida_principal}")