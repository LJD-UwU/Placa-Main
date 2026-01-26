import os
import openpyxl
from backend.utils.txt_to_xlsx import MAINBOARD_FILES_FOLDER


def limpiar_excel_mainboard(ruta_xlsx: str):
    """
    Limpia UN archivo XLSX del ensamble de MAINBOARD
    """

    wb = openpyxl.load_workbook(ruta_xlsx)
    ws = wb.active

    #  Eliminar columna A
    ws.delete_cols(1)

    # Eliminar primeras 9 filas
    ws.delete_rows(1, 9)

    #  Insertar celdas de niveles
    ws.insert_cols(3)
    ws["B2"] = "3TE"
    ws["C2"] = "Ensamble Mainboard"
    ws["B1"] = "Nivel 1"
    ws["B1"] = "Nivel 2"

    # Insertar columnas
    ws.insert_cols(1, 2)
    ws["E1"] = "" 
    ws["F1"] = ""

    #  Mover columnas J y K después de descripción en inglés
    headers = [cell.value for cell in ws[1]]

    try:
        idx_desc_en = headers.index("Description (EN)") + 1

        data_j = [ws.cell(row=r, column=10).value for r in range(1, ws.max_row + 1)]
        data_k = [ws.cell(row=r, column=11).value for r in range(1, ws.max_row + 1)]

        ws.delete_cols(10, 2)
        ws.insert_cols(idx_desc_en + 1, 2)

        for r in range(1, ws.max_row + 1):
            ws.cell(row=r, column=idx_desc_en + 1).value = data_j[r - 1]
            ws.cell(row=r, column=idx_desc_en + 2).value = data_k[r - 1]

    except ValueError:
        pass  # No existe descripción en inglés

    # 6. Eliminar columnas posteriores a 项目文本行 
    headers = [cell.value for cell in ws[1]]

    if "项目文本行 " in headers:
        idx_fin = headers.index("项目文本行") + 1
        if ws.max_column > idx_fin:
            ws.delete_cols(idx_fin + 1, ws.max_column - idx_fin)

    wb.save(ruta_xlsx)


def limpiar_todos_los_mainboard():
    """
    Limpia AUTOMÁTICAMENTE todos los XLSX de la carpeta MAINBOARD
    """

    for archivo in os.listdir(MAINBOARD_FILES_FOLDER):
        if archivo.lower().endswith(".xlsx"):
            ruta = os.path.join(MAINBOARD_FILES_FOLDER, archivo)

            try:
                limpiar_excel_mainboard(ruta)
                print(f"[OK] Limpio: {archivo}")
            except Exception as e:
                print(f"[ERROR] {archivo} → {e}")
