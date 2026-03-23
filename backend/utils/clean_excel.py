import os
import time
import openpyxl
from backend.utils.txt_to_xlsx import MAINBOARD_1_FILES_FOLDER


def mover_columnas_completas(ws, columnas_originales, nueva_pos):
    n = len(columnas_originales)

    #! Guardar todos los datos de las columnas
    datos = [[ws.cell(row=r, column=c).value for r in range(1, ws.max_row + 1)]
    for c in columnas_originales]

    #! Eliminar columnas originales (de mayor a menor para no desordenar)
    for c in sorted(columnas_originales, reverse=True):
        ws.delete_cols(c)

    #! Insertar columnas en la nueva posición
    ws.insert_cols(nueva_pos, n)

    #! Colocar los datos en las nuevas columnas
    for i, col_data in enumerate(datos):
        for r in range(1, ws.max_row + 1):
            ws.cell(row=r, column=nueva_pos + i).value = col_data[r - 1]


def limpiar_excel_mainboard(ruta_xlsx: str):
    wb = openpyxl.load_workbook(ruta_xlsx)
    ws = wb.active

    #! Eliminar columnas y filas según tu lógica original
    ws.delete_cols(1)
    ws.delete_cols(9,26)
    ws.delete_rows(1, 9)

    #! Insertar cabeceras
    ws.insert_cols(1, 1)
    ws.cell(row=1, column=1).value = "LEVEL"
    ws.cell(row=1, column=2).value = "ITEM"
    ws.cell(row=1, column=3).value = "MATERIAL"
    ws.cell(row=1, column=4).value = "DESCRIPTION IN CHINESE"
    ws.cell(row=1, column=5).value = "DESCRIPTION IN ENGLISH"
    ws.cell(row=1, column=6).value = "QTY"
    ws.cell(row=1, column=7).value = "UN"
    ws.cell(row=1, column=8).value = "LINE 1"
    ws.cell(row=1, column=9).value = "LINE 2"
    ws.cell(row=1, column=10).value = "SORT STRING"

    wb.save(ruta_xlsx)

def limpiar_todos_los_mainboard():
    for archivo in os.listdir(MAINBOARD_1_FILES_FOLDER):
        if archivo.lower().endswith(".xlsx"):
            ruta = os.path.join(MAINBOARD_1_FILES_FOLDER, archivo)

            try:
                limpiar_excel_mainboard(ruta)
                print(f"[OK] Limpio: {archivo}\n")
                time.sleep(1)  #! espera entre archivos
            except Exception as e:
                print(f"[ERROR] {archivo} → {e}")
