import os
import time
import re
from glob import glob

import pandas as pd
import openpyxl
from backend.utils.txt_to_xlsx import MAINBOARD_2_FILES_FOLDER

# ======================================================
# CONFIG
# ======================================================
EXTRAER_ARCHIVO = r"Z:\IE-SAP\1) BOM files\e) Submaterial Usage"

# ======================================================
# BLOQUE 1: LIMPIEZA BASE EXCEL
# ======================================================
def mover_columnas_completas_2(ws, columnas_originales, nueva_pos):
    n = len(columnas_originales)
    datos = [[ws.cell(row=r, column=c).value for r in range(1, ws.max_row + 1)]
             for c in columnas_originales]
    for c in sorted(columnas_originales, reverse=True):
        ws.delete_cols(c)
    ws.insert_cols(nueva_pos, n)
    for i, col_data in enumerate(datos):
        for r in range(1, ws.max_row + 1):
            ws.cell(row=r, column=nueva_pos + i).value = col_data[r - 1]

def limpiar_excel_mainboard_2(ruta_xlsx: str):
    wb = openpyxl.load_workbook(ruta_xlsx)
    ws = wb.active

    ws.delete_cols(1)
    ws.delete_cols(9, 26)
    ws.delete_rows(1, 9)

    headers = [
        "LEVEL", "ITEM", "MATERIAL",
        "DESCRIPTION IN CHINESE", "DESCRIPTION IN ENGLISH",
        "QTY", "UN", "LINE 1", "LINE 2", "SORT STRING"
    ]

    ws.insert_cols(1, 1)
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    wb.save(ruta_xlsx)

def limpiar_todos_los_mainboard():
    for archivo in os.listdir(MAINBOARD_2_FILES_FOLDER):
        if archivo.lower().endswith(".xlsx"):
            ruta = os.path.join(MAINBOARD_2_FILES_FOLDER, archivo)
            try:
                limpiar_excel_mainboard_2(ruta)
                print(f"[OK] Limpio: {archivo}")
                time.sleep(1)
            except Exception as e:
                print(f"[ERROR] {archivo} → {e}")

# ======================================================
# UTILIDADES
# ======================================================
def set_cell(df, row, col, value):
    df.loc[row - 1, col] = value

def limpiar_item(x):
    if pd.isna(x) or str(x).strip() == "":
        return "X"
    return str(x).strip()

def extraer_codigo_pcb(texto):
    if isinstance(texto, str) and "PCB" in texto:
        match = re.search(r"\.(\d+)(\\|$)", texto)
        if match:
            return match.group(1)
    return None

# ======================================================
# BLOQUE 2 + SUBMATERIALES
# ======================================================
def procesar_archivo_principal_mainboard_2(
    ruta_excel_principal: str,
    ruta_salida_principal: str
):
    nombre_archivo = os.path.splitext(os.path.basename(ruta_excel_principal))[0]
    df = pd.read_excel(ruta_excel_principal)

    # --------------------------------------------------
    # Insertar 2 filas vacías al inicio
    # --------------------------------------------------
    df = pd.concat([pd.DataFrame(columns=df.columns, index=range(2)), df], ignore_index=True)

    # --------------------------------------------------
    # Configurar primeras filas
    # --------------------------------------------------
    set_cell(df, 1, "LEVEL", "0")
    set_cell(df, 1, "ITEM", "X")
    set_cell(df, 1, "MATERIAL", "3TE")

    set_cell(df, 2, "LEVEL", "1")
    set_cell(df, 2, "ITEM", "X")
    set_cell(df, 2, "MATERIAL", nombre_archivo)

    set_cell(df, 3, "LEVEL", "1")
    set_cell(df, 3, "ITEM", "X")
    set_cell(df, 3, "MATERIAL", nombre_archivo)

    filas_protegidas = {0, 1, 2}

    # --------------------------------------------------
    # Limpiar ITEM y inicializar LEVEL
    # --------------------------------------------------
    df["ITEM"] = df["ITEM"].apply(limpiar_item).astype(str)
    df.loc[~df.index.isin(filas_protegidas), "LEVEL"] = 0

    contador_bloque = 10
    for i, val in df["ITEM"].items():
        if i in filas_protegidas:
            continue
        if val == "X":
            contador_bloque = 10
            continue
        df.at[i, "ITEM"] = str(contador_bloque)
        contador_bloque += 10

    nivel_actual = 1
    for i in range(len(df)):
        if i in filas_protegidas:
            continue
        if df.at[i, "ITEM"] == "X":
            nivel_actual += 1
        else:
            df.at[i, "LEVEL"] = nivel_actual + 1

    for i in range(1, len(df)):
        if i in filas_protegidas:
            continue
        if df.at[i, "LEVEL"] == 0:
            df.at[i, "LEVEL"] = df.at[i - 1, "LEVEL"]

    # --------------------------------------------------
    # Manejo de SORT STRING
    # --------------------------------------------------
    if "SORT STRING" not in df.columns:
        df["SORT STRING"] = None
    df["SORT STRING"] = df["SORT STRING"].astype("object")
    df.loc[1:, "SORT STRING"] = df.loc[1:, "LINE 2"]
    df.loc[1:, "LINE 2"] = None

    # --------------------------------------------------
    # Extraer PCB y filtrar BOM
    # --------------------------------------------------
    df["PCB_CODE"] = df["DESCRIPTION IN CHINESE"].apply(extraer_codigo_pcb)
    lista_pcb = df["PCB_CODE"].dropna().tolist()

    if lista_pcb:
        archivos = glob(os.path.join(EXTRAER_ARCHIVO, "*.xls*"))
        archivo_bom = max(archivos, key=os.path.getmtime)
        df_bom = pd.read_excel(archivo_bom)
        df_bom["PCB_clean"] = df_bom["PCB"].astype(str).str.strip()

        mask = df_bom["PCB_clean"].apply(lambda x: any(code in x for code in lista_pcb))
        df_filtrado = df_bom.loc[mask, ["PCB", "Part #", "ZH Description", "EN Description", "QTY", "UNIT"]].reset_index(drop=True)

        # Separar Part # especiales
        finales = ["1063182", "L600022"]
        filas_final = df_filtrado[df_filtrado["Part #"].isin(finales)]
        filas_a_insertar = df_filtrado[~df_filtrado["Part #"].isin(finales)]

        # Determinar índice de inserción antes de la última X en nivel 2
        df["LEVEL"] = df["LEVEL"].astype(str)
        nivel_2_indices = df.index[df["LEVEL"].str.startswith("2")].tolist()
        indices_x = df.index[df["ITEM"].str.upper() == "X"].tolist()
        indice_insercion = max([i for i in nivel_2_indices if i < indices_x[-1]])

        # Mapeo columnas BOM -> Excel principal
        col_mapping = {
            "PCB": "ITEM",
            "Part #": "MATERIAL",
            "ZH Description": "DESCRIPTION IN CHINESE",
            "EN Description": "DESCRIPTION IN ENGLISH",
            "QTY": "QTY",
            "UNIT": "UN"
        }

        def crear_filas_pegando_datos(filas_filtradas):
            df_nuevo = pd.DataFrame(columns=df.columns)
            for i in range(len(filas_filtradas)):
                fila = {col_mapping[col]: filas_filtradas[col].iloc[i] for col in filas_filtradas.columns if col in col_mapping}
                df_nuevo = pd.concat([df_nuevo, pd.DataFrame([fila], columns=df.columns)], ignore_index=True)
            return df_nuevo

        df_a_insertar_completo = crear_filas_pegando_datos(filas_a_insertar)
        df_final_insertar_completo = crear_filas_pegando_datos(filas_final)

        df = pd.concat([
            df.iloc[:indice_insercion+1],
            df_a_insertar_completo,
            df.iloc[indice_insercion+1:],
            df_final_insertar_completo
        ], ignore_index=True)

    # --------------------------------------------------
    # Limpiar y reasignar ITEM / LEVEL final
    # --------------------------------------------------
    df["ITEM"] = df["ITEM"].apply(limpiar_item).astype(str)
    filas_protegidas = {0, 1, 2}
    df.loc[~df.index.isin(filas_protegidas), "LEVEL"] = 0

    contador_bloque = 10
    for i, val in df["ITEM"].items():
        if i in filas_protegidas:
            continue
        if val == "X":
            contador_bloque = 10
            continue
        df.at[i, "ITEM"] = str(contador_bloque)
        contador_bloque += 10

    nivel_actual = 1
    for i in range(len(df)):
        if i in filas_protegidas:
            continue
        if df.at[i, "ITEM"] == "X":
            nivel_actual += 1
        else:
            df.at[i, "LEVEL"] = nivel_actual + 1

    for i in range(1, len(df)):
        if i in filas_protegidas:
            continue
        if df.at[i, "LEVEL"] == 0:
            df.at[i, "LEVEL"] = df.at[i - 1, "LEVEL"]

    # Limpiar columnas auxiliares
    df.drop(columns=["PCB_CODE", "PCB_clean"], errors="ignore", inplace=True)

    # Guardar archivo final
    df.to_excel(ruta_salida_principal, index=False)
    print(f"[OK] Mainboard P2 COMPLETO: {ruta_salida_principal}")
