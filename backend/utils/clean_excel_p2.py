import os
import re
from glob import glob
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ----------------- FUNCIONES AUXILIARES -----------------
def limpiar_item(x):
    if pd.isna(x) or str(x).strip() == "":
        return "X"
    return str(x).strip()

def extraer_codigo_pcb(texto):
    if isinstance(texto, str) and ("PCB" in texto or "印制板" in texto):
        match = re.search(r"\.(\d+)(\\|$)", texto)
        if match:
            return match.group(1)
    return None

def tiene_chino(texto):
    if isinstance(texto, str):
        return any('\u4e00' <= c <= '\u9fff' for c in texto)
    return False

def safe_float(x):
    try:
        return float(str(x).replace(",", ""))
    except (ValueError, AttributeError):
        return x

# ----------------- LIMPIEZA INICIAL CON OPENPYXL -----------------
def limpiar_excel_inicial(ruta_excel, ruta_salida):
    wb = openpyxl.load_workbook(ruta_excel)
    ws = wb.active

    # Eliminar filas y columnas innecesarias
    ws.delete_cols(10, 26)
    ws.delete_rows(1, 9)

    # Encabezados
    headers = [
        "LEVEL", "ITEM", "MATERIAL",
        "DESCRIPTION IN CHINESE", "DESCRIPTION IN ENGLISH",
        "QTY", "UN", "LINE 1", "LINE 2", "SORTSTRNG"
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    # Guardar archivo limpio
    wb.save(ruta_salida)
    print(f"[OK] Limpieza inicial completada → {ruta_salida}")

# ----------------- LOGICA PRINCIPAL CON PANDAS -----------------
def procesar_excel_con_logica(ruta_excel_principal, ruta_salida_principal, carpeta_extraer=None):
    # Primero limpiar el archivo para asegurar encabezados correctos
    limpiar_excel_inicial(ruta_excel_principal, ruta_excel_principal)

    nombre_archivo = os.path.splitext(os.path.basename(ruta_excel_principal))[0]
    df = pd.read_excel(ruta_excel_principal)

    # Insertar 2 filas vacías al inicio
    df = pd.concat([pd.DataFrame(columns=df.columns, index=range(2)), df], ignore_index=True)

    # Configurar primeras filas manuales
    df.loc[0, "ITEM"] = "X"
    df.loc[0, "MATERIAL"] = "3TE"
    df.loc[1, "LEVEL"] = 1
    df.loc[1, "ITEM"] = ""
    df.loc[1, "MATERIAL"] = nombre_archivo
    df.loc[2, "LEVEL"] = 1
    df.loc[2, "ITEM"] = "X"
    df.loc[2, "MATERIAL"] = nombre_archivo

    filas_protegidas = {0, 1, 2}

    # Convertir columnas numéricas de forma segura
    for col in ["LEVEL", "QTY"]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: safe_float(x) if pd.notna(x) else x)

    # Reinicio ITEM/LEVEL
    df["ITEM"] = df["ITEM"].apply(limpiar_item)
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
            df.at[i, "LEVEL"] = df.at[i-1, "LEVEL"]

    # Mover datos con caracteres chinos a SORTSTRNG
    if "SORTSTRNG" not in df.columns:
        df["SORTSTRNG"] = None
    for col in ["LINE 1", "LINE 2"]:
        if col in df.columns:
            mask = df[col].apply(tiene_chino)
            df.loc[mask, "SORTSTRNG"] = df.loc[mask, col]
            df.loc[mask, col] = None

    # Submateriales y filtrado PCB
    if "DESCRIPTION IN CHINESE" in df.columns:
        df["PCB_CODE"] = df["DESCRIPTION IN CHINESE"].apply(extraer_codigo_pcb)
        lista_pcb = df["PCB_CODE"].dropna().tolist()
        if lista_pcb and carpeta_extraer:
            archivos = glob(os.path.join(carpeta_extraer, "*.xls*"))
            if archivos:
                archivo_bom = max(archivos, key=os.path.getmtime)
                df_bom = pd.read_excel(archivo_bom)
                df_bom["PCB_clean"] = df_bom["PCB"].astype(str).str.strip()
                if "USE/NO USE" in df_bom.columns:
                    df_bom = df_bom[df_bom["USE/NO USE"].astype(str).str.strip().str.upper() != "NO USE"]
                mask = df_bom["PCB_clean"].apply(lambda x: any(code in x for code in lista_pcb))
                df_filtrado = df_bom.loc[mask, ["PCB","Part #","ZH Description","EN Description","QTY","UNIT"]].reset_index(drop=True)

                col_mapping = {
                    "PCB": "ITEM",
                    "Part #": "MATERIAL",
                    "ZH Description": "DESCRIPTION IN CHINESE",
                    "EN Description": "DESCRIPTION IN ENGLISH",
                    "QTY": "QTY",
                    "UNIT": "UN"
                }

                df_nuevo = pd.DataFrame(columns=df.columns)
                for i in range(len(df_filtrado)):
                    fila = {col_mapping[c]: df_filtrado[c].iloc[i] for c in df_filtrado.columns if c in col_mapping}
                    df_nuevo = pd.concat([df_nuevo, pd.DataFrame([fila], columns=df.columns)], ignore_index=True)
                df_nuevo["LEVEL"] = 2
                df = pd.concat([df, df_nuevo], ignore_index=True)

    # Guardar Excel
    df.drop(columns=["PCB_CODE"], errors="ignore", inplace=True)
    df.to_excel(ruta_salida_principal, index=False)

    # Colorear filas con SORTSTRNG
    wb = load_workbook(ruta_salida_principal)
    ws = wb.active
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    if "SORTSTRNG" in df.columns:
        col_sort_string = df.columns.get_loc("SORTSTRNG") + 1
        for idx, tiene in enumerate(df["SORTSTRNG"].notna(), start=2):
            if tiene:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=idx, column=col).fill = amarillo
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=col_sort_string).value = None

    ws.title = "BOMlist"
    wb.save(ruta_salida_principal)
    print(f"[OK] Mainboard COMPLETO → {ruta_salida_principal}")