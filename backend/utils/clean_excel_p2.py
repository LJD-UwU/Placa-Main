import os 
import time
import re
from glob import glob
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from backend.utils.txt_to_xlsx import MAINBOARD_2_FILES_FOLDER
from backend.config.sap_config import EXTRAER_ARCHIVO

# BLOQUE 1: LIMPIEZA BASE EXCEL

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
        "QTY", "UN", "LINE 1", "LINE 2", "SORTSTRNG" 
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

# UTILIDADES

def set_cell(df, row, col, value):
    df.loc[row - 1, col] = value

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

# BLOQUE 2 + SUBMATERIALES

def procesar_archivo_principal_mainboard_2(
    ruta_excel_principal: str,
    ruta_salida_principal: str
):

    nombre_archivo = os.path.splitext(os.path.basename(ruta_excel_principal))[0]
    df = pd.read_excel(ruta_excel_principal)

    # Insertar 2 filas vacías al inicio
    df = pd.concat([pd.DataFrame(columns=df.columns, index=range(2)), df], ignore_index=True)

    # Configurar primeras filas “manuales”
    set_cell(df, 1, "ITEM", "X")
    set_cell(df, 1, "MATERIAL", "3TE")

    set_cell(df, 2, "LEVEL", "1")
    set_cell(df, 2, "ITEM", "")
    set_cell(df, 2, "MATERIAL", nombre_archivo)  

    set_cell(df, 3, "LEVEL", "1")
    set_cell(df, 3, "ITEM", "X")
    set_cell(df, 3, "MATERIAL", nombre_archivo)

    filas_protegidas = {0, 1, 2}

    # PRIMER REFRESH: CONVERTIR VALORES MANUALES A FLOAT 
    columnas_a_float = ["LEVEL", "MATERIAL", "ITEM", "QTY"]

    for col in columnas_a_float:
        if col in df.columns:
            def convertir_valor(x):
                try:
                    if pd.notna(x) and str(x).strip() not in {"X", "3TE"} and str(x).strip() != "":
                        return float(str(x).replace(",", ""))
                    return x
                except:
                    return x
            df[col] = df[col].apply(convertir_valor)
    # FIN PRIMER REFRESH

    # LOGICA ORIGINAL DE REINICIO ITEM / LEVEL
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

    # DETECCIÓN DE CHINO
    if "SORTSTRNG" not in df.columns:
        df["SORTSTRNG"] = None
    df["SORTSTRNG"] = df["SORTSTRNG"].astype("object")

    def tiene_chino(texto):
        if isinstance(texto, str):
            return any('\u4e00' <= c <= '\u9fff' for c in texto)
        return False

    for col in ["LINE 1", "LINE 2"]:
        mask = df[col].apply(tiene_chino)
        df.loc[mask, "SORTSTRNG"] = df.loc[mask, col]
        df.loc[mask, col] = None

    # SUBMATERIALES + FILTRO USE / NO USE
    df["PCB_CODE"] = df["DESCRIPTION IN CHINESE"].apply(extraer_codigo_pcb)
    lista_pcb = df["PCB_CODE"].dropna().tolist()

    if lista_pcb:
        archivos = glob(os.path.join(EXTRAER_ARCHIVO, "*.xls*"))
        archivo_bom = max(archivos, key=os.path.getmtime)
        df_bom = pd.read_excel(archivo_bom)
        df_bom["PCB_clean"] = df_bom["PCB"].astype(str).str.strip()

        if "USE/NO USE" in df_bom.columns:
            df_bom["USE/NO USE"] = df_bom["USE/NO USE"].astype(str).str.strip().str.upper()
            df_bom = df_bom[df_bom["USE/NO USE"] != "NO USE"]

        mask = df_bom["PCB_clean"].apply(lambda x: any(code in x for code in lista_pcb))
        df_filtrado = df_bom.loc[mask, ["PCB","USE/NO USE","Part #","ZH Description",
                                        "EN Description","QTY","UNIT"]].reset_index(drop=True)

        finales = {"L600022","1063182"}
        df_filtrado["Part #"] = df_filtrado["Part #"].astype(str)
        filas_final = df_filtrado[df_filtrado["Part #"].isin(finales)]
        filas_a_insertar = df_filtrado[~df_filtrado["Part #"].isin(finales)]

        df["LEVEL"] = df["LEVEL"].astype(str)
        nivel_2_indices = df.index[df["LEVEL"].str.startswith("2")].tolist()
        indices_x = df.index[df["ITEM"].str.upper() == "X"].tolist()
        indice_insercion = max([i for i in nivel_2_indices if i < indices_x[-1]])

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
                fila = {col_mapping[col]: filas_filtradas[col].iloc[i]
                        for col in filas_filtradas.columns if col in col_mapping}
                df_nuevo = pd.concat([df_nuevo,
                                      pd.DataFrame([fila], columns=df.columns)],
                                      ignore_index=True)
            return df_nuevo

        # CREAR SUBMATERIALES
        df_sub_normales = crear_filas_pegando_datos(filas_a_insertar)
        df_sub_final = crear_filas_pegando_datos(filas_final)

        # FILAS MANUALES FINALES
        filas_manuales = [
            {
                "ITEM": "73467",
                "MATERIAL": "L100022",
                "DESCRIPTION IN CHINESE": "",
                "DESCRIPTION IN ENGLISH": "MB BARCODE LABEL (28mm*8mm)",
                "QTY": "1,000",
                "UN": "PC",
                "LEVEL": 2
            },
            {
                "ITEM": "7353742",
                "MATERIAL": "L600006",
                "DESCRIPTION IN CHINESE": "",
                "DESCRIPTION IN ENGLISH": "RIBBON\\110mm*450m\\LOCAL 556",
                "QTY": "556",
                "UN": "",
                "LEVEL": 2
            }
        ]

        df_manuales = pd.DataFrame(filas_manuales, columns=df.columns)

        # CONCATENAR SUBMATERIALES Y MANUALES
        df = pd.concat([
            df.iloc[:indice_insercion+1],
            df_sub_normales,
            df_manuales,
            df.iloc[indice_insercion+1:],
            df_sub_final
        ], ignore_index=True)

    # REINICIO LOGICA ITEM / LEVEL DESPUES DE INSERTAR FILAS
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

    # GUARDAR EXCEL
    df.drop(columns=["PCB_CODE","PCB_clean"], errors="ignore", inplace=True)
    df.to_excel(ruta_salida_principal, index=False)
    wb = load_workbook(ruta_salida_principal)
    ws = wb.active

    # Amarillo → filas con chino
    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for idx, tiene in enumerate(df["SORTSTRNG"].notna(), start=2):
        if tiene:
            for col in range(1, 10):
                ws.cell(row=idx, column=col).fill = amarillo

    # Limpiar contenido de la columna SORTSTRNG
    col_sort_string = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "SORTSTRNG":
            col_sort_string = col
            break
    if col_sort_string:
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=col_sort_string).value = None

    # Cambiar el nombre de la hoja existente
    ws.title = "BOMlist"

    # TEXTOS FIJOS
    ws["A2"] = "0"
    ws["F2"] = "1000"
    ws["F3"] = "1000"
    ws["B3"] = " "
    ws["J3"] = "HIMEX"
    ws["G3"] = "PC"
    ws["D3"] = " MAINBOARD\INTERNAL MODEL\ROH"
    ws["D4"] = "MAINBOARD SMT PART\INTERNAL MODEL\ROH"

    # --- SEGUNDO REFRESH FINAL: CONVERTIR DATOS NUMÉRICOS ---
    columnas_numericas = ["LEVEL", "ITEM", "QTY"]
    mapa_columnas = {str(ws.cell(row=1, column=c).value).strip().upper(): openpyxl.utils.get_column_letter(c)
                     for c in range(1, ws.max_column + 1)
                     if str(ws.cell(row=1, column=c).value).strip().upper() in columnas_numericas}

    for nombre, letra in mapa_columnas.items():
        for cell in ws[letra][1:]:  # fila 1 es header
            if cell.value is not None:
                valor = str(cell.value).strip()
                if valor != "":
                    try:
                        cell.value = float(valor.replace(",", ""))
                    except:
                        pass

    # Agregar dos nuevas hojas vacías
    wb.create_sheet("BOMHeader")
    wb.create_sheet("BOMItem")
    wb.save(ruta_salida_principal)

    print(f"[OK] Mainboard P2 COMPLETO: {ruta_salida_principal}")
