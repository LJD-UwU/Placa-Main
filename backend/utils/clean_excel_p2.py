import os
import openpyxl
from openpyxl.styles import PatternFill, Font
import pandas as pd
from glob import glob
from backend.config.sap_config import EXTRAER_ARCHIVO
import re

# ================== FUNCIONES ==================
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
    """Aplica la lógica de marcar ITEM vacíos con 'X', numeración por bloques y LEVEL jerárquico."""
    col_indices = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    col_item = col_indices.get("ITEM")
    col_level = col_indices.get("LEVEL")
    if not col_item or not col_level:
        return

    filas_protegidas = {2, 3, 4}
    bold_font = Font(bold=True)

    # 1. ITEM vacío → X
    for row in range(2, ws.max_row + 1):
        if row in filas_protegidas:
            continue
        val = ws.cell(row=row, column=col_item).value
        if val is None or str(val).strip() == "":
            ws.cell(row=row, column=col_item).value = "X"
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).font = bold_font

    # 2. Numeración por bloques
    contador = 10
    for row in range(2, ws.max_row + 1):
        if row in filas_protegidas:
            continue
        val = ws.cell(row=row, column=col_item).value
        if val == "X":
            contador = 10
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).font = bold_font
            continue
        ws.cell(row=row, column=col_item).value = str(contador)
        contador += 10

    # 3. LEVEL jerárquico
    nivel_actual = 1
    for row in range(2, ws.max_row + 1):
        if row in filas_protegidas:
            continue
        val = ws.cell(row=row, column=col_item).value
        if val == "X":
            nivel_actual += 1
        else:
            ws.cell(row=row, column=col_level).value = nivel_actual + 1

    # 4. Rellenar LEVEL vacío
    for row in range(3, ws.max_row + 1):
        if ws.cell(row=row, column=col_level).value in (None, ""):
            ws.cell(row=row, column=col_level).value = ws.cell(row=row - 1, column=col_level).value

# ================== SUBMATERIALES ==================
def agregar_submateriales(df_main, ws):
    """
    Agrega submateriales de BOM y manuales dentro del LEVEL 2 antes de la X correspondiente.
    Aplica color gris y fuente Calibri 11 sin negrita a todos los submateriales.
    """
    gris_submaterial = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fuente_normal = Font(name="Calibri", size=11, bold=False)

    # ===== Extraer códigos PCB =====
    df_main["PCB_CODE"] = df_main.apply(
        lambda row: extraer_codigo_pcb(row["MATERIAL"], row["DESCRIPTION IN CHINESE"]), axis=1
    )
    lista_pcb = df_main["PCB_CODE"].dropna().tolist()
    if not lista_pcb:
        df_main.drop(columns=["PCB_CODE"], inplace=True, errors="ignore")
        return df_main

    # ===== Cargar penúltimo archivo BOM =====
    archivos = [f for f in glob(os.path.join(EXTRAER_ARCHIVO, "*.xlsx")) if not os.path.basename(f).startswith("~$")]
    if len(archivos) < 2:
        print("[WARN] No hay suficientes archivos BOM para tomar el anterior al más reciente.")
        df_main.drop(columns=["PCB_CODE"], inplace=True, errors="ignore")
        return df_main
    archivo_bom = sorted(archivos, key=os.path.getmtime)[-1]
    print(f"[INFO] Archivo BOM tomado: {archivo_bom}")

    try:
        df_bom = pd.read_excel(archivo_bom, engine="openpyxl")
    except Exception as e:
        print(f"[ERROR] No se pudo leer {archivo_bom}: {e}")
        df_main.drop(columns=["PCB_CODE"], inplace=True, errors="ignore")
        return df_main

    # ===== Filtrar USE =====
    df_bom["PCB_clean"] = df_bom["PCB"].astype(str).str.strip()
    if "USE/NO USE" in df_bom.columns:
        df_bom["USE/NO USE"] = df_bom["USE/NO USE"].astype(str).str.strip().str.upper()
        df_bom = df_bom[df_bom["USE/NO USE"] != "NO USE"]

    # ===== Filtrar submateriales relacionados con PCB =====
    mask = df_bom["PCB_clean"].apply(lambda x: any(code in x for code in lista_pcb))
    cols_interes = ["PCB","Part #","ZH Description","EN Description","QTY","UNIT"]
    df_filtrado = df_bom.loc[mask, cols_interes].reset_index(drop=True)

    # ===== Separar finales y normales =====
    finales = {"L600022","1063182"}
    df_filtrado["Part #"] = df_filtrado["Part #"].astype(str)
    df_finales = df_filtrado[df_filtrado["Part #"].isin(finales)]
    df_normales = df_filtrado[~df_filtrado["Part #"].isin(finales)]

    # ===== Mapear columnas BOM → Excel =====
    col_map = {
        "PCB": "ITEM",
        "Part #": "MATERIAL",
        "ZH Description": "DESCRIPTION IN CHINESE",
        "EN Description": "DESCRIPTION IN ENGLISH",
        "QTY": "QTY",
        "UNIT": "UN"
    }

    def mapear_filas(df_sub):
        df_nuevo = pd.DataFrame(columns=df_main.columns)
        for _, fila in df_sub.iterrows():
            nueva = {col_map[col]: fila[col] for col in df_sub.columns if col in col_map}
            df_nuevo = pd.concat([df_nuevo, pd.DataFrame([nueva], columns=df_main.columns)], ignore_index=True)
        df_nuevo["LEVEL"] = 2
        df_nuevo["_SUBMATERIAL"] = True
        return df_nuevo

    df_sub_normales = mapear_filas(df_normales)
    df_sub_finales = mapear_filas(df_finales)

    # ===== Filas manuales =====
    filas_manuales = [
        {
            "ITEM": "73467",
            "MATERIAL": "L100022",
            "DESCRIPTION IN CHINESE": "",
            "DESCRIPTION IN ENGLISH": "MB BARCODE LABEL (28mm*8mm)",
            "QTY": "1,000",
            "UN": "PC",
            "LEVEL": 2,
            "_SUBMATERIAL": True
        },
        {
            "ITEM": "7353742",
            "MATERIAL": "L600006",
            "DESCRIPTION IN CHINESE": "",
            "DESCRIPTION IN ENGLISH": "RIBBON\\110mm*450m\\LOCAL 556",
            "QTY": "556",
            "UN": "",
            "LEVEL": 2,
            "_SUBMATERIAL": True
        }
    ]
    df_manuales = pd.DataFrame(filas_manuales, columns=df_sub_normales.columns)
    df_sub_normales = pd.concat([df_sub_normales, df_manuales], ignore_index=True)

    # ===== Insertar submateriales antes de X del bloque LEVEL 2 =====
    indices_level2 = [i for i, v in enumerate(df_main["LEVEL"]) if v == 2]
    for idx in reversed(indices_level2):
        x_index = None
        for j in range(idx, len(df_main)):
            if df_main.at[j, "ITEM"] == "X" and df_main.at[j, "LEVEL"] == 2:
                x_index = j
                break  
        if x_index is not None:
            df_main = pd.concat([
                df_main.iloc[:x_index],
                df_sub_normales,
                df_main.iloc[x_index:]
            ], ignore_index=True)

            # Aplicar color y fuente directamente en ws
            for i in range(len(df_sub_normales)):
                fila_excel = x_index + 2 + i
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=fila_excel, column=c).fill = gris_submaterial
                    ws.cell(row=fila_excel, column=c).font = fuente_normal
            break

    # ===== Agregar submateriales finales al final =====
    df_main = pd.concat([df_main, df_sub_finales], ignore_index=True)
    fila_inicio = len(df_main) - len(df_sub_finales) + 2
    for i in range(len(df_sub_finales)):
        fila_excel = fila_inicio + i
        for c in range(1, ws.max_column + 1):
            ws.cell(row=fila_excel, column=c).fill = gris_submaterial
            ws.cell(row=fila_excel, column=c).font = fuente_normal

    # ===== Limpiar columna temporal =====
    df_main.drop(columns=["PCB_CODE","_SUBMATERIAL"], inplace=True, errors="ignore")

    return df_main

# ================== PROCESO PRINCIPAL ==================
def procesar_archivo_principal_mainboard_2(
    ruta_excel_principal: str, 
    ruta_salida_principal: str,
    internal_model: str =""
    ):
    
    wb = openpyxl.load_workbook(ruta_excel_principal)
    ws = wb.active

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

    # ===== DataFrame =====
    df_main = pd.DataFrame(ws.values)
    df_main.columns = df_main.iloc[0]
    df_main = df_main[1:].reset_index(drop=True)

    df_main = agregar_submateriales(df_main, ws)

    filas_protegidas = {0,1,2}
    df_main["ITEM"] = df_main["ITEM"].apply(lambda v: str(v).strip() if v else "")
    df_main.loc[~df_main.index.isin(filas_protegidas), "LEVEL"] = 0

    contador_bloque = 10
    for i, val in df_main["ITEM"].items():
        if i in filas_protegidas:
            continue
        if val == "X":
            contador_bloque = 10
            continue
        df_main.at[i, "ITEM"] = str(contador_bloque)
        contador_bloque += 10

    nivel_actual = 1
    for i in range(len(df_main)):
        if i in filas_protegidas:
            continue
        if df_main.at[i, "ITEM"] == "X":
            nivel_actual += 1
        else:
            df_main.at[i, "LEVEL"] = nivel_actual + 1

    for i in range(1, len(df_main)):
        if i in filas_protegidas:
            continue
        if df_main.at[i, "LEVEL"] == 0:
            df_main.at[i, "LEVEL"] = df_main.at[i - 1, "LEVEL"]

    gris_submaterial = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

    for r_idx, fila in enumerate(df_main.itertuples(index=False), start=2):
        is_submaterial = getattr(fila, "_SUBMATERIAL", False)
        for c_idx, val in enumerate(fila, start=1):
            ws.cell(row=r_idx, column=c_idx).value = val
            if is_submaterial:
                ws.cell(row=r_idx, column=c_idx).fill = gris_submaterial

    if "_SUBMATERIAL" in df_main.columns:
        df_main.drop(columns=["_SUBMATERIAL"], inplace=True)

    bold_font = Font(bold=True)
    col_indices = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
    col_item = col_indices.get("ITEM")
    if col_item:
        for row in range(2, ws.max_row + 1):
            if str(ws.cell(row=row, column=col_item).value).strip() == "X":
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).font = bold_font

        ws.title = "BOMlist"
        ws["A2"] = "0"
        ws["F2"] = "1000"
        ws["F3"] = "1000"
        ws["J3"] = "HIMEX"
        ws["G3"] = "PC"

        texto_modelo = internal_model.strip() if internal_model else ""

        ws["E3"] = f"MAIN BOARD\\{texto_modelo}\\ROH"
        ws["E4"] = f"MAIN BOARD\\{texto_modelo}\\ROH"

        valor = ws["D5"].value

        if valor and "\\" in valor:
            parte = valor.split("\\", 1)[1]   
            ws["E5"] = "MAINBOARD SMT PART\\" + parte
        else:
            ws["E5"] = "MAINBOARD SMT PART\\"
    
    if "BOMHeader" not in wb.sheetnames:
        ws_header = wb.create_sheet("BOMHeader")
        encabezados_header = ["BOMID","MATNR","WERKS","STLAN","STLAL","ZTEXT","BMENG","STKTX"]
        for col, header in enumerate(encabezados_header, start=1):
            ws_header.cell(row=1, column=col).value = header
    if "BOMItem" not in wb.sheetnames:
        ws_item = wb.create_sheet("BOMItem")
        encabezados_item = ["BOMID","POSNR","POSTP","IDNRK","MENGE","MEINS","SORTF","POTX1","POTX2"]
        for col, header in enumerate(encabezados_item, start=1):
            ws_item.cell(row=1, column=col).value = header

    columnas_numericas = ["LEVEL", "ITEM", "QTY","MATERIAL"]
    mapa_columnas = {str(ws.cell(row=1, column=c).value).strip().upper(): openpyxl.utils.get_column_letter(c)
                     for c in range(1, ws.max_column + 1)
                     if str(ws.cell(row=1, column=c).value).strip().upper() in columnas_numericas}
    for nombre, letra in mapa_columnas.items():
        for cell in ws[letra][1:]:
            if cell.value is not None:
                valor = str(cell.value).strip()
                if valor != "":
                    try:
                        cell.value = float(valor.replace(",", ""))
                    except:
                        pass

    wb.save(ruta_salida_principal)
    print(f"[OK] Proceso completo {ruta_salida_principal}")