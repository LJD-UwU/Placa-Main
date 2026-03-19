import os
import pandas as pd
import shutil
from backend.utils.txt_to_xlsx import (
    exportar_bom_a_xls,
    convertir_xls_a_xlsx,
    MAINBOARD_2_FILES_FOLDER
)
from tkinter import filedialog
from openpyxl import load_workbook
from backend.config.sap_config import TRANSACCION
from backend.utils.sap_utils import acceso_bom_exitoso


from openpyxl import load_workbook

def actualizar_excel_mainboard(mother, materiales, ruta_excel):
    """
    Actualiza un archivo Excel en la columna 'MAINBOARD PART NUMBER' para el modelo dado.

    Args:
        mother (str): Modelo a actualizar.
        materiales (list): Lista de materiales a poner en la celda.
        ruta_excel (str): Ruta del archivo Excel a actualizar.
    """
    wb = load_workbook(ruta_excel)
    ws = wb.active

    col_material = None
    col_mainboard = None

    # Buscar las columnas necesarias
    for i, cell in enumerate(ws[1], start=1):
        nombre = str(cell.value).strip().upper()
        if nombre == "MOTHERBOARD PART NUMBER":
            col_material = i
        if nombre == "MAINBOARD PART NUMBER":
            col_mainboard = i

    if not col_material or not col_mainboard:
        raise Exception("No se encontraron las columnas necesarias en el Excel.")

    fila_objetivo = None

    # Buscar fila vacía correspondiente al modelo
    for row in ws.iter_rows(min_row=2):
        material = str(row[col_material - 1].value).strip()
        valor_actual = row[col_mainboard - 1].value
        if material == str(mother).strip() and not valor_actual:
            fila_objetivo = row[0].row
            break

    if not fila_objetivo:
        raise Exception(f"No hay fila disponible para el modelo '{mother}'.")

    # Escribir materiales o "NOT FOUND"
    ws.cell(row=fila_objetivo, column=col_mainboard).value = ", ".join(materiales) if materiales else "NOT FOUND"

    wb.save(ruta_excel)
    print(f"[OK] Archivo actualizado: {ruta_excel}")

def leer_excel_sap_fallback(ruta_xls):
    """
    Intenta leer cualquier archivo SAP exportado aunque tenga formato extraño.
    """
    try:
        return pd.read_excel(ruta_xls, engine="openpyxl")
    except:
        try:
            return pd.read_excel(ruta_xls, engine="xlrd")
        except:
            try:
                return pd.read_html(ruta_xls)[0]
            except Exception as e:
                print(f"[WARNING] No se pudo leer XLS original: {e}")
                return pd.DataFrame()


def procesar_material_desde_mainboard(session, ruta_mainboard_xlsx, uso, planta, mother):

    ruta_mainboard_xlsx = str(ruta_mainboard_xlsx)

    if not os.path.exists(ruta_mainboard_xlsx):
        raise FileNotFoundError(f"No existe el archivo mainboard: {ruta_mainboard_xlsx}")

    df = pd.read_excel(ruta_mainboard_xlsx, engine="openpyxl")

    if df.empty:
        raise Exception("El archivo mainboard está vacío")

    posibles_columnas = ["MATERIAL", "Material", "MATNR", "Component", "Componente"]
    columna_material = next((c for c in posibles_columnas if c in df.columns), None)

    if not columna_material:
        raise Exception("No se encontró columna MATERIAL en el mainboard")

    material = str(df[columna_material].dropna().iloc[0]).strip()

    if not material:
        raise Exception("Material detectado vacío")

    print(f"[INFO] Material detectado: {material} | Planta: {planta}")
    print(f"[DEBUG] Archivo cargado: {ruta_mainboard_xlsx}")

    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = planta
        session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = material
        session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = uso
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        if not acceso_bom_exitoso(session):
            print(f"[WARNING] No se pudo acceder al BOM {material}")
            return None

        # Generación de XLS/XLSX del BOM
        nombre_xls = f"{material}_{planta}.xls"
        nombre_xlsx = f"{material}.xlsx"
        ruta_xls_destino = os.path.join(MAINBOARD_2_FILES_FOLDER, nombre_xls)
        ruta_xlsx = os.path.join(MAINBOARD_2_FILES_FOLDER, nombre_xlsx)

        if not os.path.exists(ruta_xls_destino):
            ruta_xls = exportar_bom_a_xls(session=session, material=material, mainboard=False)
            if not ruta_xls or not os.path.exists(ruta_xls):
                print(f"[WARNING] Falló exportación BOM {material}")
                return None
            shutil.move(ruta_xls, ruta_xls_destino)
            ruta_xls = ruta_xls_destino
        else:
            ruta_xls = ruta_xls_destino

        try:
            convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
        except:
            df_temp = leer_excel_sap_fallback(ruta_xls)
            df_temp.to_excel(ruta_xlsx, index=False)

        print(f"[INFO] BOM generado correctamente: {ruta_xlsx}")

        # Extraer componentes
        materiales_extraidos = []
        try:
            df_bom = pd.read_excel(ruta_xlsx, engine="openpyxl")
            posibles_cols = ["Component", "Componente", "MATERIAL", "Material"]
            col_comp = next((c for c in posibles_cols if c in df_bom.columns), None)
            if col_comp:
                materiales_extraidos = df_bom[col_comp].dropna().astype(str).str.strip().unique().tolist()
        except Exception as e:
            print(f"[WARNING] No se pudieron extraer materiales: {e}")

        return ruta_xlsx, material

    except Exception as e:
        print(f"[ERROR] Planta {planta}: {e}")
        return None