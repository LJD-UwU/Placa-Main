import os
import shutil
import pandas as pd
import xlwings as xw
from tkinter import filedialog
from openpyxl import load_workbook
from backend.config.sap_config import TRANSACCION
from backend.utils.sap_utils import acceso_bom_exitoso
from backend.utils.txt_to_xlsx import (exportar_bom_a_xls,convertir_xls_a_xlsx,MAINBOARD_2_FILES_FOLDER)

def limpiar_valor(valor):
    
    if valor is None:
        return ""
    valor = str(valor).strip()
    if valor.endswith(".0"):
        valor = valor[:-2]
    return valor


def actualizar_excel_mainboard(mother, materiales, ruta_excel, app=None, desc_mother="", desc_main=""):
    import xlwings as xw

    cerrar_app = False

    if app is None:
        app = xw.App(visible=False)
        cerrar_app = True

    try:
        wb = app.books.open(ruta_excel)
        sheet = wb.sheets.active

        data = sheet.used_range.value

        if not data:
            raise ValueError("El archivo Excel está vacío")

        #! Forzar estructura lista de listas
        if not isinstance(data[0], list):
            data = [data]

        header = [limpiar_valor(h).upper() for h in data[0]]

        try:
            col_mother_pn = header.index("MOTHERBOARD PART NUMBER") + 1
            col_main_pn = header.index("MAINBOARD PART NUMBER") + 1
            
            try:
                col_mother_desc = header.index("MOTHERBOARD DESCR") + 1
            except ValueError:
                col_mother_desc = None
                
            try:
                col_main_desc = header.index("MAINBOARD DESCR") + 1
            except ValueError:
                col_main_desc = None
                
        except ValueError:
            raise Exception("No se encontraron las columnas necesarias.")

        fila_objetivo = None

        for i, row in enumerate(data[1:], start=2):
            if not row:
                continue

            material = limpiar_valor(row[col_mother_pn - 1])
            valor_actual = row[col_main_pn - 1]

            if material == limpiar_valor(mother):
                if not valor_actual:
                    fila_objetivo = i
                    break

        if not fila_objetivo:
            raise Exception(f"No hay fila disponible para el modelo '{mother}'.")

        #! Escritura optimizada
        resultado = ", ".join(materiales) if materiales else "NOT FOUND"
        sheet.cells(fila_objetivo, col_main_pn).value = resultado
        
        if col_mother_desc and desc_mother:
            sheet.cells(fila_objetivo, col_mother_desc).value = desc_mother
            
        if col_main_desc and desc_main:
            sheet.cells(fila_objetivo, col_main_desc).value = desc_main

        wb.save()
        wb.close()

        print(f"[OK] Archivo actualizado: {ruta_excel}")

    finally:
        if cerrar_app:
            app.quit()

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

    material_row = df[df[columna_material].notna()].iloc[0]
    material = str(material_row[columna_material]).strip()

    posibles_desc = ["DESCRIPTION IN CHINESE", "DESCRIPCION", "MAKTX", "Description"]
    columna_desc = next((c for c in posibles_desc if c in df.columns), None)
    descripcion = str(material_row[columna_desc]).strip() if columna_desc else ""

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

        #! Generación de XLS/XLSX del BOM
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

        #! Extraer componentes
        materiales_extraidos = []
        try:
            df_bom = pd.read_excel(ruta_xlsx, engine="openpyxl")
            posibles_cols = ["Component", "Componente", "MATERIAL", "Material"]
            col_comp = next((c for c in posibles_cols if c in df_bom.columns), None)
            if col_comp:
                materiales_extraidos = df_bom[col_comp].dropna().astype(str).str.strip().unique().tolist()
        except Exception as e:
            print(f"[WARNING] No se pudieron extraer materiales: {e}")

        return ruta_xlsx, material, descripcion

    except Exception as e:
        print(f"[ERROR] Planta {planta}: {e}")
        return None

    except Exception as e:
        print(f"[ERROR] Planta {planta}: {e}")
        return None