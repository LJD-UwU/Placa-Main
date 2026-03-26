import os
import shutil
import pandas as pd
from openpyxl import load_workbook

from backend.config.sap_config import TRANSACCION
from backend.utils.sap_utils import acceso_bom_exitoso
from backend.utils.txt_to_xlsx import (exportar_bom_a_xls,convertir_xls_a_xlsx,MAINBOARD_2_FILES_FOLDER)

def leer_excel_sap_fallback(ruta_xls):
    """
    Intenta leer cualquier archivo SAP exportado, aunque tenga formato extraño.
    """
    try:
        return pd.read_excel(ruta_xls, engine='openpyxl')
    except Exception:
        try:
            return pd.read_excel(ruta_xls, engine='xlrd')
        except Exception:
            try:
                return pd.read_html(ruta_xls)[0]
            except Exception as e:
                print(f"[WARNING] No se pudo leer XLS original: {e}")
                return pd.DataFrame()

def actualizar_excel_mainboard_2(ruta_excel, modelo, materiales):
    import xlwings as xw
    app = xw.App(visible=False)
    try:
        wb = app.books.open(ruta_excel)
        sheet = wb.sheets.active
        
        # Leer datos para buscar columnas
        data = sheet.used_range.value
        if not data:
            raise ValueError("El archivo Excel está vacío")
            
        header = [str(h).strip().upper() for h in data[0]]
        try:
            col_material = header.index("MATERIAL") + 1
            col_mainboard = header.index("MAINBOARD PART NUMBER") + 1
        except ValueError:
            raise ValueError("No se encontraron columnas 'MATERIAL' o 'MAINBOARD PART NUMBER'")

        fila_objetivo = None
        # Buscar fila vacía del mismo modelo
        for i, row in enumerate(data[1:], start=2):
            material = str(row[col_material - 1]).strip()
            # row index matches enumerate start=2
            valor_actual = row[col_mainboard - 1]

            if material == str(modelo).strip():
                if not valor_actual:
                    fila_objetivo = i
                    break

        if not fila_objetivo:
            app.quit()
            raise Exception(f"No hay fila disponible para {modelo}")

        if materiales:
            sheet.cells(fila_objetivo, col_mainboard).value = ", ".join(materiales)
        else:
            sheet.cells(fila_objetivo, col_mainboard).value = "NOT FOUND"

        wb.save()
        wb.close()
    finally:
        app.quit()

def procesar_material_desde_mainboard(session, ruta_mainboard_xlsx, uso, planta):
    """
    Procesa un material para una sola planta.
    Retorna la ruta del archivo XLSX generado, o None si falló.
    """
    ruta_mainboard_xlsx = str(ruta_mainboard_xlsx)

    if not os.path.exists(ruta_mainboard_xlsx):
        raise FileNotFoundError(f"No existe el archivo mainboard: {ruta_mainboard_xlsx}")

    #!  LEER MAINBOARD NIVEL 1 
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

    print(f"[INFO] Material detectado desde mainboard: {material}, Planta: {planta}")
    materiales_detectados = []

    if material:
        materiales_detectados.append(material)
    try:
        #!  ENTRAR A TRANSACCIÓN SAP 
        session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = planta
        session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = material
        session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = uso
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        if not acceso_bom_exitoso(session):
            print(f"[WARNING] No se pudo acceder al BOM de {material} en planta {planta}")
            return None

        #!  VERIFICAR SI YA EXISTE XLS 
        nombre_xls_esperado = f"{material}_{planta}.xls"
        ruta_xls_destino = os.path.join(MAINBOARD_2_FILES_FOLDER, nombre_xls_esperado)

        if os.path.exists(ruta_xls_destino):
            print(f"[INFO] XLS ya existe, no se descargará de SAP: {ruta_xls_destino}")
            ruta_xls = ruta_xls_destino

        else:
            #!  EXPORTAR BOM DESDE SAP 
            ruta_xls = exportar_bom_a_xls(
                session=session,
                material=material,
                mainboard=False
            )

            if not ruta_xls or not os.path.exists(ruta_xls):
                print(f"[WARNING] Falló exportación BOM planta {planta}")
                return None

            #!  MOVER XLS A MAINBOARD_2_FILES_FOLDER 
            try:
                shutil.move(ruta_xls, ruta_xls_destino)
                ruta_xls = ruta_xls_destino
                print(f"[INFO] XLS movido a: {ruta_xls}")
            except Exception as e:
                print(f"[WARNING] No se pudo mover el XLS: {e}")

        #!  CONVERTIR XLS → XLSX 
        nombre_base = f"{material}"
        ruta_xlsx = os.path.join(MAINBOARD_2_FILES_FOLDER, f"{nombre_base}.xlsx")

        try:
            convertir_xls_a_xlsx(str(ruta_xls), str(ruta_xlsx))
        except Exception as e:
            print(f"[WARNING] Convertir XLS→XLSX falló, usando fallback: {e}")
            df_temp = leer_excel_sap_fallback(ruta_xls)
            if df_temp.empty:
                print("[WARNING] Archivo XLS no pudo ser leído, se continuará con limpieza base vacía")
                df_temp = pd.DataFrame()
            df_temp.to_excel(ruta_xlsx, index=False)

        print(f"[INFO] BOM procesado correctamente: {ruta_xlsx}")
        return material

    except Exception as e:
        print(f"[ERROR] Planta {planta}: {e}")
        return None