import os
import pandas as pd
import shutil

from backend.utils.txt_to_xlsx import (
    exportar_bom_a_xls,
    convertir_xls_a_xlsx,
    MAINBOARD_2_FILES_FOLDER
)

from backend.config.sap_config import TRANSACCION
from backend.utils.sap_utils import acceso_bom_exitoso


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


def procesar_material_desde_mainboard(session, ruta_mainboard_xlsx, uso, planta):

    ruta_mainboard_xlsx = str(ruta_mainboard_xlsx)

    if not os.path.exists(ruta_mainboard_xlsx):
        raise FileNotFoundError(f"No existe el archivo mainboard: {ruta_mainboard_xlsx}")

    # ===== LEER MAINBOARD NIVEL 1 =====
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

    try:

        # ===== ENTRAR A TRANSACCIÓN =====
        session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = planta
        session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = material
        session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = uso

        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        if not acceso_bom_exitoso(session):
            print(f"[WARNING] No se pudo acceder al BOM {material} planta {planta}")
            return None

        # ===== NOMBRE ARCHIVOS =====
        nombre_xls = f"{material}_{planta}.xls"
        nombre_xlsx = f"{material}.xlsx"

        ruta_xls_destino = os.path.join(MAINBOARD_2_FILES_FOLDER, nombre_xls)
        ruta_xlsx = os.path.join(MAINBOARD_2_FILES_FOLDER, nombre_xlsx)

        # ===== VERIFICAR SI YA EXISTE =====
        if os.path.exists(ruta_xls_destino):

            print(f"[INFO] XLS ya existe: {ruta_xls_destino}")
            ruta_xls = ruta_xls_destino

        else:

            # ===== EXPORTAR DESDE SAP =====
            ruta_xls = exportar_bom_a_xls(
                session=session,
                material=material,
                mainboard=False
            )

            if not ruta_xls or not os.path.exists(ruta_xls):
                print(f"[WARNING] Falló exportación BOM {material}")
                return None

            try:
                shutil.move(ruta_xls, ruta_xls_destino)
                ruta_xls = ruta_xls_destino
                print(f"[INFO] XLS movido a {ruta_xls}")
            except Exception as e:
                print(f"[WARNING] No se pudo mover el XLS: {e}")

        # ===== CONVERTIR XLS → XLSX =====
        try:

            convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)

        except Exception as e:

            print(f"[WARNING] Conversión falló, usando fallback: {e}")

            df_temp = leer_excel_sap_fallback(ruta_xls)

            if df_temp.empty:
                print("[WARNING] XLS vacío, creando archivo base")
                df_temp = pd.DataFrame()

            df_temp.to_excel(ruta_xlsx, index=False)

        print(f"[INFO] BOM generado correctamente: {ruta_xlsx}")

        return ruta_xlsx

    except Exception as e:

        print(f"[ERROR] Planta {planta}: {e}")
        return None