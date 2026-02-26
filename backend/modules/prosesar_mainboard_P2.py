import os
import pandas as pd

from backend.utils.txt_to_xlsx import (
    exportar_bom_a_xls,
    convertir_xls_a_xlsx,
    MAINBOARD_2_FILES_FOLDER
)

from backend.config.sap_config import TRANSACCION
from backend.utils.sap_utils import acceso_bom_exitoso

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
            # SAP a veces exporta HTML disfrazado de XLS
            try:
                return pd.read_html(ruta_xls)[0]
            except Exception as e:
                print(f"[WARNING] No se pudo leer XLS original: {e}")
                return pd.DataFrame()  # retorna vacío para continuar flujo


def procesar_material_desde_mainboard(session, ruta_mainboard_xlsx, uso, plantas):
    ruta_mainboard_xlsx = str(ruta_mainboard_xlsx)

    if not os.path.exists(ruta_mainboard_xlsx):
        raise FileNotFoundError(f"No existe el archivo mainboard: {ruta_mainboard_xlsx}")

    #! ===== LEER MAINBOARD NIVEL 1 =====
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

    print(f"[INFO] Material detectado desde mainboard: {material}")

    rutas_finales = []

    #! ===== CS11 POR CADA PLANTA =====
    for planta in plantas:
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION
            session.findById("wnd[0]").sendVKey(0)

            session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = planta
            session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = material
            session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = uso
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            if not acceso_bom_exitoso(session):
                print(f"[WARNING] No se pudo acceder al BOM de {material} en planta {planta}")
                continue

            #! ===== EXPORTAR BOM =====
            ruta_xls = exportar_bom_a_xls(session=session, material=material, mainboard=False)
            if not ruta_xls or not os.path.exists(ruta_xls):
                print(f"[WARNING] Falló exportación BOM planta {planta}")
                continue

            #! ===== CONVERTIR XLS → XLSX =====
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

        except Exception as e:
            print(f"[ERROR] Planta {planta}: {e}")
            continue

    if not rutas_finales:
        print(f"[ERROR] No se pudo procesar el BOM de {material} en ninguna planta")
        return []

    return rutas_finales