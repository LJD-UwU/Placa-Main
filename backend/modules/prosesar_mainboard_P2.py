import os
import pandas as pd
from backend.utils.txt_to_xlsx import (
    exportar_bom_a_xls,
    convertir_xls_a_xlsx,
    MAINBOARD_2_FILES_FOLDER
)
from backend.config.sap_config import TRANSACCION
from backend.utils.sap_utils import acceso_bom_exitoso

def procesar_material_desde_mainboard(session, ruta_mainboard_xlsx, uso, plantas):
    """
    Flujo completo nivel 2:
    1) Detecta material desde mainboard nivel 1
    2) Accede a CS11 en SAP por cada planta
    3) Exporta BOM nivel 2
    4) Convierte XLS → XLSX
    5) Limpia estructura base
    6) Inserta submateriales desde BOM
    7) Devuelve lista de rutas XLSX procesadas
    """

    ruta_mainboard_xlsx = str(ruta_mainboard_xlsx)

    if not os.path.exists(ruta_mainboard_xlsx):
        raise FileNotFoundError(f"No existe el archivo mainboard: {ruta_mainboard_xlsx}")

    # Leer Excel asegurando engine openpyxl
    try:
        df = pd.read_excel(ruta_mainboard_xlsx, engine="openpyxl")
    except Exception as e:
        raise Exception(f"Error leyendo {ruta_mainboard_xlsx}: {e}")

    if df.empty:
        raise Exception("El archivo mainboard está vacío")

    # Buscar columna de material
    posibles_columnas = ["MATERIAL", "Material", "MATNR", "Component", "Componente"]
    columna_material = next((c for c in posibles_columnas if c in df.columns), None)
    if not columna_material:
        raise Exception("No se encontró columna MATERIAL en el mainboard")

    material = str(df[columna_material].dropna().iloc[0]).strip()
    if not material:
        raise Exception("Material detectado vacío")

    print(f"[INFO] Material detectado desde mainboard: {material}")

    rutas_procesadas = []

    # Acceso SAP CS11 por cada planta
    for planta in plantas:
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION
            session.findById("wnd[0]").sendVKey(0)

            session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = material
            session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = planta
            session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = uso
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            if not acceso_bom_exitoso(session):
                print(f"[WARNING] No se pudo acceder al BOM de {material} en planta {planta}")
                continue

            # Exportar BOM nivel 2
            ruta_xls = exportar_bom_a_xls(session=session, material=material, mainboard=False)
            if not ruta_xls or not os.path.exists(ruta_xls):
                print(f"[WARNING] Falló la exportación del BOM desde SAP para planta {planta}")
                continue

            # Convertir a XLSX
            nombre_xlsx = f"{material}_{planta}.xlsx"
            ruta_xlsx = os.path.join(MAINBOARD_2_FILES_FOLDER, nombre_xlsx)
            convertir_xls_a_xlsx(str(ruta_xls), str(ruta_xlsx))

            print(f"[OK] Mainboard nivel 2 procesado COMPLETO: {ruta_xlsx}")
            rutas_procesadas.append(ruta_xlsx)

        except Exception as e:
            print(f"[ERROR] Planta {planta}: {e}")
            continue

    if not rutas_procesadas:
        raise Exception(f"No se pudo procesar el BOM de {material} en ninguna planta")

    return rutas_procesadas
