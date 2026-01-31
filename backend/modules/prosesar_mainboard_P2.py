import os
import pandas as pd

from backend.utils.txt_to_xlsx import (
    exportar_bom_a_xls,
    convertir_xls_a_xlsx,
    MAINBOARD_2_FILES_FOLDER
)
from backend.utils.sap_utils import acceso_bom_exitoso
from backend.utils.clean_excel_p2 import (
    limpiar_excel_mainboard_2,
    procesar_archivo_principal_mainboard_2
)

# ============================================================
# PROCESAR MATERIAL DESDE MAINBOARD (NIVEL 2)
# ============================================================
def procesar_material_desde_mainboard(session, ruta_mainboard_xlsx, uso):
    """
    Flujo completo:
    1) Detecta material desde mainboard nivel 1
    2) Accede a CS11 en SAP
    3) Exporta BOM nivel 2
    4) Convierte XLS → XLSX
    5) Limpia estructura base
    6) Inserta submateriales desde BOM
    7) Devuelve archivo final listo
    """

    ruta_mainboard_xlsx = str(ruta_mainboard_xlsx)

    if not os.path.exists(ruta_mainboard_xlsx):
        raise FileNotFoundError(f"No existe el archivo mainboard: {ruta_mainboard_xlsx}")

    # --------------------------------------------------------
    # DETECTAR MATERIAL DESDE MAINBOARD
    # --------------------------------------------------------
    df = pd.read_excel(ruta_mainboard_xlsx)

    if df.empty:
        raise Exception("El archivo mainboard está vacío")

    posibles_columnas = [
        "MATERIAL", "Material", "MATNR", "Component", "Componente"
    ]

    columna_material = next(
        (c for c in posibles_columnas if c in df.columns),
        None
    )

    if not columna_material:
        raise Exception("No se encontró columna MATERIAL en el mainboard")

    material = str(df[columna_material].dropna().iloc[0]).strip()

    if not material:
        raise Exception("Material detectado vacío")

    print(f"[INFO] Material detectado desde mainboard: {material}")

    # --------------------------------------------------------
    # ACCESO SAP CS11
    # --------------------------------------------------------
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCS11"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = material
    session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = "2000"
    session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = uso
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    if not acceso_bom_exitoso(session):
        raise Exception(f"No se pudo acceder al BOM de {material}")

    # --------------------------------------------------------
    # EXPORTAR BOM NIVEL 2
    # --------------------------------------------------------
    ruta_xls = exportar_bom_a_xls(
        session=session,
        material=material,
        mainboard=False
    )

    if not ruta_xls:
        raise Exception("Falló la exportación del BOM desde SAP")

    # --------------------------------------------------------
    # CONVERTIR A XLSX
    # --------------------------------------------------------
    ruta_xlsx = os.path.join(
        MAINBOARD_2_FILES_FOLDER,
        f"{material}.xlsx"
    )

    convertir_xls_a_xlsx(str(ruta_xls), str(ruta_xlsx))

    # --------------------------------------------------------
    # LIMPIEZA BASE (HEADERS / COLUMNAS)
    # --------------------------------------------------------
    limpiar_excel_mainboard_2(str(ruta_xlsx))

    # --------------------------------------------------------
    # PROCESAMIENTO FINAL + SUBMATERIALES
    # (usa BOM automáticamente desde clean_excel_p2.py)
    # --------------------------------------------------------
    procesar_archivo_principal_mainboard_2(
        ruta_excel_principal=str(ruta_xlsx),
        ruta_salida_principal=str(ruta_xlsx)
    )

    print(f"[OK] Mainboard nivel 2 procesado COMPLETO: {ruta_xlsx}")
    return ruta_xlsx
