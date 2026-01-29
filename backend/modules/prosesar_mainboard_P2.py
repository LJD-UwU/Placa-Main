import os
import pandas as pd
from backend.utils.txt_to_xlsx import exportar_bom_a_xls, convertir_xls_a_xlsx
from backend.utils.sap_utils import acceso_bom_exitoso
from backend.utils.clean_excel_p2 import limpiar_excel_mainboard_2
from backend.utils.txt_to_xlsx import MAINBOARD_2_FILES_FOLDER

def procesar_material_desde_mainboard(session, ruta_mainboard_xlsx, uso):

    ruta_mainboard_xlsx = str(ruta_mainboard_xlsx)  # Asegurarse que sea str

    if not os.path.exists(ruta_mainboard_xlsx):
        raise FileNotFoundError(f"No existe el archivo mainboard: {ruta_mainboard_xlsx}")

    df = pd.read_excel(ruta_mainboard_xlsx)

    posibles_columnas = ["MATERIAL", "Material", "MATNR", "Component", "Componente"]
    columna_material = next((c for c in posibles_columnas if c in df.columns), None)

    if not columna_material:
        raise Exception("No se encontró columna MATERIAL en mainboard")

    material = str(df[columna_material].dropna().iloc[0]).strip()
    print(f"[INFO] Material detectado desde mainboard: {material}")

    # Abrir SAP CS11
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCS11"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = material
    session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = "2000"
    session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = uso
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    if not acceso_bom_exitoso(session):
        raise Exception(f"No se accedió al BOM de {material}")

    # Exportar segundo nivel → MAINBOARD_2_FILES
    ruta_xls = exportar_bom_a_xls(session=session, material=material, mainboard=False)
    if not ruta_xls:
        raise Exception("Falló exportación SAP")

    ruta_xlsx = os.path.join(MAINBOARD_2_FILES_FOLDER, f"{material}.xlsx")
    ruta_xlsx_str = str(ruta_xlsx)  # <-- Convertir Path a str antes de usar
    convertir_xls_a_xlsx(ruta_xls, ruta_xlsx_str)

    # Limpiar automáticamente el Excel generado
    limpiar_excel_mainboard_2(ruta_xlsx_str)

    print(f"[OK] Nivel 2 exportado y limpiado correctamente: {ruta_xlsx_str}")
    return ruta_xlsx_str
