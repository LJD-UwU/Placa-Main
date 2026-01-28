import os
import pandas as pd
from backend.utils.txt_to_xlsx import exportar_bom_a_xls, convertir_xls_a_xlsx, MAINBOARD_2_FILES_FOLDER
from backend.utils.sap_utils import acceso_bom_exitoso
from backend.utils.clean_excel import limpiar_excel_mainboard  # 🔹 Importar limpieza

def procesar_material_desde_mainboard(session, ruta_mainboard_xlsx, uso):
    """
    Procesa el segundo nivel de material desde un mainboard ya generado (nivel 1).
    
    Parámetros:
    - session: sesión de SAP abierta
    - ruta_mainboard_xlsx: ruta completa del archivo XLSX del mainboard nivel 1
    - uso: CAPID (uso del material en SAP)
    
    Flujo:
    1️⃣ Leer el primer material del archivo mainboard nivel 1
    2️⃣ Abrir CS11 en SAP con ese material
    3️⃣ Exportar BOM a MAINBOARD_2_FILES
    4️⃣ Limpiar automáticamente el Excel resultante
    """

    if not os.path.exists(ruta_mainboard_xlsx):
        raise FileNotFoundError(f"No existe el archivo mainboard: {ruta_mainboard_xlsx}")

    # Leer el mainboard nivel 1
    df = pd.read_excel(ruta_mainboard_xlsx)

    posibles_columnas = ["MATERIAL", "Material", "MATNR", "Component", "Componente"]
    columna_material = next((c for c in posibles_columnas if c in df.columns), None)

    if not columna_material:
        raise Exception("No se encontró columna MATERIAL en mainboard")

    # Tomar el primer material
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

    # Convertir a XLSX y guardar en MAINBOARD_2_FILES
    ruta_xlsx = os.path.join(MAINBOARD_2_FILES_FOLDER, f"{material}.xlsx")
    convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)

    # 🔹 Limpiar automáticamente el Excel generado
    limpiar_excel_mainboard(ruta_xlsx)

    print(f"[OK] Nivel 2 exportado y limpiado correctamente: {ruta_xlsx}")
    return ruta_xlsx
