import os
import xlwings as xw
import pandas as pd
from backend.config.sap_config import EXPORT_FINAL_PATH

# --- Carpetas para los archivos ---
XLS_FOLDER = os.path.join(EXPORT_FINAL_PATH, "XLS")
CSV_FOLDER = os.path.join(EXPORT_FINAL_PATH, "CSV")
XLSX_FOLDER = os.path.join(EXPORT_FINAL_PATH, "XLSX")

os.makedirs(XLS_FOLDER, exist_ok=True)
os.makedirs(CSV_FOLDER, exist_ok=True)
os.makedirs(XLSX_FOLDER, exist_ok=True)


def convertir_xls_a_csv_y_xlsx(ruta_xls: str, ruta_csv: str, ruta_xlsx: str):
    """
    Convierte un archivo XLS a CSV y luego a XLSX.
    - ruta_xls: archivo original
    - ruta_csv: ruta destino CSV
    - ruta_xlsx: ruta destino XLSX
    """
    try:
        if not os.path.exists(ruta_xls):
            print(f"[ERROR] No se encontró el XLS: {ruta_xls}")
            return None

        # --- Abrir XLS con xlwings ---
        wb = xw.Book(ruta_xls)

        # --- Guardar como CSV usando Excel API ---
        wb.api.SaveAs(ruta_csv, FileFormat=6)  # 6 = xlCSV
        wb.close()
        print(f"[INFO] CSV generado: {ruta_csv}")

        # --- Leer CSV con encoding GB2312 para soportar caracteres chinos ---
        df = pd.read_csv(ruta_csv, encoding="gb2312")

        # --- Guardar como XLSX ---
        df.to_excel(ruta_xlsx, index=False, engine="openpyxl")
        print(f"[INFO] XLSX generado: {ruta_xlsx}")

        return ruta_xlsx

    except Exception as e:
        print(f"[ERROR] Falló conversión XLS → CSV → XLSX: {e}")
        return None


def exportar_bom_a_xls(session, modelo):
    """
    Exporta el BOM de CS11 a XLS en la carpeta XLS_FOLDER.
    """
    import re
    nombre_limpio = re.sub(r'[\\/*?:"<>|]', "_", modelo)
    xls_name = f"{nombre_limpio}.XLS"
    ruta_xls_tmp = os.path.join(EXPORT_FINAL_PATH, xls_name)
    ruta_xls_final = os.path.join(XLS_FOLDER, xls_name)

    try:
        session.findById("wnd[0]").maximize()

        # --- Exportar desde SAP a XLS ---
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = xls_name
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = EXPORT_FINAL_PATH
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # --- Esperar a que SAP genere el archivo ---
        if not esperar_archivo(ruta_xls_tmp, timeout=40):
            print(f"[ERROR] SAP no generó el XLS para {modelo}")
            return None

        # --- Mover a carpeta final XLS ---
        os.replace(ruta_xls_tmp, ruta_xls_final)
        print(f"[INFO] XLS guardado en: {ruta_xls_final}")
        return ruta_xls_final

    except Exception as e:
        print(f"[ERROR] Falló exportación BOM {modelo}: {e}")
        return None


def esperar_archivo(path, timeout=60):
    """
    Espera hasta que el archivo exista y tenga contenido.
    """
    import time
    inicio = time.time()
    while time.time() - inicio < timeout:
        if os.path.exists(path) and os.path.getsize(path) > 0:
            time.sleep(1)  # buffer extra
            return True
        time.sleep(1)
    return False
