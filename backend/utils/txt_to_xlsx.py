import os
import time
import re
import xlwings as xw
import pandas as pd
from backend.config.sap_config import EXPORT_FINAL_PATH

# CARPETAS
MODEL_FILES_FOLDER = os.path.join(EXPORT_FINAL_PATH, "MODEL_BOM")
MAINBOARD_1_FILES_FOLDER = os.path.join(EXPORT_FINAL_PATH, "MOTHERBOARD__BOM")
MAINBOARD_2_FILES_FOLDER = os.path.join(EXPORT_FINAL_PATH, "MAINBOARD_FINAL_BOM")

os.makedirs(MODEL_FILES_FOLDER, exist_ok=True)
os.makedirs(MAINBOARD_1_FILES_FOLDER, exist_ok=True)
os.makedirs(MAINBOARD_2_FILES_FOLDER, exist_ok=True) 

# CONVERSIÓN XLS (CP936) → CSV → XLSX

def convertir_xls_a_xlsx(ruta_xls: str, ruta_xlsx: str):
    if not os.path.exists(ruta_xls):
        print(f"[ERROR] No existe el archivo XLS: {ruta_xls}")
        return None

    ruta_csv = None
    app = None

    try:
        carpeta = os.path.dirname(ruta_xls)
        base = os.path.splitext(os.path.basename(ruta_xls))[0]
        ruta_csv = os.path.join(carpeta, f"{base}.csv")

        # 1️⃣ Abrir Excel invisible
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(ruta_xls)

        # 2️⃣ Guardar como CSV (CP936)
        wb.api.SaveAs(ruta_csv, FileFormat=6)  # xlCSV
        wb.close()
        app.quit()

        # 3️⃣ Leer CSV en chino
        df = pd.read_csv(ruta_csv, encoding="gb2312")

        # 4️⃣ Guardar como XLSX 
        df.to_excel(ruta_xlsx, index=False, engine="openpyxl")

        print(f"[OK] XLS → XLSX (CP936 preservado): {ruta_xlsx}")
        return ruta_xlsx

    except Exception as e:
        print(f"[ERROR] Falló conversión XLS → XLSX: {e}")
        try:
            if app:
                app.quit()
        except:
            pass
        return None

    finally:
        # 🧹 LIMPIEZA DEL CSV TEMPORAL
        if ruta_csv and ruta_xls and os.path.exists(ruta_csv) and os.path.exists(ruta_xls):
            try:
                os.remove(ruta_csv)
                os.remove(ruta_xls)
            except:
                pass
            


# EXPORTAR BOM DESDE SAP

def exportar_bom_a_xls(session, material, mainboard=False):
    """
    Exporta el BOM de CS11 a XLS (CP936).
    Guarda en carpeta de MODELOS o MAINBOARD.
    """
    nombre_limpio = re.sub(r'[\\/*?:"<>|]', "_", material)
    xls_name = f"{nombre_limpio}.XLS"

    carpeta_destino = MAINBOARD_1_FILES_FOLDER if mainboard else MODEL_FILES_FOLDER
    ruta_xls_final = os.path.join(carpeta_destino, xls_name)

    try:
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        time.sleep(1)

        # Opción Spreadsheet si aparece
        try:
            session.findById(
                "wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]"
            ).select()
        except:
            pass

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1)

        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = xls_name
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = carpeta_destino
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        if not esperar_archivo(ruta_xls_final, timeout=60):
            raise RuntimeError("SAP no generó el archivo XLS")

        print(f"[OK] XLS exportado correctamente: {ruta_xls_final}")
        return ruta_xls_final

    except Exception as e:
        print(f"[ERROR] Exportación SAP falló ({material}): {e}")
        return None


# ESPERA DE ARCHIVO

def esperar_archivo(path, timeout=60):
    inicio = time.time()
    while time.time() - inicio < timeout:
        if os.path.exists(path) and os.path.getsize(path) > 0:
            time.sleep(1)
            return True
        time.sleep(1)
    return False
