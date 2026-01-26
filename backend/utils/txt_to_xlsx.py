import os
import xlwings as xw
import pandas as pd
from backend.config.sap_config import EXPORT_FINAL_PATH

# --- Carpetas para los archivos ---
XLS_FOLDER = os.path.join(EXPORT_FINAL_PATH, "XLS")
XLSX_FOLDER = os.path.join(EXPORT_FINAL_PATH, "XLSX")

os.makedirs(XLS_FOLDER, exist_ok=True)
os.makedirs(XLSX_FOLDER, exist_ok=True)


import os
import xlwings as xw

import os
import xlwings as xw
import pandas as pd

def convertir_xls_a_xlsx(ruta_xls: str, ruta_xlsx: str):
    """
    Conversión segura SAP:
    XLS (CodePage 936) → CSV (Excel) → XLSX (Unicode)
    SIN perder caracteres chinos.
    """
    try:
        if not os.path.exists(ruta_xls):
            print(f"[ERROR] No existe el archivo XLS: {ruta_xls}")
            return None

        carpeta = os.path.dirname(ruta_xls)
        base = os.path.splitext(os.path.basename(ruta_xls))[0]
        ruta_csv = os.path.join(carpeta, f"{base}.csv")

        # 🧠 1) Abrir XLS con Excel
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(ruta_xls)

        # 🔥 2) Guardar como CSV desde Excel (mantiene 936)
        wb.api.SaveAs(ruta_csv, FileFormat=6)  # 6 = xlCSV
        wb.close()
        app.quit()

        # 🧠 3) Leer CSV con encoding chino
        df = pd.read_csv(ruta_csv, encoding="gb2312")

        # 🧠 4) Exportar a XLSX (Unicode real)
        df.to_excel(ruta_xlsx, index=False, engine="openpyxl")

        print(f"[OK] Conversión correcta XLS → XLSX (GB2312 preservado): {ruta_xlsx}")
        return ruta_xlsx

    except Exception as e:
        print(f"[ERROR] Falló conversión XLS → XLSX: {e}")
        try:
            app.quit()
        except:
            pass
        return None


def exportar_bom_a_xls(session, modelo):
    """
    Exporta el BOM de CS11 a XLS (CodePage 936).
    """
    import re
    import time

    nombre_limpio = re.sub(r'[\\/*?:"<>|]', "_", modelo)
    xls_name = f"{nombre_limpio}.XLS"
    ruta_xls_tmp = os.path.join(EXPORT_FINAL_PATH, xls_name)
    ruta_xls_final = os.path.join(XLS_FOLDER, xls_name)

    try:
        session.findById("wnd[0]").maximize()

        # 📤 Exportar lista
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        time.sleep(1)

        # 🧠 SAP puede mostrar diferentes pantallas
        try:
            # Opción "Spreadsheet"
            session.findById(
                "wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]"
            ).select()
        except:
            pass

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1)

        # 📁 Nombre y ruta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = xls_name
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = EXPORT_FINAL_PATH
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # ⏳ Esperar archivo
        if not esperar_archivo(ruta_xls_tmp, timeout=60):
            raise RuntimeError("SAP no generó el archivo XLS")

        # 📦 Mover a carpeta XLS
        os.replace(ruta_xls_tmp, ruta_xls_final)
        print(f"[OK] XLS exportado correctamente: {ruta_xls_final}")

        return ruta_xls_final

    except Exception as e:
        print(f"[ERROR] Exportación SAP falló ({modelo}): {e}")
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
