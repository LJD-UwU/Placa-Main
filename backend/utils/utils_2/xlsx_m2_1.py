import os
import re
import time
import pandas as pd
import xlwings as xw
from datetime import datetime
from backend.utils.txt_to_xlsx import MAINBOARD_2_FILES_FOLDER


def convertir_xls_a_xlsx_1(ruta_xls: str, ruta_xlsx: str):
    if not os.path.exists(ruta_xls):
        print(f"[ERROR] No existe el archivo XLS: {ruta_xls}")
        return None

    ruta_csv1 = None
    app1 = None

    try:
        carpeta = os.path.dirname(ruta_xls)
        base = os.path.splitext(os.path.basename(ruta_xls))[0]
        ruta_csv1 = os.path.join(carpeta, f"{base}.csv")
        app1 = xw.App(visible=False)
        app1.display_alerts = False
        app1.screen_updating = False
        wb = app1.books.open(ruta_xls)
        wb.api.SaveAs(ruta_csv1, FileFormat=6)
        wb.close()
        app1.quit()
        df = pd.read_csv(ruta_csv1, encoding="gb2312")

        #! Guardar como XLSX en la carpeta nueva
        df.to_excel(ruta_xlsx, index=False, engine="openpyxl")
        print(f"[OK] XLS → XLSX (CP936 preservado): {ruta_xlsx}")
        return ruta_xlsx
    except Exception as e:
        print(f"[ERROR] Falló conversión XLS → XLSX: {e}")
        try:
            if app1:
                app1.quit()
        except:
            pass
        return None
    finally:
        if ruta_csv1 and os.path.exists(ruta_csv1):
            try:
                os.remove(ruta_csv1)
            except:
                pass


#! Función: Exportar BOM desde SAP
def exportar_bom_a_xls_1(session, material):
    nombre_limpio_1 = re.sub(r'[\\/*?:"<>|]', "_", material)
    xls_name_1 = f"{nombre_limpio_1}.XLS"
    carpeta_destino_1 = MAINBOARD_2_FILES_FOLDER
    ruta_xls_final_1 = os.path.join(carpeta_destino_1, xls_name_1)
    try:

        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        time.sleep(1)
        
        #! Opción Spreadsheet si aparece
        try:
            session.findById(
                "wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]"
            ).select()
        except:
            pass

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1)
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = xls_name_1
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = carpeta_destino_1
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        if not esperar_archivo_1(ruta_xls_final_1, timeout=60):
            raise RuntimeError("SAP no generó el archivo XLS")
        print(f"[OK] XLS exportado correctamente: {ruta_xls_final_1}")
        return ruta_xls_final_1
    except Exception as e:
        print(f"[ERROR] Exportación SAP falló ({material}): {e}")
        return None


#! Espera de archivo
def esperar_archivo_1(path, timeout=60):
    inicio1 = time.time()
    while time.time() - inicio1 < timeout:
        if os.path.exists(path) and os.path.getsize(path) > 0:
            time.sleep(1)
            return True
        time.sleep(1)
    return False