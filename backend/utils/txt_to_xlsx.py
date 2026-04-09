import os
import re
import time
import pandas as pd
import xlwings as xw
import win32com.client
from datetime import datetime
from backend.config.sap_config import EXPORT_FINAL_PATH

#! CARPETA ORIGINAL
BASE_BOM_FOLDER = os.path.join(EXPORT_FINAL_PATH, "BOM_FILES")

#! CARPETAS SECUNDARIAS
MODEL_FILES_FOLDER = os.path.join(BASE_BOM_FOLDER, "1TE_FILES")
MAINBOARD_1_FILES_FOLDER = os.path.join(BASE_BOM_FOLDER, "MOTHERBOARD")
MAINBOARD_2_FILES_FOLDER = os.path.join(BASE_BOM_FOLDER, "MAIN_BOARD_FINAL")

#! CREAR UN ACCEDO DIRECTO AL DESKTOP PARA BOM FILES

def crear_estructura_y_acceso_directo():
    #!Crear carpetas
    os.makedirs(MODEL_FILES_FOLDER, exist_ok=True)
    os.makedirs(MAINBOARD_1_FILES_FOLDER, exist_ok=True)
    os.makedirs(MAINBOARD_2_FILES_FOLDER, exist_ok=True)

    print("[INFO] Carpetas creadas correctamente")

    #! Crear acceso directo en el escritorio
    escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
    ruta_acceso = os.path.join(escritorio, "BOM_FILES.lnk")

    #! Evitar duplicados
    if not os.path.exists(ruta_acceso):
        shell = win32com.client.Dispatch("WScript.Shell")
        acceso = shell.CreateShortCut(ruta_acceso)
        acceso.TargetPath = BASE_BOM_FOLDER
        acceso.WorkingDirectory = BASE_BOM_FOLDER
        acceso.IconLocation = "shell32.dll, 3"  #! icono de carpeta
        acceso.save()

        print("[INFO] Acceso directo creado en el escritorio")
    else:
        print("[INFO] El acceso directo ya existe")
        
crear_estructura_y_acceso_directo()

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

        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(ruta_xls)
        wb.api.SaveAs(ruta_csv, FileFormat=6) 
        wb.close()
        app.quit()
        df = pd.read_csv(ruta_csv, encoding="gb2312")

        #! Guardar como XLSX
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
        if ruta_csv and os.path.exists(ruta_csv):
            try:
                os.remove(ruta_csv)
            except:
                pass

#! Función: Exportar BOM desde SAP
def exportar_bom_a_xls(
    session, material, 
    mainboard=False,
    ):
    """
    Exporta el BOM de CS11 a XLS.
    Guarda automáticamente en la carpeta correspondiente.
    """
    nombre_limpio = re.sub(r'[\\/*?:"<>|]', "_", material)
    fecha = datetime.now().strftime("%Y-%m-%d-%M-%S")
    xls_name = f"{fecha}-{nombre_limpio}.XLS"

    carpeta_destino = MAINBOARD_1_FILES_FOLDER if mainboard else MODEL_FILES_FOLDER
    ruta_xls_final = os.path.join(carpeta_destino, xls_name)

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

#! Espera de archivo
def esperar_archivo(path, timeout=60):
    inicio = time.time()
    while time.time() - inicio < timeout:
        if os.path.exists(path) and os.path.getsize(path) > 0:
            time.sleep(1)
            return True
        time.sleep(1)
    return False
