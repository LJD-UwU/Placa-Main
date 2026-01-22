import win32com.client
import time
import re
import os
import pandas as pd

from backend.config.sap_config import (
    SAP_LOGON_PATH,
    SAP_SYSTEM_NAME,
    SAP_USER,
    SAP_PASSWORD,
    SAP_TMP_PATH,
    EXPORT_FINAL_PATH
)

# ==============================
# UTILIDADES GENERALES SAP
# ==============================

def pausar(segundos=1):
    time.sleep(segundos)

def esperar_sap(session, timeout=15):
    for _ in range(timeout * 10):
        if not session.Busy:
            return
        time.sleep(0.1)
    raise Exception("SAP no respondió a tiempo")

def esperar_id(session, id_control, timeout=15):
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            return session.findById(id_control)
        except:
            time.sleep(0.2)
    raise Exception(f"Control no encontrado: {id_control}")

def escribir_campo(session, id_campo, texto):
    campo = esperar_id(session, id_campo)
    campo.text = texto
    campo.caretPosition = len(texto)
    return campo

def ejecutar_busqueda(session):
    esperar_sap(session)
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
    except:
        session.findById("wnd[0]").sendVKey(8)
    esperar_sap(session)

# ==============================
# CONEXIÓN SAP
# ==============================

def conectar_sap():
    try:
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except:
            import subprocess
            subprocess.Popen(SAP_LOGON_PATH)
            time.sleep(5)
            SapGuiAuto = win32com.client.GetObject("SAPGUI")

        app = SapGuiAuto.GetScriptingEngine

        connection = None
        for i in range(app.Children.Count):
            c = app.Children.Item(i)
            if SAP_SYSTEM_NAME.lower() in c.Description.lower():
                connection = c
                break

        if connection is None:
            connection = app.OpenConnection(SAP_SYSTEM_NAME, True)
            time.sleep(3)

        session = connection.Children.Item(0)

        # Login automático
        try:
            if session.findById("wnd[0]/usr/txtRSYST-BNAME").text == "":
                session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USER
                session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASSWORD
                session.findById("wnd[0]/tbar[0]/btn[0]").press()
                time.sleep(2)
        except:
            pass

        session.findById("wnd[0]").maximize()
        print("[INFO] Conectado a SAP ✅")
        return session

    except Exception as e:
        print(f"[ERROR] No se pudo conectar a SAP: {e}")
        return None

# ==============================
# EXPORTACIÓN BOM CS11
# ==============================

def exportar_bom_a_excel(session, modelo):
    """
    Exporta el BOM de CS11 a XLS, lo mueve a carpeta final
    y lo convierte a XLSX usando el nombre del modelo
    """

    nombre_limpio = re.sub(r'[\\/*?:"<>|]', "_", modelo)
    nombre_xls = f"{nombre_limpio}.XLS"

    ruta_tmp_xls = os.path.join(SAP_TMP_PATH, nombre_xls)
    ruta_final_xls = os.path.join(EXPORT_FINAL_PATH, nombre_xls)

    try:
        session.findById("wnd[0]").maximize()
        esperar_sap(session)

        # --- Exportar ---
        session.findById("wnd[0]/tbar[1]/btn[45]").press()
        esperar_sap(session)

        session.findById(
            "wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]"
        ).select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_xls
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = SAP_TMP_PATH
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        print(f"[INFO] Esperando exportación SAP: {ruta_tmp_xls}")

        if not esperar_archivo(ruta_tmp_xls, timeout=40):
            print("[ERROR] SAP no generó el XLS a tiempo")
            return None

        # --- MOVER A RUTA FINAL ---
        os.makedirs(EXPORT_FINAL_PATH, exist_ok=True)
        os.replace(ruta_tmp_xls, ruta_final_xls)

        print(f"[INFO] Archivo movido a: {ruta_final_xls}")

        # --- CONVERTIR usando MODELO ---
        return convertir_por_modelo(nombre_limpio)

    except Exception as e:
        print(f"[ERROR] Falló exportación BOM {modelo}: {e}")
        return None


# ==============================
# CONVERSIÓN XLS → XLSX
# ==============================

def convert_xls_to_xlsx(ruta_xls, eliminar_xls=True):
    try:
        if not os.path.exists(ruta_xls):
            print(f"[ERROR] No existe el archivo: {ruta_xls}")
            return None

        ruta_xlsx = ruta_xls.replace(".XLS", ".xlsx").replace(".xls", ".xlsx")

        df = pd.read_excel(ruta_xls, engine="xlrd")
        df.to_excel(ruta_xlsx, index=False, engine="openpyxl")

        if eliminar_xls:
            os.remove(ruta_xls)

        print(f"[INFO] Convertido a XLSX: {ruta_xlsx} ✅")
        return ruta_xlsx

    except Exception as e:
        print(f"[ERROR] Conversión XLS → XLSX falló: {e}")
        return None
    
    # ==============================
# UTILIDADES ESPECÍFICAS CS11
# ==============================

def esperar_cs11_completo(session, timeout=30):
    """
    Espera a que el grid de CS11 cargue completamente
    y devuelve el objeto grid
    """
    esperar_sap(session, timeout)

    posibles_grids = [
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell"
    ]

    inicio = time.time()
    while time.time() - inicio < timeout:
        for grid_id in posibles_grids:
            try:
                grid = session.findById(grid_id)
                if grid.RowCount > 0:
                    return grid
            except:
                pass
        time.sleep(0.5)

    raise Exception("CS11 no terminó de cargar el grid")


def validar_planta(session, planta):
    """
    Verifica que SAP haya aceptado la planta ingresada
    """
    try:
        campo = session.findById("wnd[0]/usr/ctxtRC29L-WERKS")
        return campo.text.strip() == planta
    except:
        return False

def tiene_parentesis_numericos(material: str) -> bool:
    """
    Detecta materiales con formato XXXXX(1), XXXXX(2), etc
    """
    return bool(re.search(r"\(\d+\)$", material))

def esperar_archivo(ruta_archivo, timeout=30):
    """
    Espera hasta que el archivo exista físicamente en disco.
    Útil para exportaciones SAP que tardan en terminar.
    """
    inicio = time.time()
    while time.time() - inicio < timeout:
        if os.path.exists(ruta_archivo):
            return True
        time.sleep(0.5)
    return False

def convertir_por_modelo(modelo_limpio):
    """
    Busca el XLS por nombre de modelo en la carpeta final
    y lo convierte a XLSX
    """
    ruta_xls = os.path.join(EXPORT_FINAL_PATH, f"{modelo_limpio}.XLS")
    ruta_xlsx = os.path.join(EXPORT_FINAL_PATH, f"{modelo_limpio}.xlsx")

    try:
        if not os.path.exists(ruta_xls):
            print(f"[ERROR] No se encontró XLS del modelo: {ruta_xls}")
            return None

        df = pd.read_csv(
            ruta_xls,
            sep="\t",
            encoding="latin1",
            engine="python"
        )

        df.to_excel(ruta_xlsx, index=False, engine="openpyxl")
        os.remove(ruta_xls)

        print(f"[INFO] XLSX generado correctamente: {ruta_xlsx} ✅")
        return ruta_xlsx

    except Exception as e:
        print(f"[ERROR] Conversión por modelo falló: {e}")
        return None


