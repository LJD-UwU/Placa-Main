import win32com.client
import time
import re
from backend.config.sap_config import SAP_LOGON_PATH, SAP_SYSTEM_NAME, SAP_USER, SAP_PASSWORD

# --- Funciones de SAP ---
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
            control = session.findById(id_control)
            return control
        except:
            time.sleep(0.2)
    raise Exception(f"Control no encontrado: {id_control}")

def escribir_campo(session, id_campo, texto, mover_caret=True):
    campo = esperar_id(session, id_campo)
    campo.text = texto
    if mover_caret:
        campo.caretPosition = len(texto)
    return campo

def ejecutar_busqueda(session):
    esperar_sap(session)
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
    except:
        session.findById("wnd[0]").sendVKey(8)
    esperar_sap(session)

def esperar_cs11_completo(session, timeout=30):
    esperar_sap(session, timeout)
    posibles_grids = [
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell"
    ]
    inicio = time.time()
    while time.time() - inicio < timeout:
        for g in posibles_grids:
            try:
                grid = session.findById(g)
                if grid.RowCount > 0:
                    return grid
            except:
                pass
        time.sleep(0.5)
    raise Exception("CS11 no terminó de cargar el grid")

def validar_planta(session, planta):
    try:
        werks_field = session.findById("wnd[0]/usr/ctxtRC29L-WERKS")
        return werks_field.text.strip() == planta
    except:
        return False

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
        con = None
        for i in range(app.Children.Count):
            conexion = app.Children.Item(i)
            if SAP_SYSTEM_NAME.lower() in conexion.Description.lower():
                con = conexion
                break
        if con is None:
            con = app.OpenConnection(SAP_SYSTEM_NAME, True)
            time.sleep(2)
        if con.Children.Count == 0:
            ses = con.Children.Add(0)
        else:
            ses = con.Children.Item(0)

        # Login automático
        try:
            if ses.findById("wnd[0]/usr/txtRSYST-BNAME").text == "":
                ses.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USER
                ses.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASSWORD
                ses.findById("wnd[0]/tbar[0]/btn[0]").press()
                time.sleep(2)
        except:
            pass

        ses.findById("wnd[0]").maximize()
        print("[INFO] Conexión SAP establecida ✅")
        return ses
    except Exception as e:
        print(f"[ERROR] No se pudo conectar a SAP: {e}")
        return None

# --- Funciones de material ---
def quitar_parentesis(material):
    return re.sub(r"\([^)]*\)$", "", material)

def tiene_parentesis_numericos(material: str) -> bool:
    return bool(re.search(r"\(\d+\)$", material))
