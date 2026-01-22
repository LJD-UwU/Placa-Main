import sys
import time
import subprocess
import win32com.client
from backend.config.sap_config import SAP_LOGON_PATH, SAP_USER, SAP_PASSWORD, SAP_SYSTEM_NAME
from backend.utils.sap_utils import esperar_sap, escribir_campo

def abrir_sap_y_login(timeout=60):
    try:
        sap_gui_auto = None
        application = None

        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            if sap_gui_auto:
                application = sap_gui_auto.GetScriptingEngine
                if application.Connections.Count > 0:
                    connection = application.Connections(0)
                    if connection.Children.Count > 0:
                        session = connection.Children(0)
                        print("[INFO] Sesión de SAP activa detectada. Reutilizando... ✅")
                        session.findById("wnd[0]").maximize()
                        return session
        except Exception:
            print("[INFO] No se detectó sesión activa. Procediendo a abrir SAP...")

        print(f"[INFO] Iniciando SAP Logon en {SAP_LOGON_PATH}...")
        subprocess.Popen(SAP_LOGON_PATH)

        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                sap_gui_auto = win32com.client.GetObject("SAPGUI")
                if sap_gui_auto: break
            except:
                time.sleep(1)

        if not sap_gui_auto:
            raise TimeoutError("No se pudo conectar con el objeto SAPGUI")

        application = sap_gui_auto.GetScriptingEngine
        print(f"[INFO] Conectando a: {SAP_SYSTEM_NAME}...")
        connection = application.OpenConnection(SAP_SYSTEM_NAME, True)
        session = connection.Children(0)
        session.findById("wnd[0]").maximize()

        try:
            escribir_campo(session, "wnd[0]/usr/txtRSYST-BNAME", SAP_USER)
            escribir_campo(session, "wnd[0]/usr/pwdRSYST-BCODE", SAP_PASSWORD)
            session.findById("wnd[0]").sendVKey(0)
            esperar_sap(session)
            print("[INFO] Login SAP exitoso ✅")
        except Exception:
            print("[INFO] El sistema ya inició sesión o la pantalla de login no apareció.")

        return session

    except Exception as e:
        print(f"[ERROR] Falló la apertura o login de SAP: {e}")
        sys.exit(1)
