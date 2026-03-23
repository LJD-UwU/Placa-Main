import sys
import time
import subprocess
import win32com.client

from backend.config.sap_config import SAP_LOGON_PATH
from backend.utils.sap_utils import esperar_sap, escribir_campo
from backend.config.credenciales_loader import cargar_credenciales

def abrir_sap_y_login(timeout=60, max_intentos=3):

    for intento in range(1, max_intentos + 1):
        try:
            print(f"[INFO] Intento {intento} de {max_intentos} para iniciar SAP...")

            #! Cargar credenciales
            cred = cargar_credenciales()
            SAP_SYSTEM_NAME = cred.get("SAP_SYSTEM_NAME")
            SAP_USER = cred.get("SAP_USER")
            SAP_PASSWORD = cred.get("SAP_PASSWORD")

            if not all([SAP_SYSTEM_NAME, SAP_USER, SAP_PASSWORD]):
                raise ValueError("Credenciales SAP incompletas en Credenciales.json")

            sap_gui_auto = None
            application = None

            #! Reutilizar sesión existente
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
                print("[INFO] No se detectó sesión activa. Abriendo SAP...")

            #! Abrir SAP
            print(f"[INFO] Iniciando SAP Logon en {SAP_LOGON_PATH}...")
            subprocess.Popen(SAP_LOGON_PATH)

            start_time = time.time()
            while time.time() - start_time < timeout:
                try:
                    sap_gui_auto = win32com.client.GetObject("SAPGUI")
                    if sap_gui_auto:
                        break
                except Exception:
                    time.sleep(1)

            if not sap_gui_auto:
                raise TimeoutError("No se pudo conectar con el objeto SAPGUI")

            application = sap_gui_auto.GetScriptingEngine

            #! Conectar al sistema
            print(f"[INFO] Conectando a: {SAP_SYSTEM_NAME}...")
            connection = application.OpenConnection(SAP_SYSTEM_NAME, True)
            session = connection.Children(0)
            session.findById("wnd[0]").maximize()

            #! Login
            try:
                escribir_campo(session, "wnd[0]/usr/txtRSYST-BNAME", SAP_USER)
                escribir_campo(session, "wnd[0]/usr/pwdRSYST-BCODE", SAP_PASSWORD)
                session.findById("wnd[0]").sendVKey(0)
                esperar_sap(session)
                print("[INFO] Login SAP exitoso ✅")
            except Exception:
                print("[INFO] El sistema ya inició sesión o no apareció el login.")

            return session

        except Exception as e:
            print(f"[ERROR] Intento {intento} falló: {e}")

            if intento < max_intentos:
                print("[INFO] Reintentando en 5 segundos...")
                time.sleep(5)
            else:
                print("[ERROR] Se agotaron los intentos para iniciar SAP.")
                sys.exit(1)