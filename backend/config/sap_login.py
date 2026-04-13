import sys
import time
import subprocess
import win32com.client

from backend.config.sap_config import SAP_LOGON_PATH
from backend.utils.sap_utils import esperar_sap, escribir_campo
from backend.config.credenciales_loader import cargar_credenciales


def obtener_sesion_existente(application, target_system_name):
    """
    Busca sesión activa en SAP usando coincidencia flexible
    (ej: "HQ" dentro de "HQ PRD")
    """
    try:
        for i in range(application.Connections.Count):
            connection = application.Connections(i)

            for j in range(connection.Children.Count):
                session = connection.Children(j)

                if session.Busy:
                    continue

                info = session.Info

                system_name = str(info.SystemName).upper()
                connection_name = str(connection.Description).upper()

                # 🔥 MATCH FLEXIBLE
                if (
                    target_system_name.upper() in system_name
                    or target_system_name.upper() in connection_name
                ):
                    print(f"[INFO] Sesión encontrada: {connection.Description} ✅")
                    session.findById("wnd[0]").maximize()
                    return session

        return None

    except Exception as e:
        print(f"[WARN] Error buscando sesión existente: {e}")
        return None


def abrir_sap_y_login(timeout=60, max_intentos=3):

    for intento in range(1, max_intentos + 1):
        try:
            print(f"[INFO] Intento {intento} de {max_intentos} para iniciar SAP...")

            # 🔐 Cargar credenciales
            cred = cargar_credenciales()
            SAP_SYSTEM_NAME = cred.get("SAP_SYSTEM_NAME")  # Ej: "HQ"
            SAP_USER = cred.get("SAP_USER")
            SAP_PASSWORD = cred.get("SAP_PASSWORD")

            if not all([SAP_SYSTEM_NAME, SAP_USER, SAP_PASSWORD]):
                raise ValueError("Credenciales SAP incompletas en Credenciales.json")

            sap_gui_auto = None
            application = None

            # INTENTAR REUTILIZAR SESIÓN EXISTENTE (CORRECTA)
            try:
                sap_gui_auto = win32com.client.GetObject("SAPGUI")

                if sap_gui_auto:
                    application = sap_gui_auto.GetScriptingEngine

                    session = obtener_sesion_existente(application, SAP_SYSTEM_NAME)

                    if session:
                        print("[INFO] Reutilizando sesión existente... ✅")
                        return session

            except Exception:
                print("[INFO] No se detectó sesión activa válida.")

            # 🚀 2. ABRIR SAP
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

            # 🔌 3. CONECTAR AL SISTEMA (HQ)
            print(f"[INFO] Conectando a: {SAP_SYSTEM_NAME}...")
            connection = application.OpenConnection(SAP_SYSTEM_NAME, True)
            session = connection.Children(0)
            session.findById("wnd[0]").maximize()

           # 🔑 4. LOGIN (SI ES NECESARIO) + INGRSO SIN CERRAR LA SESSION DE LOS DEMAS USUARIOS
            try:
                escribir_campo(session, "wnd[0]/usr/txtRSYST-BNAME", SAP_USER)
                escribir_campo(session, "wnd[0]/usr/pwdRSYST-BCODE", SAP_PASSWORD)
                session.findById("wnd[0]").sendVKey(0)
                esperar_sap(session)


                try:
                    popup = session.findById("wnd[1]")

                    popup.findById("usr/radMULTI_LOGON_OPT2").select()
                    popup.findById("tbar[0]/btn[0]").press()

                    print("[INFO] Multi-logon detectado, opción seleccionada ✅")

                except Exception:
                    pass

                print("[INFO] Login SAP exitoso ✅")

            except Exception:
                print("[INFO] El sistema ya tenía sesión iniciada o no mostró login.")

            return session

        except Exception as e:
                print(f"[ERROR] Intento {intento} falló: {e}")

                if intento < max_intentos:
                    print("[INFO] Reintentando en 5 segundos...")
                    time.sleep(5)
                else:
                    print("[ERROR] Se agotaron los intentos para iniciar SAP.")
                    sys.exit(1)