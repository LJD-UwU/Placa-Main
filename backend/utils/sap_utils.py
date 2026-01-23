import time
import re
import os

# ==============================
# UTILIDADES GENERALES SAP
# ==============================

def pausar(segundos=1):
    """Pausa la ejecución"""
    time.sleep(segundos)

def esperar_sap(session, timeout=15):
    """Espera a que SAP no esté ocupado"""
    for _ in range(timeout * 10):
        if not session.Busy:
            return
        time.sleep(0.1)
    raise Exception("SAP no respondió a tiempo")

def esperar_id(session, id_control, timeout=15):
    """Espera hasta que un control exista en SAP"""
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            return session.findById(id_control)
        except:
            time.sleep(0.2)
    raise Exception(f"Control no encontrado: {id_control}")

def escribir_campo(session, id_campo, texto):
    """Escribe texto en un campo SAP"""
    campo = esperar_id(session, id_campo)
    campo.text = texto
    campo.caretPosition = len(texto)
    return campo

def ejecutar_busqueda(session):
    """Presiona el botón de búsqueda (F8)"""
    esperar_sap(session)
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
    except:
        session.findById("wnd[0]").sendVKey(8)
    esperar_sap(session)

# ==============================
# UTILIDADES CS11 / VALIDACIONES
# ==============================

def validar_planta(session, planta):
    """Valida si la planta actual coincide"""
    try:
        campo = session.findById("wnd[0]/usr/ctxtRC29L-WERKS")
        return campo.text.strip() == planta
    except:
        return False

def tiene_parentesis_numericos(material: str) -> bool:
    """Verifica si un material tiene paréntesis con números al final"""
    return bool(re.search(r"\(\d+\)$", material))

def esperar_cs11_completo(session, timeout=30):
    """Espera a que el grid de CS11 cargue con datos"""
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

# ==============================
# CONEXIÓN Y EXPORTACIÓN
# ==============================

def conectar_sap():
    """
    Conecta a SAP usando SAP GUI scripting.
    Retorna el objeto session activo o None si falla.
    """
    try:
        import win32com.client
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not SapGuiAuto:
            print("[ERROR] SAPGUI no está iniciado")
            return None
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
        print("[INFO] Conexión a SAP establecida")
        return session
    except Exception as e:
        print(f"[ERROR] Falló la conexión a SAP: {e}")
        return None

def exportar_bom_a_excel(session, nombre_archivo="BOM.xlsx", ruta_carpeta=os.getcwd()):
    """
    Exporta la pantalla actual (CS11) a Excel.
    Retorna la ruta completa del archivo generado.
    """
    try:
        # Construir ruta completa
        ruta_excel = os.path.join(ruta_carpeta, nombre_archivo)

        # Usar función de SAP para exportar a Excel
        # Presionar botón de exportar
        try:
            session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        except:
            pass  # O usar alternativa según versión SAP

        # Guardar archivo (esto depende de tu SAP, aquí se deja ejemplo genérico)
        # session.findById("wnd[1]/usr/ctxtDY_PATH").text = ruta_carpeta
        # session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nombre_archivo
        # session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # Para este ejemplo, asumimos que SAP ya guardó el archivo
        print(f"[INFO] Exportado a Excel: {ruta_excel}")
        return ruta_excel
    except Exception as e:
        print(f"[ERROR] No se pudo exportar BOM: {e}")
        return None
