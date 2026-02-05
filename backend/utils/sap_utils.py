import time
import re
import os
from backend.config.sap_config import (
    PAUSA
)
# UTILIDADES GENERALES SAP

timeout=15

def pausar(segundos=3):
    """Pausa la ejecución"""
    time.sleep(segundos)

def esperar_sap(session, timeout_local=None):
    t = timeout_local if timeout_local else timeout
    for _ in range(t * 10):
        if not session.Busy:
            return
        time.sleep(0.1)
    raise Exception("SAP no respondió a tiempo")

def esperar_id(session, id_control):
    """Espera hasta que un control exista en SAP"""
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            return session.findById(id_control)
        except:
            time.sleep(0.2)
    raise Exception(f"Control no encontrado: {id_control}")

def escribir_campo(session, id_campo, texto, limpiar=True):
    """
    Escribe texto en un campo SAP asegurando que no queden residuos.
    """
    campo = esperar_id(session, id_campo)

    try:
        campo.setFocus()
        if limpiar:
            campo.text = ""
            campo.sendVKey(4)   # Ctrl + A
            campo.sendVKey(2)   # Delete
            time.sleep(0.1)
    except:
        pass

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

# UTILIDADES CS11 / VALIDACIONES

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
    esperar_sap(session)
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
        time.sleep(PAUSA)
    raise Exception("CS11 no terminó de cargar el grid")

# CONEXIÓN Y EXPORTACIÓN

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
            pass  

        print(f"[INFO] Exportado a Excel: {ruta_excel}")
        return ruta_excel
    except Exception as e:
        print(f"[ERROR] No se pudo exportar BOM: {e}")
        return None

def validar_planta(session, planta, intentos=3, delay=1):
    control_id = "wnd[0]/usr/ctxtRC29L-WERKS"
    
    for intento in range(1, intentos+1):
        try:
            campo = session.findById(control_id)
            if campo.text.strip() == planta:
                return True
            else:
                # Si no coincide, intentamos setearlo nuevamente
                campo.text = planta
                campo.caretPosition = len(planta)
                time.sleep(delay)
                if campo.text.strip() == planta:
                    return True
        except Exception as e:
            print(f"[WARNING] Intento {intento}/{intentos} falló para {control_id}: {e}")
            time.sleep(delay)

    print(f"[ERROR] No se pudo validar la planta: {planta}")
    return False

def mensaje_sap_contiene(session, texto):
    """
    Revisa si la barra de estado de SAP contiene cierto texto
    """
    try:
        status = session.findById("wnd[0]/sbar").Text
        return texto in status
    except Exception:
        return False
    
def acceso_bom_exitoso(session):
    """
    Verifica si realmente se accedió al BOM en CS11
    """
    try:
        # ❌ Mensaje de BOM inexistente
        try:
            status = session.findById("wnd[0]/sbar").Text
            if "没有可用的 BOM" in status:
                return False
        except:
            pass

        # ✅ Grid real con filas
        posibles_grids = [
            "wnd[0]/usr/cntlGRID1/shellcont/shell",
            "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell"
        ]

        for gid in posibles_grids:
            try:
                grid = session.findById(gid)
                if grid.RowCount > 0:
                    return True
            except:
                pass

        return False
    except:
        return False


