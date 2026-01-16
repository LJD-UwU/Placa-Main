import subprocess
import os
import time
from Backend.Settings.Rutas import RUTAS_SAP
from Backend.Utils.Executor import ejecutar_json

def iniciar_proceso_bom(ruta_json):
    print(f"--- Iniciando SAP ---")
    sap_abierto = False
    for ruta in RUTAS_SAP:
        if os.path.exists(ruta):
            subprocess.Popen(ruta)
            sap_abierto = True
            break
    
    if not sap_abierto:
        print("❌ No se encontró saplogon.exe")
        return False
        
    time.sleep(5) # Espera a que cargue la interfaz de SAP
    try:
        ejecutar_json(ruta_json)
        return True
    except Exception as e:
        print(f"Error en ejecución: {e}")
        return False
