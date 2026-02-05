import os
from datetime import datetime
import pandas as pd
from backend.config.sap_config import EXPORT_FINAL_PATH

# Carpeta y archivo de historial dentro de EXPORT_FINAL_PATH
HISTORIAL_FOLDER = os.path.join(EXPORT_FINAL_PATH, "HISTORIAL")
os.makedirs(HISTORIAL_FOLDER, exist_ok=True)
HISTORIAL_FILE = "historial.xlsx"

def _ruta_historial():
    return os.path.join(HISTORIAL_FOLDER, HISTORIAL_FILE)

def registrar_historial_excel(
    archivo: str,
    proceso: str,
    paso: str,
    estado: str,
    detalle: str = "",
    tipo: str = "Modelo"   # NUEVO: tipo de registro
):
    """
    Registra una fila detallada en el historial de procesamiento en Excel.
    Se agrega el tipo: Modelo, Motherboard o Mainboard.
    """
    ahora = datetime.now()

    nueva_fila = {
        "Fecha": ahora.strftime("%Y-%m-%d"),
        "Hora": ahora.strftime("%H:%M:%S"),
        "Tipo": tipo,          # NUEVO
        "Archivo": archivo,
        "Proceso": proceso,
        "Paso": paso,
        "Estado": estado,
        "Detalle": detalle
    }

    ruta = _ruta_historial()

    if os.path.exists(ruta):
        df = pd.read_excel(ruta, engine="openpyxl")
        df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)
    else:
        df = pd.DataFrame([nueva_fila])

    df.to_excel(ruta, index=False, engine="openpyxl")
