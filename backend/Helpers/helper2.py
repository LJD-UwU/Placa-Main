import json
import os
from backend.utils.txt_to_xlsx import MOTHERBOARD_2_FILES_FOLDER

PROCESADOS_FILE = os.path.join(MOTHERBOARD_2_FILES_FOLDER, "archivos_procesados.json")

def cargar_archivos_procesados():
    if os.path.exists(PROCESADOS_FILE):
        with open(PROCESADOS_FILE, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()

def guardar_archivo_procesado(nombre_archivo):
    procesados = cargar_archivos_procesados()
    procesados.add(nombre_archivo)
    with open(PROCESADOS_FILE, "w", encoding="utf-8") as f:
        json.dump(list(procesados), f, ensure_ascii=False, indent=2)