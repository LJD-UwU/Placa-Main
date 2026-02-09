import json 
import os

RUTA_CREDENCIALES = os.path.join(
    os.path.dirname(__file__),
    "Credenciales.json"
)

def cargar_credenciales():
    with open(RUTA_CREDENCIALES, "r", encoding="utf-8") as f:
        return json.load(f)

def guardar_credenciales(data):
    with open(RUTA_CREDENCIALES, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
