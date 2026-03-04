import json
import os

RUTA_CREDENCIALES = os.path.join(
    os.path.dirname(__file__),
    "Credenciales.json"
)

def cargar_credenciales():
    #! Si no existe el json de credenciales de crea desde cero 
    if not os.path.exists(RUTA_CREDENCIALES):
        credenciales_vacias = {
            "SAP_SYSTEM_NAME": "",
            "SAP_USER": "",
            "SAP_PASSWORD": ""
        }
        with open(RUTA_CREDENCIALES, "w", encoding="utf-8") as f:
            json.dump(credenciales_vacias, f, indent=2)
        return credenciales_vacias

    #! Si existe no se crea y se usa el qeu ya esta
    with open(RUTA_CREDENCIALES, "r", encoding="utf-8") as f:
        return json.load(f)

def guardar_credenciales(data):
    with open(RUTA_CREDENCIALES, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
