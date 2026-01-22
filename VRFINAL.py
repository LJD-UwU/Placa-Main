import pandas as pd
import time
from backend.utils.sap_utils import conectar_sap, tiene_parentesis_numericos
from backend.scripts.ejecutar_cs11 import ejecutar_cs11
from backend.utils.sap_utils import exportar_bom_a_excel

# Configuración
RUTA_EXCEL = r"C:\Users\admin\Documents\Practicante Archivos Main\Automatizacion\modelos.xlsx"
PLANTAS = ["2000", "2900"]  # Lista de plantas en orden de preferencia
COMPONENTE_DEFAULT = "1TE*"
USO_DEFAULT = "PP01"
PAUSA = 0.5  # Pausa entre acciones de SAP

def procesar_modelos():
    # --- Leer Excel ---
    try:
        df = pd.read_excel(RUTA_EXCEL)
        modelos = df.iloc[:, 0].dropna().astype(str).tolist()
    except Exception as e:
        print(f"[ERROR] No se pudo leer el Excel: {e}")
        return

    # --- Conectar a SAP ---
    session = conectar_sap()
    if session is None:
        print("[ERROR] No se pudo conectar a SAP, abortando")
        return

    # --- Procesar cada modelo ---
    for i, modelo in enumerate(modelos, start=1):
        print(f"\n========== {i}/{len(modelos)}: {modelo} ==========")
        try:
            # Intentar CS11
            grid = ejecutar_cs11(
                session,
                material=modelo,
                componente=COMPONENTE_DEFAULT,
                uso=USO_DEFAULT,
                plantas=PLANTAS,
                pausa_entre_acciones=PAUSA
            )

            if grid:
                print(f"[INFO] Modelo {modelo} procesado exitosamente.")
                
                ruta_excel_exportado = exportar_bom_a_excel(session, nombre_archivo=f"{modelo}_BOM.xlsx")
                if ruta_excel_exportado:
                        print(f"[INFO] BOM del modelo {modelo} guardado en {ruta_excel_exportado}")
            else:
                # Si el material tiene paréntesis y CS11 no funcionó, se reporta
                if tiene_parentesis_numericos(modelo):
                    print(f"[WARNING] CS11 no encontró BOM para {modelo} → CS03 se ejecutó previamente si era necesario")
                else:
                    print(f"[WARNING] Modelo {modelo} no tiene BOM disponible en ninguna planta.")

            print("[INFO] Esperando 5 segundos antes del siguiente modelo...")
            time.sleep(5)

        except Exception as e:
            print(f"[ERROR] Falló al procesar modelo {modelo}: {e}")
            time.sleep(3)

    print("\n[FIN] Todos los modelos procesados ✅")

if __name__ == "__main__":
    procesar_modelos()
