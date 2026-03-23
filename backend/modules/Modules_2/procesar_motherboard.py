import os
import time
import pandas as pd

from backend.utils.sap_utils import acceso_bom_exitoso
from backend.config.sap_config import FILTRO, TRANSACCION, PAUSA
from backend.utils.utils_2.xlsx_m2 import exportar_bom_a_xls, convertir_xls_a_xlsx


def procesar_numbers_desde_listas(session, mother_list, plant_list, excel_output, capid=FILTRO, loop_multiple=False):
    
    #! Detectar si es lista de listas
    if loop_multiple:
        for i, (m_list, p_list) in enumerate(zip(mother_list, plant_list)):
            print(f"\n[INFO] Iniciando batch {i+1}/{len(mother_list)}")

            procesar_numbers_desde_listas(
                session,
                m_list,
                p_list,
                f"{excel_output.rstrip('.xlsx')}_{i+1}.xlsx",
                capid
            )
        return

    #! Validación básica
    if len(mother_list) != len(plant_list):
        raise ValueError("Las listas mother_list y plant_list deben tener la misma longitud")

    df_final = pd.DataFrame(columns=["Motherboard", "Plant", "Ruta_XLSX"])

    for idx, mother in enumerate(mother_list):

        plant = plant_list[idx]

        print(f"[INFO] Procesando {mother} en planta {plant}")

        try:

            #! Abrir transacción SAP
            session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION
            session.findById("wnd[0]").sendVKey(0)

            #! Ingresar datos
            session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = mother
            session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = plant
            session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = capid

            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            time.sleep(PAUSA)

            #! Validación BOM
            if not acceso_bom_exitoso(session):

                print(f"[INFO] No se accedió al BOM para {mother} en planta {plant}")

                continue

            #! Exportar BOM (ya guarda en MAINBOARD_2_FILES_FOLDER)
            ruta_xls = exportar_bom_a_xls(session, mother)

            if not ruta_xls or not os.path.exists(ruta_xls):

                print(f"[WARNING] No se generó XLS para {mother} en planta {plant}")

                continue

            #! Convertir a XLSX
            base, _ = os.path.splitext(ruta_xls)
            ruta_xlsx = base + ".xlsx"

            convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
            print(f"[OK] Mainboard procesado: {mother} | XLSX: {ruta_xlsx}")
        except Exception as e:

            print(f"[ERROR] Error procesando {mother} en planta {plant}: {e}")

    #! Guardar Excel final
    if not df_final.empty:

        df_final.to_excel(excel_output, index=False, engine="openpyxl")

        print(f"\n[INFO] Procesamiento completado ✅")
        print(f"Archivo final guardado en: {excel_output}")
