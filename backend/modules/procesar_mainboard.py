import os
import time
import pandas as pd
from backend.utils.txt_to_xlsx import exportar_bom_a_xls, convertir_xls_a_xlsx, MAINBOARD_FILES_FOLDER
import time
from backend.utils.txt_to_xlsx import exportar_bom_a_xls
from backend.utils.sap_utils import acceso_bom_exitoso
# ==============================
# FUNCION PRINCIPAL PARA UN NUMBER
# ==============================
def procesar_number(session, number, planta, capid):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCS11"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = number
    session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = planta
    session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = capid
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    time.sleep(0.8)

    # 🔎 VALIDACIÓN REAL
    if not acceso_bom_exitoso(session):
        raise Exception("No se pudo acceder al BOM")

    ruta_xls = exportar_bom_a_xls(session, number, mainboard=True)
    if not ruta_xls:
        raise Exception("Falló exportación XLS")

    return ruta_xls
MENSAJE_SIN_BOM = "没有可用的 BOM"

def procesar_number_mainboard(session, number, capid):
    plantas = ["2000", "2900"]
    secuencia = ["2000", "2900", "2000"]  # 🔁 como pediste

    for planta in secuencia:
        try:
            print(f"[INFO] Intentando {number} en planta {planta}")

            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "/NCS11"
            session.findById("wnd[0]").sendVKey(0)

            session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = number
            session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = planta
            session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = capid
            session.findById("wnd[0]/tbar[1]/btn[8]").press()

            time.sleep(0.8)

            # 🔎 VALIDACIÓN REAL
            if not acceso_bom_exitoso(session):
                print(f"[INFO] No se accedió al BOM en planta {planta}")
                continue  # 🔁 cambiar planta

            # ✅ BOM REALMENTE CARGADO
            ruta_xls = exportar_bom_a_xls(session, number, mainboard=True)
            ruta_xlsx = ruta_xls.replace(".XLS", ".xlsx")
            convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)

            print(f"[OK] BOM obtenido para {number} en planta {planta}")
            return ruta_xlsx

        except Exception as e:
            print(f"[WARNING] Error en planta {planta}: {e}")

    raise Exception(f"No se pudo acceder al BOM de {number} en ninguna planta")

# ==============================
# FUNCION PARA PROCESAR EXCEL COMPLETO
# ==============================
def procesar_numbers_desde_excel(session, excel_input, excel_output, plantas=["2000","2900"], capid="PP01"):
    """
    Procesa todos los Numbers de un Excel:
    1. Procesa modelos internos primero
    2. Al final de cada Number, genera Mainboard XLS/XLSX
    """
    if not os.path.exists(excel_input):
        print(f"[ERROR] No existe el Excel: {excel_input}")
        return

    df = pd.read_excel(excel_input)
    df = df.dropna(subset=["Number", "Descripcion"])
    if df.empty:
        print("[INFO] No hay Numbers para procesar")
        return

    df_final = pd.DataFrame(columns=["Number", "Descripcion", "Planta", "Ruta_XLSX"])

    for idx, row in df.iterrows():
        number = str(row["Number"]).strip()
        descripcion = row["Descripcion"]

        # --- Primero procesar todos los modelos internos ---
        exito = False
        for planta in plantas:
            if procesar_number(session, number, planta, capid):
                exito = True

        if not exito:
            print(f"[WARNING] No se procesó ningún modelo interno para {number}")
            continue

        # --- Al final, exportar Mainboard una sola vez ---
        try:
            ruta_xls = exportar_bom_a_xls(session, number, mainboard=True)
            if not ruta_xls or not os.path.exists(ruta_xls):
                print(f"[WARNING] No se generó XLS de Mainboard para {number}")
                continue

            ruta_xlsx = os.path.join(MAINBOARD_FILES_FOLDER, os.path.basename(ruta_xls).replace(".XLS", ".xlsx"))
            convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)

            df_final = pd.concat([df_final, pd.DataFrame([{
                "Number": number,
                "Descripcion": descripcion,
                "Planta": ",".join(plantas),
                "Ruta_XLSX": ruta_xlsx
            }])], ignore_index=True)

            print(f"[OK] Mainboard procesado: {number} | XLSX: {ruta_xlsx}")

        except Exception as e:
            print(f"[ERROR] No se pudo generar Mainboard para {number}: {e}")

    # --- Guardar Excel final ---
    if not df_final.empty:
        df_final.to_excel(excel_output, index=False, engine="openpyxl")
        print(f"\n[INFO] Procesamiento completado ✅\nExcel final guardado en: {excel_output}")
        
def acceso_bom_exitoso(session):
    """
    Determina si realmente se accedió al BOM en CS11
    """
    try:
        # 1️⃣ No debe existir mensaje de BOM inexistente
        try:
            status = session.findById("wnd[0]/sbar").Text
            if "没有可用的 BOM" in status:
                return False
        except:
            pass

        # 2️⃣ Debe existir grid con filas
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
    except Exception:
        return False

