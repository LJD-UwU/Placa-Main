import os
import time
import pandas as pd
from openpyxl.styles import PatternFill
from backend.utils.txt_to_xlsx import exportar_bom_a_xls, convertir_xls_a_xlsx
from backend.utils.sap_utils import acceso_bom_exitoso
from backend.config.sap_config import (
MENSAJE_SIN_BOM,
RESULT_COLUMNS,
FILTRO,
TRANSACCION,
PAUSA,
SECUENCIA,
)

# FUNCION PARA MODELOS INTERNOS

def procesar_number(session, number, plantas, capid):
    """
    Procesa un Number en SAP para un modelo interno.
    Retorna la ruta del XLS exportado.
    """
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION
        session.findById("wnd[0]").sendVKey(0)

        session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = number
        session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = plantas
        session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = capid
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(0.8)

        if not acceso_bom_exitoso(session):
            print(f"[INFO] No se accedió al BOM para {number} en planta {plantas}")
            return None

        ruta_xls = exportar_bom_a_xls(session, number, mainboard=True)
        if not ruta_xls or not os.path.exists(ruta_xls):
            print(f"[WARNING] Falló exportación XLS para {number} en planta {plantas}")
            return None

        return ruta_xls
    except Exception as e:
        print(f"[ERROR] Error procesando {number} en planta {plantas}: {e}")
        return None


# FUNCION PARA MAINBOARD CON BLOQUE DINAMICO

def procesar_number_mainboard(session, number, capid):
    """
    Procesa un Number para obtener su Mainboard en SAP.
    Incluye bloque adicional dinámico para seleccionar fila y presionar botones.
    """


    for planta in SECUENCIA:
        try:
            print(f"[INFO] Intentando {number} en planta {planta}")

            # Abrir CS11
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION
            session.findById("wnd[0]").sendVKey(0)

            # Ingresar datos
            session.findById("wnd[0]/usr/ctxtRC29L-MATNR").text = number
            session.findById("wnd[0]/usr/ctxtRC29L-WERKS").text = planta
            session.findById("wnd[0]/usr/ctxtRC29L-CAPID").text = capid
            session.findById("wnd[0]/tbar[1]/btn[8]").press()
            time.sleep(0.8)

            # Validación BOM
            if not acceso_bom_exitoso(session):
                print(f"[INFO] No se accedió al BOM en planta {planta}")
                continue

            # Exportar BOM
            ruta_xls = exportar_bom_a_xls(session, number, mainboard=True)
            if not ruta_xls or not os.path.exists(ruta_xls):
                print(f"[WARNING] No se generó XLS para {number} en planta {planta}")
                continue

            ruta_xlsx = ruta_xls.replace(".XLS", ".xlsx")
            convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)

            # ==== BLOQUE ADICIONAL SAP GUI ====
            try:

                session.findById("wnd[0]/tbar[1]/btn[33]").press()
                time.sleep(1)  

                #  Acceder al grid
                grid = session.findById(
                    "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/"
                    "cntlG51_CONTAINER/shellcont/shell"
                )

                # Seleccionar fila 81 si existe, si no seleccionar última fila
                fila_objetivo = 81
                if grid.RowCount > fila_objetivo:
                    fila = fila_objetivo
                else:
                    fila = grid.RowCount - 1
                    print(f"[WARNING] La fila 81 no existe, seleccionando última fila {fila}")

                # Seleccionar celda
                grid.currentCellRow = fila
                grid.currentCellColumn = 0  # primera columna, ajustar si es otra columna
                grid.selectedRows = str(fila)
                grid.clickCurrentCell()
                time.sleep(PAUSA)

                #  Presionar botón siguiente
                session.findById("wnd[0]/tbar[1]/btn[45]").press()
                time.sleep(PAUSA)

                #  Seleccionar radio button y presionar OK
                radio = session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]")
                radio.select()
                radio.setFocus()
                session.findById("wnd[1]/tbar[0]/btn[0]").press()

                print("[OK] Bloque adicional ejecutado correctamente")

            except Exception as e:
                print(f"[ERROR] No se pudo ejecutar el bloque adicional: {e}")


            # Retornar XLSX
            return ruta_xlsx

        except Exception as e:
            print(f"[WARNING] Error en planta {planta}: {e}")

    raise Exception(f"No se pudo acceder al BOM de {number} en ninguna planta")


# FUNCION PARA PROCESAR EXCEL COMPLETO
def procesar_numbers_desde_excel(session, excel_input, excel_output, plantas=None, capid=FILTRO):
    if plantas is None or not plantas:
        raise ValueError("Debes pasar la lista de plantas a procesar")

    """
    Procesa todos los Numbers de un Excel:
    - Primero modelos internos
    - Luego Mainboard
    """
    if not os.path.exists(excel_input):
        print(f"[ERROR] No existe el Excel: {excel_input}")
        return

    df = pd.read_excel(excel_input)
    df = df.dropna(subset=RESULT_COLUMNS)
    if df.empty:
        print("[INFO] No hay Numbers para procesar")
        return

    df_final = pd.DataFrame(columns=["Number", "Descripcion", "Planta", "Ruta_XLSX"])

    for idx, row in df.iterrows():
        number = str(row["Number"]).strip()
        descripcion = row["Descripcion"]

        # --- Modelos internos ---
        exito = False
        for planta in plantas:
            if procesar_number(session, number, planta, capid):
                exito = True
        if not exito:
            print(f"[WARNING] No se procesó ningún modelo interno para {number}")
            continue

        # --- Mainboard ---
        try:
            ruta_xlsx = procesar_number_mainboard(session, number, capid)
            if not ruta_xlsx or not os.path.exists(ruta_xlsx):
                print(f"[WARNING] No se generó Mainboard para {number}")
                continue

            df_final = pd.concat([df_final, pd.DataFrame([{
                "Number": number,
                "Descripcion": descripcion,
                "Planta": ",".join(plantas),
                "Ruta_XLSX": ruta_xlsx
            }])], ignore_index=True)

            print(f"[OK] Mainboard procesado: {number} | XLSX: {ruta_xlsx}")

        except Exception as e:
            print(f"[ERROR] No se pudo generar Mainboard para {number}: {e}")

    # Guardar Excel final
    if not df_final.empty:
        df_final.to_excel(excel_output, index=False, engine="openpyxl")
        print(f"\n[INFO] Procesamiento completado ✅\nExcel final guardado en: {excel_output}")


# FUNCION PARA VALIDAR BOM

def acceso_bom_exitoso(session):
    """
    Determina si realmente se accedió al BOM en CS11
    """
    try:
        # Mensaje de BOM inexistente
        try:
            status = session.findById("wnd[0]/sbar").Text
            if MENSAJE_SIN_BOM in status:
                return False
        except:
            pass

        # Grid con filas
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
 