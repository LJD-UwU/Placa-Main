from backend.utils.sap_utils import esperar_id, escribir_campo, ejecutar_busqueda, esperar_cs11_completo, validar_planta
import pandas as pd

def ejecutar_cs11(session, material, componente="**", uso="PP01", plantas=["2000", "2900"]):
    try:
        session.findById("wnd[0]").maximize()
        okcd = esperar_id(session, "wnd[0]/tbar[0]/okcd")
        okcd.text = "/NCS11"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(4)

        sel_componente = "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]"
        sel_material = "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]"

        escribir_campo(session, sel_componente, componente)
        escribir_campo(session, sel_material, f"{material}*")

        ctrl_material = esperar_id(session, sel_material)
        ctrl_material.setFocus()
        ctrl_material.caretPosition = len(material)

        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]").sendVKey(2)

        for planta in plantas:
            try:
                escribir_campo(session, "wnd[0]/usr/ctxtRC29L-WERKS", planta)
                escribir_campo(session, "wnd[0]/usr/ctxtRC29L-CAPID", uso)

                if not validar_planta(session, planta):
                    print(f"[WARNING] Planta '{planta}' no es válida. Se omite.")
                    continue

                ejecutar_busqueda(session)
                grid = esperar_cs11_completo(session, timeout=30)
                datos = []
                for r in range(grid.RowCount):
                    fila = [grid.GetCellValue(r, c) for c in range(grid.ColumnCount)]
                    datos.append(fila)

                df = pd.DataFrame(datos, columns=[grid.GetColumnName(c) for c in range(grid.ColumnCount)])
                print(f"[INFO] CS11 completado para {material} en planta {planta}, filas: {len(df)} ✅")

            except Exception as e_planta:
                print(f"[ERROR] No se pudo procesar {material} en planta {planta}: {e_planta}")

        return True

    except Exception as e:
        raise Exception(f"No se pudo procesar el modelo '{material}': {e}")
