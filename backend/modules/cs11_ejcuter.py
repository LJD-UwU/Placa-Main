from backend.utils.sap_utils import (
    escribir_campo,
    ejecutar_busqueda,
    esperar_cs11_completo,
    pausar,
    tiene_parentesis_numericos,
    validar_planta
)
from backend.modules.cs03 import ejecutar_cs03_corregir_material
from backend.config.sap_config import (
    TRANSACCION,
    FILTRO_SAP,
    FILTRO,
    PAUSA,
    PLANTAS
)

def ejecutar_cs11(session, material, componente=FILTRO_SAP, uso=FILTRO, plantas=None, pausa_entre_acciones=PAUSA):
  
    if PLANTAS is None:
        PLANTAS

    print(f"[INFO] Iniciando CS11 para: {material}")

    # --- Ir a CS11 ---
    session.findById("wnd[0]").maximize()
    pausar(pausa_entre_acciones)

    session.findById("wnd[0]/tbar[0]/okcd").text = TRANSACCION
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]").sendVKey(0)
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]").sendVKey(4) 
    #pausar(pausa_entre_acciones)

    # Colocar el material dentro de **
    escribir_campo(session,
        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/"
        "ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/"
        "txtG_SELFLD_TAB-LOW[0,24]",
        f"*.{material}.*")
    pausar(pausa_entre_acciones)

    escribir_campo(session,
        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/"
        "ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/"
        "txtG_SELFLD_TAB-LOW[2,24]",
        componente)
    pausar(pausa_entre_acciones)

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    pausar(pausa_entre_acciones)
    session.findById("wnd[1]").sendVKey(2)
    pausar(pausa_entre_acciones)

    material_original = material

    for planta in plantas:
        # Validar planta
        if not validar_planta(session, planta):
            print(f"[WARNING] Planta {planta} no válida, se omite")
            continue

        escribir_campo(session, "wnd[0]/usr/ctxtRC29L-CAPID", uso)
        pausar(pausa_entre_acciones)
        ejecutar_busqueda(session)
        pausar(pausa_entre_acciones)

        try:
            grid = esperar_cs11_completo(session, timeout=15)
            print(f"[INFO] CS11 cargado para {material} en planta {planta}")
            # --- Detener bucle tan pronto como encontramos un grid válido ---
            return [(planta, grid)]
        except Exception:
            print(f"[WARNING] CS11 falló para {material} en planta {planta}")

            # Si tiene paréntesis numéricos, intentar corregir con CS03
            if tiene_parentesis_numericos(material_original):
                print("[INFO] Paréntesis numéricos detectados → ejecutando CS03")
                material = ejecutar_cs03_corregir_material(session, material, componente, planta)
                ejecutar_busqueda(session)
                pausar(pausa_entre_acciones)
                try:
                    grid = esperar_cs11_completo(session, timeout=15)
                    print(f"[INFO] CS11 corregido exitosamente en planta {planta}")
                    return [(planta, grid)]
                except Exception:
                    print("[ERROR] CS03 no resolvió el problema")
            else:
                print("[INFO] No aplica CS03 para este modelo")

    print(f"[ERROR] No se pudo obtener BOM para {material}")
    return None
