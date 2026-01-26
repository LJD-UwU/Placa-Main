from backend.utils.sap_utils import (
    escribir_campo,
    ejecutar_busqueda,
    esperar_cs11_completo,
    pausar,
    tiene_parentesis_numericos,
    validar_planta
)
from backend.modules.cs03 import ejecutar_cs03_corregir_material

def ejecutar_cs11(session, material, componente="1TE*", uso="PP01", plantas=None, pausa_entre_acciones=0.5):
    """
    Ejecuta CS11 para un material en varias plantas y maneja automáticamente CS03
    si el material tiene paréntesis numéricos.
    Retorna una lista de tuples: [(planta, grid), ...]
    """
    if plantas is None:
        plantas = ["2000", "2900"]

    print(f"[INFO] Iniciando CS11 para: {material}")
    session.findById("wnd[0]").maximize()
    pausar(pausa_entre_acciones)

    # Ir a CS11
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCS11"
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]").sendVKey(0)
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]").sendVKey(4)  # Selección múltiple
    pausar(pausa_entre_acciones)

    # Escribir material y componente
    escribir_campo(
        session,
        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/"
        "sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]",
        f"*{material}*"
    )
    pausar(pausa_entre_acciones)

    escribir_campo(
        session,
        "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/"
        "sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]",
        componente
    )
    pausar(pausa_entre_acciones)

    # Confirmar selección múltiple
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    pausar(pausa_entre_acciones)
    session.findById("wnd[1]").sendVKey(2)
    pausar(pausa_entre_acciones)

    resultados = []
    bom_obtenido = False
    material_original = material

    for planta in plantas:
        if bom_obtenido:
            print(f"[INFO] BOM ya obtenido, saltando planta {planta}")
            break

        try:
            # Intentar setear control de planta de manera segura
            set_werks(session, planta)
            set_capid(session, uso)

            ejecutar_busqueda(session)
            pausar(pausa_entre_acciones)

            grid = esperar_cs11_completo(session, timeout=15)
            print(f"[INFO] CS11 cargado para {material} en planta {planta} ({grid.RowCount} filas)")
            resultados.append((planta, grid))
            bom_obtenido = True

        except Exception as e:
            print(f"[WARNING] CS11 falló para {material} en planta {planta}: {e}")

            if not bom_obtenido and tiene_parentesis_numericos(material_original):
                print("[INFO] Paréntesis numéricos detectados → ejecutando CS03")
                material = ejecutar_cs03_corregir_material(session, material, componente, planta)

                try:
                    set_werks(session, planta)
                    set_capid(session, uso)
                    ejecutar_busqueda(session)
                    grid = esperar_cs11_completo(session, timeout=15)
                    print(f"[INFO] CS11 corregido exitosamente en planta {planta}")
                    resultados.append((planta, grid))
                    bom_obtenido = True
                except Exception as e2:
                    print(f"[ERROR] CS03 no resolvió el problema en planta {planta}: {e2}")

    if not resultados:
        print(f"[ERROR] No se pudo obtener BOM para {material_original}")
        return None

    return resultados


# ----------------- FUNCIONES AUXILIARES SEGURAS -----------------

def set_werks(session, planta, intentos=3, pausa=0.5):
    """Intenta setear el campo de planta varias veces si falla"""
    for i in range(intentos):
        try:
            escribir_campo(session, "wnd[0]/usr/ctxtRC29L-WERKS", planta)
            return
        except Exception:
            print(f"[WARNING] Intento {i+1}/{intentos} falló para wnd[0]/usr/ctxtRC29L-WERKS")
            pausar(pausa)
    raise Exception("No se pudo setear la planta en SAP")

def set_capid(session, uso, intentos=3, pausa=0.5):
    """Intenta setear el campo de uso varias veces si falla"""
    for i in range(intentos):
        try:
            escribir_campo(session, "wnd[0]/usr/ctxtRC29L-CAPID", uso)
            return
        except Exception:
            print(f"[WARNING] Intento {i+1}/{intentos} falló para wnd[0]/usr/ctxtRC29L-CAPID")
            pausar(pausa)
    raise Exception("No se pudo setear CAPID en SAP")
