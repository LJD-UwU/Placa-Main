from backend.utils.sap_utils import (
    escribir_campo,
    ejecutar_busqueda,
    esperar_cs11_completo,
    pausar
)
from backend.config.sap_config import (
    TRANSACCION,
    FILTRO_SAP,
    FILTRO,
    PAUSA
)

def ejecutar_cs11(session, material, planta=None, alternativa=None, componente=FILTRO_SAP, uso=FILTRO, plantas=None, pausa_entre_acciones=PAUSA):
    # Si se pasó una planta individual, convertirla en lista
    if planta is not None and (plantas is None or not plantas):
        plantas = [planta]

    if plantas is None or not plantas:
        raise ValueError("Debes pasar la lista de plantas a ejecutar_cs11")

    print(f"[INFO] Iniciando CS11 para: {material}")
    session.findById("wnd[0]").maximize()
    pausar(pausa_entre_acciones)

    # Ir a CS11
    print(f"[INFO] Ingresando transacción {TRANSACCION}")
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCS11"
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]").sendVKey(0)
    print("[INFO] Ejecutada transacción")
    pausar(pausa_entre_acciones)

    resultados = []
    bom_obtenido = False

    for idx, planta_actual in enumerate(plantas):
        print(f"[INFO] Procesando planta: {planta_actual}")

        try:
            # Material
            campo_mat = session.findById("wnd[0]/usr/ctxtRC29L-MATNR")
            campo_mat.setFocus()
            campo_mat.text = material
            campo_mat.caretPosition = len(material)

            # Planta
            campo_plant = session.findById("wnd[0]/usr/ctxtRC29L-WERKS")
            campo_plant.setFocus()
            campo_plant.text = planta_actual
            campo_plant.caretPosition = len(planta_actual)

            # Alternativa
            #campo_alt = session.findById("wnd[0]/usr/txtRC29L-STLAL")
            #campo_alt.setFocus()
            #campo_alt.text = alternativa
            #campo_alt.caretPosition = len(alternativa)

            # Filtro (uso)
            campo_uso = session.findById("wnd[0]/usr/ctxtRC29L-CAPID")
            campo_uso.setFocus()
            campo_uso.text = uso
            campo_uso.caretPosition = len(uso)

            # Ejecutar búsqueda y esperar resultados
            ejecutar_busqueda(session)
            pausar(pausa_entre_acciones)
            grid = esperar_cs11_completo(session, timeout=15)
            resultados.append((planta_actual, grid))
            bom_obtenido = True
            print(f"[INFO] CS11 cargado para {material} en planta {planta_actual}")

        except Exception as e:
            print(f"[WARNING] CS11 falló para {material} en planta {planta_actual}: {e}")

        if bom_obtenido:
            print(f"[INFO] BOM ya obtenido, saltando resto de plantas")
            break

    if not resultados:
        print(f"[ERROR] No se pudo obtener BOM para {material}")
        return None

    print(f"[INFO] CS11 finalizado con éxito para {material}")
    return resultados
