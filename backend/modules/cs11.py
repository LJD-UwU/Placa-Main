from backend.utils.sap_utils import (
    ejecutar_busqueda,
    esperar_cs11_completo,
    pausar
)
from backend.config.sap_config import (
    TRANSACCION,
    FILTRO,
    PAUSA
)

def ejecutar_cs11(session, material, plantas, altboms, uso=FILTRO, pausa_entre_acciones=PAUSA):
    if isinstance(plantas, str):
        plantas = [plantas]

    if not plantas:
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

    for idx, planta in enumerate(plantas):
        print(f"[INFO] Procesando planta: {planta}")
        try:
            def set_campo(campo_id, valor):
                campo = session.findById(campo_id)
                campo.setFocus()
                campo.text = valor
                campo.caretPosition = len(valor)

            set_campo("wnd[0]/usr/ctxtRC29L-MATNR", material)
            set_campo("wnd[0]/usr/ctxtRC29L-WERKS", planta)
            set_campo("wnd[0]/usr/txtRC29L-STLAL", altboms)
            set_campo("wnd[0]/usr/ctxtRC29L-CAPID", uso)

            ejecutar_busqueda(session)
            pausar(pausa_entre_acciones)
            grid = esperar_cs11_completo(session, timeout=15)
            resultados.append((planta, grid))
            bom_obtenido = True
            print(f"[INFO] CS11 cargado para {material} en planta {planta}")

        except Exception as e:
            print(f"[WARNING] CS11 falló para {material} en planta {planta}: {e}")

        if bom_obtenido:
            print(f"[INFO] BOM ya obtenido, saltando resto de plantas")
            break

    if not resultados:
        print(f"[ERROR] No se pudo obtener BOM para {material}")
        return None

    print(f"[INFO] CS11 finalizado con éxito para {material}")
    return resultados