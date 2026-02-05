from backend.utils.sap_utils import escribir_campo, ejecutar_busqueda, pausar
from backend.config.sap_config import (
    FILTRO_SAP,
    PLANTA1,
    FILTRO,
    PAUSA
)

def ejecutar_cs03_corregir_material(session, material, componente= FILTRO_SAP, planta=PLANTA1, uso=FILTRO, pausa_entre_acciones=PAUSA):
    
    print(f"[INFO] Ejecutando CS03 para {material} en planta {planta}")
    session.findById("wnd[0]").maximize()
    pausar(pausa_entre_acciones)

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCS03"
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]").sendVKey(0)
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]").sendVKey(4)  
    pausar(pausa_entre_acciones)

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

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    pausar(pausa_entre_acciones)
    session.findById("wnd[1]").sendVKey(2)
    pausar(pausa_entre_acciones)

    from backend.modules.cs11 import set_werks, set_capid
    set_werks(session, planta)
    set_capid(session, uso)

    ejecutar_busqueda(session)
    pausar(pausa_entre_acciones)

    material_corregido = material
    if "(" in material and ")" in material:
        material_corregido = material.split("(")[0]

    print(f"[INFO] Material corregido: {material_corregido}")
    return material_corregido
