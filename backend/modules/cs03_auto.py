from backend.utils.sap_utils import escribir_campo, ejecutar_busqueda, pausar

def ejecutar_cs03_corregir_material(session, material, componente="1TE*", planta="2000", uso="PP01", pausa_entre_acciones=0.5):
    """
    Ejecuta la transacción CS03 solo para materiales con paréntesis numéricos.
    Devuelve el material corregido (sin paréntesis si es necesario).
    """
    print(f"[INFO] Ejecutando CS03 para {material} en planta {planta}")

    session.findById("wnd[0]").maximize()
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCS03"
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]").sendVKey(0)
    pausar(pausa_entre_acciones)
    session.findById("wnd[0]").sendVKey(4)  # selección múltiple
    pausar(pausa_entre_acciones)

    escribir_campo(session,
                   "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/"
                   "ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/"
                   "txtG_SELFLD_TAB-LOW[0,24]",
                   f"*{material}*")
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

    escribir_campo(session, "wnd[0]/usr/ctxtRC29L-WERKS", planta)
    pausar(pausa_entre_acciones)
    escribir_campo(session, "wnd[0]/usr/ctxtRC29L-CAPID", uso)
    pausar(pausa_entre_acciones)
    ejecutar_busqueda(session)
    pausar(pausa_entre_acciones)

    # Corregir material quitando paréntesis
    material_corregido = material
    if "(" in material and ")" in material:
        material_corregido = material.split("(")[0]

    print(f"[INFO] Material corregido: {material_corregido}")
    return material_corregido
