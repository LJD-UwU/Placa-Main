import argparse
import logging
from backend.config.sap_login import abrir_sap_y_login
from backend.modules.sap_cs11 import ejecutar_cs11

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("sap_automation.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

def main(material, componente, uso, plantas):
    try:
        logging.info("Iniciando sesión SAP...")
        session = abrir_sap_y_login()
        logging.info("Sesión SAP lista ✅")

        logging.info(f"Iniciando CS11 con Material='{material}', Componente='{componente}', Plantas={plantas}")
        grids_por_planta = ejecutar_cs11(session, material, componente, uso, plantas)

        for planta, grid in grids_por_planta.items():
            logging.info(f"Planta {planta}: {grid.RowCount} filas extraídas ✅")

    except Exception as e:
        logging.error(f"Error en automatización CS11: {e}", exc_info=True)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Automatización SAP CS11")
    parser.add_argument("--material", default="58Q60SUR(01)")
    parser.add_argument("--componente", default="1TE*")
    parser.add_argument("--uso", default="PP01")
    parser.add_argument("--plantas", default="2000,2900")
    args = parser.parse_args()

    plantas_list = [p.strip() for p in args.plantas.split(",") if p.strip()]
    main(args.material, args.componente, args.uso, plantas_list)
