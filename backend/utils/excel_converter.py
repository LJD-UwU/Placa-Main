import os
import pandas as pd

def convert_xls_to_xlsx(ruta_xls, eliminar_xls=True):
    try:
        if not os.path.exists(ruta_xls):
            print(f"[ERROR] No existe el archivo: {ruta_xls}")
            return None

        ruta_xlsx = ruta_xls.replace(".XLS", ".xlsx").replace(".xls", ".xlsx")

        # ⚠️ SAP exporta como TEXTO, no Excel real
        df = pd.read_csv(
            ruta_xls,
            sep="\t",
            encoding="latin1",
            engine="python"
        )

        df.to_excel(ruta_xlsx, index=False, engine="openpyxl")

        if eliminar_xls:
            os.remove(ruta_xls)

        print(f"[INFO] Convertido correctamente a XLSX: {ruta_xlsx} ✅")
        return ruta_xlsx

    except Exception as e:
        print(f"[ERROR] Conversión XLS → XLSX falló: {e}")
        return None

