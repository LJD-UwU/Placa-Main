import pandas as pd
import re
from backend.config.sap_config import (
    RESULT_COLUMNS,
)

def extract_descripcion_numbers(input_xlsx, modelo, descripcion_a_buscar, skiprows=0):
    # Aceptar string o lista
    if isinstance(descripcion_a_buscar, str):
        descripcion_a_buscar = [descripcion_a_buscar]

    # Leer Excel
    try:
        df = pd.read_excel(input_xlsx, header=None, skiprows=skiprows)
    except Exception as e:
        print(f"[ERROR] No se pudo abrir {input_xlsx}: {e}")
        return pd.DataFrame(columns=RESULT_COLUMNS)

    resultados = []

    for row in df.itertuples(index=False):
        for i, cell in enumerate(row):
            if pd.isna(cell):
                continue
            cell_str = str(cell)

            # ✅ FILTRO: Debe contener al menos uno de los caracteres chinos y el modelo
            if any(chino in cell_str for chino in descripcion_a_buscar) and modelo in cell_str:
                number = None
                if i > 0:
                    left_cell = row[i-1]
                    if pd.notna(left_cell):
                        match = re.search(r'\d+', str(left_cell))
                        if match:
                            number = match.group()
                resultados.append([number, cell_str])

    df_result = pd.DataFrame(resultados, columns=RESULT_COLUMNS)
    return df_result
