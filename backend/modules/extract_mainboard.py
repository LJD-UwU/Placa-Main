import pandas as pd
import re
from backend.config.sap_config import (
    RESULT_COLUMNS,
)
def extract_descripcion_numbers(input_xlsx, internal_models, descripcion_a_buscar, skiprows=0):
    if isinstance(descripcion_a_buscar, str):
        descripcion_a_buscar = [descripcion_a_buscar]

    #! Asegurarnos que internal_models sea un dato string
    internal_models = str(internal_models) if internal_models else ""

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

            #! Filtrar solo si contiene descripción y modelo
            if any(desc in cell_str for desc in descripcion_a_buscar) and internal_models in cell_str:
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
