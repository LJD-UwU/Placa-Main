import pandas as pd
import re

def extract_descripcion_numbers(input_xlsx, descripcion_a_buscar, output_xlsx=None, skiprows=0):
    """
    Busca en el XLSX las filas que contengan la parte fija de la descripción indicada
    y toma el número de la celda a la izquierda, guardando todo en un DataFrame.
    
    Parámetros:
        input_xlsx (str): Ruta del archivo XLSX a analizar.
        descripcion_a_buscar (str o lista): Parte fija de la descripción a buscar.
        output_xlsx (str, opcional): Ruta para guardar el resultado en Excel.
        skiprows (int): Filas a saltar al leer el Excel.
        
    Retorna:
        pd.DataFrame con columnas ["Number", "Descripcion"]
    """
    # Aceptar lista o string
    if isinstance(descripcion_a_buscar, str):
        descripcion_a_buscar = [descripcion_a_buscar]

    # Leer todo el Excel
    try:
        df = pd.read_excel(input_xlsx, header=None, skiprows=skiprows)
    except Exception as e:
        print(f"[ERROR] No se pudo abrir {input_xlsx}: {e}")
        return pd.DataFrame(columns=["Number", "Descripcion"])

    resultados = []

    # Iterar todas las filas
    for row in df.itertuples(index=False):
        for i, cell in enumerate(row):
            if pd.isna(cell):
                continue
            cell_str = str(cell)
            # Buscar si contiene la parte fija
            if any(desc in cell_str for desc in descripcion_a_buscar):
                # Tomar número de la celda izquierda
                number = None
                if i > 0:
                    left_cell = row[i-1]
                    if pd.notna(left_cell):
                        match = re.search(r'\d+', str(left_cell))
                        if match:
                            number = match.group()
                resultados.append([number, cell_str])

    df_result = pd.DataFrame(resultados, columns=["Number", "Descripcion"])

    # Guardar Excel si se pidió
    if output_xlsx and not df_result.empty:
        try:
            with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
                df_result.to_excel(writer, index=False)
            print(f"[INFO] Extracción completada y guardada en: {output_xlsx}")
        except Exception as e:
            print(f"[ERROR] No se pudo guardar Excel: {e}")

    return df_result
