import os

#! Ruta SAP Logon
SAP_LOGON_PATH = r""

#! Ruta donde se exportarán archivos
EXPORT_FINAL_PATH = os.path.join(os.path.expanduser("~"), "Documents")

#! Ruta archivos de los submateriales
EXTRAER_ARCHIVO = r""

# * CONFIGURACIONES MODULOS
DESCRIPCIONES = ["主板大组件\\", "主板总成\\", "主板组件\\"]
RESULT_COLUMNS = ["Number", "Descripcion"]
MENSAJE_SIN_BOM = "没有可用的 BOM"
PLANTAS = ["2000", "2900"]
SECUENCIA = ["2000", "2900", "2000"]
TRANSACCION = "/NCS11"
FILTRO_SAP = "1TE*"
PLANTA1 = "2000"
FILTRO = "PP01"
DEFAULT = 0
PAUSA = 0.8