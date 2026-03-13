import os

#! Ruta SAP Logon
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

#! Ruta donde se exportarán archivos
EXPORT_FINAL_PATH = os.path.join(os.path.expanduser("~"), "Documents")

#! Ruta archivos de los submateriales
EXTRAER_ARCHIVO = r""

#* CONFIGURACIONES MODULOS NO MOVER A MENOS DE QUE SEA CAMBIOS O MODFICACIONES REUQERIDAS
DESCRIPCIONES = ["主板大组件\\", "主板总成\\", "主板组件\\"]
RESULT_COLUMNS = ["Number", "Descripcion"]
MENSAJE_SIN_BOM = "没有可用的 BOM"
SECUENCIA = ["2000", "2900", "2000"]
TRANSACCION = "/NCS11"
FILTRO_SAP = "1TE*"
FILTRO = "PP01"
DEFAULT = 0
PAUSA = 0.8