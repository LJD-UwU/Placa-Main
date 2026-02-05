#* CONFIGURACIÓN SAP
#! Ruta SAP Logon
SAP_LOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
SAP_SYSTEM_NAME = "HQ"

EXPORT_FINAL_PATH = r"C:\Users\admin\Documents\Practicante Archivos Main\Automatizacion"

#! Ruta archivos de los submateriales
EXTRAER_ARCHIVO = r"Z:\IE-SAP\1) BOM files\e) Submaterial Usage"

#! Credenciales SAP
SAP_USER = "RD"
SAP_PASSWORD = "123"


#* CONFIGURACIONES MODULOS
DESCRIPCIONES = ["主板大组件\\", "主板总成\\", "主板组件\\"]
RESULT_COLUMNS = ["Number", "Descripcion"]
MENSAJE_SIN_BOM = "没有可用的 BOM"
PLANTAS = ["2000", "2900"]
TRANSACCION = "/NCS11"
FILTRO_SAP = "1TE*"
PLANTA1 = "2000"
FILTRO = "PP01"
DEFAULT = 0
PAUSA = 0.8
