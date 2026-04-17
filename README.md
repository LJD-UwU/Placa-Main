# ⚡ MBAutomator

Herramienta de automatización para el procesamiento de **BOMs (Bill of Materials)** de motherboards desde SAP, con interfaz gráfica en Tkinter.

---

## 📋 Requisitos del sistema

| Requisito | Detalle |
|-----------|---------|
| **OS** | Windows 10 / 11 (64-bit) — obligatorio |
| **Python** | 3.11 a 3.13 |
| **Microsoft Excel** | Instalado y con licencia activa |
| **SAP GUI** | Instalado y configurado en el equipo |

> ⚠️ Este proyecto **no es compatible con Linux/macOS** porque depende de `pywin32` (COM) y SAP GUI para Windows.

---

## 🗂️ Estructura del proyecto

```
Practicante-Placa-Main/
├── backend/
│   ├── config/
│   │   ├── sap_config.py           # Constantes SAP (DESCRIPCIONES, FILTRO)
│   │   ├── credenciales_loader.py  # Carga y guarda credenciales SAP
│   │   └── sap_login.py            # Abre SAP y hace login automático
│   ├── Helpers/
│   │   └── helper.py               # Registro de archivos procesados
│   ├── modules/
│   │   ├── Modules_2/
│   │   │   ├── procesar_motherboard.py
│   │   │   └── procesar_mainboard.py
│   │   ├── cs11.py                         # Transacción CS11 en SAP
│   │   ├── extract_mainboard.py            # Extrae números de descripción
│   │   ├── procesar_motherboard_P1.py      # Procesamiento Fase 1
│   │   └── procesar_mainboard_P2.py        # Procesamiento Fase 2
│   ├── UI/
│   │   └── motherboard_app.py      # Ventana secundaria de motherboards
│   ├── utils/
│   │   ├── utils_2/
│   │   │   └── xlsx_m2.py
│   │   ├── clean_excel_p2.py
│   │   ├── clean_excel.py          # Limpieza de Excel mainboard
│   │   ├── sap_utils.py
│   │   └── txt_to_xlsx.py          # Conversión TXT/XLS → XLSX y rutas base
│   └── IMG/
│       ├── logo.png                # Ícono de la app
│       └── bg.png                  # Fondo de la ventana (opcional)
├── .gitignore
├── README.md
├── LICENSE
├── UI.py                           # Punto de entrada principal (UI estándar)
├── PRUEVAS.py                      # Versión alternativa con tema oscuro
└── requirements.txt
```

---

## ⚙️ Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/Practicante-Placa-Main.git
cd Practicante-Placa-Main
```

### 2. Crear entorno virtual (recomendado)

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 4. Post-instalación de pywin32 (obligatorio)

```bash
python Scripts/pywin32_postinstall.py -install
```

> Si el comando anterior falla, búscalo en: `.venv\Scripts\pywin32_postinstall.py`

---

## ▶️ Uso

```bash
python UI.py
```

Al iniciar la app por primera vez, aparecerá un aviso solicitando las credenciales de SAP. Ve a **🔐 Login SAP** e ingrésalas antes de procesar.

### Flujo de trabajo

1. **Seleccionar Excel** — archivo con la lista de materiales (columnas `MATERIAL`, `PROCESS`, etc.)
2. **Login SAP** — ingresa tus credenciales; la app abre SAP automáticamente
3. **Procesar 1TE** — extrae BOMs desde SAP para cada modelo
4. **Motherboard** — procesa y actualiza las columnas de motherboard en el Excel
5. **Resultados** — abre la carpeta con los archivos generados

---

## 📦 Dependencias principales

| Paquete | Uso |
|---------|-----|
| `pandas` | Manipulación de datos tabulares |
| `openpyxl` | Lectura y escritura de `.xlsx` |
| `xlwings` | Automatización de Excel vía COM |
| `pywin32` | Interfaz COM con SAP GUI y Excel (`pythoncom`) |
| `Pillow` | Logo e íconos de la interfaz gráfica |

---

## 🪵 Logs y errores

La consola integrada en la app muestra mensajes en tiempo real:

- 🔵 `INFO` — operaciones en curso
- 🟢 `OK` — paso completado con éxito
- 🔴 `ERROR` — fallo en algún paso (el proceso continúa con el siguiente)
- 🟡 `WARNING` — advertencia no crítica

Los errores también se guardan con `logging` estándar de Python para depuración.

---

## 📄 Licencia

Ver archivo [LICENSE](LICENSE).