# 🤖 MB-Automator (Automatización SAP + Excel)

Aplicación de automatización basada en Python diseñada para procesar **datos de Mainboard y Motherboard desde SAP**, limpiar archivos Excel y generar estructuras listas para análisis.

Combina automatización SAP, procesamiento de datos e interfaz gráfica para ejecutar todo el flujo de trabajo de manera eficiente.

---

## 🚀 Funcionalidades

- 🔐 Inicio de sesión automático en SAP
- 📥 Ejecución de la transacción **CS11**
- 📊 Exportación de BOM desde SAP
- 🔄 Conversión de archivos `.xls` → `.xlsx`
- 🧹 Limpieza avanzada de Excel
- 🧠 Procesamiento de:
  - Mainboard
  - Motherboard
- 🧩 Integración de submateriales (BOM)
- 🖥️ Interfaz gráfica (Tkinter)
- 📁 Organización automática de archivos
- 📌 Actualización automática del Excel principal

---

## 🧠 Flujo de trabajo del sistema

```text
1. El usuario carga el archivo Excel (modelos)
2. Inicio de sesión automático en SAP
3. Ejecución de CS11 para cada modelo
4. Exportación del BOM
5. Conversión a Excel
6. Limpieza de datos
7. Procesamiento de Mainboard / Motherboard
8. Integración de submateriales
9. Actualización del Excel principal
10. Generación de resultados finales
```

---

## 📂 Estructura del proyecto

```
backend/
│
├── config/              # Configuración de SAP
│   ├── sap_login.py
│   ├── credenciales_loader.py
│   └── sap_config.py
│
├── modules/             # Lógica principal de procesamiento
│   ├── cs11.py
│   ├── extract_mainboard.py
│   ├── procesar_mainboard_P2.py
│   ├── procesar_motherboard_P1.py
│   └── Modules_2/
│         ├── procesar_motherboard.py
│         └── procesar_mainboard.py
│
├── utils/               # Utilidades y limpieza de datos
│   ├── clean_excel.py
│   ├── clean_excel_p2.py
│   ├── sap_utils.py
│   ├── txt_to_xlsx.py
│   └── utils_2/
│           └── xlsx_m2.py
│
├── UI/                  # Interfaz secundaria
│   └── motherboard_app.py
│
├── Helpers/             # Funciones auxiliares
│   └── helper.py
│
├── IMG/                 # Recursos visuales
│   └── logo.png
└── UI.py                # Interfaz principal
```

---

## 🖥️ Interfaz gráfica

La aplicación incluye una interfaz basada en Tkinter que permite:

- Seleccionar un archivo Excel
- Ejecutar procesos:
  - Procesar Mainboard
  - Procesar Motherboard
  - Limpieza de archivos
- Inicio de sesión en SAP
- Visualización de logs en tiempo real

---

## 📋 Requisitos del sistema

| Requisito | Detalle |
|-----------|---------|
| **OS** | Windows 10 / 11 (64-bit) — obligatorio |
| **Python** | 3.11 a 3.13 |
| **Microsoft Excel** | Instalado y con licencia activa |
| **SAP GUI** | Instalado y configurado |

> ⚠️ Este proyecto no es compatible con Linux/macOS debido a la dependencia de pywin32 (COM) y SAP GUI para Windows.

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

> Si el comando falla, búscalo en: `.venv\Scripts\pywin32_postinstall.py`

---

## ▶️ Uso

```bash
python UI.py
```

Al iniciar la app por primera vez, aparecerá un aviso solicitando las credenciales de SAP. Ve a **🔐 Login SAP** e ingrésalas antes de procesar.

### Flujo de trabajo

1. **Seleccionar Excel** — archivo con los materiales (MATERIAL, PROCESS, etc.)
2. **Login SAP** — ingresa tus credenciales; SAP se abrirá automáticamente
3. **Procesar 1TE** — extrae BOMs desde SAP para cada modelo
4. **Motherboard** — procesa y actualiza las columnas de motherboard en el Excel
5. **Resultados** — abre la carpeta con los archivos generados

---

## 📊 Procesamiento de datos

### 🔹 Mainboard

- Limpieza de columnas
- Identificación de materiales
- Lógica de LEVEL
- Detección de caracteres chinos
- Extracción de PCB
- Integración de BOM

---

### 🔹 Motherboard

- Procesamiento estructurado por modelo
- Separación de materiales
- Aplicación de lógica de negocio

---

### 🔹 Limpieza de Excel

- Eliminación de filas/columnas innecesarias
- Reorganización de estructura
- Formato automático

---

## ⚙️ Funciones clave

- `cs11.py` → Automatización SAP
- `clean_excel_p2.py` → Limpieza avanzada
- `procesar_mainboard_P2.py` → Lógica principal de Mainboard
- `sap_utils.py` → Funciones auxiliares de SAP
- `txt_to_xlsx.py` → Conversión de archivos

---

## 📁 Archivos generados

- Archivos BOM procesados
- Excel actualizado con:
  - MATERIAL
  - PROCESS
  - MAINBOARD PART NUMBER
- Carpetas organizadas automáticamente

---

## ✅ Resultado

✔ Automatización completa SAP + Excel
✔ Reducción del trabajo manual
✔ Procesamiento masivo de modelos
✔ Datos listos para análisis

---