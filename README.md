# 🤖 MB-Automator (SAP + Excel Automation)

Python-based automation application designed to process **Mainboard and Motherboard data from SAP**, clean Excel files, and generate analysis-ready structures.

It combines SAP automation, data processing, and a graphical interface to execute the entire workflow efficiently.

---

## 🚀 Features

- 🔐 Automatic SAP login
- 📥 Execution of transaction **CS11**
- 📊 BOM export from SAP
- 🔄 File conversion `.xls` → `.xlsx`
- 🧹 Advanced Excel cleaning
- 🧠 Processing of:
  - Mainboard
  - Motherboard
- 🧩 Sub-material (BOM) integration
- 🖥️ Graphical interface (Tkinter)
- 📁 Automatic file organization
- 📌 Automatic main Excel update

---

## 🧠 System Workflow

```text
1. User loads Excel file (models)
2. Automatic SAP login
3. CS11 execution for each model
4. BOM export
5. Conversion to Excel
6. Data cleaning
7. Mainboard / Motherboard processing
8. Sub-material integration
9. Main Excel update
10. Final results generation
```

---

## 📂 Estructura del proyecto

```
backend/
│
├── config/              # SAP configuration
│   ├── sap_login.py
│   ├── credenciales_loader.py
│   └── sap_config.py
│
├── modules/             # Core processing logic
│   ├── cs11.py
│   ├── extract_mainboard.py
│   ├── procesar_mainboard_P2.py
│   ├── procesar_motherboard_P1.py
│   └── Modules_2/
│         ├── procesar_motherboard.py
│         └── procesar_mainboard.py
│
├── utils/               # Utilities and data cleaning
│   ├── clean_excel.py
│   ├── clean_excel_p2.py
│   ├── sap_utils.py
│   ├── txt_to_xlsx.py
│   └── utils_2/
│           └── xlsx_m2.py
│
├── UI/                  # Secondary interface
│   └── motherboard_app.py
│
├── Helpers/             # Helper functions
│   └── helper.py
│
├── IMG/                 # Visual assets
│   └── logo.png
└── UI.py                # Main interface
```

---

## 🖥️ Graphical Interface

The application includes a Tkinter-based UI that allows:

- Selecting an Excel file
- Running processes:
  - Process Mainboard
  - Process Motherboard
  - File cleaning
- SAP login
- Real-time log visualization

---

## 📋 System Requirements

| Requirement | Details |
|-----------|---------|
| **OS** | Windows 10 / 11 (64-bit) — required |
| **Python** | 3.11 to 3.13 |
| **Microsoft Excel** | Installed with active license |
| **SAP GUI** | Installed and configured |

> ⚠️ This project is not compatible with Linux/macOS due to dependency on pywin32 (COM) and SAP GUI for Windows.


---

## 📦 Main Dependencies

| Package | Purpose |
|---------|-----|
| `pandas` | Tabular data manipulation |
| `openpyxl` | Read/write `.xlsx` |
| `xlwings` | Excel automation via COM |
| `pywin32` | COM interface with SAP GUI and Excel (`pythoncom`) |
| `Pillow` | LUI logo and icons |

---

## ⚙️ Installation

### 1. Clone the repository

```bash
git clone https://github.com/tu-usuario/Practicante-Placa-Main.git
cd Practicante-Placa-Main
```

### 2. Create virtual environment (recommended)

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Post-install pywin32 (required)

```bash
python Scripts/pywin32_postinstall.py -install
```

> If the command fails, locate it at: `.venv\Scripts\pywin32_postinstall.py`

---

## ▶️ Usage

```bash
python UI.py
```

On first launch, the app will prompt for SAP credentials. Go to 🔐 SAP Login and enter them before processing.

### Workflow

1. **Select Exce**l — file containing materials (MATERIAL, PROCESS, etc.)
2. **SAP Login** — enter credentials; SAP will open automatically
3. **Process 1TE** — extract BOMs from SAP for each model
4. **Motherboard** — process and update motherboard columns in Excel
5. **Results** — open the folder with generated files

---

## 📊 Data Processing

### 🔹 Mainboard

- Column cleaning
- Material identification
- LEVEL logic
- Chinese character detection
- PCB extraction
- BOM integration

---

### 🔹 Motherboard

- Structured model-based processing
- Material separation
- Business logic application

---

### 🔹 Excel Cleaning

- Removal of unnecessary rows/columns
- Structure reorganization
- Automatic formatting

---

## ⚙️ Key Functions

- `cs11.py` → SAP automation
- `clean_excel_p2.py` → Advanced cleaning
- `procesar_mainboard_P2.py` → Mainboard core logic
- `sap_utils.py` → SAP helper functions
- `txt_to_xlsx.py` → File conversion

---

## 📁 Generated Output

- Processed BOM files
- Updated Excel with:
  - MATERIAL
  - PROCESS
- MAINBOARD PART NUMBER
- Automatically organized folders

---

## ✅ Result

✔ Full SAP + Excel automation
✔ Reduced manual work
✔ Mass model processing
✔ Analysis-ready data

---

## 🌍 Languages

- [Spanish Readme](README_spa.md)
- [Chinese Readme](README_zh.md)