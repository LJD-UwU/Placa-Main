# 🤖 MB-Automator（SAP + Excel 自动化）

基于 Python 的自动化应用程序，专为从 SAP 处理**主板（Mainboard）和母板（Motherboard）数据**、清理 Excel 文件并生成可分析的数据结构而设计。

它将 SAP 自动化、数据处理和图形界面相结合，高效执行整个工作流程。

---

## 🚀 功能特性

- 🔐 自动登录 SAP
- 📥 执行事务 **CS11**
- 📊 从 SAP 导出 BOM
- 🔄 文件转换 `.xls` → `.xlsx`
- 🧹 高级 Excel 清理
- 🧠 处理：
  - 主板（Mainboard）
  - 母板（Motherboard）
- 🧩 子材料（BOM）集成
- 🖥️ 图形界面（Tkinter）
- 📁 自动文件整理
- 📌 自动更新主 Excel 文件

---

## 🧠 系统工作流程

```text
1. 用户加载 Excel 文件（型号列表）
2. 自动登录 SAP
3. 为每个型号执行 CS11
4. 导出 BOM
5. 转换为 Excel 格式
6. 数据清理
7. 主板 / 母板处理
8. 子材料集成
9. 更新主 Excel 文件
10. 生成最终结果
```

---

## 📂 项目结构

```
backend/
│
├── config/              # SAP 配置
│   ├── sap_login.py
│   ├── credenciales_loader.py
│   └── sap_config.py
│
├── modules/             # 核心处理逻辑
│   ├── cs11.py
│   ├── extract_mainboard.py
│   ├── procesar_mainboard_P2.py
│   ├── procesar_motherboard_P1.py
│   └── Modules_2/
│         ├── procesar_motherboard.py
│         └── procesar_mainboard.py
│
├── utils/               # 工具与数据清理
│   ├── clean_excel.py
│   ├── clean_excel_p2.py
│   ├── sap_utils.py
│   ├── txt_to_xlsx.py
│   └── utils_2/
│           └── xlsx_m2.py
│
├── UI/                  # 辅助界面
│   └── motherboard_app.py
│
├── Helpers/             # 辅助函数
│   └── helper.py
│
├── IMG/                 # 图像资源
│   └── logo.png
└── UI.py                # 主界面
```

---

## 🖥️ 图形界面

应用程序包含基于 Tkinter 的界面，支持以下操作：

- 选择 Excel 文件
- 运行流程：
  - 处理主板（Mainboard）
  - 处理母板（Motherboard）
  - 文件清理
- SAP 登录
- 实时日志查看

---

## 📋 系统要求

| 要求 | 详情 |
|-----------|---------|
| **操作系统** | Windows 10 / 11（64位）— 必须 |
| **Python** | 3.11 至 3.13 |
| **Microsoft Excel** | 已安装并激活许可证 |
| **SAP GUI** | 已安装并配置 |

> ⚠️ 本项目不兼容 Linux/macOS，因为依赖于 pywin32（COM）和 Windows 版 SAP GUI。

---

## 📦 主要依赖

| 包名 | 用途 |
|---------|-----|
| `pandas` | 表格数据处理 |
| `openpyxl` | 读写 `.xlsx` 文件 |
| `xlwings` | 通过 COM 自动化 Excel |
| `pywin32` | 与 SAP GUI 和 Excel 的 COM 接口（`pythoncom`）|
| `Pillow` | 图形界面的图标与 Logo |

---

## ⚙️ 安装步骤

### 1. 克隆仓库

```bash
git clone https://github.com/tu-usuario/Practicante-Placa-Main.git
cd Practicante-Placa-Main
```

### 2. 创建虚拟环境（推荐）

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. pywin32 安装后配置（必须）

```bash
python Scripts/pywin32_postinstall.py -install
```

> 如果命令失败，请在以下路径查找：`.venv\Scripts\pywin32_postinstall.py`

---

## ▶️ 使用方法

```bash
python UI.py
```

首次启动时，应用会提示输入 SAP 凭据。请前往 **🔐 SAP 登录** 并在处理前完成输入。

### 操作流程

1. **选择 Excel** — 包含材料列表的文件（含 MATERIAL、PROCESS 等列）
2. **SAP 登录** — 输入凭据，SAP 将自动打开
3. **处理 1TE** — 从 SAP 为每个型号提取 BOM
4. **母板处理** — 处理并更新 Excel 中的母板列
5. **查看结果** — 打开包含生成文件的文件夹

---

## 📊 数据处理

### 🔹 主板（Mainboard）

- 列清理
- 材料识别
- LEVEL 逻辑处理
- 中文字符检测
- PCB 提取
- BOM 集成

---

### 🔹 母板（Motherboard）

- 基于型号的结构化处理
- 材料分类
- 业务逻辑应用

---

### 🔹 Excel 清理

- 删除多余行/列
- 结构重组
- 自动格式化

---

## ⚙️ 关键函数

- `cs11.py` → SAP 自动化
- `clean_excel_p2.py` → 高级清理
- `procesar_mainboard_P2.py` → 主板核心逻辑
- `sap_utils.py` → SAP 辅助函数
- `txt_to_xlsx.py` → 文件转换

---

## 📁 生成的输出文件

- 已处理的 BOM 文件
- 更新后的 Excel，包含：
  - MATERIAL
  - PROCESS
  - MAINBOARD PART NUMBER
- 自动整理的文件夹

---

## ✅ 成果

✔ SAP + Excel 全流程自动化
✔ 减少手动操作
✔ 批量型号处理
✔ 数据即时可用于分析

---