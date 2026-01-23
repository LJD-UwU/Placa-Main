import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import time, os
from datetime import datetime

from backend.config.sap_login import abrir_sap_y_login
from backend.modules.sap_cs11 import ejecutar_cs11
from backend.utils.txt_to_xlsx import exportar_bom_a_xls, convertir_xls_a_csv_y_xlsx, CSV_FOLDER, XLSX_FOLDER
from backend.scripts.extract_mainboard import extract_descripcion_numbers

DESCRIPCIONES_BUSCAR = ["主板大组件\\"]

class SAPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SAP CS11 Automatización Mainboard")
        self.root.geometry("850x600")

        # ---------------- UI ----------------
        tk.Label(root, text="Archivo Excel Modelos:").grid(row=0, column=0, sticky="w", padx=10)
        self.excel_path = tk.StringVar()
        tk.Entry(root, textvariable=self.excel_path, width=65).grid(row=0, column=1)
        tk.Button(root, text="Seleccionar", command=self.seleccionar_excel).grid(row=0, column=2)

        tk.Label(root, text="Plantas:").grid(row=1, column=0, sticky="w", padx=10)
        self.plantas_var = tk.StringVar(value="2000,2900")
        tk.Entry(root, textvariable=self.plantas_var, width=65).grid(row=1, column=1)

        tk.Label(root, text="Componente:").grid(row=2, column=0, sticky="w", padx=10)
        self.componente_var = tk.StringVar(value="1TE*")
        tk.Entry(root, textvariable=self.componente_var, width=65).grid(row=2, column=1)

        tk.Label(root, text="Uso:").grid(row=3, column=0, sticky="w", padx=10)
        self.uso_var = tk.StringVar(value="PP01")
        tk.Entry(root, textvariable=self.uso_var, width=65).grid(row=3, column=1)

        self.btn_procesar = tk.Button(root, text="INICIAR PROCESO", bg="green", fg="white", command=self.iniciar)
        self.btn_procesar.grid(row=4, column=1, pady=10)

        self.log = scrolledtext.ScrolledText(root, width=105, height=25)
        self.log.grid(row=5, column=0, columnspan=3)
        self.log.config(state="disabled")

        # ---------------- DATA ----------------
        self.df_resultado = pd.DataFrame(columns=["Modelo", "Number", "Descripcion"])
        self.df_proceso = pd.DataFrame(columns=["FechaHora", "Modelo", "Planta", "Paso", "Accion", "Estado", "Detalle"])

    # ---------- LOG ----------
    def log_msg(self, msg):
        self.log.config(state="normal")
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state="disabled")
        self.root.update()

    # ---------- REGISTRO PROCESO ----------
    def registrar(self, modelo, planta, paso, accion, estado, detalle=""):
        self.df_proceso.loc[len(self.df_proceso)] = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            modelo, planta, paso, accion, estado, detalle
        ]

    # ---------- FILE ----------
    def seleccionar_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f:
            self.excel_path.set(f)

    # ---------- START ----------
    def iniciar(self):
        self.btn_procesar.config(state="disabled")
        self.log_msg("🔵 Iniciando proceso...")

        ruta = self.excel_path.get()
        if not ruta:
            messagebox.showerror("Error", "Selecciona Excel")
            return

        df = pd.read_excel(ruta)
        self.modelos = df.iloc[:, 0].dropna().astype(str).tolist()
        self.log_msg(f"Modelos cargados: {len(self.modelos)}")

        # SAP Login
        self.log_msg("Conectando SAP...")
        self.session = abrir_sap_y_login()
        if not self.session:
            messagebox.showerror("SAP ERROR", "No login")
            return

        self.modelo_index = 0
        self.root.after(100, self.procesar_modelo)

    # ---------- MAIN LOOP ----------
    def procesar_modelo(self):
        if self.modelo_index >= len(self.modelos):
            self.guardar_excel_final()
            self.log_msg("✅ PROCESO TERMINADO")
            self.btn_procesar.config(state="normal")
            return

        modelo = self.modelos[self.modelo_index]
        plantas = [p.strip() for p in self.plantas_var.get().split(",")]

        self.log_msg(f"\n====== {modelo} ======")

        for planta in plantas:
            paso = 1
            try:
                self.registrar(modelo, planta, paso, "Iniciar CS11", "OK")
                grid = ejecutar_cs11(self.session, modelo, self.componente_var.get(), self.uso_var.get(), [planta])

                if not grid:
                    self.registrar(modelo, planta, paso, "CS11", "FAIL", "No BOM")
                    continue

                paso += 1
                self.registrar(modelo, planta, paso, "BOM cargado", "OK")

                # Export XLS
                ruta_xls = exportar_bom_a_xls(self.session, modelo)
                if not ruta_xls:
                    self.registrar(modelo, planta, paso, "Exportar XLS", "FAIL")
                    continue

                paso += 1
                self.registrar(modelo, planta, paso, "Export XLS", "OK", ruta_xls)

                nombre = os.path.basename(ruta_xls).replace(".XLS", "")
                ruta_csv = os.path.join(CSV_FOLDER, nombre + ".csv")
                ruta_xlsx = os.path.join(XLSX_FOLDER, nombre + ".xlsx")

                ruta_xlsx = convertir_xls_a_csv_y_xlsx(ruta_xls, ruta_csv, ruta_xlsx)
                paso += 1
                self.registrar(modelo, planta, paso, "Convertir XLSX", "OK", ruta_xlsx)

                # Extract
                df_modelo = extract_descripcion_numbers(ruta_xlsx, DESCRIPCIONES_BUSCAR)
                paso += 1
                if df_modelo.empty:
                    self.registrar(modelo, planta, paso, "Extracción", "NO DATA")
                else:
                    df_modelo["Modelo"] = modelo
                    df_modelo = df_modelo[["Modelo", "Number", "Descripcion"]]
                    self.df_resultado = pd.concat([self.df_resultado, df_modelo], ignore_index=True)

                    self.registrar(modelo, planta, paso, "Extracción", "OK", f"{len(df_modelo)} registros")

                    for _, r in df_modelo.iterrows():
                        self.log_msg(f"{r['Modelo']} | {r['Number']} | {r['Descripcion']}")

            except Exception as e:
                self.registrar(modelo, planta, paso, "ERROR", "FAIL", str(e))
                self.log_msg(f"❌ ERROR {modelo}: {e}")

        self.modelo_index += 1
        self.root.after(500, self.procesar_modelo)

    # ---------- SAVE ----------
    def guardar_excel_final(self):
        ruta = os.path.join(XLSX_FOLDER, "RESULTADO_FINAL_MAINBOARD.xlsx")
        with pd.ExcelWriter(ruta) as writer:
            self.df_resultado.to_excel(writer, sheet_name="RESULTADO", index=False)
            self.df_proceso.to_excel(writer, sheet_name="PROCESO_DETALLADO", index=False)

        self.log_msg(f"📁 Excel guardado: {ruta}")


# ---------- RUN ----------
if __name__ == "__main__":
    root = tk.Tk()
    SAPApp(root)
    root.mainloop()
