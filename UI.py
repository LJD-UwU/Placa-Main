import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os
import re

from backend.modules.extract_mainboard import extract_descripcion_numbers
from backend.utils.clean_excel import limpiar_excel_mainboard
from backend.modules.procesar_mainboard_P1 import procesar_number
from backend.modules.prosesar_mainboard_P2 import procesar_material_desde_mainboard
from backend.config.sap_login import abrir_sap_y_login
from backend.modules.cs11 import ejecutar_cs11

from backend.utils.txt_to_xlsx import (
    exportar_bom_a_xls,
    convertir_xls_a_xlsx,
    MAINBOARD_1_FILES_FOLDER,
    MODEL_FILES_FOLDER
)

# ============================================================
# CONFIGURACIÓN
# ============================================================
DESCRIPCIONES_BUSCAR = ["主板大组件\\", "主板总成\\", "主板组件\\"]

# ============================================================
# UI
# ============================================================
class SAPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatización SAP CS11")
        self.root.geometry("760x560")

        # ---------- UI ----------
        tk.Label(root, text="Archivo Excel de Modelos:").grid(row=0, column=0, sticky="w", padx=10)
        self.excel_path = tk.StringVar()
        tk.Entry(root, textvariable=self.excel_path, width=60).grid(row=0, column=1)
        tk.Button(root, text="Seleccionar", command=self.seleccionar_excel).grid(row=0, column=2)

        tk.Label(root, text="Plantas (coma):").grid(row=1, column=0, sticky="w", padx=10)
        self.plantas_var = tk.StringVar(value="2000,2900")
        tk.Entry(root, textvariable=self.plantas_var, width=60).grid(row=1, column=1)

        tk.Label(root, text="Componente:").grid(row=2, column=0, sticky="w", padx=10)
        self.componente_var = tk.StringVar(value="1TE*")
        tk.Entry(root, textvariable=self.componente_var, width=60).grid(row=2, column=1)

        tk.Label(root, text="Uso:").grid(row=3, column=0, sticky="w", padx=10)
        self.uso_var = tk.StringVar(value="PP01")
        tk.Entry(root, textvariable=self.uso_var, width=60).grid(row=3, column=1)

        self.btn_procesar = tk.Button(root, text="Procesar Modelos", command=self.iniciar)
        self.btn_procesar.grid(row=4, column=1, pady=10)

        self.log = scrolledtext.ScrolledText(root, width=92, height=25)
        self.log.grid(row=5, column=0, columnspan=3, padx=10)
        self.log.config(state="disabled")

        # ---------- DATA ----------
        self.modelos = []
        self.idx = 0
        self.session = None
        self.df_todos = pd.DataFrame(columns=["Modelo", "Planta", "Number", "Descripcion"])

    # ============================================================
    # HELPERS
    # ============================================================
    def log_msg(self, msg):
        self.log.config(state="normal")
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state="disabled")
        self.root.update()

    def seleccionar_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f:
            self.excel_path.set(f)

    # ============================================================
    # FLUJO PRINCIPAL
    # ============================================================
    def iniciar(self):
        if not self.excel_path.get():
            messagebox.showwarning("Atención", "Selecciona un archivo Excel primero")
            return

        self.btn_procesar.config(state="disabled")
        self.log_msg("[INFO] Automatización iniciada...")
        self.root.after(100, self.cargar_excel)

    def cargar_excel(self):
        try:
            df = pd.read_excel(self.excel_path.get())
            self.modelos = df.iloc[:, 0].dropna().astype(str).tolist()
            self.log_msg(f"[INFO] {len(self.modelos)} modelos cargados")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.btn_procesar.config(state="normal")
            return

        self.log_msg("[INFO] Conectando a SAP...")
        self.session = abrir_sap_y_login()
        if not self.session:
            messagebox.showerror("Error", "Conexión a SAP fallida")
            self.btn_procesar.config(state="normal")
            return

        self.idx = 0
        self.root.after(300, self.procesar_modelo)

    def procesar_modelo(self):
        if self.idx >= len(self.modelos):
            self.guardar_excel_final()
            self.log_msg("\n[FIN] Procesamiento de modelos completado ✅")
            self.btn_procesar.config(state="normal")
            return

        modelo = self.modelos[self.idx]
        plantas = [p.strip() for p in self.plantas_var.get().split(",") if p.strip()]

        self.log_msg(f"\n===== {self.idx + 1}/{len(self.modelos)} → {modelo} =====")

        try:
            resultados = ejecutar_cs11(
                self.session,
                material=modelo,
                componente=self.componente_var.get(),
                uso=self.uso_var.get(),
                plantas=plantas
            )

            if resultados:
                for planta, _ in resultados:
                    ruta_xls = exportar_bom_a_xls(self.session, modelo, mainboard=False)
                    base = re.sub(r'[\\/*?:"<>|]', "_", os.path.basename(ruta_xls).replace(".XLS", ""))
                    ruta_xlsx = os.path.join(MODEL_FILES_FOLDER, f"{base}.xlsx")
                    convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)

                    df_modelo = extract_descripcion_numbers(
                        input_xlsx=ruta_xlsx,
                        modelo=modelo,
                        descripcion_a_buscar=DESCRIPCIONES_BUSCAR
                    )

                    if not df_modelo.empty:
                        df_modelo["Modelo"] = modelo
                        df_modelo["Planta"] = planta
                        self.df_todos = pd.concat([self.df_todos, df_modelo], ignore_index=True)

        except Exception as e:
            self.log_msg(f"[ERROR] {e}")

        self.idx += 1
        self.root.after(500, self.procesar_modelo)

    # ============================================================
    # MAINBOARD NIVEL 1 + NIVEL 2 AUTOMÁTICO
    # ============================================================
    def guardar_excel_final(self):
        self.log_msg("\n[INFO] Procesando ensambles de mainboard...")

        for _, row in self.df_todos.iterrows():
            number = str(row["Number"]).strip()

            # Evitar reprocesar nivel 1
            if any(number in f for f in os.listdir(MAINBOARD_1_FILES_FOLDER)):
                self.log_msg(f"[SKIP] Mainboard {number} ya procesado")
                continue

            try:
                # ---------------- Nivel 1 ----------------
                ruta_xls = procesar_number(
                    session=self.session,
                    number=number,
                    planta="2000",
                    capid=self.uso_var.get()
                )

                base = re.sub(r'[\\/*?:"<>|]', "_", os.path.basename(ruta_xls).replace(".XLS", ""))
                ruta_xlsx = os.path.join(MAINBOARD_1_FILES_FOLDER, f"{base}.xlsx")
                convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
                limpiar_excel_mainboard(ruta_xlsx)
                self.log_msg(f"[OK] Mainboard {number} nivel 1 procesado y limpio")

                # ---------------- Nivel 2 ----------------
                try:
                    procesar_material_desde_mainboard(
                        session=self.session,
                        ruta_mainboard_xlsx=ruta_xlsx,
                        uso=self.uso_var.get()
                    )
                    self.log_msg(f"[OK] Segundo nivel generado desde {number}")

                except Exception as e:
                    self.log_msg(f"[WARN] Segundo nivel falló: {e}")

            except Exception as e:
                self.log_msg(f"[ERROR] Mainboard {number} falló: {e}")

        self.log_msg("\n[FIN] Flujo completo terminado ✅")

# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = SAPApp(root)
    root.mainloop()
