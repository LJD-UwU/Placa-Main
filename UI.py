import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from backend.config.sap_login import abrir_sap_y_login
from backend.modules.sap_cs11 import ejecutar_cs11
from backend.utils.sap_utils import exportar_bom_a_excel
import pandas as pd
import time
import threading

class SAPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatización SAP CS11")
        self.root.geometry("700x500")

        # --- Archivo Excel ---
        tk.Label(root, text="Archivo Excel de Modelos:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.excel_path = tk.StringVar()
        tk.Entry(root, textvariable=self.excel_path, width=60).grid(row=0, column=1, padx=10)
        tk.Button(root, text="Seleccionar", command=self.seleccionar_excel).grid(row=0, column=2, padx=10)

        # --- Plantas ---
        tk.Label(root, text="Plantas (separadas por coma):").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.plantas_var = tk.StringVar(value="2000,2900")
        tk.Entry(root, textvariable=self.plantas_var, width=60).grid(row=1, column=1, padx=10)

        # --- Componente ---
        tk.Label(root, text="Componente:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.componente_var = tk.StringVar(value="1TE*")
        tk.Entry(root, textvariable=self.componente_var, width=60).grid(row=2, column=1, padx=10)

        # --- Uso ---
        tk.Label(root, text="Uso:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
        self.uso_var = tk.StringVar(value="PP01")
        tk.Entry(root, textvariable=self.uso_var, width=60).grid(row=3, column=1, padx=10)

        # --- Botón Procesar ---
        self.btn_procesar = tk.Button(root, text="Procesar Modelos", command=self.iniciar_proceso)
        self.btn_procesar.grid(row=4, column=1, pady=10)

        # --- Log de ejecución ---
        self.log_area = scrolledtext.ScrolledText(root, width=80, height=20)
        self.log_area.grid(row=5, column=0, columnspan=3, padx=10, pady=10)
        self.log_area.config(state="disabled")

    def escribir_log(self, mensaje):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, f"{mensaje}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        self.root.update()

    def seleccionar_excel(self):
        archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if archivo:
            self.excel_path.set(archivo)

    def iniciar_proceso(self):
        # Ejecutar en hilo separado para no bloquear la GUI
        threading.Thread(target=self.procesar_modelos, daemon=True).start()

    def procesar_modelos(self):
        ruta_excel = self.excel_path.get()
        if not ruta_excel:
            messagebox.showerror("Error", "Debe seleccionar un archivo Excel")
            return

        plantas = [p.strip() for p in self.plantas_var.get().split(",") if p.strip()]
        componente = self.componente_var.get()
        uso = self.uso_var.get()

        # --- Leer Excel ---
        try:
            df = pd.read_excel(ruta_excel)
            modelos = df.iloc[:, 0].dropna().astype(str).tolist()
            self.escribir_log(f"[INFO] {len(modelos)} modelos cargados desde Excel")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el Excel: {e}")
            return

        # --- Conectar SAP ---
        self.escribir_log("[INFO] Conectando a SAP...")
        session = abrir_sap_y_login()
        if session is None:
            messagebox.showerror("Error", "No se pudo conectar a SAP")
            return

        # --- Procesar cada modelo ---
        for i, modelo in enumerate(modelos, start=1):
            self.escribir_log(f"\n========== {i}/{len(modelos)}: {modelo} ==========")
            try:
                grid = ejecutar_cs11(
                    session,
                    material=modelo,
                    componente=componente,
                    uso=uso,
                    plantas=plantas,
                    pausa_entre_acciones=0.5
                )
                
                if grid:
                    self.escribir_log(f"[INFO] Modelo {modelo} procesado exitosamente ✅")

                    ruta_excel_exportado = exportar_bom_a_excel(session, modelo)

                    if ruta_excel_exportado:
                        self.escribir_log(
                            f"[INFO] BOM exportado correctamente: {ruta_excel_exportado} ✅"
                        )
                else:
                    self.escribir_log(
                        f"[WARNING] Modelo {modelo} no tiene BOM disponible o CS03 no resolvió ❌"
                    )


            except Exception as e:
                self.escribir_log(f"[ERROR] Falló al procesar {modelo}: {e}")
                time.sleep(1)

        self.escribir_log("\n[FIN] Todos los modelos procesados ✅")


if __name__ == "__main__":
    root = tk.Tk()
    app = SAPApp(root)
    root.mainloop()
