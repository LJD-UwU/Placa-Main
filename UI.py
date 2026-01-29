import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os, re, time

from backend.modules.extract_mainboard import extract_descripcion_numbers
from backend.utils.clean_excel import limpiar_excel_mainboard
from backend.modules.procesar_mainboard_P1 import procesar_number
from backend.modules.prosesar_mainboard_P2 import procesar_material_desde_mainboard
from backend.config.sap_login import abrir_sap_y_login
from backend.modules.cs11 import ejecutar_cs11
from backend.utils.txt_to_xlsx import (
    exportar_bom_a_xls,
    convertir_xls_a_xlsx, 
    MODEL_FILES_FOLDER,
    MAINBOARD_1_FILES_FOLDER,
    MAINBOARD_2_FILES_FOLDER
)

# ================= CONFIG =================
DESCRIPCIONES_BUSCAR = ["主板大组件\\", "主板总成\\", "主板组件\\"]
PLANTAS = ["2000", "2900"]
COMPONENTE = "1TE*"
USO = "PP01"


class SAPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatización SAP")
        self.root.geometry("460x420")
        self.root.resizable(False, False)

        # ---- estado ----
        self.animando = False
        self.anim_dots = 0
        self.start_time = None

        # ---- estilo ----
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("TProgressbar", thickness=10)

        # ---- título ----
        ttk.Label(root, text="Automatización SAP", style="Title.TLabel").pack(pady=(8, 0))
        ttk.Label(
            root,
            text="Procesamiento automático de modelos y mainboards",
            foreground="gray"
        ).pack(pady=(0, 6))

        main = ttk.Frame(root, padding=6)
        main.pack(fill="both", expand=True)

        # ---- archivo ----
        fila_file = ttk.Frame(main)
        fila_file.pack(fill="x", pady=4)

        self.excel_path = tk.StringVar()
        ttk.Entry(fila_file, textvariable=self.excel_path)\
            .pack(side="left", fill="x", expand=True)

        ttk.Button(fila_file, text="📂", width=3,
        command=self.seleccionar_excel).pack(side="left", padx=4)

        # ---- progreso ----
        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.pack(fill="x", pady=6)

        # ---- botones ----
        fila_btn = ttk.Frame(main)
        fila_btn.pack(pady=4)

        self.btn_procesar = ttk.Button(fila_btn, text="▶ Procesar", command=self.iniciar)
        self.btn_procesar.pack(side="left", padx=4)

        self.btn_open = ttk.Button(
            fila_btn, text="📁 Resultados",
            command=self.abrir_resultados,
            state="disabled"
        )
        self.btn_open.pack(side="left", padx=4)

        # ---- log ----
        frame_log = ttk.LabelFrame(main, text="Log")
        frame_log.pack(fill="both", expand=True, pady=(6, 0))

        self.log = scrolledtext.ScrolledText(
            frame_log, height=9, font=("Consolas", 9)
        )
        self.log.pack(fill="both", expand=True, padx=5, pady=5)
        self.log.config(state="disabled")

        self.log.tag_config("INFO", foreground="blue")
        self.log.tag_config("OK", foreground="green")
        self.log.tag_config("ERROR", foreground="red")

        # ---- estado ----
        self.status = tk.StringVar(value="Estado: Listo")
        ttk.Label(root, textvariable=self.status, anchor="w")\
            .pack(fill="x", side="bottom", padx=6, pady=4)

        # ---- data ----
        self.modelos = []
        self.idx = 0
        self.session = None
        self.df_todos = pd.DataFrame(
            columns=["Modelo", "Planta", "Number", "Descripcion"]
        )

    # ================= LOG =================
    def log_msg(self, msg, tag="INFO"):
        self.log.config(state="normal")
        self.log.insert(tk.END, msg + "\n", tag)
        self.log.see(tk.END)
        self.log.config(state="disabled")
        self.root.update()

    # ================= ESTADO =================
    def set_status(self, msg, animar=False):
        self.animando = False
        self.anim_dots = 0
        if animar:
            self.animando = True
            self.animar_estado(msg)
        else:
            self.status.set(f"Estado: {msg}")
        self.root.update()

    def animar_estado(self, texto):
        if not self.animando:
            return
        self.anim_dots = (self.anim_dots + 1) % 4
        self.status.set(f"Estado: {texto}{'.' * self.anim_dots}")
        self.root.after(500, lambda: self.animar_estado(texto))

    def actualizar_tiempo(self):
        if not self.start_time:
            return
        m, s = divmod(int(time.time() - self.start_time), 60)
        base = self.status.get().split(" | ")[0]
        self.status.set(f"{base} | ⏱ {m:02d}:{s:02d}")
        if self.btn_procesar["state"] == "disabled":
            self.root.after(1000, self.actualizar_tiempo)

    # ================= UI =================
    def seleccionar_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f:
            self.excel_path.set(f)

    def abrir_resultados(self):
        path = os.path.abspath(MAINBOARD_2_FILES_FOLDER)
        if os.path.exists(path):
            os.startfile(path)

    # ================= FLUJO  =================
    def iniciar(self):
        if not self.excel_path.get():
            messagebox.showwarning("Atención", "Selecciona un Excel")
            return

        self.btn_procesar.config(state="disabled")
        self.btn_open.config(state="disabled")
        self.progress["value"] = 0

        self.start_time = time.time()
        self.actualizar_tiempo()

        self.log_msg("[INFO] Automatización iniciada")
        self.set_status("Cargando Excel", animar=True)
        self.root.after(100, self.cargar_excel)

    def cargar_excel(self):
        try:
            df = pd.read_excel(self.excel_path.get())
            self.modelos = df.iloc[:, 0].dropna().astype(str).tolist()
            self.log_msg(f"[OK] {len(self.modelos)} modelos cargados", "OK")
        except Exception as e:
            self.log_msg(f"[ERROR] {e}", "ERROR")
            self.btn_procesar.config(state="normal")
            return

        self.set_status("Conectando a SAP", animar=True)
        self.session = abrir_sap_y_login()
        self.animando = False

        if not self.session:
            self.log_msg("[ERROR] Conexión SAP fallida", "ERROR")
            self.btn_procesar.config(state="normal")
            return

        self.log_msg("[OK] Conectado a SAP", "OK")
        self.idx = 0
        self.root.after(200, self.procesar_modelo)

    def procesar_modelo(self):
        total = len(self.modelos)

        if self.idx >= total:
            self.log_msg("\n[INFO] Iniciando procesamiento de mainboards y limpieza", "INFO")
            self.guardar_excel_final()
            self.set_status("Finalizado ✅")
            self.progress["value"] = 100
            self.btn_open.config(state="normal")
            self.btn_procesar.config(state="normal")
            self.log_msg("[OK] Proceso completo", "OK")
            return

        modelo = self.modelos[self.idx]
        self.progress["value"] = int(((self.idx + 1) / total) * 100)

        self.set_status(f"Modelo {self.idx + 1}/{total}")
        self.log_msg(f"\n▶ Modelo {self.idx + 1}/{total}: {modelo}", "INFO")

        try:
            self.log_msg("  • Ejecutando CS11...", "INFO")
            resultados = ejecutar_cs11(
                self.session,
                material=modelo,
                componente=COMPONENTE,
                uso=USO,
                plantas=PLANTAS
            )

            if not resultados:
                self.log_msg("  ⚠ Sin resultados CS11", "INFO")

            for planta, _ in resultados:
                self.log_msg(f"  • Planta {planta}: exportando BOM", "INFO")

                ruta_xls = exportar_bom_a_xls(self.session, modelo, mainboard=False)
                self.log_msg("    ✓ BOM exportado", "OK")

                base = re.sub(r'[\\/*?:"<>|]', "_",
                    os.path.basename(ruta_xls).replace(".XLS", ""))
                ruta_xlsx = os.path.join(MODEL_FILES_FOLDER, f"{base}.xlsx")

                convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
                self.log_msg("    ✓ Convertido a XLSX", "OK")

                self.log_msg("    • Analizando descripciones", "INFO")
                df_modelo = extract_descripcion_numbers(
                    input_xlsx=ruta_xlsx,
                    modelo=modelo,
                    descripcion_a_buscar=DESCRIPCIONES_BUSCAR
                )

                if df_modelo.empty:
                    self.log_msg("    ⚠ Sin mainboards encontrados", "INFO")
                else:
                    self.log_msg(f"    ✓ {len(df_modelo)} mainboards encontrados", "OK")
                    df_modelo["Modelo"] = modelo
                    df_modelo["Planta"] = planta
                    self.df_todos = pd.concat(
                        [self.df_todos, df_modelo], ignore_index=True
                    )

        except Exception as e:
            self.log_msg(f"[ERROR] {e}", "ERROR")

        self.idx += 1
        self.root.after(200, self.procesar_modelo)


    def guardar_excel_final(self):
        self.set_status("Procesando mainboards")

        for folder in [MODEL_FILES_FOLDER,
            MAINBOARD_1_FILES_FOLDER,
            MAINBOARD_2_FILES_FOLDER]:
            os.makedirs(folder, exist_ok=True)

        for _, row in self.df_todos.iterrows():
            number = str(row["Number"]).strip()
            if any(number in f for f in os.listdir(MAINBOARD_1_FILES_FOLDER)):
                continue

            try:
                ruta_xls = procesar_number(self.session, number, "2000", USO)
                base = re.sub(r'[\\/*?:"<>|]', "_",
                os.path.basename(ruta_xls).replace(".XLS", ""))
                ruta_xlsx = os.path.join(MAINBOARD_1_FILES_FOLDER, f"{base}.xlsx")
                convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
                limpiar_excel_mainboard(ruta_xlsx)
                procesar_material_desde_mainboard(self.session, ruta_xlsx, USO)

            except Exception as e:
                self.log_msg(f"[ERROR] Mainboard {number}: {e}", "ERROR")


if __name__ == "__main__":
    root = tk.Tk()
    app = SAPApp(root)
    root.mainloop()
