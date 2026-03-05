import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from PIL import Image, ImageTk
import os, re
from datetime import datetime
import pandas as pd
import threading
from backend.config.sap_login import abrir_sap_y_login
from backend.config.sap_config import FILTRO
from backend.utils.txt_to_xlsx import MOTHERBOARD_FILES,MOTHERBOARD_1_FILES_FOLDER,MOTHERBOARD_2_FILES_FOLDER
from backend.Modules_2.procesar_motherboard import procesar_numbers_desde_listas

class MainboardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MBAutomator - Motherboard")
        self.root.geometry("410x420")
        self.root.resizable(False, False)

        #! Icono
        try:
            img = Image.open("IMG/logo.png").resize((256, 256))
            icon = ImageTk.PhotoImage(img)
            self.root.iconphoto(True, icon)
        except:
            pass

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("TProgressbar", thickness=10)

        ttk.Label(root, text="Automatización SAP", style="Title.TLabel").pack(pady=(8, 0))
        ttk.Label(root, text="Procesamiento de Excel para Mainboards", foreground="gray").pack(pady=(0, 6))

        main = ttk.Frame(root, padding=6)
        main.pack(fill="both", expand=True)

        #! Campo de selección de Excel
        fila_file = ttk.Frame(main)
        fila_file.pack(fill="x", pady=4)
        self.excel_path = tk.StringVar()
        ttk.Entry(fila_file, textvariable=self.excel_path).pack(side="left", fill="x", expand=True)
        ttk.Button(fila_file, text="📂", width=3, command=self.seleccionar_excels).pack(side="left", padx=4)

        #! Barra de progreso opcional
        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.pack(fill="x", pady=6)

        #! Botones principales
        fila_btn = ttk.Frame(main)
        fila_btn.pack(pady=4)
        self.btn_procesar = ttk.Button(fila_btn, text="▶ Procesar", command=self.iniciar_procesamiento, state="normal")
        self.btn_procesar.pack(side="left", padx=4)
        self.btn_limpiar = ttk.Button(fila_btn, text="🧹 Limpiar", command=self.limpiar_log)
        self.btn_limpiar.pack(side="left", padx=4)
        self.btn_resultados = ttk.Button(fila_btn, text="📁 Resultados", command=self.abrir_resultados, state="disabled")
        self.btn_resultados.pack(side="left", padx=4)

        #! Consola
        frame_log = ttk.LabelFrame(main, text="CONSOLA")
        frame_log.pack(fill="both", expand=True, pady=(6, 0))
        self.log = scrolledtext.ScrolledText(frame_log, height=10, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=5, pady=5)
        self.log.config(state="disabled")
        self.log.tag_config("INFO", foreground="blue")
        self.log.tag_config("OK", foreground="green")
        self.log.tag_config("ERROR", foreground="red")
        self.log.tag_config("WARNING", foreground="orange")

        #! Estado
        self.status = tk.StringVar(value="Estado: Listo")
        ttk.Label(root, textvariable=self.status, anchor="w").pack(fill="x", side="bottom", padx=6, pady=4)

        #! Datos
        self.session = None
        self.mother = []
        self.plants = []
        self.altboms = []
        self.intermodel = []

    #! =================== Funciones UI ===================
    def log_msg(self, msg, tag="INFO"):
        self.log.config(state="normal")
        self.log.insert(tk.END, msg + "\n", tag)
        self.log.see(tk.END)
        self.log.config(state="disabled")
        self.root.update()

    def seleccionar_excels(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx")])
        if files:
            self.excel_paths = list(files)
            self.log_msg(f"{len(files)} archivos seleccionados", "OK")
            
    def limpiar_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.config(state="disabled")
        self.log_msg("Log limpiado", "INFO")

    def abrir_resultados(self):
        os.makedirs(MOTHERBOARD_1_FILES_FOLDER, exist_ok=True)
        os.startfile(os.path.abspath(MOTHERBOARD_1_FILES_FOLDER))

    #! =================== Funciones principales ===================
    def conectar_sap(self):
        if not self.session:
            try:
                self.log_msg("Conectando a SAP...")
                self.session = abrir_sap_y_login()
                self.log_msg("Conectado a SAP correctamente", "OK")
                self.status.set("Estado: SAP conectado")
            except Exception as e:
                self.log_msg(f"[ERROR] {e}", "ERROR")
                self.status.set("Estado: Error de conexión")

    def cargar_excel_datos(self):
        """Carga y valida Excel"""
        if not self.excel_path.get():
            self.log_msg("No hay Excel seleccionado", "ERROR")
            return False
        try:
            df = pd.read_excel(self.excel_path.get())
            df.columns = df.columns.str.strip().str.upper()
            columnas_requeridas = ["MOTHERBOARD PART NUMBER", "PLANT", "INTERNAL MODEL"]
            faltantes = [c for c in columnas_requeridas if c not in df.columns]
            if faltantes:
                raise ValueError(f"No se encontraron las columnas: {faltantes}")

            def limpiar_columna(nombre_columna):
                return (
                    df[nombre_columna]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .str.replace(r"\.0$", "", regex=True)
                    .tolist()
                )

            self.mother = limpiar_columna("MOTHERBOARD PART NUMBER")
            self.plants = limpiar_columna("PLANT")
            self.intermodel = limpiar_columna("INTERNAL MODEL")
            self.log_msg("Excel cargado y validado correctamente", "OK")
            return True

        except Exception as e:
            self.log_msg(f"[ERROR] {e}", "ERROR")
            return False

    #! =================== Procesamiento ===================
    def iniciar_procesamiento(self):
        """Inicia el procesamiento en el hilo principal (SAP no es thread-safe)"""
        self.procesar()

    def procesar(self):
        if not hasattr(self, 'excel_paths') or not self.excel_paths:
            self.log_msg("No hay archivos Excel seleccionados", "ERROR")
            return

        self.conectar_sap()
        if not self.session:
            return

        self.log_msg("Procesamiento iniciado...", "INFO")
        os.makedirs(MOTHERBOARD_1_FILES_FOLDER, exist_ok=True)

        for i, excel_file in enumerate(self.excel_paths, start=1):
            self.log_msg(f"[INFO] Procesando archivo {i}/{len(self.excel_paths)}: {os.path.basename(excel_file)}")
            try:
                df = pd.read_excel(excel_file)
                df.columns = df.columns.str.strip().str.upper()

                # Cargar listas
                self.mother = df["MOTHERBOARD PART NUMBER"].dropna().astype(str).str.strip().tolist()
                self.plants = df["PLANT"].dropna().astype(str).str.strip().tolist()
                
                excel_salida = os.path.join(MOTHERBOARD_1_FILES_FOLDER)
                fecha = datetime.now().strftime("%Y-%m-%d")
                nombre_base = os.path.splitext(os.path.basename(excel_file))[0]
                nombre_base = re.sub(r'[\\/*?:"<>|]', "_", nombre_base)
                nombre_base = re.sub(r'^(?:\d+-)+', '', nombre_base)

                excel_salida = os.path.join(
                    MOTHERBOARD_1_FILES_FOLDER,
                    f"{fecha}-{nombre_base}.xlsx"
)
                
                # Procesar cada motherboard UNO POR UNO
                for idx, mother in enumerate(self.mother):
                    plant = self.plants[idx]
                    try:
                        self.log_msg(f"[INFO] Procesando {mother} en planta {plant}")
                        procesar_numbers_desde_listas(
                            session=self.session,
                            mother_list=[mother],
                            plant_list=[plant],
                            excel_output=excel_salida,
                            capid=FILTRO
                        )
                    except Exception as e:
                        self.log_msg(f"[ERROR] Error procesando {mother} en planta {plant}: {e}", "ERROR")

            except Exception as e:
                self.log_msg(f"[ERROR] Falló procesamiento archivo {os.path.basename(excel_file)}: {e}", "ERROR")
                
    #! Eliminar archivos .xls residuales
        for folder in [MOTHERBOARD_FILES, MOTHERBOARD_1_FILES_FOLDER, MOTHERBOARD_2_FILES_FOLDER]:
            for f in os.listdir(folder):
                ruta = os.path.join(folder, f)
                if os.path.isfile(ruta) and f.lower().endswith(".xls"):
                    try:
                        os.remove(ruta)
                    except Exception as e:
                        self.log_msg(f"[ERROR] No se pudo eliminar {f}: {e}","ERROR")

        self.log_msg("Todos los archivos procesados ✅", "OK")
        self.btn_resultados.config(state="normal")
        self.status.set("Estado: Listo")

if __name__ == "__main__":
    root = tk.Tk()
    app = MainboardApp(root)
    root.mainloop()