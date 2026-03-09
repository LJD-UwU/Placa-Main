import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from PIL import Image, ImageTk
import os
from datetime import datetime
import pandas as pd

from backend.config.sap_login import abrir_sap_y_login
from backend.config.sap_config import FILTRO
from backend.utils.txt_to_xlsx import (
    MOTHERBOARD_FILES,
    MOTHERBOARD_1_FILES_FOLDER,
    MOTHERBOARD_2_FILES_FOLDER
)
from backend.Helpers.helper2 import cargar_archivos_procesados,guardar_archivo_procesado
from backend.utils.clean_excel_p2 import procesar_archivo_principal_mainboard_2
from backend.modules.Modules_2.prosesar_mainboard import procesar_material_desde_mainboard
from backend.modules.Modules_2.procesar_motherboard import procesar_numbers_desde_listas
from backend.utils.clean_excel import limpiar_excel_mainboard
from backend.utils.utils_2.xlsx_m2 import convertir_xls_a_xlsx


class MainboardApp:

    def __init__(self, root):
        self.root = root
        self.root.title("MBAutomator - Motherboard")
        self.root.geometry("410x420")
        self.root.resizable(False, False)

        try:
            img = Image.open("backend/IMG/logo.png").resize((256, 256))
            icon = ImageTk.PhotoImage(img)
            self.root.iconphoto(True, icon)
        except:
            pass

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))

        ttk.Label(root, text="Automatización SAP", style="Title.TLabel").pack(pady=(8, 0))
        ttk.Label(root, text="Procesamiento de Excel para Mainboards", foreground="gray").pack(pady=(0, 6))

        main = ttk.Frame(root, padding=6)
        main.pack(fill="both", expand=True)

        fila_file = ttk.Frame(main)
        fila_file.pack(fill="x", pady=4)

        self.excel_path = tk.StringVar()

        ttk.Entry(fila_file, textvariable=self.excel_path).pack(side="left", fill="x", expand=True)
        ttk.Button(fila_file, text="📂", width=3, command=self.seleccionar_excels).pack(side="left", padx=4)

        fila_btn = ttk.Frame(main)
        fila_btn.pack(pady=4)

        self.btn_procesar = ttk.Button(
            fila_btn,
            text="▶ Procesar",
            command=self.iniciar_procesamiento
        )
        self.btn_procesar.pack(side="left", padx=4)

        self.btn_limpiar = ttk.Button(
            fila_btn,
            text="🧹 Limpiar",
            command=self.limpiar
        )
        self.btn_limpiar.pack(side="left", padx=4)

        self.btn_resultados = ttk.Button(
            fila_btn,
            text="📁 Resultados",
            command=self.abrir_resultados,
            state="disabled"
        )
        self.btn_resultados.pack(side="left", padx=4)

        frame_log = ttk.LabelFrame(main, text="CONSOLA")
        frame_log.pack(fill="both", expand=True, pady=(6, 0))

        self.log = scrolledtext.ScrolledText(
            frame_log,
            height=10,
            font=("Consolas", 9)
        )
        self.log.pack(fill="both", expand=True, padx=5, pady=5)

        self.log.config(state="disabled")
        self.log.tag_config("INFO", foreground="blue")
        self.log.tag_config("OK", foreground="green")
        self.log.tag_config("ERROR", foreground="red")

        self.status = tk.StringVar(value="Estado: Listo")

        ttk.Label(
            root,
            textvariable=self.status,
            anchor="w"
        ).pack(fill="x", side="bottom", padx=6, pady=4)

        self.session = None
        self.mother = []
        self.plants = []
        self.internal_models = []

    #! ================= LOG =================

    def log_msg(self, msg, tag="INFO"):
        self.log.config(state="normal")
        self.log.insert(tk.END, msg + "\n", tag)
        self.log.see(tk.END)
        self.log.config(state="disabled")
        self.root.update()

    #! ================= UI UTILS =================

    def seleccionar_excels(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx")])
        if files:
            self.excel_paths = list(files)
            self.log_msg(f"[OK] {len(files)} archivos seleccionados", "OK")

    def limpiar(self):
        # Limpiar consola
        self.log.config(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.config(state="disabled")
        self.log_msg("[INFO] Limpieza iniciada", "INFO")

        folder = MOTHERBOARD_2_FILES_FOLDER

        if not os.path.exists(folder):
            self.log_msg(f"[ERROR] La carpeta {folder} no existe", "ERROR")
            return

        # Crear carpeta final
        carpeta_final = os.path.join(folder, "ARCHIVOS_FINALES")
        os.makedirs(carpeta_final, exist_ok=True)

        # Cargar archivos ya procesados desde JSON
        archivos_procesados = cargar_archivos_procesados()

        # Obtener archivos XLSX ordenados por fecha
        archivos = [
            f for f in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, f)) and f.lower().endswith(".xlsx")
            and f not in archivos_procesados  # <-- solo archivos no procesados
        ]
        archivos.sort(key=lambda x: os.path.getmtime(os.path.join(folder, x)))

        if not archivos:
            self.log_msg("[INFO] No hay archivos nuevos para procesar", "INFO")
            return

        # Procesar archivos nuevos
        for i, f in enumerate(archivos):
            ruta_excel = os.path.join(folder, f)
            salida_excel = os.path.join(carpeta_final, f"MB-BMM-{f}")

            try:
                self.log_msg(f"[INFO] Procesando archivo: {f}", "INFO")

                internal = self.internal_models[i] if i < len(self.internal_models) else ""
                plantas = self.plants[i] if i < len(self.plants) else ""

                procesar_archivo_principal_mainboard_2(
                    ruta_excel,
                    salida_excel,
                    internal,
                    plantas,
                )

                # Guardar archivo como procesado
                guardar_archivo_procesado(f)

                self.log_msg(f"[OK] Archivo procesado: {f}\n", "OK")

            except Exception as e:
                self.log_msg(f"[ERROR] No se pudo procesar {f}: {e}", "ERROR")

        self.log_msg("[INFO] Todos los archivos nuevos han sido procesados", "OK")

    def abrir_resultados(self):
        os.makedirs(MOTHERBOARD_2_FILES_FOLDER, exist_ok=True)
        os.startfile(os.path.abspath(MOTHERBOARD_2_FILES_FOLDER))

    #! ================= SAP CONEXION =================
    def conectar_sap(self):
        if not self.session:
            try:
                self.log_msg("[INFO] Conectando a SAP...")
                self.session = abrir_sap_y_login()
                self.log_msg("[OK] Conectado a SAP", "OK")
            except Exception as e:
                self.log_msg(f"[ERROR] {e}", "ERROR")

    #! ================= PROCESAMIENTO DE LA MOTHERBOARD=================

    def iniciar_procesamiento(self):
        self.log_msg("[INFO] Automatización iniciada")
        self.procesar()

    def procesar(self):
        if not hasattr(self, "excel_paths"):
            self.log_msg("[ERROR] No hay archivos Excel seleccionados", "ERROR")
            return
        self.conectar_sap()
        if not self.session:
            return
        os.makedirs(MOTHERBOARD_1_FILES_FOLDER, exist_ok=True)
        for i, excel_file in enumerate(self.excel_paths):
            try:
                self.log_msg(f"\n▶ Archivo {i+1}/{len(self.excel_paths)}: {os.path.basename(excel_file)}", "OK")
                self.log_msg("  • Leyendo Excel", "INFO")
                df = pd.read_excel(excel_file)
                df.columns = df.columns.str.strip().str.upper()
                self.mother = df["MOTHERBOARD PART NUMBER"].dropna().astype(str).tolist()
                self.plants = df["PLANT"].dropna().astype(str).tolist()
                self.internal_models = df["INTERNAL MODEL"].dropna().astype(str).tolist()
                total = len(self.mother)
                for idx, (mother, plant) in enumerate(zip(self.mother, self.plants), start=1):
                    try:
                        self.log_msg(f"\n▶ Archivo {idx}/{total}: {mother}", "OK")
                        self.log_msg(f"  • Planta {plant}")
                        fecha = datetime.now().strftime("%Y-%m-%d")
                        excel_salida = os.path.join(
                            MOTHERBOARD_1_FILES_FOLDER,
                            f"{fecha}-{mother}.xlsx"
                        )
                        
                        procesar_numbers_desde_listas(
                            session=self.session,
                            mother_list=[mother],
                            plant_list=[plant],
                            excel_output=excel_salida,
                            capid=FILTRO
                        )
                        
                        self.log_msg("    ✓ XLS generado", "OK")
                        ruta_xls = os.path.join(
                            MOTHERBOARD_1_FILES_FOLDER,
                            f"{fecha}-{mother}.XLS"
                        )
                        
                        ruta_xlsx = os.path.join(
                            MOTHERBOARD_1_FILES_FOLDER,
                            f"{fecha}-{mother}.xlsx"
                        )
                        
                        ruta_convertida = convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
                        if ruta_convertida:
                            self.log_msg("    ✓ Convertido a XLSX", "OK")
                            limpiar_excel_mainboard(ruta_convertida)
                            self.log_msg("    ✓ Excel limpiado", "OK")

                            #! ================= PROCESAMIENTO DE LAS MAIN BOARDS=================
                            try:
                                self.log_msg("    • Procesando nivel 2", "INFO")
                                resultado = procesar_material_desde_mainboard(
                                    session=self.session,
                                    ruta_mainboard_xlsx=ruta_convertida,
                                    uso=FILTRO,
                                    planta=plant
                                )
                                
                                if resultado:
                                    self.log_msg(
                                        f"    ✓ BOM nivel 2 generado: {os.path.basename(resultado)}",
                                        "OK"
                                    )
                                else:
                                    self.log_msg(
                                        f"    [ERROR] No se pudo generar BOM nivel 2 para {mother}",
                                        "ERROR"
                                    )
                                    
                            except Exception as e:
                                self.log_msg(
                                    f"[ERROR] Segundo proceso {mother}: {e}",
                                    "ERROR"
                                )
                                
                    except Exception as e:
                        self.log_msg(f"[ERROR] {mother}: {e}", "ERROR")
            except Exception as e:
                self.log_msg(
                    f"[ERROR] Falló procesamiento archivo {os.path.basename(excel_file)}: {e}",
                    "ERROR"
                )

        #! ================= LIMPIAR XLS =================
        for folder in [
            MOTHERBOARD_FILES,
            MOTHERBOARD_1_FILES_FOLDER,
            MOTHERBOARD_2_FILES_FOLDER
        ]:
            if not os.path.exists(folder):
                continue
            for f in os.listdir(folder):
                if f.lower().endswith(".xls"):
                    try:
                        os.remove(os.path.join(folder, f))
                    except:
                        pass
        self.log_msg("\n[OK] Proceso completo", "OK")
        self.btn_resultados.config(state="normal")


if __name__ == "__main__":

    root = tk.Tk()
    app = MainboardApp(root)
    root.mainloop()