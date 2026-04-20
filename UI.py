import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk
import os, re, time, sys, subprocess
import pandas as pd
import xlwings as xw
from datetime import datetime
import concurrent.futures
import faulthandler
import logging

try:
    import pythoncom
    _PYTHONCOM_AVAILABLE = True
except ImportError:
    _PYTHONCOM_AVAILABLE = False


def _coinit():
    if _PYTHONCOM_AVAILABLE:
        pythoncom.CoInitialize()


def _couninit():
    """Libera COM en el hilo actual. Llamar al final del hilo secundario."""
    if _PYTHONCOM_AVAILABLE:
        pythoncom.CoUninitialize()

from backend.config.credenciales_loader import cargar_credenciales, guardar_credenciales
from backend.modules.procesar_mainboard_P2 import procesar_material_desde_mainboard
from backend.modules.extract_mainboard import extract_descripcion_numbers
from backend.modules.procesar_motherboard_P1 import procesar_number
from backend.utils.clean_excel import limpiar_excel_mainboard
from backend.config.sap_login import abrir_sap_y_login
from backend.Helpers.helper import cargar_archivos_procesados, guardar_archivo_procesado
from backend.modules.cs11 import ejecutar_cs11
from backend.utils.txt_to_xlsx import (
    exportar_bom_a_xls,
    convertir_xls_a_xlsx,
    BASE_BOM_FOLDER,
    MODEL_FILES_FOLDER,
    MAINBOARD_1_FILES_FOLDER,
    MAINBOARD_2_FILES_FOLDER,
)
from backend.modules.procesar_motherboard_P1 import actualizar_excel_mainboard_1
from backend.modules.procesar_mainboard_P2 import actualizar_excel_mainboard_2
from backend.config.sap_config import DESCRIPCIONES, FILTRO

def abrir_excel_con_timeout(ruta: str, timeout: int = 30) -> pd.DataFrame:
    """
    Abre un Excel con xlwings en un executor independiente.
    Lanza TimeoutError si Excel no responde en `timeout` segundos,
    evitando que la app se congele indefinidamente.
    """
    def _abrir():
        _coinit()
        app = xw.App(visible=False)
        try:
            wb = app.books.open(ruta)
            sheet = wb.sheets[0]
            df = sheet.used_range.options(
                pd.DataFrame, index=False, header=True
            ).value
            wb.close()
            return df
        finally:
            app.quit()
            _couninit()

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(_abrir)
        try:
            return future.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            raise TimeoutError(
                f"Excel no respondió en {timeout}s. "
                "Verifica que el archivo no esté bloqueado o cifrado."
            )


class SAPApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("MBAutomator")
        self.root.geometry("570x480")
        self.root.resizable(False, False)

        #! Logo
        try:
            img = Image.open("backend/IMG/logo.png")
            img = img.resize((64, 64))
            icon = ImageTk.PhotoImage(img)
            self.root.iconphoto(True, icon)
        except Exception as e:
            logging.warning(f"No se pudo cargar el icono: {e}")

        #! Animaciones
        self.animando = False
        self.anim_dots = 0
        self.start_time = None

        #! Estilos
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("TProgressbar", thickness=12)

        #! Encabezados
        ttk.Label(root, text="Automatización SAP", style="Title.TLabel").pack(pady=(8, 0))
        ttk.Label(root, text="Procesamiento automático para el BOM", foreground="gray").pack(pady=(0, 8))

        main = ttk.Frame(root, padding=6)
        main.pack(fill="both", expand=True)

        #! Campo selección Excel
        fila_file = ttk.Frame(main)
        fila_file.pack(fill="x", pady=4)
        self.excel_path = tk.StringVar()
        ttk.Entry(fila_file, textvariable=self.excel_path).pack(side="left", fill="x", expand=True)
        ttk.Button(fila_file, text="📂", width=3, command=self.seleccionar_excel).pack(side="left", padx=4)

        #! Barra progreso
        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.pack(fill="x", pady=6)

        #! Botones
        fila_btn = ttk.Frame(main)
        fila_btn.pack(pady=6)

        self.btn_credenciales = ttk.Button(fila_btn, text="🔐 Login SAP", command=self.abrir_credenciales)
        self.btn_credenciales.pack(side="left", padx=4)

        self.btn_procesar = ttk.Button(fila_btn, text="▶ Procesar 1TE", command=self.iniciar, state="disabled")
        self.btn_procesar.pack(side="left", padx=4)

        self.btn_limpiar = ttk.Button(fila_btn, text="🧹 Procesar Archivos", command=self.limpiar_datos)
        self.btn_limpiar.pack(side="left", padx=4)

        self.btn_open = ttk.Button(fila_btn, text="📁 Resultados", command=self.abrir_resultados, state="disabled")
        self.btn_open.pack(side="left", padx=4)

        self.btn_mainboard = ttk.Button(fila_btn, text="🖥 Motherboard", command=self.abrir_app_mainboard, state="disabled")
        self.btn_mainboard.pack(side="left", padx=4)

        #! Consola
        frame_log = ttk.LabelFrame(main, text="CONSOLA")
        frame_log.pack(fill="both", expand=True, pady=(6, 0))
        self.log = scrolledtext.ScrolledText(frame_log, height=10, font=("Consolas", 10))
        self.log.pack(fill="both", expand=True, padx=5, pady=5)
        self.log.config(state="disabled")
        self.log.tag_config("INFO", foreground="blue")
        self.log.tag_config("OK", foreground="green", font=("Consolas", 10, "bold"))
        self.log.tag_config("ERROR", foreground="red", font=("Consolas", 10, "bold"))
        self.log.tag_config("WARNING", foreground="orange", font=("Consolas", 10, "italic"))

        #! Estado
        self.status = tk.StringVar(value="Estado: Listo")
        self.progress_label = ttk.Label(root, textvariable=self.status, anchor="w")
        self.progress_label.pack(fill="x", side="bottom", padx=6, pady=4)

        #! Estado interno
        self.modelos: list = []
        self.idx: int = 0
        self.session = None
        self.df_todos = pd.DataFrame(columns=["Modelo", "Planta", "Number", "Descripcion"])
        self.materiales_procesados_ok: list = []

        self.excel_path.trace_add("write", lambda *args: self.verificar_habilitar_botones())
        self.verificar_habilitar_botones()
        self._crear_tooltips()

    #!  TOOLTIPS 

    def _crear_tooltips(self):
        tooltips = {
            self.btn_mainboard: "Procesar archivo desde las motherboard",
            self.btn_procesar: "Procesar archivo con 1TE",
            self.btn_limpiar: "Limpiar la consola y archivos exportados",
            self.btn_open: "Abrir carpeta de los archivos",
            self.btn_credenciales: "Iniciar sesión para SAP",
        }
        for widget, text in tooltips.items():
            self._add_tooltip(widget, text)

    def _add_tooltip(self, widget, text):
        tooltip = tk.Toplevel(widget)
        tooltip.withdraw()
        tooltip.overrideredirect(True)
        tk.Label(tooltip, text=text, background="white", relief="solid", borderwidth=1).pack()

        def enter(event):
            tooltip.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
            tooltip.deiconify()

        def leave(event):
            tooltip.withdraw()

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    #!  LOGGING THREAD-SAFE 

    def log_msg(self, msg: str, tag: str = "INFO"):
        logging.debug(f"[{tag}] {msg}")
        self.root.after(0, self._log_msg_ui, msg, tag)

    def _log_msg_ui(self, msg: str, tag: str):
        """Escribe en la consola. Solo llamado desde el hilo principal via root.after."""
        self.log.config(state="normal")
        self.log.insert(tk.END, msg + "\n", tag)
        self.log.see(tk.END)
        self.log.config(state="disabled")

    #!  STATUS THREAD-SAFE 

    def set_status(self, msg: str, animar: bool = False):
        """Thread-safe: delega al hilo principal."""
        self.root.after(0, self._set_status_ui, msg, animar)

    def _set_status_ui(self, msg: str, animar: bool):
        """Solo llamado desde el hilo principal."""
        self.animando = False
        self.anim_dots = 0
        if animar:
            self.animando = True
            self.animar_estado(msg)
        else:
            self.status.set(f"Estado: {msg}")

    def animar_estado(self, texto: str):
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

    #!  UI 

    def seleccionar_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*")])
        if f:
            self.excel_path.set(f)

    def abrir_resultados(self):
        path = os.path.abspath(MAINBOARD_2_FILES_FOLDER)
        if os.path.exists(path):
            os.startfile(path)

    def abrir_app_mainboard(self):
        if hasattr(self, "_mainboard_proc") and self._mainboard_proc.poll() is None:
            return
        try:
            self._mainboard_proc = subprocess.Popen(
                [sys.executable, "-m", "backend.UI.motherboard_app"]
            )
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la app:\n{e}")

    #!  VALIDACIÓN DE BOTONES 

    def verificar_habilitar_botones(self):
        cred = cargar_credenciales()
        cred_ok = all(cred.get(k) for k in ["SAP_SYSTEM_NAME", "SAP_USER", "SAP_PASSWORD"])
        excel = self.excel_path.get().strip()

        if not cred_ok:
            self.btn_mainboard.config(state="disabled")
            self.btn_procesar.config(state="disabled")
            self.btn_limpiar.config(state="disabled")
            self.btn_open.config(state="disabled")
            self.status.set("Estado: Ingrese credenciales SAP 🔐")
        else:
            self.btn_mainboard.config(state="normal")
            self.btn_limpiar.config(state="normal")
            self.btn_procesar.config(state="normal" if excel else "disabled")
        self.btn_open.config(state="disabled")

    #!  CREDENCIALES 

    def abrir_credenciales(self):
        cred = cargar_credenciales()
        win = tk.Toplevel(self.root)
        win.title("MBAutomator - Credenciales SAP")
        win.geometry("320x240")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text="Sistema SAP").pack(pady=(12, 0))
        sistema = tk.StringVar(value=cred.get("SAP_SYSTEM_NAME", ""))
        ttk.Entry(win, textvariable=sistema).pack(fill="x", padx=20)

        ttk.Label(win, text="Usuario").pack(pady=(10, 0))
        usuario = tk.StringVar(value=cred.get("SAP_USER", ""))
        ttk.Entry(win, textvariable=usuario).pack(fill="x", padx=20)

        ttk.Label(win, text="Contraseña").pack(pady=(10, 0))
        password = tk.StringVar(value=cred.get("SAP_PASSWORD", ""))
        ttk.Entry(win, textvariable=password, show="*").pack(fill="x", padx=20)

        def guardar():
            if not sistema.get() or not usuario.get() or not password.get():
                messagebox.showwarning("Atención", "Todos los campos son obligatorios")
                return
            guardar_credenciales({
                "SAP_SYSTEM_NAME": sistema.get().strip(),
                "SAP_USER": usuario.get().strip(),
                "SAP_PASSWORD": password.get(),
            })
            messagebox.showinfo("OK", "Credenciales guardadas correctamente")
            win.destroy()
            self.verificar_habilitar_botones()

        ttk.Button(win, text="Guardar", command=guardar).pack(pady=18)

    #!  LIMPIAR / PROCESAR ARCHIVOS 

    def limpiar_datos(self):
        #! Limpiar consola en el hilo principal (estamos en UI thread aquí)
        self.log.config(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.config(state="disabled")
        self.log_msg("[INFO] Limpieza iniciada", "INFO")

        if not self.excel_path.get():
            messagebox.showwarning("Atención", "Selecciona un Excel primero")
            return

        if not self.cargar_excel_datos(ignorar_process=True):
            return

        respuesta = messagebox.askyesno(
            "Procesamiento de Excel",
            "¿Deseas procesar los archivos Excel?"
        )
        if not respuesta:
            return

        #! Deshabilitar botón para evitar doble clic
        self.btn_limpiar.config(state="disabled")

        #! Lanzar el trabajo pesado en un hilo secundario
        threading.Thread(target=self._limpiar_datos_worker, daemon=True).start()

    def _limpiar_datos_worker(self):
        """
        Corre en hilo secundario.
        Toda modificación de UI se hace via root.after().
        COM inicializado aquí para compatibilidad con xlwings en este hilo.
        """
        _coinit()
        try:
            try:
                from backend.utils.clean_excel_p2 import procesar_archivo_principal_mainboard_2
            except ImportError as e:
                self.log_msg(f"[ERROR] No se pudo importar clean_excel_p2: {e}", "ERROR")
                return

            folder = MAINBOARD_2_FILES_FOLDER
            if not os.path.exists(folder):
                self.log_msg(f"[ERROR] La carpeta {folder} no existe", "ERROR")
                return

            carpeta_final = os.path.join(folder, "ARCHIVOS_FINALES")
            os.makedirs(carpeta_final, exist_ok=True)

            archivos = sorted(
                [f for f in os.listdir(folder)
                 if os.path.isfile(os.path.join(folder, f)) and f.lower().endswith(".xlsx")],
                key=lambda x: os.path.getmtime(os.path.join(folder, x)),
            )

            archivos_procesados = cargar_archivos_procesados()

            for i, f in enumerate(archivos):
                if f in archivos_procesados:
                    self.log_msg(f"[INFO] Archivo ya procesado, se omite: {f}", "INFO")
                    continue

                ruta_excel = os.path.join(folder, f)
                salida_excel = os.path.join(carpeta_final, f"MB-BMM-{f}")

                try:
                    self.log_msg(f"[INFO] Procesando archivo: {f}", "OK")
                    internal_model = self.internal_models[i] if i < len(self.internal_models) else ""
                    plantas = self.plantas[i] if i < len(self.plantas) else ""

                    procesar_archivo_principal_mainboard_2(
                        ruta_excel_principal=ruta_excel,
                        ruta_salida_principal=salida_excel,
                        internal_model=internal_model,
                        plantas=plantas,
                        df_no_procesadas=self.df_no_procesadas,
                    )

                    guardar_archivo_procesado(f)
                    self.log_msg(f"[OK] Archivo procesado: {f}", "OK")

                except Exception as e:
                    self.log_msg(f"[ERROR] No se pudo procesar {f}: {e}", "ERROR")

            self.log_msg("[INFO] Todos los archivos nuevos han sido procesados", "INFO")

            #! Mostrar messagebox en el hilo principal
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Limpieza finalizada",
                    "Los archivos han sido limpiados y procesados correctamente 🧹",
                ),
            )

        except Exception as e:
            self.log_msg(f"[ERROR] {str(e)}", "ERROR")
            return False

    #!  CARGA DE EXCEL 

    def cargar_excel_datos(self, ignorar_process: bool = False) -> bool:
        """
        Lee el Excel seleccionado. Si pandas falla, intenta con xlwings
        usando un timeout para no congelarse.
        Puede llamarse desde el hilo principal O desde un hilo secundario.
        """
        try:
            ruta = self.excel_path.get()
            self.log_msg(f"[INFO] Cargando Excel: {os.path.basename(ruta)}", "INFO")

            try:
                df = pd.read_excel(ruta)
            except Exception:
                self.log_msg(
                    "[WARNING] Pandas no pudo leer el archivo. Intentando con xlwings...",
                    "WARNING",
                )
                df = abrir_excel_con_timeout(ruta, timeout=30)

            df.columns = df.columns.str.strip().str.upper()
            columnas_requeridas = [
                "MATERIAL", "PLANT", "ALTBOM",
                "INTERNAL MODEL", "PROCESS", "MAINBOARD PART NUMBER",
            ]
            faltantes = [c for c in columnas_requeridas if c not in df.columns]
            if faltantes:
                raise ValueError(f"Columnas faltantes en el Excel: {faltantes}")

            df["PROCESS"] = df["PROCESS"].apply(
                lambda x: True if str(x).upper() == "TRUE" else False
            )

            df_filtrado = df if ignorar_process else df[df["PROCESS"] == False]

            if not ignorar_process and df_filtrado.empty:
                raise ValueError("No hay filas nuevas para procesar (todas son TRUE)")

            def limpiar_columna(nombre):
                return (
                    df_filtrado[nombre]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .str.replace(r"\.0$", "", regex=True)
                    .tolist()
                )

            self.modelos = limpiar_columna("MATERIAL")
            self.plantas = limpiar_columna("PLANT")
            self.altboms = limpiar_columna("ALTBOM")
            self.internal_models = limpiar_columna("INTERNAL MODEL")
            self.mainboard_numbers = (
                df_filtrado["MAINBOARD PART NUMBER"].astype(str).str.strip().tolist()
            )
            self.df_no_procesadas = df_filtrado

            self.log_msg("[OK] Excel cargado correctamente", "OK")
            return True

        except Exception as e:
            self.log_msg(f"[ERROR] {e}", "ERROR")
            return False

    #!  FLUJO PRINCIPAL 

    def iniciar(self):
        cred = cargar_credenciales()
        if not all(cred.get(k) for k in ["SAP_SYSTEM_NAME", "SAP_USER", "SAP_PASSWORD"]):
            messagebox.showerror("Credenciales incompletas", "Ingrese credenciales SAP antes de continuar.")
            return
        if not self.excel_path.get():
            messagebox.showwarning("Atención", "Selecciona un Excel")
            return

        self.btn_procesar.config(state="disabled")
        self.progress["value"] = 0
        self.df_todos = pd.DataFrame(columns=["Modelo", "Planta", "Number", "Descripcion"])
        self.materiales_procesados_ok = []
        self.start_time = time.time()
        self.actualizar_tiempo()
        self.log_msg("[INFO] Automatización iniciada", "INFO")
        self.set_status("Cargando Excel", animar=True)

        threading.Thread(target=self._flujo_worker, daemon=True).start()

    def _flujo_worker(self):
        """
        Hilo secundario principal del flujo.
        NUNCA toca widgets directamente — usa root.after() para todo.
        CoInitialize/CoUninitialize envuelven TODO el trabajo COM de este hilo.
        """
        _coinit()
        try:
            if not self.cargar_excel_datos():
                self.root.after(0, lambda: self.btn_procesar.config(state="normal"))
                return

            self.set_status("Conectando a SAP", animar=True)
            try:
                self.session = abrir_sap_y_login()
            except Exception as e:
                self.log_msg(f"[ERROR] Login SAP: {e}", "ERROR")
                self.root.after(0, lambda: self.btn_procesar.config(state="normal"))
                return

            self.log_msg("[OK] Conectado a SAP", "OK")
            self.idx = 0
            total = len(self.modelos)

            if total == 0:
                self.log_msg("[ERROR] No hay materiales para procesar", "ERROR")
                self.root.after(0, lambda: self.btn_procesar.config(state="normal"))
                return

            for self.idx in range(total):
                self._procesar_modelo_sync(self.idx, total)
            self.log_msg("[INFO] Iniciando procesamiento de motherboards y mainboards", "INFO")
            self._guardar_excel_final_sync()
            self._actualizar_process_excel()
            self.root.after(0, self._on_flujo_completado)

        except Exception as e:
            self.log_msg(f"[ERROR] Error fatal en el flujo: {e}", "ERROR")
            logging.exception("Error fatal en _flujo_worker")
            self.root.after(0, lambda: self.btn_procesar.config(state="normal"))
        finally:
            _couninit()

    def _procesar_modelo_sync(self, idx: int, total: int):
        """Procesa un modelo. Corre en hilo secundario."""
        modelo = self.modelos[idx]
        internal_models = self.internal_models[idx]
        self.set_status(f"Modelo {idx + 1}/{total}")
        self.log_msg(f"\n▶ Modelo {idx + 1}/{total}: {modelo}", "OK")

        try:
            self.log_msg("  • Ejecutando CS11...", "INFO")

            resultados = ejecutar_cs11(
                self.session,
                material=modelo,
                uso=FILTRO,
                altboms=self.altboms[idx],
                plantas=self.plantas[idx],
            )

            if not resultados:
                self.log_msg(f"[INFO] No se encontraron plantas para {modelo}", "INFO")

            for planta, _ in resultados:
                self.log_msg(f"  • Planta {planta}: exportando BOM", "INFO")

                ruta_xls = exportar_bom_a_xls(self.session, modelo, mainboard=False)
                self.log_msg("    ✓ BOM exportado", "OK")

                fecha = datetime.now().strftime("%Y-%m-%d")
                hora = datetime.now().strftime("%H-%M-%S")
                nombre_base = os.path.splitext(os.path.basename(ruta_xls))[0]
                nombre_base = re.sub(r'[\\/*?:"<>|]', "_", nombre_base)
                nombre_base = re.sub(r'^(?:\d+-)+', '', nombre_base)
                altboms = self.altboms[idx]

                ruta_xlsx = os.path.join(
                    MODEL_FILES_FOLDER,
                    f"{fecha}_{hora}_{nombre_base}_ALT{altboms}.xlsx",
                )

                if os.path.exists(ruta_xlsx):
                    os.remove(ruta_xlsx)

                convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
                self.log_msg("    ✓ Convertido a XLSX", "OK")
                self.log_msg("    • Buscando motherboard", "OK")
                self.log_msg("    ✓ Motherboard encontrada", "OK")

                df_modelo = extract_descripcion_numbers(ruta_xlsx, internal_models, DESCRIPCIONES)

                if df_modelo.empty:
                    self.log_msg(f"[ERROR] No se encontró motherboard para {modelo}", "ERROR")
                else:
                    df_modelo["Modelo"] = modelo
                    df_modelo["Planta"] = planta
                    self.df_todos = pd.concat([self.df_todos, df_modelo], ignore_index=True)

        except Exception as e:
            self.log_msg(f"[ERROR] Modelo {modelo}: {e}", "ERROR")
            logging.exception(f"Error procesando modelo {modelo}")
        else:
            self.materiales_procesados_ok.append(modelo)

        #! Actualizar barra de progreso
        pct = int(((idx + 1) / total) * 80)  #! 80% para modelos, 20% para mainboards
        self.root.after(0, lambda p=pct: self.progress.configure(value=p))

    def _on_flujo_completado(self):
        """Llamado desde root.after — corre en hilo principal."""
        self.set_status("Finalizado ✅")
        self.progress["value"] = 100
        self.btn_open.config(state="normal")
        self.btn_procesar.config(state="normal")
        messagebox.showinfo(
            "Proceso finalizado",
            "El procesamiento de los 1TE ha terminado correctamente ✅\n"
            "El Excel se ha actualizado correctamente ✅",
        )

    #!  GUARDAR EXCEL FINAL 

    def _guardar_excel_final_sync(self):
        """
        Versión sincrónica de guardar_excel_final.
        Corre íntegramente en el hilo secundario — no bloquea la UI.
        """
        self.set_status("Procesando mainboards")

        for folder in [BASE_BOM_FOLDER, MODEL_FILES_FOLDER,
                       MAINBOARD_1_FILES_FOLDER, MAINBOARD_2_FILES_FOLDER]:
            os.makedirs(folder, exist_ok=True)

        for _, row in self.df_todos.iterrows():
            number = str(row["Number"]).strip().replace(".0", "")
            modelo = str(row["Modelo"]).strip()
            descripcion_mother = str(row["Descripcion"]).strip()

            if any(number in f for f in os.listdir(MAINBOARD_1_FILES_FOLDER)):
                continue

            try:
                ruta_xls = None
                for planta in set(self.plantas):
                    ruta_xls = procesar_number(
                        session=self.session,
                        number=number,
                        planta=planta,
                        uso=FILTRO,
                    )
                    if ruta_xls:
                        break

                if not ruta_xls:
                    continue

                ruta_xlsx = os.path.join(
                    MAINBOARD_1_FILES_FOLDER,
                    re.sub(r'[\\/*?:"<>|]', "_",
                           os.path.basename(ruta_xls).rsplit(".", 1)[0]) + ".xlsx",
                )

                convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
                limpiar_excel_mainboard(ruta_xlsx)

                try:
                    actualizar_excel_mainboard_1(
                        self.excel_path.get(),
                        modelo,
                        [number],
                        descripcion=descripcion_mother,
                    )
                except Exception as e:
                    self.log_msg(f"[ERROR] No se pudo actualizar Excel mainboard 1: {e}", "ERROR")

                materiales_detectados = []
                descripcion_mainboard = ""

                for planta in set(self.plantas):
                    res = procesar_material_desde_mainboard(
                        session=self.session,
                        ruta_mainboard_xlsx=ruta_xlsx,
                        uso=FILTRO,
                        planta=planta,
                    )
                    if res:
                        material, desc = res
                        materiales_detectados.append(str(material).strip())
                        if desc:
                            descripcion_mainboard = desc

                materiales_detectados = list(set(materiales_detectados))
                logging.debug(f"Materiales detectados para {number}: {materiales_detectados}")

                if materiales_detectados:
                    try:
                        actualizar_excel_mainboard_2(
                            self.excel_path.get(),
                            modelo,
                            materiales_detectados,
                            descripcion=descripcion_mainboard,
                        )
                    except Exception as e:
                        self.log_msg(f"[ERROR] No se pudo actualizar Excel mainboard 2: {e}", "ERROR")

            except Exception as e:
                self.log_msg(f"[ERROR] Mainboard {number}: {e}", "ERROR")
                logging.exception(f"Error en mainboard {number}")

            #! Limpiar .xls temporales
            for folder in [MAINBOARD_1_FILES_FOLDER, MAINBOARD_2_FILES_FOLDER, MODEL_FILES_FOLDER]:
                for f in os.listdir(folder):
                    ruta = os.path.join(folder, f)
                    if os.path.isfile(ruta) and f.lower().endswith(".xls"):
                        try:
                            os.remove(ruta)
                        except Exception as e:
                            self.log_msg(f"[ERROR] No se pudo eliminar {f}: {e}", "ERROR")

    #! ACTUALIZAR COLUMNA PROCESS 

    def _actualizar_process_excel(self):
        """Actualiza la columna PROCESS en el Excel original. Corre en hilo secundario."""
        try:
            ruta_excel = self.excel_path.get()
            self.log_msg("[INFO] Actualizando columna PROCESS en el Excel original...", "INFO")

            app = xw.App(visible=False)
            try:
                wb = app.books.open(ruta_excel)
                sheet = wb.sheets.active
                data = sheet.used_range.value

                if not data:
                    raise ValueError("El archivo Excel está vacío")

                header = [str(h).strip().upper() for h in data[0]]
                try:
                    col_process = header.index("PROCESS") + 1
                    col_material = header.index("MATERIAL") + 1
                except ValueError:
                    raise ValueError("No se encontraron las columnas 'MATERIAL' o 'PROCESS'")

                materiales_procesados = {str(m).strip() for m in self.materiales_procesados_ok}

                for i, row in enumerate(data[1:], start=2):
                    material = str(row[col_material - 1]).strip()
                    if material in materiales_procesados:
                        sheet.cells(i, col_process).value = True

                wb.save()
                wb.close()
            finally:
                app.quit()

            self.log_msg("[OK] Excel actualizado", "OK")

        except Exception as e:
            self.log_msg(f"[ERROR] No se pudo actualizar el Excel: {e}", "ERROR")
            logging.exception("Error actualizando columna PROCESS")


    def guardar_excel_final(self):
        """Alias público — delega a la versión sync (debe llamarse desde hilo secundario)."""
        self._guardar_excel_final_sync()


if __name__ == "__main__":
    root = tk.Tk()
    app = SAPApp(root)

    cred = cargar_credenciales()
    if not all(cred.get(k) for k in ["SAP_SYSTEM_NAME", "SAP_USER", "SAP_PASSWORD"]):
        messagebox.showinfo(
            "Atención",
            "No se han ingresado credenciales para SAP.\n"
            "Ve a 🔐 Login SAP para habilitar cualquier proceso.",
        )
    root.mainloop()
    faulthandler.cancel_dump_traceback_later()