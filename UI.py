import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os, re, time

from backend.modules.extract_mainboard import extract_descripcion_numbers
from backend.utils.historial import registrar_historial_excel
from backend.utils.clean_excel import limpiar_excel_mainboard
from backend.modules.procesar_mainboard_P1 import procesar_number
from backend.modules.prosesar_mainboard_P2 import procesar_material_desde_mainboard
from backend.config.sap_login import abrir_sap_y_login
from backend.config.credenciales_loader import cargar_credenciales, guardar_credenciales
from backend.modules.cs11 import ejecutar_cs11
from backend.utils.txt_to_xlsx import (
    exportar_bom_a_xls,
    convertir_xls_a_xlsx,
    MODEL_FILES_FOLDER,
    MAINBOARD_1_FILES_FOLDER,
    MAINBOARD_2_FILES_FOLDER,
    MODEL_FILES_FOLDER,
    HISTORIAL_FOLDER
)
from backend.config.sap_config import (
    DESCRIPCIONES,
    PLANTAS,
    FILTRO_SAP,
    FILTRO
)

class SAPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MBAutomator")
        self.root.geometry("460x420")
        self.root.resizable(False, False)
        try:
                self.root.iconbitmap(r"IMG/logo.ico") 
        except Exception as e:
                print(f"No se pudo cargar el icono: {e}")

        self.animando = False
        self.anim_dots = 0
        self.start_time = None

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("TProgressbar", thickness=10)

        ttk.Label(root, text="Automatización SAP", style="Title.TLabel").pack(pady=(8, 0))
        ttk.Label(root, text="Procesamiento automático de modelos y mainboards", foreground="gray").pack(pady=(0, 6))

        main = ttk.Frame(root, padding=6)
        main.pack(fill="both", expand=True)

        fila_file = ttk.Frame(main)
        fila_file.pack(fill="x", pady=4)
        self.excel_path = tk.StringVar()
        ttk.Entry(fila_file, textvariable=self.excel_path).pack(side="left", fill="x", expand=True)
        ttk.Button(fila_file, text="📂", width=3, command=self.seleccionar_excel).pack(side="left", padx=4)

        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.pack(fill="x", pady=6)

        fila_btn = ttk.Frame(main)
        fila_btn.pack(pady=4)

        self.btn_procesar = ttk.Button(fila_btn, text="▶ Procesar", command=self.iniciar)
        self.btn_procesar.pack(side="left", padx=4)

        self.btn_open = ttk.Button(fila_btn, text="📁 Resultados", command=self.abrir_resultados, state="disabled")
        self.btn_open.pack(side="left", padx=4)

        ttk.Button(fila_btn, text="🔐 Credenciales", command=self.abrir_credenciales).pack(side="left", padx=4)

        frame_log = ttk.LabelFrame(main, text="CONSOLA")
        frame_log.pack(fill="both", expand=True, pady=(6, 0))
        self.log = scrolledtext.ScrolledText(frame_log, height=9, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=5, pady=5)
        self.log.config(state="disabled")
        self.log.tag_config("INFO", foreground="blue")
        self.log.tag_config("OK", foreground="green")
        self.log.tag_config("ERROR", foreground="red")

        self.status = tk.StringVar(value="Estado: Listo")
        ttk.Label(root, textvariable=self.status, anchor="w").pack(fill="x", side="bottom", padx=6, pady=4)

        self.modelos = []
        self.idx = 0
        self.session = None
        self.df_todos = pd.DataFrame(columns=["Modelo", "Planta", "Number", "Descripcion"])

    # ================= CREDENCIALES =================
    def abrir_credenciales(self):
        cred = cargar_credenciales()

        win = tk.Toplevel(self.root)
        win.title("Credenciales SAP")
        win.geometry("320x240")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        # --- Cambiar el icono de la ventana de credenciales ---
        try:
            win.iconbitmap(r"IMG/logo.ico")  # reemplaza el icono por tu logo
        except Exception as e:
            print(f"No se pudo cambiar el icono de la ventana de credenciales: {e}")

        # --- Campos de credenciales ---
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
                "SAP_PASSWORD": password.get()
            })

            messagebox.showinfo("OK", "Credenciales guardadas correctamente")
            win.destroy()

        ttk.Button(win, text="Guardar", command=guardar).pack(pady=18)

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

    # ================= FLUJO =================
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

        try:
            # Intentar abrir SAP
            self.session = abrir_sap_y_login()

        except Exception as e:
            msg = str(e)

            # Detectar error de credenciales incompletas
            if "Credenciales SAP incompletas" in msg:
                self.log_msg(f"[ERROR] {msg}", "ERROR")
                messagebox.showerror(
                    "Error SAP",
                    "Falló la apertura o login de SAP:\nCredenciales SAP incompletas.\nPor favor revisa tus credenciales."
                )
                # Habilitar botón para reintentar
                self.btn_procesar.config(state="normal")
                self.animando = False
                self.session = None
                return

            # Otros errores de SAP (opcional)
            self.log_msg(f"[ERROR] {msg}", "ERROR")
            messagebox.showerror(
                "Error SAP",
                f"No se pudo conectar a SAP:\n{msg}"
            )
            self.btn_procesar.config(state="normal")
            self.animando = False
            self.session = None
            return

        # Si todo salió bien
        self.animando = False
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
                componente=FILTRO_SAP,
                uso=FILTRO,
                plantas=PLANTAS
            )

            registrar_historial_excel(
                archivo=modelo,
                proceso="Modelo",
                paso="CS11",
                estado="OK" if resultados else "INFO",
                detalle=f"Plantas encontradas: {len(resultados)}" if resultados else "Sin resultados"
            )

            for planta, _ in resultados:
                self.log_msg(f"  • Planta {planta}: exportando BOM", "INFO")
                ruta_xls = exportar_bom_a_xls(self.session, modelo, mainboard=False)
                self.log_msg("    ✓ BOM exportado", "OK")
                registrar_historial_excel(
                    archivo=os.path.basename(ruta_xls),
                    proceso="Modelo",
                    paso="Exportar BOM",
                    estado="OK",
                    detalle=f"Planta {planta}"
                )

                ruta_xlsx = os.path.join(MODEL_FILES_FOLDER, re.sub(r'[\\/*?:"<>|]', "_", os.path.basename(ruta_xls).replace(".XLS","")) + ".xlsx")
                convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
                self.log_msg("    ✓ Convertido a XLSX", "OK")
                registrar_historial_excel(
                    archivo=os.path.basename(ruta_xlsx),
                    proceso="Modelo",
                    paso="Conversión XLSX",
                    estado="OK",
                    detalle="Codificación preservada"
                )

                self.log_msg("    • Analizando descripciones", "INFO")
                df_modelo = extract_descripcion_numbers(ruta_xlsx, modelo, DESCRIPCIONES)
                if not df_modelo.empty:
                    df_modelo["Modelo"] = modelo
                    df_modelo["Planta"] = planta
                    self.df_todos = pd.concat([self.df_todos, df_modelo], ignore_index=True)
                    registrar_historial_excel(
                        archivo=modelo,
                        proceso="Modelo",
                        paso="Buscar Mainboards",
                        estado="OK",
                        detalle=f"Mainboards encontrados: {len(df_modelo)}"
                    )

        except Exception as e:
            self.log_msg(f"[ERROR] {e}", "ERROR")
            registrar_historial_excel(
                archivo=modelo,
                proceso="Modelo",
                paso="Error general",
                estado="ERROR",
                detalle=str(e)
            )

        self.idx += 1
        self.root.after(200, self.procesar_modelo)

    def guardar_excel_final(self):
        self.set_status("Procesando mainboards")

        # Crear carpetas si no existen
        for folder in [MODEL_FILES_FOLDER, MAINBOARD_1_FILES_FOLDER, MAINBOARD_2_FILES_FOLDER]:
            os.makedirs(folder, exist_ok=True)

        # Procesamiento normal de mainboards
        for _, row in self.df_todos.iterrows():
            number = str(row["Number"]).strip()
            if any(number in f for f in os.listdir(MAINBOARD_1_FILES_FOLDER)):
                continue
            try:
                ruta_xls = procesar_number(self.session, number, "2000", FILTRO)
                ruta_xlsx = os.path.join(MAINBOARD_1_FILES_FOLDER, re.sub(r'[\\/*?:"<>|]', "_", os.path.basename(ruta_xls).replace(".XLS","")) + ".xlsx")
                convertir_xls_a_xlsx(ruta_xls, ruta_xlsx)
                limpiar_excel_mainboard(ruta_xlsx)

                materiales_detectados = procesar_material_desde_mainboard(self.session, ruta_xlsx, FILTRO)
                if materiales_detectados:
                    registrar_historial_excel(
                        archivo=number,
                        proceso="Mainboard",
                        paso="Material detectado",
                        estado="OK",
                        detalle=f"Materiales: {', '.join(materiales_detectados)}"
                    )

                registrar_historial_excel(
                    archivo=number,
                    proceso="Mainboard",
                    paso="Procesamiento completo",
                    estado="OK",
                    detalle="Mainboard exportado y analizado"
                )

                registrar_historial_excel(
                    archivo=os.path.basename(ruta_xlsx),
                    proceso="Exportación final SAP",
                    paso="Exportación XLSX",
                    estado="OK",
                    detalle="Archivo final generado desde SAP"
                )

            except Exception as e:
                self.log_msg(f"[ERROR] Mainboard {number}: {e}", "ERROR")
                registrar_historial_excel(
                    archivo=number,
                    proceso="Mainboard",
                    paso="Error",
                    estado="ERROR",
                    detalle=str(e)
                )
                
        for folder in [HISTORIAL_FOLDER, MAINBOARD_1_FILES_FOLDER, MAINBOARD_2_FILES_FOLDER,MODEL_FILES_FOLDER]:
            for f in os.listdir(folder):
                ruta = os.path.join(folder, f)
                # Comprobar si es un archivo .xls (mayúscula o minúscula)
                if os.path.isfile(ruta) and f.lower().endswith(".xls"):
                    try:
                        os.remove(ruta)
                        
                    except Exception as e:
                        self.log_msg(f"[ERROR] No se pudo eliminar {f}: {e}", "ERROR")

if __name__ == "__main__":
    root = tk.Tk()
    app = SAPApp(root)
    root.mainloop()
