import sys
import os
import json
import tkinter as tk
from tkinter import messagebox
import threading
import stat

# --- IMPORTACIONES ---
from Backend.UI.ui_consola import ConsoleWindow
from Backend.UI.ui_dinamico import DynamicFooter
from Backend.Utils.console_writer import ConsoleWriter
from Backend.Utils import Executor
from Backend.UI.App_Logic import iniciar_proceso_bom

class BOM_Interface:
    def __init__(self, root):
        self.root = root
        self.root.title("BOM APP")
        self.root.geometry("400x350") 
        self.root.configure(bg="#000000")
        self.root.resizable(False, False)

        self._console_writer = ConsoleWriter(sys.stdout)
        sys.stdout = self._console_writer
        sys.stderr = self._console_writer
        self.console_ui = ConsoleWindow(self.root, self._console_writer)

        main_frame = tk.Frame(root, padx=15, pady=10, bg="#000000")
        main_frame.pack(fill="both", expand=True)

        tk.Label(main_frame, text="Mainboard BOM Automator", 
        font=("Segoe UI", 12, "bold"), bg="#000000", fg="white").pack(pady=(5, 10))

        self.frame_creds = tk.LabelFrame(main_frame, text=" Configuración del sistema SAP ", 
        bg="#000000", fg="#4CAF50", padx=10, pady=10)
        self.frame_creds.pack(fill="x", pady=5)

        self._crear_campo(self.frame_creds, "Usuario:", "ent_user")
        self._crear_campo(self.frame_creds, "Contraseña:", "ent_pass", show="*")
        self._crear_campo(self.frame_creds, "Modelo Interno:", "ent_modelo")

        self._setup_grid_botones(main_frame)

    def _crear_campo(self, parent, label_text, var_name, show=None):
        frame = tk.Frame(parent, bg="#000000")
        frame.pack(fill="x", pady=2)
        tk.Label(frame, text=label_text, bg="#000000", fg="white", width=15, anchor="w").pack(side="left")
        entry = tk.Entry(frame, bg="#333333", fg="white", insertbackground="white", borderwidth=0)
        if show: entry.config(show=show)
        entry.pack(side="right", fill="x", expand=True)
        setattr(self, var_name, entry)

    def _setup_grid_botones(self, parent):
        grid = tk.Frame(parent, bg="#000000")
        grid.pack(pady=0)
        
        btn_style = {"font": ("Segoe UI", 9, "bold"), "width": 18, "height": 2, "cursor": "hand2"}
        
        self.btn_run = tk.Button(grid, text="▶ INICIO", bg="#4CAF50", fg="white", 
                                 command=self.iniciar_hilo, **btn_style)
        self.btn_run.grid(row=0, column=0, padx=5, pady=5)

        self.btn_stop = tk.Button(grid, text="⏹ DETENER (F10)", bg="#f44336", fg="white", 
                                  command=self.detener_bot, **btn_style)
        self.btn_stop.grid(row=0, column=1, padx=5, pady=5)

        self.btn_save = tk.Button(grid, text="💾 GUARDAR EN JSON", bg="#3f51b5", fg="white", 
                                  command=self.guardar_datos, **btn_style)
        self.btn_save.grid(row=1, columnspan=2, sticky="ew", padx=5, pady=5)

        tk.Button(grid, text="📂 VISUALIZAR CONSOLA", bg="#9C27B0", fg="white", 
                  command=self.console_ui.toggle, **btn_style).grid(row=2, columnspan=2, sticky="ew", padx=5, pady=5)

    def guardar_datos(self):
        _, externa = self._obtener_rutas()
        CARPETA_JSON = os.path.join(externa, "Backend", "Json")
        
        RUTA_1 = os.path.join(CARPETA_JSON, "Primer-pass.json")
        RUTA_2 = os.path.join(CARPETA_JSON, "Segundo-pass.json")
        
        # --- CORRECCIÓN AQUÍ: Uso de .strip() para limpiar saltos de línea ---
        u = self.ent_user.get().strip()
        p = self.ent_pass.get().strip()
        m = self.ent_modelo.get().strip()

        if not all([u, p, m]):
            messagebox.showwarning("Campos Vacíos", "Completa los campos.")
            return

        try:
            # --- ACTUALIZAR JSON 1: Credenciales ---
            with open(RUTA_1, 'r', encoding='utf-8') as f:
                data1 = json.load(f)
            
            data1[3]['valor'] = u
            data1[6]['valor'] = p
            
            with open(RUTA_1, 'w', encoding='utf-8') as f:
                json.dump(data1, f, indent=4, ensure_ascii=False)

            # --- ACTUALIZAR JSON 2: Modelo ---
            with open(RUTA_2, 'r', encoding='utf-8') as f:
                data2 = json.load(f)
            
            data2[8]['valor'] = m
            
            with open(RUTA_2, 'w', encoding='utf-8') as f:
                json.dump(data2, f, indent=4, ensure_ascii=False)

            print(f"[INFO] JSONs actualizados: User '{u}' en J1[3], Pass en J1[6] y Modelo '{m}' en J2[8].")
            messagebox.showinfo("Éxito", "Datos inyectados correctamente en las secuencias.")
            
            self.ent_user.delete(0, tk.END)
            self.ent_pass.delete(0, tk.END)
            self.ent_modelo.delete(0, tk.END)

        except FileNotFoundError as e:
            print(f"[ERROR] No se encontró el archivo: {e.filename}")
            messagebox.showerror("Error de Archivo", f"Asegúrate de que {os.path.basename(e.filename)} exista.")
        except IndexError:
            print("[ERROR] Los archivos JSON no tienen el tamaño suficiente para los índices 3, 6 o 8.")
            messagebox.showerror("Error de Estructura", "El JSON no coincide con la cantidad de pasos esperada.")
        except Exception as e:
            print(f"[ERROR] Error inesperado: {e}")

    def iniciar_hilo(self):
        self.btn_run.config(state="disabled", bg="#222222")
        threading.Thread(target=self.ejecutar_proyecto, daemon=True).start()

    def ejecutar_proyecto(self):
        _, externa = self._obtener_rutas()
        j1 = os.path.join(externa, "Backend", "Json", "Primer-pass.json")
        j2 = os.path.join(externa, "Backend", "Json", "Segundo-pass.json")
        try:
            iniciar_proceso_bom([j1, j2])
        finally:
            self.btn_run.config(state="normal", bg="#4CAF50")

    def detener_bot(self):
        Executor.corriendo = False
        print("[WARNING] Solicitud de parada enviada.")

    def _obtener_rutas(self):
        if hasattr(sys, '_MEIPASS'): return sys._MEIPASS, os.path.dirname(sys.executable)
        raiz = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        return raiz, raiz

    def _on_update(self):
        print("[INFO] Buscando actualizaciones...")
