import os
import time
import pandas as pd
import tkinter as tk
import xlwings as xw
from datetime import datetime
from PIL import Image, ImageTk
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from tkinter import ttk, filedialog, scrolledtext, messagebox

from backend.config.sap_config import FILTRO
from backend.config.sap_login import abrir_sap_y_login
from backend.utils.clean_excel import limpiar_excel_mainboard
from backend.utils.utils_2.xlsx_m2 import convertir_xls_a_xlsx
from backend.modules.Modules_2.procesar_mainboard import actualizar_excel_mainboard
from backend.modules.Modules_2.procesar_motherboard import procesar_numbers_desde_listas
from backend.utils.txt_to_xlsx import (MAINBOARD_1_FILES_FOLDER,MAINBOARD_2_FILES_FOLDER)
from backend.modules.Modules_2.procesar_mainboard import procesar_material_desde_mainboard


def limpiar_valor(valor):
    if valor is None:
        return ""
    valor = str(valor).strip()
    if valor.endswith(".0"):
        valor = valor[:-2]
    return valor

def abrir_excel_seguro(path_excel):
    """
    Abre Excel con fallback:
    1. openpyxl (rápido)
    2. xlwings (soporta cifrado)
    """
    try:
        wb = load_workbook(path_excel)
        return wb, "openpyxl"
    except Exception:
        # fallback a xlwings
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False

        try:
            wb_xw = app.books.open(path_excel)
            return (app, wb_xw), "xlwings"
        except Exception as e:
            app.quit()
            raise Exception(f"No se pudo abrir el archivo (ni openpyxl ni xlwings): {e}")

#!  FILTRO POR COLOR 
def es_amarillo(celda):
    if not celda.fill:
        return False
    fg = celda.fill.fgColor
    try:
        if fg.rgb:
            rgb = str(fg.rgb)
            if rgb.upper().endswith("FFFF00"):
                return True
        if fg.indexed is not None and fg.indexed == 6:
            return True
    except Exception:
        pass
    return False

def leer_filas_amarillas(path_excel):

    obj, engine = abrir_excel_seguro(path_excel)

    mothers, plants, internals = [], [], []

    try:
        # 🔹 OPENPYXL (rápido)
        if engine == "openpyxl":
            wb = obj
            ws = wb.active

            headers = {}
            for col in range(1, ws.max_column + 1):
                nombre = ws.cell(row=1, column=col).value
                if nombre:
                    headers[str(nombre).strip().upper()] = col

            required = ["MOTHERBOARD PART NUMBER", "PLANT", "INTERNAL MODEL"]
            for r in required:
                if r not in headers:
                    raise Exception(f"No se encontró la columna: {r}")

            col_mother = headers["MOTHERBOARD PART NUMBER"]
            col_plant = headers["PLANT"]
            col_internal = headers["INTERNAL MODEL"]

            for row in range(2, ws.max_row + 1):
                celda = ws.cell(row=row, column=col_mother)

                if es_amarillo(celda) and celda.value:
                    mothers.append(limpiar_valor(celda.value))
                    plants.append(limpiar_valor(ws.cell(row=row, column=col_plant).value))
                    internals.append(limpiar_valor(ws.cell(row=row, column=col_internal).value))

        # 🔹 XLWINGS (cifrado)
        else:
            app, wb = obj
            sheet = wb.sheets[0]

            data = sheet.used_range.value
            if not data:
                raise Exception("Excel vacío")

            headers = [str(h).strip().upper() for h in data[0]]

            def col_idx(name):
                if name not in headers:
                    raise Exception(f"No se encontró la columna: {name}")
                return headers.index(name)

            col_mother = col_idx("MOTHERBOARD PART NUMBER")
            col_plant = col_idx("PLANT")
            col_internal = col_idx("INTERNAL MODEL")

            # Leer colores con xlwings
            for i in range(2, len(data) + 1):
                celda = sheet.cells(i, col_mother + 1)

                color = celda.color  # RGB tuple

                es_amarillo_xw = color and (
                    (color[0] > 200 and color[1] > 200 and color[2] < 100)
                )

                if es_amarillo_xw:
                    row = data[i - 1]

                    mother = limpiar_valor(row[col_mother])
                    plant = limpiar_valor(row[col_plant])
                    internal = limpiar_valor(row[col_internal])

                    if mother:
                        mothers.append(mother)
                        plants.append(plant)
                        internals.append(internal)

    finally:
        #  limpieza segura
        if engine == "openpyxl":
            obj.close()
        else:
            app, wb = obj
            try:
                wb.close()
            except:
                pass
            app.quit()

    return mothers, plants, internals

#!  MARCAR MOTHERBOARD PROCESADA 
def marcar_procesado(path_excel, mother_name):
    try:
        app = xw.App(visible=False)
        wb = app.books.open(path_excel)
        sheet = wb.sheets[0]

        data = sheet.used_range.value
        headers = [str(h).strip().upper() for h in data[0]]

        col_mother = headers.index("MOTHERBOARD PART NUMBER")

        for i in range(2, len(data) + 1):
            celda = sheet.cells(i, col_mother + 1)

            if limpiar_valor(celda.value) == limpiar_valor(mother_name):
                
                #! VERDE CORRECTO
                celda.color = (198, 224, 180)
                break

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        print(f"[ERROR] {e}")

#!  ELIMINAR ARCHIVOS XLS 
def eliminar_xls_carpeta(carpeta):
    for f in os.listdir(carpeta):
        if f.lower().endswith(".xls"):
            ruta = os.path.join(carpeta, f)
            try:
                os.remove(ruta)
                print(f"[OK] Eliminado: {ruta}")
            except Exception as e:
                print(f"[ERROR] No se pudo eliminar {ruta}: {e}")

#!  MAINBOARD APP 
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
        ttk.Button(fila_file, text="📂", width=3, command=self.seleccionar_excel).pack(side="left", padx=4)

        fila_btn = ttk.Frame(main)
        fila_btn.pack(pady=4)
        self.btn_procesar = ttk.Button(fila_btn, text="▶ Procesar", command=self.iniciar_procesamiento)
        self.btn_procesar.pack(side="left", padx=4)

        frame_log = ttk.LabelFrame(main, text="CONSOLA")
        frame_log.pack(fill="both", expand=True, pady=(6, 0))
        self.log = scrolledtext.ScrolledText(frame_log, height=10, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=5, pady=5)
        self.log.config(state="disabled")
        self.log.tag_config("INFO", foreground="blue")
        self.log.tag_config("OK", foreground="green")
        self.log.tag_config("ERROR", foreground="red")

        self.status = tk.StringVar(value="Estado: Listo")
        ttk.Label(root, textvariable=self.status, anchor="w").pack(fill="x", side="bottom", padx=6, pady=4)

        self.session = None
        self.mother = []
        self.plants = []
        self.internal_models = []

    def log_msg(self, msg, tag="INFO"):
        self.log.config(state="normal")
        self.log.insert(tk.END, msg + "\n", tag)
        self.log.see(tk.END)
        self.log.config(state="disabled")
        self.root.update()

    def seleccionar_excel(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx")])
        if files:
            self.excel_paths = list(files)
            self.log_msg(f"[OK] {len(files)} archivo seleccionado", "OK")

    def conectar_sap(self):
        if not self.session:
            try:
                self.log_msg("[INFO] Conectando a SAP...")
                self.session = abrir_sap_y_login()
                self.log_msg("[OK] Conectado a SAP", "OK")
            except Exception as e:
                self.log_msg(f"[ERROR] {e}", "ERROR")

    def iniciar_procesamiento(self):
        self.log_msg("[INFO] Automatización iniciada\n")

        if not hasattr(self, "excel_paths"):
            self.log_msg("[ERROR] No hay archivos Excel seleccionados", "ERROR")
            return

        self.conectar_sap()
        if not self.session:
            return

        os.makedirs(MAINBOARD_1_FILES_FOLDER, exist_ok=True)

        for excel_file in self.excel_paths:

            eliminar_xls_carpeta(MAINBOARD_1_FILES_FOLDER)
            eliminar_xls_carpeta(MAINBOARD_2_FILES_FOLDER)

            archivo_nombre = os.path.basename(excel_file)
            self.log_msg(f"▶ Archivo : {archivo_nombre}")
            self.log_msg("  • Leyendo Excel")

            try:
                self.mother, self.plants, self.internal_models = leer_filas_amarillas(excel_file)
                total = len(self.mother)

                if total == 0:
                    self.log_msg("    ⚠ No hay filas amarillas → archivo omitido", "ERROR")
                    continue

                self.log_msg(f"    ✓ {total} Motherboard a procesar\n", "OK")

            except Exception as e:
                self.log_msg(f"    [ERROR] No se pudieron leer filas amarillas: {e}", "ERROR")
                continue

            for idx, (mother, plant) in enumerate(zip(self.mother, self.plants), start=1):

                self.log_msg(f"▶ Motherboard {idx}/{total}: {mother}")
                self.log_msg(f"  • Planta {plant}")

                fecha = datetime.now().strftime("%Y-%m-%d")
                excel_salida = os.path.join(MAINBOARD_1_FILES_FOLDER, f"{fecha}-{mother}.xlsx")

                try:
                    procesar_numbers_desde_listas(
                        session=self.session,
                        mother_list=[mother],
                        plant_list=[plant],
                        excel_output=excel_salida,
                        capid=FILTRO
                    )

                    ruta_xls = os.path.join(MAINBOARD_1_FILES_FOLDER, f"{fecha}-{mother}.XLS")

                    for _ in range(20):
                        if os.path.exists(ruta_xls):
                            break
                        time.sleep(1)
                    else:
                        raise Exception("El XLS no fue generado por SAP")

                    convertir_xls_a_xlsx(ruta_xls, excel_salida)
                    limpiar_excel_mainboard(excel_salida)

                    resultado = procesar_material_desde_mainboard(
                        session=self.session,
                        ruta_mainboard_xlsx=excel_salida,
                        uso=FILTRO,
                        planta=plant,
                        mother=mother
                    )

                    if resultado is None:
                        self.log_msg(f"    [ERROR] No se pudo generar BOM para {mother}\n", "ERROR")
                        continue

                    ruta_xlsx, material_detectado = resultado

                    try:
                        actualizar_excel_mainboard(
                            mother,
                            [material_detectado],
                            ruta_excel=excel_file
                        )
                        self.log_msg(f"    ✓ Excel actualizado con materiales de {mother}", "OK")
                    except Exception as e:
                        self.log_msg(f"    [ERROR] No se pudo actualizar Excel: {e}", "ERROR")

                    marcar_procesado(excel_file, mother)
                    self.log_msg(f"    ✓ BOM generado correctamente\n", "OK")

                except Exception as e:
                    self.log_msg(f"    [ERROR] Error procesando {mother}: {e}\n", "ERROR")

        self.log_msg("[INFO] Todos las motherboard se procesaron")
        self.log_msg("[OK] Proceso completo")

        messagebox.showinfo("Proceso terminado", "El proceso de Motherboard ha finalizado correctamente")

        self.root.update()
        time.sleep(5)

        #! LIMPIEZA FINAL DE XLS
        try:
            self.log_msg("[INFO] Limpiando archivos XLS antes de cerrar...")
            eliminar_xls_carpeta(MAINBOARD_1_FILES_FOLDER)
            eliminar_xls_carpeta(MAINBOARD_2_FILES_FOLDER)
            self.log_msg("[OK] Archivos XLS eliminados", "OK")
        except Exception as e:
            self.log_msg(f"[ERROR] No se pudieron eliminar los XLS: {e}", "ERROR")

        self.root.destroy()

        messagebox.showinfo(
            "Siguiente paso",
            "Cargar archivo procesado en la ventana principal y dar clic en el botón 'Limpiar'"
        )

if __name__ == "__main__":
    root = tk.Tk()
    app = MainboardApp(root)
    root.mainloop()