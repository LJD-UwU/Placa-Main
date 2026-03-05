import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from PIL import Image, ImageTk
import os

from backend.config.sap_login import abrir_sap_y_login


class MainboardApp:

    def __init__(self, root):
        self.root = root
        self.root.title("MBAutomator - Motherboards")
        self.root.geometry("410x360")
        self.root.resizable(False, False)

        #! icono
        try:
            img = Image.open("IMG/logo.png")
            img = img.resize((256,256))
            icon = ImageTk.PhotoImage(img)
            self.root.iconphoto(True,icon)
        except Exception as e:
            print(f"No se pudo cargar el icono: {e}")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Title.TLabel", font=("Segoe UI",14,"bold"))
        ttk.Label(
            root,
            text="Automatización SAP",
            style="Title.TLabel"
        ).pack(pady=(10,0))
        ttk.Label(
            root,
            text="Herramienta de conexión y carga de Excel",
            foreground="gray"
        ).pack(pady=(0,8))
        main = ttk.Frame(root,padding=8)
        main.pack(fill="both",expand=True)
        
        #! Seleccionar Excel
        fila_file = ttk.Frame(main)
        fila_file.pack(fill="x",pady=5)
        self.excel_path = tk.StringVar()
        ttk.Entry(
            fila_file,
            textvariable=self.excel_path
        ).pack(side="left",fill="x",expand=True)
        ttk.Button(
            fila_file,
            text="📂",
            width=3,
            command=self.seleccionar_excel
        ).pack(side="left",padx=4)

       
        #! Botón conectar SAP
        self.btn_conectar = ttk.Button(
            main,
            text="🔌 Conectarse a SAP",
            command=self.conectar_sap
        )
        self.btn_conectar.pack(pady=10)

     
        #! Consola
        frame_log = ttk.LabelFrame(main,text="CONSOLA")
        frame_log.pack(fill="both",expand=True,pady=(6,0))
        self.log = scrolledtext.ScrolledText(
            frame_log,
            height=10,
            font=("Consolas",9)
        )
        self.log.pack(fill="both",expand=True,padx=5,pady=5)
        self.log.config(state="disabled")
        self.log.tag_config("INFO",foreground="blue")
        self.log.tag_config("OK",foreground="green")
        self.log.tag_config("ERROR",foreground="red")

        #! Estado
        self.status = tk.StringVar(value="Estado: Listo")
        ttk.Label(
            root,
            textvariable=self.status,
            anchor="w"
        ).pack(fill="x",side="bottom",padx=6,pady=4)
        self.session = None

    #! LOG
    def log_msg(self,msg,tag="INFO"):
        self.log.config(state="normal")
        self.log.insert(tk.END,msg+"\n",tag)
        self.log.see(tk.END)
        self.log.config(state="disabled")

    #! Seleccionar Excel
    def seleccionar_excel(self):
        file = filedialog.askopenfilename(
            filetypes=[("Excel","*.xlsx")]
        )
        if file:
            self.excel_path.set(file)
            self.log_msg(f"Excel seleccionado: {os.path.basename(file)}","OK")

    #! Conectar SAP
    def conectar_sap(self):

        try:
            self.log_msg("Intentando conectar a SAP...")
            self.session = abrir_sap_y_login()
            self.log_msg("Conectado a SAP correctamente","OK")
            self.status.set("Estado: SAP conectado")
        except Exception as e:
            self.log_msg(f"[ERROR] {e}","ERROR")
            self.status.set("Estado: Error de conexión")


if __name__ == "__main__":
    root = tk.Tk()
    app = MainboardApp(root)
    root.mainloop()