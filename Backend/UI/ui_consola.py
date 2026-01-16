import tkinter as tk
from tkinter.scrolledtext import ScrolledText

class ConsoleWindow:
    def __init__(self, parent, console_writer):
        self.parent = parent
        self.console_writer = console_writer
        self.window = None
        self.text_widget = None

    def toggle(self):
        if self.window and tk.Toplevel.winfo_exists(self.window):
            self.close()
            return
        
        self.window = tk.Toplevel(self.parent)
        self.window.title("SAP Bot Console")
        self.window.geometry("600x400")
        self.window.configure(bg="#1e1e1e")

        self.text_widget = ScrolledText(
            self.window, bg="#1e1e1e", fg="#d4d4d4", 
            insertbackground="white", font=("Consolas", 10),
            wrap=tk.WORD, state=tk.DISABLED
        )
        self.text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Configurar colores para etiquetas
        self.text_widget.tag_configure("INFO", foreground="#4CAF50")
        self.text_widget.tag_configure("ERROR", foreground="#f44336")
        self.text_widget.tag_configure("WARNING", foreground="#FF9800")

        self.console_writer.add_target(self.text_widget)
        print("[INFO] Consola de monitoreo activa.")
        
        self.window.protocol("WM_DELETE_WINDOW", self.close)

    def close(self):
        if self.text_widget:
            self.console_writer.remove_target(self.text_widget)
        if self.window:
            self.window.destroy()
        self.window = None
        self.text_widget = None
