import sys
import tkinter as tk

class ConsoleWriter:
    def __init__(self, original_stdout):
        self.original_stdout = original_stdout
        self.targets = []

    def add_target(self, widget):
        if widget not in self.targets:
            self.targets.append(widget)

    def remove_target(self, widget):
        if widget in self.targets:
            self.targets.remove(widget)

    def write(self, message):
        # Escribir en la terminal real 
        self.original_stdout.write(message)
        
        # Escribir en todos los widgets de la interfaz
        for target in self.targets:
            try:
                target.config(state=tk.NORMAL)
                
                # Detectar etiquetas de color
                tag = None
                if "[INFO]" in message: tag = "INFO"
                elif "[ERROR]" in message: tag = "ERROR"
                elif "[WARNING]" in message: tag = "WARNING"
                
                target.insert(tk.END, message, tag)
                target.see(tk.END)
                target.config(state=tk.DISABLED)
            except:
                pass

    def flush(self):
        self.original_stdout.flush()
