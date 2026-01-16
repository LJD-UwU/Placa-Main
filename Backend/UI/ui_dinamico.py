import tkinter as tk

class DynamicFooter:
    def __init__(self, parent, config_cb, update_cb):
        # Cambiamos side=tk.BOTTOM por side=tk.TOP
        self.frame = tk.Frame(parent, bg="#000000")
        self.frame.pack(side=tk.TOP, fill=tk.X, pady=(5, 5)) 

        # Botón para guardar configuración
        self.btn_config = tk.Button(
            self.frame, text="💾 Guardar Cambios en JSON", 
            bg="#3f51b5", fg="white", font=("Segoe UI", 9, "bold"),
            height=2, cursor="hand2", command=config_cb
        )
        
        # Botón de actualización
        self.btn_update = tk.Button(
            self.frame, text="🚀 Actualización Disponible - Instalar", 
            bg="#FF9800", fg="black", font=("Segoe UI", 10, "bold"),
            height=2, cursor="hand2", command=update_cb
        )
        
        self.btn_config.pack(fill=tk.X)

    def switch_mode(self, hay_update: bool):
        if hay_update:
            self.btn_config.pack_forget()
            self.btn_update.pack(fill=tk.X)
        else:
            self.btn_update.pack_forget()
            self.btn_config.pack(fill=tk.X)
