import sys
import os
import tkinter as tk

# Agregar el directorio raíz al PATH
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

try:
    # Formato: from Carpeta.Subcarpeta.Archivo import Clase
    from Backend.UI.UI_App import BOM_Interface
except ImportError as e:
    print(f"❌ Error de ruta: {e}")
    sys.exit(1)

if __name__ == "__main__":
    root = tk.Tk()
    app = BOM_Interface(root)
    root.mainloop()
