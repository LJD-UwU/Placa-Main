import json
import time
import os
from pynput import mouse, keyboard
from pynput.mouse import Button, Controller as MouseController
from pynput.keyboard import Key, Controller as KeyboardController

mouse_ctrl = MouseController()
teclado_ctrl = KeyboardController()
corriendo = True

def monitorear_parada(key):
    global corriendo
    if key == keyboard.Key.f10:
        print("\n🛑 PARADA DE EMERGENCIA (F10)")
        corriendo = False
        return False

def ejecutar_json(archivo_o_lista):
    global corriendo
    corriendo = True
    
    # Si recibimos una ruta única, la convertimos en lista para el loop
    lista_rutas = [archivo_o_lista] if isinstance(archivo_o_lista, str) else archivo_o_lista
    
    listener = keyboard.Listener(on_press=monitorear_parada)
    listener.start()

    try:
        for ruta in lista_rutas:
            if not corriendo: break
            
            print(f"🚀 Cargando secuencia: {os.path.basename(ruta)}")
            with open(ruta, 'r', encoding='utf-8') as f:
                pasos = json.load(f)

            pasos.sort(key=lambda x: x.get('delay', 0))
            tiempo_inicio = time.time()

            for paso in pasos:
                if not corriendo: break

                tiempo_objetivo = paso.get('delay', 0)
                while (time.time() - tiempo_inicio) < tiempo_objetivo:
                    if not corriendo: break
                    time.sleep(0.01)

                if not corriendo: break
                tipo = paso['tipo']
                
                if tipo in ['click', 'double_click']:
                    mouse_ctrl.position = (paso['x'], paso['y'])
                    boton_str = paso.get('boton', 'Button.left').split('.')[-1].lower()
                    boton = Button.left if boton_str == 'left' else Button.right
                    num_clicks = 2 if tipo == 'double_click' else 1
                    mouse_ctrl.click(boton, num_clicks)
                    print(f"🖱️ {tipo}: ({paso['x']}, {paso['y']})")

                elif tipo == 'escritura':
                    teclado_ctrl.type(str(paso['valor']))
                    print(f"📝 Escribiendo: {paso['valor']}")

                elif tipo == 'tecla_especial':
                    nombre_tecla = paso['valor'].split('.')[-1].lower()
                    tecla_obj = getattr(Key, nombre_tecla)
                    teclado_ctrl.press(tecla_obj)
                    teclado_ctrl.release(tecla_obj)
                    print(f"⌨️ Tecla: {nombre_tecla}")

    except Exception as e:
        print(f"❌ Error crítico en ejecución: {e}")
    finally:
        listener.stop()
        print("\n🏁 Proceso finalizado.")
