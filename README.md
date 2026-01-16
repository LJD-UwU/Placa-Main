# 📑 Procesamiento de BOM (Mainboard - Nuevos Modelos)

Este documento describe el procedimiento estandarizado para la extracción, limpieza y formateo de la Lista de Materiales (BOM) desde **SAP HQ** hasta su entrega final en Excel.

---

## 🛠 1. Extracción en SAP

1.  **Acceso:** Iniciar sesión en el sistema SAP y acceder al módulo **HQ**.
2.  **Transacción:** Ejecutar el comando `/NCS11`.
3.  **Búsqueda de Material:**
    *   Clic en el icono de búsqueda del campo **Material**.
    *   Filtrar el modelo interno usando asteriscos (Ej: `*MODELO*`) en el primer recuadro.
    *   Filtrar por patrón `1TE*`.
    *   Doble clic en el material correspondiente.
4.  **Parámetros de Planta:**
    *   **Planta:** `2000` o `2900`.
    *   **BOM Application:** `PP01`.
    *   Presionar el icono del **Reloj (Ejecutar)**.

## 🔍 2. Identificación del Ensamble

1.  Dentro del listado, localizar el símbolo chino **"主"**.
2.  Copiar el **ID** numérico asociado a ese símbolo.
3.  Regresar al inicio de la transacción y colocar este ID en el campo de material principal.
4.  Verificar que los datos de Planta y BOM sigan correctos y ejecutar nuevamente.

## 📤 3. Exportación y Codificación

1.  **Formato:** Seleccionar la vista `MB VERIFY`.
2.  **Exportar:** Elegir la segunda opción de exportación, asignar nombre y extensión `.XLS`.
3.  **Apertura en Excel:**
    *   Al abrir, confirmar "Sí" al mensaje de advertencia.
    *   **IMPORTANTE:** En el origen del archivo (*File Origin*), seleccionar:  
      `936 : Chinese Simplified (GB2312)` para evitar errores de lectura.
4.  **Guardado:** Cambiar inmediatamente el formato de `.XLS` a **`.xlsx`**.

---

## 🧹 4. Limpieza y Reestructuración

### Ajuste de Celdas
*   Eliminar la **Columna A**.
*   Eliminar las **filas 1 a la 9**.
*   **Columnas de Nivel:**
    *   Insertar 3 columnas tras los encabezados originales.
    *   *Nivel 0:* Llenar con `x` -> `3TE`.
    *   *Nivel 1:* Llenar con `x` -> ID Ensamble Mainboard.
*   **Reorganización:**
    *   Agregar 2 columnas vacías a partir de la Columna E.
    *   Mover columnas **J y K** (descripciones) después de la descripción en inglés.
    *   Borrar todo después de la columna `项目文本行 2`.

### Encabezados Estándar
Renombrar las columnas de la siguiente manera:

---
## 📂 Estructura del Proyecto
```text
📦 Proyecto SAP
┣ 📂 Backend            
┃ ┣ 📂 Json                # Archivos de configuración
┃ ┃ ┃
┃ ┃ ┗ 📜 Primer-pass.json
┃ ┣ 📂 Settings            # Configuraciones globales y constantes
┃ ┃ ┃
┃ ┃ ┗ 📜 Rutas.py
┃ ┣ 📂 UI                  # Capa de presentación y lógica de interfaz
┃ ┃ ┣ 📜 App_Logic.py
┃ ┃ ┣ 📜 UI_App.py
┃ ┃ ┣ 📜 ui_console.py
┃ ┃ ┗ 📜 ui_dinamico.py
┃ ┗ 📂 Utils               # Herramientas de soporte y ejecución
┃   ┣ 📜 console_writer.py
┃   ┗ 📜 Executor.py
┣ 📂 Data  
┣ 📜 .gitignore 
┣ 📜 LICENSE  
┣ 📜 Main_SAP.py 
┗ 📜 README.md             






