# Compresor Recursivo de Subcarpetas - PDFs y Archivos de Office

Este programa recorre recursivamente una carpeta especificada, inspeccionando su contenido y todas las subcarpetas para localizar archivos PDF, Word, Excel, etc. Luego realiza:

- ✅ Compresión de archivos PDF usando **Ghostscript**
- ✅ Conversión de archivos Word a PDF usando **LibreOffice**
- ✅ Limpieza de archivos ocultos/macOS
- ✅ Compresión de subcarpetas divididas en partes si superan cierto tamaño
- ✅ Renombrado de estructuras de carpetas repetidas
- ✅ Generación de ZIPs individuales por subcarpeta

## Requisitos para ejecutar

1. **Python 3.11 o 3.13**

   - Descarga: [https://www.python.org/downloads/windows/](https://www.python.org/downloads/windows/)
   - Marcar la opción "Add Python to PATH" al instalar

2. **Ghostscript**

   - Descarga: [https://www.ghostscript.com/download/gsdnld.html](https://www.ghostscript.com/download/gsdnld.html)
   - Durante la instalación asegúrate de agregar al PATH
   - Ejecutable requerido: `gswin64c` (Windows) o `gs` (macOS/Linux)

3. **LibreOffice**

   - Descarga: [https://www.libreoffice.org/download/](https://www.libreoffice.org/download/)
   - Selecciona "Agregar LibreOffice al PATH" en la instalación
   - Ejecutable requerido:

     - Windows: `soffice.exe`
     - macOS: `/Applications/LibreOffice.app/Contents/MacOS/soffice`

## Instalación y configuración del entorno

Puedes usar el script `agregar_path.bat` para configurar el entorno de forma automática:

```bat
@echo off
setlocal EnableDelayedExpansion

:: Detecta y agrega al PATH los siguientes componentes si existen:
:: - LibreOffice
:: - Ghostscript (busca versión instalada)
:: - Python 3.11 y 3.13 (y Scripts)

:: Ejecutar como ADMINISTRADOR
:: Reiniciar la terminal tras ejecutar
```

> El script busca las rutas comúnmente usadas y usa `setx` para agregarlas al PATH si no están presentes.

## Ejecución del programa

### Opcion 1: Desde Python (con interfaz)

1. Instala las dependencias usando:

   ```bash
   pip install -r requirements.txt
   ```

2. Ejecuta el script principal:

   ```bash
   python app.py
   ```

3. Usa la interfaz para:

   - Seleccionar carpeta origen (con subcarpetas)
   - Seleccionar carpeta destino
   - Presionar "Comprimir Subcarpetas"

### Opcion 2: Ejecutable `.exe`

Si ya tienes `CompresorArchivos.exe`, solo haz doble clic para usarlo. No requiere Python instalado.

Para generar el `.exe`:

```bash
pyinstaller --noconsole --onefile app.py
```

## Consideraciones adicionales

- Archivos `.doc` y `.docx` se convierten a PDF antes de ser comprimidos.
- Archivos `.xlsx`, `.xls` se copian tal cual.
- Archivos `.DS_Store`, `__MACOSX` y similares se eliminan.
- Estructuras de carpetas repetidas se renombran con sufijos `_1`, `_2`, etc.

## Verificación manual del entorno

En la terminal (CMD):

```cmd
python --version
soffice --version
gswin64c --version
```

Si alguno de ellos falla, asegúrate de que:

- Esté correctamente instalado
- Su ruta esté en el `PATH`

## Soporte multiplataforma

- ✅ Windows
- ✅ macOS

El script usa `platform.system()` y `shutil.which()` para detectar y adaptar los comandos según el sistema operativo.

---

© Proyecto Compresor Recursivo de Subcarpetas - 2025
