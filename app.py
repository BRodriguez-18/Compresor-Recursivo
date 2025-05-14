import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import zipfile
import shutil
from collections import defaultdict
import platform

# Parámetros
MAX_ZIP_SIZE_MB = 4
EXTENSIONES_VALIDAS = [".pdf", ".doc", ".docx", ".xls", ".xlsx"]

# Función para comprimir un PDF usando Ghostscript
def comprimir_pdf_ghostscript(entrada, salida):
    opciones = ["/screen", "/ebook", "/printer", "/prepress"]

    sistema = platform.system()
    if sistema == "Windows":
        ghostscript_exec = shutil.which("gswin64c") or shutil.which("gswin32c")
    else:
        ghostscript_exec = shutil.which("gs")

    if not ghostscript_exec:
        print("Ghostscript no encontrado. Asegúrate de que esté instalado y en el PATH.")
        return False        

    for opcion in opciones:
        comando = [
            ghostscript_exec,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS={opcion}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={salida}",
            entrada
        ]
        try:
            subprocess.run(comando, check=True)
            print(f"Compresión exitosa con {opcion} para: {entrada}")
            return True  # Éxito con esta opción
        except Exception as e:
            print(f"Error al comprimir {entrada} con {opcion}: {e}")
    return False  # Ninguna opción funcionó

# Función que genera la ruta interna transformada para el ZIP.
# Ahora sólo se antepone new_root y se mantiene la estructura relativa.
def crear_arcname(archivo, origen_carpeta, new_root, dyn_index):
    # Relativa completa del archivo dentro de origen
    # rel = os.path.relpath(archivo, origen_carpeta)
    # Dentro del ZIP, la carpeta raíz incluirá el new_root (con índice) y luego la ruta relativa original
    # return os.path.join(new_root, rel)
    return os.path.relpath(archivo, origen_carpeta)

def revisar_estructuras_repetidas(destino):
    from pathlib import Path
    carpeta_descomprimidos = os.path.join(destino, "DESCOMPRIMIDOS")
    if not os.path.exists(carpeta_descomprimidos):
        print("No se encontró la carpeta DESCOMPRIMIDOS.")
        return

    sufijos_encontrados = defaultdict(list)

    for carpeta in os.listdir(carpeta_descomprimidos):
        if not os.path.isdir(os.path.join(carpeta_descomprimidos, carpeta)):
            continue

        match = re.match(r"(Punto \d+)", carpeta)
        if match:
            sufijo_base = match.group(1)
            ruta = os.path.join(carpeta_descomprimidos, carpeta)
            sufijos_encontrados[sufijo_base].append(ruta)

    coincidencias = 0

    for sufijo, rutas in sufijos_encontrados.items():
        print(f"\n🔎 Revisando grupo con sufijo: {sufijo}")
        rutas_relativas = defaultdict(list)

        for base_ruta in rutas:
            for root, dirs, _ in os.walk(base_ruta):
                for d in dirs:
                    abs_path = os.path.join(root, d)
                    rel_path = os.path.relpath(abs_path, base_ruta)
                    rutas_relativas[rel_path].append((base_ruta, abs_path))  # guardamos contexto

        for subruta, ubicaciones in rutas_relativas.items():
            if len(ubicaciones) > 1:
                coincidencias += 1
                print(f"\n📁 Ruta repetida encontrada: '{subruta}'")

                for idx, (base_ruta, full_path) in enumerate(ubicaciones, start=1):
                    nombre_base = os.path.basename(base_ruta)
                    partes = Path(subruta).parts
                    print(f"{nombre_base}")
                    indent = "    "

                    nueva_ruta = base_ruta

                    for nivel, parte in enumerate(partes):
                        actual = os.path.join(nueva_ruta, parte)
                        nuevo_nombre = f"{parte}_{idx}"
                        nuevo_path = os.path.join(nueva_ruta, nuevo_nombre)

                        try:
                            if os.path.exists(actual):
                                os.rename(actual, nuevo_path)
                                print(f"{indent * (nivel+1)}{nuevo_nombre} {idx}")
                            else:
                                print(f"{indent * (nivel+1)}{parte} {idx} (no encontrado)")
                        except Exception as e:
                            print(f"⚠️ Error al renombrar {actual}: {e}")
                            break

                        nueva_ruta = nuevo_path


    if coincidencias == 0:
        print("\n✅ No se encontraron carpetas repetidas en ningún nivel.")
    else:
        print(f"\n✅ Revisión finalizada. Se encontraron {coincidencias} coincidencias totales.")

def eliminar_archivos_metadata(carpeta):
    for root, dirs, files in os.walk(carpeta):
        for f in files:
            if f.startswith(".DS_Store") or f.startswith("._"):
                try:
                    os.remove(os.path.join(root, f))
                except Exception:
                    pass
        for d in dirs:
            if d == "__MACOSX":
                try:
                    shutil.rmtree(os.path.join(root, d))
                except Exception:
                    pass

def comprimir_descomprimidos(destino):
    from zipfile import ZipFile, ZIP_DEFLATED

    carpeta_objetivo = os.path.join(destino, "DESCOMPRIMIDOS")
    carpeta_zips = os.path.join(destino, "ZIPS_FINAL")
    os.makedirs(carpeta_zips, exist_ok=True)

    for nombre_carpeta in os.listdir(carpeta_objetivo):
        ruta_carpeta = os.path.join(carpeta_objetivo, nombre_carpeta)
        if not os.path.isdir(ruta_carpeta):
            continue

        print(f"📦 Comprimiendo: {nombre_carpeta}")
        eliminar_archivos_metadata(ruta_carpeta)

        ruta_zip = os.path.join(carpeta_zips, f"{nombre_carpeta}.zip")

        with ZipFile(ruta_zip, 'w', ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(ruta_carpeta):
                for f in files:
                    if f.startswith('.') or f.startswith('._'):
                        continue
                    abs_path = os.path.join(root, f)
                    rel_path = os.path.relpath(abs_path, ruta_carpeta)
                    zipf.write(abs_path, rel_path)

        print(f"✅ ZIP creado: {ruta_zip}")

def depurar_destino_final(destino):
    carpeta_zips = os.path.join(destino, "ZIPS_FINAL")

    for item in os.listdir(destino):
        ruta = os.path.join(destino, item)

        # No borrar la carpeta de ZIPS_FINAL
        if os.path.abspath(ruta) == os.path.abspath(carpeta_zips):
            continue

        try:
            if os.path.isdir(ruta):
                shutil.rmtree(ruta)
                print(f"🗑️ Carpeta eliminada: {ruta}")
            elif os.path.isfile(ruta):
                os.remove(ruta)
                print(f"🗑️ Archivo eliminado: {ruta}")
        except Exception as e:
            print(f"⚠️ Error al eliminar {ruta}: {e}")

    print("\n✅ Depuración completa. Solo queda la carpeta ZIPS_FINAL.")        

def convertir_word_a_pdf(ruta_word, carpeta_destino):
    import platform

    if platform.system() == "Windows":
        libreoffice_exec = "soffice.exe"
    elif platform.system() == "Darwin":  # macOS
        libreoffice_exec = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    else:
        libreoffice_exec = "libreoffice"

    if not os.path.exists(libreoffice_exec) and not shutil.which(libreoffice_exec):
        messagebox.showerror("LibreOffice no encontrado", f"No se encontró LibreOffice.\nRuta esperada:\n{libreoffice_exec}")
        return False

    try:
        subprocess.run([
            libreoffice_exec,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", carpeta_destino,
            ruta_word
        ], check=True)
        print(f"📝 Convertido a PDF: {ruta_word}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"⚠️ Error al convertir {ruta_word} a PDF: {e}")
        return False



# Función principal para procesar archivos (copiar/comprimir) en la carpeta de origen
def comprimir_pdfs():
    origen = entrada_origen.get()
    destino = entrada_destino.get()

    if not origen or not destino:
        messagebox.showerror("Error", "Debes seleccionar ambas carpetas.")
        return

    archivos_a_procesar = []
    for root, _, files in os.walk(origen):
        for file in files:
            if any(file.lower().endswith(ext) for ext in EXTENSIONES_VALIDAS):
                archivos_a_procesar.append(os.path.join(root, file))

    total_archivos = len(archivos_a_procesar)
    if total_archivos == 0:
        messagebox.showinfo("Sin archivos", "No se encontraron archivos para procesar.")
        return

    progress["maximum"] = total_archivos
    progress["value"] = 0

    # Copiamos archivos (compress PDF o copy otros)
    destino_final = destino
    for i, ruta_archivo in enumerate(archivos_a_procesar, start=1):
        rel_path = os.path.relpath(os.path.dirname(ruta_archivo), origen)
        destino_dir = os.path.join(destino_final, rel_path)
        os.makedirs(destino_dir, exist_ok=True)
        salida_archivo = os.path.join(destino_dir, os.path.basename(ruta_archivo))
        try:
            if ruta_archivo.lower().endswith(".pdf"):
                if not comprimir_pdf_ghostscript(ruta_archivo, salida_archivo):
                    print(f"No se pudo comprimir {ruta_archivo} con ninguna opción.")
            elif ruta_archivo.lower().endswith((".doc", ".docx")):
                if convertir_word_a_pdf(ruta_archivo, destino_dir):
                    print(f"Convertido Word ➜ PDF en: {destino_dir}")
                else:
                    print(f"No se pudo convertir Word: {ruta_archivo}")
            else:
                shutil.copy2(ruta_archivo, salida_archivo)
                print(f"Copiado: {ruta_archivo} ➜ {salida_archivo}")
        except Exception as e:
            print(f"Error procesando {ruta_archivo}: {e}")
        progress["value"] = i
        ventana.update_idletasks()

    messagebox.showinfo("Listo", "Todos los archivos han sido procesados.")

    # Procesamos subcarpetas para generar ZIPs
    subcarpetas = [
        d for d in os.listdir(origen)
        if os.path.isdir(os.path.join(origen, d))
    ]
    if not subcarpetas:
        messagebox.showinfo("Sin subcarpetas", "No se encontraron subcarpetas en la carpeta origen.")
        return

    for sub in subcarpetas:
        path_sub = os.path.join(destino_final, sub)
        comprimir_subcarpeta(path_sub, destino_final)

    messagebox.showinfo("Proceso finalizado", "Compresión y generación de ZIPs finalizada.\nRevisa la carpeta ZIPS en el destino.")
    descomprimir_zips(destino)
    revisar_estructuras_repetidas(destino)
    eliminar_archivos_metadata(destino)
    comprimir_descomprimidos(destino)
    depurar_destino_final(destino)

def descomprimir_zips(destino):
    """
    Busca todos los .zip en destino/ZIPS y los extrae a destino/DESCOMPRIMIDOS/<nombre_zip_sin_ext>.
    """
    carpeta_zips = os.path.join(destino, "ZIPS")
    carpeta_salida = os.path.join(destino, "DESCOMPRIMIDOS")
    os.makedirs(carpeta_salida, exist_ok=True)

    for nombre in os.listdir(carpeta_zips):
        if nombre.lower().endswith(".zip"):
            ruta_zip = os.path.join(carpeta_zips, nombre)
            nombre_base = os.path.splitext(nombre)[0]
            ruta_extraccion = os.path.join(carpeta_salida, nombre_base)
            os.makedirs(ruta_extraccion, exist_ok=True)

            try:
                with zipfile.ZipFile(ruta_zip, 'r') as zf:
                    zf.extractall(ruta_extraccion)
                print(f"Descomprimido: {ruta_zip} ➜ {ruta_extraccion}")
            except zipfile.BadZipFile as e:
                print(f"Error al descomprimir {ruta_zip}: {e}")



def comprimir_subcarpeta(path_subcarpeta, destino):
    nombre_sub = os.path.basename(path_subcarpeta)
    archivos = []
    for root, _, files in os.walk(path_subcarpeta):
        for f in files:
            if any(f.lower().endswith(ext) for ext in EXTENSIONES_VALIDAS):
                archivos.append(os.path.join(root, f))
    if not archivos:
        print(f"No se encontraron archivos en {path_subcarpeta}")
        return

    carpeta_zips = os.path.join(destino, "ZIPS")
    os.makedirs(carpeta_zips, exist_ok=True)
    dividir_y_comprimir_carpeta(path_subcarpeta, carpeta_zips, nombre_sub)

# Función para dividir los archivos de una subcarpeta en partes y generar ZIP(s).
def dividir_y_comprimir_carpeta(origen_carpeta, destino_zip, base_name):
    archivos = []
    for root, _, files in os.walk(origen_carpeta):
        for f in files:
            ruta = os.path.join(root, f)
            size_mb = os.path.getsize(ruta) / (1024*1024)
            archivos.append((ruta, size_mb))
    if not archivos:
        print(f"No hay archivos en {origen_carpeta}")
        return

    partes = []
    parte_actual = []
    tam_actual = 0
    for ruta, size_mb in archivos:
        if tam_actual + size_mb > MAX_ZIP_SIZE_MB and parte_actual:
            partes.append(parte_actual)
            parte_actual = []
            tam_actual = 0
        parte_actual.append(ruta)
        tam_actual += size_mb
    if parte_actual:
        partes.append(parte_actual)
    total_partes = len(partes)

    for i, parte in enumerate(partes, start=1):
        dyn_index = f"{i}-{total_partes}"
        new_root = f"{base_name} ({dyn_index})"
        nombre_zip = f"{new_root}.zip"
        ruta_zip = os.path.join(destino_zip, nombre_zip)
        with zipfile.ZipFile(ruta_zip, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
            for archivo in parte:
                arcname = crear_arcname(archivo, origen_carpeta, new_root, dyn_index)
                zipf.write(archivo, arcname)
        print(f"ZIP creado: {ruta_zip}")

# Funciones de selección de carpetas

def seleccionar_carpeta_origen():
    carpeta = filedialog.askdirectory()
    if carpeta:
        entrada_origen.set(carpeta)

def seleccionar_carpeta_destino():
    carpeta = filedialog.askdirectory()
    if carpeta:
        entrada_destino.set(carpeta)

# Interfaz Tkinter
ventana = tk.Tk()
ventana.title("Compresor de Subcarpetas - PDFs y Office")

entrada_origen = tk.StringVar()
entrada_destino = tk.StringVar()

tk.Label(ventana, text="Carpeta de origen (contiene subcarpetas a procesar):").pack(pady=5)
tk.Entry(ventana, textvariable=entrada_origen, width=60).pack()
tk.Button(ventana, text="Seleccionar carpeta origen", command=seleccionar_carpeta_origen).pack(pady=5)

tk.Label(ventana, text="Carpeta destino:").pack(pady=5)
tk.Entry(ventana, textvariable=entrada_destino, width=60).pack()
tk.Button(ventana, text="Seleccionar carpeta destino", command=seleccionar_carpeta_destino).pack(pady=5)

tk.Button(ventana, text="Comprimir Subcarpetas", command=comprimir_pdfs, bg="#4CAF50", fg="white").pack(pady=20)

progress = ttk.Progressbar(ventana, orient="horizontal", length=400, mode="determinate")
progress.pack(pady=10)

ventana.mainloop()