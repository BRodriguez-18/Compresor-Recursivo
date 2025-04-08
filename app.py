import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import zipfile
import shutil

# Parámetros
MAX_ZIP_SIZE_MB = 4
EXTENSIONES_VALIDAS = [".pdf", ".doc", ".docx", ".xls", ".xlsx"]

# Función para comprimir un PDF usando Ghostscript
def comprimir_pdf_ghostscript(entrada, salida):
    opciones = ["/screen", "/ebook", "/printer", "/prepress"]
    for opcion in opciones:
        comando = [
            "gs",
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
# new_root: cadena dinámica (ej. "Pepsi (1-2)")
# dyn_index: el índice dinámico que se aplicará en cada subnivel (ej. "1-2" o "2-2")
def crear_arcname(archivo, origen_carpeta, new_root, dyn_index):
    rel = os.path.relpath(archivo, origen_carpeta)  # Ej: "proveedores/pólizas de almacen/archivo1.pdf"
    dir_rel = os.path.dirname(rel)
    file_name = os.path.basename(rel)
    if not dir_rel:
        return os.path.join(new_root, file_name)
    componentes = dir_rel.split(os.sep)
    acumulado = new_root
    carpetas = [new_root]
    for comp in componentes:
        acumulado = f"{acumulado} {comp} ({dyn_index})"
        carpetas.append(acumulado)
    return os.path.join(*carpetas, file_name)

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

    # Se copia el contenido interno (lo que está adentro de la carpeta origen) al destino seleccionado
    destino_final = destino  # Trabajamos directamente con lo de adentro
    for i, ruta_archivo in enumerate(archivos_a_procesar, start=1):
        rel_path = os.path.relpath(os.path.dirname(ruta_archivo), origen)
        destino_dir = os.path.join(destino_final, rel_path)
        os.makedirs(destino_dir, exist_ok=True)
        salida_archivo = os.path.join(destino_dir, os.path.basename(ruta_archivo))
        try:
            if ruta_archivo.lower().endswith(".pdf"):
                if not comprimir_pdf_ghostscript(ruta_archivo, salida_archivo):
                    print(f"No se pudo comprimir {ruta_archivo} con ninguna opción.")
            else:
                shutil.copy2(ruta_archivo, salida_archivo)
                print(f"Copiado: {ruta_archivo} ➜ {salida_archivo}")
        except Exception as e:
            print(f"Error procesando {ruta_archivo}: {e}")
        progress["value"] = i
        ventana.update_idletasks()

    messagebox.showinfo("Listo", "Todos los archivos han sido procesados.")

    # Suponemos que los contenidos adentro de la carpeta origen ya fueron copiados al destino.
    # Ahora, para comprimir, se procesarán las subcarpetas de la carpeta origen (es decir, se ignora la carpeta base).
    subcarpetas = [
        d for d in os.listdir(origen)
        if os.path.isdir(os.path.join(origen, d))
    ]
    if not subcarpetas:
        messagebox.showinfo("Sin subcarpetas", "No se encontraron subcarpetas en la carpeta origen.")
        return

    for sub in subcarpetas:
        path_sub = os.path.join(destino_final, sub)  # trabajar con lo que se copió
        comprimir_subcarpeta(path_sub, destino_final)

    messagebox.showinfo("Proceso finalizado", "Compresión y generación de ZIPs finalizada.\nRevisa la carpeta ZIPS en el destino.")

# Función para procesar y comprimir cada subcarpeta (por ejemplo, "Punto1", "Punto2", etc.)
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

    # Se generan ZIP(s) para esta subcarpeta en la carpeta ZIPS
    carpeta_zips = os.path.join(destino, "ZIPS")
    os.makedirs(carpeta_zips, exist_ok=True)
    dividir_y_comprimir_carpeta(path_subcarpeta, carpeta_zips, nombre_sub)

# Función para dividir los archivos de una subcarpeta en partes (si es mayor a 4 MB) y generar ZIP(s).
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
        # El índice dinámico será "i-total_partes", que queremos usar en toda la estructura interna.
        dyn_index = f"{i}-{total_partes}"
        new_root = f"{base_name} ({dyn_index})"
        nombre_zip = f"{new_root}.zip"
        ruta_zip = os.path.join(destino_zip, nombre_zip)
        with zipfile.ZipFile(ruta_zip, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
            for archivo in parte:
                arcname = crear_arcname(archivo, origen_carpeta, new_root, dyn_index)
                zipf.write(archivo, arcname)
        print(f"ZIP creado: {ruta_zip}")

# Funciones de selección de carpetas en la interfaz
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
