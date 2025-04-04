import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import zipfile

# Parámetros
MAX_ZIP_SIZE_MB = 4

# Comprimir un archivo PDF usando Ghostscript
def comprimir_pdf_ghostscript(entrada, salida):
    comando = [
        "gs",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        "-dPDFSETTINGS=/screen",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        f"-sOutputFile={salida}",
        entrada
    ]
    subprocess.run(comando)

# Función principal para comprimir PDFs
def comprimir_pdfs():
    origen = entrada_origen.get()
    destino = entrada_destino.get()

    if not origen or not destino:
        messagebox.showerror("Error", "Debes seleccionar ambas carpetas.")
        return

    pdfs_a_procesar = []
    for root, _, files in os.walk(origen):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdfs_a_procesar.append(os.path.join(root, file))

    total_archivos = len(pdfs_a_procesar)
    if total_archivos == 0:
        messagebox.showinfo("Sin archivos", "No se encontraron archivos PDF para comprimir.")
        return

    progress["maximum"] = total_archivos
    progress["value"] = 0

    for i, ruta_pdf in enumerate(pdfs_a_procesar, start=1):
        rel_path = os.path.relpath(os.path.dirname(ruta_pdf), origen)
        destino_dir = os.path.join(destino, rel_path)
        os.makedirs(destino_dir, exist_ok=True)

        salida_pdf = os.path.join(destino_dir, os.path.basename(ruta_pdf))

        try:
            comprimir_pdf_ghostscript(ruta_pdf, salida_pdf)
            print(f"Comprimido: {ruta_pdf} ➜ {salida_pdf}")
        except Exception as e:
            print(f"Error al comprimir {ruta_pdf}: {e}")

        progress["value"] = i
        ventana.update_idletasks()

    messagebox.showinfo("Listo", "Todos los archivos han sido comprimidos.")

    # Crear carpeta para ZIPs
    zip_destino = os.path.join(destino, "ZIPS")
    os.makedirs(zip_destino, exist_ok=True)
    dividir_y_comprimir_por_grupo(destino, zip_destino)

    messagebox.showinfo("ZIPs creados", f"Los archivos ZIP se guardaron en:\n{zip_destino}")

# Agrupar por carpetas y crear ZIPs
def dividir_y_comprimir_por_grupo(carpeta_base, salida_zip):
    for root, dirs, _ in os.walk(carpeta_base):
        for dir in dirs:
            carpeta_objetivo = os.path.join(root, dir)
            if "ZIPS" not in carpeta_objetivo:  # evitar recursividad con carpeta ZIP
                comprimir_por_tamaño(carpeta_objetivo, salida_zip)

# Dividir los archivos en grupos de máximo 4 MB y generar ZIPs
def comprimir_por_tamaño(origen_carpeta, destino_zip):
    archivos = []
    nombre_base = os.path.basename(origen_carpeta)

    for root, _, files in os.walk(origen_carpeta):
        for file in files:
            ruta = os.path.join(root, file)
            size_mb = os.path.getsize(ruta) / (1024 * 1024)
            archivos.append((ruta, size_mb))

    partes = []
    parte_actual = []
    tamaño_actual = 0

    for ruta, size_mb in archivos:
        if tamaño_actual + size_mb > MAX_ZIP_SIZE_MB and parte_actual:
            partes.append(parte_actual)
            parte_actual = []
            tamaño_actual = 0

        parte_actual.append(ruta)
        tamaño_actual += size_mb

    if parte_actual:
        partes.append(parte_actual)

    for i, parte in enumerate(partes, start=1):
        nombre_zip = f"{nombre_base} ({i}-{len(partes)}).zip"
        ruta_zip = os.path.join(destino_zip, nombre_zip)

        with zipfile.ZipFile(ruta_zip, "w", compression=zipfile.ZIP_DEFLATED) as zipf:
            for archivo in parte:
                arc_rel = os.path.relpath(archivo, origen_carpeta)
                zipf.write(archivo, arc_rel)

        print(f"ZIP creado: {ruta_zip}")

# Selección de carpetas
def seleccionar_carpeta_origen():
    carpeta = filedialog.askdirectory()
    if carpeta:
        entrada_origen.set(carpeta)

def seleccionar_carpeta_destino():
    carpeta = filedialog.askdirectory()
    if carpeta:
        entrada_destino.set(carpeta)
        with open("carpeta_destino.txt", "w") as f:
            f.write(carpeta)

# Interfaz
ventana = tk.Tk()
ventana.title("Compresor de PDFs + Zips (Ghostscript + Tkinter)")

entrada_origen = tk.StringVar()
entrada_destino = tk.StringVar()

tk.Label(ventana, text="Carpeta a analizar:").pack(pady=5)
tk.Entry(ventana, textvariable=entrada_origen, width=50).pack()
tk.Button(ventana, text="Seleccionar carpeta origen", command=seleccionar_carpeta_origen).pack(pady=5)

tk.Label(ventana, text="Carpeta destino:").pack(pady=5)
tk.Entry(ventana, textvariable=entrada_destino, width=50).pack()
tk.Button(ventana, text="Seleccionar carpeta destino", command=seleccionar_carpeta_destino).pack(pady=5)

tk.Button(ventana, text="Comprimir PDFs y generar ZIPs", command=comprimir_pdfs, bg="#4CAF50", fg="white").pack(pady=20)

progress = ttk.Progressbar(ventana, orient="horizontal", length=300, mode="determinate")
progress.pack(pady=10)

ventana.mainloop()
