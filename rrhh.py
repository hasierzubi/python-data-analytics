#!/usr/bin/env python3
"""GUI paso a paso (versión 6).

Cambios destacados:
- El mensaje de éxito ahora distingue entre carpeta creada y carpeta ya existente. reportlab 
"""

import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path

# ---------------------------------------------------------------------------
# Carpeta base donde se creará/verificará la subcarpeta con la fecha de hoy.
BASE_DIR = Path(r"C:\Users\hzubi\OneDrive\Escritorio\airbus\gui")  # <— AJUSTA AQUÍ
# ---------------------------------------------------------------------------

BASE_DIR.mkdir(parents=True, exist_ok=True)


def select_file() -> None:
    """Abre un cuadro de diálogo para elegir un archivo y muestra la ruta."""
    filepath = filedialog.askopenfilename(
        title="Selecciona un archivo",
        filetypes=[("Todos los archivos", "*.*")],
    )
    if filepath:
        file_var.set(filepath)


def get_today_folder() -> Path:
    """Devuelve la ruta BASE_DIR/YYYY-MM-DD como objeto Path."""
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    return BASE_DIR / today_str


def visualize() -> None:
    """Crea/verifica la carpeta con la fecha de hoy y muestra un mensaje adecuado."""
    dest_folder = get_today_folder()
    existed_before = dest_folder.exists()

    try:
        dest_folder.mkdir(parents=True, exist_ok=True)
        if existed_before:
            msg = f"La carpeta ya existía:\n{dest_folder}"
        else:
            msg = f"La carpeta se creó correctamente:\n{dest_folder}"
        messagebox.showinfo(title="Resultado", message=msg)
    except Exception as exc:
        messagebox.showerror(
            title="Error al crear carpeta",
            message=f"No se pudo crear la carpeta:\n{exc}",
        )


def main() -> None:
    root = tk.Tk()
    root.title("Seleccionar archivo y número de días")

    # --- Marco principal ---
    main_frame = ttk.Frame(root, padding=10)
    main_frame.grid(row=0, column=0, sticky="nsew")
    main_frame.columnconfigure(1, weight=1)  # La columna de la ruta se expande

    # Variable global para la ruta del archivo seleccionada
    global file_var
    file_var = tk.StringVar()

    # Botón para seleccionar archivo
    ttk.Button(
        main_frame,
        text="Selecciona un archivo",
        command=select_file,
    ).grid(row=0, column=0, padx=(0, 5))

    # Entrada que muestra la ruta seleccionada
    ttk.Entry(
        main_frame,
        textvariable=file_var,
        width=40,
    ).grid(row=0, column=1, sticky="ew", padx=(0, 10))

    # Etiqueta y entrada para el número de días
    ttk.Label(main_frame, text="Introduce el número de días:").grid(row=0, column=2, padx=(0, 5))
    number_var = tk.StringVar()
    ttk.Entry(main_frame, textvariable=number_var, width=10).grid(row=0, column=3)

    # Botón Visualizar
    ttk.Button(main_frame, text="Visualizar", command=visualize).grid(row=1, column=0, columnspan=4, pady=(10, 0))

    root.mainloop()


if __name__ == "__main__":
    main()
