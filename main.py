import tkinter as tk
from tkinter import filedialog, messagebox
from consolidar_tablas import consolidar
from comparar_libros_excel import comparar


def select_file(entry):
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)


def save_file(entry):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[
                                             ("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)


def open_consolidar_window():
    consolidar_window = tk.Toplevel(root)
    consolidar_window.title("Consolidar Excel")

    tk.Label(consolidar_window, text="Archivo de entrada:").grid(
        row=0, column=0, sticky=tk.W)
    entry_input = tk.Entry(consolidar_window, width=50)
    entry_input.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(consolidar_window, text="Seleccionar", command=lambda: select_file(
        entry_input)).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(consolidar_window, text="Archivo de salida:").grid(
        row=1, column=0, sticky=tk.W)
    entry_output = tk.Entry(consolidar_window, width=50)
    entry_output.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(consolidar_window, text="Guardar como", command=lambda: save_file(
        entry_output)).grid(row=1, column=2, padx=5, pady=5)

    def run_consolidar():
        ruta_base = entry_input.get()
        ruta_salida = entry_output.get()
        if not ruta_base or not ruta_salida:
            messagebox.showerror(
                "Error", "Por favor, selecciona los archivos de entrada y salida")
            return
        try:
            consolidar(ruta_base, ruta_salida)
            messagebox.showinfo("Éxito", "El proceso se completó exitosamente")
        except Exception as e:
            messagebox.showerror("Error", f"Ha ocurrido un error: {e}")

    tk.Button(consolidar_window, text="Ejecutar", command=run_consolidar).grid(
        row=2, column=0, columnspan=3, pady=10)


def open_comparar_window():
    comparar_window = tk.Toplevel(root)
    comparar_window.title("Comparar archivos")

    tk.Label(comparar_window, text="Archivo de entrada 1:").grid(
        row=0, column=0, sticky=tk.W)
    entry_input1 = tk.Entry(comparar_window, width=50)
    entry_input1.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(comparar_window, text="Seleccionar", command=lambda: select_file(
        entry_input1)).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(comparar_window, text="Archivo de entrada 2:").grid(
        row=1, column=0, sticky=tk.W)
    entry_input2 = tk.Entry(comparar_window, width=50)
    entry_input2.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(comparar_window, text="Seleccionar", command=lambda: select_file(
        entry_input2)).grid(row=1, column=2, padx=5, pady=5)

    tk.Label(comparar_window, text="Archivo de salida:").grid(
        row=2, column=0, sticky=tk.W)
    entry_output = tk.Entry(comparar_window, width=50)
    entry_output.grid(row=2, column=1, padx=5, pady=5)
    tk.Button(comparar_window, text="Guardar como", command=lambda: save_file(
        entry_output)).grid(row=2, column=2, padx=5, pady=5)

    def run_comparar():
        ruta_libro1 = entry_input1.get()
        ruta_libro2 = entry_input2.get()
        ruta_salida = entry_output.get()
        if not ruta_libro1 or not ruta_libro2 or not ruta_salida:
            messagebox.showerror(
                "Error", "Por favor, selecciona los archivos de entrada y salida")
            return
        try:
            comparar(ruta_libro1, ruta_libro2, ruta_salida)
            messagebox.showinfo("Éxito", "El proceso se completó exitosamente")
        except Exception as e:
            messagebox.showerror("Error", f"Ha ocurrido un error: {e}")

    tk.Button(comparar_window, text="Ejecutar", command=run_comparar).grid(
        row=3, column=0, columnspan=3, pady=10)


root = tk.Tk()
root.geometry("300x200")

root.title("Gestor de Archivos Excel")

tk.Button(root, text="Consolidar Excel",
          command=open_consolidar_window).pack(pady=30, padx=10)
tk.Button(root, text="Comparar Excel",
          command=open_comparar_window).pack(pady=10, padx=10)

root.mainloop()
