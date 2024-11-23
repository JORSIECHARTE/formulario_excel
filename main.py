import tkinter as tk
from tkinter import messagebox
import re
from openpyxl import Workbook


def validar_numeros(texto):
    return texto.isdigit() or texto == ""


def validar_email():
    email = entry_email.get()
    patron_email = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(patron_email, email)


def guardar_datos():
    nombre = entry_nombre.get()
    edad = entry_edad.get()
    email = entry_email.get()
    telefono = entry_telefono.get()
    direccion = entry_direccion.get()

    if not nombre:
        messagebox.showerror("Error", "El campo 'Nombre' es obligatorio.")
        return
    if not edad:
        messagebox.showerror("Error", "El campo 'Edad' es obligatorio.")
        return
    if not email or not validar_email():
        messagebox.showerror("Error", "Por favor, ingresa un email válido.")
        return
    if not telefono:
        messagebox.showerror("Error", "El campo 'Teléfono' es obligatorio.")
        return

    try:
        wb = Workbook()
        ws = wb.active
        if ws.max_row == 1:
            ws.append(["Nombre", "Edad", "Email", "Teléfono", "Dirección"])

        ws.append([nombre, edad, email, telefono, direccion])
        wb.save("datos_formulario.xlsx")

        messagebox.showinfo("Éxito", "Los datos se han guardado correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al guardar los datos: {e}")

    limpiar_campos()


def limpiar_campos():
    entry_nombre.delete(0, tk.END)
    entry_edad.delete(0, tk.END)
    entry_email.delete(0, tk.END)
    entry_telefono.delete(0, tk.END)
    entry_direccion.delete(0, tk.END)


root = tk.Tk()
root.title("Formulario de Entrada de Datos")
root.configure(bg='#4B6587')

label_style = {"bg": '#4B6587', "fg": "white"}
entry_style = {"bg": '#D3D3D3', "fg": "black"}

vcmd_numeros = root.register(validar_numeros)

label_nombre = tk.Label(root, text="Nombre", **label_style)
label_nombre.grid(row=0, column=0, padx=10, pady=5)
entry_nombre = tk.Entry(root, **entry_style)
entry_nombre.grid(row=0, column=1, padx=10, pady=5)

label_edad = tk.Label(root, text="Edad", **label_style)
label_edad.grid(row=1, column=0, padx=10, pady=5)
entry_edad = tk.Entry(root, validate="key", validatecommand=(vcmd_numeros, '%P'), **entry_style)
entry_edad.grid(row=1, column=1, padx=10, pady=5)

label_email = tk.Label(root, text="Email", **label_style)
label_email.grid(row=2, column=0, padx=10, pady=5)
entry_email = tk.Entry(root, **entry_style)
entry_email.grid(row=2, column=1, padx=10, pady=5)

label_telefono = tk.Label(root, text="Teléfono", **label_style)
label_telefono.grid(row=3, column=0, padx=10, pady=5)
entry_telefono = tk.Entry(root, validate="key", validatecommand=(vcmd_numeros, '%P'), **entry_style)
entry_telefono.grid(row=3, column=1, padx=10, pady=5)

label_direccion = tk.Label(root, text="Dirección", **label_style)
label_direccion.grid(row=4, column=0, padx=10, pady=5)
entry_direccion = tk.Entry(root, **entry_style)
entry_direccion.grid(row=4, column=1, padx=10, pady=5)

btn_guardar = tk.Button(root, text="Guardar", command=guardar_datos, bg='#1e90ff', fg='white')
btn_guardar.grid(row=5, column=0, columnspan=2, pady=10)

root.mainloop()
