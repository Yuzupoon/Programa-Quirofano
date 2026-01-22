import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook

# Diccionario de pacientes
pacientes = {}

# Función para agregar paciente
def agregar_paciente():
    nombre = entry_paciente.get().strip()
    item = entry_item.get().strip()
    if nombre and item:
        if nombre not in pacientes:
            pacientes[nombre] = []
        pacientes[nombre].append({"item": item, "estado": False})
        actualizar_lista()
        entry_item.delete(0, tk.END)
    else:
        messagebox.showwarning("Atención", "Debe ingresar paciente e ítem.")

# Función para marcar ítem como realizado
def marcar_item():
    seleccion = lista_items.selection()
    if seleccion:
        paciente, idx = lista_items.item(seleccion[0], "values")[0], int(lista_items.item(seleccion[0], "values")[1])
        pacientes[paciente][idx]["estado"] = True
        actualizar_lista()

# Función para exportar a Excel
def exportar_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist"
    ws.append(["Paciente", "Ítem", "Estado"])
    
    for paciente, items in pacientes.items():
        for i, item in enumerate(items):
            estado = "✔️" if item["estado"] else "❌"
            ws.append([paciente, item["item"], estado])
    
    wb.save("checklist_pacientes.xlsx")
    messagebox.showinfo("Éxito", "Datos exportados a checklist_pacientes.xlsx")

# Función para actualizar la lista visual
def actualizar_lista():
    for row in lista_items.get_children():
        lista_items.delete(row)
    for paciente, items in pacientes.items():
        for i, item in enumerate(items):
            estado = "✔️" if item["estado"] else "❌"
            lista_items.insert("", tk.END, values=(paciente, i, item["item"], estado))

# Interfaz gráfica
root = tk.Tk()
root.title("Checklist de Pacientes")

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Label(frame, text="Paciente:").grid(row=0, column=0)
entry_paciente = tk.Entry(frame)
entry_paciente.grid(row=0, column=1)

tk.Label(frame, text="Ítem:").grid(row=1, column=0)
entry_item = tk.Entry(frame)
entry_item.grid(row=1, column=1)

btn_agregar = tk.Button(frame, text="Agregar", command=agregar_paciente)
btn_agregar.grid(row=2, column=0, columnspan=2, pady=5)

# Tabla de ítems
cols = ("Paciente", "Índice", "Ítem", "Estado")
lista_items = ttk.Treeview(root, columns=cols, show="headings")
for col in cols:
    lista_items.heading(col, text=col)
lista_items.pack(pady=10)

btn_marcar = tk.Button(root, text="Marcar como realizado", command=marcar_item)
btn_marcar.pack(pady=5)

btn_exportar = tk.Button(root, text="Exportar a Excel", command=exportar_excel)
btn_exportar.pack(pady=5)

root.mainloop()