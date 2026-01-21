from openpyxl import Workbook

# Datos de pacientes y cosas a encargar
pacientes = {
    "Juan Pérez": ["Medicamento A", "Estudio de sangre", "Radiografía"],
    "María López": ["Medicamento B", "Electrocardiograma"],
    "Carlos Gómez": ["Medicamento C", "Resonancia magnética"]
}

# Diccionario para guardar qué está tildado
tildado = {paciente: [False] * len(items) for paciente, items in pacientes.items()}

# Función para marcar ítems
def marcar_item(paciente, numero_item):
    if 1 <= numero_item <= len(pacientes[paciente]):
        tildado[paciente][numero_item-1] = True
        print(f"Se marcó '{pacientes[paciente][numero_item-1]}' para {paciente}.")
    else:
        print("Número de ítem inválido.")

# Función para exportar a Excel
def exportar_excel(nombre_archivo="checklist.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist"

    # Encabezados
    ws.append(["Paciente", "Ítem", "Estado"])

    # Volcar datos
    for paciente, items in pacientes.items():
        for i, item in enumerate(items):
            estado = "✔️" if tildado[paciente][i] else "❌"
            ws.append([paciente, item, estado])

    wb.save(nombre_archivo)
    print(f"Datos exportados a {nombre_archivo}")

# Ejemplo de uso
marcar_item("Juan Pérez", 2)
exportar_excel("encargos_pacientes.xlsx")