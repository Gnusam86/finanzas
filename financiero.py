import tkinter as tk
from tkinter import messagebox, filedialog
import matplotlib.pyplot as plt
from openpyxl import Workbook

# Función para formatear números como moneda
def formato_moneda(valor):
    return f"${valor:,.2f}"

# Función para calcular el análisis financiero
def calcular_analisis():
    try:
        # Obtener datos de la inversión inicial
        desarrollo_app = float(entry_desarrollo.get())
        marketing_inicial = float(entry_marketing.get())
        infraestructura = float(entry_infraestructura.get())
        gastos_legales = float(entry_gastos_legales.get())
        reserva_efectivo = float(entry_reserva.get())

        inversion_inicial = desarrollo_app + marketing_inicial + infraestructura + gastos_legales + reserva_efectivo

        # Obtener datos de la proyección de ingresos y gastos
        meses = int(entry_meses.get())
        ingresos = []
        costos_fijos = []
        costos_variables = []

        for i in range(meses):
            ingresos.append(float(entry_ingresos[i].get()))
            costos_fijos.append(float(entry_costos_fijos[i].get()))
            costos_variables.append(float(entry_costos_variables[i].get()))

        # Calcular flujo de efectivo
        flujo_efectivo = []
        saldo_acumulado = -inversion_inicial  # Incluye la inversión inicial como saldo negativo

        for i in range(meses):
            flujo_mes = ingresos[i] - costos_fijos[i] - costos_variables[i]
            saldo_acumulado += flujo_mes
            flujo_efectivo.append(flujo_mes)

        # Calcular punto de equilibrio
        costos_totales = sum(costos_fijos) + sum(costos_variables)
        ingresos_totales = sum(ingresos)
        punto_equilibrio = costos_totales  # Punto de equilibrio en términos de ingresos

        # Balance general
        activos = saldo_acumulado if saldo_acumulado > 0 else 0
        pasivos = -saldo_acumulado if saldo_acumulado < 0 else 0
        patrimonio_neto = activos - pasivos

        # Mostrar resultados en una ventana emergente
        resultado = f"=== REPORTE FINANCIERO ===\n\n"
        resultado += f"--- Inversión Inicial ---\n"
        resultado += f"Desarrollo de la aplicación: {formato_moneda(desarrollo_app)}\n"
        resultado += f"Marketing inicial: {formato_moneda(marketing_inicial)}\n"
        resultado += f"Infraestructura tecnológica: {formato_moneda(infraestructura)}\n"
        resultado += f"Gastos legales y administrativos: {formato_moneda(gastos_legales)}\n"
        resultado += f"Reserva de efectivo: {formato_moneda(reserva_efectivo)}\n"
        resultado += f"Total inversión inicial: {formato_moneda(inversion_inicial)}\n\n"

        resultado += f"--- Proyección de Ingresos y Gastos ---\n"
        for mes in range(meses):
            resultado += f"\nMes {mes + 1}:\n"
            resultado += f"Ingresos: {formato_moneda(ingresos[mes])}\n"
            resultado += f"Costos fijos: {formato_moneda(costos_fijos[mes])}\n"
            resultado += f"Costos variables: {formato_moneda(costos_variables[mes])}\n"
            resultado += f"Flujo de efectivo neto: {formato_moneda(flujo_efectivo[mes])}\n"

        resultado += f"\n--- Flujo de Efectivo Acumulado ---\n"
        resultado += f"Saldo acumulado después de {meses} meses: {formato_moneda(saldo_acumulado)}\n\n"

        resultado += f"--- Balance General ---\n"
        resultado += f"Activos: {formato_moneda(activos)}\n"
        resultado += f"Pasivos: {formato_moneda(pasivos)}\n"
        resultado += f"Patrimonio Neto: {formato_moneda(patrimonio_neto)}\n\n"

        resultado += f"--- Análisis de Rentabilidad ---\n"
        resultado += f"Ingresos totales: {formato_moneda(ingresos_totales)}\n"
        resultado += f"Costos totales: {formato_moneda(costos_totales)}\n"
        resultado += f"Punto de equilibrio (ingresos necesarios para cubrir costos): {formato_moneda(punto_equilibrio)}\n\n"

        if saldo_acumulado > 0:
            resultado += "¡El emprendimiento es rentable! 🎉"
        else:
            resultado += "El emprendimiento no es rentable en este momento. Considera ajustar costos o aumentar ingresos. 💡"

        messagebox.showinfo("Resultados del Análisis Financiero", resultado)

        # Generar gráficos
        generar_graficos(meses, ingresos, costos_fijos, costos_variables, flujo_efectivo)

        # Exportar a Excel
        exportar_excel(meses, ingresos, costos_fijos, costos_variables, flujo_efectivo, saldo_acumulado)

    except ValueError:
        messagebox.showerror("Error", "Por favor, ingresa valores numéricos válidos.")

# Función para generar gráficos
def generar_graficos(meses, ingresos, costos_fijos, costos_variables, flujo_efectivo):
    meses_lista = list(range(1, meses + 1))

    plt.figure(figsize=(12, 6))

    # Gráfico de ingresos y costos
    plt.subplot(1, 2, 1)
    plt.plot(meses_lista, ingresos, label="Ingresos", marker="o")
    plt.plot(meses_lista, costos_fijos, label="Costos Fijos", marker="o")
    plt.plot(meses_lista, costos_variables, label="Costos Variables", marker="o")
    plt.title("Ingresos vs Costos")
    plt.xlabel("Meses")
    plt.ylabel("Monto ($)")
    plt.legend()
    plt.grid()

    # Gráfico de flujo de efectivo
    plt.subplot(1, 2, 2)
    plt.plot(meses_lista, flujo_efectivo, label="Flujo de Efectivo", marker="o", color="green")
    plt.title("Flujo de Efectivo")
    plt.xlabel("Meses")
    plt.ylabel("Monto ($)")
    plt.legend()
    plt.grid()

    plt.tight_layout()
    plt.show()

# Función para exportar a Excel
def exportar_excel(meses, ingresos, costos_fijos, costos_variables, flujo_efectivo, saldo_acumulado):
    try:
        # Crear un nuevo libro de Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Análisis Financiero"

        # Escribir encabezados
        ws.append(["Mes", "Ingresos", "Costos Fijos", "Costos Variables", "Flujo de Efectivo"])
        for i in range(meses):
            ws.append([i + 1, ingresos[i], costos_fijos[i], costos_variables[i], flujo_efectivo[i]])

        # Escribir saldo acumulado
        ws.append([])
        ws.append(["Saldo Acumulado", formato_moneda(saldo_acumulado)])

        # Guardar el archivo
        archivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if archivo:
            wb.save(archivo)
            messagebox.showinfo("Éxito", f"El archivo se ha guardado en {archivo}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo exportar el archivo: {e}")

# Crear la interfaz gráfica
root = tk.Tk()
root.title("Análisis Financiero para Aplicación Móvil")

# Campos para la inversión inicial
tk.Label(root, text="Desarrollo de la aplicación:").grid(row=0, column=0)
entry_desarrollo = tk.Entry(root)
entry_desarrollo.grid(row=0, column=1)

tk.Label(root, text="Marketing inicial:").grid(row=1, column=0)
entry_marketing = tk.Entry(root)
entry_marketing.grid(row=1, column=1)

tk.Label(root, text="Infraestructura tecnológica:").grid(row=2, column=0)
entry_infraestructura = tk.Entry(root)
entry_infraestructura.grid(row=2, column=1)

tk.Label(root, text="Gastos legales y administrativos:").grid(row=3, column=0)
entry_gastos_legales = tk.Entry(root)
entry_gastos_legales.grid(row=3, column=1)

tk.Label(root, text="Reserva de efectivo:").grid(row=4, column=0)
entry_reserva = tk.Entry(root)
entry_reserva.grid(row=4, column=1)

# Campos para la proyección de ingresos y gastos
tk.Label(root, text="Número de meses para la proyección:").grid(row=5, column=0)
entry_meses = tk.Entry(root)
entry_meses.grid(row=5, column=1)

# Crear campos dinámicos para ingresos y gastos
entry_ingresos = []
entry_costos_fijos = []
entry_costos_variables = []

def crear_campos_proyeccion():
    meses = int(entry_meses.get())
    for i in range(meses):
        tk.Label(root, text=f"Mes {i + 1} - Ingresos:").grid(row=6 + i, column=0)
        entry_ingresos.append(tk.Entry(root))
        entry_ingresos[i].grid(row=6 + i, column=1)

        tk.Label(root, text=f"Mes {i + 1} - Costos Fijos:").grid(row=6 + i, column=2)
        entry_costos_fijos.append(tk.Entry(root))
        entry_costos_fijos[i].grid(row=6 + i, column=3)

        tk.Label(root, text=f"Mes {i + 1} - Costos Variables:").grid(row=6 + i, column=4)
        entry_costos_variables.append(tk.Entry(root))
        entry_costos_variables[i].grid(row=6 + i, column=5)

# Botón para crear campos de proyección
tk.Button(root, text="Crear Campos de Proyección", command=crear_campos_proyeccion).grid(row=5, column=2, columnspan=2)

# Botón para calcular el análisis
tk.Button(root, text="Calcular Análisis", command=calcular_analisis).grid(row=100, column=0, columnspan=6)

# Ejecutar la interfaz gráfica
root.mainloop()