import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference

# =====================
# FUNCIÓN PARA CREAR EL EXCEL
# =====================
def convertir_excel(df_filtrado, resumen_general, resumen_por_esquema):
    output = BytesIO()
    wb = Workbook()

    # --- Hoja 1: Datos filtrados
    ws1 = wb.active
    ws1.title = "Filtrado"
    for r in dataframe_to_rows(df_filtrado, index=False, header=True):
        ws1.append(r)

    # --- Hoja 2: Resumen general
    ws2 = wb.create_sheet("Resumen General")
    for r in dataframe_to_rows(resumen_general, index=False, header=True):
        ws2.append(r)

    # Crear gráfico de líneas en "Resumen General"
    if not resumen_general.empty:
        # Suponiendo que la primera columna es "Año" y la segunda es "Promedio"
        chart = LineChart()
        chart.title = "Promedio por Año"
        chart.style = 10
        chart.y_axis.title = "Valor Promedio"
        chart.x_axis.title = "Año"
        chart.marker = True
        chart.smooth = True

        # Datos para el gráfico
        data = Reference(ws2, min_col=2, min_row=1, max_col=2, max_row=len(resumen_general) + 1)
        cats = Reference(ws2, min_col=1, min_row=2, max_row=len(resumen_general) + 1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        ws2.add_chart(chart, "E2")

    # --- Hoja 3: Resumen por esquema
    ws3 = wb.create_sheet("Resumen por Esquema")
    for r in dataframe_to_rows(resumen_por_esquema, index=False, header=True):
        ws3.append(r)

    # Crear gráfico de líneas por esquema si hay datos
    if not resumen_por_esquema.empty:
        # Suponiendo que la estructura es: ['Esquema', 'Año', 'Promedio']
        # Agregamos un gráfico por cada esquema
        esquemas = resumen_por_esquema['Esquema'].unique()
        start_row = 2
        for esquema in esquemas:
            df_temp = resumen_por_esquema[resumen_por_esquema['Esquema'] == esquema]
            if df_temp.empty:
                continue

            chart = LineChart()
            chart.title = f"{esquema} - Promedio por Año"
            chart.style = 10
            chart.y_axis.title = "Valor Promedio"
            chart.x_axis.title = "Año"
            chart.marker = True
            chart.smooth = True

            # Insertar datos en la hoja si no existen
            for r in dataframe_to_rows(df_temp, index=False, header=True):
                ws3.append(r)

            # Calcular rango de datos
            end_row = start_row + len(df_temp)
            data = Reference(ws3, min_col=3, min_row=start_row, max_col=3, max_row=end_row)
            cats = Reference(ws3, min_col=2, min_row=start_row + 1, max_row=end_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)

            # Insertar gráfico
            ws3.add_chart(chart, f"E{start_row}")

            # Mover inicio para el siguiente esquema
            start_row = end_row + 3

    # Guardar archivo en memoria
    wb.save(output)
    output.seek(0)
    return output


# =====================
# INTERFAZ STREAMLIT
# =====================
st.title("📈 Exportar Datos y Gráficos a Excel")

# Ejemplo de datos (puedes reemplazar con tus propios DataFrames)
data = {
    'Año': [2020, 2021, 2022, 2023],
    'Promedio': [100, 120, 90, 150]
}
df_filtrado = pd.DataFrame(data)
resumen_general = df_filtrado.copy()

resumen_por_esquema = pd.DataFrame({
    'Esquema': ['A', 'A', 'B', 'B'],
    'Año': [2020, 2021, 2020, 2021],
    'Promedio': [80, 110, 90, 130]
})

# Convertir a Excel
data = convertir_excel(df_filtrado, resumen_general, resumen_por_esquema)

# Botón de descarga
st.download_button(
    label="📥 Descargar Excel con Gráficos de Líneas",
    data=data,
    file_name="reporte_cnsf.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
