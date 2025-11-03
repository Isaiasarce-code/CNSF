import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows 

st.set_page_config(page_title="Filtro de Datos", layout="wide")

st.title("📊 Filtro de tabla por banderas (todas las hojas)")

# --- Cargar archivo Excel ---
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])

if uploaded_file:
    # Cargar todas las hojas del Excel como diccionario de DataFrames
    hojas = pd.read_excel(uploaded_file, sheet_name=None)
    nombres_hojas = list(hojas.keys())

    # Verificar columnas en la primera hoja como referencia
    columnas_requeridas = ["Entidad", "Modalidad", "Ciclo", "Cultivo"]
    primera_hoja = list(hojas.values())[0]

    if all(col in primera_hoja.columns for col in columnas_requeridas):
        
        # --- Filtros dinámicos (basados en la primera hoja) ---
        st.sidebar.header("Filtros")
        
        entidades_disponibles = sorted(primera_hoja["Entidad"].dropna().unique().tolist())
        entidad_opcion = st.sidebar.selectbox("Selecciona Entidad", ["(Todos)"] + entidades_disponibles)
        
        modalidades_disponibles = sorted(primera_hoja["Modalidad"].dropna().unique().tolist())
        modalidad_opcion = st.sidebar.selectbox("Selecciona Modalidad", ["(Todos)"] + modalidades_disponibles)
        
        ciclos_disponibles = sorted(primera_hoja["Ciclo"].dropna().unique().tolist())
        ciclo_opcion = st.sidebar.selectbox("Selecciona Ciclo", ["(Todos)"] + ciclos_disponibles)
        
        cultivos_disponibles = sorted(primera_hoja["Cultivo"].dropna().unique().tolist())
        cultivo_opcion = st.sidebar.selectbox("Selecciona Cultivo", ["(Todos)"] + cultivos_disponibles)

        # --- Aplicar filtros a cada hoja ---
        hojas_filtradas = {}
        for nombre, df in hojas.items():
            df_filtrado = df.copy()
            if entidad_opcion != "(Todos)":
                df_filtrado = df_filtrado[df_filtrado["Entidad"] == entidad_opcion]
            if modalidad_opcion != "(Todos)":
                df_filtrado = df_filtrado[df_filtrado["Modalidad"] == modalidad_opcion]
            if ciclo_opcion != "(Todos)":
                df_filtrado = df_filtrado[df_filtrado["Ciclo"] == ciclo_opcion]
            if cultivo_opcion != "(Todos)":
                df_filtrado = df_filtrado[df_filtrado["Cultivo"] == cultivo_opcion]
            hojas_filtradas[nombre] = df_filtrado

        # --- Mostrar resultados en pantalla ---
        for nombre, df_filtrado in hojas_filtradas.items():
            st.write(f"### 📄 Resultados filtrados - Hoja: {nombre}")
            if not df_filtrado.empty:
                st.dataframe(df_filtrado[columnas_requeridas], use_container_width=True)
            else:
                st.warning(f"La hoja **{nombre}** no tiene datos con los filtros seleccionados.")

        # --- Función para crear el Excel final ---
        def convertir_excel_todas_hojas(hojas_filtradas):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for nombre, df_filtrado in hojas_filtradas.items():
                    if df_filtrado.empty:
                        continue
                    df_filtrado.to_excel(writer, index=False, sheet_name=f"{nombre[:28]}_Filtrado")

                    # Si existe Tarifa2, crear resumen en la misma hoja
                    if "Tarifa2" in df_filtrado.columns and not df_filtrado["Tarifa2"].empty:
                        promedio = df_filtrado["Tarifa2"].mean()
                        desviacion = df_filtrado["Tarifa2"].std()
                        maximo = df_filtrado["Tarifa2"].max()

                        resumen_general = pd.DataFrame({
                            "Métrica": ["Promedio", "Desviación estándar", "Máximo"],
                            "Valor": [promedio, desviacion, maximo]
                        })
                        start_row = len(df_filtrado) + 3
                        resumen_general.to_excel(writer, index=False, sheet_name=f"{nombre[:28]}_Filtrado", startrow=start_row)

                        # Resumen por esquema si existe
                        if "Esquema de aseguramiento" in df_filtrado.columns:
                            resumen_por_esquema = (
                                df_filtrado.groupby("Esquema de aseguramiento")["Tarifa2"]
                                .agg(["mean", "std", "max"])
                                .reset_index()
                                .rename(columns={
                                    "mean": "Promedio",
                                    "std": "Desviación estándar",
                                    "max": "Máximo"
                                })
                            )
                            resumen_por_esquema.to_excel(
                                writer,
                                index=False,
                                sheet_name=f"{nombre[:28]}_Filtrado",
                                startrow=start_row + len(resumen_general) + 3
                            )
            # Añadir gráficos
            output.seek(0)
            wb = load_workbook(output)
            for nombre in wb.sheetnames:
                ws = wb[nombre]
                if "Tarifa2" not in [cell.value for cell in ws[1]]:
                    continue

                # Buscar datos para graficar
                col_tarifa = None
                col_anio = None
                for i, cell in enumerate(ws[1], 1):
                    if cell.value == "Tarifa2":
                        col_tarifa = i
                    if cell.value == "Año":
                        col_anio = i
                if not col_tarifa:
                    continue

                # Gráfico por Año
                if col_anio:
                    max_row = ws.max_row
                    chart = LineChart()
                    chart.title = "Promedio Tarifa2 por Año"
                    chart.y_axis.title = "Promedio"
                    chart.x_axis.title = "Año"
                    data = Reference(ws, min_col=col_tarifa, min_row=2, max_row=max_row)
                    cats = Reference(ws, min_col=col_anio, min_row=2, max_row=max_row)
                    chart.add_data(data, titles_from_data=False)
                    chart.set_categories(cats)
                    ws.add_chart(chart, "L2")

            new_output = BytesIO()
            wb.save(new_output)
            return new_output.getvalue()

        # --- Botón de descarga ---
        st.download_button(
            label="📥 Descargar Excel con hojas filtradas y resúmenes",
            data=convertir_excel_todas_hojas(hojas_filtradas),
            file_name="resultado_todas_las_hojas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error(f"Tu archivo no tiene todas las columnas necesarias: {columnas_requeridas}")

else:
    st.info("Sube un archivo Excel para comenzar.")
