import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows 

st.set_page_config(page_title="Filtro de Datos", layout="wide")

st.title("📊 Filtro de tabla por banderas")

# --- Cargar archivo Excel ---
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    columnas_requeridas = ["Entidad", "Modalidad", "Ciclo", "Cultivo"]
    if all(col in df.columns for col in columnas_requeridas):

        # --- Filtros ---
        st.sidebar.header("Filtros")
        
        entidad_opcion = st.sidebar.selectbox(
            "Selecciona Entidad",
            options=["(Todos)"] + sorted(df["Entidad"].dropna().unique().tolist())
        )
        
        modalidad_opcion = st.sidebar.selectbox(
            "Selecciona Modalidad",
            options=["(Todos)"] + sorted(df["Modalidad"].dropna().unique().tolist())
        )
        
        ciclo_opcion = st.sidebar.selectbox(
            "Selecciona Ciclo",
            options=["(Todos)"] + sorted(df["Ciclo"].dropna().unique().tolist())
        )
        
        cultivo_opcion = st.sidebar.selectbox(
            "Selecciona Cultivo",
            options=["(Todos)"] + sorted(df["Cultivo"].dropna().unique().tolist())
        )
        
        # --- Aplicar filtros ---
        df_filtrado = df.copy()
        
        if entidad_opcion != "(Todos)":
            df_filtrado = df_filtrado[df_filtrado["Entidad"] == entidad_opcion]
        if modalidad_opcion != "(Todos)":
            df_filtrado = df_filtrado[df_filtrado["Modalidad"] == modalidad_opcion]
        if ciclo_opcion != "(Todos)":
            df_filtrado = df_filtrado[df_filtrado["Ciclo"] == ciclo_opcion]
        if cultivo_opcion != "(Todos)":
            df_filtrado = df_filtrado[df_filtrado["Cultivo"] == cultivo_opcion]

        # --- Mostrar resultado ---
        st.write("### Resultado filtrado")
        st.dataframe(df_filtrado[columnas_requeridas], use_container_width=True)

        # --- Calcular estadísticas ---
        if "Tarifa2" in df_filtrado.columns and not df_filtrado["Tarifa2"].empty:
            promedio = df_filtrado["Tarifa2"].mean()
            desviacion = df_filtrado["Tarifa2"].std()
            maximo = df_filtrado["Tarifa2"].max()

            resumen_general = pd.DataFrame({
                "Métrica": ["Promedio", "Desviación estándar", "Máximo"],
                "Valor": [promedio, desviacion, maximo]
            })

            st.write("### 📈 Resumen general (Tarifa2)")
            st.dataframe(resumen_general, use_container_width=True)

            # --- Resumen por Esquema ---
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
                st.write("### 📊 Resumen por Esquema de aseguramiento")
                st.dataframe(resumen_por_esquema, use_container_width=True)
            else:
                resumen_por_esquema = pd.DataFrame(columns=["Esquema de aseguramiento", "Promedio", "Desviación estándar", "Máximo"])
                st.warning("No se encontró la columna 'Esquema de aseguramiento'.")

        else:
            resumen_general = pd.DataFrame({
                "Métrica": ["Promedio", "Desviación estándar", "Máximo"],
                "Valor": ["N/A", "N/A", "N/A"]
            })
            resumen_por_esquema = pd.DataFrame(columns=["Esquema de aseguramiento", "Promedio", "Desviación estándar", "Máximo"])
            st.warning("No se encontró la columna 'Tarifa2' o no tiene datos numéricos.")

        # --- Función para generar Excel con gráficos ---
        def convertir_excel(df_filtrado, resumen_general, resumen_por_esquema):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Datos Filtrados')

                # Escribimos las tablas
                resumen_general.to_excel(writer, index=False, sheet_name='Resumen Tarifa2', startrow=0)
                startrow = len(resumen_general) + 3
                resumen_por_esquema.to_excel(writer, index=False, sheet_name='Resumen Tarifa2', startrow=startrow)

            # Cargar libro para añadir gráficos
            output.seek(0)
            wb = load_workbook(output)
            ws = wb["Resumen Tarifa2"]

            # --- Gráfico 1: Promedio general ---
            if "Año" in df_filtrado.columns:
                df_por_anio = df_filtrado.groupby("Año")["Tarifa2"].mean().reset_index()
                chart_data_row = startrow + len(resumen_por_esquema) + 2
                for r in dataframe_to_rows(df_por_anio, index=False, header=True):
                    ws.append(r)

                chart1 = LineChart()
                chart1.title = "Promedio por Año"
                chart1.style = 10  # estilo predeterminado con líneas suaves
                chart1.y_axis.title = "Valor Promedio"
                chart1.x_axis.title = "Año"
                chart1.marker = True  # para mostrar puntos en cada año

                data_ref = Reference(ws, min_col=2, min_row=chart_data_row+1, max_row=chart_data_row+len(df_por_anio))
                cats_ref = Reference(ws, min_col=1, min_row=chart_data_row+1, max_row=chart_data_row+len(df_por_anio))
                chart1.add_data(data_ref, titles_from_data=False)
                chart1.set_categories(cats_ref)
                ws.add_chart(chart1, f"E{chart_data_row}")

            # --- Gráfico 2: Promedio por Esquema ---
            if not resumen_por_esquema.empty:
                chart2 = LineChart()
                chart2.title = "Promedio Tarifa2 por Esquema de aseguramiento"
                chart2.style = 10  # estilo predeterminado con líneas suaves
                chart2.y_axis.title = "Promedio"
                chart2.x_axis.title = "Esquema"
                chart2.marker = True  # para mostrar puntos en cada año
        
                data_ref2 = Reference(ws, min_col=2, min_row=startrow+2, max_row=startrow+1+len(resumen_por_esquema))
                cats_ref2 = Reference(ws, min_col=1, min_row=startrow+2, max_row=startrow+1+len(resumen_por_esquema))
                chart2.add_data(data_ref2, titles_from_data=False)
                chart2.set_categories(cats_ref2)
                ws.add_chart(chart2, "E10")

            # Guardar
            new_output = BytesIO()
            wb.save(new_output)
            return new_output.getvalue()

        # --- Botón de descarga ---
        st.download_button(
            label="📥 Descargar Excel con tablas y gráficos",
            data=convertir_excel(
                df_filtrado,
                resumen_general,
                resumen_por_esquema
            ),
            file_name="resultado_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error(f"Tu archivo no tiene todas las columnas necesarias: {columnas_requeridas}")
else:
    st.info("Sube un archivo Excel para comenzar.")
