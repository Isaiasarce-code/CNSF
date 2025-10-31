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

        # --- Filtros dinámicos ---
        st.sidebar.header("Filtros")
        
        # Filtro 1: Entidad
        entidades_disponibles = sorted(df["Entidad"].dropna().unique().tolist())
        entidad_opcion = st.sidebar.selectbox("Selecciona Entidad", ["(Todos)"] + entidades_disponibles)
        
        df_entidad = df if entidad_opcion == "(Todos)" else df[df["Entidad"] == entidad_opcion]
        
        # Filtro 2: Modalidad
        modalidades_disponibles = sorted(df_entidad["Modalidad"].dropna().unique().tolist())
        modalidad_opcion = st.sidebar.selectbox("Selecciona Modalidad", ["(Todos)"] + modalidades_disponibles)
        
        df_modalidad = df_entidad if modalidad_opcion == "(Todos)" else df_entidad[df_entidad["Modalidad"] == modalidad_opcion]
        
        # Filtro 3: Ciclo
        ciclos_disponibles = sorted(df_modalidad["Ciclo"].dropna().unique().tolist())
        ciclo_opcion = st.sidebar.selectbox("Selecciona Ciclo", ["(Todos)"] + ciclos_disponibles)
        
        df_ciclo = df_modalidad if ciclo_opcion == "(Todos)" else df_modalidad[df_modalidad["Ciclo"] == ciclo_opcion]
        
        # Filtro 4: Cultivo
        cultivos_disponibles = sorted(df_ciclo["Cultivo"].dropna().unique().tolist())
        cultivo_opcion = st.sidebar.selectbox("Selecciona Cultivo", ["(Todos)"] + cultivos_disponibles)
        
        # --- Aplicar filtros finales ---
        df_filtrado = df_ciclo if cultivo_opcion == "(Todos)" else df_ciclo[df_ciclo["Cultivo"] == cultivo_opcion]
        
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
            # --- Gráfico 1: Promedio por Año ---
            if "Año" in df_filtrado.columns:
                df_por_anio = df_filtrado.groupby("Año")["Tarifa2"].mean().reset_index()
                df_por_anio["Tarifa2"] = df_por_anio["Tarifa2"] * 100
                df_por_anio.columns = ["Año", "Valor Promedio"]
                chart_data_row = startrow + len(resumen_por_esquema) + 2
                for r in dataframe_to_rows(df_por_anio, index=False, header=True):
                    ws.append(r)
            
                chart1 = LineChart()
                chart1.title = "Promedio por Año"
                chart1.style = 2
                chart1.y_axis.title = "Valor Promedio"
                chart1.x_axis.title = "Año"
                chart1.marker = True
                chart1.legend = None
                from openpyxl.drawing.fill import SolidFillProperties
                chart1.graphicalProperties.line.solidFill = "0070C0"
            
                data_ref = Reference(ws, min_col=2, min_row=chart_data_row+1, max_row=chart_data_row+len(df_por_anio))
                cats_ref = Reference(ws, min_col=1, min_row=chart_data_row+1, max_row=chart_data_row+len(df_por_anio))
                
                chart1.add_data(data_ref, titles_from_data=True)
                chart1.set_categories(cats_ref)
            
                from openpyxl.chart.label import DataLabelList
                chart1.dataLabels = DataLabelList()
                chart1.dataLabels.showVal = True
                chart1.dataLabels.numberFormat = "0.00%"
            
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
