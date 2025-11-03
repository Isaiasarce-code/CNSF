import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference

st.set_page_config(page_title="Filtro de CSVs Rápido", layout="wide")
st.title("📊 Filtro de tablas CSV por banderas (Optimizado y rápido)")

# --- Cargar archivos ---
st.write("💡 Puedes subir uno o varios archivos CSV dentro de un archivo ZIP (uno por hoja).")

uploaded_file = st.file_uploader("Sube un archivo ZIP con tus CSVs", type=["zip"])

if uploaded_file:
    # Leer el ZIP sin extraer a disco
    with zipfile.ZipFile(uploaded_file) as z:
        lista_csv = [f for f in z.namelist() if f.endswith(".csv")]
        
        if not lista_csv:
            st.error("El archivo ZIP no contiene archivos CSV.")
        else:
            # Leer solo el primero para construir filtros
            with z.open(lista_csv[0]) as f:
                df_base = pd.read_csv(f, low_memory=False)

            columnas_requeridas = ["Entidad", "Modalidad", "Ciclo", "Cultivo"]

            if all(col in df_base.columns for col in columnas_requeridas):
                # --- Filtros dinámicos ---
                st.sidebar.header("Filtros")
                entidad_opcion = st.sidebar.selectbox("Entidad", ["(Todos)"] + sorted(df_base["Entidad"].dropna().unique().tolist()))
                modalidad_opcion = st.sidebar.selectbox("Modalidad", ["(Todos)"] + sorted(df_base["Modalidad"].dropna().unique().tolist()))
                ciclo_opcion = st.sidebar.selectbox("Ciclo", ["(Todos)"] + sorted(df_base["Ciclo"].dropna().unique().tolist()))
                cultivo_opcion = st.sidebar.selectbox("Cultivo", ["(Todos)"] + sorted(df_base["Cultivo"].dropna().unique().tolist()))
                generar_graficos = st.sidebar.checkbox("Generar gráficos (puede tardar más)", value=False)

                # --- Crear libro Excel en memoria ---
                wb = Workbook()
                wb.remove(wb.active)

                for nombre_csv in lista_csv:
                    with z.open(nombre_csv) as f:
                        df = pd.read_csv(f, low_memory=False)
                        if not all(col in df.columns for col in columnas_requeridas):
                            continue

                        # Filtro rápido usando máscara booleana
                        mask = pd.Series(True, index=df.index)
                        if entidad_opcion != "(Todos)":
                            mask &= df["Entidad"] == entidad_opcion
                        if modalidad_opcion != "(Todos)":
                            mask &= df["Modalidad"] == modalidad_opcion
                        if ciclo_opcion != "(Todos)":
                            mask &= df["Ciclo"] == ciclo_opcion
                        if cultivo_opcion != "(Todos)":
                            mask &= df["Cultivo"] == cultivo_opcion
                        df_filtrado = df[mask]

                        if df_filtrado.empty:
                            continue

                        # Crear hoja nueva
                        ws = wb.create_sheet(title=nombre_csv.replace(".csv", "")[:28])

                        # Escribir datos rápido
                        for r_idx, row in enumerate([df_filtrado.columns.tolist()] + df_filtrado.values.tolist(), start=1):
                            ws.append(row)

                        # Calcular estadísticas si hay Tarifa2
                        if "Tarifa2" in df_filtrado.columns and df_filtrado["Tarifa2"].notna().any():
                            prom = df_filtrado["Tarifa2"].mean()
                            desv = df_filtrado["Tarifa2"].std()
                            maxv = df_filtrado["Tarifa2"].max()
                            start_row = len(df_filtrado) + 3
                            ws.cell(row=start_row, column=1, value="Promedio")
                            ws.cell(row=start_row, column=2, value=prom)
                            ws.cell(row=start_row + 1, column=1, value="Desviación estándar")
                            ws.cell(row=start_row + 1, column=2, value=desv)
                            ws.cell(row=start_row + 2, column=1, value="Máximo")
                            ws.cell(row=start_row + 2, column=2, value=maxv)

                            # --- Gráfico (opcional) ---
                            if generar_graficos and "Año" in df_filtrado.columns:
                                df_por_anio = df_filtrado.groupby("Año")["Tarifa2"].mean().reset_index()
                                data_row = start_row + 5
                                for r in df_por_anio.itertuples(index=False):
                                    ws.append(r)

                                chart = LineChart()
                                chart.title = "Promedio Tarifa2 por Año"
                                chart.y_axis.title = "Promedio"
                                chart.x_axis.title = "Año"

                                data = Reference(ws, min_col=2, min_row=data_row + 1, max_row=data_row + len(df_por_anio))
                                cats = Reference(ws, min_col=1, min_row=data_row + 1, max_row=data_row + len(df_por_anio))
                                chart.add_data(data, titles_from_data=False)
                                chart.set_categories(cats)
                                ws.add_chart(chart, f"E{data_row}")

                # --- Guardar en memoria ---
                output = BytesIO()
                wb.save(output)
                output.seek(0)

                st.download_button(
                    "📥 Descargar Excel filtrado y optimizado",
                    data=output,
                    file_name="resultado_filtrado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.error(f"Faltan columnas requeridas: {columnas_requeridas}")

else:
    st.info("Sube un archivo ZIP con tus CSVs para comenzar.")
