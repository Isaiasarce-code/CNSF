import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(page_title="Filtro múltiple de CSVs", layout="wide")
st.title("📊 Filtro de 4 hojas CSV con filtros comunes")

# --- Subir archivos ---
st.write("💡 Sube exactamente 4 archivos CSV, cada uno representando una hoja de Excel.")
uploaded_files = st.file_uploader("Sube tus 4 archivos CSV:", type=["csv"], accept_multiple_files=True)

# --- Validación de cantidad ---
if uploaded_files and len(uploaded_files) == 4:
    hojas = {}
    for i, file in enumerate(uploaded_files, start=1):
        try:
            df = pd.read_csv(file, encoding="latin1")
            hojas[f"Hoja{i}"] = df
        except Exception as e:
            st.error(f"Error al leer el archivo {file.name}: {e}")

    # --- Verificar columnas requeridas ---
    columnas_requeridas = ["Cultivo", "Modalidad", "Ciclo", "Entidad"]
    if all(all(col in df.columns for col in columnas_requeridas) for df in hojas.values()):
        st.sidebar.header("🎯 Filtros comunes")

        # --- Filtros dinámicos basados en la primera hoja ---
        base_df = list(hojas.values())[0]

        cultivo_opcion = st.sidebar.multiselect("Cultivo", sorted(base_df["Cultivo"].dropna().unique()))
        modalidad_opcion = st.sidebar.multiselect("Modalidad", sorted(base_df["Modalidad"].dropna().unique()))
        ciclo_opcion = st.sidebar.multiselect("Ciclo", sorted(base_df["Ciclo"].dropna().unique()))
        entidad_opcion = st.sidebar.multiselect("Entidad", sorted(base_df["Entidad"].dropna().unique()))

        # --- Función de filtrado ---
        def aplicar_filtros(df):
            if cultivo_opcion:
                df = df[df["Cultivo"].isin(cultivo_opcion)]
            if modalidad_opcion:
                df = df[df["Modalidad"].isin(modalidad_opcion)]
            if ciclo_opcion:
                df = df[df["Ciclo"].isin(ciclo_opcion)]
            if entidad_opcion:
                df = df[df["Entidad"].isin(entidad_opcion)]
            return df

        # --- Aplicar filtros a cada hoja ---
        hojas_filtradas = {nombre: aplicar_filtros(df) for nombre, df in hojas.items()}

        # --- Mostrar resultados ---
        for nombre, df in hojas_filtradas.items():
            st.write(f"### 📄 {nombre} ({len(df):,} filas)")
            st.dataframe(df, use_container_width=True)

        # --- Generar Excel con 4 hojas ---
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            for nombre, df in hojas_filtradas.items():
                df.to_excel(writer, index=False, sheet_name=nombre)

        st.download_button(
            label="📥 Descargar Excel con hojas filtradas",
            data=buffer.getvalue(),
            file_name="hojas_filtradas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error(f"Todos los archivos deben contener las columnas: {columnas_requeridas}")
elif uploaded_files:
    st.warning("Por favor sube exactamente 4 archivos CSV.")
else:
    st.info("Sube tus archivos para comenzar.")
