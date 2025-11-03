import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference

st.set_page_config(page_title="Filtro de CSVs Rápido", layout="wide")
st.title("📊 Filtro de tablas CSV por banderas (Optimizado y rápido)")

# --- Cargar archivos ---
st.write("💡 Puedes subir uno o varios archivos CSV comprimidos en un ZIP o directamente archivos individuales.")
uploaded_files = st.file_uploader("Sube tus archivos CSV o un ZIP con varios:", type=["csv", "zip"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []

    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".zip"):
            with zipfile.ZipFile(uploaded_file) as z:
                for filename in z.namelist():
                    if filename.endswith(".csv"):
                        with z.open(filename) as f:
                            df = pd.read_csv(f, low_memory=False)
                            all_dfs.append(df)
        elif uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, low_memory=False)
            all_dfs.append(df)

    if all_dfs:
        # --- Unir todos los archivos ---
        df = pd.concat(all_dfs, ignore_index=True)
        st.success(f"✅ Archivos cargados correctamente. Total de filas: {len(df):,}")

        # --- Mostrar columnas ---
        st.write("### 🔍 Selecciona las columnas de bandera para filtrar")
        columnas_disponibles = list(df.columns)
        columnas_banderas = st.multiselect("Selecciona las columnas de banderas", columnas_disponibles)

        # --- Aplicar filtros ---
        if columnas_banderas:
            filtros = {}
            for col in columnas_banderas:
                opciones = sorted(df[col].dropna().unique().tolist())
                seleccion = st.multiselect(f"Selecciona valores para '{col}'", opciones)
                if seleccion:
                    filtros[col] = seleccion

            # Aplicar todos los filtros seleccionados
            if filtros:
                mask = pd.Series([True] * len(df))
                for col, vals in filtros.items():
                    mask &= df[col].isin(vals)
                df_filtrado = df[mask]
            else:
                df_filtrado = df

            st.write(f"### 📉 Resultado del filtrado ({len(df_filtrado):,} filas)")
            st.dataframe(df_filtrado, use_container_width=True)

            # --- Descargar resultado ---
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name="Filtrado")
            st.download_button(
                label="📥 Descargar resultado en Excel",
                data=buffer.getvalue(),
                file_name="resultado_filtrado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("Selecciona al menos una columna para aplicar filtros.")
