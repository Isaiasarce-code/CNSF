import streamlit as st
import pandas as pd

st.set_page_config(page_title="Filtro de Datos", layout="wide")

st.title("📊 Filtro de tabla por banderas")

# --- Cargar archivo Excel ---
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Verificar que las columnas existan
    columnas_requeridas = ["Entidad", "Modalidad", "Ciclo", "Cultivo"]
    if all(col in df.columns for col in columnas_requeridas):

        # --- Filtros ---
        st.sidebar.header("Filtros")

        entidad_opciones = st.sidebar.multiselect(
            "Selecciona Entidad",
            options=sorted(df["Entidad"].dropna().unique())
        )

        modalidad_opciones = st.sidebar.multiselect(
            "Selecciona Modalidad",
            options=sorted(df["Modalidad"].dropna().unique())
        )

        ciclo_opciones = st.sidebar.multiselect(
            "Selecciona Ciclo",
            options=sorted(df["Ciclo"].dropna().unique())
        )

        cultivo_opciones = st.sidebar.multiselect(
            "Selecciona Cultivo",
            options=sorted(df["Cultivo"].dropna().unique())
        )

        # --- Aplicar filtros ---
        df_filtrado = df.copy()

        if entidad_opciones:
            df_filtrado = df_filtrado[df_filtrado["Entidad"].isin(entidad_opciones)]

        if modalidad_opciones:
            df_filtrado = df_filtrado[df_filtrado["Modalidad"].isin(modalidad_opciones)]

        if ciclo_opciones:
            df_filtrado = df_filtrado[df_filtrado["Ciclo"].isin(ciclo_opciones)]

        if cultivo_opciones:
            df_filtrado = df_filtrado[df_filtrado["Cultivo"].isin(cultivo_opciones)]

        # --- Mostrar resultado ---
        st.write("### Resultado filtrado")
        st.dataframe(df_filtrado[columnas_requeridas], use_container_width=True)

        # --- Botón para descargar resultado ---
        def convertir_excel(df):
            from io import BytesIO
            output = BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            return output.getvalue()

        st.download_button(
            label="📥 Descargar resultado filtrado en Excel",
            data=convertir_excel(df_filtrado[columnas_requeridas]),
            file_name="resultado_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error(f"Tu archivo no tiene todas las columnas necesarias: {columnas_requeridas}")
else:
    st.info("Sube un archivo Excel para comenzar.")
