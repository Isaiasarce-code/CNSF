import streamlit as st
import pandas as pd

st.set_page_config(page_title="Filtro Dinámico CSV", layout="wide")
st.title("📋 Filtro dinámico de datos CSV")

# --- Cargar archivo CSV ---
uploaded_file = st.file_uploader("Sube tu archivo CSV (delimitado por comas)", type=["csv"])

if uploaded_file:
   #df = pd.read_csv(uploaded_file)
    df = pd.read_csv(uploaded_file, encoding="latin1")

    df.columns = df.columns.str.strip()  # Limpia espacios en nombres de columnas

    columnas_requeridas = ["Entidad", "Modalidad", "Ciclo", "Cultivo"]
    if all(col in df.columns for col in columnas_requeridas):

        # --- Filtros dinámicos ---
        st.sidebar.header("Filtros dinámicos")

        # Filtro 1: Entidad
        entidades = sorted(df["Entidad"].dropna().unique())
        entidad_opcion = st.sidebar.selectbox("Entidad", ["(Todos)"] + entidades)
        df_filtrado = df if entidad_opcion == "(Todos)" else df[df["Entidad"] == entidad_opcion]

        # Filtro 2: Modalidad
        modalidades = sorted(df_filtrado["Modalidad"].dropna().unique())
        modalidad_opcion = st.sidebar.selectbox("Modalidad", ["(Todos)"] + modalidades)
        df_filtrado = df_filtrado if modalidad_opcion == "(Todos)" else df_filtrado[df_filtrado["Modalidad"] == modalidad_opcion]

        # Filtro 3: Ciclo
        ciclos = sorted(df_filtrado["Ciclo"].dropna().unique())
        ciclo_opcion = st.sidebar.selectbox("Ciclo", ["(Todos)"] + ciclos)
        df_filtrado = df_filtrado if ciclo_opcion == "(Todos)" else df_filtrado[df_filtrado["Ciclo"] == ciclo_opcion]

        # Filtro 4: Cultivo
        cultivos = sorted(df_filtrado["Cultivo"].dropna().unique())
        cultivo_opcion = st.sidebar.selectbox("Cultivo", ["(Todos)"] + cultivos)
        df_filtrado = df_filtrado if cultivo_opcion == "(Todos)" else df_filtrado[df_filtrado["Cultivo"] == cultivo_opcion]

        # --- Mostrar resultado ---
        st.write("### 🔍 Datos filtrados")
        st.dataframe(df_filtrado, use_container_width=True)

        # --- Botón de descarga ---
        csv_filtrado = df_filtrado.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Descargar CSV filtrado",
            data=csv_filtrado,
            file_name="datos_filtrados.csv",
            mime="text/csv"
        )

    else:
        st.error(f"Tu archivo debe contener las columnas: {columnas_requeridas}")
else:
    st.info("Sube un archivo CSV para comenzar.")
