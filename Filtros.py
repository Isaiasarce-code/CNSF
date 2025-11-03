import streamlit as st
import pandas as pd

st.set_page_config(page_title="Filtro Dinámico CSV", layout="wide")
st.title("📋 Filtro dinámico de datos CSV")

# --- Opciones de carga ---
st.sidebar.subheader("Configuración de carga")
encoding_opcion = st.sidebar.selectbox(
    "Codificación del archivo",
    ["utf-8", "latin1", "cp1252", "ISO-8859-1"],
    index=1  # latin1 como predeterminado
)

# --- Cargar archivo CSV con caché ---
@st.cache_data
def cargar_csv(file, encoding):
    return pd.read_csv(file, encoding=encoding)

uploaded_file = st.file_uploader("Sube tu archivo CSV (delimitado por comas)", type=["csv"])

if uploaded_file:
    try:
        df = cargar_csv(uploaded_file, encoding_opcion)
        df.columns = df.columns.str.strip()  # Limpia espacios en nombres de columnas

        columnas_requeridas = ["Entidad", "Modalidad", "Ciclo", "Cultivo"]
        if all(col in df.columns for col in columnas_requeridas):

            # --- Filtros dinámicos segmentados ---
            st.sidebar.header("Filtros dinámicos")

            entidad_opcion = st.sidebar.selectbox(
                "Entidad", ["(Todos)"] + sorted(df["Entidad"].dropna().unique())
            )
            df_entidad = df if entidad_opcion == "(Todos)" else df[df["Entidad"] == entidad_opcion]

            modalidad_opcion = st.sidebar.selectbox(
                "Modalidad", ["(Todos)"] + sorted(df_entidad["Modalidad"].dropna().unique())
            )
            df_modalidad = df_entidad if modalidad_opcion == "(Todos)" else df_entidad[df_entidad["Modalidad"] == modalidad_opcion]

            ciclo_opcion = st.sidebar.selectbox(
                "Ciclo", ["(Todos)"] + sorted(df_modalidad["Ciclo"].dropna().unique())
            )
            df_ciclo = df_modalidad if ciclo_opcion == "(Todos)" else df_modalidad[df_modalidad["Ciclo"] == ciclo_opcion]

            cultivo_opcion = st.sidebar.selectbox(
                "Cultivo", ["(Todos)"] + sorted(df_ciclo["Cultivo"].dropna().unique())
            )
            df_filtrado = df_ciclo if cultivo_opcion == "(Todos)" else df_ciclo[df_ciclo["Cultivo"] == cultivo_opcion]

            # --- Mostrar resultado ---
            st.write("### 🔍 Datos filtrados")
            if len(df_filtrado) > 5000:
                st.warning(f"Demasiadas filas para mostrar ({len(df_filtrado)}). Se muestran solo las primeras 500.")
                st.dataframe(df_filtrado.head(500), use_container_width=True)
            else:
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

    except UnicodeDecodeError:
        st.error(f"No se pudo leer el archivo con la codificación '{encoding_opcion}'. Prueba con otra.")
        st.stop()
else:
    st.info("Sube un archivo CSV para comenzar.")
