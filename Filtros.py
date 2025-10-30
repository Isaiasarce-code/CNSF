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
