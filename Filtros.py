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

        # --- Calcular estadísticas si existe la columna Tarifa2 ---
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

            # --- Resumen por "Esquema de aseguramiento" ---
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

        # --- Función para crear el archivo Excel con dos hojas ---
        def convertir_excel(df_filtrado, resumen_general, resumen_por_esquema):
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Datos Filtrados')

                # Escribimos el resumen general
                resumen_general.to_excel(writer, index=False, sheet_name='Resumen Tarifa2', startrow=0)

                # Dejamos un espacio y escribimos el resumen por esquema
                startrow = len(resumen_general) + 3
                resumen_por_esquema.to_excel(writer, index=False, sheet_name='Resumen Tarifa2', startrow=startrow)

            return output.getvalue()

        # --- Botón para descargar resultado ---
        st.download_button(
            label="📥 Descargar Excel con resultados",
            data=convertir_excel(
                df_filtrado[columnas_requeridas + ['Tarifa2', 'Esquema de aseguramiento']] 
                if all(col in df_filtrado.columns for col in ['Tarifa2', 'Esquema de aseguramiento']) 
                else df_filtrado,
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
