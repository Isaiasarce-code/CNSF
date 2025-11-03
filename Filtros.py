import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO, TextIOWrapper
from openpyxl import Workbook

st.set_page_config(page_title="Filtro múltiple ZIP", layout="wide")
st.title("📦 Filtro de múltiples CSV desde ZIP")

# --- Configuración de carga ---
st.sidebar.subheader("Configuración de carga")
encoding_opcion = st.sidebar.selectbox(
    "Codificación de los archivos",
    ["utf-8", "latin1", "cp1252", "ISO-8859-1"],
    index=1
)

# --- Filtros fijos ---
st.sidebar.subheader("Filtros aplicados a todos los archivos")
entidad = st.sidebar.text_input("Entidad (dejar vacío para todos)")
modalidad = st.sidebar.text_input("Modalidad")
ciclo = st.sidebar.text_input("Ciclo")
cultivo = st.sidebar.text_input("Cultivo")

# --- Cargar ZIP ---
uploaded_zip = st.file_uploader("Sube un archivo ZIP con varios CSVs", type=["zip"])

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip) as z:
        csv_files = [f for f in z.namelist() if f.endswith(".csv")]

        if not csv_files:
            st.error("El ZIP no contiene archivos CSV.")
        else:
            st.success(f"Se encontraron {len(csv_files)} archivos CSV en el ZIP.")

            # --- Crear libro Excel ---
            wb = Workbook()
            wb.remove(wb.active)  # Quitar hoja por defecto

            for file_name in csv_files:
                with z.open(file_name) as f:
                    try:
                        df = pd.read_csv(TextIOWrapper(f, encoding=encoding_opcion))
                        df.columns = df.columns.str.strip()

                        columnas_requeridas = ["Entidad", "Modalidad", "Ciclo", "Cultivo"]
                        if not all(col in df.columns for col in columnas_requeridas):
                            st.warning(f"{file_name} no tiene todas las columnas requeridas.")
                            continue

                        # --- Aplicar filtros fijos ---
                        df_filtrado = df.copy()
                        if entidad:
                            df_filtrado = df_filtrado[df_filtrado["Entidad"] == entidad]
                        if modalidad:
                            df_filtrado = df_filtrado[df_filtrado["Modalidad"] == modalidad]
                        if ciclo:
                            df_filtrado = df_filtrado[df_filtrado["Ciclo"] == ciclo]
                        if cultivo:
                            df_filtrado = df_filtrado[df_filtrado["Cultivo"] == cultivo]

                        # --- Agregar hoja al Excel ---
                        ws = wb.create_sheet(title=file_name.replace(".csv", "")[:31])
                        for r in df_filtrado.itertuples(index=False):
                            ws.append(list(r))
                        ws.insert_rows(1)
                        ws.append(list(df_filtrado.columns))

                    except Exception as e:
                        st.error(f"Error procesando {file_name}: {e}")

            # --- Descargar Excel ---
            output = BytesIO()
            wb.save(output)
            st.download_button(
                label="📥 Descargar Excel con hojas filtradas",
                data=output.getvalue(),
                file_name="filtrado_multiples_archivos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Sube un archivo ZIP para comenzar.")
