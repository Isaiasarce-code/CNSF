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

# --- Cargar ZIP ---
uploaded_zip = st.file_uploader("Sube un archivo ZIP con varios CSVs", type=["zip"])

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip) as z:
        csv_files = [f for f in z.namelist() if f.endswith(".csv")]

        if not csv_files:
            st.error("El ZIP no contiene archivos CSV.")
        else:
            st.success(f"Se encontraron {len(csv_files)} archivos CSV en el ZIP.")

            # --- Cargar todos los datos temporalmente para obtener valores únicos ---
            dfs = []
            for file_name in csv_files:
                with z.open(file_name) as f:
                    try:
                        df = pd.read_csv(TextIOWrapper(f, encoding=encoding_opcion))
                        df.columns = df.columns.str.strip()
                        if all(col in df.columns for col in ["Entidad", "Modalidad", "Ciclo", "Cultivo"]):
                            dfs.append(df)
                    except Exception:
                        continue

            if not dfs:
                st.error("Ningún archivo tiene todas las columnas requeridas (Entidad, Modalidad, Ciclo, Cultivo).")
            else:
                df_total = pd.concat(dfs, ignore_index=True)

                # --- Filtros dinámicos ---
                st.sidebar.subheader("Filtros dinámicos")

                entidad_sel = st.sidebar.multiselect(
                    "Entidad", sorted(df_total["Entidad"].dropna().unique().tolist())
                )
                modalidad_sel = st.sidebar.multiselect(
                    "Modalidad", sorted(df_total["Modalidad"].dropna().unique().tolist())
                )
                ciclo_sel = st.sidebar.multiselect(
                    "Ciclo", sorted(df_total["Ciclo"].dropna().unique().tolist())
                )
                cultivo_sel = st.sidebar.multiselect(
                    "Cultivo", sorted(df_total["Cultivo"].dropna().unique().tolist())
                )

                # --- Crear libro Excel ---
                wb = Workbook()
                wb.remove(wb.active)  # Quitar hoja por defecto

                for file_name in csv_files:
                    with z.open(file_name) as f:
                        try:
                            df = pd.read_csv(TextIOWrapper(f, encoding=encoding_opcion))
                            df.columns = df.columns.str.strip()

                            if not all(col in df.columns for col in ["Entidad", "Modalidad", "Ciclo", "Cultivo"]):
                                st.warning(f"{file_name} no tiene todas las columnas requeridas.")
                                continue

                            df_filtrado = df.copy()

                            # --- Aplicar filtros dinámicos ---
                            if entidad_sel:
                                df_filtrado = df_filtrado[df_filtrado["Entidad"].isin(entidad_sel)]
                            if modalidad_sel:
                                df_filtrado = df_filtrado[df_filtrado["Modalidad"].isin(modalidad_sel)]
                            if ciclo_sel:
                                df_filtrado = df_filtrado[df_filtrado["Ciclo"].isin(ciclo_sel)]
                            if cultivo_sel:
                                df_filtrado = df_filtrado[df_filtrado["Cultivo"].isin(cultivo_sel)]

                            # --- Agregar hoja al Excel ---
                            ws = wb.create_sheet(title=file_name.replace(".csv", "")[:31])
                            ws.append(list(df_filtrado.columns))
                            for r in df_filtrado.itertuples(index=False):
                                ws.append(list(r))

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
