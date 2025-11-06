import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO, TextIOWrapper
from openpyxl import Workbook

st.set_page_config(page_title="Filtro ZIP por etapas", layout="wide")
st.title("📦 Filtro de múltiples CSV desde ZIP (Primero selecciona filtros)")

# --- Paso 1: Seleccionar filtros antes de subir archivos ---
st.sidebar.header("🎯 Selecciona los filtros antes de subir")

entidad_sel = st.sidebar.multiselect(
    "Entidad (opcional)", ["Aguascalientes", "Yucatán", "Chiapas", "CDMX", "Jalisco"], placeholder="Selecciona..."
)
modalidad_sel = st.sidebar.multiselect(
    "Modalidad (opcional)", ["Riego", "Temporal"], placeholder="Selecciona..."
)
ciclo_sel = st.sidebar.multiselect(
    "Ciclo (opcional)", ["Otoño-Invierno", "Primavera-Verano"], placeholder="Selecciona..."
)
cultivo_sel = st.sidebar.multiselect(
    "Cultivo (opcional)", ["Maíz", "Sorgo", "Caña", "Trigo", "Soja"], placeholder="Selecciona..."
)

st.sidebar.info("💡 Primero elige tus filtros, luego sube el ZIP para aplicar.")

# --- Paso 2: Subir ZIP ---
uploaded_zip = st.file_uploader("Sube un archivo ZIP con varios CSV", type=["zip"])

if not uploaded_zip:
    st.info("Sube tu archivo ZIP para aplicar los filtros seleccionados.")
    st.stop()

# --- Paso 3: Procesar archivos uno por uno ---
encoding_opcion = "latin1"
wb = Workbook()
wb.remove(wb.active)

try:
    with zipfile.ZipFile(uploaded_zip) as z:
        csv_files = [f for f in z.namelist() if f.endswith(".csv")]
        if not csv_files:
            st.error("El ZIP no contiene archivos CSV.")
            st.stop()

        procesados = 0
        for file_name in csv_files:
            try:
                with z.open(file_name) as f:
                    df = pd.read_csv(TextIOWrapper(f, encoding=encoding_opcion))
                    df.columns = df.columns.str.strip()

                    if not all(c in df.columns for c in ["Entidad", "Modalidad", "Ciclo", "Cultivo"]):
                        continue

                    # --- Aplicar filtros seleccionados ---
                    if entidad_sel:
                        df = df[df["Entidad"].isin(entidad_sel)]
                    if modalidad_sel:
                        df = df[df["Modalidad"].isin(modalidad_sel)]
                    if ciclo_sel:
                        df = df[df["Ciclo"].isin(ciclo_sel)]
                    if cultivo_sel:
                        df = df[df["Cultivo"].isin(cultivo_sel)]

                    if df.empty:
                        continue

                    ws = wb.create_sheet(title=file_name.replace(".csv", "")[:31])
                    ws.append(list(df.columns))
                    for row in df.itertuples(index=False):
                        ws.append(row)
                    procesados += 1

            except Exception as e:
                st.warning(f"⚠️ Error procesando {file_name}: {e}")

        if procesados == 0:
            st.warning("Ningún archivo cumplió con los filtros seleccionados.")
            st.stop()

        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="📥 Descargar Excel filtrado",
            data=output.getvalue(),
            file_name="filtrado_resultados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success(f"✅ {procesados} archivos procesados correctamente.")
except Exception as e:
    st.error(f"Error al procesar el ZIP: {e}")
