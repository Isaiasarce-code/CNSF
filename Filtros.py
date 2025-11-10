import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
from openpyxl import Workbook

st.set_page_config(page_title="Filtro ZIP de Parquet", layout="wide")
st.title("📦 Filtro de múltiples PARQUET desde ZIP (Primero selecciona filtros)")

# --- Paso 1: Seleccionar filtros antes de subir archivos ---
st.sidebar.header("🎯 Selecciona los filtros antes de subir")

entidad_sel = st.sidebar.selectbox(
    "Entidad (opcional)",
    ["", "Aguascalientes", "Yucatán", "Chiapas", "CDMX", "Jalisco"]
)

modalidad_sel = st.sidebar.selectbox(
    "Modalidad (opcional)",
    ["", "Riego", "Temporal"]
)

ciclo_sel = st.sidebar.selectbox(
    "Ciclo (opcional)",
    ["", "Otoño-Invierno", "Primavera-Verano"]
)

cultivo_sel = st.sidebar.selectbox(
    "Cultivo (opcional)",
    ["", "Maíz", "Sorgo", "Caña", "Trigo", "Soja", "Ajo"]
)

st.sidebar.info("💡 Primero elige tus filtros, luego sube el ZIP con archivos PARQUET.")

# --- Paso 2: Subir ZIP ---
uploaded_zip = st.file_uploader("Sube un archivo ZIP con varios PARQUET", type=["zip"])

if not uploaded_zip:
    st.info("Sube tu archivo ZIP para aplicar los filtros seleccionados.")
    st.stop()

# --- Paso 3: Procesar archivos uno por uno ---
wb = Workbook()
wb.remove(wb.active)

try:
    with zipfile.ZipFile(uploaded_zip) as z:
        parquet_files = [f for f in z.namelist() if f.endswith(".parquet")]
        if not parquet_files:
            st.error("El ZIP no contiene archivos PARQUET.")
            st.stop()

        procesados = 0
        for file_name in parquet_files:
            try:
                with z.open(file_name) as f:
                    df = pd.read_parquet(f)
                    df.columns = df.columns.str.strip()

                    if not all(c in df.columns for c in ["Entidad", "Modalidad", "Ciclo", "Cultivo"]):
                        continue

                    # --- Aplicar filtros seleccionados ---
                    if entidad_sel:
                        df = df[df["Entidad"] == entidad_sel]
                    if modalidad_sel:
                        df = df[df["Modalidad"] == modalidad_sel]
                    if ciclo_sel:
                        df = df[df["Ciclo"] == ciclo_sel]
                    if cultivo_sel:
                        df = df[df["Cultivo"] == cultivo_sel]

                    if df.empty:
                        continue

                    ws = wb.create_sheet(title=file_name.replace(".parquet", "")[:31])
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
