import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
import zipfile

# ---------- CONFIGURACIÓN GENERAL ----------
st.set_page_config(
    page_title="Generador de Oficios",
    page_icon="📄",
    layout="wide"
)

# ---------- ESTILO PERSONALIZADO ----------
st.markdown("""
<style>
    /* Fondo general de la app */
    .stApp {
        background-color: #f6f8fb; /* gris-azulado muy suave */
        color: #1f3b57;
    }

    /* Quitar fondo azul de los elementos */
    [data-testid="stDataFrame"],
    [data-testid="stMarkdownContainer"]

    /* Contenedor principal de cada bloque */
    .block {
        background-color: #ffffff; /* blanco puro */
        border: 1px solid #e3e8ef;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        margin-bottom: 0.8rem;
    }

    /* Títulos */
    h1, h2, h3, h4 {
        color: #2a4d69;
    }

    /* Botones estilo moderno */
    .stButton>button {
        background: linear-gradient(135deg, #4c8bf5, #6fa8ff);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6em 1.2em;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #3a73d9, #558eff);
        transform: translateY(-1px);
    }

    /* Inputs */
    .stFileUploader label {
        font-weight: 600;
        color: #1f3b57;
    }
</style>
""", unsafe_allow_html=True)


st.title("📄 Generador de Oficios")
st.markdown("#### Simplifica la creación de oficios personalizados a partir de tus archivos Excel y Word.")

# ---------- SECCIÓN 1: ARCHIVOS BASE ----------
st.subheader("📊 Archivos base")
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Subir archivo Excel", type=["xlsm", "xlsx"])
with col2:
    word_file = st.file_uploader("Subir plantilla Word (.docx)", type=["docx"])

# ---------- PROCESAR EXCEL ----------
razones_seleccionadas = []
if excel_file is not None:
    try:
        df = pd.read_excel(excel_file, sheet_name="General", engine="openpyxl")
        df.columns = df.columns.str.strip()
        razones = df["RAZON SOCIAL"].dropna().unique().tolist()
        st.subheader("🏢 Seleccione empresas")
        razones_seleccionadas = st.multiselect("Empresas disponibles:", razones)
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        st.stop()

# ---------- SECCIÓN 3: PDFS OPCIONALES ----------
st.subheader("📎 Adjuntar PDFs opcionales")

col_pdf1, col_pdf2 = st.columns(2)
with col_pdf1:
    pdf_transmision = st.file_uploader("PDF para Transmisión", type=["pdf"])
    pdf_distribucion = st.file_uploader("PDF para Distribución", type=["pdf"])
with col_pdf2:
    pdf_generacion = st.file_uploader("PDF para Generación", type=["pdf"])
    pdf_cliente_libre = st.file_uploader("PDF para Cliente Libre", type=["pdf"])

pdf_files = {
    "Transmisión": pdf_transmision,
    "Generación": pdf_generacion,
    "Distribución": pdf_distribucion,
    "Cliente Libre": pdf_cliente_libre
}

# ---------- SECCIÓN 4: GENERAR ----------
st.subheader("⚙️ Generación de oficios")

if st.button("🚀 Generar y descargar ZIP"):
    if not excel_file or not word_file or not razones_seleccionadas:
        st.warning("Debe subir el Excel, la plantilla Word y seleccionar al menos una empresa.")
        st.stop()

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for razon in razones_seleccionadas:
            fila = df[df["RAZON SOCIAL"] == razon]
            if fila.empty:
                st.warning(f"Razón Social {razon} no encontrada.")
                continue

            nombre_destinatario = fila["GERENTE GENERAL"].values[0]
            cargo = fila["CARGO DEL REPRESENTANTE"].values[0]
            entidad = fila["RAZON SOCIAL"].values[0]
            direccion = fila["DIRECCIÓN"].values[0]
            distrito = fila["Distrito"].values[0]
            actividad = fila["ACTIVIDAD"].values[0]
            codigo = fila["CODIGO"].values[0]

            documento = Document(word_file)
            for parrafo in documento.paragraphs:
                for run in parrafo.runs:
                    if "[Nombre del Destinatario]" in run.text:
                        run.text = run.text.replace("[Nombre del Destinatario]", nombre_destinatario)
                        run.font.bold = True
                        run.font.name = "Poppins"
                        run.font.size = Pt(9.5)
                    if "[Cargo]" in run.text:
                        run.text = run.text.replace("[Cargo]", cargo)
                        run.font.name = "Poppins"
                        run.font.size = Pt(9.5)
                    if "[Entidad]" in run.text:
                        run.text = run.text.replace("[Entidad]", str(entidad))
                        run.font.bold = True
                        run.font.name = "Poppins"
                        run.font.size = Pt(9.5)
                    if "[Dirección]" in run.text:
                        run.text = run.text.replace("[Dirección]", direccion)
                        run.font.name = "Poppins"
                        run.font.size = Pt(9.5)
                    if "[Distrito]" in run.text:
                        run.text = run.text.replace("[Distrito]", distrito)
                        run.font.underline = True
                        run.font.name = "Poppins"
                        run.font.size = Pt(9.5)

            doc_buffer = io.BytesIO()
            documento.save(doc_buffer)
            doc_buffer.seek(0)

            nombre_documento = f"OFICIO-{entidad.replace(' ', '_')}.docx"
            ruta_carpeta_empresa = f"{actividad}/{codigo}/{nombre_documento}"
            zip_file.writestr(ruta_carpeta_empresa, doc_buffer.read())

            pdf_file = pdf_files.get(actividad)
            if pdf_file is not None:
                pdf_bytes = pdf_file.read()
                ruta_pdf = f"{actividad}/{codigo}/{pdf_file.name}"
                zip_file.writestr(ruta_pdf, pdf_bytes)

    zip_buffer.seek(0)
    st.success("✅ Oficios generados correctamente.")
    st.download_button(
        label="📦 Descargar ZIP",
        data=zip_buffer,
        file_name="oficios_generados.zip",
        mime="application/zip"
    )
    