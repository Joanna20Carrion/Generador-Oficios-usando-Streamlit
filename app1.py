import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
import zipfile
import requests

def limpiar_texto(texto):
    return str(texto).replace(" ", "_").replace("/", "").replace("\\", "")

# ---------- CONFIGURACIÓN GENERAL ----------
st.set_page_config(
    page_title="Generador de Oficios",
    page_icon="📄",
    layout="wide"
)

# ---------- ESTILO ----------
st.markdown("""
<style>
.stApp {
    background-color: #f6f8fb;
    color: #1f3b57;
}

h1, h2, h3 {
    color: #2a4d69;
}

.stButton>button {
    background: linear-gradient(135deg, #4c8bf5, #6fa8ff);
    color: white;
    border-radius: 8px;
    padding: 0.6em 1.2em;
    border: none;
}

.stButton>button:hover {
    background: linear-gradient(135deg, #3a73d9, #558eff);
}
</style>
""", unsafe_allow_html=True)

st.title("📄 Generador de Oficios")

st.info("""
Plataforma para la generación automatizada de oficios institucionales utilizando una plantilla Word y datos 
provenientes de un directorio de empresas, permitiendo producir documentos oficiales de manera rápida, uniforme 
y eficiente.
""")

# ---------- GOOGLE SHEETS ----------
SHEET_URL = "https://script.google.com/macros/s/AKfycbwgeHn7HqMoFH4VkqyMRGZ8v-B7YAFA8_PgH1XYnIzHfwmSEIGaweuPFjjdbMFTjC_0rg/exec"

try:

    response = requests.get(SHEET_URL)

    if response.status_code != 200:
        st.error("No se pudo conectar con Google Sheets")
        st.stop()

    data = response.json()

    df = pd.DataFrame(data)

    df.columns = df.columns.str.strip()

except Exception as e:

    st.error(f"Error al conectar con Google Sheets: {e}")

    st.stop()

# ---------- EMPRESAS ----------
st.subheader("🏢 Seleccione empresas")

razones = df["Nombre de la Empresa"].dropna().unique().tolist()

razones_seleccionadas = st.multiselect(
    "Empresas disponibles:",
    razones
)

st.divider()

# ---------- PLANTILLA ----------
st.subheader("📄 Subir plantilla Word")

st.info("""
Para que el sistema reemplace los datos automáticamente, copie y pegue exactamente
las siguientes variables dentro del documento Word:
""")

st.code("""
Señor
[Nombre del Destinatario]
[Cargo]
[Entidad]
[Dirección]
[Distrito]

Asunto: [Asunto]
""", language="text")

word_file = st.file_uploader(
    "Plantilla Word (.docx)",
    type=["docx"]
)

st.divider()

# ---------- DATOS ADICIONALES ----------
st.subheader("✏️ Información adicional")

asunto = st.text_input(
    "Asunto del oficio",
    placeholder="Ejemplo: Solicitud de información técnica"
)

procedimiento = st.text_input(
    "Procedimiento (solo para el nombre del archivo)",
    placeholder="Ejemplo: ERACMF"
)

st.divider()

# ---------- ARCHIVOS ADJUNTOS ----------
st.subheader("📎 Adjuntar archivos por tipo de actividad")

tipos_archivo = [
    "pdf","docx","doc","xlsx","xls","txt","zip","rar","csv"
]

col1, col2 = st.columns(2)

with col1:

    archivos_transmision = st.file_uploader(
        "Archivos para Transmisión",
        type=tipos_archivo,
        accept_multiple_files=True
    )

    archivos_distribucion = st.file_uploader(
        "Archivos para Distribución",
        type=tipos_archivo,
        accept_multiple_files=True
    )

with col2:

    archivos_generacion = st.file_uploader(
        "Archivos para Generación",
        type=tipos_archivo,
        accept_multiple_files=True
    )

    archivos_cliente_libre = st.file_uploader(
        "Archivos para Cliente Libre",
        type=tipos_archivo,
        accept_multiple_files=True
    )

archivos_por_actividad = {
    "Transmisión": archivos_transmision,
    "Generación": archivos_generacion,
    "Distribución": archivos_distribucion,
    "Cliente Libre": archivos_cliente_libre
}

st.divider()

# ---------- GENERAR ----------
st.subheader("⚙️ Generar Oficios")

if st.button("🚀 Generar y descargar ZIP"):

    if not word_file or not razones_seleccionadas:

        st.warning("Debe subir la plantilla Word y seleccionar empresas.")

        st.stop()

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:

        for razon in razones_seleccionadas:

            fila = df[df["Nombre de la Empresa"] == razon]

            if fila.empty:
                continue

            nombre_destinatario = fila["GERENTE GENERAL"].values[0]
            cargo = fila["CARGO DEL REPRESENTANTE"].values[0]
            entidad = fila["Nombre de la Empresa"].values[0]
            direccion = fila["DIRECCIÓN"].values[0]
            distrito = fila["Distrito"].values[0]
            actividad = fila["ACTIVIDAD"].values[0]
            codigo = fila["CODIGO"].values[0]

            documento = Document(word_file)

            archivos = archivos_por_actividad.get(actividad, [])

            for parrafo in documento.paragraphs:

                if "[Nombre del Destinatario]" in parrafo.text:
                    partes = parrafo.text.split("[Nombre del Destinatario]")

                    parrafo.clear()

                    run1 = parrafo.add_run(partes[0])
                    run1.font.name = "Poppins"
                    run1.font.size = Pt(9.5)

                    run2 = parrafo.add_run(str(nombre_destinatario))
                    run2.bold = True
                    run2.font.name = "Poppins"
                    run2.font.size = Pt(9.5)

                    if len(partes) > 1:
                        run3 = parrafo.add_run(partes[1])
                        run3.font.name = "Poppins"
                        run3.font.size = Pt(9.5)

                if "[Cargo]" in parrafo.text:
                    parrafo.text = parrafo.text.replace("[Cargo]", str(cargo))

                if "[Entidad]" in parrafo.text:
                    partes = parrafo.text.split("[Entidad]")

                    parrafo.clear()

                    run1 = parrafo.add_run(partes[0])
                    run1.font.name = "Poppins"
                    run1.font.size = Pt(9.5)

                    run2 = parrafo.add_run(str(entidad))
                    run2.bold = True
                    run2.font.name = "Poppins"
                    run2.font.size = Pt(9.5)

                    if len(partes) > 1:
                        run3 = parrafo.add_run(partes[1])
                        run3.font.name = "Poppins"
                        run3.font.size = Pt(9.5)

                if "[Dirección]" in parrafo.text:
                    parrafo.text = parrafo.text.replace("[Dirección]", str(direccion))

                if "[Distrito]" in parrafo.text:
                    parrafo.text = parrafo.text.replace("[Distrito]", str(distrito))

                if "[Asunto]" in parrafo.text:
                    parrafo.text = parrafo.text.replace("[Asunto]", str(asunto))

                for run in parrafo.runs:
                    run.font.name = "Poppins"
                    run.font.size = Pt(9.5)

            doc_buffer = io.BytesIO()

            documento.save(doc_buffer)

            doc_buffer.seek(0)

            empresa_limpia = limpiar_texto(entidad)
            asunto_limpio = limpiar_texto(asunto)
            procedimiento_limpio = limpiar_texto(procedimiento)

            nombre_documento = f"Oficio_{procedimiento_limpio}_{empresa_limpia}_{asunto_limpio}.docx"

            ruta_doc = f"{actividad}/{codigo}/{nombre_documento}"

            zip_file.writestr(
                ruta_doc,
                doc_buffer.read()
            )

            if archivos:

                for archivo in archivos:

                    zip_file.writestr(
                        f"{actividad}/{codigo}/{archivo.name}",
                        archivo.read()
                    )

    zip_buffer.seek(0)

    st.success("✅ Oficios generados correctamente")

    st.download_button(
        label="📦 Descargar ZIP",
        data=zip_buffer,
        file_name="oficios_generados.zip",
        mime="application/zip"
    )
    