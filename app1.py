import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
import zipfile
import requests


def limpiar_texto(texto):
    return str(texto).replace(" ", "_").replace("/", "").replace("\\", "")


# ============================================================
# REEMPLAZAR VARIABLE SIN MODIFICAR EL RESTO DEL FORMATO
# ============================================================

def reemplazar_variable(parrafo, variable, valor, negrita=False, subrayado=False):
    """
    Reemplaza únicamente la variable indicada.

    El texto que NO es variable conserva:
    - tamaño
    - fuente
    - negrita
    - cursiva
    - subrayado
    - color
    - demás formato

    El valor reemplazado se establece en:
    - Poppins
    - 9 pt
    """

    valor = "" if valor is None else str(valor)

    # Buscar primero si la variable está completa dentro de un solo run
    for run in parrafo.runs:

        if variable in run.text:

            texto_original = run.text
            partes = texto_original.split(variable)

            # Guardar formato original del run
            fuente_original = run.font.name
            tamano_original = run.font.size
            negrita_original = run.bold
            cursiva_original = run.italic
            subrayado_original = run.underline
            color_original = None

            if run.font.color and run.font.color.rgb:
                color_original = run.font.color.rgb

            # Limpiar únicamente este run
            run.text = ""

            # Texto anterior a la variable
            if partes[0]:
                r = run._element
                nuevo_run = run._parent.add_r()
                nuevo_run.text = partes[0]

                nuevo_run.font.name = fuente_original
                nuevo_run.font.size = tamano_original
                nuevo_run.bold = negrita_original
                nuevo_run.italic = cursiva_original
                nuevo_run.underline = subrayado_original

                if color_original:
                    nuevo_run.font.color.rgb = color_original

            # Valor reemplazado
            nuevo_run = run._parent.add_r()
            nuevo_run.text = valor

            nuevo_run.font.name = "Poppins"
            nuevo_run.font.size = Pt(9)
            nuevo_run.bold = negrita
            nuevo_run.underline = subrayado

            # Texto posterior a la variable
            if len(partes) > 1 and partes[1]:

                nuevo_run = run._parent.add_r()
                nuevo_run.text = partes[1]

                nuevo_run.font.name = fuente_original
                nuevo_run.font.size = tamano_original
                nuevo_run.bold = negrita_original
                nuevo_run.italic = cursiva_original
                nuevo_run.underline = subrayado_original

                if color_original:
                    nuevo_run.font.color.rgb = color_original

            return True

    return False


# ============================================================
# REEMPLAZO DE VARIABLES QUE PUEDEN ESTAR DIVIDIDAS EN RUNS
# ============================================================

def reemplazar_variables_en_parrafo(parrafo, reemplazos):
    """
    Maneja variables aunque Word las haya dividido en varios runs.
    """

    # Primero intentamos reemplazo normal
    for variable, datos in reemplazos.items():

        valor = datos["valor"]
        negrita = datos.get("negrita", False)
        subrayado = datos.get("subrayado", False)

        if reemplazar_variable(
            parrafo,
            variable,
            valor,
            negrita,
            subrayado
        ):
            continue

        # ----------------------------------------------------
        # Si la variable está dividida entre varios runs
        # ----------------------------------------------------

        texto_completo = "".join(run.text for run in parrafo.runs)

        if variable not in texto_completo:
            continue

        posicion_inicio = texto_completo.find(variable)
        posicion_fin = posicion_inicio + len(variable)

        acumulado = 0
        runs_afectados = []

        for i, run in enumerate(parrafo.runs):

            inicio_run = acumulado
            fin_run = acumulado + len(run.text)

            if fin_run > posicion_inicio and inicio_run < posicion_fin:
                runs_afectados.append(i)

            acumulado = fin_run

        if not runs_afectados:
            continue

        primer_run = runs_afectados[0]

        # Texto completo antes y después de la variable
        texto_antes = texto_completo[:posicion_inicio]
        texto_despues = texto_completo[posicion_fin:]

        # Eliminar runs afectados
        for i in reversed(runs_afectados):
            elemento = parrafo.runs[i]._element
            elemento.getparent().remove(elemento)

        # Crear nuevamente el contenido respetando el texto
        if texto_antes:

            run_antes = parrafo.add_run(texto_antes)

        run_variable = parrafo.add_run(str(valor))

        run_variable.font.name = "Poppins"
        run_variable.font.size = Pt(9)
        run_variable.bold = negrita
        run_variable.underline = subrayado

        if texto_despues:
            run_despues = parrafo.add_run(texto_despues)

        return


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Generador de Oficios",
    page_icon="📄",
    layout="wide"
)


# ============================================================
# ESTILO
# ============================================================

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


# ============================================================
# GOOGLE SHEETS
# ============================================================

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


# ============================================================
# EMPRESAS
# ============================================================

st.subheader("🏢 Seleccione empresas")

razones = (
    df["Nombre de la Empresa"]
    .dropna()
    .unique()
    .tolist()
)

razones_seleccionadas = st.multiselect(
    "Empresas disponibles:",
    razones
)

st.divider()


# ============================================================
# PLANTILLA
# ============================================================

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


# ============================================================
# DATOS ADICIONALES
# ============================================================

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


# ============================================================
# ARCHIVOS ADJUNTOS
# ============================================================

st.subheader("📎 Adjuntar archivos por tipo de actividad")

tipos_archivo = [
    "pdf",
    "docx",
    "doc",
    "xlsx",
    "xls",
    "txt",
    "zip",
    "rar",
    "csv"
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


# ============================================================
# GENERAR
# ============================================================

st.subheader("⚙️ Generar Oficios")


if st.button("🚀 Generar y descargar ZIP"):

    if not word_file or not razones_seleccionadas:

        st.warning(
            "Debe subir la plantilla Word y seleccionar empresas."
        )

        st.stop()


    zip_buffer = io.BytesIO()


    with zipfile.ZipFile(
        zip_buffer,
        "w"
    ) as zip_file:


        for razon in razones_seleccionadas:

            fila = df[
                df["Nombre de la Empresa"] == razon
            ]


            if fila.empty:
                continue


            nombre_destinatario = fila[
                "GERENTE GENERAL"
            ].values[0]

            cargo = fila[
                "CARGO DEL REPRESENTANTE"
            ].values[0]

            entidad = fila[
                "Nombre de la Empresa"
            ].values[0]

            direccion = fila[
                "DIRECCIÓN"
            ].values[0]

            distrito = fila[
                "Distrito"
            ].values[0]

            actividad = fila[
                "ACTIVIDAD"
            ].values[0]

            codigo = fila[
                "CODIGO"
            ].values[0]


            # ------------------------------------------------
            # CREAR DOCUMENTO DESDE LA PLANTILLA
            # ------------------------------------------------

            documento = Document(word_file)


            # ------------------------------------------------
            # VARIABLES
            # ------------------------------------------------

            reemplazos = {

                "[Nombre del Destinatario]": {
                    "valor": nombre_destinatario,
                    "negrita": True,
                    "subrayado": False
                },

                "[Cargo]": {
                    "valor": cargo,
                    "negrita": False,
                    "subrayado": False
                },

                "[Entidad]": {
                    "valor": entidad,
                    "negrita": True,
                    "subrayado": False
                },

                "[Dirección]": {
                    "valor": direccion,
                    "negrita": False,
                    "subrayado": False
                },

                "[Distrito]": {
                    "valor": distrito,
                    "negrita": False,
                    "subrayado": True
                },

                "[Asunto]": {
                    "valor": asunto,
                    "negrita": False,
                    "subrayado": False
                }

            }


            # ------------------------------------------------
            # PROCESAR PÁRRAFOS
            # ------------------------------------------------

            for parrafo in documento.paragraphs:

                reemplazar_variables_en_parrafo(
                    parrafo,
                    reemplazos
                )


            # ------------------------------------------------
            # GUARDAR DOCUMENTO
            # ------------------------------------------------

            doc_buffer = io.BytesIO()

            documento.save(doc_buffer)

            doc_buffer.seek(0)


            # ------------------------------------------------
            # NOMBRE DEL ARCHIVO
            # ------------------------------------------------

            empresa_limpia = limpiar_texto(
                entidad
            )

            asunto_limpio = limpiar_texto(
                asunto
            )

            procedimiento_limpio = limpiar_texto(
                procedimiento
            )


            nombre_documento = (
                f"Oficio_"
                f"{procedimiento_limpio}_"
                f"{empresa_limpia}_"
                f"{asunto_limpio}.docx"
            )


            ruta_doc = (
                f"{actividad}/"
                f"{codigo}/"
                f"{nombre_documento}"
            )


            zip_file.writestr(
                ruta_doc,
                doc_buffer.read()
            )


            # ------------------------------------------------
            # ADJUNTOS
            # ------------------------------------------------

            archivos = archivos_por_actividad.get(
                actividad,
                []
            )


            if archivos:

                for archivo in archivos:

                    zip_file.writestr(
                        f"{actividad}/{codigo}/{archivo.name}",
                        archivo.read()
                    )


    # ========================================================
    # DESCARGA
    # ========================================================

    zip_buffer.seek(0)


    st.success(
        "✅ Oficios generados correctamente"
    )


    st.download_button(
        label="📦 Descargar ZIP",
        data=zip_buffer,
        file_name="oficios_generados.zip",
        mime="application/zip"
    )
