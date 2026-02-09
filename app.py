import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt
import io

# 1. Configuración de API
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    model = genai.GenerativeModel('gemini-1.5-flash-latest')
except Exception as e:
    st.error("Error de configuración de API.")

# 2. Configuración de Google Sheets
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

@st.cache_data(ttl=60)
def cargar_alumnos():
    df = pd.read_csv(SHEET_URL)
    # Limpieza básica de nombres de columnas
    df.columns = [c.strip() for c in df.columns]
    return df

def crear_docx_adecuado(texto_ia, diagnostico):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    
    # Lógica de formato por diagnóstico
    diag_str = str(diagnostico).lower()
    if "dislexia" in diag_str:
        font.name = 'Arial'
        font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.5
    else:
        font.name = 'Verdana'
        font.size = Pt(11)
        style.paragraph_format.line_spacing = 1.15

    for linea in texto_ia.split('\n'):
        if linea.strip():
            p = doc.add_paragraph()
            partes = linea.split("**")
            for i, parte in enumerate(partes):
                run = p.add_run(parte)
                if i % 2 != 0:
                    run.bold = True
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 3. Interfaz
st.title("Adaptación de Contenidos v2.1")

try:
    df = cargar_alumnos()
    
    # Mapeo dinámico de columnas por posición para evitar errores de texto largo
    # Columna 2: Nombre, Columna 3: Grupo, Columna 4: Emergente (ajustar si varía)
    col_nombre = df.columns[2]
    col_grupo = df.columns[3]
    col_emergente = df.columns[4]

    alumno_selec = st.selectbox("Seleccione Alumno:", df[col_nombre].unique())
    
    datos = df[df[col_nombre] == alumno_selec].iloc[-1]
    grupo = datos[col_grupo]
    emergente = datos[col_emergente]

    st.info(f"**Grupo:** {grupo} | **Dificultad:** {emergente}")

    uploaded_file = st.file_uploader("Subir examen .docx", type="docx")

    if uploaded_file and st.button("Procesar Adecuación"):
        doc_orig = Document(uploaded_file)
        texto_orig = "\n".join([p.text for p in doc_orig.paragraphs])

        prompt = f"""
        Re-escribe este examen para el alumno {alumno_selec}.
        Grupo: {grupo}
        Dificultad: {emergente}
        
        Instrucciones:
        - Si es Dislexia: Usa frases cortas y resalta verbos en negrita.
        - Si es Discalculia: Desglosa problemas y usa listas de datos.
        - Si es TDAH: Una consigna por oración y elimina ruido visual.
        - Si es Grupo A: Aumenta el andamiaje docente.
        - No incluyas saludos ni introducciones, solo el examen.

        Examen:
        {texto_orig}
        """

        with st.spinner("Generando archivo..."):
            response = model.generate_content(prompt)
            docx_file = crear_docx_adecuado(response.text, emergente)

            st.download_button(
                label="Descargar Examen en Word",
                data=docx_file,
                file_name=f"Examen_{alumno_selec}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

except Exception as e:
    st.error(f"Error al leer la planilla. Verifica que las columnas coincidan. Detalle: {e}")
