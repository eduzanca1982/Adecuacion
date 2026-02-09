import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt
import io

# 1. Configuración de API con modelo estable
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    # Cambio a versión estable para evitar el error 404
    model = genai.GenerativeModel('gemini-1.5-flash-latest')
except Exception as e:
    st.error("Error de configuración de API.")

# 2. Configuración de Google Sheets
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

@st.cache_data(ttl=60)
def cargar_alumnos():
    return pd.read_csv(SHEET_URL)

def crear_docx_adecuado(texto_ia, diagnostico):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    
    # Aplicación de reglas de formato según diagnóstico
    if "Dislexia" in str(diagnostico):
        font.name = 'Arial'
        font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.5 # Interlineado 1.5
    else:
        font.name = 'Verdana'
        font.size = Pt(11)
        style.paragraph_format.line_spacing = 1.15

    for linea in texto_ia.split('\n'):
        if linea.strip():
            p = doc.add_paragraph()
            # Procesar negritas de Markdown a Word
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
st.title("Adaptación de Contenidos v2.0")

try:
    df = cargar_alumnos()
    alumno_selec = st.selectbox("Alumno:", df["Nombre y apellido del alumno"].unique())
    
    # Extraer datos del registro
    datos = df[df["Nombre y apellido del alumno"] == alumno_selec].iloc[-1]
    grupo = datos["Los grupos base se categorizan según su autonomía y dinámica de trabajo."]
    emergente = datos["Casos emergentes: alumnos con dificultades específicas que necesitan un acompañamiento personalizado para aprender con éxito."]

    st.info(f"Dificultad detectada: {emergente}")

    uploaded_file = st.file_uploader("Subir examen .docx", type="docx")

    if uploaded_file and st.button("Procesar Adecuación"):
        # Leer archivo original
        doc_orig = Document(uploaded_file)
        texto_orig = "\n".join([p.text for p in doc_orig.paragraphs])

        # Prompt con instrucciones de adecuación
        prompt = f"""
        Re-escribe este examen para el alumno {alumno_selec}.
        Grupo: {grupo}
        Dificultad: {emergente}
        
        Instrucciones:
        - Si es Dislexia: Usa frases cortas y resalta verbos en negrita.
        - Si es Discalculia: Desglosa problemas y usa listas de datos.
        - Si es TDAH: Una consigna por oración y elimina ruido visual.
        - Si es Grupo A: Aumenta el andamiaje docente.
        
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
    st.error(f"Error técnico: {e}")
