import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt
import io

# 1. Configuración de API con el modelo verificado en tu lista
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    # Usamos el modelo 2.0-flash que aparece en tu posición #2
    model = genai.GenerativeModel('gemini-2.0-flash')
except Exception as e:
    st.error(f"Error al configurar el modelo: {e}")

# 2. Conexión a Google Sheets
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

@st.cache_data(ttl=60)
def cargar_alumnos():
    df = pd.read_csv(SHEET_URL)
    df.columns = [c.strip() for c in df.columns]
    return df

def crear_docx_adecuado(texto_ia, diagnostico):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    
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

# 3. Interfaz de Usuario
st.title("Motor Pedagógico v2.4 🚀")

try:
    df = cargar_alumnos()
    col_nombre = df.columns[2]
    col_grupo = df.columns[3]
    col_emergente = df.columns[4]

    alumno_selec = st.selectbox("Seleccione Alumno:", df[col_nombre].unique())
    datos = df[df[col_nombre] == alumno_selec].iloc[-1]
    grupo = datos[col_grupo]
    emergente = datos[col_emergente]

    st.info(f"**Alumno:** {alumno_selec} | **Grupo:** {grupo} | **Dificultad:** {emergente}")

    uploaded_file = st.file_uploader("Subir examen original (.docx)", type="docx")

    if uploaded_file and st.button("Generar Examen Adecuado"):
        doc_orig = Document(uploaded_file)
        texto_orig = "\n".join([p.text for p in doc_orig.paragraphs])

        # Prompt con tus instrucciones psicopedagógicas
        prompt = f"""
        Actúa como un experto en educación inclusiva. Re-escribe el siguiente examen.
        
        PERFIL DEL ESTUDIANTE:
        - Nivel de autonomía: {grupo}
        - Diagnóstico: {emergente}
        
        INSTRUCCIONES:
        - Si tiene Dislexia: Usa oraciones simples, resalta verbos de acción en negrita.
        - Si tiene Discalculia: Desglosa problemas, usa viñetas para datos, deja espacios amplios para cálculos.
        - Si es Grupo A: Reduce la complejidad visual y aumenta el apoyo en las consignas.
        - No incluyas explicaciones para el docente ni saludos, solo el contenido del examen adaptado.

        EXAMEN ORIGINAL:
        {texto_orig}
        """

        with st.spinner("Gemini 2.0 está adaptando el contenido..."):
            try:
                response = model.generate_content(prompt)
                
                # Crear el archivo Word
                docx_file = crear_docx_adecuado(response.text, emergente)

                st.success("✅ ¡Adecuación lista!")
                st.download_button(
                    label="⬇️ Descargar Examen Word",
                    data=docx_file,
                    file_name=f"Examen_Adaptado_{alumno_selec}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                with st.expander("Ver vista previa del texto"):
                    st.write(response.text)
                    
            except Exception as api_err:
                st.error(f"Error en la generación: {api_err}")

except Exception as e:
    st.error(f"Error general: {e}")
