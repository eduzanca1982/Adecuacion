import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt
import io

# 1. Configuración de API con nombre de modelo estándar
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    # Cambiamos a 'gemini-1.5-flash' que es el nombre más compatible
    model = genai.GenerativeModel('gemini-1.5-flash')
except Exception as e:
    st.error("Error de configuración de API.")

# 2. Configuración de Google Sheets
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
st.title("Adaptación de Contenidos v2.2")

try:
    df = cargar_alumnos()
    
    # Índices de columnas basados en tu planilla
    col_nombre = df.columns[2]
    col_grupo = df.columns[3]
    col_emergente = df.columns[4]

    alumno_selec = st.selectbox("Seleccione Alumno:", df[col_nombre].unique())
    
    datos = df[df[col_nombre] == alumno_selec].iloc[-1]
    grupo = datos[col_grupo]
    emergente = datos[col_emergente]

    st.info(f"**Grupo:** {grupo} | **Dificultad:** {emergente}")

    uploaded_file = st.file_uploader("Subir examen .docx", type="docx")

    if uploaded_file:
        if st.button("Generar Adecuación"):
            # Leer texto del Word original
            doc_orig = Document(uploaded_file)
            texto_orig = "\n".join([p.text for p in doc_orig.paragraphs])

            prompt = f"""
            Eres un experto en psicopedagogía. Tu tarea es adaptar el examen adjunto.
            ALUMNO: {alumno_selec}
            GRUPO: {grupo}
            DIAGNÓSTICO: {emergente}
            
            INSTRUCCIONES DE ADAPTACIÓN:
            1. Si tiene Dislexia: Usa frases cortas, resalta verbos en negrita, simplifica vocabulario.
            2. Si tiene Discalculia: Desglosa problemas en pasos, usa viñetas, deja espacio para cálculos.
            3. Si es Grupo A: Reduce carga cognitiva y aumenta el andamiaje.
            4. No incluyas introducciones ni despedidas.
            
            EXAMEN A ADAPTAR:
            {texto_orig}
            """

            try:
                with st.spinner("Procesando con Gemini..."):
                    # Llamada directa al método de generación
                    response = model.generate_content(prompt)
                    
                    if response.text:
                        st.success("¡Adecuación generada con éxito!")
                        docx_file = crear_docx_adecuado(response.text, emergente)

                        st.download_button(
                            label="⬇️ Descargar Examen Adaptado (.docx)",
                            data=docx_file,
                            file_name=f"Examen_Adaptado_{alumno_selec}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error("La IA no devolvió contenido.")
            except Exception as api_err:
                st.error(f"Error en la API de Google: {api_err}")

except Exception as e:
    st.error(f"Error en la aplicación: {e}")
