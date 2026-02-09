import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt
import io

# --- 1. CONFIGURACIÓN Y DIAGNÓSTICO DE MODELOS ---
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    
    # Listar modelos disponibles para diagnóstico
    modelos_disponibles = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    
    # Intentar usar el primero de la lista que sea Flash, o el primero disponible
    modelo_seleccionado = 'models/gemini-1.5-flash' # Valor por defecto
    for m in modelos_disponibles:
        if '1.5-flash' in m:
            modelo_seleccionado = m
            break
            
    model = genai.GenerativeModel(modelo_seleccionado)
    st.sidebar.success(f"Modelo activo: {modelo_seleccionado}")
except Exception as e:
    st.error(f"Error en la API Key o al listar modelos: {e}")

# --- 2. CONEXIÓN A GOOGLE SHEETS ---
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

# --- 3. INTERFAZ DE USUARIO ---
st.title("Motor Pedagógico v2.3")

try:
    df = cargar_alumnos()
    col_nombre = df.columns[2]
    col_grupo = df.columns[3]
    col_emergente = df.columns[4]

    alumno_selec = st.selectbox("Seleccione Alumno:", df[col_nombre].unique())
    datos = df[df[col_nombre] == alumno_selec].iloc[-1]
    grupo = datos[col_grupo]
    emergente = datos[col_emergente]

    st.info(f"**Grupo:** {grupo} | **Dificultad:** {emergente}")

    uploaded_file = st.file_uploader("Subir examen .docx", type="docx")

    if uploaded_file and st.button("Generar Adecuación"):
        doc_orig = Document(uploaded_file)
        texto_orig = "\n".join([p.text for p in doc_orig.paragraphs])

        prompt = f"Adapta este examen para un alumno con {emergente} del {grupo}: {texto_orig}"

        with st.spinner("Procesando..."):
            try:
                response = model.generate_content(prompt)
                docx_file = crear_docx_adecuado(response.text, emergente)

                st.download_button(
                    label="⬇️ Descargar Word",
                    data=docx_file,
                    file_name=f"Adecuacion_{alumno_selec}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as api_err:
                st.error(f"Error específico de la IA: {api_err}")
                st.write("Modelos que tu cuenta permite:", modelos_disponibles)

except Exception as e:
    st.error(f"Error general: {e}")
