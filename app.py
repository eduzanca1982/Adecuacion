import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
import time

# 1. Configuración de API - ROTACIÓN DE MODELO PARA EVITAR QUOTA 429
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    
    # Intentamos usar la versión Lite que tiene cuotas distintas
    # Si falla, puedes cambiarlo manualmente por 'models/gemini-1.5-flash'
    MODELO_SELECCIONADO = 'gemini-2.0-flash-lite' 
    model = genai.GenerativeModel(MODELO_SELECCIONADO)
except Exception as e:
    st.error(f"Error de configuración: {e}")

# ... (Funciones cargar_alumnos y crear_docx_adecuado se mantienen igual que v2.4) ...

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
        font.name = 'Arial'; font.size = Pt(12); style.paragraph_format.line_spacing = 1.5
    else:
        font.name = 'Verdana'; font.size = Pt(11); style.paragraph_format.line_spacing = 1.15
    for linea in texto_ia.split('\n'):
        if linea.strip():
            p = doc.add_paragraph()
            partes = linea.split("**")
            for i, parte in enumerate(partes):
                run = p.add_run(parte); run.bold = (i % 2 != 0)
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio

# --- INTERFAZ ---
st.title("Motor Pedagógico v2.5 🚀")

try:
    df = cargar_alumnos()
    col_nombre, col_grupo, col_emergente = df.columns[2], df.columns[3], df.columns[4]
    alumno_selec = st.selectbox("Seleccione Alumno:", df[col_nombre].unique())
    datos = df[df[col_nombre] == alumno_selec].iloc[-1]
    grupo, emergente = datos[col_grupo], datos[col_emergente]

    st.info(f"**Alumno:** {alumno_selec} | **Dificultad:** {emergente}")
    uploaded_file = st.file_uploader("Subir examen original (.docx)", type="docx")

    if uploaded_file and st.button("Generar Examen Adecuado"):
        doc_orig = Document(uploaded_file)
        texto_orig = "\n".join([p.text for p in doc_orig.paragraphs])

        prompt = f"Actúa como experto en psicopedagogía. Adapta este examen para {alumno_selec} ({emergente}, {grupo}): {texto_orig}"

        with st.spinner(f"Usando {MODELO_SELECCIONADO}... Espera un momento."):
            try:
                # Pequeña pausa de seguridad antes de llamar para refrescar cuota por minuto
                time.sleep(2) 
                response = model.generate_content(prompt)
                
                docx_file = crear_docx_adecuado(response.text, emergente)
                st.success("✅ ¡Adecuación lista!")
                st.download_button("⬇️ Descargar Word", docx_file, f"Examen_{alumno_selec}.docx")
                
            except Exception as api_err:
                if "429" in str(api_err):
                    st.error("⚠️ Cuota excedida. Por favor, espera 60 segundos y vuelve a intentarlo. Google limita las peticiones gratuitas por minuto.")
                else:
                    st.error(f"Error: {api_err}")

except Exception as e:
    st.error(f"Error general: {e}")
