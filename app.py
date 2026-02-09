import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
import time

# --- 1. CONFIGURACIÓN DE API ---
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    # Probamos con la versión Lite para maximizar la cuota disponible
    MODELO_SELECCIONADO = 'gemini-2.0-flash-lite' 
    model = genai.GenerativeModel(MODELO_SELECCIONADO)
except Exception as e:
    st.error(f"Error de configuración de API: {e}")

# --- 2. CONFIGURACIÓN DE LA PLANILLA ---
# Esta es la línea que faltaba y causaba el error
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

@st.cache_data(ttl=60)
def cargar_alumnos():
    try:
        df = pd.read_csv(SHEET_URL)
        df.columns = [c.strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"No se pudo leer la planilla de Google: {e}")
        return pd.DataFrame()

def crear_docx_adecuado(texto_ia, diagnostico):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    
    diag_str = str(diagnostico).lower()
    # Aplicación de formato visual según diagnóstico
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
            # Traducir negritas de la IA al formato Word
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
st.title("Motor Pedagógico v2.6 🚀")

try:
    df = cargar_alumnos()
    
    if not df.empty:
        # Mapeo de columnas por posición
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

            prompt = f"""
            Eres un experto en psicopedagogía. Adapta este examen.
            ALUMNO: {alumno_selec}
            DIAGNÓSTICO: {emergente}
            GRUPO: {grupo}
            
            INSTRUCCIONES:
            1. Si tiene Dislexia: Frases cortas, resalta verbos en negrita.
            2. Si tiene Discalculia: Desglosa problemas, usa viñetas, deja espacio para cálculos.
            3. Si es Grupo A: Simplifica consignas y reduce carga cognitiva.
            4. No incluyas textos introductorios, solo el examen.

            CONTENIDO:
            {texto_orig}
            """

            with st.spinner("Procesando..."):
                try:
                    time.sleep(1) # Pausa para evitar errores de cuota por segundo
                    response = model.generate_content(prompt)
                    
                    docx_file = crear_docx_adecuado(response.text, emergente)
                    st.success("✅ ¡Adecuación lista!")
                    st.download_button(
                        label="⬇️ Descargar Word",
                        data=docx_file,
                        file_name=f"Adecuacion_{alumno_selec}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as api_err:
                    if "429" in str(api_err):
                        st.error("⚠️ Cuota excedida. Espera 60 segundos antes de intentar otro examen.")
                    else:
                        st.error(f"Error de la IA: {api_err}")
    else:
        st.warning("La planilla de alumnos parece estar vacía o inaccesible.")

except Exception as e:
    st.error(f"Error técnico en la aplicación: {e}")
