import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt
import PyPDF2
import io
import zipfile
import time

# 1. Configuración de API (Usando Gemini 2.0 que es el que tienes habilitado)
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    model = genai.GenerativeModel('gemini-2.0-flash')
except Exception as e:
    st.error("Error en API Key")

# 2. Configuración de Planilla
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

@st.cache_data(ttl=60)
def cargar_datos():
    df = pd.read_csv(SHEET_URL)
    df.columns = [c.strip() for c in df.columns]
    return df

def extraer_texto(archivo):
    ext = archivo.name.split('.')[-1].lower()
    if ext == 'docx':
        doc = Document(archivo)
        return "\n".join([p.text for p in doc.paragraphs])
    elif ext == 'pdf':
        reader = PyPDF2.PdfReader(archivo)
        return "\n".join([page.extract_text() for page in reader.pages])
    return None

def crear_docx_fiel(texto_ia, nombre_alumno, diagnostico):
    doc = Document()
    diag = str(diagnostico).lower()
    
    # Configuración estética base
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial' if "dislexia" in diag else 'Verdana'
    font.size = Pt(12 if "dislexia" in diag else 11)
    style.paragraph_format.line_spacing = 1.5 if "dislexia" in diag else 1.15

    # Encabezado limpio
    p = doc.add_paragraph()
    p.add_run(f"Estudiante: {nombre_alumno}").bold = True
    doc.add_paragraph("-" * 30)

    for linea in texto_ia.split('\n'):
        if not linea.strip(): continue
        para = doc.add_paragraph()
        
        # Detectar títulos (líneas cortas sin punto final)
        es_titulo = len(linea) < 50 and not linea.endswith('.')
        
        partes = linea.split("**")
        for i, parte in enumerate(partes):
            run = para.add_run(parte)
            if i % 2 != 0: run.bold = True
            if es_titulo: 
                run.bold = True
                run.font.size = Pt(14)

        # Espacio para Discalculia
        if "discalculia" in diag and any(c.isdigit() for c in linea):
            doc.add_paragraph("\n" * 2)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 3. Interfaz
st.title("Motor Pedagógico Multiformato 🚀")

try:
    df = cargar_datos()
    col_grado, col_nombre, col_emergente = df.columns[1], df.columns[2], df.columns[4]
    
    grado_selec = st.sidebar.selectbox("Seleccione el Grado:", df[col_grado].unique())
    alumnos_adecuar = df[(df[col_grado] == grado_selec) & (df[col_emergente].notna()) & (df[col_emergente] != "Ninguna")]

    st.sidebar.info(f"Alumnos a procesar en {grado_selec}: {len(alumnos_adecuar)}")

    archivo_base = st.file_uploader("Subir Examen (DOCX o PDF)", type=["docx", "pdf"])

    if archivo_base and st.button(f"Generar Adecuaciones de {grado_selec}"):
        texto_base = extraer_texto(archivo_base)
        
        if texto_base:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                bar = st.progress(0)
                
                for i, (idx, fila) in enumerate(alumnos_adecuar.iterrows()):
                    nombre, dificultad = fila[col_nombre], fila[col_emergente]
                    
                    prompt = f"Adapta este examen para {nombre} ({dificultad}). Mantén estética y títulos: {texto_base}"
                    
                    time.sleep(4) # Evitar Error 429
                    response = model.generate_content(prompt)
                    
                    doc_final = crear_docx_fiel(response.text, nombre, dificultad)
                    zip_file.writestr(f"Adecuacion_{nombre}.docx", doc_final.getvalue())
                    bar.progress((i + 1) / len(alumnos_adecuar))
            
            st.success("¡Proceso completado!")
            st.download_button("Descargar Todo (.zip)", zip_buffer.getvalue(), f"Examenes_{grado_selec}.zip")

except Exception as e:
    st.error(f"Error: {e}")
