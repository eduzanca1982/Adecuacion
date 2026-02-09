import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile
import time

# 1. Configuración de API (Gemini 2.0 Flash)
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-2.0-flash')

# 2. Configuración de Planilla
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

@st.cache_data(ttl=60)
def cargar_datos():
    df = pd.read_csv(SHEET_URL)
    df.columns = [c.strip() for c in df.columns]
    return df

def crear_docx_humano(texto_ia, nombre_alumno, diagnostico):
    doc = Document()
    
    # Configuración de estilos base
    style = doc.styles['Normal']
    font = style.font
    diag_str = str(diagnostico).lower()
    
    # Adecuaciones Tipográficas Silenciosas
    if "dislexia" in diag_str:
        font.name = 'Arial'
        font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.5
    else:
        font.name = 'Calibri'
        font.size = Pt(11)
        style.paragraph_format.line_spacing = 1.15

    # Encabezado tradicional de examen
    header = doc.add_paragraph()
    run_h = header.add_run(f"Nombre: {nombre_alumno} {'_'*25} Grado: {'_'*10}")
    run_h.bold = True
    doc.add_paragraph("\n") # Espacio para el título

    lineas = texto_ia.split('\n')
    for linea in lineas:
        linea = linea.strip()
        if not linea: continue
        
        p = doc.add_paragraph()
        
        # Detectar títulos por formato (Líneas cortas sin punto final)
        es_titulo = len(linea) < 60 and not linea.endswith('.') and not linea[0].isdigit()
        
        # Limpiar marcas de Markdown y procesar negritas
        limpia = linea.replace('#', '').strip()
        partes = limpia.split("**")
        
        for i, parte in enumerate(partes):
            run = p.add_run(parte)
            if i % 2 != 0: run.bold = True # Negritas de la IA
            
            if es_titulo:
                run.bold = True
                run.font.size = Pt(13)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT
                
        # Si es Discalculia, agregar espacios para el desarrollo
        if "discalculia" in diag_str and any(char.isdigit() for char in linea):
            doc.add_paragraph("\n" * 2) # Deja espacio en blanco para cálculos

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 3. Interfaz
st.title("Adecuación Curricular Humana 🎓")

try:
    df = cargar_datos()
    col_grado, col_nombre, col_emergente = df.columns[1], df.columns[2], df.columns[4]

    grado_selec = st.sidebar.selectbox("Seleccionar Grado para Procesar:", df[col_grado].unique())
    alumnos_grado = df[df[col_grado] == grado_selec]
    
    # Filtro: Solo alumnos con alguna dificultad marcada
    alumnos_emergentes = alumnos_grado[alumnos_grado[col_emergente].notna() & (alumnos_grado[col_emergente] != "Ninguna")]

    st.sidebar.write(f"Alumnos a adecuar en {grado_selec}: {len(alumnos_emergentes)}")

    archivo_orig = st.file_uploader("Subir Examen Base (.docx)", type="docx")

    if archivo_orig and st.button("Generar Adecuaciones"):
        doc_base = Document(archivo_orig)
        texto_base = "\n".join([p.text for p in doc_base.paragraphs])
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            progress_bar = st.progress(0)
            
            for i, (idx, alumno) in enumerate(alumnos_emergentes.iterrows()):
                nombre = alumno[col_nombre]
                diag = alumno[col_emergente]
                
                # PROMPT DE ALTA FIDELIDAD
                prompt = f"""
                Re-escribe el examen adjunto para un alumno con {diag}. 
                OBJETIVO: Que el examen sea idéntico en espíritu al original, pero con adecuaciones de acceso.
                
                REGLAS ESTRICTAS:
                1. MANTÉN el vocabulario del docente. No uses frases de IA como "Aquí tienes tu examen".
                2. Si es Dislexia: Divide párrafos largos en oraciones cortas. Usa negrita SOLO en verbos de consigna (ej: Calcula, Lee, Une).
                3. Si es Discalculia: No cambies los números, solo organiza la información de forma visual.
                4. Estética: Mantén la numeración de los ejercicios (1, 2, 3...).
                5. No agregues contenido nuevo, solo adecua la forma del existente.

                EXAMEN ORIGINAL:
                {texto_base}
                """
                
                time.sleep(4) # Control de cuota
                response = model.generate_content(prompt)
                
                docx_buffer = crear_docx_humano(response.text, nombre, diag)
                zip_file.writestr(f"Adecuacion_{nombre}.docx", docx_buffer.getvalue())
                
                progress_bar.progress((i + 1) / len(alumnos_emergentes))
        
        st.success(f"Hecho. Se generaron {len(alumnos_emergentes)} exámenes.")
        st.download_button("Descargar Carpeta (.zip)", zip_buffer.getvalue(), f"Examenes_{grado_selec}.zip")

except Exception as e:
    st.error(f"Error técnico: {e}")
