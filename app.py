import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import PyPDF2
import io
import zipfile
import time

# ----------------------------
# 1. Configuración de API y Prompt
# ----------------------------
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-2.0-flash')

SYSTEM_PROMPT = """Actúa como un diseñador editorial pedagógico. Tu misión es adecuar exámenes de primaria.
REGLAS ESTÉTICAS:
1. No escribas introducciones. Empieza directo en el examen.
2. MANTÉN la numeración original (1, 2, 3 o 1.a, 1.b).
3. IMÁGENES: Escribe [MANTENER IMAGEN AQUÍ] donde el texto original las mencione.
4. Si el ejercicio es de completar, usa líneas largas de puntos: ........................

ADECUACIONES:
- DISLEXIA: Frases cortas. Negrita SOLO en verbos de consigna.
- DISCALCULIA: Datos en listas. Añade espacios amplios entre ejercicios.
- TDAH: Una consigna por párrafo. Pasos numerados.
- GRUPO A: Simplifica sintaxis sin perder el objetivo pedagógico."""

# ----------------------------
# 2. Funciones de Maquetación y Diseño
# ----------------------------
def extraer_texto(archivo):
    ext = archivo.name.split(".")[-1].lower()
    if ext == "docx":
        doc = Document(archivo)
        return "\n".join([p.text for p in doc.paragraphs])
    elif ext == "pdf":
        reader = PyPDF2.PdfReader(archivo)
        return "\n".join([p.extract_text() for p in reader.pages])
    return ""

def crear_docx_premium(texto_ia, nombre, diagnostico, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    color_institucional = RGBColor(31, 73, 125) # Azul Marino Profesional

    # --- ENCABEZADO CON LOGO Y TABLA ---
    section = doc.sections[0]
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.5)
    
    # Celda del Logo
    if logo_bytes:
        run_logo = table.rows[0].cells[0].paragraphs[0].add_run()
        run_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
    
    # Celda de Datos del Alumno
    datos_cell = table.rows[0].cells[1]
    p_datos = datos_cell.paragraphs[0]
    p_datos.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_datos = p_datos.add_run(f"ESTUDIANTE: {nombre.upper()}\nEVALUACIÓN ADAPTADA")
    run_datos.bold = True
    run_datos.font.color.rgb = color_institucional
    run_datos.font.size = Pt(11)

    doc.add_paragraph() # Espacio

    # --- CONFIGURACIÓN DE FUENTE SEGÚN DIAGNÓSTICO ---
    style = doc.styles['Normal']
    font = style.font
    if "dislexia" in diag:
        font.name = 'Arial'; font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.5
    else:
        font.name = 'Verdana'; font.size = Pt(11)
        style.paragraph_format.line_spacing = 1.15

    # --- PROCESAMIENTO DE CONTENIDO ---
    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea: continue
        
        para = doc.add_paragraph()
        es_titulo = len(linea) < 60 and not linea.endswith('.')
        
        partes = linea.split("**")
        for i, parte in enumerate(partes):
            run = para.add_run(parte)
            if i % 2 != 0: run.bold = True
            
            if es_titulo:
                run.bold = True
                run.font.size = Pt(14)
                run.font.color.rgb = color_institucional
                para.space_before = Pt(12)

        # Espaciado extra para Discalculia
        if "discalculia" in diag and any(c.isdigit() for c in linea):
            doc.add_paragraph("\n" + "." * 80).runs[0].font.color.rgb = RGBColor(200, 200, 200)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ----------------------------
# 3. Interfaz de Streamlit
# ----------------------------
st.title("Centro de Adecuación Curricular v4.0 🍎")

# Carga de Planilla
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
try:
    df = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv")
    df.columns = [c.strip() for c in df.columns]
    
    col_grado = df.columns[1]
    col_nombre = df.columns[2]
    col_emergente = df.columns[4]

    grado = st.sidebar.selectbox("Seleccione Grado:", df[col_grado].unique())
    alumnos = df[(df[col_grado] == grado) & (df[col_emergente].str.lower() != "ninguna")]
    
    st.sidebar.metric("Alumnos detectados:", len(alumnos))
    logo_file = st.sidebar.file_uploader("Subir Logo de la Escuela", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None

    archivo_base = st.file_uploader("Subir Examen Base (PDF/DOCX)", type=["pdf", "docx"])

    if archivo_base and st.button(f"Generar carpeta para {grado}"):
        texto_base = extraer_texto(archivo_base)
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            progreso = st.progress(0)
            
            for i, (_, fila) in enumerate(alumnos.iterrows()):
                nombre, diag = fila[col_nombre], fila[col_emergente]
                
                prompt = f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} ({diag})\n\nEXAMEN:\n{texto_base}"
                
                time.sleep(2) # Respetar cuota API
                response = model.generate_content(prompt)
                
                doc_bytes = crear_docx_premium(response.text, nombre, diag, logo_bytes)
                zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_bytes.getvalue())
                
                progreso.progress((i + 1) / len(alumnos))
        
        st.success(f"Se generaron {len(alumnos)} exámenes.")
        st.download_button("Descargar ZIP con Adecuaciones", zip_buffer.getvalue(), f"Examenes_{grado}.zip")

except Exception as e:
    st.error(f"Error técnico: {e}")
