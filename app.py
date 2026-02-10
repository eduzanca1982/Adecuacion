import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile
import time
import re

# 1. CONFIGURACIÓN ESTRUCTURAL
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
OPCIONES_MODELOS = ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash"]

st.set_page_config(page_title="Motor Pedagógico v6.8", layout="centered")

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    model_ai = genai.GenerativeModel('gemini-2.0-flash')
except Exception as e:
    st.error(f"Falta configurar la API KEY: {e}")

# PROMPT MAESTRO (Control de Calidad)
SYSTEM_PROMPT = """Eres un Diseñador Editorial Pedagógico. Genera el examen FINAL.

REGLAS DE SILENCIO:
1. PROHIBIDO incluir análisis, introducciones o explicaciones.
2. NO resaltes conectores (y, con, por, de, fue, el, la).
3. SÓLO resalta en **negrita** la información nuclear de la respuesta.

REGLAS DE ESPACIO:
1. SOLO usa [CUADRICULA] donde el alumno deba escribir.
2. NO agregues líneas de puntos al azar."""

# 2. FUNCIONES DE DISEÑO
def limpiar_nombre_archivo(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", str(nombre)).replace(" ", "_")

def extraer_texto(archivo):
    ext = archivo.name.split('.')[-1].lower()
    if ext == 'docx':
        doc = Document(archivo)
        return "\n".join([p.text for p in doc.paragraphs])
    elif ext == 'pdf':
        import PyPDF2
        reader = PyPDF2.PdfReader(archivo)
        return "\n".join([p.extract_text() for p in reader.pages])
    return ""

def crear_docx_final(texto_ia, nombre, diagnostico, grupo, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    grupo = str(grupo).upper()
    color_inst = RGBColor(31, 73, 125)

    # Encabezado con Tabla
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.5)
    if logo_bytes:
        try:
            run_logo = table.rows[0].cells[0].paragraphs[0].add_run()
            run_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    
    cell_info = table.rows[0].cells[1]
    p = cell_info.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"ESTUDIANTE: {nombre.upper()}\nAPOYO: {diagnostico.upper()} | GRUPO: {grupo}")
    run.bold = True
    run.font.color.rgb = color_inst

    # Fuente OpenDyslexic
    style = doc.styles['Normal']
    font = style.font
    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo == "A"
    font.name = 'OpenDyslexic' if is_apo else 'Verdana'
    font.size = Pt(12 if is_apo else 11)
    style.paragraph_format.line_spacing = 1.5 if is_apo else 1.15

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        # Filtro de basura (Análisis e intros)
        if any(x in linea.lower() for x in ["análisis:", "ayuda:", "aquí tienes", "analisis:"]): continue
        if re.match(r'^\.*$', linea): continue
        if not linea: continue
        
        para = doc.add_paragraph()
        if "[CUADRICULA]" in linea:
            for _ in range(2):
                p_g = doc.add_paragraph()
                p_g.add_run(" " + "." * 70).font.color.rgb = RGBColor(215, 215, 215)
            continue

        if "💡" in linea:
            run_p = para.add_run(linea)
            run_p.font.color.rgb = RGBColor(0, 102, 0)
            run_p.italic = True
            continue

        es_titulo = "[TITULO]" in linea or (len(linea) < 55 and not linea.endswith('.'))
        texto_limpio = linea.replace("[TITULO]", "").strip()
        partes = texto_limpio.split("**")
        for i, parte in enumerate(partes):
            run_part = para.add_run(parte)
            if i % 2 != 0: run_part.bold = True
            if es_titulo:
                run_part.bold = True
                run_part.font.size = Pt(13)
                run_part.font.color.rgb = color_inst

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 3. INTERFAZ (RESTURADA)
st.title("Motor Pedagógico v6.8 🎓")

try:
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    col_grado, col_nombre, col_grupo, col_emergente = df.columns[1], df.columns[2], df.columns[3], df.columns[4]
    
    st.sidebar.header("Control de IA")
    modelo_ini = st.sidebar.selectbox("Modelo:", OPCIONES_MODELOS)
    grado_sel = st.sidebar.selectbox("Grado:", df[col_grado].unique())
    alumnos_grado = df[(df[col_grado] == grado_sel) & (df[col_emergente].str.lower() != "ninguna")]
    
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    
    archivo_base = st.file_uploader("Subir Examen Original", type=["docx", "pdf"])

    if archivo_base and st.button(f"Adecuar {len(alumnos_grado)} alumnos"):
        texto_base = extraer_texto(archivo_base)
        zip_buffer = io.BytesIO()
        cascada = [modelo_ini] + [m for m in OPCIONES_MODELOS if m != modelo_ini]

        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            archivo_base.seek(0)
            zip_f.writestr(f"ORIGINAL_{archivo_base.name}", archivo_base.read())
            
            progreso = st.progress(0)
            status = st.empty()

            for i, (_, fila) in enumerate(alumnos_grado.iterrows()):
                nombre, diag, grupo = str(fila[col_nombre]), str(fila[col_emergente]), str(fila[col_grupo])
                status.text(f"Procesando: {nombre}...")
                
                success = False
                for m_name in cascada:
                    if success: break
                    try:
                        m_gen = genai.GenerativeModel(m_name)
                        time.sleep(2)
                        prompt = f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} ({diag}, Grupo {grupo})\n\nEXAMEN:\n{texto_base}"
                        res = m_gen.generate_content(prompt)
                        doc_bytes = crear_docx_final(res.text, nombre, diag, grupo, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{limpiar_nombre_archivo(nombre)}.docx", doc_bytes.getvalue())
                        success = True
                    except: continue
                
                progreso.progress((i + 1) / len(alumnos_grado))

        st.success("ZIP generado correctamente.")
        st.download_button("Descargar ZIP", zip_buffer.getvalue(), f"Examenes_{grado_sel}.zip")

except Exception as e:
    st.error(f"Error en la carga: {e}")
