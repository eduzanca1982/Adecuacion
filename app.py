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
import re

# 1. IDENTIFICADOR DE PLANILLA
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"

# 2. CONFIGURACIÓN DE IA
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    model = genai.GenerativeModel('gemini-2.0-flash')
except:
    st.error("Error de configuración de API.")

SYSTEM_PROMPT = """Actúa como un Diseñador Editorial Pedagógico experto.
OBJETIVO: Adaptar exámenes manteniendo la JERARQUÍA y ESTÉTICA original.

REGLAS DE FORMATO:
1. Sin introducciones. Empieza con el título del examen.
2. JERARQUÍA: Usa [TITULO] para encabezados y [SUBTITULO] para secciones.
3. ICONOS: 📖 para lectura, 🔢 para cálculos, ✍️ para escritura.
4. IMÁGENES: Mantén la referencia con [MANTENER IMAGEN AQUÍ].

ADECUACIONES POR PERFIL:
- DISLEXIA/DISCALCULIA/DIFICULTAD GENERAL: Fuente OpenDyslexic, interlineado 1.5, verbos en negrita.
- TDAH: Consignas de un solo paso. Elimina decoraciones irrelevantes.
- GRUPO C: Mantén el desafío pero mejora la legibilidad visual."""

# 3. FUNCIONES DE APOYO
def extraer_texto(archivo):
    ext = archivo.name.split('.')[-1].lower()
    if ext == 'docx':
        doc = Document(archivo)
        return "\n".join([p.text for p in doc.paragraphs])
    elif ext == 'pdf':
        reader = PyPDF2.PdfReader(archivo)
        return "\n".join([p.extract_text() for p in reader.pages])
    return ""

def limpiar_nombre_archivo(nombre):
    # Elimina caracteres que Windows/Linux no permiten en nombres de archivo
    return re.sub(r'[\\/*?:"<>|]', "", nombre).replace(" ", "_")

def crear_docx_fiel(texto_ia, nombre, diagnostico, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    color_inst = RGBColor(31, 73, 125)

    # --- ENCABEZADO ---
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.6)
    
    if logo_bytes:
        try:
            run_logo = table.rows[0].cells[0].paragraphs[0].add_run()
            run_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(1.1))
        except: pass
    
    cell_info = table.rows[0].cells[1]
    p = cell_info.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"ALUMNO: {nombre.upper()}\nADECUACIÓN: {diagnostico.upper()}")
    run.bold = True
    run.font.color.rgb = color_inst
    run.font.size = Pt(11)

    doc.add_paragraph()

    # --- CONFIGURACIÓN DE FUENTE ---
    style = doc.styles['Normal']
    font = style.font
    if any(x in diag for x in ["dislexia", "discalculia", "general"]):
        font.name = 'OpenDyslexic'
        font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.5
    else:
        font.name = 'Verdana'
        font.size = Pt(11)
        style.paragraph_format.line_spacing = 1.15

    # --- RENDERIZADO ---
    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea: continue
        
        para = doc.add_paragraph()
        es_titulo = "[TITULO]" in linea or (len(linea) < 50 and not linea.endswith('.'))
        es_sub = "[SUBTITULO]" in linea
        
        linea_limpia = linea.replace("[TITULO]", "").replace("[SUBTITULO]", "").strip()
        
        partes = linea_limpia.split("**")
        for i, parte in enumerate(partes):
            run_p = para.add_run(parte)
            if i % 2 != 0: run_p.bold = True
            
            if es_titulo:
                run_p.bold = True
                run_p.font.size = Pt(14)
                run_p.font.color.rgb = color_inst
                para.space_before = Pt(12)
            elif es_sub:
                run_p.bold = True
                run_p.font.size = Pt(12)
                para.space_before = Pt(8)

        if ("discalculia" in diag or "general" in diag) and any(c.isdigit() for c in linea_limpia):
            doc.add_paragraph("\n" + "." * 65).runs[0].font.color.rgb = RGBColor(215, 215, 215)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 4. INTERFAZ
st.title("Motor de Adecuación v5.2 🎓")

try:
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    col_grado, col_nombre, col_emergente = df.columns[1], df.columns[2], df.columns[4]
    
    grado = st.sidebar.selectbox("Grado:", df[col_grado].unique())
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    
    alumnos = df[(df[col_grado] == grado) & (df[col_emergente].str.lower() != "ninguna")]
    
    archivo_base = st.file_uploader("Examen Original (DOCX o PDF)", type=["docx", "pdf"])

    if archivo_base and st.button(f"Procesar {len(alumnos)} alumnos"):
        texto_base = extraer_texto(archivo_base)
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            # --- PARTE NUEVA: AGREGAR EL ORIGINAL AL ZIP ---
            archivo_base.seek(0) # Volver al inicio del archivo
            zip_f.writestr(f"ORIGINAL_{archivo_base.name}", archivo_base.read())
            
            progreso = st.progress(0)
            status = st.empty()
            
            for i, (idx, fila) in enumerate(alumnos.iterrows()):
                nombre = str(fila[col_nombre])
                diag = str(fila[col_emergente])
                status.text(f"Adecuando: {nombre}...")
                
                success = False
                for intento in range(3):
                    try:
                        time.sleep(2 + (intento * 3))
                        p = f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} ({diag})\n\nEXAMEN:\n{texto_base}"
                        res = model.generate_content(p)
                        
                        doc_bytes = crear_docx_fiel(res.text, nombre, diag, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{limpiar_nombre_archivo(nombre)}.docx", doc_bytes.getvalue())
                        success = True
                        break
                    except Exception as e:
                        if "429" in str(e): continue
                        else: st.error(f"Error con {nombre}: {e}"); break
                
                progreso.progress((i + 1) / len(alumnos))
        
        st.success(f"¡Carpeta generada! Incluye el original y {len(alumnos)} adecuaciones.")
        st.download_button("Descargar ZIP Completo", zip_buffer.getvalue(), f"Examenes_{grado}.zip")

except Exception as e:
    st.error(f"Error: {e}")
