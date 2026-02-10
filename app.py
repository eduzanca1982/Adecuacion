import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import PyPDF2
import io
import zipfile
import time

# 1. IDENTIFICADOR DE PLANILLA (Asegúrate que sea el correcto)
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"

# 2. CONFIGURACIÓN DE IA
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    model = genai.GenerativeModel('gemini-2.0-flash')
except:
    st.error("Revisa la API KEY en los Secrets de Streamlit.")

SYSTEM_PROMPT = """Eres un Diseñador Instruccional Pedagógico. Tu misión es adaptar exámenes para primaria.
REGLAS DE DISEÑO:
1. Sin introducciones. Empieza con el título del examen.
2. MANTÉN la numeración original (1, 1.a, 2...).
3. ICONOGRAFÍA: Si es Dislexia, usa [LECTURA] antes de textos largos. Si es Discalculia, usa [CÁLCULO] antes de operaciones.
4. MANTÉN referencias a imágenes con: [VER IMAGEN/GRÁFICO AQUÍ].
5. Si es completar, usa: ................................

ADECUACIONES:
- DISLEXIA: Frases cortas, vocabulario simple, verbos en negrita.
- DISCALCULIA: Datos en listas, espacios muy amplios para resolver.
- TDAH: Instrucciones de un solo paso. Dividir tareas largas."""

# 3. FUNCIONES TÉCNICAS
def extraer_texto(archivo):
    ext = archivo.name.split('.')[-1].lower()
    if ext == 'docx':
        doc = Document(archivo)
        return "\n".join([p.text for p in doc.paragraphs])
    elif ext == 'pdf':
        reader = PyPDF2.PdfReader(archivo)
        return "\n".join([p.extract_text() for p in reader.pages])
    return ""

def crear_docx_premium(texto_ia, nombre, diagnostico, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    color_inst = RGBColor(31, 73, 125) # Azul Institucional

    # --- ENCABEZADO ESTÉTICO ---
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.5)
    
    # Celda Logo
    if logo_bytes:
        try:
            run_logo = table.rows[0].cells[0].paragraphs[0].add_run()
            run_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(1.2))
        except: pass
    
    # Celda Info Alumno
    cell_info = table.rows[0].cells[1]
    p = cell_info.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"ESTUDIANTE: {nombre.upper()}\nEVALUACIÓN ADAPTADA")
    run.bold = True
    run.font.color.rgb = color_inst
    run.font.size = Pt(12)

    doc.add_paragraph() # Espacio

    # --- LÓGICA DE FUENTE ACCESIBLE ---
    style = doc.styles['Normal']
    font = style.font
    # Prioridad OpenDyslexic (requiere instalación en PC destino)
    if "dislexia" in diag or "discalculia" in diag:
        font.name = 'OpenDyslexic'
        font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.5
    else:
        font.name = 'Verdana'
        font.size = Pt(11)
        style.paragraph_format.line_spacing = 1.15

    # --- PROCESAMIENTO DE TEXTO ---
    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea: continue
        
        para = doc.add_paragraph()
        
        # Iconos visuales según contexto
        if "[LECTURA]" in linea: linea = "📖 " + linea.replace("[LECTURA]", "")
        if "[CÁLCULO]" in linea: linea = "🔢 " + linea.replace("[CÁLCULO]", "")
        
        # Títulos
        es_titulo = (len(linea) < 60 and not linea.endswith('.')) or linea.isupper()
        
        partes = linea.split("**")
        for i, parte in enumerate(partes):
            run_p = para.add_run(parte)
            if i % 2 != 0: run_p.bold = True
            
            if es_titulo:
                run_p.bold = True
                run_p.font.size = Pt(14)
                run_p.font.color.rgb = color_inst
                para.space_before = Pt(12)

        # Espaciado extra para Discalculia (Cuadrícula visual sutil)
        if "discalculia" in diag and any(c.isdigit() for c in linea):
            for _ in range(2):
                p_espacio = doc.add_paragraph()
                p_espacio.add_run(" " + "." * 60).font.color.rgb = RGBColor(220, 220, 220)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 4. INTERFAZ DE USUARIO
st.title("Motor de Adecuación v5.0 Premium 🎓")

try:
    SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(SHEET_URL)
    df.columns = [c.strip() for c in df.columns]
    
    idx_grado, idx_nombre, idx_emergente = 1, 2, 4
    
    grado = st.sidebar.selectbox("Seleccione Grado:", df.iloc[:, idx_grado].unique())
    logo_file = st.sidebar.file_uploader("Subir Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    
    alumnos = df[(df.iloc[:, idx_grado] == grado) & (df.iloc[:, idx_emergente].str.lower() != "ninguna")]
    st.sidebar.metric("Alumnos a adecuar", len(alumnos))

    archivo_base = st.file_uploader("Subir Examen (DOCX o PDF)", type=["docx", "pdf"])

    if archivo_base and st.button(f"Generar exámenes para todo {grado}"):
        texto_base = extraer_texto(archivo_base)
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            progreso = st.progress(0)
            status = st.empty()
            
            for i, (idx, fila) in enumerate(alumnos.iterrows()):
                nombre = fila.iloc[idx_nombre]
                diag = fila.iloc[idx_emergente]
                
                status.text(f"Adecuando para: {nombre}...")
                
                # Manejo de Reintentos para Error 429
                success = False
                for intento in range(3):
                    try:
                        time.sleep(2 + (intento * 3))
                        prompt = f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} ({diag})\n\nEXAMEN ORIGINAL:\n{texto_base}"
                        response = model.generate_content(prompt)
                        
                        doc_bytes = crear_docx_premium(response.text, nombre, diag, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_bytes.getvalue())
                        success = True
                        break
                    except Exception as e:
                        if "429" in str(e): continue
                        else: st.error(f"Error con {nombre}: {e}"); break
                
                progreso.progress((i + 1) / len(alumnos))
        
        st.success("¡Proceso completado con éxito!")
        st.download_button("⬇️ Descargar Archivo ZIP", zip_buffer.getvalue(), f"Adecuaciones_{grado}.zip")

except Exception as e:
    st.error(f"Error técnico: {e}")
