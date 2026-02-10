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

# 1. CONFIGURACIÓN
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
OPCIONES_MODELOS = ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash"]

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
except Exception as e:
    st.error(f"Error de configuración: {e}")

# PROMPT MAESTRO: FOCO EN INTENCIONALIDAD DOCENTE
SYSTEM_PROMPT = """Eres un Psicopedagogo de apoyo en aula. Tu misión es analizar la intención de la docente en cada ejercicio y personalizar la ayuda.

PROCESO DE PENSAMIENTO:
1. Analiza cada pregunta: ¿Qué espera la docente que el alumno responda?
2. Diseña la ayuda: Basado en el perfil del alumno (Dislexia, TDAH, etc.), resalta en el texto original SOLAMENTE lo que ayuda a cumplir ese objetivo.
3. Personaliza la pista: Crea una [PISTA] breve que funcione como un andamio (no la respuesta, sino el cómo llegar).

REGLAS DE FORMATO:
- DISLEXIA: Resalta conectores y palabras clave. Fuente OpenDyslexic.
- TDAH: Segmenta y numera pasos. Pistas de "parar y revisar".
- MATEMÁTICA: Resalta datos numéricos y verbos operativos.
- ICONOS: 📖 (Lectura), 🔢 (Matemática), 💡 (Pista), ❓ (Pregunta).
- MANTÉN TÍTULOS CON [TITULO] Y LÍNEAS CON [CUADRICULA]."""

# 2. FUNCIONES EDITORIALES
def limpiar_nombre(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", str(nombre)).replace(" ", "_")

def crear_docx_fiel(texto_ia, nombre, diagnostico, grupo, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    grupo = str(grupo).upper()
    color_inst = RGBColor(31, 73, 125)
    color_pista = RGBColor(0, 102, 0)

    # Encabezado
    table = doc.add_table(rows=1, cols=2)
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

    # Fuente
    style = doc.styles['Normal']
    font = style.font
    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo == "A"
    font.name = 'OpenDyslexic' if is_apo else 'Verdana'
    font.size = Pt(12 if is_apo else 11)
    style.paragraph_format.line_spacing = 1.5 if is_apo else 1.15

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea: continue
        
        para = doc.add_paragraph()
        
        if "[PISTA]" in linea or "💡" in linea:
            txt = linea.replace("[PISTA]", "").replace("💡", "").strip()
            run_p = para.add_run(f"💡 PISTA: {txt}")
            run_p.font.color.rgb = color_pista
            run_p.italic = True
            continue

        if "[CUADRICULA]" in linea:
            for _ in range(3):
                p_g = doc.add_paragraph()
                p_g.add_run(" " + "." * 75).font.color.rgb = RGBColor(215, 215, 215)
            continue

        es_titulo = "[TITULO]" in linea or (len(linea) < 55 and not linea.endswith('.'))
        texto_limpio = linea.replace("[TITULO]", "").strip()
        
        partes = texto_limpio.split("**")
        for i, parte in enumerate(partes):
            run_part = para.add_run(parte)
            if i % 2 != 0: run_part.bold = True
            if es_titulo:
                run_part.bold = True
                run_part.font.size = Pt(14)
                run_part.font.color.rgb = color_inst

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 3. INTERFAZ
st.title("Motor Pedagógico v6.5 🎓")

try:
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    col_grado, col_nombre, col_grupo, col_emergente = df.columns[1], df.columns[2], df.columns[3], df.columns[4]
    
    st.sidebar.header("Inteligencia Artificial")
    modelo_ini = st.sidebar.selectbox("Modelo:", OPCIONES_MODELOS)
    grado_sel = st.sidebar.selectbox("Grado:", df[col_grado].unique())
    alumnos_grado = df[(df[col_grado] == grado_sel) & (df[col_emergente].str.lower() != "ninguna")]
    
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    archivo_base = st.file_uploader("Examen Original", type=["docx", "pdf"])

    if archivo_base and st.button(f"Adecuar {len(alumnos_grado)} alumnos"):
        from docx import Document as DocRead
        doc_read = DocRead(archivo_base)
        texto_base = "\n".join([p.text for p in doc_read.paragraphs])
        
        zip_buffer = io.BytesIO()
        cascada = [modelo_ini] + [m for m in OPCIONES_MODELOS if m != modelo_ini]

        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            archivo_base.seek(0)
            zip_f.writestr(f"ORIGINAL_{archivo_base.name}", archivo_base.read())
            
            progreso = st.progress(0)
            status = st.empty()

            for i, (_, fila) in enumerate(alumnos_grado.iterrows()):
                nombre, diag, grupo = str(fila[col_nombre]), str(fila[col_emergente]), str(fila[col_grupo])
                status.text(f"Analizando intención docente para: {nombre}...")
                
                success = False
                for m_name in cascada:
                    if success: break
                    try:
                        m_gen = genai.GenerativeModel(m_name)
                        time.sleep(2)
                        p_prompt = f"{SYSTEM_PROMPT}\n\nPERFIL ALUMNO: {nombre} ({diag}, Grupo {grupo})\n\nEXAMEN A ANALIZAR:\n{texto_base}"
                        res = m_gen.generate_content(p_prompt)
                        
                        doc_bytes = crear_docx_fiel(res.text, nombre, diag, grupo, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{limpiar_nombre(nombre)}.docx", doc_bytes.getvalue())
                        success = True
                        break
                    except: continue
                
                progreso.progress((i + 1) / len(alumnos_grado))

        st.success("Adecuación finalizada con éxito.")
        st.download_button("Descargar ZIP", zip_buffer.getvalue(), f"Examenes_{grado_sel}.zip")

except Exception as e:
    st.error(f"Error: {e}")
