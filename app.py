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
    st.error(f"Falta GOOGLE_API_KEY: {e}")

# PROMPT MAESTRO CON LÓGICA DE PISTAS
SYSTEM_PROMPT = """Eres un Psicopedagogo experto. Tu objetivo es el andamiaje cognitivo.

REGLAS DE APOYO (SOLO PARA GRUPO A O DIFICULTADES):
1. PISTAS PEDAGÓGICAS: Para problemas de matemáticas o preguntas de comprensión, añade una breve "Pista" o "Idea" que ayude a iniciar el proceso (ej: "Para resolver esto, pensá si tenés que agregar o quitar").
2. AYUDA MEMORIA: Si el ejercicio requiere un concepto específico, inclúyelo de forma simple (ej: "Recordá: 1 metro = 100 cm").
3. RESALTE DE RESPUESTAS: En los textos, marca en **negrita** las oraciones donde está la respuesta.
4. MATEMÁTICAS: Marca datos numéricos en **negrita** y la pregunta central con ❓.
5. FORMATO: Usa [PISTA] para estas ayudas y [CUADRICULA] para espacios de resolución.

GENERAL:
- No uses introducciones. Títulos con [TITULO].
- Iconos: 📖 (Lectura), 🔢 (Cálculo), 💡 (Pista)."""

# 2. FUNCIONES DE MAQUETACIÓN PREMIUM
def limpiar_nombre(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", str(nombre)).replace(" ", "_")

def crear_docx_andamiaje(texto_ia, nombre, diagnostico, grupo, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    grupo = str(grupo).upper()
    color_inst = RGBColor(31, 73, 125)
    color_pista = RGBColor(0, 102, 0) # Verde oscuro pedagógico

    # Encabezado con Tabla
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try:
            run_logo = table.rows[0].cells[0].paragraphs[0].add_run()
            run_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(1.1))
        except: pass
    
    cell_info = table.rows[0].cells[1]
    p = cell_info.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"ALUMNO: {nombre.upper()}\nGRUPO: {grupo} | APOYO: {diagnostico.upper()}")
    run.bold = True
    run.font.color.rgb = color_inst

    # Configuración de fuente accesible
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
        
        # Tratamiento de PISTAS (💡)
        if "[PISTA]" in linea or "💡" in linea:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run_p = para.add_run(linea.replace("[PISTA]", "💡 PISTA: "))
            run_p.font.color.rgb = color_pista
            run_p.italic = True
            continue

        # Tratamiento de CUADRICULA
        if "[CUADRICULA]" in linea:
            para.add_run("\n" + "." * 70 + "\n" + "." * 70).font.color.rgb = RGBColor(210, 210, 210)
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

# 3. INTERFAZ STREAMLIT
st.title("Motor de Adecuación v6.2 Premium 🎓💡")

try:
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    col_grado = df.columns[1]
    col_nombre = df.columns[2]
    col_grupo = df.columns[3]
    col_emergente = df.columns[4]
    
    st.sidebar.header("Control de IA")
    modelo_ini = st.sidebar.selectbox("Modelo Principal:", OPCIONES_MODELOS)
    grado_sel = st.sidebar.selectbox("Seleccionar Grado:", df[col_grado].unique())
    alumnos_grado = df[(df[col_grado] == grado_sel) & (df[col_emergente].str.lower() != "ninguna")]
    
    logo_file = st.sidebar.file_uploader("Logo Institucional", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    archivo_base = st.file_uploader("Subir Examen (DOCX/PDF)", type=["docx", "pdf"])

    if archivo_base and st.button(f"Procesar Grado ({len(alumnos_grado)} alumnos)"):
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
                nombre = str(fila[col_nombre])
                diag = str(fila[col_emergente])
                grupo = str(fila[col_grupo])
                status.text(f"Generando andamiaje para: {nombre}...")
                
                success = False
                for m_name in cascada:
                    if success: break
                    try:
                        m_gen = genai.GenerativeModel(m_name)
                        time.sleep(2)
                        p_prompt = f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} (Grupo {grupo}, {diag})\n\nEXAMEN:\n{texto_base}"
                        res = m_gen.generate_content(p_prompt)
                        
                        doc_bytes = crear_docx_andamiaje(res.text, nombre, diag, grupo, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{limpiar_nombre(nombre)}.docx", doc_bytes.getvalue())
                        success = True
                        break
                    except: continue
                
                progreso.progress((i + 1) / len(alumnos_grado))

        st.success("¡Adecuación con Pistas completada!")
        st.download_button("Descargar ZIP", zip_buffer.getvalue(), f"Examenes_{grado_sel}.zip")

except Exception as e:
    st.error(f"Error técnico: {e}")
