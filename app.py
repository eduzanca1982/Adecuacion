import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile
import time
import random
import re
import unicodedata

# 1. CONFIGURACIÓN Y UTILIDADES (Robustez ChatGPT)
st.set_page_config(page_title="Motor Pedagógico v10 🚀", layout="wide")
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"

def nrm(s: str) -> str:
    s = str(s or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def with_backoff(call, max_tries=6):
    """Reintento exponencial ante errores de cuota 429."""
    for attempt in range(max_tries):
        try:
            return call()
        except Exception as e:
            if "429" in str(e) and attempt < max_tries - 1:
                wait = (2 ** attempt) + random.uniform(0, 1)
                time.sleep(wait)
                continue
            raise e

# 2. PROMPT MAESTRO (Identidad Pedagógica Reforzada)
SYSTEM_PROMPT = """Eres un Diseñador Editorial para Inclusión Escolar. 
Tu tarea es ADECUAR el examen original de forma empática y clara.

REGLAS DE ORO:
1. FIDELIDAD: Copia cada ejercicio. NO los resuelvas.
2. EMOJIS: Usa emojis para hacer el examen más amigable (ej: 🔢 para matemática, 📖 para lectura).
3. PISTAS: Inserta 💡 en verde itálico debajo de las preguntas según la dificultad del alumno.
4. IMÁGENES: Inserta  para conceptos que necesiten apoyo visual.
5. RESALTE: **Negrita** solo para información nuclear de la respuesta. No resaltes conectores.
6. SILENCIO: Prohibido incluir análisis para la docente."""

# 3. GENERACIÓN DE IMÁGENES
def generar_imagen_v10(descripcion):
    try:
        model = genai.GenerativeModel("imagen-3.0")
        res = model.generate_content(f"Estilo educativo escolar, fondo blanco, minimalista: {descripcion}")
        return io.BytesIO(res.candidates[0].content.parts[0].inline_data.data)
    except:
        return None

# 4. CREACIÓN DEL DOCUMENTO (Diseño Inclusivo)
def crear_docx_v10(texto_ia, nombre, diagnostico, grupo, logo_bytes, gen_img):
    doc = Document()
    diag, grupo_v = str(diagnostico).lower(), str(grupo).upper()
    color_inst, color_pista = RGBColor(31, 73, 125), RGBColor(0, 102, 0)

    # Header Fiel
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try: table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    p = table.rows[0].cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diagnostico.upper()} | GRUPO: {grupo_v}").bold = True

    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo_v == "A"
    
    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea or any(x in linea.lower() for x in ["análisis:", "ayuda:"]): continue

        # Inserción de Imagen con IA
        if "[IMAGEN:" in linea and gen_img:
            desc = linea.split("[IMAGEN:")[1].split("]")[0]
            img_bytes = generar_imagen_v10(desc)
            if img_bytes:
                para_i = doc.add_paragraph()
                para_i.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para_i.add_run().add_picture(img_bytes, width=Inches(2.5))
            continue

        para = doc.add_paragraph()
        if "💡" in linea:
            run = para.add_run(linea)
            run.font.color.rgb, run.italic = color_pista, True
        else:
            partes = linea.split("**")
            for i, parte in enumerate(partes):
                run = para.add_run(parte)
                if i % 2 != 0: run.bold = True
                run.font.name = 'OpenDyslexic' if is_apo else 'Verdana'
                run.font.size = Pt(12 if is_apo else 11)

    bio = io.BytesIO()
    doc.save(bio)
    return bio

# 5. INTERFAZ STREAMLIT
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    # Detección inteligente de columnas
    col_grado = [c for c in df.columns if "grado" in nrm(c)][0]
    col_nombre = [c for c in df.columns if "nombre" in nrm(c)][0]
    col_grupo = [c for c in df.columns if "grupo" in nrm(c)][0]
    col_casos = [c for c in df.columns if "casos" in nrm(c) or "emergente" in nrm(c)][0]

    st.sidebar.header("Opciones de IA")
    grado_sel = st.sidebar.selectbox("Grado:", df[col_grado].unique())
    df_grado = df[df[col_grado] == grado_sel]
    
    seleccionados = st.sidebar.multiselect("Alumnos:", df_grado[col_nombre].tolist())
    alumnos_final = df_grado[df_grado[col_nombre].isin(seleccionados)] if seleccionados else df_grado

    activar_img = st.sidebar.checkbox("Generar Apoyos Visuales 🖼️", value=True)
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    archivo_base = st.file_uploader("Examen Original (docx)", type=["docx"])

    if archivo_base and st.button("🚀 Iniciar Adecuación"):
        from docx import Document as DocRead
        texto_base = "\n".join([p.text for p in DocRead(archivo_base).paragraphs])
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            progreso = st.progress(0)
            status = st.empty()
            
            for i, (_, fila) in enumerate(alumnos_final.iterrows()):
                nombre, diag, grupo = str(fila[col_nombre]), str(fila[col_casos]), str(fila[col_grupo])
                status.text(f"Adecuando: {nombre}...")
                
                def llamar_ia():
                    m = genai.GenerativeModel("gemini-1.5-flash")
                    return m.generate_content(f"{SYSTEM_PROMPT}\n\nALUMNO: {nombre} ({diag}, Grupo {grupo})\n\nEXAMEN:\n{texto_base}")

                res = with_backoff(llamar_ia)
                doc_res = crear_docx_v10(res.text, nombre, diag, grupo, logo_bytes, activar_img)
                zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_res.getvalue())
                progreso.progress((i + 1) / len(alumnos_final))

        st.success("¡Lote v10 finalizado!")
        st.download_button("📥 Descargar ZIP", zip_buffer.getvalue(), f"Adecuaciones_{grado_sel}.zip")

except Exception as e:
    st.error(f"Error de sistema: {e}")
