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

# 1. CONFIGURACIÓN Y ROBUSTEZ
st.set_page_config(page_title="Motor Pedagógico v10.2 🚀", layout="wide")
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"

def nrm(s: str) -> str:
    s = str(s or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def with_backoff(call, max_tries=6):
    for attempt in range(max_tries):
        try:
            return call()
        except Exception as e:
            if "429" in str(e) and attempt < max_tries - 1:
                time.sleep((2 ** attempt) + random.uniform(0, 1))
                continue
            raise e

# 2. INICIALIZACIÓN DINÁMICA (Adiós al 404)
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    modelos_disponibles = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    MODELO_TEXTO = next((m for m in modelos_disponibles if "gemini-1.5-flash" in m), modelos_disponibles[0])
    MODELO_IMG = next((m for m in modelos_disponibles if "imagen" in m), None)
except Exception as e:
    st.error(f"Error de conexión: {e}")
    st.stop()

# 3. PROMPT MAESTRO
SYSTEM_PROMPT = """Eres un Diseñador Editorial para Inclusión Escolar. 
1. TRANSCRIBE: Copia cada ejercicio original. NO los resuelvas.
2. EMOJIS: Usa emojis amigables (🔢, 📖, ✍️).
3. PISTAS: 💡 en verde itálico debajo de las preguntas.
4. IMÁGENES:  para apoyos visuales.
5. RESALTE: **Negrita** solo en evidencia nuclear. No resaltes conectores."""

# 4. FUNCIONES DE GENERACIÓN
def generar_imagen_ia(descripcion):
    if not MODELO_IMG: return None
    try:
        model = genai.GenerativeModel(MODELO_IMG)
        res = model.generate_content(f"Estilo dibujo escolar, fondo blanco: {descripcion}")
        return io.BytesIO(res.candidates[0].content.parts[0].inline_data.data)
    except: return None

def crear_docx_v10(texto_ia, nombre, diagnostico, grupo, logo_bytes, gen_img):
    doc = Document()
    diag, grupo_v = str(diagnostico).lower(), str(grupo).upper()
    color_inst, color_pista = RGBColor(31, 73, 125), RGBColor(0, 102, 0)

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

        if "[IMAGEN:" in linea and gen_img:
            desc = linea.split("[IMAGEN:")[1].split("]")[0]
            img_bytes = generar_imagen_ia(desc)
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

# 5. INTERFAZ
try:
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    col_grado = [c for c in df.columns if "grado" in nrm(c)][0]
    col_nombre = [c for c in df.columns if "nombre" in nrm(c)][0]
    col_grupo = [c for c in df.columns if "grupo" in nrm(c)][0]
    col_casos = [c for c in df.columns if "casos" in nrm(c) or "emergente" in nrm(c)][0]

    st.sidebar.header("🎯 Selección de Alumnos")
    grado_sel = st.sidebar.selectbox("Grado:", df[col_grado].unique())
    df_grado = df[df[col_grado] == grado_sel]
    
    # NUEVO: Lógica de selección de alumnos solicitada
    alcance = st.sidebar.radio("¿A quiénes adecuar?", ["Todos los alumnos", "Seleccionar cuáles"])
    
    if alcance == "Seleccionar cuáles":
        seleccionados = st.sidebar.multiselect("Elige los alumnos:", df_grado[col_nombre].tolist())
        alumnos_final = df_grado[df_grado[col_nombre].isin(seleccionados)]
    else:
        alumnos_final = df_grado

    st.sidebar.divider()
    activar_img = st.sidebar.checkbox("Generar Apoyos Visuales 🖼️", value=True)
    logo_file = st.sidebar.file_uploader("Logo", type=["png", "jpg"])
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
                n, d, g = str(fila[col_nombre]), str(fila[col_casos]), str(fila[col_grupo])
                status.text(f"Adecuando: {n}...")
                
                def llamar_ia():
                    m = genai.GenerativeModel(MODELO_TEXTO)
                    return m.generate_content(f"{SYSTEM_PROMPT}\n\nALUMNO: {n} ({d}, Grupo {g})\n\nEXAMEN:\n{texto_base}")

                res = with_backoff(llamar_ia)
                doc_res = crear_docx_v10(res.text, n, d, g, logo_bytes, activar_img)
                zip_f.writestr(f"Adecuacion_{n.replace(' ', '_')}.docx", doc_res.getvalue())
                progreso.progress((i + 1) / len(alumnos_final))

        st.success(f"Lote finalizado con modelo: {MODELO_TEXTO}")
        st.download_button("📥 Descargar ZIP", zip_buffer.getvalue(), "Adecuaciones.zip")

except Exception as e:
    st.error(f"Fallo en el sistema: {e}")
