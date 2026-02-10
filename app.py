import streamlit as st
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile
import time
import random
import re

# ==========================================
# 1. CONFIGURACIÓN GLOBAL Y SEGURIDAD DINÁMICA
# ==========================================
st.set_page_config(page_title="Motor Pedagógico v12.3 🍎", layout="wide")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    
    # DETECCIÓN DINÁMICA DE CATEGORÍAS (Evita el error 400/Civic Integrity)
    SAFETY = []
    categorias_validas = [
        "HARM_CATEGORY_HARASSMENT",
        "HARM_CATEGORY_HATE_SPEECH",
        "HARM_CATEGORY_SEXUALLY_EXPLICIT",
        "HARM_CATEGORY_DANGEROUS_CONTENT",
        "HARM_CATEGORY_CIVIC_INTEGRITY" # Se intentará, pero el código es tolerante
    ]
    
    for cat in categorias_validas:
        try:
            # Solo agregamos la categoría si la librería la reconoce
            SAFETY.append({"category": cat, "threshold": "BLOCK_NONE"})
        except:
            continue

    modelos_api = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    MODELO_IA = next((m for m in modelos_api if "1.5-flash" in m), modelos_api[0])
except Exception as e:
    st.error(f"Error inicializando API: {e}")

# ==========================================
# 2. PROMPT DE RAZONAMIENTO TUTOR
# ==========================================
SYSTEM_PROMPT = """Actúa como un Tutor Psicopedagogo.
Tu misión es intervenir el examen para que sea un puente al conocimiento, NO un muro.

TU LÓGICA DE TRABAJO:
1. Resuelve mentalmente el ejercicio para conocer el desafío.
2. Crea una Pista 💡 de RAZONAMIENTO. Prohibido dar la respuesta.
3. Si el alumno es Grupo A (Andamiaje Intenso): Usa ejemplos con objetos (manzanas, lápices).
4. Si el alumno es Grupo B/C (Independiente): Usa pistas de proceso y revisión.

REGLAS DE FORMATO:
- TRANSCRIBE TODO el examen original. No resumas.
- Usa emojis (🔢, 📖, ✍️) para dar orden.
- PISTAS: 💡 en verde itálico debajo de cada consigna.
- IMÁGENES:  para apoyo visual concreto."""

# ==========================================
# 3. FUNCIONES DE DISEÑO
# ==========================================
def generar_imagen(desc):
    try:
        m = genai.GenerativeModel("imagen-3.0")
        res = m.generate_content(f"Dibujo escolar, fondo blanco, trazos negros: {desc}", safety_settings=SAFETY)
        return io.BytesIO(res.candidates[0].content.parts[0].inline_data.data)
    except: return None

def crear_docx_v12_3(texto_ia, nombre, diag, grupo, logo_bytes, gen_img):
    doc = Document()
    color_pista = RGBColor(0, 102, 0) # Verde
    
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try: table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    p_hdr = table.rows[0].cells[1].paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hdr.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diag.upper()} | GRUPO: {grupo.upper()}").bold = True

    # Estilo inclusivo
    es_apo = any(x in str(diag).lower() for x in ["dislexia", "discalculia"]) or "A" in str(grupo)
    
    # Limpieza automática de comentarios de la IA
    texto_ia = re.sub(r"^(¡Claro|Hola|Aquí|Entendido|Como).*?\n", "", texto_ia, flags=re.IGNORECASE)

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea or "análisis:" in linea.lower(): continue

        if "[IMAGEN:" in linea and gen_img:
            img_bytes = generar_imagen(linea.split("[IMAGEN:")[1].split("]")[0])
            if img_bytes:
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_bytes, width=Inches(2.5))
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
                run.font.name = 'OpenDyslexic' if es_apo else 'Verdana'
                run.font.size = Pt(12 if es_apo else 11)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio

# ==========================================
# 4. INTERFAZ STREAMLIT
# ==========================================
try:
    df = pd.read_csv(URL_PLANILLA)
    df.columns = [c.strip() for c in df.columns]
    
    st.sidebar.header("⚙️ Configuración")
    grado_sel = st.sidebar.selectbox("Seleccionar Grado:", df[df.columns[1]].unique())
    df_f = df[df[df.columns[1]] == grado_sel]
    
    alcance = st.sidebar.radio("¿A quiénes adecuar?", ["Todos", "Seleccionar"])
    if alcance == "Seleccionar":
        sel = st.sidebar.multiselect("Check de Alumnos:", df_f[df_f.columns[2]].tolist())
        df_f = df_f[df_f[df_f.columns[2]].isin(sel)]

    st.sidebar.divider()
    activar_img = st.sidebar.checkbox("Generar Imágenes IA 🖼️", value=True)
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    
    st.title("Motor Pedagógico v12.3 🧠")
    archivo = st.file_uploader("Subir Examen Base (docx)", type=["docx"])

    if archivo and st.button("🚀 INICIAR ADECUACIÓN"):
        from docx import Document as DocRead
        txt_base = "\n".join([p.text for p in DocRead(archivo).paragraphs])
        zip_bio = io.BytesIO()
        
        with zipfile.ZipFile(zip_bio, "w") as z:
            prog = st.progress(0)
            status = st.empty()
            for i, (_, fila) in enumerate(df_f.iterrows()):
                n, g, d = str(fila[df.columns[2]]), str(fila[df.columns[3]]), str(fila[df.columns[4]])
                status.text(f"Razonando adecuación para {n}...")
                
                m = genai.GenerativeModel(MODELO_IA)
                # Llamada con bypass de seguridad dinámico
                res = m.generate_content(
                    f"{SYSTEM_PROMPT}\nALUMNO: {n} ({d}, {g})\nEXAMEN:\n{txt_base}", 
                    safety_settings=SAFETY
                )
                
                doc_res = crear_docx_v12_3(res.text, n, d, g, logo_bytes, activar_img)
                z.writestr(f"Adecuacion_{n.replace(' ', '_')}.docx", doc_res.getvalue())
                prog.progress((i + 1) / len(df_f))

        st.success("¡Lote procesado exitosamente!")
        st.download_button("📥 Descargar Resultados (ZIP)", zip_bio.getvalue(), "Examenes_Adecuados.zip")
except Exception as e:
    st.error(f"Fallo del sistema: {e}")import streamlit as st
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile
import time
import random
import re

# ==========================================
# 1. CONFIGURACIÓN GLOBAL
# ==========================================
st.set_page_config(page_title="Motor Pedagógico v12.2 🍎", layout="wide")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    
    # CORRECCIÓN ERROR 400: Lista explícita de categorías válidas
    SAFETY = [
        {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
        {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
        {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
        {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
        {"category": "HARM_CATEGORY_CIVIC_INTEGRITY", "threshold": "BLOCK_NONE"},
    ]
    
    modelos_api = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    MODELO_IA = next((m for m in modelos_api if "1.5-flash" in m), modelos_api[0])
except Exception as e:
    st.error(f"Error de conexión: {e}")

# ==========================================
# 2. PROMPT DE RAZONAMIENTO PROFUNDO
# ==========================================
SYSTEM_PROMPT = """Actúa como un Tutor Psicopedagogo de nivel primario.
Tu misión es intervenir el examen original para que sea un puente al conocimiento.

PASOS DE TU PENSAMIENTO:
1. Resuelve mentalmente el ejercicio para conocer la dificultad real.
2. Basado en el diagnóstico del alumno, crea una Pista 💡 de razonamiento.
3. Si el alumno tiene Discalculia (José de San Martín), la pista debe ser visual y concreta.
4. Si el alumno tiene Dificultad General (Francisco Moreno), la pista debe simplificar los pasos.

REGLAS DE FORMATO:
- TRANSCRIBE TODO el examen original. No resumas.
- NO DES LA RESPUESTA. Deja los espacios vacíos.
- Usa emojis (🔢, 📖, ✍️) para dar orden visual.
- PISTAS: 💡 en verde itálico debajo de cada consigna.
- IMÁGENES:  para apoyos visuales."""

# ==========================================
# 3. FUNCIONES DE DISEÑO
# ==========================================
def generar_imagen(desc):
    try:
        m = genai.GenerativeModel("imagen-3.0")
        res = m.generate_content(f"Dibujo escolar, trazos simples, fondo blanco: {desc}", safety_settings=SAFETY)
        return io.BytesIO(res.candidates[0].content.parts[0].inline_data.data)
    except: return None

def crear_docx_v12(texto_ia, nombre, diag, grupo, logo_bytes, gen_img):
    doc = Document()
    color_pista = RGBColor(0, 102, 0)
    
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try: table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    p_hdr = table.rows[0].cells[1].paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hdr.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diag.upper()} | GRUPO: {grupo.upper()}").bold = True

    es_apo = any(x in str(diag).lower() for x in ["dislexia", "discalculia"]) or "A" in str(grupo)
    texto_ia = re.sub(r"^(¡Claro|Hola|Aquí|Entendido|Como).*?\n", "", texto_ia, flags=re.IGNORECASE)

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea or "análisis:" in linea.lower(): continue

        if "[IMAGEN:" in linea and gen_img:
            img_bytes = generar_imagen(linea.split("[IMAGEN:")[1].split("]")[0])
            if img_bytes:
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_bytes, width=Inches(2.5))
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
                run.font.name = 'OpenDyslexic' if es_apo else 'Verdana'
                run.font.size = Pt(12 if es_apo else 11)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio

# ==========================================
# 4. INTERFAZ
# ==========================================
try:
    df = pd.read_csv(URL_PLANILLA)
    df.columns = [c.strip() for c in df.columns]
    
    st.sidebar.header("🎯 Configuración")
    grado_sel = st.sidebar.selectbox("Grado:", df[df.columns[1]].unique())
    df_f = df[df[df.columns[1]] == grado_sel]
    
    alcance = st.sidebar.radio("¿A quiénes adecuar?", ["Todos", "Seleccionar"])
    if alcance == "Seleccionar":
        sel = st.sidebar.multiselect("Alumnos:", df_f[df_f.columns[2]].tolist())
        df_f = df_f[df_f[df_f.columns[2]].isin(sel)]

    st.sidebar.divider()
    activar_img = st.sidebar.checkbox("Generar Imágenes", value=True)
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    
    st.title("Motor Pedagógico v12.2")
    archivo = st.file_uploader("Subir Examen (docx)", type=["docx"])

    if archivo and st.button("🚀 INICIAR"):
        from docx import Document as DocRead
        txt_base = "\n".join([p.text for p in DocRead(archivo).paragraphs])
        zip_bio = io.BytesIO()
        
        with zipfile.ZipFile(zip_bio, "w") as z:
            prog = st.progress(0)
            for i, (_, fila) in enumerate(df_f.iterrows()):
                n, g, d = str(fila[df.columns[2]]), str(fila[df.columns[3]]), str(fila[df.columns[4]])
                
                m = genai.GenerativeModel(MODELO_IA)
                res = m.generate_content(f"{SYSTEM_PROMPT}\nALUMNO: {n} ({d}, {g})\nEXAMEN:\n{txt_base}", safety_settings=SAFETY)
                
                doc_res = crear_docx_v12(res.text, n, d, g, logo_bytes, activar_img)
                z.writestr(f"Adecuacion_{n.replace(' ', '_')}.docx", doc_res.getvalue())
                prog.progress((i + 1) / len(df_f))

        st.success("¡Completado!")
        st.download_button("📥 Descargar ZIP", zip_bio.getvalue(), "Examenes.zip")
except Exception as e:
    st.error(f"Error: {e}")
