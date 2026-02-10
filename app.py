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
import unicodedata

# 1. CONFIGURACIÓN DE SEGURIDAD Y CONEXIÓN
st.set_page_config(page_title="Motor Pedagógico v11.5 🚀", layout="wide")
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"

# Bypass de Seguridad: Evita bloqueos en contenidos escolares (falsos positivos)
SAFETY_SETTINGS = {
    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
}

# 2. UTILIDADES DE ROBUSTEZ (Detección de Columnas y Reintentos)
def nrm(s: str) -> str:
    s = str(s or "").strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

def with_backoff(call, max_tries=6):
    """Estrategia de reintento exponencial para eliminar el error 429."""
    for attempt in range(max_tries):
        try:
            return call()
        except Exception as e:
            if "429" in str(e) and attempt < max_tries - 1:
                wait = (2 ** attempt) + random.uniform(0, 1)
                time.sleep(wait)
                continue
            raise e

# 3. PROMPT MAESTRO DE RAZONAMIENTO (Cerebro de la IA)
SYSTEM_PROMPT = """Actúa como un Especialista en Neurodiversidad y Diseñador Instruccional. 
Tu misión es INTERVENIR el examen original para que el alumno pueda razonar la respuesta.

PROCESO DE PENSAMIENTO OBLIGATORIO:
1. ANALIZA: ¿Qué se evalúa en este ejercicio? (Ej: Multiplicación, Comprensión Lectora).
2. RAZONA: Resuelve tú mismo el ejercicio internamente.
3. ADECUACIÓN POR GRUPO:
   - GRUPO A (Andamiaje Intenso): El alumno requiere apoyos visuales y lenguaje ultra-simple. Usa ejemplos de la vida real.
   - GRUPO B (Andamiaje Moderado): Pistas de proceso. Divide tareas largas en pasos.
   - GRUPO C (Metacognición): Desafíos para evitar el aburrimiento. Pistas de revisión.

REGLAS DE DISEÑO:
- EMOJIS: Usa iconos como anclas visuales (🔢 Matemática, 📖 Lectura, ✍️ Escribir).
- PISTAS 💡: Deben ser de razonamiento, nunca des la respuesta.
- IMÁGENES: Usa la etiqueta  para representar el proceso mental.
- RESALTE: **Negrita** solo en datos y verbos de acción. No resaltes conectores.
- SILENCIO: Prohibido saludar o dar introducciones como "Aquí tienes el examen"."""

# 4. FUNCIONES DE GENERACIÓN DE CONTENIDO
def generar_imagen_ia(descripcion):
    try:
        model = genai.GenerativeModel("imagen-3.0")
        res = model.generate_content(f"Estilo dibujo escolar, fondo blanco, minimalista: {descripcion}", safety_settings=SAFETY_SETTINGS)
        return io.BytesIO(res.candidates[0].content.parts[0].inline_data.data)
    except: return None

def crear_docx_v11(texto_ia, nombre, diagnostico, grupo, logo_bytes, gen_img):
    doc = Document()
    diag, grupo_v = str(diagnostico).lower(), str(grupo).upper()
    color_inst, color_pista = RGBColor(31, 73, 125), RGBColor(0, 102, 0)

    # Header Fiel al Motor
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try: table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    p = table.rows[0].cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diagnostico.upper()} | GRUPO: {grupo_v}").bold = True

    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo_v == "A"
    
    # Limpieza de "charla" de la IA
    texto_filtrado = re.sub(r"^(¡Claro|Hola|Aquí tienes|Entendido|Como maquetador).*?\n", "", texto_ia, flags=re.IGNORECASE)

    for linea in texto_filtrado.split('\n'):
        linea = linea.strip()
        if not linea or "análisis:" in linea.lower(): continue

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

# 5. LÓGICA DE INTERFAZ Y PROCESAMIENTO
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    # Escaneo de modelos para evitar Error 404
    modelos = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    MODELO_IA = next((m for m in modelos if "gemini-1.5-flash" in m), modelos[0])

    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    # Detección inteligente de columnas
    col_grado = [c for c in df.columns if "grado" in nrm(c)][0]
    col_nombre = [c for c in df.columns if "nombre" in nrm(c)][0]
    col_casos = [c for c in df.columns if "casos" in nrm(c) or "emergente" in nrm(c)][0]
    col_grupo = [c for c in df.columns if "grupo" in nrm(c)][0]

    st.sidebar.header("🎯 Selección de Alumnos")
    grado_sel = st.sidebar.selectbox("Grado:", df[col_grado].unique())
    df_grado = df[df[col_grado] == grado_sel]
    
    alcance = st.sidebar.radio("¿A quiénes adecuar?", ["Todos los alumnos", "Seleccionar cuáles"])
    alumnos_final = df_grado if alcance == "Todos los alumnos" else df_grado[df_grado[col_nombre].isin(st.sidebar.multiselect("Elige los alumnos:", df_grado[col_nombre].tolist()))]

    st.sidebar.divider()
    activar_img = st.sidebar.checkbox("Generar Apoyos Visuales 🖼️", value=True)
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    archivo_base = st.file_uploader("Examen Original (docx)", type=["docx"])

    if archivo_base and not alumnos_final.empty and st.button("🚀 Iniciar Procesamiento v11.5"):
        from docx import Document as DocRead
        texto_base = "\n".join([p.text for p in DocRead(archivo_base).paragraphs])
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            progreso = st.progress(0)
            for i, (_, fila) in enumerate(alumnos_final.iterrows()):
                n, d, g = str(fila[col_nombre]), str(fila[col_casos]), str(fila[col_grupo])
                try:
                    def llamar_ia():
                        m = genai.GenerativeModel(MODELO_IA)
                        return m.generate_content(f"{SYSTEM_PROMPT}\nALUMNO: {n} ({d}, Grupo {g})\nEXAMEN:\n{texto_base}", safety_settings=SAFETY_SETTINGS)
                    
                    res = with_backoff(llamar_ia)
                    doc_res = crear_docx_v11(res.text, n, d, g, logo_bytes, activar_img)
                    zip_f.writestr(f"Adecuacion_{n.replace(' ', '_')}.docx", doc_res.getvalue())
                except Exception as e:
                    st.sidebar.error(f"Fallo en {n}: {e}")
                progreso.progress((i + 1) / len(alumnos_final))

        st.success("Lote completado exitosamente.")
        st.download_button("📥 Descargar ZIP", zip_buffer.getvalue(), f"Adecuaciones_{grado_sel}.zip")

except Exception as e:
    st.error(f"Fallo técnico: {e}")
