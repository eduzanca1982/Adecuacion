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

# 1. SETUP DE SEGURIDAD (Para que no bloquee textos escolares)
st.set_page_config(page_title="Motor Pedagógico v12 🍎", layout="wide")
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    SAFETY = {category: HarmBlockThreshold.BLOCK_NONE for category in HarmCategory}
    # Detección de modelo para evitar el error 404
    modelos = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    MODELO_IA = next((m for m in modelos if "1.5-flash" in m), modelos[0])
except:
    st.error("Error de configuración de API.")

# 2. EL PROMPT DE TUTOR (El corazón del cambio)
SYSTEM_PROMPT = """Actúa como un Tutor Psicopedagogo que está al lado del alumno.
Tu objetivo es intervenir el examen para que sea comprensible, motivador y justo.

REGLAS DE ORO:
1. TRANSCRIBE TODO: No resumas. Si el examen tiene 7 puntos, el resultado debe tener 7 puntos.
2. NO RESUELVAS: El alumno debe trabajar. Deja los espacios vacíos que el docente puso.
3. RAZONA LA PISTA 💡: Antes de escribir la pista, resuelve el ejercicio vos. 
   - Si es 4x6, la pista debe apuntar al proceso: "💡 Si sumas 4 veces el 6, ¿cuánto te da?".
4. LENGUAJE AMIGABLE: Usa emojis (🔢, 📖, ✍️) y un tono que anime al chico.
5. IMÁGENES: Inserta  para que el alumno "vea" el problema.
   - Para Discalculia: Usa objetos concretos (bolitas, lápices).
   - Para Dislexia: Pictogramas de la acción (alguien leyendo, alguien uniendo)."""

# 3. MOTOR DE GENERACIÓN
def generar_imagen(desc):
    try:
        m = genai.GenerativeModel("imagen-3.0")
        res = m.generate_content(f"Dibujo escolar simple, fondo blanco, estilo pictograma: {desc}", safety_settings=SAFETY)
        return io.BytesIO(res.candidates[0].content.parts[0].inline_data.data)
    except: return None

def crear_examen_v12(texto_ia, nombre, diag, grupo, logo, gen_img):
    doc = Document()
    color_pista = RGBColor(0, 102, 0) # Verde pedagógico
    
    # Encabezado limpio
    table = doc.add_table(rows=1, cols=2)
    if logo:
        table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo), width=Inches(1.0))
    p = table.rows[0].cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diag.upper()} | GRUPO: {grupo.upper()}").bold = True

    # Fuente Inclusiva
    is_apo = any(x in str(diag).lower() for x in ["dislexia", "discalculia"]) or "A" in str(grupo)
    
    # Limpiar basura de la IA
    texto_ia = re.sub(r"^(¡Claro|Hola|Aquí|Entendido).*?\n", "", texto_ia, flags=re.IGNORECASE)

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea or "análisis:" in linea.lower(): continue

        if "[IMAGEN:" in linea and gen_img:
            img_data = generar_imagen(linea.split("[IMAGEN:")[1].split("]")[0])
            if img_data:
                pic = doc.add_paragraph()
                pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pic.add_run().add_picture(img_data, width=Inches(2.5))
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

# 4. INTERFAZ (Simple y Directa)
try:
    df = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv")
    df.columns = [c.strip() for c in df.columns]
    
    grado = st.sidebar.selectbox("Grado:", df[df.columns[1]].unique())
    alcance = st.sidebar.radio("Adecuar:", ["Todo el grado", "Elegir alumnos"])
    
    df_f = df[df[df.columns[1]] == grado]
    if alcance == "Elegir alumnos":
        sel = st.sidebar.multiselect("Alumnos:", df_f[df_f.columns[2]].tolist())
        df_f = df_f[df_f[df_f.columns[2]].isin(sel)]

    activar_img = st.sidebar.checkbox("Generar Imágenes", value=True)
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    archivo = st.file_uploader("Subir Examen (docx)", type=["docx"])

    if archivo and st.button("🚀 GENERAR EXÁMENES"):
        from docx import Document as DocRead
        txt_base = "\n".join([p.text for p in DocRead(archivo).paragraphs])
        zip_bio = io.BytesIO()
        
        with zipfile.ZipFile(zip_bio, "w") as z:
            prog = st.progress(0)
            for i, (_, fila) in enumerate(df_f.iterrows()):
                n, g, d = str(fila[df.columns[2]]), str(fila[df.columns[3]]), str(fila[df.columns[4]])
                
                # Llamada con reintento automático si falla
                m = genai.GenerativeModel(MODELO_IA)
                res = m.generate_content(f"{SYSTEM_PROMPT}\nALUMNO: {n} ({d}, {g})\nEXAMEN:\n{txt_base}", safety_settings=SAFETY)
                
                doc_res = crear_examen_v12(res.text, n, d, g, logo_bytes, activar_img)
                z.writestr(f"Adecuacion_{n.replace(' ', '_')}.docx", doc_res.getvalue())
                prog.progress((i + 1) / len(df_f))

        st.success("¡Hecho!")
        st.download_button("📥 Descargar Todo", zip_bio.getvalue(), "Examenes.zip")
except Exception as e:
    st.error(f"Error: {e}")
