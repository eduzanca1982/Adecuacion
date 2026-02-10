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
# 1. CONFIGURACIÓN GLOBAL (Fija e Inamovible)
# ==========================================
st.set_page_config(page_title="Motor Pedagógico v12.1 🍎", layout="wide")

# ID de la planilla y conexión
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    SAFETY = {category: HarmBlockThreshold.BLOCK_NONE for category in HarmCategory}
    # Detección de modelos disponibles
    modelos_api = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    MODELO_IA = next((m for m in modelos_api if "1.5-flash" in m), modelos_api[0])
except Exception as e:
    st.error(f"Error de conexión con la API de Google: {e}")

# ==========================================
# 2. EL PROMPT DE TUTOR PEDAGÓGICO
# ==========================================
SYSTEM_PROMPT = """Actúa como un Tutor Psicopedagogo experto.
Tu objetivo es intervenir el examen para que sea comprensible, motivador y justo.

REGLAS DE ORO:
1. TRANSCRIBE TODO: No resumas. El examen debe mantener todos sus puntos originales.
2. NO RESUELVAS: Deja los espacios de respuesta vacíos.
3. RAZONA LA PISTA 💡: Antes de escribir la pista, resuelve el ejercicio tú mismo. 
   - La pista debe guiar el proceso, no dar el resultado.
4. LENGUAJE AMIGABLE: Usa emojis (🔢, 📖, ✍️) para guiar al alumno.
5. IMÁGENES: Inserta  donde el alumno necesite "ver" para entender.
   - Para Discalculia: Usa apoyo visual concreto (lápices, cajas, billetes).
6. RESALTE: Usa **negrita** para verbos y datos clave."""

# ==========================================
# 3. FUNCIONES DE GENERACIÓN
# ==========================================
def generar_imagen(desc):
    try:
        m = genai.GenerativeModel("imagen-3.0")
        res = m.generate_content(f"Pictograma escolar simple, fondo blanco: {desc}", safety_settings=SAFETY)
        return io.BytesIO(res.candidates[0].content.parts[0].inline_data.data)
    except:
        return None

def crear_docx_adecuado(texto_ia, nombre, diag, grupo, logo_bytes, gen_img):
    doc = Document()
    color_pista = RGBColor(0, 102, 0)
    
    # Encabezado
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try:
            table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    p_hdr = table.rows[0].cells[1].paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hdr.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diag.upper()} | GRUPO: {grupo.upper()}").bold = True

    # Estilo de fuente según dificultad
    es_dislexico = any(x in str(diag).lower() for x in ["dislexia", "discalculia"]) or "A" in str(grupo)
    
    # Limpieza de comentarios iniciales de la IA
    texto_ia = re.sub(r"^(¡Claro|Hola|Aquí|Entendido|Como).*?\n", "", texto_ia, flags=re.IGNORECASE)

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea or "análisis:" in linea.lower(): continue

        if "[IMAGEN:" in linea and gen_img:
            desc_img = linea.split("[IMAGEN:")[1].split("]")[0]
            img_bytes = generar_imagen(desc_img)
            if img_bytes:
                para_pic = doc.add_paragraph()
                para_pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para_pic.add_run().add_picture(img_bytes, width=Inches(2.5))
            continue

        para = doc.add_paragraph()
        if "💡" in linea:
            run_p = para.add_run(linea)
            run_p.font.color.rgb, run_p.italic = color_pista, True
        else:
            partes = linea.split("**")
            for i, parte in enumerate(partes):
                run = para.add_run(parte)
                if i % 2 != 0: run.bold = True
                run.font.name = 'OpenDyslexic' if es_dislexico else 'Verdana'
                run.font.size = Pt(12 if es_dislexico else 11)
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio

# ==========================================
# 4. INTERFAZ STREAMLIT
# ==========================================
try:
    df = pd.read_csv(URL_PLANILLA)
    df.columns = [c.strip() for c in df.columns]
    
    # Menú Lateral - Configuración
    st.sidebar.header("🎯 Selección de Alumnos")
    grados_disponibles = df[df.columns[1]].unique()
    grado_sel = st.sidebar.selectbox("Elegir Grado:", grados_disponibles)
    
    df_grado = df[df[df.columns[1]] == grado_sel]
    
    # Opciones de alcance solicitadas
    alcance = st.sidebar.radio("¿A quiénes adecuar?", ["Todos los alumnos", "Seleccionar cuáles"])
    
    if alcance == "Seleccionar cuáles":
        alumnos_nombres = df_grado[df_grado.columns[2]].tolist()
        seleccionados = st.sidebar.multiselect("Checkbox de alumnos:", alumnos_nombres)
        alumnos_final = df_grado[df_grado[df_grado.columns[2]].isin(seleccionados)]
    else:
        alumnos_final = df_grado

    st.sidebar.divider()
    activar_img = st.sidebar.checkbox("Generar Imágenes con IA", value=True)
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    
    # Área Principal
    st.title("Motor de Adecuación v12.1 🎓")
    archivo_docx = st.file_uploader("Subir Examen Original (docx)", type=["docx"])

    if archivo_docx and st.button("🚀 INICIAR PROCESAMIENTO"):
        if alumnos_final.empty:
            st.warning("No hay alumnos seleccionados para procesar.")
        else:
            from docx import Document as DocRead
            doc_original = DocRead(archivo_docx)
            texto_base = "\n".join([p.text for p in doc_original.paragraphs])
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as z:
                prog = st.progress(0)
                status_text = st.empty()
                
                for i, (_, fila) in enumerate(alumnos_final.iterrows()):
                    nombre = str(fila[df.columns[2]])
                    grupo = str(fila[df.columns[3]])
                    diagnostico = str(fila[df.columns[4]])
                    
                    status_text.text(f"Razonando adecuación para: {nombre}...")
                    
                    # Llamada a la IA
                    modelo = genai.GenerativeModel(MODELO_IA)
                    respuesta = modelo.generate_content(
                        f"{SYSTEM_PROMPT}\nALUMNO: {nombre} ({diagnostico}, Grupo {grupo})\nEXAMEN:\n{texto_base}",
                        safety_settings=SAFETY
                    )
                    
                    doc_final = crear_docx_adecuado(respuesta.text, nombre, diagnostico, grupo, logo_bytes, activar_img)
                    z.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_final.getvalue())
                    prog.progress((i + 1) / len(alumnos_final))

            st.success("¡Lote completado!")
            st.download_button("📥 Descargar Resultados (ZIP)", zip_buffer.getvalue(), f"Adecuaciones_{grado_sel}.zip")

except Exception as e:
    st.error(f"Fallo general del sistema: {e}")
