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

# 1. CONFIGURACIÓN ESTRUCTURAL
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
st.set_page_config(page_title="Motor Pedagógico v8.5", layout="wide")

# Conexión y Escaneo de Modelos
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    def obtener_modelos():
        disponibles = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        prioridad = ["models/gemini-2.0-flash", "models/gemini-1.5-flash"]
        return [p for p in prioridad if p in disponibles] or disponibles
    MODELOS_OK = obtener_modelos()
except:
    MODELOS_OK = []

# PROMPT MAESTRO
SYSTEM_PROMPT = """Eres un Psicopedagogo experto en adecuación curricular.
REGLAS:
1. PISTAS: Usa 💡 en verde itálico. Solo si el Grupo (A o B) lo requiere.
2. IMÁGENES: Si el interruptor está activo, inserta .
   - Grupo A: Pictogramas simples, fondo blanco.
   - Grupo B: Esquemas organizadores.
   - Grupo C: Imágenes complejas para análisis/desafío.
3. RESALTE: **Negrita** solo en palabras clave de la respuesta. No resaltes conectores.
4. LIMPIEZA: Prohibido análisis internos o introducciones."""

# 2. FUNCIONES TÉCNICAS
def generar_imagen_ia(prompt_visual):
    try:
        # Nota: Requiere modelo con capacidad 'imagen-3.0' o similar habilitado en API Key
        model_img = genai.GenerativeModel("imagen-3.0")
        result = model_img.generate_content(prompt_visual)
        return io.BytesIO(result.candidates[0].content.parts[0].inline_data.data)
    except:
        return None

def crear_docx(texto_ia, nombre, diagnostico, grupo, logo_bytes=None, gen_img=False):
    doc = Document()
    diag, grupo_val = str(diagnostico).lower(), str(grupo).upper()
    color_inst, color_pista = RGBColor(31, 73, 125), RGBColor(0, 102, 0)

    # Encabezado
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try:
            table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    
    p = table.rows[0].cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_h = p.add_run(f"ESTUDIANTE: {nombre.upper()}\nAPOYO: {diagnostico.upper()} | GRUPO: {grupo_val}")
    run_h.bold, run_h.font.color.rgb = True, color_inst

    # Estilo base
    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo_val == "A"
    font_name = 'OpenDyslexic' if is_apo else 'Verdana'
    
    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea or any(x in linea.lower() for x in ["análisis:", "ayuda:"]): continue

        # Gestión de Imágenes
        if "[IMAGEN:" in linea and gen_img:
            desc = linea.split("[IMAGEN:")[1].split("]")[0]
            img_data = generar_imagen_ia(desc)
            if img_data:
                para_img = doc.add_paragraph()
                para_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para_img.add_run().add_picture(img_data, width=Inches(2.5))
                continue

        para = doc.add_paragraph()
        para.style.font.name = font_name
        para.paragraph_format.line_spacing = 1.5 if is_apo else 1.15

        if "💡" in linea:
            run_p = para.add_run(linea)
            run_p.font.color.rgb, run_p.italic, run_p.font.size = color_pista, True, Pt(11)
        elif "[CUADRICULA]" in linea:
            for _ in range(2): doc.add_paragraph().add_run(" " + "." * 75).font.color.rgb = RGBColor(200, 200, 200)
        else:
            partes = linea.split("**")
            for i, parte in enumerate(partes):
                run = para.add_run(parte)
                if i % 2 != 0: run.bold = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio

# 3. INTERFAZ STREAMLIT
st.title("Motor Pedagógico v8.5 💡")

if MODELOS_OK:
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
        df = pd.read_csv(url)
        df.columns = [c.strip() for c in df.columns]
        
        # Selectores en Barra Lateral
        col_grado, col_nombre, col_grupo, col_emergente = df.columns[1], df.columns[2], df.columns[3], df.columns[4]
        grado_sel = st.sidebar.selectbox("Seleccionar Grado:", df[col_grado].unique())
        
        # Filtro de Alumnos
        df_grado = df[df[col_grado] == grado_sel]
        modo_seleccion = st.sidebar.radio("Alcance:", ["Todo el grado", "Alumnos seleccionados"])
        
        if modo_seleccion == "Alumnos seleccionados":
            alumnos_sel = st.sidebar.multiselect("Elegir alumnos:", df_grado[col_nombre].tolist())
            alumnos_final = df_grado[df_grado[col_nombre].isin(alumnos_sel)]
        else:
            alumnos_final = df_grado

        # Configuración de IA
        activar_img = st.sidebar.checkbox("Generar Apoyos Visuales (IA)", value=False)
        logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
        logo_bytes = logo_file.read() if logo_file else None
        
        archivo_base = st.file_uploader("Subir Examen Original (docx)", type=["docx"])

        if archivo_base and not alumnos_final.empty and st.button("🚀 Iniciar Procesamiento"):
            from docx import Document as DocRead
            doc_read = DocRead(archivo_base)
            texto_base = "\n".join([p.text for p in doc_read.paragraphs])
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_f:
                progreso = st.progress(0)
                status = st.empty()
                
                for i, (_, fila) in enumerate(alumnos_final.iterrows()):
                    nombre, diag, grupo = str(fila[col_nombre]), str(fila[col_emergente]), str(fila[col_grupo])
                    status.text(f"Procesando: {nombre}...")
                    
                    # Lógica de Reintento/Cascada
                    for m_name in MODELOS_OK:
                        try:
                            time.sleep(2)
                            m_gen = genai.GenerativeModel(m_name)
                            prompt = f"{SYSTEM_PROMPT}\nIMG_ACTIVA: {activar_img}\nPERFIL: {nombre} ({diag}, Grupo {grupo})\nEXAMEN:\n{texto_base}"
                            res = m_gen.generate_content(prompt)
                            
                            doc_res = crear_docx(res.text, nombre, diag, grupo, logo_bytes, activar_img)
                            zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_res.getvalue())
                            break 
                        except: continue
                    
                    progreso.progress((i + 1) / len(alumnos_final))
                
            st.success("Procesamiento completado.")
            st.download_button("📥 Descargar Lote (.zip)", zip_buffer.getvalue(), f"Adecuaciones_{grado_sel}.zip")

    except Exception as e:
        st.error(f"Error de datos: {e}")
else:
    st.warning("Configura tu GOOGLE_API_KEY en los secrets.")
