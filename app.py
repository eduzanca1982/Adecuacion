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
MODELOS_CASCADA = ["models/gemini-2.0-flash", "models/gemini-2.0-flash-lite", "models/gemini-1.5-flash"]

st.set_page_config(page_title="Motor Pedagógico v8.7", layout="wide")

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    def obtener_modelos_disponibles():
        disponibles = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        return [m for m in MODELOS_CASCADA if m in disponibles] or disponibles
    MODELOS_OK = obtener_modelos_disponibles()
except Exception as e:
    st.error(f"Error de API: {e}")
    MODELOS_OK = []

# PROMPT MAESTRO: BLINDAJE PEDAGÓGICO
SYSTEM_PROMPT = """Eres un Maquetador Pedagógico. Tu tarea es ADECUAR el examen, NO resolverlo.

REGLAS CRÍTICAS:
1. EXAMEN EN BLANCO: No respondas preguntas ni resuelvas cálculos.
2. SIN CONSEJOS: Prohibido incluir textos para la docente o análisis psicopedagógicos.
3. IMÁGENES: Si se solicita, inserta .
   - Grupo A: Pictogramas lineales simples.
   - Grupo B: Esquemas de flujo o tablas.
   - Grupo C: Imágenes de desafío o curiosidades.
4. PISTAS: 💡 en verde itálico debajo de las consignas.
5. RESALTE: **Negrita** solo para evidencia nuclear. No resaltes conectores."""

# 2. FUNCIONES DE GENERACIÓN Y DISEÑO
def generar_imagen_ia(descripcion):
    try:
        model = genai.GenerativeModel("imagen-3.0") 
        res = model.generate_content(descripcion)
        return io.BytesIO(res.candidates[0].content.parts[0].inline_data.data)
    except:
        return None

def crear_docx_final(texto_ia, nombre, diagnostico, grupo, logo_bytes=None, gen_img=False):
    doc = Document()
    diag, grupo_v = str(diagnostico).lower(), str(grupo).upper()
    color_inst, color_pista = RGBColor(31, 73, 125), RGBColor(0, 102, 0)

    # Encabezado
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try:
            table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    
    p = table.rows[0].cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_h = p.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diagnostico.upper()} | GRUPO: {grupo_v}")
    run_h.bold, run_h.font.color.rgb = True, color_inst

    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo_v == "A"
    
    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        # Filtro de seguridad: no respuestas, no consejos a docente
        if any(x in linea.lower() for x in ["análisis:", "ayuda:", "respuesta:", "resultado:", "docente:"]): continue
        if not linea: continue

        # Inserción de Imagen
        if "[IMAGEN:" in linea and gen_img:
            desc = linea.split("[IMAGEN:")[1].split("]")[0]
            img_data = generar_imagen_ia(desc)
            if img_data:
                para_i = doc.add_paragraph()
                para_i.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para_i.add_run().add_picture(img_data, width=Inches(2.5))
            continue

        para = doc.add_paragraph()
        if "💡" in linea:
            run_p = para.add_run(f"💡 PISTA: {linea.replace('💡', '').strip()}")
            run_p.font.color.rgb, run_p.italic = color_pista, True
            continue

        if "[CUADRICULA]" in linea or "___" in linea:
            for _ in range(3): # Espacio justo de 3 líneas
                doc.add_paragraph().add_run(" " + "." * 75).font.color.rgb = RGBColor(210, 210, 210)
            continue

        es_titulo = (len(linea) < 55 and not linea.endswith('.')) or "[TITULO]" in linea
        partes = linea.replace("[TITULO]", "").strip().split("**")
        for i, parte in enumerate(partes):
            run = para.add_run(parte)
            if i % 2 != 0: run.bold = True
            run.font.name = 'OpenDyslexic' if is_apo else 'Verdana'
            run.font.size = Pt(12 if is_apo else 11)
            if es_titulo:
                run.bold, run.font.size, run.font.color.rgb = True, Pt(13), color_inst

    bio = io.BytesIO()
    doc.save(bio)
    return bio

# 3. INTERFAZ STREAMLIT COMPLETA
st.title("Motor Pedagógico v8.7 🧠")

if MODELOS_OK:
    try:
        url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
        df = pd.read_csv(url)
        df.columns = [c.strip() for c in df.columns]
        
        col_grado, col_nombre, col_grupo, col_emergente = df.columns[1], df.columns[2], df.columns[3], df.columns[4]
        
        # BARRA LATERAL
        st.sidebar.header("Configuración")
        grado_sel = st.sidebar.selectbox("Grado:", df[col_grado].unique())
        df_grado = df[df[col_grado] == grado_sel]
        
        modo_alcance = st.sidebar.radio("Alcance:", ["Todo el grado", "Selección manual"])
        
        alumnos_final = df_grado
        if modo_alcance == "Selección manual":
            seleccionados = st.sidebar.multiselect("Alumnos:", df_grado[col_nombre].tolist())
            alumnos_final = df_grado[df_grado[col_nombre].isin(seleccionados)]

        activar_img = st.sidebar.checkbox("Generar Imágenes con IA", value=False)
        logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
        logo_bytes = logo_file.read() if logo_file else None
        
        archivo_base = st.file_uploader("Subir Examen Base (docx)", type=["docx"])

        if archivo_base and not alumnos_final.empty and st.button("🚀 Iniciar Adecuación"):
            from docx import Document as DocRead
            doc_read = DocRead(archivo_base)
            texto_base = "\n".join([p.text for p in doc_read.paragraphs])
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_f:
                progreso = st.progress(0)
                status = st.empty()
                
                for i, (_, fila) in enumerate(alumnos_final.iterrows()):
                    nombre, diag, grupo = str(fila[col_nombre]), str(fila[col_emergente]), str(fila[col_grupo])
                    status.text(f"Generando: {nombre}...")
                    
                    success = False
                    for m_name in MODELOS_OK:
                        if success: break
                        try:
                            time.sleep(4) # Enfriamiento cuota
                            m_gen = genai.GenerativeModel(m_name)
                            res = m_gen.generate_content(f"{SYSTEM_PROMPT}\n\nALUMNO: {nombre} ({diag}, Grupo {grupo})\n\nEXAMEN:\n{texto_base}")
                            
                            doc_res = crear_docx_final(res.text, nombre, diag, grupo, logo_bytes, activar_img)
                            zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_res.getvalue())
                            success = True
                        except: continue
                    progreso.progress((i + 1) / len(alumnos_final))

            st.success("¡Lote completo finalizado!")
            st.download_button("📥 Descargar ZIP", zip_buffer.getvalue(), f"Examenes_{grado_sel}.zip")

    except Exception as e:
        st.error(f"Error de datos: {e}")
