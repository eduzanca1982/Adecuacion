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
import traceback

# 1. CONFIGURACIÓN ESTRUCTURAL
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
MODELOS_CASCADA = ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash"]

st.set_page_config(page_title="Motor Pedagógico v7.1", layout="wide")

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
except Exception as e:
    st.error("Error: No se encontró la API KEY en los Secrets.")

# PROMPT MAESTRO (RESTAURACIÓN DE PISTAS + CONTROL)
SYSTEM_PROMPT = """Eres un Psicopedagogo y Diseñador Editorial. Tu misión es generar la adecuación FINAL.

REGLAS DE ANDAMIAJE (PISTAS):
1. PISTAS: Genera pistas breves [PISTA] solo cuando el alumno lo necesite por su perfil. 
   - La pista debe ayudar a ENTENDER qué hacer (ej: "Buscá el dato en el 2do párrafo").
   - NO des la respuesta. Usa un tono alentador.
2. RESALTE: Marca en **negrita** la información nuclear del texto que responde a las preguntas. NO resaltes conectores.

REGLAS DE LIMPIEZA:
1. PROHIBIDO: No incluyas intros, saludos ni análisis técnicos.
2. ESPACIOS: Usa [CUADRICULA] solo donde el alumno deba escribir. Máximo 2 líneas de puntos por ejercicio.
3. ESTÉTICA: No agregues líneas de puntos al azar entre párrafos."""

# 2. FUNCIONES DE DISEÑO Y LIMPIEZA
def limpiar_output_ia(texto):
    """Filtra líneas vacías de puntos y textos de análisis interno."""
    lineas = texto.split('\n')
    limpias = []
    for l in lineas:
        if re.match(r'^[\s\.*]*$', l): continue # Elimina líneas que solo tienen puntos
        if any(x in l.lower() for x in ["análisis:", "ayuda:", "aquí tienes", "analisis pedagógico"]): continue
        limpias.append(l)
    return "\n".join(limpias)

def crear_docx_premium(texto_ia, nombre, diagnostico, grupo, logo_bytes=None):
    doc = Document()
    diag, grupo = str(diagnostico).lower(), str(grupo).upper()
    color_inst = RGBColor(31, 73, 125)
    color_pista = RGBColor(0, 102, 0) # Verde oscuro para pistas

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

    # Tipografía Dinámica
    style = doc.styles['Normal']
    font = style.font
    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo == "A"
    font.name = 'OpenDyslexic' if is_apo else 'Verdana'
    font.size = Pt(12 if is_apo else 11)
    style.paragraph_format.line_spacing = 1.5 if is_apo else 1.15

    texto_limpio = limpiar_output_ia(texto_ia)
    for linea in texto_limpio.split('\n'):
        linea = linea.strip()
        if not linea: continue
        
        para = doc.add_paragraph()
        
        # Formato de Pistas
        if "[PISTA]" in linea or "💡" in linea:
            txt_pista = linea.replace("[PISTA]", "").replace("💡", "").strip()
            run_p = para.add_run(f"💡 PISTA: {txt_pista}")
            run_p.font.color.rgb = color_pista
            run_p.italic = True
            continue

        # Formato de Cuadrícula Controlada
        if "[CUADRICULA]" in linea:
            for _ in range(2):
                p_g = doc.add_paragraph()
                p_g.add_run(" " + "." * 75).font.color.rgb = RGBColor(215, 215, 215)
                p_g.paragraph_format.space_after = Pt(0)
            continue

        es_titulo = (len(linea) < 55 and not linea.endswith('.')) or "[TITULO]" in linea
        linea_final = linea.replace("[TITULO]", "").strip()
        
        partes = linea_final.split("**")
        for i, parte in enumerate(partes):
            run_part = para.add_run(parte)
            if i % 2 != 0: run_part.bold = True
            if es_titulo:
                run_part.bold = True
                run_part.font.size = Pt(13)
                run_part.font.color.rgb = color_inst

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 3. INTERFAZ Y CONSOLA DE DIAGNÓSTICO
st.title("Motor Pedagógico v7.1 🚀")

try:
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    col_grado, col_nombre, col_grupo, col_emergente = df.columns[1], df.columns[2], df.columns[3], df.columns[4]
    
    st.sidebar.header("Opciones de IA")
    modelo_ini = st.sidebar.selectbox("Prioridad de Modelo:", MODELOS_CASCADA)
    grado_sel = st.sidebar.selectbox("Grado:", df[col_grado].unique())
    alumnos_grado = df[(df[col_grado] == grado_sel) & (df[col_emergente].str.lower() != "ninguna")]
    
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    archivo_base = st.file_uploader("Subir Examen Base", type=["docx", "pdf"])

    if archivo_base and st.button(f"Procesar Grupo"):
        from docx import Document as DocRead
        doc_read = DocRead(archivo_base)
        texto_base = "\n".join([p.text for p in doc_read.paragraphs])
        
        zip_buffer = io.BytesIO()
        procesados_ok = 0
        debug_logs = []

        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            archivo_base.seek(0)
            zip_f.writestr(f"ORIGINAL_{archivo_base.name}", archivo_base.read())
            
            progreso = st.progress(0)
            status = st.empty()
            
            orden_modelos = [modelo_ini] + [m for m in MODELOS_CASCADA if m != modelo_ini]

            for i, (_, fila) in enumerate(alumnos_grado.iterrows()):
                nombre, diag, grupo = str(fila[col_nombre]), str(fila[col_emergente]), str(fila[col_grupo])
                status.text(f"Generando: {nombre}...")
                
                success = False
                errores_acumulados = []

                for m_name in orden_modelos:
                    if success: break
                    try:
                        time.sleep(3) 
                        m_gen = genai.GenerativeModel(m_name)
                        p_prompt = f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} ({diag}, Grupo {grupo})\n\nEXAMEN ORIGINAL:\n{texto_base}"
                        res = m_gen.generate_content(p_prompt)
                        
                        doc_final = crear_docx_premium(res.text, nombre, diag, grupo, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_final.getvalue())
                        success = True
                        procesados_ok += 1
                    except Exception as e:
                        errores_acumulados.append(f"{m_name}: {str(e)[:150]}")
                        continue
                
                if not success:
                    debug_logs.append({"alumno": nombre, "errores": errores_acumulados})
                
                progreso.progress((i + 1) / len(alumnos_grado))

        if procesados_ok > 0:
            st.success(f"Éxito: {procesados_ok} archivos generados.")
            st.download_button("Descargar ZIP Completo", zip_buffer.getvalue(), f"Examenes_{grado_sel}.zip")
        
        if debug_logs:
            with st.expander("🔍 Ventana de Diagnóstico"):
                for log in debug_logs:
                    st.warning(f"Alumno: {log['alumno']}")
                    for err in log['errores']:
                        st.write(f"- {err}")
                    st.divider()

except Exception as e:
    st.error(f"Error general: {e}")
