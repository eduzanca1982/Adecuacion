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
# Nombres corregidos para evitar el error 404
MODELOS_CASCADA = ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash", "gemini-1.5-pro"]

st.set_page_config(page_title="Motor Pedagógico v7.3", layout="wide")

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
except Exception as e:
    st.error("Error: Configura la API KEY en los Secrets.")

SYSTEM_PROMPT = """Eres un Psicopedagogo experto. Genera la adecuación FINAL del examen.

REGLAS DE ANDAMIAJE:
1. PISTAS: Para alumnos de Grupo A o Dificultad General, genera pistas breves [PISTA] (máx 2 renglones) que ayuden al razonamiento.
2. RESALTE: Marca en **negrita** la información nuclear del texto que responde a las preguntas. NO resaltes conectores.

REGLAS DE LIMPIEZA:
1. PROHIBIDO: No incluyas intros ni análisis técnicos.
2. ESPACIOS: Usa [CUADRICULA] solo donde el alumno deba escribir (máx 2 líneas)."""

# 2. FUNCIONES TÉCNICAS
def limpiar_output(texto):
    lineas = texto.split('\n')
    limpias = [l for l in lineas if not any(x in l.lower() for x in ["análisis:", "ayuda:", "analisis:"]) and not re.match(r'^[\s\.*]*$', l)]
    return "\n".join(limpias)

def crear_docx(texto_ia, nombre, diagnostico, grupo, logo_bytes=None):
    doc = Document()
    diag, grupo = str(diagnostico).lower(), str(grupo).upper()
    color_inst, color_pista = RGBColor(31, 73, 125), RGBColor(0, 102, 0)

    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try:
            run_logo = table.rows[0].cells[0].paragraphs[0].add_run()
            run_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    
    cell_info = table.rows[0].cells[1]
    p = cell_info.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diagnostico.upper()} | GRUPO: {grupo}")
    run.bold, run.font.color.rgb = True, color_inst

    style = doc.styles['Normal']
    font = style.font
    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo == "A"
    font.name = 'OpenDyslexic' if is_apo else 'Verdana'
    font.size = Pt(12 if is_apo else 11)
    style.paragraph_format.line_spacing = 1.5 if is_apo else 1.15

    texto_limpio = limpiar_output(texto_ia)
    for linea in texto_limpio.split('\n'):
        linea = linea.strip()
        if not linea: continue
        para = doc.add_paragraph()
        
        if "[PISTA]" in linea or "💡" in linea:
            run_p = para.add_run(f"💡 PISTA: {linea.replace('[PISTA]', '').replace('💡', '').strip()}")
            run_p.font.color.rgb, run_p.italic = color_pista, True
            continue

        if "[CUADRICULA]" in linea:
            for _ in range(2):
                doc.add_paragraph().add_run(" " + "." * 70).font.color.rgb = RGBColor(215, 215, 215)
            continue

        es_titulo = (len(linea) < 55 and not linea.endswith('.')) or "[TITULO]" in linea
        linea_final = linea.replace("[TITULO]", "").strip()
        partes = linea_final.split("**")
        for i, parte in enumerate(partes):
            run_part = para.add_run(parte)
            if i % 2 != 0: run_part.bold = True
            if es_titulo:
                run_part.bold, run_part.font.size, run_part.font.color.rgb = True, Pt(13), color_inst

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 3. INTERFAZ
st.title("Motor Pedagógico v7.3 🎓")

try:
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    col_grado, col_nombre, col_grupo, col_emergente = df.columns[1], df.columns[2], df.columns[3], df.columns[4]
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
        procesados_ok, debug_logs = 0, []

        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            archivo_base.seek(0)
            zip_f.writestr(f"ORIGINAL_{archivo_base.name}", archivo_base.read())
            
            progreso = st.progress(0)
            status = st.empty()

            for i, (_, fila) in enumerate(alumnos_grado.iterrows()):
                nombre, diag, grupo = str(fila[col_nombre]), str(fila[col_emergente]), str(fila[col_grupo])
                status.text(f"Generando: {nombre}...")
                
                success = False
                errores_alumno = []

                for m_name in MODELOS_CASCADA:
                    if success: break
                    # Intentar hasta 2 veces por modelo si es error de cuota
                    for intento in range(2):
                        try:
                            time.sleep(2) 
                            m_gen = genai.GenerativeModel(m_name)
                            res = m_gen.generate_content(f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} ({diag}, Grupo {grupo})\n\nEXAMEN:\n{texto_base}")
                            
                            doc_final = crear_docx(res.text, nombre, diag, grupo, logo_bytes)
                            zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_final.getvalue())
                            success, procesados_ok = True, procesados_ok + 1
                            break
                        except Exception as e:
                            err_str = str(e)
                            if "429" in err_str:
                                status.warning(f"Saturación. Esperando 30s para {nombre}...")
                                time.sleep(30)
                            else:
                                errores_alumno.append(f"{m_name}: {err_str[:150]}")
                                break 
                
                if not success: debug_logs.append({"alumno": nombre, "errores": errores_alumno})
                progreso.progress((i + 1) / len(alumnos_grado))

        if procesados_ok > 0:
            st.success(f"Éxito: {procesados_ok} generados.")
            st.download_button("Descargar ZIP", zip_buffer.getvalue(), f"Examenes_{grado_sel}.zip")
        
        if debug_logs:
            with st.expander("🔍 Ventana de Diagnóstico"):
                for log in debug_logs:
                    st.warning(f"Alumno: {log['alumno']}")
                    for err in log['errores']: st.write(f"- {err}")
                    st.divider()

except Exception as e:
    st.error(f"Error general: {e}")
