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

# 1. CONFIGURACIÓN ESTRUCTURAL Y ESCANEO
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
st.set_page_config(page_title="Motor Pedagógico v7.5", layout="wide")

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    
    # --- FUNCIÓN DE ESCANEO PREVENTIVO ---
    def obtener_modelos_disponibles():
        disponibles = []
        for m in genai.list_models():
            # Filtramos solo los modelos que soportan generación de texto (GenerateContent)
            if 'generateContent' in m.supported_generation_methods:
                disponibles.append(m.name)
        # Priorizamos 2.0 y 1.5 si están en la lista
        prioridad = ["models/gemini-2.0-flash", "models/gemini-2.0-flash-lite", "models/gemini-1.5-flash"]
        final = [p for p in prioridad if p in disponibles]
        # Agregamos el resto que no esté en prioridad
        final += [d for d in disponibles if d not in final]
        return final

    MODELOS_VALIDOS = obtener_modelos_disponibles()
except Exception as e:
    st.error(f"Error de conexión o API Key: {e}")
    MODELOS_VALIDOS = []

SYSTEM_PROMPT = """Eres un Psicopedagogo experto. Genera la adecuación FINAL del examen.
REGLAS:
1. PISTAS: Breves [PISTA] solo para Grupo A o Dificultad General.
2. RESALTE: En **negrita** solo información nuclear de la respuesta.
3. LIMPIEZA: Prohibido intros, análisis o líneas de puntos excesivas."""

# 2. FUNCIONES DE DISEÑO
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

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea or any(x in linea.lower() for x in ["análisis:", "ayuda:", "analisis:"]): continue
        para = doc.add_paragraph()
        if "[PISTA]" in linea or "💡" in linea:
            run_p = para.add_run(f"💡 PISTA: {linea.replace('[PISTA]', '').replace('💡', '').strip()}")
            run_p.font.color.rgb, run_p.italic = color_pista, True
            continue
        if "[CUADRICULA]" in linea:
            for _ in range(2): doc.add_paragraph().add_run(" " + "." * 70).font.color.rgb = RGBColor(215, 215, 215)
            continue
        es_titulo = (len(linea) < 55 and not linea.endswith('.')) or "[TITULO]" in linea
        partes = linea.replace("[TITULO]", "").strip().split("**")
        for i, parte in enumerate(partes):
            run_part = para.add_run(parte)
            if i % 2 != 0: run_part.bold = True
            if es_titulo: run_part.bold, run_part.font.size, run_part.font.color.rgb = True, Pt(13), color_inst

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 3. INTERFAZ
st.title("Motor Pedagógico v7.5 🎓")

if MODELOS_VALIDOS:
    st.sidebar.success(f"Modelos detectados: {len(MODELOS_VALIDOS)}")
    with st.sidebar.expander("Ver modelos disponibles"):
        for m in MODELOS_VALIDOS: st.write(f"- {m}")
else:
    st.error("No se detectaron modelos disponibles. Revisa tu facturación o API Key.")

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

                for m_name in MODELOS_VALIDOS:
                    if success: break
                    try:
                        time.sleep(5) # Enfriamiento RPM
                        m_gen = genai.GenerativeModel(m_name)
                        res = m_gen.generate_content(f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} ({diag}, Grupo {grupo})\n\nEXAMEN:\n{texto_base}")
                        
                        doc_final = crear_docx(res.text, nombre, diag, grupo, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_final.getvalue())
                        success, procesados_ok = True, procesados_ok + 1
                    except Exception as e:
                        err_str = str(e)
                        if "429" in err_str:
                            status.warning(f"Límite en {m_name}. Saltando...")
                            time.sleep(10)
                        else:
                            errores_alumno.append(f"{m_name}: {err_str[:100]}")
                
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

except Exception as e:
    st.error(f"Error general: {e}")
