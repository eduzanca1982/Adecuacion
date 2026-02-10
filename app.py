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

# 1. CONFIGURACIÓN Y MODELOS DISPONIBLES
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"

# Lista de modelos que tu dashboard mostró como activos
OPCIONES_MODELOS = [
    "gemini-2.0-flash", 
    "gemini-2.0-flash-lite", 
    "gemini-1.5-flash", 
    "gemini-1.5-pro"
]

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
except Exception as e:
    st.error(f"Falta GOOGLE_API_KEY en los Secrets: {e}")

SYSTEM_PROMPT = """Eres un experto en adecuación curricular. Adapta el contenido manteniendo la jerarquía de ejercicios.
- Usa [TITULO] para encabezados principales.
- Usa iconos: 📖 (Lectura), 🔢 (Cálculo), ✍️ (Escritura).
- Perfil Dificultad General: Simplifica oraciones y usa negrita en verbos de acción."""

# 2. FUNCIONES DE DISEÑO
def limpiar_nombre(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", str(nombre)).replace(" ", "_")

def crear_docx_adecuado(texto_ia, nombre, diagnostico, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    color_inst = RGBColor(31, 73, 125)

    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try:
            run_logo = table.rows[0].cells[0].paragraphs[0].add_run()
            run_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(1.1))
        except: pass
    
    cell_info = table.rows[0].cells[1]
    p = cell_info.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"ALUMNO: {nombre.upper()}\nADECUACIÓN: {diagnostico.upper()}")
    run.bold = True
    run.font.color.rgb = color_inst

    style = doc.styles['Normal']
    font = style.font
    font.name = 'OpenDyslexic' if any(x in diag for x in ["dislexia", "discalculia", "general"]) else 'Verdana'
    font.size = Pt(12) if font.name == 'OpenDyslexic' else Pt(11)
    style.paragraph_format.line_spacing = 1.5 if font.name == 'OpenDyslexic' else 1.15

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea: continue
        para = doc.add_paragraph()
        es_titulo = "[TITULO]" in linea or (len(linea) < 50 and not linea.endswith('.'))
        
        texto_limpio = linea.replace("[TITULO]", "").strip()
        partes = texto_limpio.split("**")
        for i, parte in enumerate(partes):
            run_p = para.add_run(parte)
            if i % 2 != 0: run_p.bold = True
            if es_titulo:
                run_p.bold = True
                run_p.font.size = Pt(14)
                run_p.font.color.rgb = color_inst

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 3. INTERFAZ
st.title("Motor Pedagógico v5.6 🚀")

try:
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    
    col_grado, col_nombre, col_emergente = df.columns[1], df.columns[2], df.columns[4]
    
    # --- SELECTOR DE MODELOS ---
    st.sidebar.header("Configuración de IA")
    modelo_principal = st.sidebar.selectbox(
        "Modelo de inicio:", 
        OPCIONES_MODELOS, 
        index=0,
        help="Si este modelo falla, la app probará automáticamente con los demás."
    )
    
    grado_sel = st.sidebar.selectbox("Grado:", df[col_grado].unique())
    alumnos_grado = df[(df[col_grado] == grado_sel) & (df[col_emergente].str.lower() != "ninguna")]
    
    logo_file = st.sidebar.file_uploader("Logo Colegio", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None
    archivo_base = st.file_uploader("Subir Examen Base", type=["docx", "pdf"])

    if archivo_base and st.button(f"Procesar Grado ({len(alumnos_grado)} alumnos)"):
        from docx import Document as DocRead
        doc_read = DocRead(archivo_base)
        texto_base = "\n".join([p.text for p in doc_read.paragraphs])
        
        zip_buffer = io.BytesIO()
        procesados_ok = 0
        debug_logs = []

        # Crear lista de cascada: primero el elegido, luego el resto
        cascada_modelos = [modelo_principal] + [m for m in OPCIONES_MODELOS if m != modelo_principal]

        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            archivo_base.seek(0)
            zip_f.writestr(f"ORIGINAL_{archivo_base.name}", archivo_base.read())
            
            progreso = st.progress(0)
            status = st.empty()

            for i, (_, fila) in enumerate(alumnos_grado.iterrows()):
                nombre = str(fila[col_nombre])
                diag = str(fila[col_emergente])
                status.text(f"Adecuando: {nombre}...")
                
                success = False
                intentos_log = []

                # ITERACIÓN EN CASCADA
                for m_name in cascada_modelos:
                    if success: break
                    
                    try:
                        m_gen = genai.GenerativeModel(m_name)
                        time.sleep(2) # Pausa mínima preventiva
                        p_prompt = f"{SYSTEM_PROMPT}\n\nALUMNO: {nombre} ({diag})\n\nCONTENIDO:\n{texto_base}"
                        res = m_gen.generate_content(p_prompt)
                        
                        doc_bytes = crear_docx_adecuado(res.text, nombre, diag, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{limpiar_nombre(nombre)}.docx", doc_bytes.getvalue())
                        success = True
                        procesados_ok += 1
                        intentos_log.append(f"Éxito con {m_name}")
                    except Exception as e:
                        err_msg = str(e)
                        intentos_log.append(f"Fallo en {m_name}: {err_msg[:100]}")
                        if "429" in err_msg:
                            time.sleep(5) # Si es cuota, esperar un poco antes de saltar al siguiente
                        continue
                
                if not success:
                    debug_logs.append({"alumno": nombre, "historial": intentos_log})
                
                progreso.progress((i + 1) / len(alumnos_grado))

        if procesados_ok > 0:
            st.success(f"Proceso finalizado. {procesados_ok} archivos generados.")
            st.download_button("Descargar ZIP Completo", zip_buffer.getvalue(), f"Examenes_{grado_sel}.zip")
        
        if debug_logs:
            with st.expander("🔍 Detalle de la Cascada de Modelos"):
                for log in debug_logs:
                    st.write(f"**Alumno:** {log['alumno']}")
                    for msg in log['historial']:
                        st.write(f"- {msg}")
                    st.divider()

except Exception as e:
    st.error(f"Fallo general: {e}")
