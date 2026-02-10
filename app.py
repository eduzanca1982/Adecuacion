import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import PyPDF2
import io
import zipfile
import time

# ----------------------------
# 1. Configuración de API y ID de Planilla
# ----------------------------
# Reemplaza con tu ID de planilla real
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    model = genai.GenerativeModel('gemini-2.0-flash')
except Exception as e:
    st.error("Error al configurar la API Key en Streamlit Secrets.")

SYSTEM_PROMPT = """Actúa como un diseñador editorial pedagógico. Tu misión es adecuar exámenes de primaria.
REGLAS ESTÉTICAS:
1. No escribas introducciones. Empieza directo en el examen.
2. MANTÉN la numeración original y la jerarquía de títulos.
3. IMÁGENES: Escribe [MANTENER IMAGEN AQUÍ] donde el texto original las mencione.
4. Si el ejercicio es de completar, usa líneas de puntos: ........................

ADECUACIONES SEGÚN PERFIL:
- DISLEXIA/DISCALCULIA: Frases cortas. Negrita SOLO en verbos de consigna. 
- TDAH: Una consigna por párrafo. Pasos numerados.
- GRUPO A: Simplifica sintaxis sin perder el objetivo pedagógico."""

# ----------------------------
# 2. Funciones de Maquetación y Formato
# ----------------------------
def extraer_texto(archivo):
    ext = archivo.name.split('.')[-1].lower()
    if ext == 'docx':
        doc = Document(archivo)
        return "\n".join([p.text for p in doc.paragraphs])
    elif ext == 'pdf':
        reader = PyPDF2.PdfReader(archivo)
        return "\n".join([p.extract_text() for p in reader.pages])
    return ""

def crear_docx_premium(texto_ia, nombre, diagnostico, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    color_institucional = RGBColor(31, 73, 125)

    # Encabezado con Tabla para Estética
    section = doc.sections[0]
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.5)
    
    if logo_bytes:
        try:
            run_logo = table.rows[0].cells[0].paragraphs[0].add_run()
            run_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except:
            pass # Si el logo falla, continúa sin él
    
    datos_cell = table.rows[0].cells[1]
    p_datos = datos_cell.paragraphs[0]
    p_datos.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_datos = p_datos.add_run(f"ESTUDIANTE: {nombre.upper()}\nEVALUACIÓN ADAPTADA")
    run_datos.bold = True
    run_datos.font.color.rgb = color_institucional

    # --- CONFIGURACIÓN DE FUENTE OPENDYSLEXIC ---
    style = doc.styles['Normal']
    font = style.font
    if "dislexia" in diag or "discalculia" in diag:
        font.name = 'OpenDyslexic' # Debe estar instalada en la PC que imprime
        font.size = Pt(12)
        style.paragraph_format.line_spacing = 1.5
    else:
        font.name = 'Verdana'
        font.size = Pt(11)
        style.paragraph_format.line_spacing = 1.15

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea: continue
        
        para = doc.add_paragraph()
        es_titulo = len(linea) < 60 and not linea.endswith('.')
        
        partes = linea.split("**")
        for i, parte in enumerate(partes):
            run = para.add_run(parte)
            if i % 2 != 0: run.bold = True
            if es_titulo:
                run.bold = True
                run.font.size = Pt(14)
                run.font.color.rgb = color_institucional

        # Guía visual para Discalculia
        if "discalculia" in diag and any(c.isdigit() for c in linea):
            doc.add_paragraph("\n" + "." * 70).runs[0].font.color.rgb = RGBColor(210, 210, 210)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ----------------------------
# 3. Interfaz de Usuario
# ----------------------------
st.title("Motor Pedagógico v4.2 🍎")

try:
    # Carga de Planilla
    SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    df = pd.read_csv(SHEET_URL)
    df.columns = [c.strip() for c in df.columns]
    
    col_grado = df.columns[1]
    col_nombre = df.columns[2]
    col_emergente = df.columns[4]

    grado = st.sidebar.selectbox("Grado:", df[col_grado].unique())
    alumnos = df[(df[col_grado] == grado) & (df[col_emergente].str.lower() != "ninguna")]
    
    st.sidebar.info(f"Alumnos en {grado}: {len(alumnos)}")
    logo_file = st.sidebar.file_uploader("Subir Logo Escuela", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None

    archivo_base = st.file_uploader("Subir Examen Base (DOCX o PDF)", type=["docx", "pdf"])

    if archivo_base and st.button(f"Generar carpeta para {grado}"):
        texto_base = extraer_texto(archivo_base)
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            progreso = st.progress(0)
            status = st.empty()
            
            for i, (_, fila) in enumerate(alumnos.iterrows()):
                nombre, diag = fila[col_nombre], fila[col_emergente]
                status.text(f"Procesando: {nombre}...")
                
                prompt = f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} ({diag})\n\nEXAMEN:\n{texto_base}"
                
                # --- MANEJO DE ERROR 429 Y REINTENTOS ---
                exito = False
                intentos = 0
                while not exito and intentos < 3:
                    try:
                        time.sleep(1 + intentos * 2) # Pausa mínima
                        response = model.generate_content(prompt)
                        doc_bytes = crear_docx_premium(response.text, nombre, diag, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_bytes.getvalue())
                        exito = True
                    except Exception as e:
                        if "429" in str(e):
                            intentos += 1
                            status.warning(f"Límite excedido para {nombre}. Reintentando ({intentos}/3)...")
                            time.sleep(10 * intentos) # Espera mayor si hay saturación
                        else:
                            st.error(f"Fallo crítico con {nombre}: {e}")
                            break
                
                progreso.progress((i + 1) / len(alumnos))
        
        st.success(f"Finalizado. Procesados {len(alumnos)} alumnos.")
        st.download_button("Descargar ZIP de Adecuaciones", zip_buffer.getvalue(), f"Examenes_{grado}.zip")

except Exception as e:
    st.error(f"Error técnico: {e}")
