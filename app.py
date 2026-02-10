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
# 1. Configuración de API y Prompt Maestro
# ----------------------------
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-2.0-flash')

SYSTEM_PROMPT = """Actúa como un diseñador editorial pedagógico. Tu misión es adecuar exámenes de primaria.
REGLAS ESTÉTICAS:
1. No escribas introducciones. Empieza directo en el examen.
2. MANTÉN la numeración original.
3. IMÁGENES: Escribe [MANTENER IMAGEN AQUÍ] donde el texto original las mencione.
4. Si el ejercicio es de completar, usa líneas largas de puntos.

ADECUACIONES SEGÚN PERFIL:
- DISLEXIA/DISCALCULIA: Frases cortas. Negrita SOLO en verbos de consigna. 
- TDAH: Una consigna por párrafo. Pasos numerados.
- GRUPO A: Simplifica sintaxis sin perder el objetivo pedagógico."""

# ----------------------------
# 2. Funciones de Maquetación y Formato
# ----------------------------
def extraer_texto(archivo):
    ext = archivo.name.split(".")[-1].lower()
    if ext == "docx":
        doc = Document(archivo)
        return "\n".join([p.text for p in doc.paragraphs])
    elif ext == "pdf":
        reader = PyPDF2.PdfReader(archivo)
        return "\n".join([p.extract_text() for p in reader.pages])
    return ""

def crear_docx_premium(texto_ia, nombre, diagnostico, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    color_institucional = RGBColor(31, 73, 125)

    # Encabezado con Tabla
    section = doc.sections[0]
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Inches(1.5)
    
    if logo_bytes:
        run_logo = table.rows[0].cells[0].paragraphs[0].add_run()
        run_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
    
    datos_cell = table.rows[0].cells[1]
    p_datos = datos_cell.paragraphs[0]
    p_datos.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_datos = p_datos.add_run(f"ESTUDIANTE: {nombre.upper()}\nEVALUACIÓN ADAPTADA")
    run_datos.bold = True
    run_datos.font.color.rgb = color_institucional

    # --- LÓGICA DE FUENTE OPENDYSLEXIC ---
    style = doc.styles['Normal']
    font = style.font
    
    # Nota: Para que funcione, OpenDyslexic debe estar instalada en la PC que abre el archivo.
    # En el código se define el nombre para que Word lo reconozca al abrirse.
    if "dislexia" in diag or "discalculia" in diag:
        font.name = 'OpenDyslexic'
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

        if "discalculia" in diag and any(c.isdigit() for c in linea):
            doc.add_paragraph("\n" + "." * 70).runs[0].font.color.rgb = RGBColor(210, 210, 210)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ----------------------------
# 3. Interfaz con Manejo de Errores 429
# ----------------------------
st.title("Motor Pedagógico v4.1 🍎")

try:
    df = pd.read_csv(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv")
    df.columns = [c.strip() for c in df.columns]
    
    col_grado, col_nombre, col_emergente = df.columns[1], df.columns[2], df.columns[4]
    grado = st.sidebar.selectbox("Seleccione Grado:", df[col_grado].unique())
    alumnos = df[(df[col_grado] == grado) & (df[col_emergente].str.lower() != "ninguna")]
    
    logo_file = st.sidebar.file_uploader("Logo Escuela", type=["png", "jpg"])
    logo_bytes = logo_file.read() if logo_file else None

    archivo_base = st.file_uploader("Subir Examen Base", type=["pdf", "docx"])

    if archivo_base and st.button(f"Generar carpeta para {grado}"):
        texto_base = extraer_texto(archivo_base)
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_f:
            progreso = st.progress(0)
            status_text = st.empty()
            
            for i, (_, fila) in enumerate(alumnos.iterrows()):
                nombre, diag = fila[col_nombre], fila[col_emergente]
                status_text.text(f"Procesando: {nombre}...")
                
                prompt = f"{SYSTEM_PROMPT}\n\nPERFIL: {nombre} ({diag})\n\nEXAMEN:\n{texto_base}"
                
                # --- SOLUCIÓN ERROR 429: Reintentos ---
                exito = False
                intentos = 0
                while not exito and intentos < 3:
                    try:
                        time.sleep(2 + intentos * 2) # Espera incremental
                        response = model.generate_content(prompt)
                        doc_bytes = crear_docx_premium(response.text, nombre, diag, logo_bytes)
                        zip_f.writestr(f"Adecuacion_{nombre.replace(' ', '_')}.docx", doc_bytes.getvalue())
                        exito = True
                    except Exception as e:
                        if "429" in str(e):
                            intentos += 1
                            status_text.warning(f"Límite alcanzado para {nombre}. Reintentando en {intentos * 5}s...")
                            time.sleep(5 * intentos)
                        else:
                            st.error(f"Error con {nombre}: {e}")
                            break
                
                progreso.progress((i + 1) / len(alumnos))
        
        st.success(f"Finalizado. Se generaron {len(alumnos)} exámenes.")
        st.download_button("Descargar ZIP", zip_buffer.getvalue(), f"Examenes_{grado}.zip")

except Exception as e:
    st.error(f"Error técnico: {e}")
