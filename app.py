import streamlit as st
import google.generativeai as genai
import pandas as pd
from docx import Document
import PyPDF2
import io
import zipfile
import time

# 1. Configuración de API
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
model = genai.GenerativeModel('gemini-2.0-flash')

# 2. Funciones de Extracción de Texto
def extraer_texto(archivo):
    nombre = archivo.name.lower()
    if nombre.endswith('.docx'):
        doc = Document(archivo)
        return "\n".join([p.text for p in doc.paragraphs])
    elif nombre.endswith('.pdf'):
        pdf_reader = PyPDF2.PdfReader(archivo)
        texto = ""
        for pagina in pdf_reader.pages:
            texto += pagina.extract_text()
        return texto
    else:
        st.error("Formato no soportado. Usa .docx o .pdf")
        return None

# 3. Función de Creación de Word (Conserva estética)
def crear_docx_fiel(texto_ia, nombre_alumno, diagnostico):
    doc = Document()
    diag_str = str(diagnostico).lower()
    
    # Estilo base por diagnóstico
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial' if "dislexia" in diag_str else 'Verdana'
    font.size = Pt(12 if "dislexia" in diag_str else 11)

    # Encabezado
    p = doc.add_paragraph()
    p.add_run(f"ESTUDIANTE: {nombre_alumno}").bold = True
    doc.add_paragraph("-" * 20)

    for linea in texto_ia.split('\n'):
        if linea.strip():
            para = doc.add_paragraph()
            # Lógica de negritas y títulos
            if "**" in linea:
                partes = linea.split("**")
                for i, parte in enumerate(partes):
                    run = para.add_run(parte)
                    if i % 2 != 0: run.bold = True
            else:
                para.add_run(linea)
            
            # Espacios extra para Discalculia
            if "discalculia" in diag_str and any(c.isdigit() for c in linea):
                doc.add_paragraph("\n" * 2)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- Interfaz Principal ---
st.title("Motor Pedagógico Multiformato 📚")

# Carga de planilla (usando tu SHEET_ID ya configurado)
# ... (mantener lógica de cargar_datos y filtrado de grado de v3.2) ...

archivo_examen = st.file_uploader("Subir Examen (DOCX o PDF)", type=["docx", "pdf"])

if archivo_examen and st.button("Procesar Todo el Grado"):
    texto_base = extraer_texto(archivo_examen)
    
    if texto_base:
        # (Mantener lógica de bucle 'for' y creación de ZIP de v3.2)
        # El prompt se envía como definimos en el punto 1
        pass
