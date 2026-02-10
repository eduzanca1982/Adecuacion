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

# 1. CONFIGURACIÓN
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
st.set_page_config(page_title="Motor Pedagógico v9.2", layout="wide")

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    def obtener_modelos():
        disponibles = [m.name for m in genai.list_models()]
        return [m for m in ["models/gemini-2.0-flash", "models/gemini-1.5-flash"] if m in disponibles] or disponibles
    MODELOS_OK = obtener_modelos()
except:
    MODELOS_OK = []

# PROMPT MAESTRO (FIDELIDAD Y APOYOS)
SYSTEM_PROMPT = """Eres un Diseñador de Inclusión Escolar. 
1. TRANSCRIBE: Copia cada ejercicio del original. NO los resuelvas.
2. IMÁGENES: Si hay conceptos abstractos (reparto, lectura), inserta .
3. PISTAS: 💡 en verde itálico solo para Grupo A y B.
4. FORMATO: Usa [CUADRICULA] para espacios de respuesta."""

# 2. FUNCIONES DE GENERACIÓN
def intentar_generar_imagen(descripcion):
    """
    Intenta generar imagen con Gemini. 
    Retorna bytes si tiene éxito, o el error si falla.
    """
    try:
        # Intentamos con el modelo de imagen específico
        model_img = genai.GenerativeModel("imagen-3.0")
        response = model_img.generate_content(descripcion)
        # Verificamos si hay datos en la respuesta
        return io.BytesIO(response.candidates[0].content.parts[0].inline_data.data), "OK"
    except Exception as e:
        return None, str(e)

def crear_docx_v9_2(texto_ia, nombre, diagnostico, grupo, logo_bytes=None, gen_img=False):
    doc = Document()
    diag, grupo_v = str(diagnostico).lower(), str(grupo).upper()
    color_inst, color_pista = RGBColor(31, 73, 125), RGBColor(0, 102, 0)
    reporte_imagenes = []

    # Encabezado
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try: table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(1.0))
        except: pass
    p = table.rows[0].cells[1].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diagnostico.upper()} | GRUPO: {grupo_v}").bold = True

    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo_v == "A"
    
    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        if not linea or any(x in linea.lower() for x in ["análisis:", "ayuda:", "respuesta:"]): continue

        # LÓGICA DE IMAGEN CON REPORTE
        if "[IMAGEN:" in linea and gen_img:
            desc = linea.split("[IMAGEN:")[1].split("]")[0]
            img_data, status = intentar_generar_imagen(desc)
            if img_data:
                para_i = doc.add_paragraph()
                para_i.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para_i.add_run().add_picture(img_data, width=Inches(2.5))
                reporte_imagenes.append(f"✅ Imagen generada: {desc[:30]}...")
            else:
                reporte_imagenes.append(f"❌ Error en imagen ({desc[:20]}): {status}")
            continue

        para = doc.add_paragraph()
        if "💡" in linea:
            run = para.add_run(linea)
            run.font.color.rgb, run.italic = color_pista, True
            continue

        if "___" in linea or "[CUADRICULA]" in linea:
            for _ in range(3): doc.add_paragraph().add_run(" " + "." * 75).font.color.rgb = RGBColor(215, 215, 215)
            continue

        # Texto base
        partes = linea.replace("[TITULO]", "").split("**")
        for i, parte in enumerate(partes):
            run = para.add_run(parte)
            if i % 2 != 0: run.bold = True
            run.font.name = 'OpenDyslexic' if is_apo else 'Verdana'
            run.font.size = Pt(12 if is_apo else 11)

    bio = io.BytesIO()
    doc.save(bio)
    return bio, reporte_imagenes

# 3. INTERFAZ
# ... (Mantenemos selectores de v9.1)

        if archivo_base and st.button("Procesar"):
            # ... (Lectura de archivo y loop de alumnos)
            # Dentro del loop:
            doc_res, log_img = crear_docx_v9_2(res.text, nombre, diag, grupo, logo_bytes, activar_img)
            # Mostramos el reporte de imágenes en la UI
            if log_img:
                with st.expander(f"Estado de imágenes para {nombre}"):
                    for item in log_img: st.write(item)
