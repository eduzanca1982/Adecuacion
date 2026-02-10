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
OPCIONES_MODELOS = ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash"]

try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
except Exception as e:
    st.error(f"Falta GOOGLE_API_KEY: {e}")

# PROMPT MAESTRO v6.7 (ESTRICTO)
SYSTEM_PROMPT = """Eres un Diseñador Editorial Pedagógico. Genera el examen FINAL.

REGLAS DE SILENCIO:
1. PROHIBIDO incluir frases de análisis, ayuda, intros o explicaciones como "*Análisis:*", "*Ayuda:*", "Aquí tienes".
2. Si no hay una adecuación que aporte valor real, devuelve el texto original exacto.

REGLAS DE RESALTE (SÓLO NÚCLEO):
1. NO resaltes conectores (y, con, por, de, fue, el, la).
2. SÓLO resalta en **negrita** la información que responde a las preguntas (ej: "al parque", "con su hermano").

REGLAS DE ESPACIO:
1. SOLO usa [CUADRICULA] donde el alumno deba escribir.
2. NO agregues líneas de puntos al principio o final de los textos de lectura."""

# 2. FUNCIONES DE DISEÑO
def limpiar_nombre(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", str(nombre)).replace(" ", "_")

def crear_docx_final(texto_ia, nombre, diagnostico, grupo, logo_bytes=None):
    doc = Document()
    diag = str(diagnostico).lower()
    grupo = str(grupo).upper()
    color_inst = RGBColor(31, 73, 125)

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
    run = p.add_run(f"ALUMNO: {nombre.upper()}\nAPOYO: {diagnostico.upper()} | GRUPO: {grupo}")
    run.bold = True
    run.font.color.rgb = color_inst

    style = doc.styles['Normal']
    font = style.font
    is_apo = any(x in diag for x in ["dislexia", "discalculia", "general"]) or grupo == "A"
    font.name = 'OpenDyslexic' if is_apo else 'Verdana'
    font.size = Pt(12 if is_apo else 11)
    style.paragraph_format.line_spacing = 1.5 if is_apo else 1.15

    for linea in texto_ia.split('\n'):
        linea = linea.strip()
        
        # --- FILTRO ANTIBASURA ---
        # Si la línea contiene palabras de análisis o solo puntos, la ignoramos.
        if any(x in linea.lower() for x in ["análisis:", "ayuda:", "analisis:", "ayuda memoria", "recordá:"]):
            continue
        if re.match(r'^\.*$', linea): # Ignora líneas que son solo puntos
            continue
        
        if not linea: continue
        
        para = doc.add_paragraph()
        
        if "[CUADRICULA]" in linea:
            for _ in range(2):
                p_g = doc.add_paragraph()
                p_g.add_run(" " + "." * 70).font.color.rgb = RGBColor(215, 215, 215)
                p_g.paragraph_format.space_after = Pt(0)
            continue

        if "💡" in linea:
            run_p = para.add_run(linea)
            run_p.font.color.rgb = RGBColor(0, 102, 0)
            run_p.italic = True
            continue

        es_titulo = "[TITULO]" in linea or (len(linea) < 55 and not linea.endswith('.'))
        texto_limpio = linea.replace("[TITULO]", "").strip()
        
        partes = texto_limpio.split("**")
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

# 3. INTERFAZ (Se mantiene la lógica de modelos y descarga)
# ... (Igual que v6.6)
