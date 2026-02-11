import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import zipfile
import time
import random
import hashlib
from datetime import datetime
from typing import Any, Dict, List, Optional
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# 1. CONFIGURACIÓN Y BOOT ESTRATÉGICO
# ============================================================
st.set_page_config(page_title="Motor Pedagógico Opal v19.0", layout="wide", page_icon="🍎")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

@st.cache_resource
def boot_scan():
    """Detecta el mejor hardware lógico disponible en la cuenta."""
    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
        modelos = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        txt_m = next((m for m in modelos if "1.5-flash" in m), modelos[0])
        img_m = next((m for m in modelos if "imagen" in m or "image" in m), None)
        return {"txt": txt_m, "img": img_m}
    except: return {"txt": None, "img": None}

CONFIG = boot_scan()

# --- PROMPT CORE (Basado en Estándar Opal Inclusivo) ---
SYSTEM_PROMPT_OPAL = """
Actúa como un Senior Inclusive UX Designer y Tutor Psicopedagogo.
OBJETIVO: Crear fichas de trabajo inclusivas con estética "Opal" (limpia, visual, estructurada).

REGLAS DE DISEÑO:
- ICONOS: Inicia enunciados con ✍️ (completar), 📖 (leer), 🔢 (calcular), 🎨 (dibujar).
- LENGUAJE: Verdana 14, interlineado 1.5. Sin itálicas. **Negrita** solo en palabras clave.
- PISTAS: Micro-pasos de acción (ej: 'Usa tus dedos para contar 3').
- IMAGEN: 'Pictograma estilo ARASAAC, trazos negros gruesos, fondo blanco de: [OBJETO]'.

SALIDA: JSON puro con: objetivo_aprendizaje, consigna_adaptada, items (enunciado, opciones, pista_visual), adecuaciones_aplicadas.
"""

# ============================================================
# 2. MOTOR DE DISEÑO VISUAL (DOCX STYLING)
# ============================================================
def set_cell_shading(cell, color):
    """Aplica el look de 'Card' de Google (fondo gris suave)."""
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color)
    cell._tc.get_or_add_tcPr().append(shd)

def render_ficha_opal(data, logo_b, img_m_id):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Verdana'
    style.font.size = Pt(14)
    
    # Encabezado de Ficha
    header = doc.add_table(rows=1, cols=2)
    header.width = Inches(6)
    if logo_b:
        header.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_b), width=Inches(0.8))
    
    p_info = header.rows[0].cells[1].paragraphs[0]
    p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    al = data.get("alumno", {})
    p_info.add_run(f"ALUMNO: {al['nombre']}\nAPOYO: {al['diagnostico']}").bold = True

    # Bloque Objetivo (Card sombreada)
    doc.add_paragraph()
    obj_table = doc.add_table(rows=1, cols=1)
    obj_table.style = 'Table Grid'
    cell_obj = obj_table.rows[0].cells[0]
    set_cell_shading(cell_obj, "F8F8F8")
    p_obj = cell_obj.paragraphs[0]
    p_obj.add_run("🎯 OBJETIVO: ").bold = True
    p_obj.add_run(data.get("objetivo_aprendizaje", ""))

    # Consigna y Actividades
    doc.add_paragraph()
    p_cons = doc.add_paragraph()
    p_cons.add_run("CONSIGNA: ").bold = True
    p_cons.add_run(data.get("consigna_adaptada", ""))

    for it in data.get("items", []):
        it_table = doc.add_table(rows=1, cols=1)
        it_table.style = 'Table Grid'
        c = it_table.rows[0].cells[0]
        set_cell_shading(c, "FFFFFF") # Fondo blanco para ítems
        
        p_enun = c.paragraphs[0]
        p_enun.add_run(it.get("enunciado", "")).bold = True
        
        opts = it.get("opciones", [])
        if opts:
            for opt in opts: doc.add_paragraph(f"  ☐ {opt}", style='List Bullet')
        else:
            doc.add_paragraph("\n✍️ Mi respuesta: __________________________")

        p_pista = doc.add_paragraph()
        p_pista.add_run(f" 💡 {it.get('pista_visual', '')}").font.color.rgb = RGBColor(0, 102, 0)
        doc.add_paragraph()

    # Footer Pedagógico
    doc.add_page_break()
    doc.add_heading("Justificación de Adecuaciones", level=2)
    doc.add_paragraph(", ".join(data.get("adecuaciones_aplicadas", [])))

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ============================================================
# 3. INTERFAZ DE USUARIO (UX OPTIMIZADA)
# ============================================================
def main():
    if not CONFIG["txt"]: st.error("Sin acceso a API Google."); return

    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except: st.error("Planilla no encontrada."); return

    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/learning.png", width=80)
        st.header("Configuración")
        grado = st.selectbox("Grado Escolar", df.iloc[:, 1].unique())
        df_f = df[df.iloc[:, 1] == grado]
        
        al_sel = st.multiselect("Alumnos específicos", df_f.iloc[:, 2].unique(), help="Si no eliges ninguno, se procesará todo el grado.")
        df_final = df_f[df_f.iloc[:, 2].isin(al_sel)] if al_sel else df_f

        st.divider()
        logo = st.file_uploader("Logo Institucional", type=["png", "jpg"])
        l_bytes = logo.read() if logo else None

    # CUERPO PRINCIPAL - EL CENTRO DE CONTROL
    st.title("Generador Pedagógico Opal v19.0")
    
    modo = st.tabs(["🔄 Adaptar DOCX", "✨ Crear desde Cero"])
    
    with modo[0]:
        archivo = st.file_uploader("Sube el examen de la maestra", type=["docx"])
        instrucciones_adaptar = st.text_area("Indicaciones extra para la adaptación:", placeholder="Ej: Hazlo más visual, usa solo 4 opciones para los múltiples choices...")

    with modo[1]:
        brief = st.text_area("¿Qué actividad necesitas crear hoy?", height=150, placeholder="Ej: Crea una actividad de división por dos cifras con temática de Minecraft, incluye 5 ejercicios y una pista visual por cada uno.")

    if st.button("🚀 INICIAR GENERACIÓN POR LOTE"):
        # Validación de entrada
        input_data = ""
        if archivo:
            from docx import Document as DocRead
            input_data = "\n".join([p.text for p in DocRead(archivo).paragraphs])
        
        if not input_data and not brief:
            st.warning("Por favor, sube un archivo o escribe una idea en el modo 'Crear'.")
            return

        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as zf:
            prog = st.progress(0.0)
            status = st.empty()
            
            for i, (_, row) in enumerate(df_final.iterrows()):
                n, g, d = str(row.iloc[2]), str(row.iloc[3]), str(row.iloc[4])
                status.info(f"Generando ficha Opal para: {n}")
                
                try:
                    m = genai.GenerativeModel(CONFIG["txt"])
                    ctx = brief if brief else f"Adaptar examen: {input_data}\nInstrucciones extra: {instrucciones_adaptar}"
                    full_p = f"{SYSTEM_PROMPT_OPAL}\nALUMNO: {n} ({d})\nCONTEXTO: {ctx}"
                    
                    res = m.generate_content(full_p, generation_config={"response_mime_type": "application/json"})
                    data = json.loads(res.text)
                    data["alumno"] = {"nombre": n, "diagnostico": d}
                    
                    docx_b = render_ficha_opal(data, l_bytes, CONFIG["img"])
                    zf.writestr(f"Actividad_Inclusiva_{n.replace(' ', '_')}.docx", docx_b)
                except Exception as e:
                    zf.writestr(f"ERROR_{n}.txt", str(e))
                prog.progress((i+1)/len(df_final))

        st.success(f"¡Lote de {len(df_final)} alumnos finalizado!")
        st.download_button("📥 Descargar Todas las Adecuaciones (ZIP)", zip_io.getvalue(), f"Opal_Adecuaciones_{grado}.zip")

if __name__ == "__main__": main()
