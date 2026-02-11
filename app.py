import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import zipfile
import time
from datetime import datetime
from typing import Any, Dict, List, Optional
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# 1. CONFIGURACIÓN Y NANO BANANA BOOT
# ============================================================
st.set_page_config(page_title="Nano Opal v20.0", layout="wide", page_icon="🍌")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

@st.cache_resource
def boot_nano():
    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
        # Priorizamos el nuevo modelo Nano Banana para imágenes
        return {"txt": "gemini-1.5-flash", "img": "gemini-2.5-flash-image"}
    except: return {"txt": None, "img": None}

CONFIG = boot_nano()

# --- PROMPT PEDAGÓGICO REFINADO (Estilo Opal Inclusivo) ---
SYSTEM_PROMPT_NANO = """
Actúa como un Diseñador de UX Pedagógica. Genera fichas con estética de "Card".
REGLAS:
- ICONOS: 🔢 (Cálculo), 📖 (Lectura), ✍️ (Escritura), 🎨 (Arte).
- DISEÑO: Verdana 14, interlineado 1.8. Negrita para palabras clave.
- PISTAS: Micro-pasos de acción concreta.
- IMAGEN: 'Pictograma ARASAAC, trazos negros gruesos, fondo blanco de: [OBJETO]'.
SALIDA: JSON puro.
"""

# ============================================================
# 2. RENDERIZADO "CARD STYLE" (Estética Superior)
# ============================================================
def apply_card_style(cell):
    """Simula una tarjeta HTML con bordes y sombreado en Word."""
    tc_pr = cell._tc.get_or_add_tcPr()
    # Sombreado gris muy tenue (Look Google)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), "FAFAFA")
    tc_pr.append(shd)
    # Bordes finos
    tc_borders = OxmlElement('w:tcBorders')
    for b in ['top', 'left', 'bottom', 'right']:
        edge = OxmlElement(f'w:{b}')
        edge.set(qn('w:val'), 'single')
        edge.set(qn('w:sz'), '4')
        edge.set(qn('w:space'), '0')
        edge.set(qn('w:color'), 'E0E0E0')
        tc_borders.append(edge)
    tc_pr.append(tc_borders)

def render_nano_card(data, logo_b, img_m_id):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Verdana'
    style.font.size = Pt(14)
    
    # Header minimalista
    header = doc.add_table(rows=1, cols=2)
    header.width = Inches(6.5)
    if logo_b:
        header.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_b), width=Inches(0.7))
    
    info = header.rows[0].cells[1].paragraphs[0]
    info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    al = data.get("alumno", {})
    info.add_run(f"{al['nombre']}\n{al['diagnostico']}").bold = True

    # Bloque de Actividades (Iteración de Cards)
    for it in data.get("items", []):
        doc.add_paragraph()
        table = doc.add_table(rows=1, cols=1)
        table.width = Inches(6.5)
        cell = table.rows[0].cells[0]
        apply_card_style(cell)
        
        # Enunciado
        p = cell.paragraphs[0]
        p.add_run(it.get("enunciado", "")).bold = True
        p.paragraph_format.space_before = Pt(10)
        
        # Opciones o Respuesta
        opts = it.get("opciones", [])
        if opts:
            for opt in opts:
                doc.add_paragraph(f"  ○ {opt}", style='List Bullet')
        else:
            p_resp = doc.add_paragraph()
            p_resp.add_run("\n✍️ Mi respuesta: __________________________\n")

        # Pista Verde (Andamiaje)
        p_pista = doc.add_paragraph()
        run = p_pista.add_run(f" 💡 {it.get('pista_visual', '')}")
        run.font.color.rgb = RGBColor(0, 150, 0)
        run.italic = True
        
        # Imagen Nano Banana
        if data.get("visual", {}).get("habilitado") and img_m_id:
            try:
                m = genai.GenerativeModel(img_m_id)
                res = m.generate_content("Dibujo simple estilo escolar: " + data["visual"]["prompt"])
                doc.add_paragraph().add_run().add_picture(io.BytesIO(res.candidates[0].content.parts[0].inline_data.data), width=Inches(2.2))
            except: pass

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ============================================================
# 3. INTERFAZ (Tabs Estilo Opal)
# ============================================================
def main():
    st.title("Nano Opal v20.0 🧠🍌")
    
    if not CONFIG["txt"]: st.error("Falta API Key."); return

    df = pd.read_csv(URL_PLANILLA)
    df.columns = [c.strip() for c in df.columns]

    with st.sidebar:
        grado = st.selectbox("Grado", df.iloc[:, 1].unique())
        df_f = df[df.iloc[:, 1] == grado]
        al_sel = st.multiselect("Alumnos", df_f.iloc[:, 2].unique())
        df_final = df_f[df_f.iloc[:, 2].isin(al_sel)] if al_sel else df_f
        logo = st.file_uploader("Logo", type=["png", "jpg"])
        l_bytes = logo.read() if logo else None

    # TABS PARA MODOS
    tab1, tab2 = st.tabs(["🔄 Adaptar DOCX", "✨ Crear Actividad"])
    
    with tab1:
        archivo = st.file_uploader("Examen base", type=["docx"])
        inst_adapt = st.text_area("Notas de estilo para adaptar:")
        
    with tab2:
        brief = st.text_area("Escribe tu idea aquí:", height=150, placeholder="Ej: Sumas con peras y manzanas...")

    if st.button("🚀 GENERAR LOTE NANO"):
        input_text = ""
        if archivo:
            from docx import Document as DocRead
            input_text = "\n".join([p.text for p in DocRead(archivo).paragraphs])
        
        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as zf:
            prog = st.progress(0.0)
            for i, (_, row) in enumerate(df_final.iterrows()):
                n, g, d = str(row.iloc[2]), str(row.iloc[3]), str(row.iloc[4])
                try:
                    m = genai.GenerativeModel(CONFIG["txt"])
                    ctx = brief if brief else f"Adaptar: {input_text}\nNotas: {inst_adapt}"
                    full_p = f"{SYSTEM_PROMPT_NANO}\nALUMNO: {n} ({d})\nCONTEXTO: {ctx}"
                    
                    res = m.generate_content(full_p, generation_config={"response_mime_type": "application/json"})
                    data = json.loads(res.text)
                    data["alumno"] = {"nombre": n, "diagnostico": d}
                    
                    zf.writestr(f"Ficha_{n}.docx", render_nano_card(data, l_bytes, CONFIG["img"]))
                except Exception as e:
                    zf.writestr(f"ERROR_{n}.txt", str(e))
                prog.progress((i+1)/len(df_final))

        st.success("Lote Nano finalizado.")
        st.download_button("📥 Descargar ZIP", zip_io.getvalue(), "nano_opal.zip")

if __name__ == "__main__": main()
