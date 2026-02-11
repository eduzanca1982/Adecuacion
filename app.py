import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import zipfile
import time
import hashlib
from datetime import datetime
from typing import Any, Dict, List, Optional
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# 1. BOOT SCAN: ESCANEO REAL DE MODELOS (EVITA EL 404)
# ============================================================
st.set_page_config(page_title="Nano Opal v21.0", layout="wide", page_icon="🍌")

@st.cache_resource(show_spinner="Escaneando modelos autorizados...")
def boot_system():
    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
        # Listamos todos los modelos que soportan generación de contenido
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        
        # Selección dinámica de Texto (Ranking de potencia)
        txt_candidates = ["models/gemini-1.5-flash", "models/gemini-1.5-pro", "gemini-1.5-flash", "gemini-1.5-pro"]
        text_model = next((m for m in available_models if any(c in m for c in txt_candidates)), available_models[0])
        
        # Selección dinámica de Imagen (Nano Banana o Imagen)
        img_candidates = ["models/gemini-2.5-flash-image", "models/imagen-3.0", "imagen-3.0", "gemini-2.5-flash-image"]
        image_model = next((m for m in available_models if any(c in m for c in img_candidates)), None)
        
        return {"txt": text_model, "img": image_model, "all": available_models}
    except Exception as e:
        return {"txt": None, "img": None, "error": str(e)}

BOOT = boot_system()

# --- PROMPT PEDAGÓGICO OPAL ---
SYSTEM_PROMPT = """
Actúa como un Diseñador de UX Pedagógica. Genera fichas con estética de "Card".
REGLAS:
- ICONOS AL INICIO: 🔢 (Cálculo), 📖 (Lectura), ✍️ (Escritura), 🎨 (Arte).
- DISEÑO: Verdana 14, interlineado 1.8. Negrita para palabras clave.
- PISTAS: Micro-pasos de acción concreta (ej: 'Busca el número que termina en 5').
- IMAGEN: 'Pictograma ARASAAC, trazos negros gruesos, fondo blanco de: [OBJETO]'.
SALIDA: JSON puro.
"""

# ============================================================
# 2. DISEÑO VISUAL "CARD" (TABLAS CON SOMBREADO)
# ============================================================
def apply_card_style(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    # Fondo Gris Google (Opal Style)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), "FAFAFA")
    tc_pr.append(shd)
    # Bordes sutiles
    tc_borders = OxmlElement('w:tcBorders')
    for b in ['top', 'left', 'bottom', 'right']:
        edge = OxmlElement(f'w:{b}')
        edge.set(qn('w:val'), 'single')
        edge.set(qn('w:sz'), '4')
        edge.set(qn('w:color'), 'E0E0E0')
        tc_borders.append(edge)
    tc_pr.append(tc_borders)

def render_docx_v21(data, logo_b, img_model_id):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Verdana'
    style.font.size = Pt(14)
    
    # Header minimalista
    header = doc.add_table(rows=1, cols=2)
    header.width = Inches(6.5)
    if logo_b:
        header.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_b), width=Inches(0.75))
    
    info = header.rows[0].cells[1].paragraphs[0]
    info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    al = data.get("alumno", {})
    info.add_run(f"{al.get('nombre','')}\n{al.get('diagnostico','')}").bold = True

    # Render de Ítems como Tarjetas
    for it in data.get("items", []):
        doc.add_paragraph()
        table = doc.add_table(rows=1, cols=1)
        table.width = Inches(6.5)
        cell = table.rows[0].cells[0]
        apply_card_style(cell)
        
        # Enunciado
        p = cell.paragraphs[0]
        p.add_run(it.get("enunciado", "")).bold = True
        
        # Opciones o Línea de respuesta
        opts = it.get("opciones", [])
        if opts:
            for opt in opts:
                doc.add_paragraph(f"  ○ {opt}", style='List Bullet')
        else:
            p_resp = doc.add_paragraph()
            p_resp.add_run("\n✍️ Mi respuesta: __________________________\n")

        # Pista Verde
        p_pista = doc.add_paragraph()
        run = p_pista.add_run(f" 💡 Pista: {it.get('pista_visual', it.get('pista', ''))}")
        run.font.color.rgb = RGBColor(0, 150, 0)
        run.italic = True
        
        # Imagen Nano Banana
        if data.get("visual", {}).get("habilitado") and img_model_id:
            try:
                m = genai.GenerativeModel(img_model_id)
                res = m.generate_content("Dibujo escolar simple: " + data["visual"]["prompt"])
                doc.add_paragraph().add_run().add_picture(io.BytesIO(res.candidates[0].content.parts[0].inline_data.data), width=Inches(2.5))
            except: pass

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ============================================================
# 3. INTERFAZ Y PROCESO
# ============================================================
def main():
    if BOOT.get("txt") is None:
        st.error(f"Error de modelos: {BOOT.get('error')}")
        return

    SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
    URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
    
    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except: st.error("No se pudo cargar la planilla."); return

    with st.sidebar:
        st.header("⚙️ Configuración")
        st.success(f"Modelo Texto: {BOOT['txt']}")
        st.write(f"Modelo Imagen: {BOOT['img'] or 'No detectado'}")
        
        grado = st.selectbox("Grado", df.iloc[:, 1].unique())
        df_f = df[df.iloc[:, 1] == grado]
        al_sel = st.multiselect("Alumnos", df_f.iloc[:, 2].unique())
        df_final = df_f[df_f.iloc[:, 2].isin(al_sel)] if al_sel else df_f
        
        logo = st.file_uploader("Logo Colegio", type=["png", "jpg"])
        l_bytes = logo.read() if logo else None

    # TABS CENTRALES
    tab1, tab2 = st.tabs(["🔄 Adaptar DOCX", "✨ Crear desde Cero"])
    
    with tab1:
        archivo = st.file_uploader("Subir examen original", type=["docx"])
        inst_extra = st.text_area("Instrucciones de estilo (Modo Adaptar):")
        
    with tab2:
        brief = st.text_area("Describí la actividad a crear:", height=150)

    if st.button("🚀 INICIAR LOTE"):
        input_content = ""
        if archivo:
            from docx import Document as DocRead
            input_content = "\n".join([p.text for p in DocRead(archivo).paragraphs])
        
        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as zf:
            prog = st.progress(0.0)
            status = st.empty()
            
            for i, (_, row) in enumerate(df_final.iterrows()):
                n, g, d = str(row.iloc[2]), str(row.iloc[3]), str(row.iloc[4])
                status.info(f"Generando ficha para: {n}")
                try:
                    m = genai.GenerativeModel(BOOT["txt"])
                    ctx = brief if brief else f"Adaptar: {input_content}\nExtra: {inst_extra}"
                    full_p = f"{SYSTEM_PROMPT}\nALUMNO: {n} ({d}, {g})\nCONTEXTO: {ctx}"
                    
                    res = m.generate_content(full_p, generation_config={"response_mime_type": "application/json"})
                    data = json.loads(res.text)
                    data["alumno"] = {"nombre": n, "diagnostico": d, "grupo": g}
                    
                    zf.writestr(f"Ficha_{n}.docx", render_docx_v21(data, l_bytes, BOOT["img"]))
                except Exception as e:
                    zf.writestr(f"ERROR_{n}.txt", str(e))
                prog.progress((i+1)/len(df_final))

        st.success("¡Lote finalizado!")
        st.download_button("📥 Descargar ZIP", zip_io.getvalue(), "adecuaciones_v21.zip")

if __name__ == "__main__":
    main()
