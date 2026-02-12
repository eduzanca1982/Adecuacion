import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
import zipfile
import time
import random
import hashlib
import base64
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

# 

# ============================================================
# 1. CONFIGURACIÓN Y ESTILOS UI
# ============================================================
st.set_page_config(page_title="Nano Opal v25.2", layout="wide", page_icon="🧠")

# Inyección de CSS para limpiar la interfaz de Streamlit
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 12px; height: 3.5em; background-color: #7C3AED; color: white; font-weight: bold; }
    .stTextArea textarea { border-radius: 12px; border: 1px solid #E5E7EB; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #F9FAFB; border-radius: 8px 8px 0 0; gap: 1px; }
    .stTabs [aria-selected="true"] { background-color: #FFFFFF; border-bottom: 2px solid #7C3AED !format; }
    </style>
    """, unsafe_allow_html=True)

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
TEXT_MODEL_ID = "models/gemini-2.5-flash"

# ============================================================
# 2. TEMAS Y DISEÑO VISUAL (SISTEMA DE TOKENS)
# ============================================================
THEMES = {
    "Opal Clean (Dark)": {"bg": "#0B1020", "paper": "#0F172A", "card": "#111C35", "ink": "#EAF0FF", "muted": "#A8B3D6", "accent": "#7C3AED", "good": "#22C55E", "line": "rgba(255,255,255,0.1)"},
    "Paper Bright (Light)": {"bg": "#F4F6FB", "paper": "#FFFFFF", "card": "#F8FAFF", "ink": "#0B1220", "muted": "#42526E", "accent": "#2563EB", "good": "#16A34A", "line": "rgba(15,23,42,0.1)"}
}

def build_css_v25(theme_name):
    t = THEMES[theme_name]
    return f"""
    :root {{
        --bg: {t['bg']}; --paper: {t['paper']}; --card: {t['card']};
        --ink: {t['ink']}; --muted: {t['muted']}; --accent: {t['accent']};
        --good: {t['good']}; --line: {t['line']};
    }}
    body {{ background: var(--bg); color: var(--ink); font-family: Verdana, sans-serif; margin: 0; padding: 20px; }}
    .paper {{ background: var(--paper); max-width: 850px; margin: auto; padding: 30px; border-radius: 20px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); border: 1px solid var(--line); }}
    .card {{ background: var(--card); border: 1px solid var(--line); padding: 20px; border-radius: 15px; margin-bottom: 20px; page-break-inside: avoid; }}
    .pista {{ border-left: 5px solid var(--good); background: rgba(34,197,94,0.1); padding: 15px; border-radius: 8px; font-style: italic; }}
    .imgbox {{ background: rgba(255,255,255,0.05); border: 1px dashed var(--line); border-radius: 10px; padding: 10px; text-align: center; min-height: 100px; }}
    h1 {{ color: var(--accent); font-size: 24px; }}
    .badge {{ display: inline-block; padding: 4px 12px; background: var(--accent); color: white; border-radius: 99px; font-size: 12px; font-weight: bold; }}
    """

# 

# ============================================================
# 3. LÓGICA DE GENERACIÓN (BLINDADA)
# ============================================================
def clean_ai_html(raw):
    s = re.sub(r'```html\s*', '', raw, flags=re.IGNORECASE)
    s = re.sub(r'```', '', s)
    start = s.lower().find("<!doctype")
    if start == -1: start = s.lower().find("<html")
    return s[start:].strip() if start != -1 else s.strip()

def build_student_prompt(brief, alumno, grado, visual_mode):
    return f"""
    Eres un experto en educación inclusiva. Genera una ficha HTML completa.
    ALUMNO: {alumno['nombre']} (Grado: {grado}, Grupo: {alumno['grupo']})
    PERFIL DE APRENDIZAJE: {alumno['perfil']}
    
    MODO VISUAL: {'Genera un SVG simple' if visual_mode == 'SVG' else 'Crea un placeholder <div class="imgbox">'}
    
    ESTRUCTURA:
    1. Header profesional con datos del alumno.
    2. Objetivo del día.
    3. 8 ítems en formato .card: Cada uno con icono (✍️, 📖, 🔢), consigna clara y una PISTA VERDE (💡) que sea un micro-paso de acción.
    
    CONTENIDO BASE: {brief}
    Responde SOLO con código HTML.
    """

# ============================================================
# 4. APP PRINCIPAL
# ============================================================
def main():
    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error de inicio: {e}"); return

    # --- SIDEBAR (Configuración) ---
    with st.sidebar:
        st.header("Configuración")
        grado = st.selectbox("Elegir Grado", sorted(df.iloc[:, 1].dropna().unique()))
        df_f = df[df.iloc[:, 1] == grado]
        
        alcance = st.radio("Alcance", ["Todo el grado", "Seleccionar alumnos"])
        alumnos_final = df_f if alcance == "Todo el grado" else df_f[df_f.iloc[:, 2].isin(st.multiselect("Alumnos", df_f.iloc[:, 2].unique()))]
        
        st.divider()
        tema_ui = st.selectbox("Tema Visual", list(THEMES.keys()))
        visual_mode = st.selectbox("Modo de Imágenes", ["SVG (IA Dibuja)", "Placeholder"])
        logo = st.file_uploader("Logo Colegio", type=["png", "jpg"])
        l_bytes = base64.b64encode(logo.read()).decode() if logo else ""

    # --- CUERPO (Trabajo) ---
    st.title("Nano Opal v25.2")
    
    tab1, tab2 = st.tabs(["✨ Crear desde Idea", "🔄 Adaptar DOCX/Texto"])
    
    with tab1:
        brief = st.text_area("¿Qué actividad necesitas hoy?", height=200, placeholder="Ej: Sumas y restas con temática de dinosaurios para 2do grado...")
    
    with tab2:
        source_doc = st.text_area("Pega aquí el texto del examen original:", height=200)

    # 

    if st.button("🚀 GENERAR LOTE DE FICHAS"):
        content = brief if brief else source_doc
        if not content: st.warning("Ingresa contenido."); return

        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as zf:
            prog = st.progress(0.0)
            status = st.empty()
            css_base = build_css_v25(tema_ui)
            
            for i, (_, row) in enumerate(alumnos_final.iterrows()):
                nombre, grupo, perfil = str(row.iloc[2]), str(row.iloc[3]), str(row.iloc[4])
                status.info(f"Generando para {nombre}...")
                
                try:
                    m = genai.GenerativeModel(TEXT_MODEL_ID)
                    p = build_student_prompt(content, {"nombre": nombre, "grupo": grupo, "perfil": perfil}, grado, visual_mode)
                    res = m.generate_content(p, generation_config={"temperature": 0.4, "max_output_tokens": 8000})
                    
                    html_raw = clean_ai_html(res.text)
                    # Inyección de CSS y Logo
                    html_final = f"<html><head><style>{css_base}</style></head><body><div class='paper'>"
                    if l_bytes: html_final += f"<img src='data:image/png;base64,{l_bytes}' style='height:60px; float:right;'>"
                    html_final += f"{html_raw}</div></body></html>"
                    
                    zf.writestr(f"Ficha_{nombre.replace(' ', '_')}.html", html_final)
                except Exception as e:
                    zf.writestr(f"ERROR_{nombre}.txt", str(e))
                prog.progress((i + 1) / len(alumnos_final))

        st.success("¡Lote listo!")
        st.download_button("📥 Descargar ZIP", zip_io.getvalue(), f"Opal_{grado}.zip", "application/zip")

if __name__ == "__main__":
    main()
