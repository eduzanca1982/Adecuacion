import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
import zipfile
import time
import random
import hashlib
import re
from datetime import datetime

# ============================================================
# Nano Opal HTML v26.0 (SIMPLE + ESTABLE)
# - Modelo fijo: gemini-2.5-flash
# - Output: HTML completo generado por IA
# - Visuales: SVG inline (sin generación externa)
# - UI mínima
# - Sin APIs privadas de Streamlit
# - Sin PDF
# - Sin lógica híbrida
# ============================================================

st.set_page_config(page_title="Nano Opal HTML v26.0", layout="wide")

TEXT_MODEL = "models/gemini-2.5-flash"

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

RETRIES = 4
MIN_HTML_CHARS = 2000

SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

GEN_CFG = {
    "temperature": 0.4,
    "top_p": 0.9,
    "top_k": 40,
    "max_output_tokens": 8192,
}

# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def hash_text(t):
    return hashlib.sha256(t.encode("utf-8")).hexdigest()[:10]

def safe_filename(name):
    s = str(name).replace(" ", "_")
    for ch in ["/","\\",":","*","?","\"","<",">","|"]:
        s = s.replace(ch,"_")
    return s[:120]

def retry_with_backoff(fn):
    last = None
    for i in range(RETRIES):
        try:
            return fn()
        except Exception as e:
            last = e
            time.sleep(min((2**i)+random.random(), 10))
    raise last

def extract_text(resp):
    try:
        cand = resp.candidates[0]
        content = cand.content
        parts = content.parts
        out = ""
        for p in parts:
            if hasattr(p, "text") and p.text:
                out += p.text
        return out.strip()
    except Exception:
        return ""

def looks_like_html(s):
    if not s:
        return False
    s2 = s.lower()
    return "<html" in s2 and "</html>" in s2

def ensure_html(s):
    if not s:
        return ""
    start = s.lower().find("<!doctype")
    if start == -1:
        start = s.lower().find("<html")
    if start != -1:
        s = s[start:]
    if "</html>" not in s.lower():
        s += "\n</html>"
    return s

# ------------------------------------------------------------
# Prompt Builders
# ------------------------------------------------------------

def build_student_prompt(source_text, alumno, grado):
    return f"""
Devuelve UN SOLO documento HTML completo (incluye <!doctype html> ... </html>).
Prohibido markdown.
Nada fuera del HTML.

Objetivo:
Ficha visual de 60 minutos, estilo moderno tipo cards.
Neuroinclusiva (TDAH/dislexia friendly).
Micro pasos concretos.
Carga cognitiva controlada.

Cada card debe incluir:
- Encabezado: "Ítem N"
- Emoji inicial (✍️📖🔢🎨)
- 2-6 pasos concretos
- Zona Trabajo (caja de respuesta o checkboxes)
- Pista concreta
- Un SVG inline simple dentro de <div class="imgbox">...</div>

Alumno:
Nombre: {alumno["nombre"]}
Grupo: {alumno["grupo"]}
Grado: {grado}
Perfil de aprendizaje: {alumno["perfil"]}

Contenido base:
{source_text}

Salida: HTML completo.
"""

def build_teacher_prompt(student_html, alumno, grado):
    return f"""
Devuelve UN SOLO documento HTML completo.
Prohibido markdown.
Nada fuera del HTML.

Objetivo:
Solucionario docente alineado con los mismos ítems del alumno.

Debe incluir:
- Respuesta final por ítem
- Desarrollo breve
- Errores frecuentes
- Adecuaciones aplicadas
- Criterios de corrección

Alumno:
Nombre: {alumno["nombre"]}
Grupo: {alumno["grupo"]}
Grado: {grado}
Perfil: {alumno["perfil"]}

HTML del alumno:
{student_html}

Salida: HTML completo.
"""

# ------------------------------------------------------------
# Core
# ------------------------------------------------------------

def generate_html(prompt):
    m = genai.GenerativeModel(TEXT_MODEL)
    resp = retry_with_backoff(
        lambda: m.generate_content(prompt, generation_config=GEN_CFG, safety_settings=SAFETY_SETTINGS)
    )
    txt = extract_text(resp)
    html = ensure_html(txt)
    return html

def robust_generate(prompt):
    html = generate_html(prompt)
    if looks_like_html(html) and len(html) > MIN_HTML_CHARS:
        return html
    # retry más fuerte
    prompt2 = prompt + "\nLa salida anterior fue incompleta. Devuelve HTML largo y completo.\n"
    html2 = generate_html(prompt2)
    return html2

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------

def main():

    st.title("Nano Opal HTML v26.0")
    st.caption("Interfaz simple · HTML consistente · SVG inline · Modelo fijo")

    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    except Exception as e:
        st.error(f"API KEY error: {e}")
        return

    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando planilla: {e}")
        return

    grado_col = df.columns[1]
    alumno_col = df.columns[2]
    grupo_col = df.columns[3]
    perfil_col = df.columns[4]

    with st.sidebar:
        grado = st.selectbox("Grado", sorted(df[grado_col].dropna().unique()))
        df_f = df[df[grado_col] == grado]

        alumnos = st.multiselect("Alumnos", sorted(df_f[alumno_col].dropna().unique()))
        if alumnos:
            df_final = df_f[df_f[alumno_col].isin(alumnos)]
        else:
            df_final = df_f

    with st.form("form_main"):
        brief = st.text_area(
            "Prompt / Contenido base",
            height=250,
            placeholder="Ej: Actividad de 60 minutos sobre proporcionalidad directa..."
        )
        submitted = st.form_submit_button("Generar lote")

    if not submitted:
        return

    if not brief.strip():
        st.error("Contenido vacío.")
        return

    if len(df_final) == 0:
        st.error("No hay alumnos seleccionados.")
        return

    run_id = hash_text(now_str() + brief)

    zip_io = io.BytesIO()
    ok = 0
    err = 0

    with zipfile.ZipFile(zip_io, "w", compression=zipfile.ZIP_DEFLATED) as zf:

        for i, (_, row) in enumerate(df_final.iterrows(), start=1):
            nombre = str(row[alumno_col])
            grupo = str(row[grupo_col])
            perfil = str(row[perfil_col])

            alumno = {
                "nombre": nombre,
                "grupo": grupo,
                "perfil": perfil
            }

            try:
                st.write(f"Generando: {nombre}")

                p_student = build_student_prompt(brief, alumno, grado)
                student_html = robust_generate(p_student)

                p_teacher = build_teacher_prompt(student_html, alumno, grado)
                teacher_html = robust_generate(p_teacher)

                base = safe_filename(f"{grado}_{grupo}_{nombre}")

                zf.writestr(f"{base}__ALUMNO.html", student_html)
                zf.writestr(f"{base}__DOCENTE.html", teacher_html)

                ok += 1

            except Exception as e:
                err += 1
                zf.writestr(f"ERROR_{safe_filename(nombre)}.txt", str(e))

    st.success(f"Proceso finalizado. OK={ok} | Errores={err}")

    st.download_button(
        "Descargar ZIP",
        data=zip_io.getvalue(),
        file_name=f"NanoOpal_{grado}_{run_id}.zip",
        mime="application/zip",
        use_container_width=True
    )

if __name__ == "__main__":
    main()
