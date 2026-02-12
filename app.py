import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
import zipfile
import base64
import json
from datetime import datetime
from weasyprint import HTML

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(page_title="Opal Classroom v28 PRO", layout="wide")

TEXT_MODEL = "gemini-2.5-flash"
IMAGE_MODEL = "gemini-2.5-flash-image"

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

# ============================================================
# CSS PREMIUM A4
# ============================================================

GLOBAL_CSS = """
@page { size: A4; margin: 30mm 20mm 25mm 20mm; }

body {
    font-family: Verdana, sans-serif;
    background: #f4f6fb;
}

.paper {
    background: white;
    padding: 40px;
    border-radius: 12px;
}

.header {
    border-bottom: 4px solid #7C3AED;
    padding-bottom: 15px;
    margin-bottom: 30px;
}

.student-name {
    font-size: 26px;
    font-weight: bold;
}

.group-badge {
    background: #7C3AED;
    color: white;
    padding: 5px 12px;
    border-radius: 999px;
    font-size: 12px;
}

.card {
    border: 2px solid #e5e7eb;
    border-radius: 18px;
    padding: 25px;
    margin-bottom: 35px;
    line-height: 2;
}

.enunciado {
    font-size: 18px;
    font-weight: bold;
    margin-bottom: 15px;
}

.pista {
    background: #ecfdf5;
    border-left: 6px solid #10b981;
    padding: 15px;
    border-radius: 8px;
    margin-top: 20px;
    font-style: italic;
}

.answer-line {
    border-bottom: 2px solid #cbd5e1;
    height: 35px;
    margin-top: 20px;
}

.img-box {
    text-align: center;
    margin: 20px 0;
}
"""

# ============================================================
# IA – GENERACIÓN
# ============================================================

SYSTEM_PROMPT = """
Actúa como diseñador instruccional senior.

Devuelve JSON con:
{
 "items":[
   {
     "icono":"✍️",
     "enunciado":"...",
     "pista":"...",
     "prompt_imagen":"..."
   }
 ]
}

Reglas:
- Actividades con sentido real.
- Pista adaptada a dificultad del alumno.
- Imagen pedagógica clara.
- No markdown.
"""

def generar_imagen(prompt_visual):
    try:
        m = genai.GenerativeModel(IMAGE_MODEL)
        r = m.generate_content(
            f"Pictograma educativo claro, fondo blanco, estilo simple de: {prompt_visual}"
        )
        return base64.b64encode(
            r.candidates[0].content.parts[0].inline_data.data
        ).decode()
    except:
        return None

def render_html(data, alumno, logo_b64):
    html = f"""
    <html>
    <head>
        <style>{GLOBAL_CSS}</style>
    </head>
    <body>
        <div class="paper">
            <div class="header">
                <span class="group-badge">Grupo {alumno['grupo']}</span>
                <div class="student-name">{alumno['nombre']}</div>
                <div>{alumno['perfil']}</div>
            </div>
    """

    for item in data["items"]:
        img_html = ""
        if item.get("img"):
            img_html = f'<div class="img-box"><img src="data:image/png;base64,{item["img"]}" width="280"></div>'

        html += f"""
        <div class="card">
            <div class="enunciado">{item["icono"]} {item["enunciado"]}</div>
            {img_html}
            <div class="pista">💡 {item["pista"]}</div>
            <div class="answer-line"></div>
            <div class="answer-line"></div>
        </div>
        """

    html += "</div></body></html>"
    return html

# ============================================================
# UI SIMPLE
# ============================================================

def main():
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])

    df = pd.read_csv(URL_PLANILLA)
    df.columns = [c.strip() for c in df.columns]

    st.title("Opal Classroom v28 PRO")

    grado = st.selectbox("Grado", sorted(df.iloc[:,1].unique()))
    df_f = df[df.iloc[:,1] == grado]

    alumnos = st.multiselect("Seleccionar alumnos", df_f.iloc[:,2].unique())
    df_final = df_f[df_f.iloc[:,2].isin(alumnos)] if alumnos else df_f

    prompt = st.text_area("¿Qué deben aprender hoy?", height=200)

    logo = st.file_uploader("Logo", type=["png","jpg"])
    logo_b64 = base64.b64encode(logo.read()).decode() if logo else ""

    if st.button("🚀 GENERAR FICHAS", use_container_width=True):

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zf:

            for _, row in df_final.iterrows():

                nombre = str(row.iloc[2])
                grupo = str(row.iloc[3])
                perfil = str(row.iloc[4])

                m = genai.GenerativeModel(TEXT_MODEL)
                p = f"""
{SYSTEM_PROMPT}

ALUMNO:
Nombre: {nombre}
Grupo: {grupo}
Dificultad: {perfil}

CONTENIDO:
{prompt}
"""
                r = m.generate_content(p)
                data = json.loads(r.text)

                for item in data["items"]:
                    item["img"] = generar_imagen(item["prompt_imagen"])

                alumno_dict = {
                    "nombre": nombre,
                    "grupo": grupo,
                    "perfil": perfil
                }

                html = render_html(data, alumno_dict, logo_b64)

                pdf_bytes = HTML(string=html).write_pdf()

                safe_name = nombre.replace(" ", "_")

                zf.writestr(f"{safe_name}.html", html)
                zf.writestr(f"{safe_name}.pdf", pdf_bytes)

        st.success("Lote generado correctamente.")
        st.download_button(
            "Descargar ZIP",
            zip_buffer.getvalue(),
            f"Fichas_{grado}.zip",
            "application/zip"
        )

if __name__ == "__main__":
    main()
