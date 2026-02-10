import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import zipfile
import time
import random
import hashlib
from typing import Any, Dict, List, Optional
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================
st.set_page_config(page_title="Motor Pedagógico Determinista v13.4", layout="wide")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

MODEL_TEXT_DEFAULT = "gemini-1.5-flash"
MODEL_IMAGE_DEFAULT = "mini-2.0-flash-exp-image-gen"

GEN_CFG_JSON = {
    "response_mime_type": "application/json",
    "temperature": 0,
    "top_p": 1,
    "top_k": 1,
    "max_output_tokens": 4096,
}

SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

RETRIES = 6
CACHE_TTL = 6 * 60 * 60

# ============================================================
# PYDANTIC (OPCIONAL)
# ============================================================
PYDANTIC_AVAILABLE = False
try:
    from pydantic import BaseModel

    class VisualModel(BaseModel):
        habilitado: bool
        prompt: Optional[str] = None

    class ItemModel(BaseModel):
        tipo: str
        enunciado_original: str
        pista: str
        visual: VisualModel

    class AlumnoModel(BaseModel):
        nombre: str
        grupo: str
        diagnostico: str

    class AdecuacionModel(BaseModel):
        alumno: AlumnoModel
        documento: List[ItemModel]

    PYDANTIC_AVAILABLE = True
except Exception:
    PYDANTIC_AVAILABLE = False

# ============================================================
# UTILIDADES
# ============================================================
def hash_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()

def retry(fn):
    for i in range(RETRIES):
        try:
            return fn()
        except Exception:
            if i == RETRIES - 1:
                raise
            time.sleep((2 ** i) + random.uniform(0, 0.5))

def normalize_bool(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    if isinstance(v, str):
        return v.lower() in ["true", "1", "si", "sí", "yes"]
    return False

def normalize_visual(v: Any) -> Dict[str, Any]:
    if not isinstance(v, dict):
        return {"habilitado": False}
    return {
        "habilitado": normalize_bool(v.get("habilitado", False)),
        "prompt": str(v.get("prompt", "")).strip(),
    }

# ============================================================
# EXTRACCIÓN DOCX (PÁRRAFOS + TABLAS)
# ============================================================
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def extract_text(el) -> str:
    return "".join(
        n.text for n in el.iter() if n.tag == f"{W_NS}t" and n.text
    ).strip()

def extract_docx(file) -> str:
    doc = Document(file)
    out: List[str] = []
    for el in doc.element.body:
        if el.tag == f"{W_NS}p":
            t = extract_text(el)
            if t:
                out.append(t)
        elif el.tag == f"{W_NS}tbl":
            for row in el.findall(f".//{W_NS}tr"):
                cells = [extract_text(c) for c in row.findall(f".//{W_NS}tc")]
                if any(cells):
                    out.append("\t".join(cells))
            out.append("")
    return "\n".join(out).strip()

# ============================================================
# VALIDACIÓN Y GENERACIÓN DE IMÁGENES
# ============================================================
def test_image_model(model_id: str) -> bool:
    try:
        m = genai.GenerativeModel(model_id)
        r = retry(lambda: m.generate_content(
            "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: manzana",
            safety_settings=SAFETY_SETTINGS
        ))
        data = r.candidates[0].content.parts[0].inline_data.data
        return bool(data and len(data) > 500)
    except Exception:
        return False

def generate_image(model_id: str, prompt: str) -> Optional[io.BytesIO]:
    try:
        m = genai.GenerativeModel(model_id)
        r = retry(lambda: m.generate_content(prompt, safety_settings=SAFETY_SETTINGS))
        data = r.candidates[0].content.parts[0].inline_data.data
        if not data or len(data) < 500:
            return None
        return io.BytesIO(data)
    except Exception:
        return None

# ============================================================
# GEMINI → JSON ESTRICTO
# ============================================================
BASE_PROMPT = """
Devuelve EXCLUSIVAMENTE un JSON válido.

Esquema:
{
 "alumno": { "nombre": "...", "grupo": "...", "diagnostico": "..." },
 "documento": [
   {
     "tipo": "consigna",
     "enunciado_original": "texto literal",
     "pista": "pista pedagógica",
     "visual": { "habilitado": boolean, "prompt": "opcional" }
   }
 ]
}

Reglas:
- No omitir consignas
- No dar respuestas
- enunciado_original debe ser literal
- visual.prompt debe empezar con:
  "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: "
- Nada fuera del JSON
""".strip()

@st.cache_data(ttl=CACHE_TTL)
def get_json(prompt: str, model_id: str) -> Dict[str, Any]:
    m = genai.GenerativeModel(model_id)
    r = retry(lambda: m.generate_content(
        prompt,
        generation_config=GEN_CFG_JSON,
        safety_settings=SAFETY_SETTINGS
    ))
    return json.loads(r.text)

def request_json(nombre: str, diag: str, grupo: str, examen: str, model_id: str) -> Dict[str, Any]:
    p = f"{BASE_PROMPT}\nAlumno: {nombre}\nGrupo: {grupo}\nDiagnóstico: {diag}\nEXAMEN:\n{examen}"
    data = get_json(p, model_id)
    if PYDANTIC_AVAILABLE:
        AdecuacionModel.model_validate(data)
    return data

# ============================================================
# RENDER DOCX
# ============================================================
def render_docx(data: Dict[str, Any], logo: Optional[bytes], allow_images: bool, model_img: str) -> bytes:
    doc = Document()
    tbl = doc.add_table(rows=1, cols=2)

    if logo:
        tbl.rows[0].cells[0].paragraphs[0].add_run().add_picture(
            io.BytesIO(logo), width=Inches(0.8)
        )

    info = tbl.rows[0].cells[1].paragraphs[0]
    info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    a = data["alumno"]
    info.add_run(
        f"ALUMNO: {a['nombre']}\nGRUPO: {a['grupo']} | APOYO: {a['diagnostico']}"
    ).bold = True

    green = RGBColor(0, 128, 0)
    for it in data["documento"]:
        doc.add_paragraph(it["enunciado_original"])
        p = doc.add_paragraph()
        r = p.add_run("💡 " + it["pista"])
        r.font.color.rgb = green
        r.italic = True

        vis = normalize_visual(it.get("visual", {}))
        if allow_images and vis["habilitado"] and vis["prompt"]:
            img = generate_image(model_img, vis["prompt"])
            if img:
                d = doc.add_paragraph()
                d.alignment = WD_ALIGN_PARAGRAPH.CENTER
                d.add_run().add_picture(img, width=Inches(2.5))
        doc.add_paragraph()

    b = io.BytesIO()
    doc.save(b)
    return b.getvalue()

# ============================================================
# STREAMLIT UI
# ============================================================
def main():
    st.title("Motor Pedagógico Determinista v13.4")

    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    except Exception:
        st.error("GOOGLE_API_KEY no configurada en st.secrets")
        return

    df = pd.read_csv(URL_PLANILLA)
    df.columns = [c.strip() for c in df.columns]

    with st.sidebar:
        model_text = st.text_input("Modelo texto", MODEL_TEXT_DEFAULT)
        model_img = st.text_input("Modelo imagen", MODEL_IMAGE_DEFAULT)

        use_images = st.checkbox("Generar imágenes", True)
        image_ok = False
        if use_images:
            image_ok = test_image_model(model_img)
            if not image_ok:
                st.warning("Modelo de imagen no disponible. Se usará solo texto.")

        logo = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
        logo_bytes = logo.read() if logo else None

        grado = st.selectbox("Grado", df.iloc[:, 1].unique())
        df_f = df[df.iloc[:, 1] == grado]

    file = st.file_uploader("Examen DOCX", type=["docx"])
    if not file:
        return

    if st.button("Iniciar procesamiento"):
        exam = extract_docx(file)
        if not exam:
            st.error("Examen vacío")
            return

        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as z:
            bar = st.progress(0.0)
            for i, (_, r) in enumerate(df_f.iterrows()):
                try:
                    data = request_json(r.iloc[2], r.iloc[4], r.iloc[3], exam, model_text)
                    docx = render_docx(data, logo_bytes, use_images and image_ok, model_img)
                    z.writestr(
                        f"Adecuacion_{str(r.iloc[2]).replace(' ', '_')}.docx",
                        docx,
                    )
                except Exception as e:
                    z.writestr(
                        f"ERROR_{str(r.iloc[2]).replace(' ', '_')}.txt",
                        str(e),
                    )
                bar.progress((i + 1) / len(df_f))

        st.download_button("Descargar ZIP", zip_io.getvalue(), "adecuaciones.zip")

if __name__ == "__main__":
    main()
