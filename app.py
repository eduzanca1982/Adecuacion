# app.py
import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import zipfile
import time
import random
import hashlib
import re
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============================================================
# CONFIGURACIÓN TÉCNICA Y SEGURIDAD
# ============================================================
st.set_page_config(
    page_title="Motor Pedagógico Determinista v13.4", 
    layout="wide",
    initial_sidebar_state="expanded"
)

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

# Constantes de Entorno
MODEL_TEXT_DEFAULT = "gemini-1.5-flash"
MODEL_IMAGE_DEFAULT = "imagen-3.0"
RETRIES = 6
BACKOFF_BASE_SECONDS = 1.0
CACHE_TTL_SECONDS = 6 * 60 * 60

# Configuración de Generación Estricta
GEN_CFG_JSON = {
    "response_mime_type": "application/json",
    "temperature": 0.0,
    "top_p": 1.0,
    "top_k": 1,
    "max_output_tokens": 4096,
}

# Filtros de Seguridad Educativa
SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_ONLY_HIGH"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

# ============================================================
# VALIDACIÓN DE ESQUEMAS (PYDANTIC)
# ============================================================
PYDANTIC_AVAILABLE = False
try:
    from pydantic import BaseModel, Field, ValidationError

    class VisualModel(BaseModel):
        habilitado: bool = Field(..., description="Determina si requiere imagen")
        prompt: Optional[str] = Field(None, description="Prompt descriptivo para la IA de imagen")

    class ItemModel(BaseModel):
        tipo: str = Field(..., description="Categoría del item")
        enunciado_original: str = Field(..., description="Texto literal del examen")
        pista: str = Field(..., description="Andamiaje pedagógico")
        visual: VisualModel

    class AlumnoModel(BaseModel):
        nombre: str
        grupo: str
        diagnostico: str

    class AdecuacionModel(BaseModel):
        alumno: AlumnoModel
        documento: List[ItemModel]

    PYDANTIC_AVAILABLE = True
except ImportError:
    pass

# ============================================================
# UTILIDADES DE SISTEMA Y CACHÉ
# ============================================================
def get_content_hash(text: str) -> str:
    """Genera huella digital del examen para evitar re-procesamiento."""
    return hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()

def clean_json_string(text: str) -> str:
    """Elimina basura de Markdown y extrae solo el bloque JSON."""
    match = re.search(r'\{.*\}', text, re.DOTALL)
    return match.group(0) if match else text

def retry_with_backoff(fn, retries: int = RETRIES, backoff_in_seconds: float = BACKOFF_BASE_SECONDS):
    """Implementa exponencial backoff para cumplimiento de cuotas."""
    last_err = None
    for attempt in range(retries + 1):
        try:
            return fn()
        except Exception as e:
            last_err = e
            s = str(e).lower()
            retryable = any(m in s for m in ["429", "rate", "quota", "timeout", "503", "500"])
            if not retryable or attempt == retries:
                raise last_err
            time.sleep(min((backoff_in_seconds * (2 ** attempt)) + random.uniform(0, 1), 30))
    raise last_err

# ============================================================
# EXTRACCIÓN DOCX (REFACTORIZADA)
# ============================================================
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def _extract_rich_text(el) -> str:
    """Extrae texto manteniendo coherencia de espacios entre nodos."""
    return "".join(node.text for node in el.iter(f"{W_NS}t") if node.text).strip()

def extraer_contenido_completo(file) -> str:
    """Parsea el DOCX respetando el flujo de lectura (párrafos y tablas)."""
    doc = Document(file)
    buffer: List[str] = []

    for element in doc.element.body:
        if element.tag == f"{W_NS}p":
            t = _extract_rich_text(element)
            if t: buffer.append(t)
        elif element.tag == f"{W_NS}tbl":
            for row in element.findall(f".//{W_NS}tr"):
                cells = [_extract_rich_text(c) for c in row.findall(f".//{W_NS}tc")]
                line = " | ".join(cells)
                if line.strip(): buffer.append(f"[TABLA] {line}")
            buffer.append("") # Separador visual pro-contexto
            
    return "\n".join(buffer).strip()

# ============================================================
# MOTOR DE IMAGEN (DIAGNÓSTICO CACHEADO)
# ============================================================
@st.cache_resource(show_spinner="Validando capacidad de imagen...")
def validate_image_capability(model_id: str) -> Tuple[bool, str]:
    """Testea el modelo de imagen una sola vez por sesión."""
    test_prompt = "Dibujo lineal simple, fondo blanco, una manzana"
    try:
        model = genai.GenerativeModel(model_id)
        res = model.generate_content(test_prompt, safety_settings=SAFETY_SETTINGS)
        if hasattr(res.candidates[0].content.parts[0], "inline_data"):
            return True, "Capacidad de imagen confirmada."
        return False, "El modelo no devolvió datos binarios (inline_data)."
    except Exception as e:
        return False, f"Fallo de validación: {str(e)}"

def generar_imagen_ia(model_id: str, prompt: str) -> Optional[io.BytesIO]:
    """Generador desacoplado con validación de integridad de bytes."""
    try:
        model = genai.GenerativeModel(model_id)
        res = retry_with_backoff(lambda: model.generate_content(prompt, safety_settings=SAFETY_SETTINGS))
        img_data = res.candidates[0].content.parts[0].inline_data.data
        if img_data and len(img_data) > 1000:
            return io.BytesIO(img_data)
        return None
    except:
        return None

# ============================================================
# LÓGICA GEMINI (JSON + REPARACIÓN)
# ============================================================
PROMPT_SYSTEM = """
Genera un JSON ESTRUCTURADO. No incluyas explicaciones.
Esquema:
{
  "alumno": {"nombre": "string", "grupo": "string", "diagnostico": "string"},
  "documento": [
    {
      "tipo": "consigna",
      "enunciado_original": "copia literal",
      "pista": "pista pedagógica breve",
      "visual": {"habilitado": bool, "prompt": "string descriptivo"}
    }
  ]
}
Reglas:
1. enunciado_original DEBE ser literal.
2. Si visual.habilitado=true, prompt DEBE iniciar con: 'Dibujo escolar, trazos negros, fondo blanco, estilo simple de: '.
"""

@st.cache_data(ttl=CACHE_TTL_SECONDS)
def process_student_json(cache_key, model_id, prompt_full) -> Dict[str, Any]:
    """Caché inteligente de adecuaciones."""
    model = genai.GenerativeModel(model_id)
    resp = retry_with_backoff(lambda: model.generate_content(prompt_full, generation_config=GEN_CFG_JSON))
    
    raw_text = clean_json_string(resp.text)
    data = json.loads(raw_text)
    
    if PYDANTIC_AVAILABLE:
        AdecuacionModel.model_validate(data)
    return data

# ============================================================
# RENDERIZADO DETERMINISTA
# ============================================================
def render_docx(data: Dict, logo: Optional[bytes], img_enabled: bool, img_model: str) -> bytes:
    doc = Document()
    
    # Header Dinámico
    table = doc.add_table(rows=1, cols=2)
    if logo:
        try:
            table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo), width=Inches(0.85))
        except: pass
    
    c_info = table.rows[0].cells[1].paragraphs[0]
    c_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    al = data["alumno"]
    header = c_info.add_run(f"ALUMNO: {al['nombre']}\nGRUPO: {al['grupo']} | APOYO: {al['diagnostico']}")
    header.bold = True
    header.font.size = Pt(10)

    # Procesamiento de Consignas
    for item in data.get("documento", []):
        # Texto Original
        p_orig = doc.add_paragraph()
        p_orig.add_run(item["enunciado_original"]).font.size = Pt(11)
        
        # Pista Pedagógica
        p_pista = doc.add_paragraph()
        pista_run = p_pista.add_run(f"💡 {item['pista']}")
        pista_run.font.color.rgb = RGBColor(0, 120, 0)
        pista_run.italic = True
        pista_run.font.size = Pt(10)

        # Inserción de Imagen (si aplica)
        if img_enabled and item["visual"]["habilitado"]:
            p = item["visual"]["prompt"]
            img_stream = generar_imagen_ia(img_model, p)
            if img_stream:
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_stream, width=Inches(2.8))
        
        doc.add_paragraph() # Spacer

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ============================================================
# INTERFAZ DE USUARIO (STREAMLIT)
# ============================================================
def main():
    st.header("🧠 Motor Pedagógico Determinista v13.4")
    
    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    except:
        st.error("API Key no configurada en Secrets.")
        return

    # Carga de Planilla
    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando base de datos: {e}")
        return

    with st.sidebar:
        st.subheader("⚙️ Parámetros de Motor")
        m_text = st.text_input("Modelo de Texto", MODEL_TEXT_DEFAULT)
        m_img = st.text_input("Modelo de Imagen", MODEL_IMAGE_DEFAULT)
        
        st.divider()
        
        # Mapeo de Columnas (Índices Estables)
        g_col, a_col, gr_col, d_col = df.columns[1], df.columns[2], df.columns[3], df.columns[4]
        
        grado = st.selectbox("Seleccionar Grado", sorted(df[g_col].dropna().unique()))
        df_f = df[df[g_col] == grado].copy()
        
        alcance = st.radio("Alcance", ["Todos", "Selección Manual"])
        alumnos_proc = df_f[a_col].tolist()
        if alcance == "Selección Manual":
            alumnos_proc = st.multiselect("Alumnos", alumnos_proc)
        
        st.divider()
        
        req_img = st.checkbox("Generar imágenes IA", value=True)
        valid_img, img_msg = False, ""
        if req_img:
            valid_img, img_msg = validate_image_capability(m_img)
            if not valid_img: st.warning(img_status)
        
        logo = st.file_uploader("Logo Institucional", type=["png", "jpg", "jpeg"])
        logo_data = logo.read() if logo else None

    # Operación Principal
    file_docx = st.file_uploader("Examen Base (DOCX)", type=["docx"])
    
    if file_docx and st.button("🚀 Iniciar Procesamiento por Lote"):
        if not alumnos_proc:
            st.warning("Seleccione al menos un alumno.")
            return

        with st.spinner("Analizando examen..."):
            exam_content = extraer_contenido_completo(file_docx)
            exam_hash = get_content_hash(exam_content)
        
        # Filtro de Alumnos Final
        df_final = df_f[df_f[a_col].isin(alumnos_proc)].copy()
        total_a = len(df_final)
        
        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as zf:
            bar = st.progress(0.0)
            status_box = st.empty()
            
            for idx, (_, row) in enumerate(df_final.iterrows(), 1):
                nom, grp, dia = str(row[a_col]), str(row[gr_col]), str(row[d_col])
                
                status_box.info(f"Procesando {idx}/{total_a}: {nom}")
                
                try:
                    # Generación JSON
                    key = f"{exam_hash}_{nom}_{dia}"
                    prompt = f"{PROMPT_SYSTEM}\nAlumno: {nom}\nGrupo: {grp}\nDiagnóstico: {dia}\nExamen: {exam_content}"
                    
                    data_json = process_student_json(key, m_text, prompt)
                    
                    # Generación DOCX
                    docx_data = render_docx(data_json, logo_data, valid_img, m_img)
                    
                    filename = f"Adecuacion_{nom.replace(' ', '_')}.docx"
                    zf.writestr(filename, docx_data)
                    
                except Exception as ex:
                    err_msg = f"ERROR en {nom}: {str(ex)}"
                    zf.writestr(f"FALLO_{nom.replace(' ', '_')}.txt", err_msg.encode())
                    st.error(err_msg)
                
                bar.progress(idx / total_a)
        
        st.success("🏁 Lote completado.")
        st.download_button(
            label="📥 Descargar ZIP de Adecuaciones",
            data=zip_io.getvalue(),
            file_name=f"Adecuaciones_{grado}_{int(time.time())}.zip",
            mime="application/zip"
        )

if __name__ == "__main__":
    main()
