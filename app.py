import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import zipfile
import time
import random
import hashlib
import re  # IMPORTACIÓN CRÍTICA PARA EVITAR EL CRASH
from typing import Any, Dict, List, Optional, Tuple

# Verificación de Pydantic para evitar el fallo de arranque
try:
    from pydantic import BaseModel, Field, ValidationError
    PYDANTIC_AVAILABLE = True
except ImportError:
    PYDANTIC_AVAILABLE = False

# ============================================================
# 1. CONFIGURACIÓN ESTRUCTURAL Y DE SEGURIDAD
# ============================================================
st.set_page_config(
    page_title="Motor Pedagógico Determinista v13.5", 
    layout="wide",
    page_icon="🧠"
)

# Constantes de conexión
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

# Modelos y Cuotas
MODEL_TEXT_DEFAULT = "gemini-1.5-flash"
MODEL_IMAGE_DEFAULT = "imagen-3.0"
RETRIES = 6
BACKOFF_BASE = 1.0

# Configuraciones de la IA (Determinismo)
GEN_CFG_JSON = {
    "response_mime_type": "application/json",
    "temperature": 0.0,
    "top_p": 1.0,
    "max_output_tokens": 4096,
}

# ============================================================
# 2. ESQUEMAS DE DATOS (Contrato de Inteligencia)
# ============================================================
if PYDANTIC_AVAILABLE:
    class VisualSupport(BaseModel):
        habilitado: bool = Field(default=False)
        prompt: Optional[str] = Field(default="")

    class ExamenItem(BaseModel):
        tipo: str = Field(default="consigna")
        enunciado_original: str
        pista: str
        visual: VisualSupport

    class AdecuacionFinal(BaseModel):
        alumno: Dict[str, str]
        documento: List[ExamenItem]

# ============================================================
# 3. UTILIDADES DE SISTEMA
# ============================================================
def get_content_hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()

def retry_with_backoff(fn, retries: int = RETRIES):
    """Maneja el error 429 con espera exponencial."""
    for attempt in range(retries + 1):
        try:
            return fn()
        except Exception as e:
            if "429" in str(e) and attempt < retries:
                time.sleep((2 ** attempt) + random.uniform(0, 1))
                continue
            raise e

# ============================================================
# 4. EXTRACCIÓN DE DOCX (Fidelidad de Tablas)
# ============================================================
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

def extraer_docx_completo(file) -> str:
    """Extrae párrafos y tablas manteniendo el orden del documento."""
    doc = Document(file)
    buffer = []
    
    # Namespace para identificar tablas y párrafos en el XML
    W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    
    for element in doc.element.body:
        if element.tag == f"{W_NS}p":
            t = "".join(node.text for node in element.iter(f"{W_NS}t") if node.text)
            if t.strip(): buffer.append(t.strip())
        elif element.tag == f"{W_NS}tbl":
            for row in element.findall(f".//{W_NS}tr"):
                cells = ["".join(node.text for node in c.iter(f"{W_NS}t") if node.text).strip() 
                         for c in row.findall(f".//{W_NS}tc")]
                buffer.append(" | ".join(cells))
    return "\n".join(buffer)

# ============================================================
# 5. MOTOR DE INTELIGENCIA Y REPARACIÓN
# ============================================================
def generar_adecuacion_json(nombre, diag, grupo, examen_texto, model_id):
    prompt = f"""
    Eres un Tutor Psicopedagogo. Resuelve y adecua este examen para el alumno {nombre} ({diag}).
    Devuelve EXCLUSIVAMENTE un JSON con este esquema:
    {{
      "alumno": {{"nombre": "{nombre}", "diagnostico": "{diag}", "grupo": "{grupo}"}},
      "documento": [
        {{
          "tipo": "consigna",
          "enunciado_original": "copia literal del examen",
          "pista": "pista de razonamiento en verde",
          "visual": {{"habilitado": bool, "prompt": "descripcion para dibujo simple"}}
        }}
      ]
    }}
    EXAMEN:
    {examen_texto}
    """
    model = genai.GenerativeModel(model_id)
    def call():
        return model.generate_content(prompt, generation_config=GEN_CFG_JSON)
    
    resp = retry_with_backoff(call)
    data = json.loads(resp.text)
    
    if PYDANTIC_AVAILABLE:
        return AdecuacionFinal.model_validate(data).dict()
    return data

def generar_imagen_ia(model_id, prompt_img):
    try:
        model = genai.GenerativeModel(model_id)
        res = model.generate_content(prompt_img)
        return io.BytesIO(res.candidates[0].content.parts[0].inline_data.data)
    except:
        return None

# ============================================================
# 6. RENDERIZADO DOCX (Sin Regex, basado en JSON)
# ============================================================
def renderizar_adecuacion(data_json, logo_bytes, activar_img, model_img):
    doc = Document()
    
    # Header programático
    table = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(0.9))
    
    p_hdr = table.rows[0].cells[1].paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    al = data_json["alumno"]
    p_hdr.add_run(f"ALUMNO: {al['nombre']}\nAPOYO: {al['diagnostico']}").bold = True

    for item in data_json["documento"]:
        # Transcripción Fiel
        p_orig = doc.add_paragraph(item["enunciado_original"])
        
        # Pista Verde
        p_pista = doc.add_paragraph()
        run = p_pista.add_run(f"💡 {item['pista']}")
        run.font.color.rgb = RGBColor(0, 128, 0)
        run.italic = True
        
        # Apoyo Visual
        if activar_img and item["visual"]["habilitado"]:
            img_data = generar_imagen_ia(model_img, item["visual"]["prompt"])
            if img_data:
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.add_run().add_picture(img_data, width=Inches(2.5))
        
        doc.add_paragraph() # Espaciador

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ============================================================
# 7. INTERFAZ DE USUARIO (Streamlit)
# ============================================================
def main():
    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
        df = pd.read_csv(URL_PLANILLA)
    except Exception as e:
        st.error(f"Error de conexión o secretos: {e}")
        return

    # Sidebar
    with st.sidebar:
        st.header("Configuración")
        m_text = st.text_input("Modelo Texto", MODEL_TEXT_DEFAULT)
        m_img = st.text_input("Modelo Imagen", MODEL_IMAGE_DEFAULT)
        
        grado = st.selectbox("Grado", df.iloc[:, 1].unique())
        df_f = df[df.iloc[:, 1] == grado]
        
        alcance = st.radio("Adecuar para:", ["Todo el grado", "Seleccionar alumnos"])
        alumnos_final = df_f if alcance == "Todo el grado" else df_f[df_f.iloc[:, 2].isin(st.sidebar.multiselect("Alumnos", df_f.iloc[:, 2].unique()))]
        
        st.divider()
        activar_img = st.checkbox("Generar imágenes con IA", True)
        logo = st.file_uploader("Logo Colegio", type=["png", "jpg"])
        logo_b = logo.read() if logo else None

    # Área Principal
    st.title("Motor Pedagógico v13.5")
    archivo = st.file_uploader("Subir Examen (DOCX)", type=["docx"])

    if archivo and st.button("🚀 Iniciar Procesamiento"):
        texto_base = extraer_docx_completo(archivo)
        zip_io = io.BytesIO()
        
        with zipfile.ZipFile(zip_io, 'w') as zf:
            prog = st.progress(0)
            status = st.empty()
            
            for i, (_, fila) in enumerate(alumnos_final.iterrows()):
                nombre = str(fila.iloc[2])
                diag = str(fila.iloc[4])
                grupo = str(fila.iloc[3])
                
                status.text(f"Analizando pedagógicamente a: {nombre}")
                try:
                    data = generar_adecuacion_json(nombre, diag, grupo, texto_base, m_text)
                    docx = renderizar_adecuacion(data, logo_b, activar_img, m_img)
                    zf.writestr(f"Adecuacion_{nombre}.docx", docx)
                except Exception as e:
                    st.sidebar.error(f"Fallo en {nombre}: {e}")
                
                prog.progress((i + 1) / len(alumnos_final))

        st.success("¡Lote completado!")
        st.download_button("Descargar ZIP", zip_io.getvalue(), "adecuaciones.zip")

if __name__ == "__main__":
    main()
