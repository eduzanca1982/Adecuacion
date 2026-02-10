import streamlit as st
import google.generativeai as genai

import pandas as pd
import json
import io
import zipfile
import time
import random
import hashlib
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(page_title="Motor Pedagógico Determinista v13.1", layout="wide")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

# Modelos (configurables)
MODEL_TEXT_DEFAULT = "gemini-1.5-flash"
MODEL_IMAGE_DEFAULT = "imagen-3.0"  # Debe ser configurable: puede no existir en tu cuenta.

# Determinismo (parcial; el modelo puede no ser bit-a-bit determinista)
GEN_CFG_JSON = {
    "response_mime_type": "application/json",
    "temperature": 0,
    "top_p": 1,
    "top_k": 1,
    "max_output_tokens": 4096,
}

# Safety (educación: más restrictivo que lo que tenías)
SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

# Retry
RETRIES = 6
BACKOFF_BASE_SECONDS = 1.0

# Cache (Streamlit)
CACHE_TTL_SECONDS = 6 * 60 * 60  # 6h

# ============================================================
# OPTIONAL: Pydantic validation if available
# ============================================================
PYDANTIC_AVAILABLE = False
try:
    from pydantic import BaseModel, Field, ValidationError

    class VisualModel(BaseModel):
        habilitado: bool = Field(...)
        prompt: Optional[str] = None

    class ItemModel(BaseModel):
        tipo: str = Field(...)
        enunciado_original: str = Field(...)
        pista: str = Field(...)
        visual: VisualModel = Field(...)

    class AlumnoModel(BaseModel):
        nombre: str = Field(...)
        grupo: str = Field(...)
        diagnostico: str = Field(...)

    class AdecuacionModel(BaseModel):
        alumno: AlumnoModel = Field(...)
        documento: List[ItemModel] = Field(...)

    PYDANTIC_AVAILABLE = True
except Exception:
    PYDANTIC_AVAILABLE = False


# ============================================================
# UTILITIES
# ============================================================
def get_content_hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()


def _is_retryable_error(e: Exception) -> bool:
    s = str(e).lower()
    retry_markers = [
        "429",
        "rate",
        "quota",
        "resource exhausted",
        "too many requests",
        "timeout",
        "timed out",
        "deadline",
        "unavailable",
        "503",
        "500",
        "internal",
        "service",
        "temporarily",
        "connection reset",
    ]
    return any(m in s for m in retry_markers)


def retry_with_backoff(fn, retries: int = RETRIES, backoff_in_seconds: float = BACKOFF_BASE_SECONDS):
    last = None
    for attempt in range(retries + 1):
        try:
            return fn()
        except Exception as e:
            last = e
            if attempt >= retries or not _is_retryable_error(e):
                raise
            sleep = (backoff_in_seconds * (2 ** attempt)) + random.uniform(0, 0.75)
            time.sleep(min(sleep, 30))
    raise last


def normalize_bool(x: Any) -> bool:
    if isinstance(x, bool):
        return x
    if isinstance(x, (int, float)):
        return bool(x)
    if isinstance(x, str):
        v = x.strip().lower()
        if v in {"true", "1", "yes", "y", "si", "sí"}:
            return True
        if v in {"false", "0", "no", "n"}:
            return False
    return False


def normalize_visual(v: Any) -> Dict[str, Any]:
    if not isinstance(v, dict):
        return {"habilitado": False}
    habil = normalize_bool(v.get("habilitado", False))
    prompt = v.get("prompt")
    if not isinstance(prompt, str):
        prompt = None
    if not habil:
        return {"habilitado": False}
    return {"habilitado": True, "prompt": prompt or ""}


def basic_schema_validate_and_normalize(data: Any) -> Tuple[Optional[Dict[str, Any]], List[str]]:
    errors: List[str] = []
    if not isinstance(data, dict):
        return None, ["Root debe ser un objeto JSON."]
    if "alumno" not in data or "documento" not in data:
        return None, ["Faltan claves: alumno/documento."]

    alumno = data.get("alumno")
    documento = data.get("documento")
    if not isinstance(alumno, dict):
        errors.append("alumno debe ser objeto.")
        alumno = {}
    if not isinstance(documento, list):
        errors.append("documento debe ser lista.")
        documento = []

    nombre = alumno.get("nombre", "")
    grupo = alumno.get("grupo", "")
    diagnostico = alumno.get("diagnostico", "")

    if not isinstance(nombre, str) or not nombre.strip():
        errors.append("alumno.nombre debe ser string no vacío.")
    if not isinstance(grupo, str) or not grupo.strip():
        errors.append("alumno.grupo debe ser string no vacío.")
    if not isinstance(diagnostico, str):
        errors.append("alumno.diagnostico debe ser string.")

    norm_doc: List[Dict[str, Any]] = []
    for idx, it in enumerate(documento):
        if not isinstance(it, dict):
            errors.append(f"documento[{idx}] debe ser objeto.")
            continue
        tipo = it.get("tipo", "consigna")
        enun = it.get("enunciado_original", "")
        pista = it.get("pista", "")
        visual = normalize_visual(it.get("visual", {"habilitado": False}))

        if not isinstance(tipo, str) or not tipo.strip():
            tipo = "consigna"
        if not isinstance(enun, str) or not enun.strip():
            errors.append(f"documento[{idx}].enunciado_original debe ser string no vacío.")
        if not isinstance(pista, str) or not pista.strip():
            errors.append(f"documento[{idx}].pista debe ser string no vacío.")

        norm_doc.append(
            {
                "tipo": tipo,
                "enunciado_original": enun if isinstance(enun, str) else "",
                "pista": pista if isinstance(pista, str) else "",
                "visual": visual,
            }
        )

    normalized = {
        "alumno": {
            "nombre": str(nombre),
            "grupo": str(grupo),
            "diagnostico": str(diagnostico),
        },
        "documento": norm_doc,
    }
    return normalized, errors


# ============================================================
# DOCX EXTRACTION (paragraphs + tables in body order)
# ============================================================
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _extract_text_from_element(el) -> str:
    # Concatenate all w:t within this element
    parts = []
    for node in el.iter():
        if node.tag == f"{W_NS}t" and node.text:
            parts.append(node.text)
    return "".join(parts).strip()


def extraer_contenido_completo(file) -> str:
    """
    Extrae párrafos y tablas respetando el orden del body.
    Nota: no incluye headers/footers ni imágenes.
    """
    doc = Document(file)
    contenido: List[str] = []

    for element in doc.element.body:
        # Paragraph
        if element.tag == f"{W_NS}p":
            text = _extract_text_from_element(element)
            if text:
                contenido.append(text)
        # Table
        elif element.tag == f"{W_NS}tbl":
            # iterate rows
            for row in element.findall(f".//{W_NS}tr"):
                row_cells = []
                for cell in row.findall(f".//{W_NS}tc"):
                    cell_text = _extract_text_from_element(cell)
                    row_cells.append(cell_text)
                # Use TAB separator to keep reading simple; pipe also ok
                line = "\t".join([c for c in row_cells if c is not None])
                if line.strip():
                    contenido.append(line.strip())
            # spacer between tables
            contenido.append("")

    return "\n".join([c for c in contenido if c is not None]).strip()


# ============================================================
# GEMINI JSON GENERATION + REPAIR
# ============================================================
BASE_JSON_INSTRUCTIONS = """
Eres un experto en adecuación pedagógica. Debes devolver EXCLUSIVAMENTE un JSON válido (sin markdown, sin texto extra, sin comentarios).

Esquema obligatorio:
{
  "alumno": { "nombre": "string", "grupo": "string", "diagnostico": "string" },
  "documento": [
    {
      "tipo": "consigna",
      "enunciado_original": "texto literal completo de la consigna",
      "pista": "pista pedagógica breve para guiar razonamiento (no dar respuesta)",
      "visual": { "habilitado": boolean, "prompt": "string opcional" }
    }
  ]
}

Reglas:
1) Incluye TODAS las consignas del examen original SIN OMITIR ninguna.
2) enunciado_original debe ser COPIA FIEL del examen (no parafrasear).
3) Ajusta la pista por diagnóstico y grupo. No entregues la solución.
4) visual.habilitado:
   - true solo si realmente ayuda a razonar.
   - Si true, visual.prompt debe empezar EXACTAMENTE con:
     "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: "
     y luego el objeto/escena breve.
   - Si false, no incluyas prompt o déjalo vacío.
5) Prohibido: saludos, análisis, explicaciones al docente, texto fuera del JSON.
""".strip()


def build_prompt(nombre: str, diagnostico: str, grupo: str, examen_texto: str) -> str:
    # Evitar variables no deterministas (ej: timestamps)
    return f"""{BASE_JSON_INSTRUCTIONS}

Alumno:
- nombre: {nombre}
- grupo: {grupo}
- diagnostico: {diagnostico}

EXAMEN ORIGINAL (copia literal):
{examen_texto}
""".strip()


def build_repair_prompt(bad_text: str, errors: List[str]) -> str:
    err_txt = "\n".join([f"- {e}" for e in errors]) if errors else "- JSON inválido o no parseable."
    return f"""
Devuelve EXCLUSIVAMENTE un JSON válido y corregido (sin texto extra).

Errores detectados:
{err_txt}

JSON A CORREGIR:
{bad_text}

Recuerda: debe cumplir EXACTAMENTE el esquema y reglas. No agregues nada fuera del JSON.
""".strip()


@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def _cached_gemini_json(cache_key: str, model_text: str, prompt: str) -> Dict[str, Any]:
    # cache_key exists only to allow caching; do not remove.
    model = genai.GenerativeModel(model_text)

    def _call():
        return model.generate_content(
            prompt,
            generation_config=GEN_CFG_JSON,
            safety_settings=SAFETY_SETTINGS,
        )

    resp = retry_with_backoff(_call)
    text = getattr(resp, "text", "") or ""
    data = json.loads(text)

    # Validate (pydantic if possible, else basic)
    if PYDANTIC_AVAILABLE:
        try:
            AdecuacionModel.model_validate(data)
            return data
        except Exception as ve:
            # Let caller do repair; raise with context
            raise ValueError(f"Schema inválido (pydantic): {ve}")
    else:
        normalized, errs = basic_schema_validate_and_normalize(data)
        if errs:
            raise ValueError(f"Schema inválido: {errs[:6]}")
        return normalized  # normalized is Dict

    return data


def solicitar_adecuacion_json(
    nombre: str,
    diagnostico: str,
    grupo: str,
    examen_texto: str,
    exam_hash: str,
    model_text: str,
) -> Dict[str, Any]:
    prompt = build_prompt(nombre, diagnostico, grupo, examen_texto)
    cache_key = f"{exam_hash}::{nombre}::{grupo}::{diagnostico}::{model_text}"
    try:
        data = _cached_gemini_json(cache_key, model_text, prompt)
        return data
    except Exception as e:
        # Attempt repair once
        # We need the raw response text to repair; so we call once without cache to capture it.
        model = genai.GenerativeModel(model_text)

        def _call_raw():
            return model.generate_content(
                prompt,
                generation_config=GEN_CFG_JSON,
                safety_settings=SAFETY_SETTINGS,
            )

        resp = retry_with_backoff(_call_raw)
        raw_text = getattr(resp, "text", "") or ""

        parse_errors: List[str] = []
        parsed = None
        try:
            parsed = json.loads(raw_text)
        except Exception as je:
            parse_errors.append(f"No se pudo parsear JSON: {je}")

        if parsed is not None:
            if PYDANTIC_AVAILABLE:
                try:
                    AdecuacionModel.model_validate(parsed)
                    return parsed
                except Exception as ve:
                    parse_errors.append(str(ve))
            else:
                normalized, errs = basic_schema_validate_and_normalize(parsed)
                if errs:
                    parse_errors.extend(errs)
                else:
                    return normalized

        # Repair prompt
        repair_prompt = build_repair_prompt(raw_text, parse_errors)

        def _call_repair():
            return model.generate_content(
                repair_prompt,
                generation_config=GEN_CFG_JSON,
                safety_settings=SAFETY_SETTINGS,
            )

        repaired = retry_with_backoff(_call_repair)
        repaired_text = getattr(repaired, "text", "") or ""

        data2 = json.loads(repaired_text)
        if PYDANTIC_AVAILABLE:
            AdecuacionModel.model_validate(data2)
            return data2

        normalized2, errs2 = basic_schema_validate_and_normalize(data2)
        if errs2:
            raise ValueError(f"JSON reparado inválido: {errs2[:8]} (original error: {e})")
        return normalized2


# ============================================================
# IMAGE GENERATION (best effort, configurable, non-deterministic)
# ============================================================
def generar_imagen_ia(model_image: str, prompt_visual: str) -> Optional[io.BytesIO]:
    model = genai.GenerativeModel(model_image)

    def _call():
        return model.generate_content(
            prompt_visual,
            # Si tu endpoint soporta safety_settings para imagen, dejalo.
            # Si no, puede fallar. Lo mantenemos por seguridad; si falla, se atrapa.
            safety_settings=SAFETY_SETTINGS,
        )

    try:
        res = retry_with_backoff(_call)
        # Defensive parsing
        cand = res.candidates[0]
        part0 = cand.content.parts[0]
        inline = getattr(part0, "inline_data", None)
        if not inline or not getattr(inline, "data", None):
            return None
        return io.BytesIO(inline.data)
    except Exception:
        return None


# ============================================================
# DOCX RENDER (deterministic from JSON)
# ============================================================
def renderizar_docx(data_json: Dict[str, Any], logo_bytes: Optional[bytes], gen_img_bool: bool, model_image: str) -> bytes:
    doc = Document()

    # Header table
    header_table = doc.add_table(rows=1, cols=2)

    if logo_bytes:
        try:
            cell_logo = header_table.rows[0].cells[0]
            cell_logo.paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(0.85))
        except Exception:
            pass

    cell_info = header_table.rows[0].cells[1]
    p_info = cell_info.paragraphs[0]
    p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    al = data_json.get("alumno", {})
    nombre = str(al.get("nombre", "")).strip()
    grupo = str(al.get("grupo", "")).strip()
    diag = str(al.get("diagnostico", "")).strip()

    run = p_info.add_run(f"ALUMNO: {nombre}\nGRUPO: {grupo} | APOYO: {diag}")
    run.bold = True
    run.font.size = Pt(10)

    # Body
    green = RGBColor(0, 128, 0)
    items = data_json.get("documento", []) if isinstance(data_json.get("documento"), list) else []

    for item in items:
        if not isinstance(item, dict):
            continue

        enun = str(item.get("enunciado_original", "")).strip()
        pista = str(item.get("pista", "")).strip()
        visual = normalize_visual(item.get("visual", {"habilitado": False}))

        if not enun:
            continue

        # Enunciado
        p_orig = doc.add_paragraph()
        r1 = p_orig.add_run(enun)
        r1.font.size = Pt(11)
        p_orig.paragraph_format.space_after = Pt(3)

        # Pista
        if pista:
            p_pista = doc.add_paragraph()
            r2 = p_pista.add_run(f"💡 {pista}")
            r2.font.color.rgb = green
            r2.italic = True
            r2.font.size = Pt(10)
            p_pista.paragraph_format.space_after = Pt(4)

        # Imagen
        if gen_img_bool and visual.get("habilitado", False):
            prompt_visual = str(visual.get("prompt", "") or "").strip()
            # Forzar prefijo requerido si vino vacío o mal
            required_prefix = "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: "
            if prompt_visual and not prompt_visual.startswith(required_prefix):
                prompt_visual = required_prefix + prompt_visual

            if prompt_visual:
                img_data = generar_imagen_ia(model_image, prompt_visual)
                if img_data:
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    try:
                        p_img.add_run().add_picture(img_data, width=Inches(2.6))
                    except Exception:
                        pass
                    p_img.paragraph_format.space_after = Pt(6)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# STREAMLIT APP
# ============================================================
def main():
    st.title("Motor Pedagógico Determinista v13.1")

    # API key required
    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    except Exception:
        st.error("Falta GOOGLE_API_KEY en st.secrets.")
        return

    # Load CSV
    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error leyendo planilla CSV: {e}")
        return

    # Sidebar config
    with st.sidebar:
        st.header("⚙️ Configuración")

        # Model config
        model_text = st.text_input("Modelo texto", value=MODEL_TEXT_DEFAULT)
        model_image = st.text_input("Modelo imagen", value=MODEL_IMAGE_DEFAULT)

        # Filters
        grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
        grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
        diag_col = df.columns[4] if len(df.columns) > 4 else df.columns[0]

        st.caption("Filtros planilla")
        grado = st.selectbox("Grado", sorted(df[grado_col].dropna().unique().tolist()))
        df_f = df[df[grado_col] == grado].copy()

        alcance = st.radio("¿A quiénes adecuar?", ["Todos", "Seleccionar"], horizontal=True)
        if alcance == "Seleccionar":
            alumnos_sel = st.multiselect("Alumnos", df_f[alumno_col].dropna().unique().tolist())
            if alumnos_sel:
                df_f = df_f[df_f[alumno_col].isin(alumnos_sel)]
        else:
            alumnos_sel = []

        st.divider()

        gen_img = st.checkbox("Generar imágenes IA", value=True)
        logo = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
        logo_bytes = logo.read() if logo else None

        st.divider()
        st.caption(f"Pydantic: {'ON' if PYDANTIC_AVAILABLE else 'OFF'}")

    file_base = st.file_uploader("Examen base (DOCX)", type=["docx"])

    if not file_base:
        return

    if st.button("Procesar lote"):
        # Extract exam
        with st.spinner("Extrayendo contenido del DOCX..."):
            exam_text = extraer_contenido_completo(file_base)
        if not exam_text.strip():
            st.error("No se pudo extraer texto del DOCX (vacío).")
            return

        exam_hash = get_content_hash(exam_text)

        # Process
        alumnos_data = df_f[[alumno_col, grupo_col, diag_col]].dropna(subset=[alumno_col]).copy()
        total = len(alumnos_data)
        if total == 0:
            st.error("No hay alumnos para procesar con los filtros actuales.")
            return

        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
            bar = st.progress(0.0)
            status = st.empty()

            for i, (_, row) in enumerate(alumnos_data.iterrows(), start=1):
                nombre = str(row[alumno_col]).strip()
                grupo_val = str(row[grupo_col]).strip()
                diag_val = str(row[diag_col]).strip()

                status.text(f"({i}/{total}) Generando JSON para: {nombre}")

                try:
                    data_json = solicitar_adecuacion_json(
                        nombre=nombre,
                        diagnostico=diag_val,
                        grupo=grupo_val,
                        examen_texto=exam_text,
                        exam_hash=exam_hash,
                        model_text=model_text,
                    )

                    # Extra normalize (if pydantic not present, already normalized; if pydantic present, keep but normalize visual defensively)
                    if not PYDANTIC_AVAILABLE:
                        normalized, _ = basic_schema_validate_and_normalize(data_json)
                        data_json = normalized or data_json

                    status.text(f"({i}/{total}) Renderizando DOCX: {nombre}")
                    docx_bytes = renderizar_docx(data_json, logo_bytes, gen_img, model_image)

                    safe_name = nombre.replace(" ", "_").replace("/", "_").replace("\\", "_")
                    zip_file.writestr(f"Adecuacion_{safe_name}.docx", docx_bytes)

                except Exception as e:
                    # Continue batch; embed an error file for traceability
                    err_txt = f"Alumno: {nombre}\nGrupo: {grupo_val}\nDiagnóstico: {diag_val}\nError: {e}\n"
                    zip_file.writestr(f"ERROR_{nombre.replace(' ', '_')}.txt", err_txt.encode("utf-8"))

                bar.progress(i / total)

        st.success("Lote completado.")
        st.download_button(
            "Descargar ZIP",
            data=zip_io.getvalue(),
            file_name="adecuaciones.zip",
            mime="application/zip",
        )


if __name__ == "__main__":
    main()
