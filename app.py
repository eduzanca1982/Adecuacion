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
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============================================================
# Motor Pedagógico v17.0 (Opal++ Ultra)
# SUPERIOR a la versión de Gemini en:
# - Dos modos de entrada: (A) DOCX a adaptar (B) Prompt para generar desde cero
# - Diagnóstico de parseo (DOCX + JSON) antes del lote
# - Selección automática robusta de modelos (texto + imagen) con ranking y smoke test
# - Manejo de finish_reason=2 (MAX_TOKENS) y respuestas sin parts
# - JSON estricto + validación dura + reparación 1 vez + fallback compacto
# - Detección/normalización de visual.prompt + validación de bytes de imagen
# - Render DOCX “dyslexia-friendly” (Verdana 14 + interlineado + bloques cortos)
# - ZIP blindado: _REPORTE.txt, _RESUMEN.txt, ERROR_*.txt, _META_*.txt
# - Cache por examen/brief/alumno para ahorrar tokens y estabilizar resultados
# ============================================================

st.set_page_config(page_title="Motor Pedagógico v17.0 (Opal++)", layout="wide")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

RETRIES = 6
CACHE_TTL_SECONDS = 6 * 60 * 60
SMOKE_IMAGE_MIN_BYTES = 500

IMAGE_PROMPT_PREFIX = "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: "

SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

BASE_GEN_CFG_JSON = {
    "response_mime_type": "application/json",
    "temperature": 0,
    "top_p": 1,
    "top_k": 1,
}

# Escalado tokens: FULL primero, luego COMPACT
OUT_TOKEN_STEPS_FULL = [4096, 6144, 8192]
OUT_TOKEN_STEPS_COMPACT = [2048, 4096]

# ============================================================
# Pydantic opcional (si está instalado)
# ============================================================
PYDANTIC_AVAILABLE = False
try:
    from pydantic import BaseModel, Field

    class VisualModel(BaseModel):
        habilitado: bool = Field(...)
        prompt: Optional[str] = None

    class ItemModel(BaseModel):
        tipo: str = Field(...)
        enunciado: str = Field(...)
        opciones: List[str] = Field(default_factory=list)

    class AlumnoModel(BaseModel):
        nombre: str = Field(...)
        grupo: str = Field(...)
        diagnostico: str = Field(...)
        grado: str = Field(...)

    class ActividadModel(BaseModel):
        alumno: AlumnoModel = Field(...)
        contexto: Dict[str, Any] = Field(...)
        objetivo_aprendizaje: str = Field(...)
        consigna_adaptada: str = Field(...)
        items: List[ItemModel] = Field(...)
        adecuaciones_aplicadas: List[str] = Field(default_factory=list)
        sugerencias_docente: List[str] = Field(default_factory=list)
        visual: VisualModel = Field(...)
        control_calidad: Dict[str, Any] = Field(...)

    PYDANTIC_AVAILABLE = True
except Exception:
    PYDANTIC_AVAILABLE = False

# ============================================================
# Prompt Opal++ (salida JSON determinista)
# - Soporta "ADAPTAR" (con actividad base) o "CREAR" (solo brief)
# - Incluye control_calidad: conteos y garantías para auditar
# ============================================================
SYSTEM_PROMPT_OPALPP = f"""
Actúa como un Asistente Pedagógico Experto en Inclusión y Dislexia.

TU OBJETIVO:
- Si MODO=ADAPTAR: transformar una actividad original en una versión accesible e inclusiva.
- Si MODO=CREAR: crear una actividad inclusiva desde cero en base a un brief.

GUÍAS ESTRICTAS (Dyslexia-friendly):
- Tipografía recomendada: Sans Serif (Arial, Verdana, Open Sans). (El DOCX usará Verdana 14)
- Tamaño 14, interlineado 1.5 a 2.0
- Texto oscuro sobre fondo claro
- Bloques cortos (1 idea por bloque), listas numeradas/viñetas
- **Negrita** solo para palabras clave; evitar itálicas/subrayado (en la salida textual)
- Consignas claras: 1 acción por frase, pasos secuenciales
- Vocabulario concreto; evitar metáforas/ambigüedades
- Siempre dar un ejemplo breve cuando sea útil
- Reducir ítems para evitar fatiga cognitiva (prioriza comprensión)
- No penalizar ortografía si el objetivo no es ortografía
- Sugerir tiempo extra y lectura acompañada

SALIDA:
Devuelve EXCLUSIVAMENTE un JSON válido (sin markdown, sin texto extra).

ESQUEMA EXACTO:
{{
  "alumno": {{
    "nombre": "string",
    "grupo": "string",
    "diagnostico": "string",
    "grado": "string"
  }},
  "contexto": {{
    "modo": "ADAPTAR|CREAR",
    "materia": "string",
    "nivel": "string",
    "tema": "string",
    "estilo_extra": "string"
  }},
  "objetivo_aprendizaje": "string",
  "consigna_adaptada": "string",
  "items": [
    {{
      "tipo": "multiple choice|unir|completar|verdadero_falso|problema_guiado",
      "enunciado": "string",
      "opciones": ["string", "string"]
    }}
  ],
  "adecuaciones_aplicadas": ["string", "string"],
  "sugerencias_docente": ["string", "string"],
  "visual": {{
    "habilitado": boolean,
    "prompt": "string"
  }},
  "control_calidad": {{
    "items_count": number,
    "incluye_ejemplo": boolean,
    "lenguaje_concreto": boolean,
    "una_accion_por_frase": boolean
  }}
}}

REGLAS DURAS:
1) JSON puro, nada fuera del JSON.
2) items_count debe coincidir con len(items).
3) visual.prompt SOLO si visual.habilitado=true y debe empezar EXACTAMENTE con:
   "{IMAGE_PROMPT_PREFIX}"
4) Si MODO=ADAPTAR: preservar la intención pedagógica del original; NO copiar texto basura; pero sí mantener lo esencial.
5) Si MODO=CREAR: construir desde cero basado en el brief.
""".strip()

def build_prompt_opalpp(
    modo: str,
    materia: str,
    nivel: str,
    tema: str,
    estilo_extra: str,
    alumno_nombre: str,
    alumno_grupo: str,
    alumno_diag: str,
    alumno_grado: str,
    original_text: str
) -> str:
    payload = {
        "MODO": modo,
        "MATERIA": materia,
        "NIVEL": nivel,
        "TEMA": tema,
        "ESTILO_EXTRA": estilo_extra or "",
        "ALUMNO": {
            "nombre": alumno_nombre,
            "grupo": alumno_grupo,
            "diagnostico": alumno_diag,
            "grado": alumno_grado
        },
        "ORIGINAL": original_text or ""
    }
    # JSON-in-prompt reduce ambigüedad
    return f"{SYSTEM_PROMPT_OPALPP}\n\nENTRADA:\n{json.dumps(payload, ensure_ascii=False, indent=2)}"

def build_prompt_compact_opalpp(
    modo: str,
    materia: str,
    nivel: str,
    tema: str,
    estilo_extra: str,
    alumno_nombre: str,
    alumno_grupo: str,
    alumno_diag: str,
    alumno_grado: str,
    original_text: str
) -> str:
    # Fallback compacto: menos ítems, sin visual, pistas muy cortas
    payload = {
        "MODO": modo,
        "MATERIA": materia,
        "NIVEL": nivel,
        "TEMA": tema,
        "ESTILO_EXTRA": estilo_extra or "",
        "ALUMNO": {
            "nombre": alumno_nombre,
            "grupo": alumno_grupo,
            "diagnostico": alumno_diag,
            "grado": alumno_grado
        },
        "ORIGINAL": (original_text or "")[:3000]
    }
    return f"""
Devuelve SOLO JSON válido.
Genera max 6 items. Enunciados y consignas breves. visual siempre false.

ESQUEMA:
{{
  "alumno": {{"nombre":"{alumno_nombre}","grupo":"{alumno_grupo}","diagnostico":"{alumno_diag}","grado":"{alumno_grado}"}},
  "contexto": {{"modo":"{modo}","materia":"{materia}","nivel":"{nivel}","tema":"{tema}","estilo_extra":"{estilo_extra or ''}"}},
  "objetivo_aprendizaje":"string",
  "consigna_adaptada":"string",
  "items":[{{"tipo":"multiple choice","enunciado":"string","opciones":["A","B"]}}],
  "adecuaciones_aplicadas":["string"],
  "sugerencias_docente":["string"],
  "visual":{{"habilitado":false,"prompt":""}},
  "control_calidad":{{"items_count":0,"incluye_ejemplo":true,"lenguaje_concreto":true,"una_accion_por_frase":true}}
}}

ENTRADA:
{json.dumps(payload, ensure_ascii=False)}
""".strip()

def build_repair_prompt(bad: str, why: str) -> str:
    return f"""
Devuelve EXCLUSIVAMENTE un JSON válido y corregido (sin texto extra).

Problema detectado:
{why}

JSON A CORREGIR:
{bad}

Reglas:
- Debe cumplir EXACTAMENTE el esquema y reglas del sistema
- items_count debe coincidir con len(items)
- No agregues texto fuera del JSON
""".strip()

# ============================================================
# Utilidades base
# ============================================================
def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def hash_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()

def safe_filename(name: str) -> str:
    s = str(name).strip().replace(" ", "_")
    for ch in ["/", "\\", ":", "*", "?", "\"", "<", ">", "|"]:
        s = s.replace(ch, "_")
    while "__" in s:
        s = s.replace("__", "_")
    return (s or "SIN_NOMBRE")[:120]

def _is_retryable_error(e: Exception) -> bool:
    s = str(e).lower()
    markers = [
        "429", "too many requests", "rate", "quota", "resource exhausted",
        "timeout", "timed out", "deadline", "unavailable", "503", "500", "internal",
        "connection reset", "temporarily"
    ]
    return any(m in s for m in markers)

def retry_with_backoff(fn):
    last = None
    for i in range(RETRIES):
        try:
            return fn()
        except Exception as e:
            last = e
            if i == RETRIES - 1 or not _is_retryable_error(e):
                raise
            sleep = (2 ** i) + random.uniform(0, 0.75)
            time.sleep(min(sleep, 30))
    raise last

def normalize_bool(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    if isinstance(v, str):
        return v.strip().lower() in {"true", "1", "yes", "y", "si", "sí"}
    return False

def normalize_visual(v: Any) -> Dict[str, Any]:
    if not isinstance(v, dict):
        return {"habilitado": False, "prompt": ""}
    return {"habilitado": normalize_bool(v.get("habilitado", False)),
            "prompt": str(v.get("prompt", "")).strip()}

# ============================================================
# DOCX extraction (párrafos + tablas)
# ============================================================
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def _extract_text(el) -> str:
    return "".join(n.text for n in el.iter() if n.tag == f"{W_NS}t" and n.text).strip()

def extraer_texto_docx(file) -> str:
    doc = Document(file)
    out: List[str] = []
    for el in doc.element.body:
        if el.tag == f"{W_NS}p":
            t = _extract_text(el)
            if t:
                out.append(t)
        elif el.tag == f"{W_NS}tbl":
            for row in el.findall(f".//{W_NS}tr"):
                cells = [_extract_text(c) for c in row.findall(f".//{W_NS}tc")]
                if any(cells):
                    out.append(" | ".join(cells))
            out.append("")
    return "\n".join(out).strip()

def validate_exam_text(exam_text: str) -> Tuple[bool, str, Dict[str, Any]]:
    info = {
        "chars": len(exam_text or ""),
        "lines": (exam_text or "").count("\n") + (1 if exam_text else 0),
        "pipes": (exam_text or "").count("|"),
        "preview": (exam_text or "")[:1400],
    }
    if not exam_text or not exam_text.strip():
        return False, "TEXTO vacío tras extracción (posible actividad en imágenes/cuadros).", info
    if len(exam_text) < 150:
        return False, "TEXTO muy corto (<150 chars). Posible doc con imágenes/shapes.", info
    return True, "OK", info

# ============================================================
# Model scanning + selection (texto + imagen)
# ============================================================
def list_models_generate_content() -> List[str]:
    ms = genai.list_models()
    out = []
    for m in ms:
        methods = getattr(m, "supported_generation_methods", []) or []
        if "generateContent" in methods:
            out.append(m.name)
    return out

def _rank_text_model(name: str) -> Tuple[int, int]:
    n = name.lower()
    penalty = 0
    for bad in ["tts", "embedding", "embed", "audio", "rerank"]:
        if bad in n:
            penalty += 100

    if "gemini-2.5-pro" in n:
        tier = 0
    elif "gemini-2.5-flash" in n:
        tier = 1
    elif "gemini-2.0-pro" in n:
        tier = 2
    elif "gemini-2.0-flash" in n:
        tier = 3
    elif "gemini-1.5-pro" in n:
        tier = 4
    elif "gemini-1.5-flash" in n:
        tier = 5
    elif "gemini" in n and "pro" in n:
        tier = 6
    elif "gemini" in n and "flash" in n:
        tier = 7
    elif "gemini" in n:
        tier = 10
    elif "mini" in n and "pro" in n:
        tier = 12
    elif "mini" in n and "flash" in n:
        tier = 13
    elif "mini" in n:
        tier = 15
    else:
        tier = 30

    exp_penalty = 1 if any(x in n for x in ["exp", "preview", "experimental"]) else 0
    return (tier + penalty + exp_penalty, len(n))

def pick_best_text_model(models: List[str]) -> Optional[str]:
    if not models:
        return None
    return sorted(models, key=_rank_text_model)[0]

def _candidate_image_models(models: List[str]) -> List[str]:
    cands = []
    for m in models:
        n = m.lower()
        if "image" in n or "image-gen" in n or "imagen" in n:
            cands.append(m)

    def score(x: str) -> Tuple[int, int]:
        nx = x.lower()
        exp = 1 if any(t in nx for t in ["exp", "preview", "experimental"]) else 0
        return (exp, len(nx))
    return sorted(cands, key=score)

def _extract_inline_bytes_or_none(resp) -> Optional[bytes]:
    try:
        cand = resp.candidates[0]
        content = getattr(cand, "content", None)
        if not content or not getattr(content, "parts", None):
            return None
        part0 = content.parts[0]
        inline = getattr(part0, "inline_data", None)
        data = getattr(inline, "data", None) if inline else None
        return data
    except Exception:
        return None

def smoke_test_image_model(model_id: str) -> Tuple[bool, str]:
    prompt = f"{IMAGE_PROMPT_PREFIX} manzana"
    try:
        m = genai.GenerativeModel(model_id)
        r = retry_with_backoff(lambda: m.generate_content(prompt, safety_settings=SAFETY_SETTINGS))
        data = _extract_inline_bytes_or_none(r)
        if not data:
            return False, "Respuesta sin inline_data.data"
        if len(data) < SMOKE_IMAGE_MIN_BYTES:
            return False, f"inline_data muy chico ({len(data)} bytes)"
        return True, f"OK bytes={len(data)}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

def pick_best_image_model(models: List[str]) -> Tuple[Optional[str], str]:
    cands = _candidate_image_models(models)
    if not cands:
        return None, "No se detectaron candidatos de imagen en list_models()."
    for mid in cands[:10]:
        ok, msg = smoke_test_image_model(mid)
        if ok:
            return mid, f"Seleccionado por smoke test: {msg}"
    return None, "Se detectaron candidatos, pero ninguno pasó el smoke test."

@st.cache_resource(show_spinner=False)
def forced_boot_scan() -> Dict[str, Any]:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    models = list_models_generate_content()
    if not models:
        raise RuntimeError("No se detectaron modelos con generateContent. API key inválida/limitada o endpoint incompatible.")
    text_model = pick_best_text_model(models)
    if not text_model:
        raise RuntimeError("No se pudo seleccionar modelo de texto automáticamente.")
    image_model, image_reason = pick_best_image_model(models)
    return {
        "models": models,
        "text_model": text_model,
        "image_model": image_model,
        "image_reason": image_reason,
        "boot_time": now_str(),
    }

def boot_or_stop() -> Dict[str, Any]:
    try:
        return forced_boot_scan()
    except Exception as e:
        st.error("Fallo en arranque (escaneo forzado de modelos). La app no continuará.")
        st.code(f"{type(e).__name__}: {e}")
        st.stop()

BOOT = boot_or_stop()

# ============================================================
# Respuesta Gemini robusta (sin depender de response.text)
# ============================================================
def _extract_text_or_none(resp) -> Optional[str]:
    try:
        cand = resp.candidates[0]
        content = getattr(cand, "content", None)
        if not content or not getattr(content, "parts", None):
            return None
        chunks = []
        for p in content.parts:
            t = getattr(p, "text", None)
            if t:
                chunks.append(t)
        text = "".join(chunks).strip()
        return text if text else None
    except Exception:
        return None

def _finish_reason(resp) -> Optional[int]:
    try:
        return int(resp.candidates[0].finish_reason)
    except Exception:
        return None

def validate_activity_json(data: Dict[str, Any]) -> Tuple[bool, str]:
    try:
        if PYDANTIC_AVAILABLE:
            Activity = ActividadModel.model_validate(data)
            # items_count must match
            if Activity.control_calidad.get("items_count") != len(Activity.items):
                return False, "control_calidad.items_count no coincide con len(items)"
            # visual prompt rule
            if Activity.visual.habilitado and Activity.visual.prompt and not Activity.visual.prompt.startswith(IMAGE_PROMPT_PREFIX):
                return False, "visual.prompt no respeta el prefijo requerido"
            return True, "OK(pydantic)"
        # Basic validation
        if not isinstance(data, dict):
            return False, "Root no es objeto"
        for k in ["alumno", "contexto", "objetivo_aprendizaje", "consigna_adaptada", "items", "adecuaciones_aplicadas", "sugerencias_docente", "visual", "control_calidad"]:
            if k not in data:
                return False, f"Falta clave: {k}"
        if not isinstance(data["items"], list) or len(data["items"]) < 1:
            return False, "items vacío/no lista"
        cc = data.get("control_calidad", {})
        if cc.get("items_count") != len(data["items"]):
            return False, "control_calidad.items_count no coincide con len(items)"
        v = data.get("visual", {})
        if normalize_bool(v.get("habilitado", False)):
            p = str(v.get("prompt", "")).strip()
            if p and not p.startswith(IMAGE_PROMPT_PREFIX):
                return False, "visual.prompt no respeta el prefijo requerido"
        return True, "OK(basic)"
    except Exception as e:
        return False, f"Exception validando: {e}"

def generate_json_once(model_id: str, prompt: str, max_out: int) -> Dict[str, Any]:
    m = genai.GenerativeModel(model_id)
    cfg = dict(BASE_GEN_CFG_JSON)
    cfg["max_output_tokens"] = max_out
    resp = retry_with_backoff(lambda: m.generate_content(prompt, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
    text = _extract_text_or_none(resp)
    if text is None:
        fr = _finish_reason(resp)
        raise ValueError(f"Empty candidate (finish_reason={fr})")
    return json.loads(text)

def generate_json_with_repair(model_id: str, prompt: str, max_out: int) -> Dict[str, Any]:
    try:
        data = generate_json_once(model_id, prompt, max_out)
        ok, why = validate_activity_json(data)
        if ok:
            return data
        raise ValueError(f"JSON inválido: {why}")
    except Exception as e:
        m = genai.GenerativeModel(model_id)
        cfg = dict(BASE_GEN_CFG_JSON)
        cfg["max_output_tokens"] = max_out

        resp1 = retry_with_backoff(lambda: m.generate_content(prompt, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
        raw = _extract_text_or_none(resp1)
        fr = _finish_reason(resp1)
        if raw is None:
            raise ValueError(f"Empty candidate (finish_reason={fr})")

        why = f"{type(e).__name__}: {e}"
        repair_prompt = build_repair_prompt(raw, why)
        resp2 = retry_with_backoff(lambda: m.generate_content(repair_prompt, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
        raw2 = _extract_text_or_none(resp2)
        fr2 = _finish_reason(resp2)
        if raw2 is None:
            raise ValueError(f"Empty candidate after repair (finish_reason={fr2})")
        data2 = json.loads(raw2)
        ok2, why2 = validate_activity_json(data2)
        if not ok2:
            raise ValueError(f"JSON reparado inválido: {why2}")
        return data2

@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def cached_generate_activity(cache_key: str, model_id: str, prompt: str, max_out: int) -> Dict[str, Any]:
    data = generate_json_with_repair(model_id, prompt, max_out)
    return data

def request_activity_ultra(model_id: str, prompt_full: str, prompt_compact: str, cache_key: str) -> Tuple[Dict[str, Any], str, int]:
    last_err = None
    for t in OUT_TOKEN_STEPS_FULL:
        try:
            data = cached_generate_activity(cache_key + f"::FULL::{t}", model_id, prompt_full, t)
            return data, "FULL", t
        except Exception as e:
            last_err = e
            continue
    for t in OUT_TOKEN_STEPS_COMPACT:
        try:
            data = cached_generate_activity(cache_key + f"::COMPACT::{t}", model_id, prompt_compact, t)
            return data, "COMPACT", t
        except Exception as e:
            last_err = e
            continue
    raise last_err if last_err else RuntimeError("Fallo desconocido generando actividad")

# ============================================================
# Imagen (best effort + validación bytes)
# ============================================================
def generar_imagen_ia(model_id: str, prompt_img: str) -> Optional[io.BytesIO]:
    try:
        m = genai.GenerativeModel(model_id)
        resp = retry_with_backoff(lambda: m.generate_content(prompt_img, safety_settings=SAFETY_SETTINGS))
        data = _extract_inline_bytes_or_none(resp)
        if not data or len(data) < SMOKE_IMAGE_MIN_BYTES:
            return None
        return io.BytesIO(data)
    except Exception:
        return None

# ============================================================
# Render DOCX (Opal++): layout dyslexia-friendly
# - Verdana 14, line spacing 1.8
# - Secciones: Objetivo, Consigna, Actividad, Adecuaciones, Sugerencias
# - Items con checkboxes visuales
# ============================================================
def _add_heading(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.2

def _add_block(doc: Document, text: str, line_spacing: float = 1.8):
    p = doc.add_paragraph(text)
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_after = Pt(6)
    return p

def render_docx_opalpp(
    data: Dict[str, Any],
    logo_bytes: Optional[bytes],
    activar_img: bool,
    model_img_id: Optional[str]
) -> bytes:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Verdana"
    style.font.size = Pt(14)

    # Header
    header = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try:
            header.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(0.85))
        except Exception:
            pass

    info = header.rows[0].cells[1].paragraphs[0]
    info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    al = data.get("alumno", {})
    ctx = data.get("contexto", {})
    info_run = info.add_run(
        f"ALUMNO: {al.get('nombre','')}\n"
        f"GRADO: {al.get('grado','')} | GRUPO: {al.get('grupo','')}\n"
        f"APOYO: {al.get('diagnostico','')}\n"
        f"MATERIA: {ctx.get('materia','')} | TEMA: {ctx.get('tema','')}"
    )
    info_run.bold = True

    # Secciones
    _add_heading(doc, "Objetivo de aprendizaje")
    _add_block(doc, str(data.get("objetivo_aprendizaje", "")).strip())

    _add_heading(doc, "Consigna adaptada")
    _add_block(doc, str(data.get("consigna_adaptada", "")).strip())

    _add_heading(doc, "Actividad / Ítems")
    items = data.get("items", [])
    for idx, it in enumerate(items, start=1):
        tipo = str(it.get("tipo", "")).strip()
        enun = str(it.get("enunciado", "")).strip()
        opts = it.get("opciones", []) if isinstance(it.get("opciones", []), list) else []

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.8
        r1 = p.add_run(f"{idx}. ")
        r1.bold = True
        r2 = p.add_run(enun)
        r2.bold = True  # keyword emphasis; enunciado completo en negrita ayuda foco en dislexia

        p2 = doc.add_paragraph(f"Tipo: {tipo}")
        p2.paragraph_format.line_spacing = 1.6

        if opts:
            for op in opts[:10]:
                po = doc.add_paragraph(f"☐ {op}")
                po.paragraph_format.line_spacing = 1.6
        else:
            # Línea de respuesta
            _add_block(doc, "Respuesta: ________________________________", line_spacing=1.6)

        doc.add_paragraph("")

    _add_heading(doc, "Adecuaciones aplicadas")
    for a in (data.get("adecuaciones_aplicadas", []) or [])[:20]:
        pa = doc.add_paragraph(f"• {a}")
        pa.paragraph_format.line_spacing = 1.6

    _add_heading(doc, "Sugerencias para el docente")
    for s in (data.get("sugerencias_docente", []) or [])[:20]:
        ps = doc.add_paragraph(f"• {s}")
        ps.paragraph_format.line_spacing = 1.6

    # Imagen
    v = normalize_visual(data.get("visual", {}))
    if activar_img and model_img_id and v.get("habilitado") and v.get("prompt"):
        pv = str(v.get("prompt", "")).strip()
        if pv and not pv.startswith(IMAGE_PROMPT_PREFIX):
            pv = IMAGE_PROMPT_PREFIX + pv
        img = generar_imagen_ia(model_img_id, pv)
        if img:
            pic = doc.add_paragraph()
            pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                pic.add_run().add_picture(img, width=Inches(2.7))
            except Exception:
                pass

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ============================================================
# Diagnóstico (entrada + JSON)
# ============================================================
def run_diagnostic_activity(
    modo: str,
    materia: str,
    nivel: str,
    tema: str,
    estilo_extra: str,
    original_text: str,
    df_rows: pd.DataFrame,
    alumno_col: str,
    grupo_col: str,
    diag_col: str,
    grado_text: str
) -> Dict[str, Any]:
    ok_in, msg_in, info_in = validate_exam_text(original_text) if modo == "ADAPTAR" else (True, "OK", {
        "chars": len(original_text or ""),
        "lines": (original_text or "").count("\n") + (1 if original_text else 0),
        "pipes": (original_text or "").count("|"),
        "preview": (original_text or "")[:1400],
    })

    out: Dict[str, Any] = {
        "input_ok": ok_in,
        "input_msg": msg_in,
        "input_info": info_in,
        "json_ok": False,
        "json_msg": "",
        "json_preview": "",
        "mode": "",
        "max_tokens": 0,
        "model_text": BOOT["text_model"],
        "model_image": BOOT.get("image_model"),
    }
    if not ok_in:
        return out
    if len(df_rows) == 0:
        out["json_ok"] = False
        out["json_msg"] = "No hay alumnos para probar JSON."
        return out

    row = df_rows.iloc[0]
    n = str(row[alumno_col]).strip()
    g = str(row[grupo_col]).strip()
    d = str(row[diag_col]).strip()

    prompt_full = build_prompt_opalpp(modo, materia, nivel, tema, estilo_extra, n, g, d, grado_text, original_text)
    prompt_comp = build_prompt_compact_opalpp(modo, materia, nivel, tema, estilo_extra, n, g, d, grado_text, original_text)

    cache_key = f"DIAG::{hash_text(modo + materia + nivel + tema + estilo_extra + (original_text or ''))}::{BOOT['text_model']}::{n}::{g}::{d}::{grado_text}"

    try:
        data, mode_used, max_t = request_activity_ultra(BOOT["text_model"], prompt_full, prompt_comp, cache_key)
        okj, whyj = validate_activity_json(data)
        out["json_ok"] = okj
        out["json_msg"] = whyj
        out["mode"] = mode_used
        out["max_tokens"] = max_t
        out["json_preview"] = json.dumps(data, ensure_ascii=False, indent=2)[:2800]
        return out
    except Exception as e:
        out["json_ok"] = False
        out["json_msg"] = f"{type(e).__name__}: {e}"
        return out

# ============================================================
# UI + Proceso
# ============================================================
def main():
    st.title("Motor Pedagógico v17.0 (Opal++)")
    st.caption("Adaptar desde DOCX o Crear desde brief. Salida inclusiva con JSON determinista y DOCX dyslexia-friendly.")

    # Planilla
    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando planilla: {e}")
        return

    # Column mapping por posición (compat)
    grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
    grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
    diag_col = df.columns[4] if len(df.columns) > 4 else df.columns[0]

    with st.sidebar:
        st.header("Modelos (auto)")
        st.write(f"Boot: {BOOT.get('boot_time')}")
        st.write(f"Texto: {BOOT.get('text_model')}")
        if BOOT.get("image_model"):
            st.write(f"Imagen: {BOOT.get('image_model')}")
            st.caption(BOOT.get("image_reason", ""))
        else:
            st.write("Imagen: N/A")
            st.caption(BOOT.get("image_reason", ""))

        st.divider()
        st.header("Selección grupo")
        grado_sel = st.selectbox("Grado (planilla)", sorted(df[grado_col].dropna().unique().tolist()))
        df_f = df[df[grado_col] == grado_sel].copy()

        alcance = st.radio("Alcance", ["Todo el grado", "Seleccionar alumnos"], horizontal=True)
        alumnos_lista = df_f[alumno_col].dropna().unique().tolist()
        if alcance == "Seleccionar alumnos":
            seleccion = st.multiselect("Alumnos", alumnos_lista)
            alumnos_final = df_f[df_f[alumno_col].isin(seleccion)].copy() if seleccion else df_f.iloc[0:0].copy()
            if len(alumnos_final) == 0:
                st.info("Seleccioná al menos 1 alumno.")
        else:
            alumnos_final = df_f

        st.divider()
        st.header("Modo de entrada")
        modo = st.radio("Elegí modo", ["ADAPTAR (subir DOCX)", "CREAR (solo brief)"])
        modo_key = "ADAPTAR" if modo.startswith("ADAPTAR") else "CREAR"

        st.divider()
        st.header("Contexto pedagógico")
        materia = st.text_input("Materia", value="Matemática")
        nivel = st.text_input("Nivel", value="7mo grado")
        tema = st.text_input("Tema", value="División progresiva")
        estilo_extra = st.text_area("Instrucciones extra (opcional)", placeholder="Ej: incluir verdadero/falso + 1 ejemplo guiado, y reducir a 6 ítems.")

        st.divider()
        st.header("Assets / salida")
        activar_img_user = st.checkbox("Generar imagen de apoyo (si disponible)", value=True)
        activar_img = activar_img_user and (BOOT.get("image_model") is not None)
        if activar_img_user and not activar_img:
            st.warning("Imágenes desactivadas: no se detectó un modelo de imagen funcional.")

        logo = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
        logo_bytes = logo.read() if logo else None

        st.divider()
        st.header("Diagnóstico")
        diag_mode = st.checkbox("Activar diagnóstico", value=True)
        st.caption(f"Pydantic: {'ON' if PYDANTIC_AVAILABLE else 'OFF'}")

    st.subheader("Entrada")
    original_text = ""

    if modo_key == "ADAPTAR":
        archivo = st.file_uploader("Subir actividad/examen base (DOCX)", type=["docx"])
        if archivo:
            original_text = extraer_texto_docx(archivo)
            ok_in, msg_in, info_in = validate_exam_text(original_text)
            if diag_mode:
                if ok_in:
                    st.success(f"Parseo DOCX: OK ({info_in['chars']} chars, {info_in['lines']} líneas)")
                else:
                    st.error(f"Parseo DOCX: {msg_in}")
                with st.expander("Preview texto extraído", expanded=False):
                    st.text(info_in.get("preview", ""))
        else:
            st.info("Subí un DOCX para adaptar.")
    else:
        brief = st.text_area(
            "Brief de actividad (texto libre)",
            placeholder="Ej: Matemática 7mo grado. División progresiva. 1 ejemplo guiado. 6 ejercicios. 2 de selección múltiple y 2 de completar."
        )
        original_text = (brief or "").strip()
        if diag_mode:
            if original_text:
                st.success(f"Brief: OK ({len(original_text)} chars)")
                with st.expander("Preview brief", expanded=False):
                    st.text(original_text[:1400])
            else:
                st.error("Brief vacío. Escribí una descripción.")

    # Diagnóstico JSON
    if diag_mode and st.button("Probar generación (primer alumno)"):
        if len(alumnos_final) == 0:
            st.error("No hay alumnos para probar.")
        else:
            if modo_key == "ADAPTAR":
                ok_in, msg_in, _ = validate_exam_text(original_text)
                if not ok_in:
                    st.error(f"No se puede probar: {msg_in}")
                else:
                    dres = run_diagnostic_activity(modo_key, materia, nivel, tema, estilo_extra, original_text, alumnos_final, alumno_col, grupo_col, diag_col, str(grado_sel))
                    if dres["json_ok"]:
                        st.success(f"JSON OK | mode={dres['mode']} | max_tokens={dres['max_tokens']}")
                    else:
                        st.error(f"JSON FAIL: {dres['json_msg']}")
                    with st.expander("Preview JSON (truncado)", expanded=False):
                        st.text(dres.get("json_preview", ""))
            else:
                if not original_text:
                    st.error("Brief vacío.")
                else:
                    dres = run_diagnostic_activity(modo_key, materia, nivel, tema, estilo_extra, original_text, alumnos_final, alumno_col, grupo_col, diag_col, str(grado_sel))
                    if dres["json_ok"]:
                        st.success(f"JSON OK | mode={dres['mode']} | max_tokens={dres['max_tokens']}")
                    else:
                        st.error(f"JSON FAIL: {dres['json_msg']}")
                    with st.expander("Preview JSON (truncado)", expanded=False):
                        st.text(dres.get("json_preview", ""))

    # Procesamiento lote
    if st.button("🚀 GENERAR LOTE (ZIP)"):
        if len(alumnos_final) == 0:
            st.error("No hay alumnos para procesar (selección vacía).")
            return

        if modo_key == "ADAPTAR":
            ok_in, msg_in, info_in = validate_exam_text(original_text)
            if not ok_in:
                st.error(f"No se inicia: {msg_in}")
                return
            input_hash = hash_text(original_text)
        else:
            if not original_text:
                st.error("Brief vacío. No se inicia.")
                return
            info_in = {
                "chars": len(original_text),
                "lines": original_text.count("\n") + 1,
                "pipes": original_text.count("|"),
                "preview": original_text[:1400],
            }
            input_hash = hash_text(original_text)

        model_text = BOOT["text_model"]
        model_img = BOOT.get("image_model") if activar_img else None

        zip_io = io.BytesIO()
        logs: List[str] = []
        errors: List[str] = []
        ok_count = 0

        logs.append("Motor Pedagógico v17.0 (Opal++)")
        logs.append(f"Inicio: {now_str()}")
        logs.append(f"Modo: {modo_key}")
        logs.append(f"Materia: {materia} | Nivel: {nivel} | Tema: {tema}")
        logs.append(f"Estilo extra: {estilo_extra}")
        logs.append(f"Modelo texto: {model_text}")
        logs.append(f"Modelo imagen: {model_img if model_img else 'N/A'}")
        logs.append(f"Imágenes habilitadas: {bool(model_img)}")
        logs.append(f"Grado planilla: {grado_sel}")
        logs.append(f"Total alumnos: {len(alumnos_final)}")
        logs.append(f"Hash entrada: {input_hash}")
        logs.append(f"Pydantic: {'ON' if PYDANTIC_AVAILABLE else 'OFF'}")
        logs.append(f"Entrada chars: {info_in.get('chars')} | líneas: {info_in.get('lines')} | pipes: {info_in.get('pipes')}")
        logs.append("")

        with zipfile.ZipFile(zip_io, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("_REPORTE.txt", "\n".join(logs))

            prog = st.progress(0.0)
            status = st.empty()

            total = len(alumnos_final)
            for i, (_, row) in enumerate(alumnos_final.iterrows(), start=1):
                n = str(row[alumno_col]).strip()
                g = str(row[grupo_col]).strip()
                d = str(row[diag_col]).strip()

                status.info(f"Generando: {n} ({i}/{total})")

                try:
                    prompt_full = build_prompt_opalpp(
                        modo_key, materia, nivel, tema, estilo_extra,
                        alumno_nombre=n,
                        alumno_grupo=g,
                        alumno_diag=d,
                        alumno_grado=str(grado_sel),
                        original_text=original_text
                    )
                    prompt_comp = build_prompt_compact_opalpp(
                        modo_key, materia, nivel, tema, estilo_extra,
                        alumno_nombre=n,
                        alumno_grupo=g,
                        alumno_diag=d,
                        alumno_grado=str(grado_sel),
                        original_text=original_text
                    )

                    cache_key = f"{input_hash}::{model_text}::{modo_key}::{materia}::{nivel}::{tema}::{estilo_extra}::{n}::{g}::{d}::{grado_sel}"

                    data, mode_used, max_t = request_activity_ultra(model_text, prompt_full, prompt_comp, cache_key)

                    # Inyecta alumno/contexto si el modelo omitió algo mínimo (robustez extra)
                    data.setdefault("alumno", {})
                    data["alumno"].setdefault("nombre", n)
                    data["alumno"].setdefault("grupo", g)
                    data["alumno"].setdefault("diagnostico", d)
                    data["alumno"].setdefault("grado", str(grado_sel))
                    data.setdefault("contexto", {})
                    data["contexto"].setdefault("modo", modo_key)
                    data["contexto"].setdefault("materia", materia)
                    data["contexto"].setdefault("nivel", nivel)
                    data["contexto"].setdefault("tema", tema)
                    data["contexto"].setdefault("estilo_extra", estilo_extra or "")

                    # Normaliza visual
                    v = normalize_visual(data.get("visual", {}))
                    if v.get("habilitado"):
                        pv = str(v.get("prompt", "")).strip()
                        if pv and not pv.startswith(IMAGE_PROMPT_PREFIX):
                            v["prompt"] = IMAGE_PROMPT_PREFIX + pv
                    data["visual"] = v

                    okj, whyj = validate_activity_json(data)
                    if not okj:
                        raise ValueError(f"JSON final inválido: {whyj}")

                    docx_bytes = render_docx_opalpp(data, logo_bytes, activar_img=bool(model_img), model_img_id=model_img)
                    zf.writestr(f"Actividad_{safe_filename(n)}.docx", docx_bytes)

                    zf.writestr(f"_META_{safe_filename(n)}.txt", f"mode={mode_used}\nmax_tokens={max_t}\nitems={len(data.get('items',[]))}\n")
                    ok_count += 1

                except Exception as e:
                    msg = f"{n} :: {type(e).__name__} :: {e}"
                    errors.append(msg)
                    zf.writestr(f"ERROR_{safe_filename(n)}.txt", msg)

                prog.progress(i / total)

            resumen = []
            resumen.append("RESUMEN")
            resumen.append(f"Fin: {now_str()}")
            resumen.append(f"OK: {ok_count} / {total}")
            resumen.append(f"Errores: {len(errors)}")
            if errors:
                resumen.append("")
                resumen.append("ERRORES (primeros 200):")
                resumen.extend([f"- {e}" for e in errors[:200]])
                if len(errors) > 200:
                    resumen.append(f"... truncado ({len(errors)} errores totales)")
            zf.writestr("_RESUMEN.txt", "\n".join(resumen))

        st.success(f"Lote finalizado. OK: {ok_count} | Errores: {len(errors)}")
        st.download_button("📥 Descargar ZIP", zip_io.getvalue(), "actividades_opalpp_v17_0.zip", mime="application/zip")

if __name__ == "__main__":
    main()
