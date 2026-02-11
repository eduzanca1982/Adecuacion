import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import zipfile
import time
import random
import hashlib
import base64
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# Nano Opal v22.0
# Objetivo:
# - Mantener selector por GRADO (planilla con dificultades por alumno)
# - Robustez JSON (no depender de response.text) + reparación 1 vez
# - Estética "Opal Card" en DOCX (sin itálicas, Verdana 14, line spacing 1.5-1.8)
# - Imágenes: soportar gemini-2.5-flash-image aunque NO devuelva inline_data por default.
#   => Forzamos responseModalities/Image en generation_config + parsers múltiples:
#      inline_data / inlineData / data URI base64 / texto base64
# - Boot: consultar primero, ejecutar después, y desactivar imagen si smoke test falla
# - ZIP siempre con _REPORTE.txt y _RESUMEN.txt + ERROR_{alumno}.txt
# ============================================================

st.set_page_config(page_title="Nano Opal v22.0 🍌", layout="wide", page_icon="🍌")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

RETRIES = 6
CACHE_TTL_SECONDS = 6 * 60 * 60

# Visual prompt requerido por tu consigna
IMAGE_PROMPT_PREFIX = "Pictograma estilo ARASAAC, trazos negros gruesos, fondo blanco, ultra simple, sin sombras de: "
MIN_IMAGE_BYTES = 1200  # umbral realista para PNG 1024

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

OUT_TOKEN_STEPS_FULL = [4096, 6144, 8192]
OUT_TOKEN_STEPS_COMPACT = [2048, 4096]

ACTION_EMOJI_BY_TIPO = {
    "completar": "✍️",
    "escritura": "✍️",
    "multiple choice": "🔢",
    "multiple_choice": "🔢",
    "seleccion": "🔢",
    "unir": "📖",
    "lectura": "📖",
    "verdadero_falso": "📖",
    "problema_guiado": "🔢",
    "calcular": "🔢",
    "dibujar": "🎨",
    "arte": "🎨",
}

SYSTEM_PROMPT_OPALPP = f"""
Actúa como un Senior Inclusive UX Designer y Tutor Psicopedagogo.
Tu output debe ser una FICHA estilo "Card" (como HTML), neuroinclusiva, dislexia-friendly.

REGLAS DE ORO:
- ICONOGRAFÍA: Cada item en items[] debe iniciar su enunciado con un emoji de acción:
  ✍️ completar/escribir, 📖 leer, 🔢 calcular, 🎨 dibujar.
- MICRO-PASOS: pista_visual debe ser andamiaje concreto físico/visual. No teoría.
- SIN ITÁLICAS: Prohibido usar itálicas en cualquier campo.
- KEYWORDS: usa **negrita** solo como anclaje visual.
- ESTILO DE IMAGEN: Si visual.habilitado=true, visual.prompt DEBE empezar EXACTAMENTE con:
  "{IMAGE_PROMPT_PREFIX}[OBJETO]"

SALIDA: JSON puro, sin markdown, sin texto extra, sin backticks.

ESQUEMA EXACTO:
{{
  "objetivo_aprendizaje": "string",
  "consigna_adaptada": "string",
  "items": [
    {{
      "tipo": "calcular|lectura|escritura|dibujar|multiple choice|unir|completar|verdadero_falso|problema_guiado",
      "enunciado": "string (DEBE EMPEZAR con emoji de acción)",
      "opciones": ["string","string"],
      "pista_visual": "string (micro-pasos concretos)"
    }}
  ],
  "adecuaciones_aplicadas": ["string","string"],
  "sugerencias_docente": ["string","string"],
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
""".strip()


# ============================================================
# Helpers
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
        "connection reset", "temporarily", "service unavailable"
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

def ensure_action_emoji(tipo: str, enunciado: str) -> str:
    t = (tipo or "").strip().lower()
    e = (enunciado or "").strip()
    if not e:
        return e
    if any(e.startswith(x) for x in ["✍️", "📖", "🔢", "🎨"]):
        return e
    emoji = ACTION_EMOJI_BY_TIPO.get(t, "📖")
    return f"{emoji} {e}"

def normalize_visual_prompt(p: str) -> str:
    p = (p or "").strip()
    if not p:
        return p
    if p.startswith(IMAGE_PROMPT_PREFIX):
        return p
    return IMAGE_PROMPT_PREFIX + p

def validate_text_input(text: str, mode: str) -> Tuple[bool, str, Dict[str, Any]]:
    info = {
        "chars": len(text or ""),
        "lines": (text or "").count("\n") + (1 if text else 0),
        "preview": (text or "")[:1600],
    }
    if mode == "ADAPTAR":
        if not text or not text.strip():
            return False, "TEXTO vacío tras extracción (posible actividad en imágenes/shapes).", info
        if len(text) < 120:
            return False, "TEXTO muy corto (<120 chars).", info
        return True, "OK", info
    # CREAR
    if not text or not text.strip():
        return False, "Brief vacío.", info
    return True, "OK", info


# ============================================================
# DOCX extraction (párrafos + tablas)
# ============================================================
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def _extract_text_from_el(el) -> str:
    return "".join(n.text for n in el.iter() if n.tag == f"{W_NS}t" and n.text).strip()

def extraer_texto_docx(file) -> str:
    doc = Document(file)
    out: List[str] = []
    for el in doc.element.body:
        if el.tag == f"{W_NS}p":
            t = _extract_text_from_el(el)
            if t:
                out.append(t)
        elif el.tag == f"{W_NS}tbl":
            for row in el.findall(f".//{W_NS}tr"):
                cells = [_extract_text_from_el(c) for c in row.findall(f".//{W_NS}tc")]
                if any(cells):
                    out.append(" | ".join(cells))
            out.append("")
    return "\n".join(out).strip()


# ============================================================
# Robust response parsing (TEXT)
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
        out = "".join(chunks).strip()
        return out if out else None
    except Exception:
        return None

def _finish_reason(resp) -> Optional[int]:
    try:
        return int(resp.candidates[0].finish_reason)
    except Exception:
        return None


# ============================================================
# Robust image parsing (INLINE / inlineData / data URI / base64 text)
# ============================================================
DATA_URI_RE = re.compile(r"data:image/(png|jpeg|jpg|webp);base64,([A-Za-z0-9+/=\n\r]+)")

def _maybe_b64_to_bytes(x: Any) -> Optional[bytes]:
    if x is None:
        return None
    if isinstance(x, (bytes, bytearray)):
        return bytes(x)
    if isinstance(x, str):
        s = x.strip()
        # data uri
        m = DATA_URI_RE.search(s)
        if m:
            b64 = m.group(2)
            try:
                return base64.b64decode(b64, validate=False)
            except Exception:
                return None
        # plain base64 (heurística)
        if len(s) > 400 and all(c in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=\n\r" for c in s[:200]):
            try:
                return base64.b64decode(s, validate=False)
            except Exception:
                return None
    return None

def _extract_inline_bytes_or_none(resp) -> Optional[bytes]:
    """
    Intenta extraer imagen desde:
    - part.inline_data.data (SDK viejo)
    - part.inlineData.data (SDK/format alternativo)
    - part.text con data URI o base64
    """
    try:
        cand = resp.candidates[0]
        content = getattr(cand, "content", None)
        if not content or not getattr(content, "parts", None):
            return None

        for part in content.parts:
            # 1) inline_data (python google.generativeai viejo)
            inline = getattr(part, "inline_data", None)
            if inline is not None:
                data = getattr(inline, "data", None)
                b = _maybe_b64_to_bytes(data) or (data if isinstance(data, (bytes, bytearray)) else None)
                if b:
                    return b

            # 2) inlineData (otros bindings)
            inline2 = getattr(part, "inlineData", None)
            if inline2 is not None:
                data2 = getattr(inline2, "data", None)
                b2 = _maybe_b64_to_bytes(data2)
                if b2:
                    return b2

            # 3) texto con data uri / base64
            t = getattr(part, "text", None)
            b3 = _maybe_b64_to_bytes(t)
            if b3:
                return b3

        return None
    except Exception:
        return None

def _looks_like_image(b: bytes) -> bool:
    if not b or len(b) < MIN_IMAGE_BYTES:
        return False
    # PNG
    if b[:8] == b"\x89PNG\r\n\x1a\n":
        return True
    # JPEG
    if b[:3] == b"\xff\xd8\xff":
        return True
    # WEBP (RIFF....WEBP)
    if b[:4] == b"RIFF" and b[8:12] == b"WEBP":
        return True
    return False

def generate_image_bytes(model_id: str, prompt_img: str) -> Optional[bytes]:
    """
    Para gemini-2.5-flash-image es CRÍTICO forzar response modalities.
    En REST/documentación se usa generationConfig.responseModalities=["Image"].
    El SDK viejo puede ignorarlo, pero lo intentamos en variantes.
    """
    if not model_id:
        return None

    prompt_img = normalize_visual_prompt(prompt_img)

    def call_with_cfg(cfg: Optional[Dict[str, Any]]):
        m = genai.GenerativeModel(model_id)
        if cfg is None:
            return m.generate_content(prompt_img, safety_settings=SAFETY_SETTINGS)
        return m.generate_content(prompt_img, generation_config=cfg, safety_settings=SAFETY_SETTINGS)

    # Variantes de config para maximizar compatibilidad con SDK viejo
    cfg_variants = [
        {"response_modalities": ["Image"]},
        {"response_modalities": ["IMAGE"]},
        {"responseModalities": ["Image"]},   # por si el binding respeta camelCase
        {"responseModalities": ["IMAGE"]},
        None,  # último recurso
    ]

    last = None
    for cfg in cfg_variants:
        try:
            resp = retry_with_backoff(lambda: call_with_cfg(cfg))
            b = _extract_inline_bytes_or_none(resp)
            if b and _looks_like_image(b):
                return b
        except Exception as e:
            last = e
            continue
    return None


# ============================================================
# Boot Scan (mantiene preferidos y smoke test real)
# ============================================================
def smoke_test_image_model(model_id: str) -> Tuple[bool, str]:
    try:
        b = generate_image_bytes(model_id, IMAGE_PROMPT_PREFIX + "manzana")
        if not b:
            return False, "No se obtuvo imagen válida (sin bytes o formato no reconocido)"
        return True, f"OK bytes={len(b)}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

def boot_pick_models(prefer_text: str, prefer_image: str) -> Dict[str, Any]:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])

    txt = prefer_text.strip() if prefer_text else None

    img = None
    img_reason = ""
    if prefer_image and prefer_image.strip():
        ok, msg = smoke_test_image_model(prefer_image.strip())
        if ok:
            img = prefer_image.strip()
            img_reason = f"Preferido OK: {msg}"
        else:
            img = None
            img_reason = f"Preferido FAIL: {msg}"

    return {"txt": txt, "img": img, "img_reason": img_reason, "boot_time": now_str()}

@st.cache_resource(show_spinner=False)
def boot_cached(prefer_text: str, prefer_image: str) -> Dict[str, Any]:
    try:
        return boot_pick_models(prefer_text, prefer_image)
    except Exception as e:
        return {"txt": None, "img": None, "img_reason": f"Boot error: {e}", "boot_time": now_str()}


# ============================================================
# JSON generation + validación + repair
# ============================================================
def validate_activity_json(data: Dict[str, Any]) -> Tuple[bool, str]:
    try:
        if not isinstance(data, dict):
            return False, "Root no es objeto"

        required = [
            "objetivo_aprendizaje", "consigna_adaptada", "items",
            "adecuaciones_aplicadas", "sugerencias_docente", "visual", "control_calidad"
        ]
        for k in required:
            if k not in data:
                return False, f"Falta clave: {k}"

        if not isinstance(data["items"], list) or len(data["items"]) < 1:
            return False, "items vacío/no lista"

        cc = data.get("control_calidad", {})
        if not isinstance(cc, dict):
            return False, "control_calidad no es objeto"
        if cc.get("items_count") != len(data["items"]):
            return False, "control_calidad.items_count != len(items)"

        v = data.get("visual", {})
        if not isinstance(v, dict):
            return False, "visual no es objeto"
        if normalize_bool(v.get("habilitado", False)):
            p = str(v.get("prompt", "")).strip()
            if not p.startswith(IMAGE_PROMPT_PREFIX):
                return False, "visual.prompt no respeta prefijo ARASAAC"

        for i, it in enumerate(data["items"][:200]):
            if not isinstance(it, dict):
                return False, f"items[{i}] no es objeto"
            en = str(it.get("enunciado", "")).strip()
            if not en:
                return False, f"items[{i}].enunciado vacío"
            if not any(en.startswith(x) for x in ["✍️", "📖", "🔢", "🎨"]):
                return False, f"items[{i}].enunciado no inicia con emoji"
            if "pista_visual" not in it:
                return False, f"items[{i}] falta pista_visual"

        return True, "OK"
    except Exception as e:
        return False, f"Exception validando: {e}"

def build_repair_prompt(bad: str, why: str) -> str:
    return f"""
Devuelve EXCLUSIVAMENTE un JSON válido y corregido (sin texto extra).

Problema detectado:
{why}

JSON A CORREGIR:
{bad}

Reglas:
- Cumplir esquema exacto.
- control_calidad.items_count == len(items)
- items[].enunciado inicia con emoji de acción (✍️📖🔢🎨)
- items[].pista_visual presente y es micro-pasos concretos
- visual.prompt inicia con "{IMAGE_PROMPT_PREFIX}" si visual.habilitado=true
- Prohibido itálicas. Solo **negrita** para keywords.
""".strip()

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
        # capturamos la 1ra salida cruda y pedimos repair
        m = genai.GenerativeModel(model_id)
        cfg = dict(BASE_GEN_CFG_JSON)
        cfg["max_output_tokens"] = max_out

        resp1 = retry_with_backoff(lambda: m.generate_content(prompt, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
        raw = _extract_text_or_none(resp1)
        fr = _finish_reason(resp1)
        if raw is None:
            raise ValueError(f"Empty candidate (finish_reason={fr})")

        repair = build_repair_prompt(raw, f"{type(e).__name__}: {e}")
        resp2 = retry_with_backoff(lambda: m.generate_content(repair, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
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
def cached_generate(cache_key: str, model_id: str, prompt: str, max_out: int) -> Dict[str, Any]:
    return generate_json_with_repair(model_id, prompt, max_out)

def request_activity_ultra(model_id: str, prompt_full: str, prompt_compact: str, cache_key: str) -> Tuple[Dict[str, Any], str, int]:
    last_err = None
    for t in OUT_TOKEN_STEPS_FULL:
        try:
            data = cached_generate(cache_key + f"::FULL::{t}", model_id, prompt_full, t)
            return data, "FULL", t
        except Exception as e:
            last_err = e
    for t in OUT_TOKEN_STEPS_COMPACT:
        try:
            data = cached_generate(cache_key + f"::COMPACT::{t}", model_id, prompt_compact, t)
            return data, "COMPACT", t
        except Exception as e:
            last_err = e
    raise last_err if last_err else RuntimeError("Fallo desconocido generando actividad")


# ============================================================
# Render DOCX "Card style"
# ============================================================
def apply_card_style(cell):
    tc_pr = cell._tc.get_or_add_tcPr()

    # sombreado
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), "FAFAFA")
    tc_pr.append(shd)

    # bordes
    tc_borders = OxmlElement('w:tcBorders')
    for b in ['top', 'left', 'bottom', 'right']:
        edge = OxmlElement(f'w:{b}')
        edge.set(qn('w:val'), 'single')
        edge.set(qn('w:sz'), '4')
        edge.set(qn('w:space'), '0')
        edge.set(qn('w:color'), 'E0E0E0')
        tc_borders.append(edge)
    tc_pr.append(tc_borders)

def clear_paragraph(paragraph):
    p = paragraph._p
    for child in list(p):
        p.remove(child)

def add_runs_with_bold_markers(paragraph, text: str, font_name: str = "Verdana", font_size_pt: int = 14, bold_default: bool = False):
    parts = str(text or "").split("**")
    for i, part in enumerate(parts):
        run = paragraph.add_run(part)
        run.bold = (not bold_default) if (i % 2 == 1) else bold_default
        run.italic = False  # prohibido itálicas
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)

def add_response_line(paragraph):
    paragraph.paragraph_format.line_spacing = 1.6
    run = paragraph.add_run("✍️ Mi respuesta: ")
    run.bold = True
    run.italic = False
    run.font.name = "Verdana"
    run.font.size = Pt(14)
    run2 = paragraph.add_run("______________________________________________")
    run2.bold = False
    run2.italic = False
    run2.font.name = "Verdana"
    run2.font.size = Pt(14)

def render_opal_docx(data: Dict[str, Any], alumno: Dict[str, str], logo_b: Optional[bytes], img_model_id: Optional[str], enable_img: bool) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Verdana'
    style.font.size = Pt(14)

    # header
    header = doc.add_table(rows=1, cols=2)
    header.width = Inches(6.5)
    if logo_b:
        try:
            header.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_b), width=Inches(0.7))
        except Exception:
            pass

    info = header.rows[0].cells[1].paragraphs[0]
    info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = info.add_run(f"{alumno.get('nombre','')}\n{alumno.get('diagnostico','')}\nGrupo: {alumno.get('grupo','')} | Grado: {alumno.get('grado','')}")
    run.bold = True
    run.italic = False
    run.font.name = "Verdana"
    run.font.size = Pt(11)

    doc.add_paragraph("")

    # Objetivo
    p_t = doc.add_paragraph()
    rt = p_t.add_run("Objetivo de aprendizaje")
    rt.bold = True
    rt.italic = False
    rt.font.size = Pt(14)

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.6
    add_runs_with_bold_markers(p, data.get("objetivo_aprendizaje", ""))

    # Consigna
    p_t = doc.add_paragraph()
    rt = p_t.add_run("Consigna adaptada")
    rt.bold = True
    rt.italic = False
    rt.font.size = Pt(14)

    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.6
    add_runs_with_bold_markers(p, data.get("consigna_adaptada", ""))

    doc.add_paragraph("")

    # Imagen global (1 por ficha, no por item)
    img_bytes = None
    visual = data.get("visual", {}) if isinstance(data.get("visual", {}), dict) else {}
    if enable_img and img_model_id and normalize_bool(visual.get("habilitado", False)):
        pr = str(visual.get("prompt", "")).strip()
        if pr:
            img_bytes = generate_image_bytes(img_model_id, pr)

    if img_bytes:
        pimg = doc.add_paragraph()
        pimg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            pimg.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(2.2))
        except Exception:
            pass
        doc.add_paragraph("")

    # Cards
    for it in data.get("items", []):
        if not isinstance(it, dict):
            continue

        tipo = str(it.get("tipo", "")).strip()
        enunciado = ensure_action_emoji(tipo, str(it.get("enunciado", "")).strip())
        pista = str(it.get("pista_visual", "")).strip()

        opciones = it.get("opciones", [])
        if not isinstance(opciones, list):
            opciones = []

        table = doc.add_table(rows=1, cols=1)
        table.width = Inches(6.5)
        cell = table.rows[0].cells[0]
        apply_card_style(cell)
        clear_paragraph(cell.paragraphs[0])

        # Enunciado
        pe = cell.add_paragraph()
        pe.paragraph_format.line_spacing = 1.8
        add_runs_with_bold_markers(pe, enunciado, bold_default=True)

        # Opciones / respuesta
        if opciones:
            for opt in opciones[:10]:
                po = cell.add_paragraph()
                po.paragraph_format.line_spacing = 1.6
                ro = po.add_run(f"☐ {str(opt)}")
                ro.bold = False
                ro.italic = False
                ro.font.name = "Verdana"
                ro.font.size = Pt(14)
        else:
            pr = cell.add_paragraph()
            add_response_line(pr)

        # Pista verde (sin itálica)
        if pista:
            pp = cell.add_paragraph()
            pp.paragraph_format.line_spacing = 1.6
            rp = pp.add_run(f"💡 {pista}")
            rp.bold = False
            rp.italic = False
            rp.font.color.rgb = RGBColor(0, 150, 0)
            rp.font.name = "Verdana"
            rp.font.size = Pt(14)

        doc.add_paragraph("")

    # Adecuaciones
    p_t = doc.add_paragraph()
    rt = p_t.add_run("Adecuaciones aplicadas")
    rt.bold = True
    rt.italic = False
    rt.font.size = Pt(14)

    for a in (data.get("adecuaciones_aplicadas", []) or [])[:30]:
        pa = doc.add_paragraph(f"• {a}")
        pa.paragraph_format.line_spacing = 1.6

    # Sugerencias
    p_t = doc.add_paragraph()
    rt = p_t.add_run("Sugerencias para el docente")
    rt.bold = True
    rt.italic = False
    rt.font.size = Pt(14)

    for s in (data.get("sugerencias_docente", []) or [])[:30]:
        ps = doc.add_paragraph(f"• {s}")
        ps.paragraph_format.line_spacing = 1.6

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# UI + Proceso por GRADO (se mantiene)
# ============================================================
def main():
    st.title("Nano Opal v22.0 🧠🍌")
    st.caption("Opal-card DOCX + JSON determinista + imágenes robustas para gemini-2.5-flash-image")

    # Planilla
    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando planilla: {e}")
        return

    # Mapeo columnas (mismo patrón que venías usando)
    grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
    grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
    diag_col = df.columns[4] if len(df.columns) > 4 else df.columns[0]

    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuración")

        prefer_txt = st.text_input("Modelo texto (preferido)", value="gemini-1.5-flash")
        prefer_img = st.text_input("Modelo imagen (preferido)", value="gemini-2.5-flash-image")

        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Reboot (rescan)"):
                st.cache_resource.clear()
        with col_b:
            if st.button("Limpiar cache"):
                st.cache_data.clear()

        CONFIG = boot_cached(prefer_txt, prefer_img)

        st.write(f"Boot: {CONFIG.get('boot_time')}")
        st.success(f"Texto: {CONFIG.get('txt')}" if CONFIG.get("txt") else "Texto: N/A")
        if CONFIG.get("img"):
            st.success(f"Imagen: {CONFIG.get('img')}")
        else:
            st.warning("Imagen: N/A (desactivado)")
        st.caption(CONFIG.get("img_reason", ""))

        st.divider()

        # Mantener selector por GRADO
        st.subheader("Grupo / Planilla")
        grado = st.selectbox("Grado", sorted(df[grado_col].dropna().unique().tolist()))
        df_f = df[df[grado_col] == grado].copy()

        alcance = st.radio("Alcance", ["Todo el grado", "Seleccionar alumnos"], horizontal=True)
        if alcance == "Seleccionar alumnos":
            al_sel = st.multiselect("Alumnos", sorted(df_f[alumno_col].dropna().unique().tolist()))
            df_final = df_f[df_f[alumno_col].isin(al_sel)].copy() if al_sel else df_f.iloc[0:0].copy()
        else:
            df_final = df_f

        st.divider()
        enable_img = st.checkbox("Habilitar imagen", value=True)
        enable_img = enable_img and bool(CONFIG.get("img"))

        logo = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
        l_bytes = logo.read() if logo else None

        st.divider()
        inst_style = st.text_area("Instrucciones de Estilo On-the-fly", height=120, placeholder="Ej: 2 ejemplos guiados, 6 ítems, 2 multiple choice, 2 completar, usar objetos concretos.")

    if not CONFIG.get("txt"):
        st.error("Falta modelo de texto o API Key.")
        return

    tab1, tab2 = st.tabs(["🔄 Adaptar DOCX", "✨ Crear Actividad"])

    adapt_docx = None
    brief = ""

    with tab1:
        st.subheader("Adaptar (DOCX)")
        adapt_docx = st.file_uploader("Examen/actividad base (DOCX)", type=["docx"], key="docx_in")

    with tab2:
        st.subheader("Crear desde brief")
        brief = st.text_area(
            "Prompt/brief",
            height=180,
            placeholder="Ej: Matemática 7mo grado, división progresiva. 1 ejemplo guiado. 6 ítems. 2 multiple choice y 2 completar."
        )

    mode = "CREAR" if (brief and brief.strip()) else "ADAPTAR"
    input_text = ""

    if mode == "ADAPTAR":
        if adapt_docx:
            input_text = extraer_texto_docx(adapt_docx)
            ok_in, msg_in, info_in = validate_text_input(input_text, "ADAPTAR")
            if ok_in:
                st.success(f"Parseo DOCX OK ({info_in['chars']} chars)")
            else:
                st.error(f"Parseo DOCX: {msg_in}")
            with st.expander("Preview texto extraído", expanded=False):
                st.text(info_in.get("preview", ""))
        else:
            st.info("Subí un DOCX o usa el tab 'Crear Actividad'.")
    else:
        input_text = brief.strip()
        ok_in, msg_in, info_in = validate_text_input(input_text, "CREAR")
        if ok_in:
            st.success(f"Brief OK ({info_in['chars']} chars)")
        else:
            st.error(f"Brief: {msg_in}")
        with st.expander("Preview brief", expanded=False):
            st.text(info_in.get("preview", ""))

    if st.button("🚀 GENERAR LOTE"):
        if len(df_final) == 0:
            st.error("No hay alumnos (ver selección por grado / alumnos).")
            return

        if mode == "ADAPTAR" and not adapt_docx:
            st.error("Falta DOCX para adaptar.")
            return

        ok_in, msg_in, _ = validate_text_input(input_text, mode)
        if not ok_in:
            st.error(f"No se inicia: {msg_in}")
            return

        zip_io = io.BytesIO()
        logs = []
        errors = []
        ok_count = 0

        logs.append("Nano Opal v22.0")
        logs.append(f"Inicio: {now_str()}")
        logs.append(f"Modo: {mode}")
        logs.append(f"Modelo texto: {CONFIG.get('txt')}")
        logs.append(f"Modelo imagen: {CONFIG.get('img') if CONFIG.get('img') else 'N/A'}")
        logs.append(f"Imagen habilitada: {enable_img}")
        logs.append(f"Grado (planilla): {grado}")
        logs.append(f"Alumnos: {len(df_final)}")
        logs.append("")

        with zipfile.ZipFile(zip_io, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("_REPORTE.txt", "\n".join(logs))

            prog = st.progress(0.0)
            status = st.empty()

            base_hash = hash_text(f"{mode}|{grado}|{input_text}|{inst_style}|{SYSTEM_PROMPT_OPALPP}|{CONFIG.get('txt')}")

            for idx, (_, row) in enumerate(df_final.iterrows(), start=1):
                n = str(row[alumno_col]).strip()
                g = str(row[grupo_col]).strip()
                d = str(row[diag_col]).strip()

                status.info(f"Procesando: {n} ({idx}/{len(df_final)})")

                try:
                    if mode == "CREAR":
                        ctx = f"CREAR ACTIVIDAD DESDE CERO:\n{input_text}\n"
                    else:
                        ctx = f"ADAPTAR CONTENIDO ORIGINAL:\n{input_text}\n"

                    prompt_full = f"""{SYSTEM_PROMPT_OPALPP}

INSTRUCCIONES ON-THE-FLY (prioridad alta):
{inst_style}

CONTEXTO:
{ctx}

ALUMNO (planilla):
- nombre: {n}
- diagnostico: {d}
- grupo: {g}
- grado: {grado}

RECORDATORIOS:
- items[].enunciado inicia con emoji (✍️📖🔢🎨)
- pista_visual = micro-pasos concretos
- sin itálicas
- visual.prompt inicia con "{IMAGE_PROMPT_PREFIX}" si visual.habilitado=true
"""

                    prompt_compact = f"""Devuelve SOLO JSON válido.
Max 6 items. Enunciados cortos con emoji. Pistas micro-pasos. visual false.
Sin itálicas. Usa **negrita** mínimo.

INSTRUCCIONES ON-THE-FLY:
{inst_style}

CONTEXTO:
{ctx}

ALUMNO: {n} | {d} | Grupo {g} | Grado {grado}
"""

                    cache_key = f"{base_hash}::{safe_filename(n)}::{safe_filename(g)}::{safe_filename(d)}"
                    data, mode_used, max_t = request_activity_ultra(CONFIG.get("txt"), prompt_full, prompt_compact, cache_key)

                    # Normalización defensiva
                    items_norm = []
                    for it in (data.get("items", []) or []):
                        if not isinstance(it, dict):
                            continue
                        tipo_i = str(it.get("tipo", "")).strip()
                        en_i = ensure_action_emoji(tipo_i, str(it.get("enunciado", "")).strip())
                        ops = it.get("opciones", [])
                        if not isinstance(ops, list):
                            ops = []
                        pista = str(it.get("pista_visual", "")).strip()
                        items_norm.append({
                            "tipo": tipo_i,
                            "enunciado": en_i,
                            "opciones": [str(x) for x in ops],
                            "pista_visual": pista
                        })
                    data["items"] = items_norm

                    # Visual
                    v = data.get("visual", {}) if isinstance(data.get("visual", {}), dict) else {}
                    v_en = normalize_bool(v.get("habilitado", False))
                    v_pr = normalize_visual_prompt(str(v.get("prompt", "")).strip()) if v_en else ""
                    data["visual"] = {"habilitado": v_en, "prompt": v_pr}

                    # Control calidad
                    data.setdefault("control_calidad", {})
                    if isinstance(data["control_calidad"], dict):
                        data["control_calidad"]["items_count"] = len(data["items"])

                    okj, whyj = validate_activity_json(data)
                    if not okj:
                        raise ValueError(f"JSON final inválido: {whyj}")

                    alumno = {"nombre": n, "diagnostico": d, "grupo": g, "grado": str(grado)}
                    docx_bytes = render_opal_docx(data, alumno, l_bytes, CONFIG.get("img"), enable_img=enable_img)

                    zf.writestr(f"Ficha_{safe_filename(n)}.docx", docx_bytes)
                    zf.writestr(f"_META_{safe_filename(n)}.txt", f"mode={mode_used}\nmax_tokens={max_t}\nitems={len(data.get('items',[]))}\n")
                    ok_count += 1

                except Exception as e:
                    msg = f"{n} :: {type(e).__name__} :: {e}"
                    errors.append(msg)
                    zf.writestr(f"ERROR_{safe_filename(n)}.txt", msg)

                prog.progress(idx / len(df_final))

            resumen = []
            resumen.append("RESUMEN")
            resumen.append(f"Fin: {now_str()}")
            resumen.append(f"OK: {ok_count} / {len(df_final)}")
            resumen.append(f"Errores: {len(errors)}")
            if errors:
                resumen.append("")
                resumen.append("ERRORES (primeros 200):")
                resumen.extend([f"- {e}" for e in errors[:200]])
                if len(errors) > 200:
                    resumen.append(f"... truncado ({len(errors)} errores totales)")
            zf.writestr("_RESUMEN.txt", "\n".join(resumen))

        st.success(f"Lote finalizado. OK: {ok_count} | Errores: {len(errors)}")
        st.download_button("📥 Descargar ZIP", zip_io.getvalue(), "nano_opal_v22_0.zip", mime="application/zip")


if __name__ == "__main__":
    main()
