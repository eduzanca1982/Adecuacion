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
# Nano Opal v24.1
# Cambios vs v24.0:
# - "Sugerencias para el docente" y "Adecuaciones aplicadas" NO se muestran al alumno.
#   Se mueven al Solucionario DOCENTE únicamente.
# - Mantiene imágenes (best-effort) y todo lo robusto de v24.0.
# ============================================================

st.set_page_config(page_title="Nano Opal v24.1 🍌", layout="wide", page_icon="🍌")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

RETRIES = 6
CACHE_TTL_SECONDS = 6 * 60 * 60

MIN_IMAGE_BYTES = 1200
IMAGE_PROMPT_PREFIX = "Pictograma estilo ARASAAC, trazos negros gruesos, fondo blanco, ultra simple, sin sombras de: "

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

SYSTEM_PROMPT_OPAL_V241 = f"""
Actúa como un Senior Inclusive UX Designer y Tutor Psicopedagogo.

Objetivo: producir una ficha de 60 minutos neuroinclusiva (TDAH/dislexia friendly) con estética tipo "Card",
Y producir un solucionario para el docente.

REGLAS NO NEGOCIABLES:
- NO uses markdown. NO uses ** ni __ ni backticks. CERO marcadores de negrita.
- ICONOS: Cada ítem en items[] debe iniciar el enunciado con un emoji de acción:
  ✍️ completar/escribir, 📖 leer, 🔢 calcular, 🎨 dibujar.
- MICRO-PASOS: pista_visual debe ser andamiaje físico/visual, instrucciones concretas. No teoría.
- LENGUAJE: 1 acción por frase, pasos numerados cuando aplique.
- VISUAL: si visual.habilitado=true, visual.prompt debe comenzar EXACTAMENTE con:
  "{IMAGE_PROMPT_PREFIX}[OBJETO]"

SALIDA: JSON puro, sin texto extra.

ESQUEMA EXACTO:
{{
  "objetivo_aprendizaje": "string",
  "tiempo_total_min": 60,
  "consigna_general_alumno": "string (paso a paso, sin saludos)",
  "items": [
    {{
      "tipo": "calcular|lectura|escritura|dibujar|multiple choice|unir|completar|verdadero_falso|problema_guiado",
      "enunciado": "string (DEBE EMPEZAR con emoji de acción)",
      "pasos": ["string","string"],
      "opciones": ["string","string"],
      "respuesta_formato": "texto_corto|procedimiento|dibujo|multiple_choice",
      "keywords_bold": ["string","string"],
      "pista_visual": "string (micro-pasos concretos)",
      "visual": {{ "habilitado": boolean, "prompt": "string" }}
    }}
  ],
  "adecuaciones_aplicadas": ["string","string"],
  "sugerencias_docente": ["string","string"],
  "solucionario_docente": {{
    "respuestas": [
      {{
        "item_index": 1,
        "respuesta_final": "string",
        "desarrollo": ["string","string"],
        "errores_frecuentes": ["string","string"]
      }}
    ],
    "criterios_correccion": ["string","string"]
  }},
  "control_calidad": {{
    "items_count": number,
    "incluye_ejemplo": boolean,
    "lenguaje_concreto": boolean,
    "una_accion_por_frase": boolean,
    "sin_markdown": boolean
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
            return False, "TEXTO vacío tras extracción.", info
        if len(text) < 120:
            return False, "TEXTO muy corto (<120 chars).", info
        return True, "OK", info
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
# Response parsing (TEXT)
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
# Image parsing (best-effort para SDK viejo)
# ============================================================
DATA_URI_RE = re.compile(r"data:image/(png|jpeg|jpg|webp);base64,([A-Za-z0-9+/=\n\r]+)")

def _maybe_b64_to_bytes(x: Any) -> Optional[bytes]:
    if x is None:
        return None
    if isinstance(x, (bytes, bytearray)):
        return bytes(x)
    if isinstance(x, str):
        s = x.strip()
        m = DATA_URI_RE.search(s)
        if m:
            b64 = m.group(2)
            try:
                return base64.b64decode(b64, validate=False)
            except Exception:
                return None
        if len(s) > 400 and all(c in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=\n\r" for c in s[:200]):
            try:
                return base64.b64decode(s, validate=False)
            except Exception:
                return None
    return None

def _extract_inline_bytes_or_none(resp) -> Optional[bytes]:
    try:
        cand = resp.candidates[0]
        content = getattr(cand, "content", None)
        if not content or not getattr(content, "parts", None):
            return None
        for part in content.parts:
            inline = getattr(part, "inline_data", None)
            if inline is not None:
                data = getattr(inline, "data", None)
                b = _maybe_b64_to_bytes(data) or (data if isinstance(data, (bytes, bytearray)) else None)
                if b:
                    return b
            inline2 = getattr(part, "inlineData", None)
            if inline2 is not None:
                data2 = getattr(inline2, "data", None)
                b2 = _maybe_b64_to_bytes(data2)
                if b2:
                    return b2
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
    if b[:8] == b"\x89PNG\r\n\x1a\n":
        return True
    if b[:3] == b"\xff\xd8\xff":
        return True
    if b[:4] == b"RIFF" and b[8:12] == b"WEBP":
        return True
    return False

def generate_image_bytes(model_id: str, prompt_img: str) -> Optional[bytes]:
    if not model_id:
        return None
    prompt_img = normalize_visual_prompt(prompt_img)

    def call_with_cfg(cfg: Optional[Dict[str, Any]]):
        m = genai.GenerativeModel(model_id)
        if cfg is None:
            return m.generate_content(prompt_img, safety_settings=SAFETY_SETTINGS)
        return m.generate_content(prompt_img, generation_config=cfg, safety_settings=SAFETY_SETTINGS)

    cfg_variants = [
        {"response_modalities": ["Image"]},
        {"response_modalities": ["IMAGE"]},
        {"responseModalities": ["Image"]},
        {"responseModalities": ["IMAGE"]},
        None,
    ]

    for cfg in cfg_variants:
        try:
            resp = retry_with_backoff(lambda: call_with_cfg(cfg))
            b = _extract_inline_bytes_or_none(resp)
            if b and _looks_like_image(b):
                return b
        except Exception:
            continue
    return None


# ============================================================
# Boot REAL + selección de modelos
# ============================================================
def list_models_generate_content() -> List[str]:
    models = []
    for m in genai.list_models():
        try:
            if 'generateContent' in getattr(m, "supported_generation_methods", []):
                models.append(m.name)
        except Exception:
            continue
    return models

def rank_text_models(models: List[str], prefer: str) -> List[str]:
    prefer = (prefer or "").strip()
    prios = []
    if prefer:
        prios.append(prefer)
        if not prefer.startswith("models/"):
            prios.append("models/" + prefer)
    prios += [
        "models/gemini-2.5-pro",
        "models/gemini-2.5-flash",
        "models/gemini-2.0-flash",
        "models/gemini-2.0-flash-001",
        "models/gemini-1.5-pro",
        "models/gemini-1.5-flash",
        "models/gemini-pro",
    ]
    ordered = []
    used = set()
    for p in prios:
        for real in models:
            if real in used:
                continue
            if real == p or p in real or (p.startswith("models/") and p.replace("models/", "") in real):
                ordered.append(real)
                used.add(real)
    for real in models:
        if real not in used:
            ordered.append(real)
            used.add(real)
    return ordered

def smoke_test_text_model(model_id: str) -> Tuple[bool, str]:
    try:
        m = genai.GenerativeModel(model_id)
        cfg = {"temperature": 0, "max_output_tokens": 64}
        resp = retry_with_backoff(lambda: m.generate_content("Responde SOLO: OK", generation_config=cfg, safety_settings=SAFETY_SETTINGS))
        t = _extract_text_or_none(resp)
        if not t:
            fr = _finish_reason(resp)
            return False, f"Sin texto (finish_reason={fr})"
        return True, "OK"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

def smoke_test_image_model(model_id: str) -> Tuple[bool, str]:
    try:
        b = generate_image_bytes(model_id, IMAGE_PROMPT_PREFIX + "manzana")
        if not b:
            return False, "No se obtuvo imagen válida (posible incompatibilidad SDK/modelo)"
        return True, f"OK bytes={len(b)}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

def pick_image_fallback(visible: List[str], prefer_img: str) -> Tuple[Optional[str], str]:
    cands = []
    if prefer_img and prefer_img.strip():
        cands.append(prefer_img.strip())
        if not prefer_img.strip().startswith("models/"):
            cands.append("models/" + prefer_img.strip())

    for m in visible:
        ml = m.lower()
        if "imagen" in ml or ml.startswith("models/imagen") or "image-generation" in ml:
            cands.append(m)

    seen = set()
    cands = [x for x in cands if not (x in seen or seen.add(x))]

    last_msg = "No probado"
    for cand in cands:
        ok, msg = smoke_test_image_model(cand)
        last_msg = f"{cand}: {msg}"
        if ok:
            return cand, f"OK {last_msg}"
    return None, f"FAIL {last_msg}"

def boot_pick_models(prefer_text: str, prefer_image: str) -> Dict[str, Any]:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    visible = list_models_generate_content()

    if not visible:
        return {"txt": None, "img": None, "txt_reason": "No hay modelos con generateContent visibles", "img_reason": "", "visible": [], "boot_time": now_str()}

    txt = None
    txt_reason = ""
    for cand in rank_text_models(visible, prefer_text):
        ok, msg = smoke_test_text_model(cand)
        if ok:
            txt = cand
            txt_reason = f"OK: {cand}"
            break
        else:
            txt_reason = f"FAIL {cand}: {msg}"

    img, img_reason = pick_image_fallback(visible, prefer_image)

    return {
        "txt": txt,
        "img": img,
        "txt_reason": txt_reason,
        "img_reason": img_reason,
        "visible": visible[:200],
        "boot_time": now_str(),
    }

@st.cache_resource(show_spinner=False)
def boot_cached(prefer_text: str, prefer_image: str) -> Dict[str, Any]:
    try:
        return boot_pick_models(prefer_text, prefer_image)
    except Exception as e:
        return {"txt": None, "img": None, "txt_reason": f"Boot error: {e}", "img_reason": "", "visible": [], "boot_time": now_str()}


# ============================================================
# JSON validation (v24.x) + repair
# ============================================================
def _contains_markdown_markers(s: str) -> bool:
    if not s:
        return False
    return ("**" in s) or ("```" in s) or ("__" in s)

def validate_activity_json_v241(data: Dict[str, Any]) -> Tuple[bool, str]:
    try:
        if not isinstance(data, dict):
            return False, "Root no es objeto"

        req = [
            "objetivo_aprendizaje", "tiempo_total_min", "consigna_general_alumno",
            "items", "adecuaciones_aplicadas", "sugerencias_docente",
            "solucionario_docente", "control_calidad"
        ]
        for k in req:
            if k not in data:
                return False, f"Falta clave: {k}"

        if data.get("tiempo_total_min") != 60:
            return False, "tiempo_total_min debe ser 60"

        if _contains_markdown_markers(str(data.get("consigna_general_alumno", ""))):
            return False, "contiene marcadores markdown en consigna_general_alumno"

        items = data.get("items", [])
        if not isinstance(items, list) or len(items) < 1:
            return False, "items vacío/no lista"

        cc = data.get("control_calidad", {})
        if not isinstance(cc, dict):
            return False, "control_calidad no es objeto"
        if cc.get("items_count") != len(items):
            return False, "control_calidad.items_count != len(items)"
        if cc.get("sin_markdown") is not True:
            return False, "control_calidad.sin_markdown debe ser true"

        for i, it in enumerate(items[:200]):
            if not isinstance(it, dict):
                return False, f"items[{i}] no es objeto"
            en = str(it.get("enunciado", "")).strip()
            if not en:
                return False, f"items[{i}].enunciado vacío"
            if _contains_markdown_markers(en) or _contains_markdown_markers(str(it.get("pista_visual",""))):
                return False, f"items[{i}] contiene marcadores markdown"
            if not any(en.startswith(x) for x in ["✍️", "📖", "🔢", "🎨"]):
                return False, f"items[{i}].enunciado no inicia con emoji"

            kw = it.get("keywords_bold", [])
            if not isinstance(kw, list):
                return False, f"items[{i}].keywords_bold no es lista"

            v = it.get("visual", {})
            if not isinstance(v, dict):
                return False, f"items[{i}].visual no es objeto"
            if normalize_bool(v.get("habilitado", False)):
                p = str(v.get("prompt", "")).strip()
                if not p.startswith(IMAGE_PROMPT_PREFIX):
                    return False, f"items[{i}].visual.prompt no respeta prefijo ARASAAC"

        sol = data.get("solucionario_docente", {})
        if not isinstance(sol, dict):
            return False, "solucionario_docente no es objeto"
        resp = sol.get("respuestas", [])
        if not isinstance(resp, list) or len(resp) < 1:
            return False, "solucionario_docente.respuestas vacío/no lista"

        return True, "OK"
    except Exception as e:
        return False, f"Exception validando: {e}"

def build_repair_prompt_v241(bad: str, why: str) -> str:
    return f"""
Devuelve EXCLUSIVAMENTE un JSON válido y corregido (sin texto extra).

Problema detectado:
{why}

JSON A CORREGIR:
{bad}

Reglas:
- Prohibido markdown. NO usar ** ni backticks ni __.
- control_calidad.sin_markdown = true
- control_calidad.items_count == len(items)
- items[].enunciado inicia con emoji (✍️📖🔢🎨)
- En vez de negritas, usar keywords_bold[].
- visual.prompt inicia con "{IMAGE_PROMPT_PREFIX}" si visual.habilitado=true
- tiempo_total_min = 60
- El alumno NO debe ver sugerencias_docente ni adecuaciones_aplicadas (igual deben existir en JSON para el DOCENTE).
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

def generate_json_with_repair_v241(model_id: str, prompt: str, max_out: int) -> Dict[str, Any]:
    try:
        data = generate_json_once(model_id, prompt, max_out)
        ok, why = validate_activity_json_v241(data)
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

        repair = build_repair_prompt_v241(raw, f"{type(e).__name__}: {e}")
        resp2 = retry_with_backoff(lambda: m.generate_content(repair, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
        raw2 = _extract_text_or_none(resp2)
        fr2 = _finish_reason(resp2)
        if raw2 is None:
            raise ValueError(f"Empty candidate after repair (finish_reason={fr2})")

        data2 = json.loads(raw2)
        ok2, why2 = validate_activity_json_v241(data2)
        if not ok2:
            raise ValueError(f"JSON reparado inválido: {why2}")
        return data2

@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def cached_generate_v241(cache_key: str, model_id: str, prompt: str, max_out: int) -> Dict[str, Any]:
    return generate_json_with_repair_v241(model_id, prompt, max_out)

def request_activity_ultra_v241(model_id: str, prompt_full: str, prompt_compact: str, cache_key: str) -> Tuple[Dict[str, Any], str, int]:
    last_err = None
    for t in OUT_TOKEN_STEPS_FULL:
        try:
            data = cached_generate_v241(cache_key + f"::FULL::{t}", model_id, prompt_full, t)
            return data, "FULL", t
        except Exception as e:
            last_err = e
    for t in OUT_TOKEN_STEPS_COMPACT:
        try:
            data = cached_generate_v241(cache_key + f"::COMPACT::{t}", model_id, prompt_compact, t)
            return data, "COMPACT", t
        except Exception as e:
            last_err = e
    raise last_err if last_err else RuntimeError("Fallo desconocido generando actividad")


# ============================================================
# DOCX rendering (Opal cards)
# ============================================================
def apply_card_style(cell, fill_hex: str = "FAFAFA"):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tc_pr.append(shd)

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

def add_text(paragraph, text: str, bold: bool = False, color: Optional[RGBColor] = None, size_pt: int = 14):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = False
    run.font.name = "Verdana"
    run.font.size = Pt(size_pt)
    if color is not None:
        run.font.color.rgb = color
    return run

def add_text_with_keywords(paragraph, text: str, keywords: List[str], size_pt: int = 14):
    text = str(text or "")
    kws = [k.strip() for k in (keywords or []) if isinstance(k, str) and k.strip()]
    if not kws:
        add_text(paragraph, text, bold=False, size_pt=size_pt)
        return

    kws_sorted = sorted(set(kws), key=len, reverse=True)
    pat = re.compile("(" + "|".join(re.escape(k) for k in kws_sorted) + ")")

    parts = pat.split(text)
    for part in parts:
        if part in kws_sorted:
            add_text(paragraph, part, bold=True, size_pt=size_pt)
        else:
            add_text(paragraph, part, bold=False, size_pt=size_pt)

def response_box(cell, label: str = "✍️ Respuesta:", lines: int = 4):
    t = cell.add_table(rows=1, cols=1)
    c = t.rows[0].cells[0]
    apply_card_style(c, fill_hex="FFFFFF")
    clear_paragraph(c.paragraphs[0])

    p = c.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    add_text(p, label + " ", bold=True)

    p2 = c.add_paragraph()
    p2.paragraph_format.line_spacing = 1.5
    add_text(p2, "\n" + ("\n" * max(0, lines - 1)) + " ", bold=False)

def checkbox_list(cell, options: List[str], max_opts: int = 8):
    for opt in (options or [])[:max_opts]:
        p = cell.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        add_text(p, f"☐ {opt}", bold=False)

def header_block(doc: Document, alumno: Dict[str, str], logo_b: Optional[bytes], title: str):
    style = doc.styles['Normal']
    style.font.name = 'Verdana'
    style.font.size = Pt(14)

    h = doc.add_table(rows=1, cols=2)
    h.width = Inches(6.5)
    if logo_b:
        try:
            h.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_b), width=Inches(0.7))
        except Exception:
            pass
    info = h.rows[0].cells[1].paragraphs[0]
    info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_text(info, title + "\n", bold=True, size_pt=12)
    add_text(info, f"{alumno.get('nombre','')}\n", bold=True, size_pt=11)
    add_text(info, f"{alumno.get('diagnostico','')}\n", bold=False, size_pt=11)
    add_text(info, f"Grupo: {alumno.get('grupo','')} | Grado: {alumno.get('grado','')}", bold=False, size_pt=11)

    doc.add_paragraph("")


def render_alumno_docx(data: Dict[str, Any], alumno: Dict[str, str], logo_b: Optional[bytes], img_model_id: Optional[str], enable_img: bool) -> bytes:
    doc = Document()
    header_block(doc, alumno, logo_b, "FICHA DEL ALUMNO")

    # Objetivo
    p = doc.add_paragraph()
    add_text(p, "Objetivo de aprendizaje", bold=True)
    p2 = doc.add_paragraph()
    p2.paragraph_format.line_spacing = 1.5
    add_text(p2, str(data.get("objetivo_aprendizaje","")), bold=False)

    doc.add_paragraph("")

    # Consigna general
    p = doc.add_paragraph()
    add_text(p, "Consigna general (paso a paso)", bold=True)
    cg = str(data.get("consigna_general_alumno","")).strip()
    for line in [x.strip() for x in cg.split("\n") if x.strip()]:
        p3 = doc.add_paragraph()
        p3.paragraph_format.line_spacing = 1.5
        add_text(p3, line, bold=False)

    doc.add_paragraph("")

    # Cards por ítem
    for idx, it in enumerate(data.get("items", []), start=1):
        if not isinstance(it, dict):
            continue

        tipo = str(it.get("tipo","")).strip()
        enunciado = ensure_action_emoji(tipo, str(it.get("enunciado","")).strip())
        pasos = it.get("pasos", []) if isinstance(it.get("pasos", []), list) else []
        opciones = it.get("opciones", []) if isinstance(it.get("opciones", []), list) else []
        formato = str(it.get("respuesta_formato","texto_corto")).strip()
        kw = it.get("keywords_bold", []) if isinstance(it.get("keywords_bold", []), list) else []
        pista = str(it.get("pista_visual","")).strip()

        v = it.get("visual", {}) if isinstance(it.get("visual", {}), dict) else {}
        v_en = normalize_bool(v.get("habilitado", False))
        v_pr = normalize_visual_prompt(str(v.get("prompt", "")).strip()) if v_en else ""

        card = doc.add_table(rows=1, cols=1)
        card.width = Inches(6.5)
        cell = card.rows[0].cells[0]
        apply_card_style(cell, fill_hex="FAFAFA")
        clear_paragraph(cell.paragraphs[0])

        # Título ítem
        pt = cell.add_paragraph()
        pt.paragraph_format.line_spacing = 1.5
        add_text(pt, f"Ítem {idx}", bold=True, size_pt=12)

        # Consigna
        p_con = cell.add_paragraph()
        p_con.paragraph_format.line_spacing = 1.6
        add_text_with_keywords(p_con, enunciado, kw, size_pt=14)

        # Pasos
        if pasos:
            for i, step in enumerate(pasos[:8], start=1):
                ps = cell.add_paragraph()
                ps.paragraph_format.line_spacing = 1.5
                add_text(ps, f"{i}. {str(step)}", bold=False)

        # Trabajo
        sep = cell.add_paragraph()
        add_text(sep, "Trabajo", bold=True, size_pt=12)

        if opciones and (formato.lower() in {"multiple_choice", "multiple choice"}):
            checkbox_list(cell, [str(x) for x in opciones], max_opts=8)
        else:
            response_box(cell, label="✍️ Respuesta:", lines=4)

        # Pista (verde)
        if pista:
            pp = cell.add_paragraph()
            pp.paragraph_format.line_spacing = 1.5
            add_text(pp, "Pista", bold=True, size_pt=12)
            pp2 = cell.add_paragraph()
            pp2.paragraph_format.line_spacing = 1.5
            add_text(pp2, "💡 " + pista, bold=False, color=RGBColor(0, 150, 0), size_pt=14)

        # Imagen por ítem (si existe)
        if enable_img and img_model_id and v_en and v_pr:
            img_bytes = generate_image_bytes(img_model_id, v_pr)
            if img_bytes:
                pi = cell.add_paragraph()
                add_text(pi, "Apoyo visual", bold=True, size_pt=12)
                pic = cell.add_paragraph()
                pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
                try:
                    pic.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(2.2))
                except Exception:
                    pass

        doc.add_paragraph("")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def render_docente_docx(data: Dict[str, Any], alumno: Dict[str, str], logo_b: Optional[bytes]) -> bytes:
    doc = Document()
    header_block(doc, alumno, logo_b, "SOLUCIONARIO DOCENTE")

    # Resumen objetivo
    p = doc.add_paragraph()
    add_text(p, "Objetivo de aprendizaje", bold=True)
    p2 = doc.add_paragraph()
    p2.paragraph_format.line_spacing = 1.5
    add_text(p2, str(data.get("objetivo_aprendizaje","")), bold=False)

    doc.add_paragraph("")

    sol = data.get("solucionario_docente", {}) if isinstance(data.get("solucionario_docente", {}), dict) else {}
    respuestas = sol.get("respuestas", []) if isinstance(sol.get("respuestas", []), list) else []

    by_idx = {}
    for r in respuestas:
        if not isinstance(r, dict):
            continue
        idx = int(r.get("item_index", 0) or 0)
        by_idx[idx] = r

    for idx, it in enumerate(data.get("items", []), start=1):
        if not isinstance(it, dict):
            continue

        en = str(it.get("enunciado","")).strip()
        r = by_idx.get(idx, {})

        card = doc.add_table(rows=1, cols=1)
        card.width = Inches(6.5)
        cell = card.rows[0].cells[0]
        apply_card_style(cell, fill_hex="FFFFFF")
        clear_paragraph(cell.paragraphs[0])

        pt = cell.add_paragraph()
        pt.paragraph_format.line_spacing = 1.4
        add_text(pt, f"Ítem {idx}", bold=True, size_pt=12)

        pe = cell.add_paragraph()
        pe.paragraph_format.line_spacing = 1.5
        add_text(pe, en, bold=True)

        pf = cell.add_paragraph()
        pf.paragraph_format.line_spacing = 1.5
        add_text(pf, "Respuesta final: ", bold=True)
        add_text(pf, str(r.get("respuesta_final","(no provista)")), bold=False)

        des = r.get("desarrollo", []) if isinstance(r.get("desarrollo", []), list) else []
        if des:
            pd = cell.add_paragraph()
            add_text(pd, "Desarrollo:", bold=True, size_pt=12)
            for step in des[:12]:
                ps = cell.add_paragraph()
                ps.paragraph_format.line_spacing = 1.5
                add_text(ps, f"• {step}", bold=False)

        ef = r.get("errores_frecuentes", []) if isinstance(r.get("errores_frecuentes", []), list) else []
        if ef:
            pef = cell.add_paragraph()
            add_text(pef, "Errores frecuentes:", bold=True, size_pt=12)
            for e in ef[:8]:
                pex = cell.add_paragraph()
                pex.paragraph_format.line_spacing = 1.5
                add_text(pex, f"• {e}", bold=False)

        doc.add_paragraph("")

    crit = sol.get("criterios_correccion", []) if isinstance(sol.get("criterios_correccion", []), list) else []
    if crit:
        p = doc.add_paragraph()
        add_text(p, "Criterios de corrección", bold=True)
        for c in crit[:15]:
            p2 = doc.add_paragraph()
            p2.paragraph_format.line_spacing = 1.5
            add_text(p2, f"• {c}", bold=False)

    # SOLO DOCENTE: Adecuaciones aplicadas + sugerencias
    adec = data.get("adecuaciones_aplicadas", []) if isinstance(data.get("adecuaciones_aplicadas", []), list) else []
    if adec:
        doc.add_paragraph("")
        p = doc.add_paragraph()
        add_text(p, "Adecuaciones aplicadas", bold=True)
        for a in adec[:20]:
            p2 = doc.add_paragraph()
            p2.paragraph_format.line_spacing = 1.5
            add_text(p2, f"• {a}", bold=False)

    sug = data.get("sugerencias_docente", []) if isinstance(data.get("sugerencias_docente", []), list) else []
    if sug:
        doc.add_paragraph("")
        p = doc.add_paragraph()
        add_text(p, "Sugerencias para el docente", bold=True)
        for s in sug[:20]:
            p2 = doc.add_paragraph()
            p2.paragraph_format.line_spacing = 1.5
            add_text(p2, f"• {s}", bold=False)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# UI + Proceso
# ============================================================
def main():
    st.title("Nano Opal v24.1 🧠🍌")
    st.caption("Alumno NO ve sugerencias/adecuaciones. Eso va al Solucionario DOCENTE. Imágenes OK.")

    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando planilla: {e}")
        return

    grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
    grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
    diag_col = df.columns[4] if len(df.columns) > 4 else df.columns[0]

    with st.sidebar:
        st.header("⚙️ Modelos (boot real)")

        prefer_txt = st.text_input("Preferido texto", value="gemini-1.5-flash")
        prefer_img = st.text_input("Preferido imagen", value="gemini-2.5-flash-image")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("Reboot"):
                st.cache_resource.clear()
        with c2:
            if st.button("Limpiar cache"):
                st.cache_data.clear()

        try:
            genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
        except Exception as e:
            st.error(f"API Key inválida o faltante: {e}")
            return

        CONFIG = boot_cached(prefer_txt, prefer_img)
        st.write(f"Boot: {CONFIG.get('boot_time')}")
        if CONFIG.get("txt"):
            st.success(f"Texto: {CONFIG.get('txt')}")
        else:
            st.error("Texto: N/A")
        st.caption(CONFIG.get("txt_reason",""))

        if CONFIG.get("img"):
            st.success(f"Imagen: {CONFIG.get('img')}")
        else:
            st.warning("Imagen: desactivada (SDK/modelo no entregó bytes)")
        st.caption(CONFIG.get("img_reason",""))

        with st.expander("Modelos visibles (generateContent)", expanded=False):
            for m in (CONFIG.get("visible", []) or []):
                st.write(m)

        st.divider()

        st.header("📚 Grado / Alumnos (Sheets)")
        grado = st.selectbox("Grado", sorted(df[grado_col].dropna().unique().tolist()))
        df_f = df[df[grado_col] == grado].copy()

        alcance = st.radio("Alcance", ["Todo el grado", "Seleccionar alumnos"], horizontal=True)
        if alcance == "Seleccionar alumnos":
            al_sel = st.multiselect("Alumnos", sorted(df_f[alumno_col].dropna().unique().tolist()))
            df_final = df_f[df_f[alumno_col].isin(al_sel)].copy() if al_sel else df_f.iloc[0:0].copy()
        else:
            df_final = df_f

        st.divider()
        enable_img = st.checkbox("Habilitar imágenes", value=True)
        enable_img = enable_img and bool(CONFIG.get("img"))

        logo = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
        l_bytes = logo.read() if logo else None

        st.divider()
        inst_style = st.text_area("Instrucciones de Estilo On-the-fly", height=120)

    if not CONFIG.get("txt"):
        st.error("No hay modelo de texto funcional.")
        return

    tab1, tab2 = st.tabs(["🔄 Adaptar DOCX", "✨ Crear Actividad"])

    adapt_docx = None
    brief = ""

    with tab1:
        st.subheader("Adaptar (DOCX)")
        adapt_docx = st.file_uploader("Actividad base (DOCX)", type=["docx"], key="docx_in")

    with tab2:
        st.subheader("Crear desde prompt")
        brief = st.text_area(
            "Prompt",
            height=220,
            placeholder="Ej: Diseña una actividad de 60 minutos para 7mo grado sobre Proporcionalidad Directa aplicada a escalas..."
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
            st.info("Subí un DOCX o usa 'Crear Actividad'.")
    else:
        input_text = brief.strip()
        ok_in, msg_in, info_in = validate_text_input(input_text, "CREAR")
        if ok_in:
            st.success(f"Prompt OK ({info_in['chars']} chars)")
        else:
            st.error(f"Prompt: {msg_in}")
        with st.expander("Preview prompt", expanded=False):
            st.text(info_in.get("preview", ""))

    if st.button("🚀 GENERAR LOTE"):
        if len(df_final) == 0:
            st.error("No hay alumnos (ver selección por grado/alumnos).")
            return
        if mode == "ADAPTAR" and not adapt_docx:
            st.error("Falta DOCX para adaptar.")
            return
        ok_in, msg_in, _ = validate_text_input(input_text, mode)
        if not ok_in:
            st.error(f"No se inicia: {msg_in}")
            return

        ok_sm, msg_sm = smoke_test_text_model(CONFIG["txt"])
        if not ok_sm:
            st.error(f"Modelo texto no responde: {msg_sm}")
            return

        zip_io = io.BytesIO()
        logs = []
        errors = []
        ok_count = 0

        logs.append("Nano Opal v24.1")
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

            base_hash = hash_text(f"{mode}|{grado}|{input_text}|{inst_style}|{SYSTEM_PROMPT_OPAL_V241}|{CONFIG.get('txt')}")

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

                    prompt_full = f"""{SYSTEM_PROMPT_OPAL_V241}

INSTRUCCIONES ON-THE-FLY (prioridad alta):
{inst_style}

CONTEXTO:
{ctx}

ALUMNO (planilla):
- nombre: {n}
- diagnostico: {d}
- grupo: {g}
- grado: {grado}

NOTAS:
- NO usar markdown. NO usar ** en ningún campo.
- La ficha del alumno NO debe incluir sugerencias_docente ni adecuaciones_aplicadas; esas son SOLO para el solucionario.
"""

                    prompt_compact = f"""Devuelve SOLO JSON válido del esquema.
Max 6 items. Sin markdown. keywords_bold[] corto. visual.habilitado=true con prompt ARASAAC.
tiempo_total_min=60. solucionario_docente incluido.

INSTRUCCIONES ON-THE-FLY:
{inst_style}

CONTEXTO:
{ctx}

ALUMNO: {n} | {d} | Grupo {g} | Grado {grado}
"""

                    cache_key = f"{base_hash}::{safe_filename(n)}::{safe_filename(g)}::{safe_filename(d)}"
                    data, mode_used, max_t = request_activity_ultra_v241(CONFIG["txt"], prompt_full, prompt_compact, cache_key)

                    data["tiempo_total_min"] = 60

                    items_norm = []
                    for it in (data.get("items", []) or []):
                        if not isinstance(it, dict):
                            continue
                        tipo_i = str(it.get("tipo", "")).strip()
                        en_i = ensure_action_emoji(tipo_i, str(it.get("enunciado", "")).strip())
                        pasos = it.get("pasos", []) if isinstance(it.get("pasos", []), list) else []
                        ops = it.get("opciones", []) if isinstance(it.get("opciones", []), list) else []
                        rf = str(it.get("respuesta_formato", "texto_corto")).strip()
                        kw = it.get("keywords_bold", []) if isinstance(it.get("keywords_bold", []), list) else []
                        pista = str(it.get("pista_visual", "")).strip()

                        v = it.get("visual", {}) if isinstance(it.get("visual", {}), dict) else {}
                        v_en = normalize_bool(v.get("habilitado", False))
                        v_pr = normalize_visual_prompt(str(v.get("prompt", "")).strip()) if v_en else ""

                        items_norm.append({
                            "tipo": tipo_i,
                            "enunciado": en_i,
                            "pasos": [str(x) for x in pasos[:10]],
                            "opciones": [str(x) for x in ops[:10]],
                            "respuesta_formato": rf,
                            "keywords_bold": [str(x) for x in kw[:12]],
                            "pista_visual": pista,
                            "visual": {"habilitado": v_en, "prompt": v_pr}
                        })
                    data["items"] = items_norm

                    data.setdefault("control_calidad", {})
                    if isinstance(data["control_calidad"], dict):
                        data["control_calidad"]["items_count"] = len(data["items"])
                        data["control_calidad"]["sin_markdown"] = True

                    okj, whyj = validate_activity_json_v241(data)
                    if not okj:
                        raise ValueError(f"JSON final inválido: {whyj}")

                    alumno = {"nombre": n, "diagnostico": d, "grupo": g, "grado": str(grado)}

                    doc_alumno = render_alumno_docx(data, alumno, l_bytes, CONFIG.get("img"), enable_img=enable_img)
                    doc_docente = render_docente_docx(data, alumno, l_bytes)

                    zf.writestr(f"Ficha_ALUMNO_{safe_filename(n)}.docx", doc_alumno)
                    zf.writestr(f"Solucionario_DOCENTE_{safe_filename(n)}.docx", doc_docente)
                    zf.writestr(f"_META_{safe_filename(n)}.txt", f"mode={mode_used}\nmax_tokens={max_t}\nitems={len(data.get('items',[]))}\nimg_enabled={enable_img}\nimg_model={CONFIG.get('img')}\n")
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
            zf.writestr("_RESUMEN.txt", "\n".join(resumen))

        st.success(f"Lote finalizado. OK: {ok_count} | Errores: {len(errors)}")
        st.download_button("📥 Descargar ZIP", zip_io.getvalue(), "nano_opal_v24_1.zip", mime="application/zip")


if __name__ == "__main__":
    main()
