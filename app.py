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
# Nano Opal v24.3 (anti-fragile)
# Cambios solicitados:
# - Modelo texto fijo: gemini-2.5-flash (sin selector)
# - Sin menú lateral para elegir modelos (ni "on the fly")
# - NO perder jamás grupo + dificultad/diagnóstico por alumno (sheet)
# - Anti-fallo JSON: si falta una clave, se AUTORELLENA (no se aborta)
# ============================================================

st.set_page_config(page_title="Nano Opal v24.3 🍌", layout="wide", page_icon="🍌")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

DEFAULT_TEXT_MODEL = "models/gemini-2.5-flash"
# best-effort: si tu SDK/modelo no devuelve bytes de imagen, se desactiva sin romper.
DEFAULT_IMAGE_MODEL = "models/gemini-2.5-flash-image"

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

OUT_TOKEN_STEPS = [4096, 6144, 8192]

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

_ALLOWED_TIPOS = {
    "calcular", "lectura", "escritura", "dibujar",
    "multiple choice", "multiple_choice",
    "unir", "completar", "verdadero_falso", "problema_guiado"
}
_ALLOWED_RESP_FMT = {"texto_corto", "procedimiento", "dibujo", "multiple_choice"}

SYSTEM_PROMPT_OPAL_V243 = f"""
Actúa como un Senior Inclusive UX Designer y Tutor Psicopedagogo.

Objetivo: producir una ficha de 60 minutos neuroinclusiva (TDAH/dislexia friendly) con estética tipo "Card"
y un solucionario para el docente.

REGLAS:
- SALIDA: JSON puro, sin texto extra.
- NO uses markdown. NO uses ** ni __ ni backticks.
- Cada ítem debe iniciar el enunciado con emoji de acción:
  ✍️ escribir, 📖 leer, 🔢 calcular, 🎨 dibujar.
- pista_visual: micro-pasos concretos (no teoría).
- Si visual.habilitado=true, visual.prompt debe iniciar EXACTAMENTE con:
  "{IMAGE_PROMPT_PREFIX}[OBJETO]"

ESQUEMA (rellenar siempre):
{{
  "objetivo_aprendizaje": "string",
  "tiempo_total_min": 60,
  "consigna_general_alumno": "string",
  "items": [
    {{
      "tipo": "calcular|lectura|escritura|dibujar|multiple choice|unir|completar|verdadero_falso|problema_guiado",
      "enunciado": "string",
      "pasos": ["string"],
      "opciones": ["string"],
      "respuesta_formato": "texto_corto|procedimiento|dibujo|multiple_choice",
      "keywords_bold": ["string"],
      "pista_visual": "string",
      "visual": {{ "habilitado": boolean, "prompt": "string" }}
    }}
  ],
  "adecuaciones_aplicadas": ["string"],
  "sugerencias_docente": ["string"],
  "solucionario_docente": {{
    "respuestas": [
      {{
        "item_index": 1,
        "respuesta_final": "string",
        "desarrollo": ["string"],
        "errores_frecuentes": ["string"]
      }}
    ],
    "criterios_correccion": ["string"]
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


def safe_filename(name: str) -> str:
    s = str(name).strip().replace(" ", "_")
    for ch in ["/", "\\", ":", "*", "?", "\"", "<", ">", "|"]:
        s = s.replace(ch, "_")
    while "__" in s:
        s = s.replace("__", "_")
    return (s or "SIN_NOMBRE")[:120]


def hash_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()


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


def _as_str_list(v: Any, max_items: int) -> List[str]:
    if isinstance(v, list):
        out = []
        for x in v:
            if x is None:
                continue
            s = str(x).strip()
            if s:
                out.append(s)
        return out[:max_items]
    if v is None:
        return []
    s = str(v).strip()
    return [s][:max_items] if s else []


# ============================================================
# JSON sanitization + parsing (robusto)
# ============================================================
_JSON_CODEFENCE_RE = re.compile(r"```(?:json)?\s*([\s\S]*?)\s*```", re.IGNORECASE)
_TRAILING_COMMA_OBJ_RE = re.compile(r",\s*}")
_TRAILING_COMMA_ARR_RE = re.compile(r",\s*]")
_SINGLELINE_COMMENT_RE = re.compile(r"//.*?$", re.MULTILINE)
_BLOCK_COMMENT_RE = re.compile(r"/\*[\s\S]*?\*/", re.MULTILINE)
_UNQUOTED_KEY_RE = re.compile(r'(\{|,)\s*([A-Za-z_][A-Za-z0-9_ ]{0,60}?)\s*:')


def _strip_wrappers(raw: str) -> str:
    if not raw:
        return raw
    m = _JSON_CODEFENCE_RE.search(raw)
    if m:
        raw = m.group(1)
    start = raw.find("{")
    end = raw.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return raw.strip()
    return raw[start:end + 1].strip()


def _fix_unquoted_keys(s: str) -> str:
    def repl(m):
        prefix = m.group(1)
        key = m.group(2).strip()
        key = re.sub(r"\s+", "_", key)
        return f'{prefix} "{key}":'
    return _UNQUOTED_KEY_RE.sub(repl, s)


def sanitize_json_text(raw_text: str) -> str:
    if not raw_text:
        return raw_text
    s = raw_text
    s = _strip_wrappers(s)
    s = _BLOCK_COMMENT_RE.sub("", s)
    s = _SINGLELINE_COMMENT_RE.sub("", s)
    s = _TRAILING_COMMA_OBJ_RE.sub("}", s)
    s = _TRAILING_COMMA_ARR_RE.sub("]", s)
    s = _fix_unquoted_keys(s)
    return s.strip()


def safe_json_loads(raw_text: str) -> Dict[str, Any]:
    if not raw_text:
        raise ValueError("Empty JSON text")
    cleaned = sanitize_json_text(raw_text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        cleaned2 = _strip_wrappers(cleaned)
        cleaned2 = _TRAILING_COMMA_OBJ_RE.sub("}", cleaned2)
        cleaned2 = _TRAILING_COMMA_ARR_RE.sub("]", cleaned2)
        cleaned2 = _fix_unquoted_keys(cleaned2)
        return json.loads(cleaned2)


def build_parse_fix_prompt(raw: str, err: str) -> str:
    return f"""
Tu única tarea es convertir el siguiente texto en JSON válido.
No agregues texto. No agregues comentarios.
No cambies el contenido semántico, solo corrige sintaxis JSON.
Reglas: keys con comillas dobles, sin trailing commas, valores string con comillas dobles.

ERROR DE PARSEO:
{err}

TEXTO A CONVERTIR:
{raw}
""".strip()


def build_repair_prompt(bad: str, why: str) -> str:
    return f"""
Devuelve EXCLUSIVAMENTE un JSON válido del esquema.
IMPORTANTE: Si falta una clave, agrégala. No elimines claves requeridas.
No uses markdown.

Problema:
{why}

JSON A CORREGIR:
{bad}
""".strip()


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
            for row in doc.element.body.findall(f".//{W_NS}tr"):
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
# Images (best-effort)
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
# Payload normalization (NO FAIL)
# ============================================================
def normalize_activity_payload(data: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(data, dict):
        data = {}

    data.setdefault("objetivo_aprendizaje", "")
    data.setdefault("consigna_general_alumno", "")
    data["tiempo_total_min"] = 60

    data["objetivo_aprendizaje"] = str(data.get("objetivo_aprendizaje", "")).strip()
    data["consigna_general_alumno"] = str(data.get("consigna_general_alumno", "")).strip()

    data["adecuaciones_aplicadas"] = _as_str_list(data.get("adecuaciones_aplicadas", []), 40)
    data["sugerencias_docente"] = _as_str_list(data.get("sugerencias_docente", []), 40)

    items = data.get("items", [])
    if not isinstance(items, list):
        items = []

    items_norm: List[Dict[str, Any]] = []
    for it in items[:60]:
        if not isinstance(it, dict):
            continue

        tipo = str(it.get("tipo", "")).strip().lower()
        if tipo not in _ALLOWED_TIPOS:
            tipo = "lectura"

        en = ensure_action_emoji(tipo, str(it.get("enunciado", "")).strip())
        pasos = _as_str_list(it.get("pasos", []), 12)
        opciones = _as_str_list(it.get("opciones", []), 12)

        resp_fmt = str(it.get("respuesta_formato", "")).strip()
        if resp_fmt not in _ALLOWED_RESP_FMT:
            resp_fmt = "multiple_choice" if opciones else "texto_corto"

        kw = _as_str_list(it.get("keywords_bold", []), 12)
        pista = str(it.get("pista_visual", "")).strip()

        v = it.get("visual", {})
        if not isinstance(v, dict):
            v = {}
        v_en = normalize_bool(v.get("habilitado", False))
        v_pr = normalize_visual_prompt(str(v.get("prompt", "")).strip()) if v_en else ""
        if v_en and not v_pr:
            v_pr = IMAGE_PROMPT_PREFIX + "objeto"

        items_norm.append({
            "tipo": tipo,
            "enunciado": en,
            "pasos": pasos,
            "opciones": opciones,
            "respuesta_formato": resp_fmt,
            "keywords_bold": kw,
            "pista_visual": pista,
            "visual": {"habilitado": bool(v_en), "prompt": v_pr},
        })

    data["items"] = items_norm

    sol = data.get("solucionario_docente", {})
    if not isinstance(sol, dict):
        sol = {}
    resp_list = sol.get("respuestas", [])
    if not isinstance(resp_list, list):
        resp_list = []

    resp_norm: List[Dict[str, Any]] = []
    for r in resp_list[:120]:
        if not isinstance(r, dict):
            continue
        try:
            idx = int(r.get("item_index", 0) or 0)
        except Exception:
            idx = 0
        if idx <= 0:
            continue
        resp_norm.append({
            "item_index": idx,
            "respuesta_final": str(r.get("respuesta_final", "")).strip() or "(no provista)",
            "desarrollo": _as_str_list(r.get("desarrollo", []), 24),
            "errores_frecuentes": _as_str_list(r.get("errores_frecuentes", []), 16),
        })

    sol["respuestas"] = resp_norm
    sol["criterios_correccion"] = _as_str_list(sol.get("criterios_correccion", []), 24)
    data["solucionario_docente"] = sol

    cc = data.get("control_calidad", {})
    if not isinstance(cc, dict):
        cc = {}
    cc["items_count"] = len(items_norm)
    cc["sin_markdown"] = True
    cc.setdefault("incluye_ejemplo", True)
    cc.setdefault("lenguaje_concreto", True)
    cc.setdefault("una_accion_por_frase", True)
    data["control_calidad"] = cc

    return data


def autofill_required_fields(data: Dict[str, Any], brief: str, alumno: Dict[str, str]) -> Dict[str, Any]:
    """
    CERO tolerancia a fallos por claves faltantes.
    Si falta cualquier pieza, se completa con defaults coherentes.
    """
    data = normalize_activity_payload(data)

    if not data.get("objetivo_aprendizaje"):
        data["objetivo_aprendizaje"] = f"Trabajar el tema del día según el grado, ajustado al grupo {alumno.get('grupo','')}."

    if not data.get("consigna_general_alumno"):
        data["consigna_general_alumno"] = (
            "1. Lee cada consigna.\n"
            "2. Haz un paso por vez.\n"
            "3. Revisa antes de terminar."
        )

    if not data.get("items"):
        # fallback mínimo (para NO abortar lote)
        data["items"] = [{
            "tipo": "lectura",
            "enunciado": "📖 Lee el texto y responde con una frase.",
            "pasos": ["Lee 2 veces.", "Responde con 1 frase."],
            "opciones": [],
            "respuesta_formato": "texto_corto",
            "keywords_bold": ["lee", "frase"],
            "pista_visual": "Subraya 2 palabras clave.",
            "visual": {"habilitado": False, "prompt": ""},
        }]
        data = normalize_activity_payload(data)

    # Solucionario: asegurar 1:1 por items
    sol = data.get("solucionario_docente", {})
    by_idx = {}
    for r in sol.get("respuestas", []):
        by_idx[int(r.get("item_index", 0) or 0)] = r

    respuestas = []
    for i in range(1, len(data["items"]) + 1):
        r = by_idx.get(i)
        if not r:
            respuestas.append({
                "item_index": i,
                "respuesta_final": "(modelo no proveyó respuesta; corregir según criterios)",
                "desarrollo": ["Verifica pasos del alumno.", "Evalúa si cumple la consigna."],
                "errores_frecuentes": ["No sigue los pasos.", "Responde incompleto."],
            })
        else:
            respuestas.append(r)

    sol["respuestas"] = respuestas
    if not sol.get("criterios_correccion"):
        sol["criterios_correccion"] = [
            "Cumple la consigna.",
            "Un paso por vez.",
            "Respuesta legible y completa."
        ]
    data["solucionario_docente"] = sol

    data["adecuaciones_aplicadas"] = _as_str_list(data.get("adecuaciones_aplicadas", []), 40) or [
        f"Consignas cortas para {alumno.get('diagnostico','perfil del alumno')}.",
        "Un paso por frase.",
        "Apoyo visual opcional por ítem."
    ]
    data["sugerencias_docente"] = _as_str_list(data.get("sugerencias_docente", []), 40) or [
        "Modelar 1 ejemplo antes de iniciar.",
        "Reforzar rutina: leer → hacer → revisar.",
        "Dar pausa breve cada 10–12 minutos."
    ]

    data["control_calidad"]["items_count"] = len(data["items"])
    data["control_calidad"]["sin_markdown"] = True
    data["tiempo_total_min"] = 60
    return data


# ============================================================
# LLM: generate JSON with parse-fix + repair (pero SIN abortar)
# ============================================================
def generate_json_raw(model_id: str, prompt: str, max_out: int) -> Tuple[str, Optional[int]]:
    m = genai.GenerativeModel(model_id)
    cfg = dict(BASE_GEN_CFG_JSON)
    cfg["max_output_tokens"] = max_out
    resp = retry_with_backoff(lambda: m.generate_content(prompt, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
    text = _extract_text_or_none(resp)
    fr = _finish_reason(resp)
    if text is None:
        return "", fr
    return text, fr


def robust_generate_activity(model_id: str, prompt: str) -> Tuple[Dict[str, Any], str]:
    """
    Devuelve (data, path_debug).
    path_debug indica qué ruta se usó: DIRECT | PARSE_FIX | REPAIR | PARTIAL_FALLBACK
    """
    last_err = None
    for t in OUT_TOKEN_STEPS:
        try:
            raw, fr = generate_json_raw(model_id, prompt, t)
            if not raw:
                last_err = ValueError(f"Empty text (finish_reason={fr})")
                continue
            try:
                return safe_json_loads(raw), f"DIRECT@{t}"
            except Exception as pe:
                # parse-fix
                try:
                    m = genai.GenerativeModel(model_id)
                    cfg = dict(BASE_GEN_CFG_JSON)
                    cfg["max_output_tokens"] = t
                    fix_prompt = build_parse_fix_prompt(raw, f"{type(pe).__name__}: {pe}")
                    resp_fix = retry_with_backoff(lambda: m.generate_content(fix_prompt, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
                    raw_fix = _extract_text_or_none(resp_fix) or ""
                    return safe_json_loads(raw_fix), f"PARSE_FIX@{t}"
                except Exception as pe2:
                    # repair
                    try:
                        m = genai.GenerativeModel(model_id)
                        cfg = dict(BASE_GEN_CFG_JSON)
                        cfg["max_output_tokens"] = t
                        repair_prompt = build_repair_prompt(raw, f"{type(pe2).__name__}: {pe2}")
                        resp_rep = retry_with_backoff(lambda: m.generate_content(repair_prompt, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
                        raw_rep = _extract_text_or_none(resp_rep) or ""
                        return safe_json_loads(raw_rep), f"REPAIR@{t}"
                    except Exception as pe3:
                        last_err = pe3
                        continue
        except Exception as e:
            last_err = e
            continue

    # jamás abortar por JSON: fallback parcial
    return {}, f"PARTIAL_FALLBACK ({type(last_err).__name__}: {last_err})"


# ============================================================
# DOCX rendering
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


def render_alumno_docx(data: Dict[str, Any], alumno: Dict[str, str], logo_b: Optional[bytes], enable_img: bool) -> bytes:
    doc = Document()
    header_block(doc, alumno, logo_b, "FICHA DEL ALUMNO")

    p = doc.add_paragraph()
    add_text(p, "Objetivo de aprendizaje", bold=True)
    p2 = doc.add_paragraph()
    p2.paragraph_format.line_spacing = 1.5
    add_text(p2, str(data.get("objetivo_aprendizaje", "")), bold=False)

    doc.add_paragraph("")

    p = doc.add_paragraph()
    add_text(p, "Consigna general (paso a paso)", bold=True)
    cg = str(data.get("consigna_general_alumno", "")).strip()
    for line in [x.strip() for x in cg.split("\n") if x.strip()]:
        p3 = doc.add_paragraph()
        p3.paragraph_format.line_spacing = 1.5
        add_text(p3, line, bold=False)

    doc.add_paragraph("")

    for idx, it in enumerate(data.get("items", []), start=1):
        if not isinstance(it, dict):
            continue

        tipo = str(it.get("tipo", "")).strip()
        enunciado = ensure_action_emoji(tipo, str(it.get("enunciado", "")).strip())
        pasos = it.get("pasos", []) if isinstance(it.get("pasos", []), list) else []
        opciones = it.get("opciones", []) if isinstance(it.get("opciones", []), list) else []
        formato = str(it.get("respuesta_formato", "texto_corto")).strip()
        kw = it.get("keywords_bold", []) if isinstance(it.get("keywords_bold", []), list) else []
        pista = str(it.get("pista_visual", "")).strip()

        v = it.get("visual", {}) if isinstance(it.get("visual", {}), dict) else {}
        v_en = normalize_bool(v.get("habilitado", False))
        v_pr = normalize_visual_prompt(str(v.get("prompt", "")).strip()) if v_en else ""

        card = doc.add_table(rows=1, cols=1)
        card.width = Inches(6.5)
        cell = card.rows[0].cells[0]
        apply_card_style(cell, fill_hex="FAFAFA")
        clear_paragraph(cell.paragraphs[0])

        pt = cell.add_paragraph()
        pt.paragraph_format.line_spacing = 1.5
        add_text(pt, f"Ítem {idx}", bold=True, size_pt=12)

        p_con = cell.add_paragraph()
        p_con.paragraph_format.line_spacing = 1.6
        add_text_with_keywords(p_con, enunciado, kw, size_pt=14)

        if pasos:
            for i, step in enumerate(pasos[:8], start=1):
                ps = cell.add_paragraph()
                ps.paragraph_format.line_spacing = 1.5
                add_text(ps, f"{i}. {str(step)}", bold=False)

        sep = cell.add_paragraph()
        add_text(sep, "Trabajo", bold=True, size_pt=12)

        if opciones and (formato.lower() in {"multiple_choice", "multiple choice"}):
            checkbox_list(cell, [str(x) for x in opciones], max_opts=8)
        else:
            response_box(cell, label="✍️ Respuesta:", lines=4)

        if pista:
            pp = cell.add_paragraph()
            pp.paragraph_format.line_spacing = 1.5
            add_text(pp, "Pista", bold=True, size_pt=12)

            pp2 = cell.add_paragraph()
            pp2.paragraph_format.line_spacing = 1.5
            add_text(pp2, "💡 " + pista, bold=False, color=RGBColor(0, 150, 0), size_pt=14)

        if enable_img and v_en and v_pr:
            img_bytes = generate_image_bytes(DEFAULT_IMAGE_MODEL, v_pr)
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

    p = doc.add_paragraph()
    add_text(p, "Objetivo de aprendizaje", bold=True)
    p2 = doc.add_paragraph()
    p2.paragraph_format.line_spacing = 1.5
    add_text(p2, str(data.get("objetivo_aprendizaje", "")), bold=False)

    doc.add_paragraph("")

    sol = data.get("solucionario_docente", {}) if isinstance(data.get("solucionario_docente", {}), dict) else {}
    respuestas = sol.get("respuestas", []) if isinstance(sol.get("respuestas", []), list) else []

    by_idx: Dict[int, Dict[str, Any]] = {}
    for r in respuestas:
        if not isinstance(r, dict):
            continue
        try:
            idx = int(r.get("item_index", 0) or 0)
        except Exception:
            idx = 0
        if idx > 0:
            by_idx[idx] = r

    for idx, it in enumerate(data.get("items", []), start=1):
        if not isinstance(it, dict):
            continue

        en = str(it.get("enunciado", "")).strip()
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
        add_text(pf, str(r.get("respuesta_final", "(no provista)")), bold=False)

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
# Cached generation (por alumno)
# ============================================================
@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def cached_activity(cache_key: str, prompt: str) -> Tuple[Dict[str, Any], str]:
    return robust_generate_activity(DEFAULT_TEXT_MODEL, prompt)


def build_prompt(mode: str, grado: str, grupo: str, diagnostico: str, nombre: str, input_text: str) -> str:
    # NO se pierde contexto por alumno: esto SIEMPRE va adentro del prompt.
    # Sin "on the fly".
    if mode == "CREAR":
        ctx = f"CREAR ACTIVIDAD DESDE CERO:\n{input_text}\n"
    else:
        ctx = f"ADAPTAR CONTENIDO ORIGINAL:\n{input_text}\n"

    return f"""{SYSTEM_PROMPT_OPAL_V243}

CONTEXTO:
{ctx}

DATOS DEL ALUMNO (OBLIGATORIO USAR PARA ADECUAR DIFICULTAD):
- nombre: {nombre}
- diagnostico / dificultad: {diagnostico}
- grupo: {grupo}
- grado: {grado}

REQUISITOS DE PERSONALIZACIÓN:
- Ajusta dificultad al diagnóstico/dificultad del alumno.
- Ajusta ejemplos al grupo (grupo: {grupo}).
- Mantén wording breve y concreto.
""".strip()


# ============================================================
# App
# ============================================================
def main():
    st.title("Nano Opal v24.3 🧠🍌")
    st.caption("Modelo fijo: gemini-2.5-flash. Sin UI de modelos. Sin on-the-fly. Anti-fallo JSON (autofill).")

    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    except Exception as e:
        st.error(f"GOOGLE_API_KEY inválida/faltante: {e}")
        return

    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando planilla: {e}")
        return

    # columnas por posición (como tu versión previa)
    grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
    grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
    diag_col = df.columns[4] if len(df.columns) > 4 else df.columns[0]

    # Controles (sin sidebar)
    cA, cB, cC = st.columns([2, 2, 2])
    with cA:
        grado = st.selectbox("Grado", sorted(df[grado_col].dropna().unique().tolist()))
    df_f = df[df[grado_col] == grado].copy()

    with cB:
        alcance = st.radio("Alcance", ["Todo el grado", "Seleccionar alumnos"], horizontal=True)
    if alcance == "Seleccionar alumnos":
        with cC:
            al_sel = st.multiselect("Alumnos", sorted(df_f[alumno_col].dropna().unique().tolist()))
        df_final = df_f[df_f[alumno_col].isin(al_sel)].copy() if al_sel else df_f.iloc[0:0].copy()
    else:
        df_final = df_f

    st.divider()

    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        enable_img = st.checkbox("Imágenes (best-effort)", value=True)
    with c2:
        debug_save_json = st.checkbox("Guardar JSON en ZIP", value=False)
    with c3:
        logo = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
        l_bytes = logo.read() if logo else None

    st.divider()

    tab1, tab2 = st.tabs(["🔄 Adaptar DOCX", "✨ Crear Actividad"])
    adapt_docx = None
    brief = ""

    with tab1:
        adapt_docx = st.file_uploader("Actividad base (DOCX)", type=["docx"], key="docx_in")

    with tab2:
        brief = st.text_area("Prompt", height=220, placeholder="Tema + objetivo + grado. Ej: Lectoescritura 1ero: sílabas directas...")

    mode = "CREAR" if (brief and brief.strip()) else "ADAPTAR"
    input_text = ""

    if mode == "ADAPTAR":
        if not adapt_docx:
            st.info("Subí un DOCX o usa 'Crear Actividad'.")
            return
        input_text = extraer_texto_docx(adapt_docx).strip()
        if len(input_text) < 120:
            st.error("DOCX extraído muy corto. No se inicia.")
            return
    else:
        input_text = brief.strip()
        if len(input_text) < 10:
            st.error("Prompt muy corto. No se inicia.")
            return

    if st.button("🚀 GENERAR LOTE"):
        if len(df_final) == 0:
            st.error("No hay alumnos (revisar selección).")
            return

        zip_io = io.BytesIO()
        ok_count = 0
        errors: List[str] = []

        logs = [
            "Nano Opal v24.3",
            f"Inicio: {now_str()}",
            f"Modo: {mode}",
            f"Modelo texto: {DEFAULT_TEXT_MODEL}",
            f"Modelo imagen: {DEFAULT_IMAGE_MODEL if enable_img else 'OFF'}",
            f"Grado: {grado}",
            f"Alumnos: {len(df_final)}",
            "",
        ]

        resumen_lines = [
            "RESUMEN - Nano Opal v24.3",
            f"Inicio: {now_str()}",
            f"Modo: {mode}",
            f"Grado: {grado}",
            "",
        ]

        base_hash = hash_text(f"{mode}|{grado}|{input_text}|{SYSTEM_PROMPT_OPAL_V243}|{DEFAULT_TEXT_MODEL}")

        with zipfile.ZipFile(zip_io, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("_REPORTE.txt", "\n".join(logs))

            prog = st.progress(0.0)
            status = st.empty()

            for idx, (_, row) in enumerate(df_final.iterrows(), start=1):
                nombre = str(row[alumno_col]).strip()
                grupo = str(row[grupo_col]).strip()
                diagnostico = str(row[diag_col]).strip()

                status.info(f"Procesando: {nombre} ({idx}/{len(df_final)})")

                alumno_meta = {"nombre": nombre, "diagnostico": diagnostico, "grupo": grupo, "grado": str(grado)}

                try:
                    prompt = build_prompt(mode, str(grado), grupo, diagnostico, nombre, input_text)

                    cache_key = f"{base_hash}::{safe_filename(nombre)}::{safe_filename(grupo)}::{safe_filename(diagnostico)}"
                    data_raw, path_debug = cached_activity(cache_key, prompt)

                    # AUTOFILL (anti-fallo)
                    data = autofill_required_fields(data_raw, input_text, alumno_meta)

                    alumno_docx_b = render_alumno_docx(data=data, alumno=alumno_meta, logo_b=l_bytes, enable_img=enable_img)
                    docente_docx_b = render_docente_docx(data=data, alumno=alumno_meta, logo_b=l_bytes)

                    base_name = safe_filename(f"{nombre}__{grupo}__{grado}")
                    zf.writestr(f"{base_name}__ALUMNO.docx", alumno_docx_b)
                    zf.writestr(f"{base_name}__DOCENTE.docx", docente_docx_b)

                    if debug_save_json:
                        zf.writestr(f"{base_name}__DATA.json", json.dumps(data, ensure_ascii=False, indent=2))

                    ok_count += 1
                    resumen_lines.append(f"- {nombre} | OK | gen={path_debug} | grupo={grupo} | dif={diagnostico} | items={len(data.get('items', []))}")

                except Exception as e:
                    err = f"{nombre} | ERROR {type(e).__name__}: {e}"
                    errors.append(err)
                    resumen_lines.append(f"- {nombre} | ERROR {type(e).__name__}: {e}")

                prog.progress(min(1.0, idx / max(1, len(df_final))))

            resumen_lines.append("")
            resumen_lines.append(f"OK: {ok_count}/{len(df_final)}")
            resumen_lines.append(f"Errores: {len(errors)}")
            if errors:
                resumen_lines.append("")
                resumen_lines.extend(errors[:200])

            zf.writestr("_RESUMEN.txt", "\n".join(resumen_lines))

        status.success(f"Listo. OK={ok_count} | Errores={len(errors)}")
        zip_io.seek(0)
        st.download_button(
            "⬇️ Descargar ZIP",
            data=zip_io.getvalue(),
            file_name=f"NanoOpal_v24_3_{safe_filename(grado)}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
            mime="application/zip"
        )


if __name__ == "__main__":
    main()
