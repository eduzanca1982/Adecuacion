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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ============================================================
# Motor Pedagógico v18.0 (Opal++ Ficha Profesional)
# - MODO: ADAPTAR (DOCX) o CREAR (Brief)
# - Prompt interno refactor (Senior Inclusive UX + Psicopedagogo)
# - Iconografía por ítem (emoji de acción al inicio)
# - Micro-pasos físicos/visuales en pistas
# - Visual.prompt ARASAAC-like (trazos negros gruesos, fondo blanco)
# - Render “Ficha” por ítem con tabla 1 celda, borde fino y sombreado tenue
# - Sin itálicas en todo el documento
# - Preserva negritas en texto usando **bold**
# - Boot scan + selección dinámica de modelo imagen por smoke test
# - Retry/backoff + manejo MAX_TOKENS (finish_reason=2) + respuesta sin parts
# - JSON estricto + validación dura + reparación 1 vez + fallback compacto
# - ZIP siempre incluye _REPORTE.txt y _RESUMEN.txt + ERROR_*.txt
# ============================================================

st.set_page_config(page_title="Motor Pedagógico v18.0 (Opal++)", layout="wide")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

RETRIES = 6
CACHE_TTL_SECONDS = 6 * 60 * 60

# Imagen estilo ARASAAC solicitado
IMAGE_PROMPT_PREFIX = "Pictograma estilo ARASAAC, trazos negros gruesos, fondo blanco, ultra simple, sin sombras de: "

# Validación mínima de bytes para considerar que una imagen “vino”
MIN_IMAGE_BYTES = 600

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

# Acción emojis (enunciado debe empezar con uno)
ACTION_EMOJI_BY_TIPO = {
    "completar": "✍️",
    "multiple choice": "🔢",
    "multiple_choice": "🔢",
    "unir": "📖",
    "verdadero_falso": "📖",
    "problema_guiado": "🔢",
    "leer": "📖",
    "calcular": "🔢",
    "dibujar": "🎨",
}


# ============================================================
# Prompt (nuevo) + builders
# ============================================================
SYSTEM_PROMPT_OPALUX = f"""
Actúa como un Senior Inclusive UX Designer y Tutor Psicopedagogo.

REGLAS DE ORO:
1) ICONOGRAFÍA: Cada ítem del JSON (items[]) debe incluir al inicio de su enunciado un emoji de acción:
   - ✍️ completar
   - 📖 leer
   - 🔢 calcular
   - 🎨 dibujar

2) MICRO-PASOS: Las pistas NO deben ser teóricas. Deben ser instrucciones de andamiaje físico o visual
   (ej: "Dibuja 3 bolsitas...", "Marca con un círculo...", "Separa en columnas...").

3) ESTILO DE IMAGEN: Si visual.habilitado=true, visual.prompt debe pedir exactamente:
   "{IMAGE_PROMPT_PREFIX}[OBJETO]"

4) SIN ITÁLICAS: Prohibido el uso de itálicas en el contenido.

GUÍAS STRICT DYSLEXIA-FRIENDLY:
- Frases cortas, 1 acción por frase. Pasos secuenciales.
- Vocabulario concreto. Evitar metáforas.
- Bloques pequeños, listas. Dar ejemplo breve cuando sea útil.
- Reducir carga cognitiva: menos ítems, pero con mayor claridad.
- No penalizar ortografía si el objetivo no es ortografía.
- Sugerir tiempo extra y lectura acompañada.

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
      "tipo": "multiple choice|unir|completar|verdadero_falso|problema_guiado|leer|calcular|dibujar",
      "enunciado": "string (DEBE EMPEZAR con emoji de acción)",
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
- JSON puro.
- control_calidad.items_count == len(items)
- visual.prompt solo si visual.habilitado=true y debe iniciar con "{IMAGE_PROMPT_PREFIX}"
""".strip()


def build_payload(
    modo: str,
    materia: str,
    nivel: str,
    tema: str,
    estilo_extra: str,
    alumno: Dict[str, str],
    original_text: str,
) -> Dict[str, Any]:
    return {
        "MODO": modo,
        "MATERIA": materia,
        "NIVEL": nivel,
        "TEMA": tema,
        "ESTILO_EXTRA": estilo_extra or "",
        "ALUMNO": alumno,
        "ORIGINAL": original_text or "",
    }


def build_prompt_full(payload: Dict[str, Any]) -> str:
    return f"{SYSTEM_PROMPT_OPALUX}\n\nENTRADA:\n{json.dumps(payload, ensure_ascii=False, indent=2)}"


def build_prompt_compact(payload: Dict[str, Any]) -> str:
    # fallback: menos ítems, sin visual, textos cortos
    payload2 = dict(payload)
    if isinstance(payload2.get("ORIGINAL"), str):
        payload2["ORIGINAL"] = payload2["ORIGINAL"][:3000]

    return f"""
Devuelve SOLO JSON válido (sin texto extra).
Max 6 items. Enunciados cortos. Pistas con micro-pasos. visual siempre false.

ESQUEMA:
{{
  "alumno": {{
    "nombre": "{payload.get('ALUMNO',{}).get('nombre','')}",
    "grupo": "{payload.get('ALUMNO',{}).get('grupo','')}",
    "diagnostico": "{payload.get('ALUMNO',{}).get('diagnostico','')}",
    "grado": "{payload.get('ALUMNO',{}).get('grado','')}"
  }},
  "contexto": {{
    "modo": "{payload.get('MODO','')}",
    "materia": "{payload.get('MATERIA','')}",
    "nivel": "{payload.get('NIVEL','')}",
    "tema": "{payload.get('TEMA','')}",
    "estilo_extra": "{payload.get('ESTILO_EXTRA','')}"
  }},
  "objetivo_aprendizaje": "string",
  "consigna_adaptada": "string",
  "items": [
    {{"tipo":"calcular","enunciado":"🔢 ...","opciones":[]}}
  ],
  "adecuaciones_aplicadas": ["string"],
  "sugerencias_docente": ["string"],
  "visual": {{"habilitado": false, "prompt": ""}},
  "control_calidad": {{"items_count": 0, "incluye_ejemplo": true, "lenguaje_concreto": true, "una_accion_por_frase": true}}
}}

ENTRADA:
{json.dumps(payload2, ensure_ascii=False)}
""".strip()


def build_repair_prompt(bad: str, why: str) -> str:
    return f"""
Devuelve EXCLUSIVAMENTE un JSON válido y corregido (sin texto extra).

Problema detectado:
{why}

JSON A CORREGIR:
{bad}

Reglas:
- Debe cumplir EXACTAMENTE el esquema.
- control_calidad.items_count == len(items)
- enunciado debe iniciar con emoji de acción
- visual.prompt debe iniciar con "{IMAGE_PROMPT_PREFIX}" si habilitado=true
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
    return {
        "habilitado": normalize_bool(v.get("habilitado", False)),
        "prompt": str(v.get("prompt", "")).strip()
    }


def ensure_action_emoji(tipo: str, enunciado: str) -> str:
    t = (tipo or "").strip().lower()
    e = (enunciado or "").strip()
    if not e:
        return e
    first = e[:2]
    # si ya trae alguno de los emojis esperados, no tocar
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


def validate_text_input(text: str, mode: str) -> Tuple[bool, str, Dict[str, Any]]:
    info = {
        "chars": len(text or ""),
        "lines": (text or "").count("\n") + (1 if text else 0),
        "pipes": (text or "").count("|"),
        "preview": (text or "")[:1400],
    }
    if mode == "ADAPTAR":
        if not text or not text.strip():
            return False, "TEXTO vacío tras extracción (posible actividad en imágenes/cuadros).", info
        if len(text) < 150:
            return False, "TEXTO muy corto (<150 chars). Posible doc con imágenes/shapes.", info
        return True, "OK", info
    # CREAR
    if not text or not text.strip():
        return False, "Brief vacío.", info
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
    prompt = normalize_visual_prompt(IMAGE_PROMPT_PREFIX + "manzana")
    try:
        m = genai.GenerativeModel(model_id)
        r = retry_with_backoff(lambda: m.generate_content(prompt, safety_settings=SAFETY_SETTINGS))
        data = _extract_inline_bytes_or_none(r)
        if not data:
            return False, "Respuesta sin inline_data.data"
        if len(data) < MIN_IMAGE_BYTES:
            return False, f"inline_data muy chico ({len(data)} bytes)"
        return True, f"OK bytes={len(data)}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"


def pick_best_image_model(models: List[str]) -> Tuple[Optional[str], str]:
    cands = _candidate_image_models(models)
    if not cands:
        return None, "No se detectaron candidatos de imagen en list_models()."
    for mid in cands[:12]:
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
        st.error("Fallo en arranque (boot scan de modelos).")
        st.code(f"{type(e).__name__}: {e}")
        st.stop()


BOOT = boot_or_stop()


# ============================================================
# Gemini response robusta (sin depender de response.text)
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
        if not isinstance(data, dict):
            return False, "Root no es objeto"

        for k in ["alumno", "contexto", "objetivo_aprendizaje", "consigna_adaptada", "items",
                  "adecuaciones_aplicadas", "sugerencias_docente", "visual", "control_calidad"]:
            if k not in data:
                return False, f"Falta clave: {k}"

        if not isinstance(data["items"], list) or len(data["items"]) < 1:
            return False, "items vacío/no lista"

        cc = data.get("control_calidad", {})
        if not isinstance(cc, dict):
            return False, "control_calidad no es objeto"
        if cc.get("items_count") != len(data["items"]):
            return False, "control_calidad.items_count no coincide con len(items)"

        v = data.get("visual", {})
        if not isinstance(v, dict):
            return False, "visual no es objeto"
        if normalize_bool(v.get("habilitado", False)):
            p = str(v.get("prompt", "")).strip()
            if not p.startswith(IMAGE_PROMPT_PREFIX):
                return False, "visual.prompt no respeta el prefijo ARASAAC requerido"

        # Enunciado debe iniciar con emoji
        for i, it in enumerate(data["items"][:200]):
            if not isinstance(it, dict):
                return False, f"items[{i}] no es objeto"
            tipo = str(it.get("tipo", "")).strip()
            en = str(it.get("enunciado", "")).strip()
            if not en:
                return False, f"items[{i}].enunciado vacío"
            if not any(en.startswith(x) for x in ["✍️", "📖", "🔢", "🎨"]):
                # tolerancia: lo arreglamos en normalización; pero marcamos warning como invalidez para que repare
                return False, f"items[{i}].enunciado no inicia con emoji (tipo={tipo})"

        return True, "OK"
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
    return generate_json_with_repair(model_id, prompt, max_out)


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
# Imagen (best effort)
# ============================================================
def generar_imagen_ia(model_id: str, prompt_img: str) -> Optional[io.BytesIO]:
    try:
        m = genai.GenerativeModel(model_id)
        resp = retry_with_backoff(lambda: m.generate_content(prompt_img, safety_settings=SAFETY_SETTINGS))
        data = _extract_inline_bytes_or_none(resp)
        if not data or len(data) < MIN_IMAGE_BYTES:
            return None
        return io.BytesIO(data)
    except Exception:
        return None


# ============================================================
# DOCX Rendering helpers (Ficha Opal)
# ============================================================
def set_cell_shading(cell, rgb_hex: str):
    """rgb_hex: 'F8F8F8'"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), rgb_hex)


def set_cell_border(cell, color: str = "D9D9D9", size: str = "6"):
    """
    size: eighths of a point (w:sz). '6' is thin-ish.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn("w:tcBorders"))
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)

    for edge in ("top", "left", "bottom", "right"):
        elem = tcBorders.find(qn(f"w:{edge}"))
        if elem is None:
            elem = OxmlElement(f"w:{edge}")
            tcBorders.append(elem)
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:sz"), size)
        elem.set(qn("w:space"), "0")
        elem.set(qn("w:color"), color)


def clear_paragraph(paragraph):
    p = paragraph._p
    for child in list(p):
        p.remove(child)


def add_runs_with_bold_markers(paragraph, text: str, font_name: str = "Verdana", font_size_pt: int = 14, bold_default: bool = False):
    """
    Preserva **negrita** en texto.
    Ej: "Marca **solo** las palabras" -> run normal + run bold + run normal
    """
    if text is None:
        text = ""
    parts = str(text).split("**")
    for i, part in enumerate(parts):
        run = paragraph.add_run(part)
        run.bold = (not bold_default) if (i % 2 == 1) else bold_default
        run.italic = False  # prohibido itálica
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)


def add_response_box(paragraph, font_name="Verdana", font_size_pt=14):
    """
    Crea línea de respuesta estilo ficha: "✍️ Mi respuesta:" + línea larga.
    """
    paragraph.paragraph_format.line_spacing = 1.5
    run = paragraph.add_run("✍️ Mi respuesta: ")
    run.bold = True
    run.italic = False
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    run2 = paragraph.add_run("______________________________________________")
    run2.bold = False
    run2.italic = False
    run2.font.name = font_name
    run2.font.size = Pt(font_size_pt)


def render_docx_ficha_opal(
    data: Dict[str, Any],
    logo_bytes: Optional[bytes],
    activar_img: bool,
    model_img_id: Optional[str],
) -> bytes:
    doc = Document()

    # Base style
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
    info_run.italic = False
    info_run.font.name = "Verdana"
    info_run.font.size = Pt(11)

    doc.add_paragraph("")  # spacer

    # Secciones
    def add_section_title(t: str):
        p = doc.add_paragraph()
        r = p.add_run(t)
        r.bold = True
        r.italic = False
        r.font.name = "Verdana"
        r.font.size = Pt(14)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)

    add_section_title("Objetivo de aprendizaje")
    p_obj = doc.add_paragraph()
    add_runs_with_bold_markers(p_obj, str(data.get("objetivo_aprendizaje", "")).strip(), font_size_pt=14)
    p_obj.paragraph_format.line_spacing = 1.5

    add_section_title("Consigna adaptada")
    p_con = doc.add_paragraph()
    add_runs_with_bold_markers(p_con, str(data.get("consigna_adaptada", "")).strip(), font_size_pt=14)
    p_con.paragraph_format.line_spacing = 1.5

    add_section_title("Actividad / Ítems")

    items = data.get("items", [])
    for idx, it in enumerate(items, start=1):
        tipo = str(it.get("tipo", "")).strip()
        enun = ensure_action_emoji(tipo, str(it.get("enunciado", "")).strip())
        opciones = it.get("opciones", [])
        if not isinstance(opciones, list):
            opciones = []

        # “Ficha” = tabla 1x1 con sombreado y borde
        t = doc.add_table(rows=1, cols=1)
        cell = t.rows[0].cells[0]
        set_cell_shading(cell, "F8F8F8")
        set_cell_border(cell, color="D9D9D9", size="6")

        # limpiar párrafo inicial
        clear_paragraph(cell.paragraphs[0])

        # Enunciado (preserva **bold**)
        p1 = cell.add_paragraph()
        p1.paragraph_format.line_spacing = 1.5
        r_idx = p1.add_run(f"{idx}. ")
        r_idx.bold = True
        r_idx.italic = False
        r_idx.font.name = "Verdana"
        r_idx.font.size = Pt(14)

        add_runs_with_bold_markers(p1, enun, font_size_pt=14)

        # Opciones o respuesta
        if opciones:
            for op in opciones[:10]:
                p_op = cell.add_paragraph()
                p_op.paragraph_format.line_spacing = 1.5
                run = p_op.add_run(f"☐ {str(op)}")
                run.bold = False
                run.italic = False
                run.font.name = "Verdana"
                run.font.size = Pt(14)
        else:
            p_resp = cell.add_paragraph()
            add_response_box(p_resp, font_size_pt=14)

        # espacio entre fichas
        doc.add_paragraph("")

    add_section_title("Adecuaciones aplicadas")
    for a in (data.get("adecuaciones_aplicadas", []) or [])[:25]:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        run = p.add_run(f"• {a}")
        run.bold = False
        run.italic = False
        run.font.name = "Verdana"
        run.font.size = Pt(14)

    add_section_title("Sugerencias para el docente")
    for s in (data.get("sugerencias_docente", []) or [])[:25]:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        run = p.add_run(f"• {s}")
        run.bold = False
        run.italic = False
        run.font.name = "Verdana"
        run.font.size = Pt(14)

    # Imagen de apoyo
    v = normalize_visual(data.get("visual", {}))
    if activar_img and model_img_id and v.get("habilitado"):
        pv = normalize_visual_prompt(v.get("prompt", ""))
        if pv:
            img = generar_imagen_ia(model_img_id, pv)
            if img:
                doc.add_paragraph("")
                pimg = doc.add_paragraph()
                pimg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                try:
                    pimg.add_run().add_picture(img, width=Inches(2.7))
                except Exception:
                    pass

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# UI + Proceso
# ============================================================
def main():
    st.title("Motor Pedagógico v18.0 (Opal++ Ficha)")
    st.caption("DOCX tipo ficha, neuroinclusión extrema, determinismo y robustez anti-fallas.")

    # Planilla
    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando planilla: {e}")
        return

    # Column mapping (compat)
    grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
    grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
    diag_col = df.columns[4] if len(df.columns) > 4 else df.columns[0]

    with st.sidebar:
        st.header("Boot scan (consultar primero)")
        st.write(f"Boot: {BOOT.get('boot_time')}")
        st.write(f"Modelo texto (auto): {BOOT.get('text_model')}")
        if BOOT.get("image_model"):
            st.write(f"Modelo imagen (auto): {BOOT.get('image_model')}")
            st.caption(BOOT.get("image_reason", ""))
        else:
            st.write("Modelo imagen (auto): N/A")
            st.caption(BOOT.get("image_reason", ""))

        st.divider()
        st.header("Grupo")
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
        modo_ui = st.radio("Modo", ["ADAPTAR (subir DOCX)", "CREAR (solo prompt/brief)"])
        modo = "ADAPTAR" if modo_ui.startswith("ADAPTAR") else "CREAR"

        st.divider()
        st.header("Contexto")
        materia = st.text_input("Materia", value="Matemática")
        nivel = st.text_input("Nivel", value="7mo grado")
        tema = st.text_input("Tema", value="División progresiva")

        st.divider()
        st.header("Instrucciones de Estilo On-the-fly")
        estilo_extra = st.text_area(
            "Se inyecta directo al prompt",
            placeholder="Ej: reducir a 6 ítems, incluir 1 ejemplo guiado, usar 2 multiple choice y 2 completar, y 1 unir columnas."
        )

        st.divider()
        st.header("Salida")
        activar_img_user = st.checkbox("Generar pictograma de apoyo (si disponible)", value=True)
        activar_img = activar_img_user and (BOOT.get("image_model") is not None)
        if activar_img_user and not activar_img:
            st.warning("Imágenes desactivadas: no hay modelo de imagen funcional (smoke test falló).")

        logo = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
        logo_bytes = logo.read() if logo else None

        st.divider()
        diag_mode = st.checkbox("Diagnóstico (preview texto/brief)", value=True)

    st.subheader("Entrada")
    original_text = ""

    if modo == "ADAPTAR":
        archivo = st.file_uploader("Subir actividad/examen base (DOCX)", type=["docx"])
        if archivo:
            original_text = extraer_texto_docx(archivo)
            ok_in, msg_in, info_in = validate_text_input(original_text, mode="ADAPTAR")
            if diag_mode:
                if ok_in:
                    st.success(f"Parseo DOCX OK ({info_in['chars']} chars, {info_in['lines']} líneas)")
                else:
                    st.error(f"Parseo DOCX: {msg_in}")
                with st.expander("Preview texto extraído", expanded=False):
                    st.text(info_in.get("preview", ""))
        else:
            st.info("Subí un DOCX para adaptar.")
    else:
        brief = st.text_area(
            "Escribí el prompt/brief (crear desde cero)",
            placeholder="Ej: Matemática 7mo grado, división progresiva. 1 ejemplo guiado. 6 ejercicios. 2 multiple choice, 2 completar, 1 verdadero/falso, 1 unir."
        )
        original_text = (brief or "").strip()
        ok_in, msg_in, info_in = validate_text_input(original_text, mode="CREAR")
        if diag_mode:
            if ok_in:
                st.success(f"Brief OK ({info_in['chars']} chars)")
            else:
                st.error(f"Brief: {msg_in}")
            with st.expander("Preview brief", expanded=False):
                st.text(info_in.get("preview", ""))

    if st.button("🚀 GENERAR LOTE (ZIP)"):
        if len(alumnos_final) == 0:
            st.error("No hay alumnos para procesar (selección vacía).")
            return

        ok_in, msg_in, info_in = validate_text_input(original_text, mode=modo)
        if not ok_in:
            st.error(f"No se inicia: {msg_in}")
            return

        model_text = BOOT["text_model"]
        model_img = BOOT.get("image_model") if activar_img else None

        input_hash = hash_text(f"{modo}|{materia}|{nivel}|{tema}|{estilo_extra}|{original_text}")

        zip_io = io.BytesIO()
        logs: List[str] = []
        errors: List[str] = []
        ok_count = 0

        logs.append("Motor Pedagógico v18.0 (Opal++ Ficha)")
        logs.append(f"Inicio: {now_str()}")
        logs.append(f"Modo: {modo}")
        logs.append(f"Materia: {materia} | Nivel: {nivel} | Tema: {tema}")
        logs.append(f"Estilo extra: {estilo_extra}")
        logs.append(f"Modelo texto: {model_text}")
        logs.append(f"Modelo imagen: {model_img if model_img else 'N/A'}")
        logs.append(f"Imágenes habilitadas: {bool(model_img)}")
        logs.append(f"Grado planilla: {grado_sel}")
        logs.append(f"Total alumnos: {len(alumnos_final)}")
        logs.append(f"Hash entrada: {input_hash}")
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
                    alumno = {"nombre": n, "grupo": g, "diagnostico": d, "grado": str(grado_sel)}
                    payload = build_payload(modo, materia, nivel, tema, estilo_extra, alumno, original_text)

                    prompt_full = build_prompt_full(payload)
                    prompt_comp = build_prompt_compact(payload)

                    cache_key = f"{input_hash}::{model_text}::{n}::{g}::{d}::{grado_sel}"
                    data, mode_used, max_t = request_activity_ultra(model_text, prompt_full, prompt_comp, cache_key)

                    # Normalización defensiva (si el modelo omitió algo)
                    data.setdefault("alumno", {})
                    data["alumno"].setdefault("nombre", n)
                    data["alumno"].setdefault("grupo", g)
                    data["alumno"].setdefault("diagnostico", d)
                    data["alumno"].setdefault("grado", str(grado_sel))

                    data.setdefault("contexto", {})
                    data["contexto"].setdefault("modo", modo)
                    data["contexto"].setdefault("materia", materia)
                    data["contexto"].setdefault("nivel", nivel)
                    data["contexto"].setdefault("tema", tema)
                    data["contexto"].setdefault("estilo_extra", estilo_extra or "")

                    # Normaliza items: emoji + opciones list
                    norm_items = []
                    for it in (data.get("items", []) or []):
                        if not isinstance(it, dict):
                            continue
                        tipo_i = str(it.get("tipo", "")).strip()
                        en_i = ensure_action_emoji(tipo_i, str(it.get("enunciado", "")).strip())
                        ops = it.get("opciones", [])
                        if not isinstance(ops, list):
                            ops = []
                        norm_items.append({"tipo": tipo_i, "enunciado": en_i, "opciones": [str(x) for x in ops]})
                    data["items"] = norm_items

                    # Normaliza visual
                    v = normalize_visual(data.get("visual", {}))
                    if v.get("habilitado"):
                        v["prompt"] = normalize_visual_prompt(v.get("prompt", ""))
                    data["visual"] = v

                    # Arregla items_count si vino mal (para no fallar render)
                    data.setdefault("control_calidad", {})
                    if isinstance(data["control_calidad"], dict):
                        data["control_calidad"]["items_count"] = len(data.get("items", []))

                    okj, whyj = validate_activity_json(data)
                    if not okj:
                        raise ValueError(f"JSON final inválido: {whyj}")

                    docx_bytes = render_docx_ficha_opal(
                        data=data,
                        logo_bytes=logo_bytes,
                        activar_img=bool(model_img),
                        model_img_id=model_img,
                    )

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
        st.download_button("📥 Descargar ZIP", zip_io.getvalue(), "actividades_opalpp_v18_0.zip", mime="application/zip")


if __name__ == "__main__":
    main()
