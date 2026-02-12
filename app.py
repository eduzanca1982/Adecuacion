import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
import zipfile
import time
import random
import hashlib
import base64
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

# ============================================================
# Nano Opal HTML v25.0
# - Generación directa de HTML por IA (sin JSON rígido)
# - Per-alumno: respeta grupo + diagnóstico + grado SIEMPRE (inyectado en prompt)
# - Estética consistente: plantilla CSS base + tokens de tema
# - Imágenes:
#     - Modo "IA (SVG inline)": el modelo dibuja pictos en SVG (consistente, 0 dependencias)
#     - Modo "Híbrido": IA devuelve data-img-prompt y se intenta generar PNG por ítem (best-effort)
# - Export:
#     - HTML siempre
#     - PDF opcional si WeasyPrint está instalado
# - UI: submit button claro (st.form), sin Ctrl+Enter
# - Model fijo: gemini-2.5-flash (sin selector manual)
# ============================================================

st.set_page_config(page_title="Nano Opal HTML v25.0", layout="wide")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

# Modelo fijo (no se expone en UI)
TEXT_MODEL_ID = "models/gemini-2.5-flash"

# Imagen best-effort (solo si habilitás híbrido y el SDK/modelo devuelve bytes)
# Si no existe/funciona, no rompe: simplemente no inserta imágenes.
PREFERRED_IMAGE_MODEL_IDS = [
    "models/gemini-2.5-flash-image",
    "models/imagen-3.0-generate-001",
    "models/imagen-3.0-fast-generate-001",
]

RETRIES = 5
MIN_HTML_CHARS = 2200  # umbral para evitar "archivo casi vacío"

# ------------------------------------------------------------
# Optional PDF backend (WeasyPrint)
# ------------------------------------------------------------
WEASYPRINT_AVAILABLE = False
try:
    from weasyprint import HTML as WeasyHTML  # type: ignore
    WEASYPRINT_AVAILABLE = True
except Exception:
    WEASYPRINT_AVAILABLE = False

# ------------------------------------------------------------
# Google AI
# ------------------------------------------------------------
SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

# Aumentá max_output_tokens si querés más detalle visual.
BASE_GEN_CFG_HTML = {
    "temperature": 0.4,           # algo de creatividad visual sin delirios
    "top_p": 0.9,
    "top_k": 40,
    "max_output_tokens": 8192,
}

BASE_GEN_CFG_HTML_RETRY = {
    "temperature": 0.55,
    "top_p": 0.95,
    "top_k": 60,
    "max_output_tokens": 8192,
}

# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
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
        "timeout", "timed out", "deadline", "unavailable", "503", "500",
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
            time.sleep(min(sleep, 18))
    raise last

def extract_text_from_gemini(resp) -> str:
    try:
        cand = resp.candidates[0]
        content = getattr(cand, "content", None)
        if not content or not getattr(content, "parts", None):
            return ""
        chunks = []
        for p in content.parts:
            t = getattr(p, "text", None)
            if t:
                chunks.append(t)
        return "".join(chunks).strip()
    except Exception:
        return ""

def ensure_html_doc(raw: str) -> str:
    s = (raw or "").strip()
    if not s:
        return ""
    # Recorte si vino con texto extra
    start = s.lower().find("<!doctype")
    if start == -1:
        start = s.lower().find("<html")
    if start != -1:
        s = s[start:]
    # Cierre si falta
    if "</html>" not in s.lower():
        s = s + "\n</html>\n"
    return s

def looks_like_html(s: str) -> bool:
    if not s:
        return False
    sl = s.lower()
    return ("<html" in sl) and ("</html>" in sl) and ("<body" in sl)

def html_to_pdf_bytes(html_str: str) -> Optional[bytes]:
    if not WEASYPRINT_AVAILABLE:
        return None
    try:
        return WeasyHTML(string=html_str).write_pdf()
    except Exception:
        return None

# ------------------------------------------------------------
# HTML template tokens (consistencia estética)
# El modelo NO decide layout global; solo rellena contenido dentro de slots.
# ------------------------------------------------------------
THEMES = {
    "Opal Clean": {
        "bg": "#0B1020",
        "paper": "#0F172A",
        "card": "#111C35",
        "ink": "#EAF0FF",
        "muted": "#A8B3D6",
        "accent": "#7C3AED",
        "good": "#22C55E",
        "warn": "#F59E0B",
        "bad":  "#EF4444",
        "line": "rgba(255,255,255,0.10)",
        "shadow": "0 18px 50px rgba(0,0,0,0.45)",
    },
    "Paper Bright": {
        "bg": "#F4F6FB",
        "paper": "#FFFFFF",
        "card": "#F8FAFF",
        "ink": "#0B1220",
        "muted": "#42526E",
        "accent": "#2563EB",
        "good": "#16A34A",
        "warn": "#D97706",
        "bad":  "#DC2626",
        "line": "rgba(15,23,42,0.10)",
        "shadow": "0 14px 40px rgba(15,23,42,0.12)",
    },
    "Soft Pastel": {
        "bg": "#F7F7FB",
        "paper": "#FFFFFF",
        "card": "#FBFBFF",
        "ink": "#1B2333",
        "muted": "#4B5563",
        "accent": "#A855F7",
        "good": "#22C55E",
        "warn": "#F59E0B",
        "bad":  "#EF4444",
        "line": "rgba(27,35,51,0.10)",
        "shadow": "0 18px 45px rgba(27,35,51,0.10)",
    },
}

DENSITY = {
    "Compacto": {"pad": "12px", "gap": "10px", "radius": "16px", "h1": "18px", "h2": "13px", "p": "12px"},
    "Normal":   {"pad": "16px", "gap": "12px", "radius": "18px", "h1": "20px", "h2": "14px", "p": "13px"},
    "Aireado":  {"pad": "20px", "gap": "14px", "radius": "20px", "h1": "22px", "h2": "15px", "p": "14px"},
}

def build_css(theme_name: str, density_name: str, font_family: str) -> str:
    t = THEMES.get(theme_name, THEMES["Opal Clean"])
    d = DENSITY.get(density_name, DENSITY["Normal"])
    return f"""
:root {{
  --bg: {t["bg"]};
  --paper: {t["paper"]};
  --card: {t["card"]};
  --ink: {t["ink"]};
  --muted: {t["muted"]};
  --accent: {t["accent"]};
  --good: {t["good"]};
  --warn: {t["warn"]};
  --bad: {t["bad"]};
  --line: {t["line"]};
  --shadow: {t["shadow"]};
  --pad: {d["pad"]};
  --gap: {d["gap"]};
  --radius: {d["radius"]};
  --h1: {d["h1"]};
  --h2: {d["h2"]};
  --p: {d["p"]};
  --font: {font_family};
}}

@page {{
  size: A4;
  margin: 14mm 14mm 14mm 14mm;
}}

* {{
  box-sizing: border-box;
}}

html, body {{
  height: 100%;
}}

body {{
  margin: 0;
  background: var(--bg);
  color: var(--ink);
  font-family: var(--font);
}}

.wrapper {{
  max-width: 980px;
  margin: 0 auto;
  padding: 18px 14px 26px 14px;
}}

.paper {{
  background: var(--paper);
  border: 1px solid var(--line);
  border-radius: calc(var(--radius) + 8px);
  box-shadow: var(--shadow);
  overflow: hidden;
}}

.topbar {{
  padding: 16px 18px;
  background: linear-gradient(135deg, rgba(124,58,237,0.16), rgba(37,99,235,0.10));
  border-bottom: 1px solid var(--line);
  display: grid;
  grid-template-columns: 1fr auto;
  gap: 12px;
  align-items: center;
}}

.brand {{
  display: grid;
  gap: 3px;
}}

.brand h1 {{
  margin: 0;
  font-size: var(--h1);
  letter-spacing: 0.2px;
}}

.brand .meta {{
  margin: 0;
  font-size: var(--p);
  color: var(--muted);
}}

.badges {{
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
  justify-content: flex-end;
}}

.badge {{
  border: 1px solid var(--line);
  background: rgba(255,255,255,0.06);
  padding: 6px 10px;
  border-radius: 999px;
  font-size: 12px;
  color: var(--muted);
}}

.grid {{
  padding: 16px 18px 18px 18px;
  display: grid;
  gap: var(--gap);
}}

.hero {{
  border: 1px solid var(--line);
  background: rgba(255,255,255,0.06);
  border-radius: var(--radius);
  padding: var(--pad);
  display: grid;
  gap: 10px;
}}

.hero .title {{
  display: flex;
  gap: 10px;
  align-items: center;
}}

.hero .title .dot {{
  width: 10px;
  height: 10px;
  border-radius: 999px;
  background: var(--accent);
}}

.hero h2 {{
  margin: 0;
  font-size: var(--h2);
}}

.hero p {{
  margin: 0;
  font-size: var(--p);
  color: var(--muted);
  line-height: 1.45;
}}

.cards {{
  display: grid;
  gap: var(--gap);
}}

.card {{
  border: 1px solid var(--line);
  background: var(--card);
  border-radius: var(--radius);
  padding: var(--pad);
  display: grid;
  gap: 10px;
  page-break-inside: avoid;
}}

.card-head {{
  display: grid;
  grid-template-columns: 1fr auto;
  gap: 10px;
  align-items: start;
}}

.card h3 {{
  margin: 0;
  font-size: var(--h2);
}}

.chip {{
  border: 1px solid var(--line);
  background: rgba(255,255,255,0.06);
  padding: 4px 10px;
  border-radius: 999px;
  font-size: 12px;
  color: var(--muted);
  white-space: nowrap;
}}

.two {{
  display: grid;
  grid-template-columns: 1.4fr 0.6fr;
  gap: 12px;
}}

@media print {{
  body {{
    background: #ffffff;
  }}
  .wrapper {{
    max-width: none;
    padding: 0;
  }}
  .paper {{
    box-shadow: none;
    border: none;
    border-radius: 0;
  }}
}}

.content p, .content li {{
  font-size: var(--p);
  line-height: 1.55;
  margin: 0;
  color: var(--ink);
}}

.muted {{
  color: var(--muted);
}}

.list {{
  margin: 0;
  padding-left: 18px;
  display: grid;
  gap: 6px;
}}

.answerbox {{
  border: 1px dashed var(--line);
  border-radius: calc(var(--radius) - 6px);
  padding: 10px;
  min-height: 70px;
  background: rgba(255,255,255,0.04);
}}

.pista {{
  border-left: 4px solid var(--good);
  background: rgba(34,197,94,0.10);
  padding: 10px 12px;
  border-radius: calc(var(--radius) - 8px);
  font-size: var(--p);
  color: var(--ink);
}}

.alert {{
  border-left: 4px solid var(--warn);
  background: rgba(245,158,11,0.12);
  padding: 10px 12px;
  border-radius: calc(var(--radius) - 8px);
  font-size: var(--p);
  color: var(--ink);
}}

.imgbox {{
  border: 1px solid var(--line);
  border-radius: calc(var(--radius) - 8px);
  background: rgba(255,255,255,0.04);
  padding: 10px;
  display: grid;
  place-items: center;
  min-height: 130px;
  overflow: hidden;
}}

.imgbox img {{
  max-width: 100%;
  height: auto;
  display: block;
}}

.small {{
  font-size: 12px;
  color: var(--muted);
}}
""".strip()

# ------------------------------------------------------------
# Prompts: IA produce HTML completo, con CSS inyectado por nosotros.
# IMPORTANT: el prompt fuerza slots, evita divagar en layout.
# ------------------------------------------------------------

HTML_SKELETON = """<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>{TITLE}</title>
  <style>
  {CSS}
  </style>
</head>
<body>
  <div class="wrapper">
    <div class="paper">
      <div class="topbar">
        <div class="brand">
          <h1>{H1}</h1>
          <p class="meta">{META}</p>
        </div>
        <div class="badges">
          {BADGES}
        </div>
      </div>

      <div class="grid">
        <div class="hero">
          <div class="title">
            <div class="dot"></div>
            <h2>{HERO_TITLE}</h2>
          </div>
          <p>{HERO_TEXT}</p>
        </div>

        <div class="cards">
          {CARDS}
        </div>

        {FOOTER}
      </div>
    </div>
  </div>
</body>
</html>
"""

def make_badges(badges: List[str]) -> str:
    out = []
    for b in badges[:8]:
        out.append(f'<div class="badge">{b}</div>')
    return "\n".join(out)

def build_student_prompt(
    brief_or_source: str,
    modo: str,
    alumno: Dict[str, str],
    grado: str,
    visual_mode: str,
    image_quality: str,
) -> str:
    # visual_mode:
    #   - "IA (SVG inline)" => el modelo inserta SVGs en .imgbox
    #   - "Híbrido (prompts + generación)" => el modelo emite data-img-prompt y placeholders
    svg_rule = (
        "En cada card, si incluyes apoyo visual, dibuja un pictograma en SVG inline dentro de <div class=\"imgbox\">...</div> "
        "Estilo: monocromo, líneas gruesas, simple, sin sombras."
        if visual_mode == "IA (SVG inline)" else
        "En cada card, agrega <div class=\"imgbox\" data-img-prompt=\"...\">"
        "Incluye dentro un placeholder <div class=\"small\">Generando imagen...</div>. "
        "El prompt debe ser muy concreto, 6-12 palabras, en español, sin comillas."
    )

    quality_rule = {
        "Rápido": "Produce 6 ítems, cards más cortas.",
        "Estándar": "Produce 8 ítems, buen equilibrio.",
        "Premium": "Produce 10 ítems, más variedad + mejor scaffolding visual.",
    }.get(image_quality, "Produce 8 ítems, buen equilibrio.")

    return f"""
Sos un motor pedagógico y diseñador editorial.
Tu tarea: devolver UN SOLO documento HTML completo (incluye <!doctype html> ... </html>).
Prohibido markdown. Prohibido texto fuera del HTML.

Objetivo: ficha de alumno de 60 minutos, neuroinclusiva (TDAH/dislexia friendly).
Formato: CARDS visuales, micro-pasos concretos, poco texto, alta claridad.

IMPORTANTE (NO NEGOCIABLE):
- Debe quedar explícito que ajustas dificultad por DIAGNÓSTICO y GRUPO. NO lo nombres como "diagnóstico"; solo se refleja en cómo lo presentas (más micro-pasos, menos carga, etc).
- Usa SIEMPRE: nombre, grupo, grado, y el perfil del alumno que te doy abajo.
- {quality_rule}
- Cada card:
  - encabezado con "Ítem N" + chip de tipo (Lectura / Cálculo / Escritura / Elección / Dibujo / Guiado)
  - enunciado con emoji inicial (✍️📖🔢🎨)
  - lista de pasos (2-6 bullets cortos)
  - zona "Trabajo" con caja de respuesta o checkboxes
  - una "Pista" breve y concreta (micro-paso físico/visual)
  - apoyo visual según regla

Regla de apoyo visual:
- {svg_rule}

Si el modo es "Elección", incluye 4 opciones con checkboxes.
Si no, incluye una caja de respuesta (answerbox).

Contenido fuente ({modo}):
{brief_or_source}

Alumno (NO omitir nunca):
- Nombre: {alumno["nombre"]}
- Grupo: {alumno["grupo"]}
- Grado: {grado}
- Perfil de aprendizaje (texto crudo): {alumno["perfil"]}

Salida: HTML completo. Nada más.
""".strip()

def build_teacher_prompt(
    student_html: str,
    alumno: Dict[str, str],
    grado: str,
) -> str:
    return f"""
Devolvé UN SOLO documento HTML completo (incluye <!doctype html> ... </html>).
Prohibido markdown. Prohibido texto fuera del HTML.

Objetivo: SOLUCIONARIO DOCENTE + adecuaciones aplicadas + errores frecuentes.
Debe ser visual y escueto.

Entradas:
- Nombre: {alumno["nombre"]}
- Grupo: {alumno["grupo"]}
- Grado: {grado}
- Perfil: {alumno["perfil"]}
- HTML del alumno (para alinear ítems): 
{student_html}

Reglas:
- Mantener la misma cantidad de ítems y numeración que el alumno.
- Por cada ítem: Respuesta final + desarrollo en 2-5 bullets + errores frecuentes (2-4).
- Agregar bloque final: "Adecuaciones aplicadas" (5-10 bullets) y "Criterios de corrección" (5-10 bullets).
- Estética tipo cards.

Salida: HTML completo. Nada más.
""".strip()

# ------------------------------------------------------------
# Image generation (best-effort) + injection
# ------------------------------------------------------------
DATA_IMG_PROMPT_RE = re.compile(r'<div class="imgbox"\s+data-img-prompt="([^"]{3,180})"\s*>', re.IGNORECASE)

def looks_like_image(b: bytes) -> bool:
    if not b or len(b) < 1200:
        return False
    if b[:8] == b"\x89PNG\r\n\x1a\n":
        return True
    if b[:3] == b"\xff\xd8\xff":
        return True
    if b[:4] == b"RIFF" and b[8:12] == b"WEBP":
        return True
    return False

def try_generate_image_bytes(prompt_img: str) -> Optional[bytes]:
    prompt_img = (prompt_img or "").strip()
    if not prompt_img:
        return None

    def _extract_inline(resp) -> Optional[bytes]:
        try:
            cand = resp.candidates[0]
            content = getattr(cand, "content", None)
            if not content or not getattr(content, "parts", None):
                return None
            for part in content.parts:
                inline = getattr(part, "inline_data", None)
                if inline is not None:
                    data = getattr(inline, "data", None)
                    if isinstance(data, (bytes, bytearray)):
                        return bytes(data)
                    if isinstance(data, str) and data.strip():
                        try:
                            return base64.b64decode(data, validate=False)
                        except Exception:
                            return None
                t = getattr(part, "text", None)
                if isinstance(t, str) and "base64" in t.lower() and len(t) > 500:
                    m = re.search(r"base64,([A-Za-z0-9+/=\n\r]+)", t)
                    if m:
                        try:
                            return base64.b64decode(m.group(1), validate=False)
                        except Exception:
                            return None
            return None
        except Exception:
            return None

    cfg_variants = [
        {"response_modalities": ["Image"]},
        {"response_modalities": ["IMAGE"]},
        {"responseModalities": ["Image"]},
        {"responseModalities": ["IMAGE"]},
        None,
    ]

    for mid in PREFERRED_IMAGE_MODEL_IDS:
        for cfg in cfg_variants:
            try:
                m = genai.GenerativeModel(mid)
                if cfg is None:
                    resp = retry_with_backoff(lambda: m.generate_content(prompt_img, safety_settings=SAFETY_SETTINGS))
                else:
                    resp = retry_with_backoff(lambda: m.generate_content(prompt_img, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
                b = _extract_inline(resp)
                if b and looks_like_image(b):
                    return b
            except Exception:
                continue
    return None

def inject_generated_images(html_doc: str, max_images: int = 10) -> Tuple[str, int]:
    prompts = DATA_IMG_PROMPT_RE.findall(html_doc)[:max_images]
    if not prompts:
        return html_doc, 0

    replaced = 0
    out = html_doc

    for p in prompts:
        img_bytes = try_generate_image_bytes(p)
        if not img_bytes:
            continue
        mime = "image/png"
        if img_bytes[:3] == b"\xff\xd8\xff":
            mime = "image/jpeg"
        elif img_bytes[:4] == b"RIFF" and img_bytes[8:12] == b"WEBP":
            mime = "image/webp"
        b64 = base64.b64encode(img_bytes).decode("ascii", errors="ignore")

        # Reemplaza el primer placeholder de esa imgbox
        pattern = re.compile(
            r'(<div class="imgbox"\s+data-img-prompt="' + re.escape(p) + r'"\s*>)([\s\S]*?)(</div>)',
            re.IGNORECASE
        )
        def repl(m):
            nonlocal replaced
            replaced += 1
            return m.group(1) + f'\n<img alt="Apoyo visual" src="data:{mime};base64,{b64}"/>\n' + m.group(3)

        out, n = pattern.subn(repl, out, count=1)
        if n == 0:
            continue

    return out, replaced

# ------------------------------------------------------------
# Core generation
# ------------------------------------------------------------
def generate_html_once(model_id: str, prompt: str, retry_cfg: bool = False) -> str:
    cfg = dict(BASE_GEN_CFG_HTML_RETRY if retry_cfg else BASE_GEN_CFG_HTML)
    m = genai.GenerativeModel(model_id)
    resp = retry_with_backoff(lambda: m.generate_content(prompt, generation_config=cfg, safety_settings=SAFETY_SETTINGS))
    txt = extract_text_from_gemini(resp)
    return ensure_html_doc(txt)

def robust_generate_student_html(prompt: str) -> Tuple[str, str]:
    # returns (html, status)
    h = generate_html_once(TEXT_MODEL_ID, prompt, retry_cfg=False)
    if looks_like_html(h) and len(h) >= MIN_HTML_CHARS:
        return h, "OK"
    # retry con prompt reforzado
    harden = prompt + "\n\nURGENCIA: tu salida anterior fue incompleta. Devuelve HTML COMPLETO, largo y con cards. NO recortes.\n"
    h2 = generate_html_once(TEXT_MODEL_ID, harden, retry_cfg=True)
    if looks_like_html(h2) and len(h2) >= MIN_HTML_CHARS:
        return h2, "OK_RETRY"
    return (h2 if looks_like_html(h2) else h), "WEAK_HTML"

def robust_generate_teacher_html(prompt: str) -> Tuple[str, str]:
    h = generate_html_once(TEXT_MODEL_ID, prompt, retry_cfg=False)
    if looks_like_html(h) and len(h) >= 1800:
        return h, "OK"
    harden = prompt + "\n\nURGENCIA: salida incompleta. Devuelve HTML COMPLETO con cards. Nada fuera del HTML.\n"
    h2 = generate_html_once(TEXT_MODEL_ID, harden, retry_cfg=True)
    if looks_like_html(h2) and len(h2) >= 1800:
        return h2, "OK_RETRY"
    return (h2 if looks_like_html(h2) else h), "WEAK_HTML"

def wrap_with_template(content_html_fragment: str, css: str, title: str, h1: str, meta: str, badges: List[str], hero_title: str, hero_text: str) -> str:
    # El modelo devuelve HTML completo. Para consistencia, forzamos nuestro CSS igual:
    # Estrategia: si ya trae <style>, lo dejamos pero inyectamos el nuestro primero.
    # Para no romper, reemplazamos el <style> existente agregando nuestro CSS arriba.
    html = content_html_fragment
    if not looks_like_html(html):
        # fallback: encajonar texto plano
        cards = f'<div class="card"><div class="card-head"><h3>Contenido</h3><div class="chip">Fallback</div></div><div class="content"><p>{st._utils.escape_markdown(html)}</p></div></div>'  # type: ignore
        footer = '<div class="small muted">Fallback por HTML inválido.</div>'
        return HTML_SKELETON.format(
            TITLE=title,
            CSS=css,
            H1=h1,
            META=meta,
            BADGES=make_badges(badges),
            HERO_TITLE=hero_title,
            HERO_TEXT=hero_text,
            CARDS=cards,
            FOOTER=footer
        )

    # Inyectar nuestro CSS al inicio del primer <style> si existe, sino insertarlo en <head>
    if "<style" in html.lower():
        html = re.sub(r"(<style[^>]*>)", r"\1\n" + css + "\n", html, count=1, flags=re.IGNORECASE)
    else:
        html = re.sub(r"(</head>)", "<style>\n" + css + "\n</style>\n\\1", html, count=1, flags=re.IGNORECASE)

    # Enriquecer topbar si el modelo no lo puso (sin romper si ya existe):
    # No intentamos re-estructurar: solo agregamos un header al comienzo del body.
    if "class=\"topbar\"" not in html:
        header = f"""
<div class="wrapper">
  <div class="paper">
    <div class="topbar">
      <div class="brand">
        <h1>{h1}</h1>
        <p class="meta">{meta}</p>
      </div>
      <div class="badges">
        {make_badges(badges)}
      </div>
    </div>
    <div class="grid">
      <div class="hero">
        <div class="title"><div class="dot"></div><h2>{hero_title}</h2></div>
        <p>{hero_text}</p>
      </div>
"""
        footer = """
    </div>
  </div>
</div>
"""
        # Insertar header al inicio del <body>, y cerrar antes de </body>
        html = re.sub(r"(<body[^>]*>)", r"\1\n" + header, html, count=1, flags=re.IGNORECASE)
        html = re.sub(r"(</body>)", footer + r"\1", html, count=1, flags=re.IGNORECASE)

    return html

# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
def main():
    st.title("Nano Opal HTML v25.0")
    st.caption("HTML consistente + PDF opcional. Modelo fijo: gemini-2.5-flash. Personalización por alumno: grupo + perfil + grado (inyectado en prompt).")

    # API key
    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    except Exception as e:
        st.error(f"Falta/invalid GOOGLE_API_KEY en secrets: {e}")
        return

    # Load sheet
    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando planilla: {e}")
        return

    # Heurística columnas
    # Esperado: [?, Grado, Alumno, Grupo, Diagnóstico/Perfil]
    grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
    grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
    perfil_col = df.columns[4] if len(df.columns) > 4 else (df.columns[3] if len(df.columns) > 3 else df.columns[0])

    # Controls
    left, right = st.columns([0.62, 0.38], gap="large")

    with right:
        st.subheader("Apariencia")
        theme_name = st.selectbox("Tema", list(THEMES.keys()), index=0)
        density_name = st.selectbox("Densidad", list(DENSITY.keys()), index=1)
        font_family = st.selectbox("Fuente", ["Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial", "Verdana, Arial, system-ui", "Noto Sans, Arial, system-ui"], index=0)

        st.subheader("Visuales")
        visual_mode = st.selectbox("Imágenes", ["IA (SVG inline)", "Híbrido (prompts + generación)"], index=0)
        image_quality = st.selectbox("Calidad", ["Rápido", "Estándar", "Premium"], index=1)

        st.subheader("Export")
        export_pdf = st.checkbox("Generar PDF además de HTML", value=WEASYPRINT_AVAILABLE, disabled=not WEASYPRINT_AVAILABLE)
        if not WEASYPRINT_AVAILABLE:
            st.info("PDF no disponible (WeasyPrint no instalado). Se exporta HTML.")

        max_images = st.slider("Máx. imágenes a intentar (híbrido)", 0, 14, 8, disabled=(visual_mode != "Híbrido (prompts + generación)"))
        st.divider()
        st.subheader("Sheets")
        grado = st.selectbox("Grado", sorted(df[grado_col].dropna().unique().tolist()))
        df_f = df[df[grado_col] == grado].copy()

        alcance = st.radio("Alcance", ["Todo el grado", "Seleccionar alumnos"], horizontal=True)
        if alcance == "Seleccionar alumnos":
            al_sel = st.multiselect("Alumnos", sorted(df_f[alumno_col].dropna().unique().tolist()))
            df_final = df_f[df_f[alumno_col].isin(al_sel)].copy() if al_sel else df_f.iloc[0:0].copy()
        else:
            df_final = df_f

    with left:
        st.subheader("Contenido")
        tabs = st.tabs(["Crear desde prompt", "Adaptar (texto pegado)"])
        with tabs[0]:
            with st.form("form_prompt", clear_on_submit=False):
                brief = st.text_area(
                    "Prompt (se recomienda concreto: tema, objetivos, restricción de texto, ejemplos)",
                    height=260,
                    placeholder="Ej: Actividad 60 min para 1ero sobre sumar y restar hasta 20 con objetos. Incluir 2 ítems con lectura simple, 4 de cálculo, 2 de elección múltiple, 2 guiados."
                )
                submitted = st.form_submit_button("Generar lote")
        with tabs[1]:
            with st.form("form_paste", clear_on_submit=False):
                pasted = st.text_area("Pegá el contenido base a adaptar", height=260)
                submitted2 = st.form_submit_button("Generar lote (adaptar)")
                if submitted2:
                    submitted = True
                    brief = pasted

        st.divider()
        st.write("Preview CSS (para validar estética base):")
        css = build_css(theme_name, density_name, font_family)
        st.code(css[:1800] + ("\n...\n" if len(css) > 1800 else ""), language="css")

    if not (locals().get("submitted", False)):
        return

    if len(df_final) == 0:
        st.error("No hay alumnos seleccionados.")
        return

    if not brief or not str(brief).strip():
        st.error("Prompt/base vacío.")
        return

    modo = "CREAR" if tabs[0] else "ADAPTAR"
    modo = "CREAR"  # simplificación estable: el prompt define el comportamiento
    source_text = str(brief).strip()

    # ZIP build
    zip_io = io.BytesIO()
    ok_count = 0
    err_count = 0
    errors: List[str] = []

    run_id = hash_text(now_str() + source_text)[:10]

    logs = []
    logs.append("Nano Opal HTML v25.0")
    logs.append(f"Inicio: {now_str()}")
    logs.append(f"RunID: {run_id}")
    logs.append(f"Modelo texto: {TEXT_MODEL_ID}")
    logs.append(f"Tema: {theme_name} | Densidad: {density_name} | Fuente: {font_family}")
    logs.append(f"Imágenes: {visual_mode} | Calidad: {image_quality}")
    logs.append(f"PDF: {export_pdf} (weasyprint={WEASYPRINT_AVAILABLE})")
    logs.append(f"Grado: {grado}")
    logs.append(f"Alumnos: {len(df_final)}")
    logs.append("")

    prog = st.progress(0.0)
    status = st.empty()

    with zipfile.ZipFile(zip_io, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("_REPORTE.txt", "\n".join(logs))
        zf.writestr("_CSS_BASE.css", css)

        for i, (_, row) in enumerate(df_final.iterrows(), start=1):
            prog.progress(i / max(1, len(df_final)))
            nombre = str(row[alumno_col]).strip()
            grupo = str(row[grupo_col]).strip()
            perfil = str(row[perfil_col]).strip()

            status.info(f"Generando: {nombre} ({i}/{len(df_final)})")

            alumno = {
                "nombre": nombre,
                "grupo": grupo,
                "perfil": perfil
            }

            try:
                # Student HTML
                p_student = build_student_prompt(
                    brief_or_source=source_text,
                    modo="CREAR",
                    alumno=alumno,
                    grado=str(grado),
                    visual_mode=visual_mode,
                    image_quality=image_quality,
                )
                student_html_raw, st_status = robust_generate_student_html(p_student)

                # Force our CSS + header (consistencia)
                student_html = wrap_with_template(
                    content_html_fragment=student_html_raw,
                    css=css,
                    title=f"Ficha Alumno - {nombre}",
                    h1="FICHA DEL ALUMNO",
                    meta=f"{nombre} · Grupo {grupo} · Grado {grado}",
                    badges=[f"60 min", f"Grupo {grupo}", f"Grado {grado}", "Neuroinclusivo"],
                    hero_title="Objetivo",
                    hero_text="Actividad guiada con micro-pasos, carga cognitiva controlada, y apoyos visuales.",
                )

                # Híbrido: generar imágenes e inyectar
                img_injected = 0
                if visual_mode == "Híbrido (prompts + generación)" and max_images > 0:
                    student_html, img_injected = inject_generated_images(student_html, max_images=max_images)

                # Teacher HTML
                p_teacher = build_teacher_prompt(student_html=student_html, alumno=alumno, grado=str(grado))
                teacher_html_raw, tch_status = robust_generate_teacher_html(p_teacher)
                teacher_html = wrap_with_template(
                    content_html_fragment=teacher_html_raw,
                    css=css,
                    title=f"Solucionario - {nombre}",
                    h1="SOLUCIONARIO DOCENTE",
                    meta=f"{nombre} · Grupo {grupo} · Grado {grado}",
                    badges=["Solucionario", "Errores frecuentes", "Adecuaciones"],
                    hero_title="Uso docente",
                    hero_text="Respuestas y criterios. Ajustes aplicados según perfil (sin exponerlo como etiqueta).",
                )

                # Write outputs
                base = safe_filename(f"{grado}__{grupo}__{nombre}")
                zf.writestr(f"{base}__ALUMNO.html", student_html)
                zf.writestr(f"{base}__DOCENTE.html", teacher_html)

                # PDF if enabled
                if export_pdf and WEASYPRINT_AVAILABLE:
                    pdf_a = html_to_pdf_bytes(student_html)
                    if pdf_a:
                        zf.writestr(f"{base}__ALUMNO.pdf", pdf_a)
                    pdf_d = html_to_pdf_bytes(teacher_html)
                    if pdf_d:
                        zf.writestr(f"{base}__DOCENTE.pdf", pdf_d)

                # Minimal resumen
                resumen = [
                    f"Alumno: {nombre}",
                    f"Grupo: {grupo}",
                    f"Grado: {grado}",
                    f"Status alumno: {st_status}",
                    f"Imágenes inyectadas: {img_injected}",
                    f"Status docente: {tch_status}",
                    f"Chars alumno html: {len(student_html)}",
                    f"Chars docente html: {len(teacher_html)}",
                    "",
                ]
                zf.writestr(f"{base}__RESUMEN.txt", "\n".join(resumen))

                ok_count += 1

            except Exception as e:
                err_count += 1
                msg = f"{nombre} | ERROR {type(e).__name__}: {e}"
                errors.append(msg)
                zf.writestr(f"ERROR__{safe_filename(nombre)}.txt", msg)

        # Resumen global
        global_sum = []
        global_sum.append("RESUMEN - Nano Opal HTML v25.0")
        global_sum.append(f"Fin: {now_str()}")
        global_sum.append(f"OK: {ok_count}/{len(df_final)}")
        global_sum.append(f"Errores: {err_count}")
        global_sum.append("")
        global_sum.extend(errors[:200])
        zf.writestr("_RESUMEN.txt", "\n".join(global_sum))

    status.success(f"Listo. OK={ok_count} | Errores={err_count}")

    st.download_button(
        "Descargar ZIP",
        data=zip_io.getvalue(),
        file_name=f"NanoOpalHTML_{grado}_{run_id}.zip",
        mime="application/zip",
        use_container_width=True
    )

    if errors:
        st.error("Errores (primeros 20):")
        st.code("\n".join(errors[:20]))

    st.divider()
    st.subheader("Notas técnicas (para mejorar estética e imágenes)")
    st.write(
        "- Si querés consistencia extrema: mantené 'IA (SVG inline)'. Es lo más estable y siempre renderiza bien.\n"
        "- Si querés imágenes “más ricas”: usá 'Híbrido'. Es best-effort: depende del modelo/SDK; si falla, queda placeholder sin romper el HTML.\n"
        "- Para PDF consistente: instalá WeasyPrint. En Streamlit Cloud suele requerir libs del sistema.\n"
    )

if __name__ == "__main__":
    main()
