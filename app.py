import base64
import io
import json
import re
import time
import zipfile
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

import google.generativeai as genai

# =========================
# CONFIG
# =========================
APP_TITLE = "Opal Classroom vNext (HTML)"
SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

TEXT_MODEL = "gemini-2.5-flash"
IMAGE_MODEL = "gemini-2.5-flash-image"

CACHE_TTL_SHEETS = 10 * 60  # 10 min
CACHE_TTL_GEN = 6 * 60 * 60  # 6 h

# Si querés “más consistente” bajá temp a 0.0
GENCFG_TEXT = {"temperature": 0.2, "top_p": 0.9, "max_output_tokens": 4096, "response_mime_type": "application/json"}

SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

# Intento PDF vía WeasyPrint; si no está, uso fallback ReportLab.
WEASYPRINT_OK = False
WEASYPRINT_ERR = ""
try:
    from weasyprint import HTML as WEASY_HTML  # type: ignore
    WEASYPRINT_OK = True
except Exception as e:
    WEASYPRINT_ERR = f"{type(e).__name__}: {e}"
    WEASYPRINT_OK = False

REPORTLAB_OK = False
REPORTLAB_ERR = ""
try:
    from reportlab.lib.pagesizes import A4  # type: ignore
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # type: ignore
    from reportlab.lib.units import mm  # type: ignore
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage  # type: ignore
    REPORTLAB_OK = True
except Exception as e:
    REPORTLAB_ERR = f"{type(e).__name__}: {e}"
    REPORTLAB_OK = False


# =========================
# UTIL
# =========================
def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def b64e(b: bytes) -> str:
    return base64.b64encode(b).decode("utf-8")


def safe_filename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^A-Za-z0-9._-]", "_", s)
    return (s or "SIN_NOMBRE")[:120]


def strip_html(s: str) -> str:
    return re.sub(r"<[^>]+>", "", s or "").strip()


def coerce_list(x: Any) -> List[Any]:
    return x if isinstance(x, list) else []


def coerce_str(x: Any, default: str = "") -> str:
    if isinstance(x, str):
        return x.strip()
    if x is None:
        return default
    return str(x).strip()


def clamp(n: int, lo: int, hi: int) -> int:
    return max(lo, min(hi, n))


# =========================
# SHEETS
# =========================
@st.cache_data(ttl=CACHE_TTL_SHEETS, show_spinner=False)
def load_sheet_df(url: str) -> pd.DataFrame:
    df = pd.read_csv(url)
    df.columns = [c.strip() for c in df.columns]
    return df


def detect_columns(df: pd.DataFrame) -> Dict[str, str]:
    """
    Intenta mapear columnas por heurística para no romper si cambia el orden.
    Ajustá si tus headers son conocidos.
    """
    cols = list(df.columns)

    def pick(candidates: List[str], fallback_idx: int) -> str:
        low = {c.lower(): c for c in cols}
        for k in candidates:
            if k.lower() in low:
                return low[k.lower()]
        return cols[fallback_idx] if fallback_idx < len(cols) else cols[0]

    return {
        "grado": pick(["grado", "course", "level"], 1),
        "alumno": pick(["alumno", "nombre", "name", "student"], 2),
        "grupo": pick(["grupo", "grupo_base", "base_group", "group"], 3),
        "perfil": pick(["perfil", "dificultad", "diagnostico", "needs"], 4),
    }


# =========================
# PROMPT (MEJORADO)
# =========================
SYSTEM_PROMPT = """
Actúa como Diseñador Instruccional Senior + UX Editor.

Objetivo: generar una ficha por alumno con estética premium, consistente y lista para imprimir.
Salida: SOLO JSON válido (sin texto extra).

REQUISITOS DUROS:
- EXACTAMENTE 5 items.
- Cada item DEBE tener apoyo visual: prompt_imagen NO vacío.
- Enunciados concretos, verificables, sin ambigüedad.
- Dificultad debe adaptarse al PERFIL/DIFICULTAD del alumno y al GRUPO BASE.
- Evita ejercicios “sin sentido”: cada item debe conectar con el objetivo.
- No uses markdown.

ESQUEMA:
{
  "objetivo": "string",
  "items": [
    {
      "icono": "✍️|📖|🔢|🎨",
      "enunciado": "string",
      "pista": "string (microayuda accionable, 1-3 líneas)",
      "respuesta_tipo": "lineas|multiple_choice|cuadricula",
      "opciones": ["A", "B", "C", "D"],
      "prompt_imagen": "string (descriptivo; pictograma educativo; sin texto embebido)",
      "alt_imagen": "string (qué representa la imagen)"
    }
  ]
}
""".strip()


def build_user_prompt(
    brief: str,
    modo: str,
    alumno: str,
    grupo: str,
    perfil: str,
    grado: str,
) -> str:
    # “modo” te permite extender a ADAPTAR cuando vuelvas a meter DOCX.
    return f"""
{SYSTEM_PROMPT}

CONTEXTO DOCENTE:
- MODO: {modo}
- GRADO: {grado}
- BRIEF: {brief}

PERFIL ALUMNO:
- NOMBRE: {alumno}
- GRUPO BASE: {grupo}
- PERFIL/DIFICULTAD: {perfil}

REGLAS DE CALIDAD:
- Cada item debe tener: enunciado + pista + prompt_imagen + alt_imagen.
- prompt_imagen: describí objetos simples, alto contraste, estilo pictograma educativo, fondo blanco, trazos gruesos.
- Si el item es multiple_choice, completa opciones[4]. Si no, opciones[] vacío.
""".strip()


# =========================
# GEMINI CALLS
# =========================
def gemini_text_json(prompt: str) -> Tuple[Optional[Dict[str, Any]], str]:
    try:
        m = genai.GenerativeModel(TEXT_MODEL)
        r = m.generate_content(prompt, generation_config=GENCFG_TEXT, safety_settings=SAFETY_SETTINGS)
        raw = getattr(r, "text", None)
        if not raw:
            # fallback: intentar ensamblar parts
            raw = ""
            try:
                cand = r.candidates[0]
                parts = cand.content.parts or []
                raw = "".join([getattr(p, "text", "") for p in parts if getattr(p, "text", "")])
            except Exception:
                raw = ""

        raw = (raw or "").strip()
        if not raw:
            return None, "Empty response text from model"

        # Sanitización tolerante: recorta a primer { ... último }
        start = raw.find("{")
        end = raw.rfind("}")
        if start != -1 and end != -1 and end > start:
            raw = raw[start : end + 1]

        # Arreglos simples: trailing commas
        raw = re.sub(r",\s*}", "}", raw)
        raw = re.sub(r",\s*]", "]", raw)

        data = json.loads(raw)
        if not isinstance(data, dict):
            return None, "Root JSON is not an object"
        return data, "OK"
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"


def gemini_image_b64(prompt: str) -> Optional[str]:
    """
    Best-effort: si falla, devuelve None.
    """
    try:
        m = genai.GenerativeModel(IMAGE_MODEL)
        # Pedimos pictograma educativo consistente
        p = f"Pictograma educativo, fondo blanco, alto contraste, trazos negros gruesos, ultra simple, sin texto, de: {prompt}"
        r = m.generate_content(p, safety_settings=SAFETY_SETTINGS)
        # Intentar extraer inline_data bytes
        try:
            cand = r.candidates[0]
            parts = cand.content.parts or []
            for part in parts:
                inline = getattr(part, "inline_data", None) or getattr(part, "inlineData", None)
                if inline is None:
                    continue
                b = getattr(inline, "data", None)
                if isinstance(b, (bytes, bytearray)) and len(b) > 500:
                    return b64e(bytes(b))
        except Exception:
            pass

        # fallback: data-uri dentro de texto
        t = getattr(r, "text", "") or ""
        m2 = re.search(r"data:image/(png|jpeg|jpg|webp);base64,([A-Za-z0-9+/=\n\r]+)", t)
        if m2:
            return m2.group(2).strip()
        return None
    except Exception:
        return None


# =========================
# DATA NORMALIZATION (ANTI “JSON RÍGIDO”)
# =========================
@dataclass
class Item:
    icono: str
    enunciado: str
    pista: str
    respuesta_tipo: str
    opciones: List[str]
    prompt_imagen: str
    alt_imagen: str
    img_b64: Optional[str] = None


def normalize_payload(payload: Dict[str, Any]) -> Tuple[str, List[Item], List[str]]:
    """
    Nunca falla: completa defaults, fuerza 5 items.
    Devuelve (objetivo, items, warnings)
    """
    warns: List[str] = []

    objetivo = coerce_str(payload.get("objetivo"), "")
    if not objetivo:
        objetivo = "Practicar habilidades del día (según consigna)."
        warns.append("objetivo faltante: se aplicó default")

    items_raw = payload.get("items")
    if not isinstance(items_raw, list):
        items_raw = []
        warns.append("items faltante/no lista: se creó lista vacía")

    norm: List[Item] = []
    for i in range(min(len(items_raw), 5)):
        it = items_raw[i] if isinstance(items_raw[i], dict) else {}
        icono = coerce_str(it.get("icono"), "✍️")
        if icono not in {"✍️", "📖", "🔢", "🎨"}:
            icono = "✍️"

        enunciado = coerce_str(it.get("enunciado"), f"Ítem {i+1}")
        pista = coerce_str(it.get("pista"), "Hacé un paso por vez.")
        respuesta_tipo = coerce_str(it.get("respuesta_tipo"), "lineas").lower()
        if respuesta_tipo not in {"lineas", "multiple_choice", "cuadricula"}:
            respuesta_tipo = "lineas"
        opciones = [coerce_str(x) for x in coerce_list(it.get("opciones"))][:4]
        prompt_imagen = coerce_str(it.get("prompt_imagen"), "")
        alt_imagen = coerce_str(it.get("alt_imagen"), "")

        if not prompt_imagen:
            # fallback agresivo: derivar del enunciado
            prompt_imagen = strip_html(enunciado)[:160]
            warns.append(f"item {i+1}: prompt_imagen faltante, derivado del enunciado")

        if not alt_imagen:
            alt_imagen = f"Apoyo visual del ítem {i+1}"
            warns.append(f"item {i+1}: alt_imagen faltante, default")

        if respuesta_tipo == "multiple_choice":
            if len(opciones) < 4:
                # completar opciones para evitar UI rota
                base = opciones + [""] * (4 - len(opciones))
                opciones = [x if x else f"Opción {j+1}" for j, x in enumerate(base)]
                warns.append(f"item {i+1}: opciones incompletas, se completaron")
        else:
            opciones = []

        norm.append(Item(icono, enunciado, pista, respuesta_tipo, opciones, prompt_imagen, alt_imagen))

    # Forzar exactamente 5
    while len(norm) < 5:
        j = len(norm) + 1
        norm.append(
            Item(
                icono="✍️",
                enunciado=f"Ítem {j}: completá la consigna con un ejemplo.",
                pista="1) Leé. 2) Elegí datos. 3) Escribí tu respuesta.",
                respuesta_tipo="lineas",
                opciones=[],
                prompt_imagen=f"Objeto simple relacionado al ítem {j}",
                alt_imagen=f"Apoyo visual del ítem {j}",
            )
        )
        warns.append(f"se añadió item faltante hasta completar 5 (nuevo item {j})")

    return objetivo, norm[:5], warns


# =========================
# HTML RENDER (PREMIUM)
# =========================
def render_html(
    alumno: str,
    grupo: str,
    perfil: str,
    grado: str,
    objetivo: str,
    items: List[Item],
    logo_b64: str,
    theme_color: str,
) -> str:
    # Layout estable “paper”, consistente para PDF.
    # Evita grids; todo vertical.
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="height:64px; object-fit:contain;">' if logo_b64 else ""

    def answer_block(it: Item) -> str:
        if it.respuesta_tipo == "multiple_choice":
            opts = "".join([f'<div class="opt"><span class="box"></span><span>{st_html_escape(o)}</span></div>' for o in it.opciones])
            return f'<div class="answer"><div class="answer-title">Respuesta</div>{opts}</div>'
        if it.respuesta_tipo == "cuadricula":
            # 8x2 cuadritos
            cells = "".join(['<div class="cell"></div>' for _ in range(16)])
            return f'<div class="answer"><div class="answer-title">Respuesta</div><div class="grid">{cells}</div></div>'
        # lineas
        lines = "".join(['<div class="line"></div>' for _ in range(4)])
        return f'<div class="answer"><div class="answer-title">Respuesta</div>{lines}</div>'

    def img_block(it: Item) -> str:
        if it.img_b64:
            return f'<div class="imgbox"><img alt="{st_html_escape(it.alt_imagen)}" src="data:image/png;base64,{it.img_b64}" /></div>'
        # placeholder elegante
        return f'<div class="imgbox placeholder"><div class="ph">Apoyo visual</div><div class="ph2">{st_html_escape(it.alt_imagen)}</div></div>'

    cards = []
    for idx, it in enumerate(items, start=1):
        cards.append(
            f"""
            <section class="card">
              <div class="card-head">
                <div class="pill">Ítem {idx}</div>
              </div>
              <div class="enun">{st_html_escape(it.icono)} {st_html_escape(it.enunciado)}</div>
              {img_block(it)}
              <div class="hint"><span class="hint-ic">💡</span><span>{st_html_escape(it.pista)}</span></div>
              {answer_block(it)}
            </section>
            """
        )

    cards_html = "\n".join(cards)

    return f"""
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Ficha - {st_html_escape(alumno)}</title>
<style>
  :root {{
    --bg: #f6f7fb;
    --paper: #ffffff;
    --ink: #111827;
    --muted: #6b7280;
    --line: #e5e7eb;
    --soft: #f3f4f6;
    --accent: {theme_color};
    --hintbg: #ecfdf5;
    --hintbr: #10b981;
  }}

  @page {{ size: A4; margin: 16mm; }}

  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    background: var(--bg);
    font-family: Verdana, sans-serif;
    color: var(--ink);
    font-size: 14pt;
    line-height: 1.65;
  }}

  .paper {{
    width: 210mm;
    min-height: 297mm;
    margin: 0 auto;
    padding: 16mm;
    background: var(--paper);
  }}

  .header {{
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 12mm;
    padding-bottom: 8mm;
    border-bottom: 3px solid var(--accent);
    margin-bottom: 10mm;
  }}

  .meta {{
    flex: 1;
  }}

  .badge {{
    display: inline-block;
    padding: 6px 10px;
    border-radius: 999px;
    background: var(--accent);
    color: #fff;
    font-weight: 700;
    font-size: 10pt;
    letter-spacing: 0.4px;
  }}

  h1 {{
    margin: 6px 0 0 0;
    font-size: 20pt;
    line-height: 1.2;
  }}

  .sub {{
    margin-top: 4px;
    color: var(--muted);
    font-size: 11.5pt;
  }}

  .obj {{
    margin: 10mm 0 8mm 0;
    padding: 8mm;
    border: 2px solid var(--line);
    border-radius: 12px;
    background: #fff;
  }}

  .obj-title {{
    font-weight: 800;
    margin-bottom: 3mm;
  }}

  .card {{
    border: 2px solid var(--soft);
    border-radius: 16px;
    padding: 8mm;
    margin: 0 0 8mm 0;
    page-break-inside: avoid;
  }}

  .card-head {{
    display: flex;
    justify-content: space-between;
    margin-bottom: 4mm;
  }}

  .pill {{
    display: inline-block;
    padding: 4px 10px;
    border-radius: 999px;
    background: #111827;
    color: #fff;
    font-weight: 700;
    font-size: 10pt;
  }}

  .enun {{
    font-weight: 800;
    font-size: 15pt;
    margin-bottom: 4mm;
  }}

  .imgbox {{
    width: 100%;
    border-radius: 14px;
    background: #fff;
    border: 2px solid var(--line);
    padding: 6mm;
    display: flex;
    justify-content: center;
    align-items: center;
    margin: 4mm 0 4mm 0;
  }}

  .imgbox img {{
    max-width: 72mm;
    max-height: 48mm;
    object-fit: contain;
  }}

  .imgbox.placeholder {{
    background: #fafafa;
    border-style: dashed;
  }}
  .ph {{
    font-weight: 800;
    color: var(--muted);
    margin-bottom: 2mm;
    text-align: center;
  }}
  .ph2 {{
    font-size: 11pt;
    color: var(--muted);
    text-align: center;
  }}

  .hint {{
    background: var(--hintbg);
    border-left: 6px solid var(--hintbr);
    border-radius: 12px;
    padding: 5mm 6mm;
    color: #065f46;
    display: flex;
    gap: 3mm;
    align-items: flex-start;
    margin: 2mm 0 4mm 0;
    page-break-inside: avoid;
  }}

  .hint-ic {{
    font-weight: 800;
  }}

  .answer {{
    margin-top: 4mm;
  }}

  .answer-title {{
    font-weight: 800;
    margin-bottom: 2mm;
  }}

  .line {{
    height: 10mm;
    border-bottom: 2px solid #d1d5db;
    margin-top: 2.5mm;
  }}

  .opt {{
    display: flex;
    align-items: center;
    gap: 4mm;
    margin: 2mm 0;
  }}

  .box {{
    width: 6mm;
    height: 6mm;
    border: 2px solid #111827;
    border-radius: 2px;
    flex: 0 0 auto;
  }}

  .grid {{
    display: grid;
    grid-template-columns: repeat(8, 10mm);
    gap: 2mm;
  }}

  .cell {{
    width: 10mm;
    height: 10mm;
    border: 2px solid #111827;
    border-radius: 2px;
  }}

  .foot {{
    margin-top: 8mm;
    color: var(--muted);
    font-size: 10pt;
    border-top: 1px solid var(--line);
    padding-top: 4mm;
  }}
</style>
</head>
<body>
  <div class="paper">
    <div class="header">
      <div class="meta">
        <span class="badge">GRUPO {st_html_escape(grupo)} · GRADO {st_html_escape(grado)}</span>
        <h1>{st_html_escape(alumno)}</h1>
        <div class="sub">{st_html_escape(perfil)}</div>
      </div>
      <div>{logo_html}</div>
    </div>

    <div class="obj">
      <div class="obj-title">Objetivo</div>
      <div>{st_html_escape(objetivo)}</div>
    </div>

    {cards_html}

    <div class="foot">Generado: {now_str()}</div>
  </div>
</body>
</html>
""".strip()


def st_html_escape(s: str) -> str:
    s = s or ""
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    s = s.replace('"', "&quot;").replace("'", "&#039;")
    return s


# =========================
# PDF EXPORT
# =========================
def html_to_pdf_bytes(html: str) -> Tuple[Optional[bytes], str]:
    """
    Intenta WeasyPrint; si no, fallback ReportLab (simple).
    """
    if WEASYPRINT_OK:
        try:
            pdf = WEASY_HTML(string=html).write_pdf()
            return pdf, "OK (WeasyPrint)"
        except Exception as e:
            # cae al fallback
            wp_err = f"WeasyPrint fail: {type(e).__name__}: {e}"
    else:
        wp_err = f"WeasyPrint not available: {WEASYPRINT_ERR}"

    # fallback ReportLab: no renderiza HTML “perfecto”, pero no rompe.
    if REPORTLAB_OK:
        try:
            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm, topMargin=14 * mm, bottomMargin=14 * mm)
            styles = getSampleStyleSheet()
            base = ParagraphStyle("base", parent=styles["Normal"], fontName="Helvetica", fontSize=11, leading=14)

            story = []
            story.append(Paragraph("Ficha (fallback PDF)", styles["Title"]))
            story.append(Spacer(1, 8))
            # Texto plano: extrae un resumen del HTML para que al menos haya contenido.
            text = strip_html(html)
            # recortar para no explotar
            text = text[:8000]
            story.append(Paragraph(st_html_escape(text).replace("\n", "<br/>"), base))
            doc.build(story)
            return buf.getvalue(), f"OK (ReportLab fallback). {wp_err}"
        except Exception as e:
            return None, f"{wp_err} | ReportLab fail: {type(e).__name__}: {e}"

    return None, f"{wp_err} | ReportLab not available: {REPORTLAB_ERR}"


# =========================
# MAIN UI (SIMPLE)
# =========================
def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    except Exception as e:
        st.error(f"GOOGLE_API_KEY faltante/invalid: {e}")
        return

    # Info técnica mínima (no ensucia UI)
    with st.sidebar:
        st.markdown("### Config")
        st.caption(f"Modelo texto: {TEXT_MODEL}")
        st.caption(f"Modelo imagen: {IMAGE_MODEL}")
        st.caption(f"PDF WeasyPrint: {'OK' if WEASYPRINT_OK else 'NO'}")
        if not WEASYPRINT_OK:
            st.caption(f"Detalle: {WEASYPRINT_ERR}")
        st.caption(f"PDF ReportLab: {'OK' if REPORTLAB_OK else 'NO'}")
        if not REPORTLAB_OK:
            st.caption(f"Detalle: {REPORTLAB_ERR}")

    try:
        df = load_sheet_df(URL_PLANILLA)
    except Exception as e:
        st.error(f"Error cargando Google Sheet CSV: {e}")
        return

    colmap = detect_columns(df)
    grado_col = colmap["grado"]
    alumno_col = colmap["alumno"]
    grupo_col = colmap["grupo"]
    perfil_col = colmap["perfil"]

    # ===== Form principal (sin Ctrl+Enter) =====
    with st.form("opal_form", clear_on_submit=False):
        c1, c2 = st.columns([2, 1])

        with c1:
            brief = st.text_area(
                "Brief docente (qué deben aprender hoy)",
                height=160,
                placeholder="Ej: Sumas con llevadas usando caramelos. 5 ejercicios con apoyo visual. Nivel: 1er grado.",
            )

        with c2:
            grado = st.selectbox("Grado", sorted(df[grado_col].dropna().unique().tolist()))
            theme_color = st.color_picker("Color", value="#7C3AED")
            export_pdf = st.checkbox("Exportar PDF además de HTML", value=True)

        c3, c4 = st.columns([2, 1])
        with c3:
            alcance = st.radio("Alcance", ["Todo el grado", "Seleccionar alumnos"], horizontal=True)

            df_g = df[df[grado_col] == grado].copy()
            df_g[alumno_col] = df_g[alumno_col].astype(str)
            if alcance == "Seleccionar alumnos":
                sel = st.multiselect("Alumnos", sorted(df_g[alumno_col].dropna().unique().tolist()))
                df_run = df_g[df_g[alumno_col].isin(sel)].copy() if sel else df_g.iloc[0:0].copy()
            else:
                df_run = df_g

        with c4:
            logo = st.file_uploader("Logo (png/jpg)", type=["png", "jpg", "jpeg"])
            logo_b64 = b64e(logo.read()) if logo else ""

        submitted = st.form_submit_button("Generar fichas", use_container_width=True)

    if not submitted:
        return

    if not brief or not brief.strip():
        st.error("Brief vacío.")
        return

    if df_run.empty:
        st.error("No hay alumnos para generar (selección vacía).")
        return

    # ===== Procesamiento =====
    start = now_str()
    zip_io = io.BytesIO()
    resumen: List[str] = []
    resumen.append(f"RESUMEN - {APP_TITLE}")
    resumen.append(f"Inicio: {start}")
    resumen.append(f"Grado: {grado}")
    resumen.append(f"Alumnos: {len(df_run)}")
    resumen.append(f"PDF: {export_pdf} | WeasyPrint: {WEASYPRINT_OK} | ReportLab: {REPORTLAB_OK}")
    resumen.append("")

    prog = st.progress(0.0)
    status = st.empty()

    ok_count = 0
    err_count = 0

    with zipfile.ZipFile(zip_io, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for i, (_, row) in enumerate(df_run.iterrows(), start=1):
            alumno = coerce_str(row.get(alumno_col, ""))
            grupo = coerce_str(row.get(grupo_col, ""))
            perfil = coerce_str(row.get(perfil_col, ""))

            if not alumno:
                alumno = f"ALUMNO_{i}"

            status.info(f"{i}/{len(df_run)} · {alumno}")

            try:
                prompt = build_user_prompt(
                    brief=brief.strip(),
                    modo="CREAR",
                    alumno=alumno,
                    grupo=grupo,
                    perfil=perfil,
                    grado=str(grado),
                )

                payload, msg = gemini_text_json(prompt)
                if payload is None:
                    # No aborta: crea payload mínimo.
                    payload = {"objetivo": "Practicar habilidades del día.", "items": []}
                    warn0 = [f"modelo texto falló: {msg}"]
                else:
                    warn0 = []

                objetivo, items, warns = normalize_payload(payload)
                warns = warn0 + warns

                # imágenes por ítem (best-effort pero “agresivo”)
                for it in items:
                    # fuerza prompt consistente
                    pimg = it.prompt_imagen
                    # retry corto
                    img_b64 = None
                    for _ in range(2):
                        img_b64 = gemini_image_b64(pimg)
                        if img_b64:
                            break
                        time.sleep(0.2)
                    it.img_b64 = img_b64

                html = render_html(
                    alumno=alumno,
                    grupo=grupo,
                    perfil=perfil,
                    grado=str(grado),
                    objetivo=objetivo,
                    items=items,
                    logo_b64=logo_b64,
                    theme_color=theme_color,
                )

                base = f"{safe_filename(alumno)}__G{safe_filename(grupo)}"
                zf.writestr(f"{base}.html", html)

                pdf_note = ""
                if export_pdf:
                    pdf_bytes, pdf_msg = html_to_pdf_bytes(html)
                    pdf_note = pdf_msg
                    if pdf_bytes:
                        zf.writestr(f"{base}.pdf", pdf_bytes)
                    else:
                        zf.writestr(f"{base}__PDF_ERROR.txt", pdf_msg)

                # warnings por alumno (si hay)
                if warns:
                    zf.writestr(f"{base}__WARNINGS.txt", "\n".join(warns))

                resumen.append(f"- {alumno} | OK | warns={len(warns)} | pdf={pdf_note}")
                ok_count += 1

            except Exception as e:
                err_count += 1
                base = safe_filename(alumno)
                zf.writestr(f"{base}__ERROR.txt", f"{type(e).__name__}: {e}")
                resumen.append(f"- {alumno} | ERROR {type(e).__name__}: {e}")

            prog.progress(i / len(df_run))

        resumen.append("")
        resumen.append(f"OK: {ok_count}/{len(df_run)}")
        resumen.append(f"Errores: {err_count}")

        zf.writestr("_RESUMEN.txt", "\n".join(resumen))

    status.success(f"Listo. OK={ok_count} · Errores={err_count}")
    st.download_button(
        "Descargar ZIP",
        data=zip_io.getvalue(),
        file_name=f"Fichas_{safe_filename(str(grado))}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
        mime="application/zip",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
