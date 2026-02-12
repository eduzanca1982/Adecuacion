import streamlit as st
import google.generativeai as genai
import pandas as pd
import io
import zipfile
import base64
import json
import re
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

# ---------- Optional WeasyPrint (may fail on Streamlit Cloud) ----------
WEASYPRINT_OK = False
WEASYPRINT_ERR = ""
try:
    from weasyprint import HTML  # type: ignore
    WEASYPRINT_OK = True
except Exception as e:
    WEASYPRINT_OK = False
    WEASYPRINT_ERR = f"{type(e).__name__}: {e}"

# ---------- ReportLab fallback (pure Python) ----------
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import utils
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(page_title="Opal Classroom v29 INFALIBLE", layout="wide")

TEXT_MODEL = "gemini-2.5-flash"
IMAGE_MODEL = "gemini-2.5-flash-image"

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

THEME_COLOR = "#7C3AED"

# ============================================================
# HTML PREMIUM
# ============================================================

GLOBAL_CSS = f"""
@page {{ size: A4; margin: 24mm 18mm 22mm 18mm; }}

body {{
    font-family: Verdana, sans-serif;
    background: #f4f6fb;
    color: #111827;
}}

.paper {{
    background: white;
    padding: 34px;
    border-radius: 12px;
}}

.header {{
    border-bottom: 4px solid {THEME_COLOR};
    padding-bottom: 14px;
    margin-bottom: 26px;
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 20px;
}}

.student-name {{
    font-size: 24px;
    font-weight: 700;
    margin: 8px 0 4px 0;
}}

.meta {{
    font-size: 12px;
    color: #374151;
}}

.badges {{
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}}

.badge {{
    background: {THEME_COLOR};
    color: white;
    padding: 6px 10px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
}}

.objective {{
    background: #f5f3ff;
    border: 1px solid #ddd6fe;
    padding: 14px 14px;
    border-radius: 12px;
    margin-bottom: 22px;
    line-height: 1.7;
}}

.card {{
    border: 2px solid #e5e7eb;
    border-radius: 16px;
    padding: 20px;
    margin-bottom: 22px;
    line-height: 1.9;
}}

.enunciado {{
    font-size: 16px;
    font-weight: 700;
    margin-bottom: 12px;
}}

.img-box {{
    text-align: center;
    margin: 14px 0 12px 0;
    background: #f9fafb;
    border: 1px dashed #e5e7eb;
    border-radius: 12px;
    padding: 10px;
}}

.pista {{
    background: #ecfdf5;
    border-left: 6px solid #10b981;
    padding: 12px;
    border-radius: 10px;
    margin-top: 14px;
    color: #065f46;
    font-style: italic;
}}

.answer-line {{
    border-bottom: 2px solid #cbd5e1;
    height: 28px;
    margin-top: 16px;
}}

.small {{
    font-size: 11px;
    color: #6b7280;
}}
"""

# ============================================================
# JSON TOLERANTE (sin rigidez)
# ============================================================

_JSON_CODEFENCE_RE = re.compile(r"```(?:json)?\s*([\s\S]*?)\s*```", re.IGNORECASE)

def _strip_to_json_object(s: str) -> str:
    if not s:
        return s
    m = _JSON_CODEFENCE_RE.search(s)
    if m:
        s = m.group(1)
    start = s.find("{")
    end = s.rfind("}")
    if start != -1 and end != -1 and end > start:
        return s[start:end+1].strip()
    return s.strip()

def safe_json_loads_loose(text: str) -> Dict[str, Any]:
    if not text:
        raise ValueError("Empty response text")
    t = _strip_to_json_object(text)
    return json.loads(t)

def build_json_fix_prompt(raw: str, err: str) -> str:
    return f"""
Tu única tarea es devolver JSON válido (sin texto extra).
No agregues comentarios. No markdown.
Corrige SOLO sintaxis.

ERROR:
{err}

TEXTO:
{raw}
""".strip()

def generate_json_infalleable(model_id: str, prompt: str) -> Dict[str, Any]:
    m = genai.GenerativeModel(model_id)
    # Pedimos JSON, pero toleramos que falle.
    res = m.generate_content(
        prompt,
        generation_config={"response_mime_type": "application/json", "temperature": 0, "top_p": 1, "top_k": 1}
    )

    raw = getattr(res, "text", None) or ""
    try:
        return safe_json_loads_loose(raw)
    except Exception as e1:
        # JSON-fix con el mismo modelo (solo sintaxis)
        fix = build_json_fix_prompt(raw, f"{type(e1).__name__}: {e1}")
        res2 = m.generate_content(
            fix,
            generation_config={"response_mime_type": "application/json", "temperature": 0, "top_p": 1, "top_k": 1}
        )
        raw2 = getattr(res2, "text", None) or ""
        return safe_json_loads_loose(raw2)

# ============================================================
# IMAGEN (best-effort)
# ============================================================

def generar_imagen_base64(prompt_visual: str) -> Optional[str]:
    try:
        m = genai.GenerativeModel(IMAGE_MODEL)
        r = m.generate_content(
            f"Pictograma educativo claro, fondo blanco, estilo simple, alto contraste, sin sombras, de: {prompt_visual}"
        )
        b = r.candidates[0].content.parts[0].inline_data.data
        return base64.b64encode(b).decode()
    except Exception:
        return None

# ============================================================
# RENDER HTML + PDF
# ============================================================

def render_html(activity: Dict[str, Any], alumno: Dict[str, str], logo_b64: str) -> str:
    objetivo = str(activity.get("objetivo", "")).strip()
    items = activity.get("items", [])
    if not isinstance(items, list):
        items = []

    header_logo = f'<img src="data:image/png;base64,{logo_b64}" style="height:64px;">' if logo_b64 else ""

    html = f"""
    <html>
    <head><style>{GLOBAL_CSS}</style></head>
    <body>
        <div class="paper">
            <div class="header">
                <div>
                    <div class="badges">
                        <span class="badge">Grupo {alumno['grupo']}</span>
                        <span class="badge">Grado {alumno['grado']}</span>
                    </div>
                    <div class="student-name">{alumno['nombre']}</div>
                    <div class="meta">{alumno['perfil']}</div>
                </div>
                <div>{header_logo}</div>
            </div>

            <div class="objective">
                <div style="font-weight:700; margin-bottom:6px;">Objetivo</div>
                <div>{objetivo}</div>
                <div class="small" style="margin-top:10px;">Tiempo sugerido: 60 minutos</div>
            </div>
    """

    for it in items:
        if not isinstance(it, dict):
            continue
        icono = str(it.get("icono", "✍️")).strip() or "✍️"
        enunciado = str(it.get("enunciado", "")).strip()
        pista = str(it.get("pista", "")).strip()
        img_b64 = it.get("img_b64")

        img_html = ""
        if isinstance(img_b64, str) and len(img_b64) > 200:
            img_html = f'<div class="img-box"><img src="data:image/png;base64,{img_b64}" style="max-width:300px;"></div>'

        html += f"""
            <div class="card">
                <div class="enunciado">{icono} {enunciado}</div>
                {img_html}
                <div class="pista">💡 {pista}</div>
                <div class="answer-line"></div>
                <div class="answer-line"></div>
            </div>
        """

    html += "</div></body></html>"
    return html

def html_to_pdf_weasyprint(html: str) -> bytes:
    return HTML(string=html).write_pdf()

def _rl_image_from_b64(b64s: str, max_w: float) -> Optional[RLImage]:
    try:
        raw = base64.b64decode(b64s)
        bio = io.BytesIO(raw)
        img = utils.ImageReader(bio)
        iw, ih = img.getSize()
        scale = min(max_w / float(iw), 1.0)
        w = iw * scale
        h = ih * scale
        bio.seek(0)
        return RLImage(bio, width=w, height=h)
    except Exception:
        return None

def html_to_pdf_reportlab(activity: Dict[str, Any], alumno: Dict[str, str], logo_b64: str) -> bytes:
    # PDF consistente aunque WeasyPrint no exista
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=18*mm,
        rightMargin=18*mm,
        topMargin=18*mm,
        bottomMargin=18*mm,
        title=f"Ficha_{alumno['nombre']}"
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("T", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=16, leading=20, alignment=TA_LEFT, spaceAfter=6)
    meta_style  = ParagraphStyle("M", parent=styles["Normal"], fontName="Helvetica", fontSize=10, leading=13, textColor=HexColor("#374151"), spaceAfter=10)
    h_style     = ParagraphStyle("H", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=12, leading=15, spaceAfter=6)
    p_style     = ParagraphStyle("P", parent=styles["Normal"], fontName="Helvetica", fontSize=11, leading=16, spaceAfter=10)
    pista_style = ParagraphStyle("G", parent=styles["Normal"], fontName="Helvetica-Oblique", fontSize=11, leading=16, textColor=HexColor("#065f46"), spaceAfter=10)

    story: List[Any] = []

    # Header
    story.append(Paragraph(alumno["nombre"], title_style))
    story.append(Paragraph(f"Grupo {alumno['grupo']} · Grado {alumno['grado']}", meta_style))
    story.append(Paragraph(alumno["perfil"], meta_style))
    story.append(HRFlowable(width="100%", thickness=2, color=HexColor(THEME_COLOR)))
    story.append(Spacer(1, 10))

    # Objetivo
    objetivo = str(activity.get("objetivo", "")).strip()
    if objetivo:
        story.append(Paragraph("Objetivo", h_style))
        story.append(Paragraph(objetivo, p_style))
        story.append(Spacer(1, 8))

    items = activity.get("items", [])
    if not isinstance(items, list):
        items = []

    for i, it in enumerate(items, start=1):
        if not isinstance(it, dict):
            continue
        icono = str(it.get("icono", "✍️")).strip() or "✍️"
        enunciado = str(it.get("enunciado", "")).strip()
        pista = str(it.get("pista", "")).strip()
        img_b64 = it.get("img_b64")

        story.append(Paragraph(f"{i}. {icono} {enunciado}", h_style))

        if isinstance(img_b64, str) and len(img_b64) > 200:
            img = _rl_image_from_b64(img_b64, max_w=120*mm)
            if img:
                story.append(img)
                story.append(Spacer(1, 8))

        if pista:
            story.append(Paragraph(f"💡 {pista}", pista_style))

        # líneas de respuesta
        story.append(Spacer(1, 10))
        story.append(HRFlowable(width="100%", thickness=1, color=HexColor("#cbd5e1")))
        story.append(Spacer(1, 14))
        story.append(HRFlowable(width="100%", thickness=1, color=HexColor("#cbd5e1")))
        story.append(Spacer(1, 18))

    doc.build(story)
    return buf.getvalue()

# ============================================================
# PROMPT (Sheets: grupo + dificultad individual SIEMPRE)
# ============================================================

SYSTEM_PROMPT = """
Actúa como Diseñador Instruccional Senior.

Devuelve JSON con el siguiente esquema:
{
  "objetivo": "string",
  "items": [
    {
      "icono": "✍️|📖|🔢|🎨",
      "enunciado": "string",
      "pista": "string",
      "prompt_imagen": "string"
    }
  ]
}

Reglas:
- Actividades con sentido real y progresión.
- Ajusta dificultad SEGÚN PERFIL del alumno y su GRUPO.
- Pistas ultra concretas (micro pasos).
- Imagen ad-hoc: que realmente ayude a resolver el ítem.
- No markdown. No texto fuera del JSON.
""".strip()

def build_prompt(alumno: Dict[str, str], contenido: str) -> str:
    return f"""
{SYSTEM_PROMPT}

ALUMNO:
- nombre: {alumno["nombre"]}
- grupo: {alumno["grupo"]}
- grado: {alumno["grado"]}
- perfil/dificultad: {alumno["perfil"]}

CONTENIDO:
{contenido}

Salida: JSON.
""".strip()

# ============================================================
# UI SIMPLE + PIPELINE INFALIBLE
# ============================================================

def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def main():
    st.title("Opal Classroom v29 INFALIBLE")
    st.caption("HTML premium + PDF infalible (fallback). Sheets obligatorio por alumno (grupo + perfil).")

    try:
        genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    except Exception as e:
        st.error(f"GOOGLE_API_KEY faltante/ inválida: {e}")
        return

    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando Sheets: {e}")
        return

    # Heurística columnas (igual que tus versiones)
    grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
    grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
    perfil_col = df.columns[4] if len(df.columns) > 4 else df.columns[0]

    colA, colB = st.columns([1, 1])

    with colA:
        grado = st.selectbox("Grado", sorted(df[grado_col].dropna().unique().tolist()))
        df_f = df[df[grado_col] == grado].copy()

        alcance = st.radio("Alcance", ["Todo el grado", "Seleccionar alumnos"], horizontal=True)
        if alcance == "Seleccionar alumnos":
            al_sel = st.multiselect("Alumnos", sorted(df_f[alumno_col].dropna().unique().tolist()))
            df_final = df_f[df_f[alumno_col].isin(al_sel)].copy() if al_sel else df_f.iloc[0:0].copy()
        else:
            df_final = df_f

    with colB:
        logo = st.file_uploader("Logo (opcional)", type=["png", "jpg", "jpeg"])
        logo_b64 = base64.b64encode(logo.read()).decode() if logo else ""

        st.write("Salida PDF:")
        if WEASYPRINT_OK:
            st.success("WeasyPrint disponible: PDF premium desde HTML")
        else:
            st.warning("WeasyPrint NO disponible: PDF fallback ReportLab (no crashea)")
            st.caption(f"Detalle: {WEASYPRINT_ERR}")

    st.divider()

    contenido = st.text_area(
        "Prompt (no Ctrl+Enter; botón abajo)",
        height=220,
        placeholder="Ej: Sumas y restas para 1ero usando situaciones cotidianas (caramelos, monedas, juguetes)."
    )

    enable_images = st.checkbox("Generar imágenes por ítem (best-effort)", value=True)

    btn = st.button("🚀 GENERAR LOTE", use_container_width=True)

    if not btn:
        return

    if len(df_final) == 0:
        st.error("No hay alumnos para generar (revisar selección).")
        return

    if not contenido or not contenido.strip():
        st.error("Prompt vacío.")
        return

    zip_buffer = io.BytesIO()
    errores: List[str] = []
    ok_count = 0

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("_RESUMEN.txt", "\n".join([
            "Opal Classroom v29 INFALIBLE",
            f"Inicio: {now_str()}",
            f"Modelo texto: {TEXT_MODEL}",
            f"Modelo imagen: {IMAGE_MODEL}",
            f"WeasyPrint: {'OK' if WEASYPRINT_OK else 'NO'}",
            f"Grado: {grado}",
            f"Alumnos: {len(df_final)}",
        ]))

        prog = st.progress(0.0)
        status = st.empty()

        for i, (_, row) in enumerate(df_final.iterrows(), start=1):
            nombre = str(row[alumno_col]).strip()
            grupo = str(row[grupo_col]).strip()
            perfil = str(row[perfil_col]).strip()

            alumno = {"nombre": nombre, "grupo": grupo, "perfil": perfil, "grado": str(grado)}

            status.info(f"Generando: {nombre} ({i}/{len(df_final)})")

            try:
                prompt = build_prompt(alumno, contenido.strip())
                activity = generate_json_infalleable(TEXT_MODEL, prompt)

                # normalización mínima (no rígida)
                if "items" not in activity or not isinstance(activity.get("items"), list):
                    activity["items"] = []

                # imágenes por ítem
                if enable_images:
                    for it in activity["items"]:
                        if not isinstance(it, dict):
                            continue
                        pv = str(it.get("prompt_imagen", "")).strip()
                        if pv:
                            it["img_b64"] = generar_imagen_base64(pv)
                        else:
                            it["img_b64"] = None

                html = render_html(activity, alumno, logo_b64)

                # PDF (premium si se puede, fallback si no)
                if WEASYPRINT_OK:
                    try:
                        pdf_bytes = html_to_pdf_weasyprint(html)
                    except Exception:
                        pdf_bytes = html_to_pdf_reportlab(activity, alumno, logo_b64)
                else:
                    pdf_bytes = html_to_pdf_reportlab(activity, alumno, logo_b64)

                safe = re.sub(r"[^A-Za-z0-9_\-]+", "_", nombre)[:80] or "ALUMNO"
                zf.writestr(f"{safe}.html", html)
                zf.writestr(f"{safe}.pdf", pdf_bytes)

                ok_count += 1

            except Exception as e:
                msg = f"{nombre} | ERROR {type(e).__name__}: {e}"
                errores.append(msg)
                zf.writestr(f"ERROR_{re.sub(r'[^A-Za-z0-9_\\-]+','_',nombre)[:80]}.txt", msg)

            prog.progress(i / max(1, len(df_final)))

        zf.writestr("_ERRORES.txt", "\n".join(errores) if errores else "Sin errores.")

    status.empty()

    st.success(f"Listo. OK: {ok_count}/{len(df_final)}. Errores: {len(df_final)-ok_count}")
    st.download_button(
        "📥 Descargar ZIP (HTML + PDF)",
        zip_buffer.getvalue(),
        file_name=f"Fichas_{grado}.zip",
        mime="application/zip",
        use_container_width=True
    )

if __name__ == "__main__":
    main()
