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
# Motor Pedagógico Determinista v14.0
# Estrategia: "consultar primero, ejecutar después"
# - Escaneo forzado de modelos ANTES de mostrar UI
# - Selección automática del modelo de texto más potente disponible
# - Selección automática del mejor candidato a imágenes (si existe)
# - ZIP blindado: nunca sale vacío; siempre incluye reporte y resumen
# - JSON estricto + reparación 1 vez + validación (pydantic opcional)
# ============================================================

st.set_page_config(page_title="Motor Pedagógico Determinista v14.0", layout="wide")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

# Forzamos JSON + baja varianza (no garantiza bit-a-bit determinismo)
GEN_CFG_JSON = {
    "response_mime_type": "application/json",
    "temperature": 0,
    "top_p": 1,
    "top_k": 1,
    "max_output_tokens": 4096,
}

SAFETY_SETTINGS = [
    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
]

RETRIES = 6
CACHE_TTL_SECONDS = 6 * 60 * 60

# =========================
# Pydantic opcional
# =========================
PYDANTIC_AVAILABLE = False
try:
    from pydantic import BaseModel, Field

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

# =========================
# Utilidades base
# =========================
def retry(fn):
    last = None
    for i in range(RETRIES):
        try:
            return fn()
        except Exception as e:
            last = e
            if i == RETRIES - 1:
                raise
            time.sleep((2 ** i) + random.uniform(0, 0.6))
    raise last

def hash_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="ignore")).hexdigest()

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

def safe_filename(name: str) -> str:
    s = str(name).strip().replace(" ", "_")
    for ch in ["/", "\\", ":", "*", "?", "\"", "<", ">", "|"]:
        s = s.replace(ch, "_")
    while "__" in s:
        s = s.replace("__", "_")
    if not s:
        s = "SIN_NOMBRE"
    return s[:120]

# =========================
# DOCX extraction (párrafos + tablas)
# =========================
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def _extract_text(el) -> str:
    return "".join(n.text for n in el.iter() if n.tag == f"{W_NS}t" and n.text).strip()

def extract_docx(file) -> str:
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
                    out.append("\t".join(cells))
            out.append("")
    return "\n".join(out).strip()

# =========================
# Modelo: escaneo forzado (antes de UI)
# =========================
def list_models_generate_content() -> List[str]:
    ms = genai.list_models()
    out = []
    for m in ms:
        methods = getattr(m, "supported_generation_methods", []) or []
        if "generateContent" in methods:
            out.append(m.name)  # típicamente "models/...."
    return out

def _rank_text_model(name: str) -> Tuple[int, int, int]:
    """
    Ranking: menor = mejor.
    Prioridad:
    1) gemini-2.5-pro > gemini-2.5-flash > gemini-2.0-pro > gemini-2.0-flash
    2) luego otros gemini/mini
    3) penaliza tts/embedding/audio
    """
    n = name.lower()
    penalty = 0
    for bad in ["tts", "embedding", "embed", "audio", "rerank"]:
        if bad in n:
            penalty += 100

    # base score por familia
    if "gemini-2.5-pro" in n:
        tier = 0
    elif "gemini-2.5-flash" in n:
        tier = 1
    elif "gemini-2.0-pro" in n:
        tier = 2
    elif "gemini-2.0-flash" in n:
        tier = 3
    elif "gemini" in n and "pro" in n:
        tier = 4
    elif "gemini" in n and "flash" in n:
        tier = 5
    elif "mini" in n and "pro" in n:
        tier = 6
    elif "mini" in n and "flash" in n:
        tier = 7
    elif "gemini" in n or "mini" in n:
        tier = 10
    else:
        tier = 20

    # preferir versiones "stable" (sin exp/preview) si hay equivalentes
    exp_penalty = 0
    for bad2 in ["exp", "preview", "experimental"]:
        if bad2 in n:
            exp_penalty += 5

    return (tier + penalty + exp_penalty, len(n), penalty)

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
    # Orden: preferir no-exp si existe; pero muchos image-gen son exp.
    def score(x: str) -> Tuple[int, int]:
        nx = x.lower()
        exp = 1 if ("exp" in nx or "preview" in nx or "experimental" in nx) else 0
        return (exp, len(nx))
    return sorted(cands, key=score)

def smoke_test_image_model(model_id: str) -> Tuple[bool, str]:
    prompt = "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: manzana"
    try:
        m = genai.GenerativeModel(model_id)
        r = retry(lambda: m.generate_content(prompt, safety_settings=SAFETY_SETTINGS))
        cand = r.candidates[0]
        part0 = cand.content.parts[0]
        inline = getattr(part0, "inline_data", None)
        data = getattr(inline, "data", None) if inline else None
        if not data:
            return False, "Respuesta sin inline_data.data"
        if len(data) < 500:
            return False, f"inline_data muy chico ({len(data)} bytes)"
        return True, f"OK bytes={len(data)}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"

def pick_best_image_model(models: List[str]) -> Tuple[Optional[str], str]:
    cands = _candidate_image_models(models)
    if not cands:
        return None, "No se detectaron candidatos de imagen en list_models()."
    for mid in cands[:10]:  # evitamos brute-force excesivo
        ok, msg = smoke_test_image_model(mid)
        if ok:
            return mid, f"Seleccionado por smoke test: {msg}"
    return None, "Se detectaron candidatos, pero ninguno pasó el smoke test."

@st.cache_resource(show_spinner=False)
def forced_model_scan() -> Dict[str, Any]:
    """
    Se ejecuta antes de la UI. Si falla, la app se detiene con error claro.
    """
    # Configurar API key
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])

    # Listar modelos
    models = list_models_generate_content()
    if not models:
        raise RuntimeError("ListModels devolvió 0 modelos con generateContent. API key inválida/limitada o endpoint no compatible.")

    # Elegir texto
    text_model = pick_best_text_model(models)
    if not text_model:
        raise RuntimeError("No se pudo seleccionar modelo de texto automáticamente.")

    # Elegir imagen (best effort)
    img_model, img_reason = pick_best_image_model(models)

    return {
        "models": models,
        "text_model": text_model,
        "image_model": img_model,
        "image_reason": img_reason,
    }

# =========================
# JSON estricto + reparación
# =========================
BASE_PROMPT = """
Devuelve EXCLUSIVAMENTE un JSON válido (sin markdown, sin texto extra).

Esquema:
{
 "alumno": { "nombre": "string", "grupo": "string", "diagnostico": "string" },
 "documento": [
   {
     "tipo": "consigna",
     "enunciado_original": "texto literal completo",
     "pista": "pista pedagógica breve (no dar respuesta)",
     "visual": { "habilitado": boolean, "prompt": "string opcional" }
   }
 ]
}

Reglas:
1) No omitir consignas.
2) enunciado_original debe ser copia fiel (no parafrasear).
3) No dar soluciones.
4) Si visual.habilitado=true, visual.prompt debe empezar EXACTAMENTE con:
   "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: "
5) Nada fuera del JSON.
""".strip()

def build_prompt(nombre: str, diag: str, grupo: str, examen: str) -> str:
    return f"""{BASE_PROMPT}

Alumno:
- nombre: {nombre}
- grupo: {grupo}
- diagnostico: {diag}

EXAMEN:
{examen}
""".strip()

def build_repair_prompt(bad: str, why: str) -> str:
    return f"""
Devuelve EXCLUSIVAMENTE un JSON válido y corregido (sin texto extra).

Problema:
{why}

JSON A CORREGIR:
{bad}

Recuerda: debe cumplir EXACTAMENTE el esquema y reglas. No agregues nada fuera del JSON.
""".strip()

def validate_json(data: Dict[str, Any]) -> Tuple[bool, str]:
    try:
        if PYDANTIC_AVAILABLE:
            AdecuacionModel.model_validate(data)
            return True, "OK(pydantic)"
        if not isinstance(data, dict):
            return False, "Root no es objeto"
        if "alumno" not in data or "documento" not in data:
            return False, "Faltan claves alumno/documento"
        if not isinstance(data["alumno"], dict):
            return False, "alumno no es objeto"
        if not isinstance(data["documento"], list):
            return False, "documento no es lista"
        for i, it in enumerate(data["documento"][:50]):
            if not isinstance(it, dict):
                return False, f"documento[{i}] no es objeto"
            for k in ["enunciado_original", "pista", "visual"]:
                if k not in it:
                    return False, f"documento[{i}] falta {k}"
        return True, "OK(basic)"
    except Exception as e:
        return False, f"Exception validando: {e}"

@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def cached_generate_json(cache_key: str, model_text: str, prompt: str) -> Dict[str, Any]:
    m = genai.GenerativeModel(model_text)
    r = retry(lambda: m.generate_content(
        prompt,
        generation_config=GEN_CFG_JSON,
        safety_settings=SAFETY_SETTINGS
    ))
    return json.loads(r.text)

def request_json(nombre: str, diag: str, grupo: str, examen: str, model_text: str, exam_hash: str) -> Dict[str, Any]:
    prompt = build_prompt(nombre, diag, grupo, examen)
    cache_key = f"{exam_hash}::{model_text}::{nombre}::{grupo}::{diag}"

    # 1) Cache
    try:
        data = cached_generate_json(cache_key, model_text, prompt)
        ok, why = validate_json(data)
        if ok:
            return data
        raise ValueError(f"JSON inválido: {why}")
    except Exception:
        # 2) Direct + repair once
        m = genai.GenerativeModel(model_text)
        r1 = retry(lambda: m.generate_content(
            prompt,
            generation_config=GEN_CFG_JSON,
            safety_settings=SAFETY_SETTINGS
        ))
        raw = getattr(r1, "text", "") or ""
        why = ""
        try:
            data1 = json.loads(raw)
            ok1, why1 = validate_json(data1)
            if ok1:
                return data1
            why = why1
        except Exception as e:
            why = f"No parsea JSON: {e}"

        repair_prompt = build_repair_prompt(raw, why)
        r2 = retry(lambda: m.generate_content(
            repair_prompt,
            generation_config=GEN_CFG_JSON,
            safety_settings=SAFETY_SETTINGS
        ))
        data2 = json.loads(getattr(r2, "text", "") or "")
        ok2, why2 = validate_json(data2)
        if not ok2:
            raise ValueError(f"JSON reparado inválido: {why2}")
        return data2

# =========================
# Imagen (best-effort)
# =========================
def generate_image(model_img: str, prompt: str) -> Optional[io.BytesIO]:
    try:
        m = genai.GenerativeModel(model_img)
        r = retry(lambda: m.generate_content(prompt, safety_settings=SAFETY_SETTINGS))
        cand = r.candidates[0]
        part0 = cand.content.parts[0]
        inline = getattr(part0, "inline_data", None)
        data = getattr(inline, "data", None) if inline else None
        if not data or len(data) < 500:
            return None
        return io.BytesIO(data)
    except Exception:
        return None

# =========================
# Render DOCX
# =========================
def render_docx(data: Dict[str, Any], logo_bytes: Optional[bytes], allow_images: bool, model_img: Optional[str]) -> bytes:
    doc = Document()
    header = doc.add_table(rows=1, cols=2)

    if logo_bytes:
        try:
            header.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(0.85))
        except Exception:
            pass

    info = header.rows[0].cells[1].paragraphs[0]
    info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    alumno = data.get("alumno", {})
    info.add_run(
        f"ALUMNO: {alumno.get('nombre','')}\nGRUPO: {alumno.get('grupo','')} | APOYO: {alumno.get('diagnostico','')}"
    ).bold = True

    green = RGBColor(0, 128, 0)
    required_prefix = "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: "

    for it in data.get("documento", []):
        if not isinstance(it, dict):
            continue
        enun = str(it.get("enunciado_original", "")).strip()
        pista = str(it.get("pista", "")).strip()

        if enun:
            doc.add_paragraph(enun)

        p = doc.add_paragraph()
        r = p.add_run(f"💡 {pista}")
        r.font.color.rgb = green
        r.italic = True
        r.font.size = Pt(10)

        visual = normalize_visual(it.get("visual", {}))
        if allow_images and model_img and visual.get("habilitado") and visual.get("prompt"):
            pv = str(visual.get("prompt", "")).strip()
            if pv and not pv.startswith(required_prefix):
                pv = required_prefix + pv
            img = generate_image(model_img, pv) if pv else None
            if img:
                pic = doc.add_paragraph()
                pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
                try:
                    pic.add_run().add_picture(img, width=Inches(2.6))
                except Exception:
                    pass

        doc.add_paragraph("")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# =========================
# Arranque: escaneo antes de UI
# =========================
def boot_or_stop() -> Dict[str, Any]:
    try:
        return forced_model_scan()
    except Exception as e:
        st.error("Fallo en arranque (escaneo de modelos). La app no continuará hasta resolverlo.")
        st.code(f"{type(e).__name__}: {e}")
        st.stop()

BOOT = boot_or_stop()

# =========================
# UI principal
# =========================
def main():
    st.title("Motor Pedagógico Determinista v14.0")
    st.caption("Consultar primero, ejecutar después: escaneo forzado de modelos al arranque.")

    # Cargar planilla
    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error leyendo planilla CSV: {e}")
        return

    with st.sidebar:
        st.header("Modelos detectados")
        st.write(f"Modelo texto seleccionado: {BOOT['text_model']}")
        if BOOT.get("image_model"):
            st.write(f"Modelo imagen seleccionado: {BOOT['image_model']}")
            st.caption(f"Imagen: {BOOT.get('image_reason','')}")
        else:
            st.write("Modelo imagen seleccionado: N/A")
            st.caption(f"Imagen: {BOOT.get('image_reason','')}")

        st.divider()

        st.header("Controles")
        want_images = st.checkbox("Generar imágenes (si disponible)", value=True)
        allow_images = want_images and (BOOT.get("image_model") is not None)

        if want_images and not BOOT.get("image_model"):
            st.warning("Imágenes desactivadas: tu cuenta/API no expone un modelo de imagen funcional (según smoke test).")

        logo = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
        logo_bytes = logo.read() if logo else None

        st.divider()

        # Columnas por posición (manteniendo tu formato histórico)
        grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
        grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
        diag_col = df.columns[4] if len(df.columns) > 4 else df.columns[0]

        grado = st.selectbox("Grado", sorted(df[grado_col].dropna().unique().tolist()))
        df_f = df[df[grado_col] == grado].copy()

        alcance = st.radio("Alcance", ["Todos", "Seleccionar"], horizontal=True)
        if alcance == "Seleccionar":
            selected = st.multiselect("Alumnos", df_f[alumno_col].dropna().unique().tolist())
            if selected:
                df_f = df_f[df_f[alumno_col].isin(selected)]
            else:
                st.info("No seleccionaste alumnos. No se procesará nada hasta elegir al menos 1.")

        st.caption(f"Pydantic: {'ON' if PYDANTIC_AVAILABLE else 'OFF'}")
        st.caption(f"Modelos visibles (generateContent): {len(BOOT['models'])}")

    file_docx = st.file_uploader("Subir Examen (DOCX)", type=["docx"])
    if not file_docx:
        return

    if st.button("Iniciar procesamiento"):
        exam_text = extract_docx(file_docx)
        if not exam_text.strip():
            st.error("No se pudo extraer texto del DOCX (vacío).")
            return

        exam_hash = hash_text(exam_text)

        alumnos_df = df_f[[alumno_col, grupo_col, diag_col]].dropna(subset=[alumno_col]).copy()
        total = len(alumnos_df)
        if total == 0:
            st.error("No hay alumnos para procesar (filtro/selección vacía).")
            return

        model_text = BOOT["text_model"]
        model_img = BOOT.get("image_model") if allow_images else None

        zip_io = io.BytesIO()
        errors: List[str] = []
        successes = 0

        with zipfile.ZipFile(zip_io, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            meta = []
            meta.append("Motor Pedagógico Determinista v14.0")
            meta.append("Estrategia: consultar primero, ejecutar después")
            meta.append(f"Modelo texto (auto): {model_text}")
            meta.append(f"Modelo imagen (auto): {model_img if model_img else 'N/A'}")
            meta.append(f"Imágenes habilitadas: {bool(model_img)}")
            meta.append(f"Grado: {grado}")
            meta.append(f"Alumnos a procesar: {total}")
            meta.append(f"Hash examen: {exam_hash}")
            meta.append(f"Pydantic: {'ON' if PYDANTIC_AVAILABLE else 'OFF'}")
            zf.writestr("_REPORTE.txt", "\n".join(meta))

            bar = st.progress(0.0)
            status = st.empty()

            for i, (_, row) in enumerate(alumnos_df.iterrows(), start=1):
                nombre = str(row[alumno_col]).strip()
                grupo = str(row[grupo_col]).strip()
                diag = str(row[diag_col]).strip()

                status.text(f"({i}/{total}) Procesando: {nombre}")

                try:
                    data = request_json(nombre, diag, grupo, exam_text, model_text, exam_hash)
                    docx_bytes = render_docx(data, logo_bytes, allow_images=bool(model_img), model_img=model_img)
                    zf.writestr(f"Adecuacion_{safe_filename(nombre)}.docx", docx_bytes)
                    successes += 1
                except Exception as e:
                    msg = f"{nombre} :: {type(e).__name__} :: {e}"
                    errors.append(msg)
                    zf.writestr(f"ERROR_{safe_filename(nombre)}.txt", msg)

                bar.progress(i / total)

            resumen = []
            resumen.append("RESUMEN")
            resumen.append(f"Procesados: {total}")
            resumen.append(f"Exitosos: {successes}")
            resumen.append(f"Con error: {len(errors)}")
            if errors:
                resumen.append("")
                resumen.append("ERRORES:")
                resumen.extend([f"- {e}" for e in errors[:200]])
                if len(errors) > 200:
                    resumen.append(f"... truncado ({len(errors)} errores totales)")
            zf.writestr("_RESUMEN.txt", "\n".join(resumen))

        st.success(f"Listo. Exitosos: {successes} | Errores: {len(errors)}")
        st.download_button("Descargar ZIP", data=zip_io.getvalue(), file_name="adecuaciones.zip", mime="application/zip")

if __name__ == "__main__":
    main()
