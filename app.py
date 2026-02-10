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
# Motor Pedagógico v14.6 (Super Robusto)
# Estrategia: Consultar primero, ejecutar después
# - Escaneo forzado de modelos ANTES de UI (cache_resource)
# - Selección automática del mejor modelo de texto disponible (ranking)
# - Selección automática del modelo de imagen (si existe) por smoke test
# - JSON estricto (application/json) + validación + reparación 1 vez
# - Retries con exponential backoff en errores típicos (429/5xx/timeouts)
# - ZIP blindado: siempre incluye _REPORTE.txt y _RESUMEN.txt
# - Error por alumno en archivo ERROR_*.txt
# - Render inclusivo (Verdana + interlineado) + pistas verdes
# ============================================================

st.set_page_config(page_title="Motor Pedagógico v14.6", layout="wide")

SHEET_ID = "1dCZdGmK765ceVwTqXzEAJCrdSvdNLBw7t3q5Cq1Qrww"
URL_PLANILLA = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

RETRIES = 6
CACHE_TTL_SECONDS = 6 * 60 * 60

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

# ============================================================
# Pydantic (opcional)
# ============================================================
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

# ============================================================
# Prompt de tutor (JSON estricto)
# ============================================================
SYSTEM_PROMPT = """
Actúa como un Tutor Psicopedagogo de nivel primario.
TU OBJETIVO: Intervenir el examen para que el alumno razone, NO para que lo resuelva la IA.

PROCESO (interno):
1) Resuelve internamente cada consigna para conocer la respuesta correcta.
2) Diseña una pista 💡 para guiar el razonamiento sin revelar la respuesta.
3) Adapta por diagnóstico y grupo. Si Grupo A o Discalculia/Dislexia: lenguaje concreto y apoyos visuales.

SALIDA:
Devuelve EXCLUSIVAMENTE un JSON válido (sin markdown, sin texto extra).

ESQUEMA:
{
  "alumno": { "nombre": "string", "grupo": "string", "diagnostico": "string" },
  "documento": [
    {
      "tipo": "consigna",
      "enunciado_original": "copia fiel literal del examen",
      "pista": "pista pedagógica (sin dar respuesta)",
      "visual": { "habilitado": boolean, "prompt": "string opcional" }
    }
  ]
}

REGLAS:
- Transcribe el 100% de los enunciados originales (copia fiel).
- No des respuestas ni soluciones.
- Si visual.habilitado=true, visual.prompt debe empezar EXACTAMENTE con:
  "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: "
- Usa 🔢, 📖, ✍️ dentro del texto si ayuda a organizar, pero siempre dentro de los strings del JSON.
""".strip()

IMAGE_PROMPT_PREFIX = "Dibujo escolar, trazos negros, fondo blanco, estilo simple de: "

# ============================================================
# Utilidades
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
    raise last  # pragma: no cover

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

# ============================================================
# DOCX extraction (párrafos + tablas, preserva orden)
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

# ============================================================
# Model scanning + selection (consultar primero)
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

    # ranking “potencia” (mejor = menor)
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

    # penaliza “exp/preview” si hay alternativas
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

def smoke_test_image_model(model_id: str) -> Tuple[bool, str]:
    prompt = f"{IMAGE_PROMPT_PREFIX} manzana"
    try:
        m = genai.GenerativeModel(model_id)
        r = retry_with_backoff(lambda: m.generate_content(prompt, safety_settings=SAFETY_SETTINGS))
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

    # testea hasta 10 para no hacer brute force excesivo
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
# JSON generation + validación + reparación 1 vez
# ============================================================
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

        # Validación mínima de contenidos
        al = data["alumno"]
        for k in ["nombre", "grupo", "diagnostico"]:
            if k not in al or not isinstance(al[k], str):
                return False, f"alumno.{k} inválido"

        for i, it in enumerate(data["documento"][:200]):
            if not isinstance(it, dict):
                return False, f"documento[{i}] no es objeto"
            for k in ["enunciado_original", "pista", "visual"]:
                if k not in it:
                    return False, f"documento[{i}] falta {k}"
            if not isinstance(it["enunciado_original"], str) or not it["enunciado_original"].strip():
                return False, f"documento[{i}].enunciado_original vacío"
            if not isinstance(it["pista"], str) or not it["pista"].strip():
                return False, f"documento[{i}].pista vacío"
            if not isinstance(it["visual"], dict):
                return False, f"documento[{i}].visual no es objeto"
        return True, "OK(basic)"
    except Exception as e:
        return False, f"Exception validando: {e}"

def build_prompt(nombre: str, diag: str, grupo: str, examen: str) -> str:
    return f"{SYSTEM_PROMPT}\n\nALUMNO: {nombre}\nGRUPO: {grupo}\nDIAGNOSTICO: {diag}\n\nEXAMEN:\n{examen}".strip()

def build_repair_prompt(bad: str, why: str) -> str:
    return f"""
Devuelve EXCLUSIVAMENTE un JSON válido y corregido (sin texto extra).

Problema:
{why}

JSON A CORREGIR:
{bad}

Recuerda: debe cumplir el esquema y reglas del sistema. No agregues nada fuera del JSON.
""".strip()

@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def cached_generate_json(cache_key: str, model_text: str, prompt: str) -> Dict[str, Any]:
    m = genai.GenerativeModel(model_text)
    r = retry_with_backoff(lambda: m.generate_content(
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
        r1 = retry_with_backoff(lambda: m.generate_content(
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
        r2 = retry_with_backoff(lambda: m.generate_content(
            repair_prompt,
            generation_config=GEN_CFG_JSON,
            safety_settings=SAFETY_SETTINGS
        ))
        data2 = json.loads(getattr(r2, "text", "") or "")
        ok2, why2 = validate_json(data2)
        if not ok2:
            raise ValueError(f"JSON reparado inválido: {why2}")
        return data2

# ============================================================
# Imagen (best effort) + validación de bytes
# ============================================================
def generar_imagen_ia(model_id: str, prompt_img: str) -> Optional[io.BytesIO]:
    try:
        m = genai.GenerativeModel(model_id)
        r = retry_with_backoff(lambda: m.generate_content(prompt_img, safety_settings=SAFETY_SETTINGS))
        cand = r.candidates[0]
        part0 = cand.content.parts[0]
        inline = getattr(part0, "inline_data", None)
        data = getattr(inline, "data", None) if inline else None
        if not data or len(data) < 500:
            return None
        return io.BytesIO(data)
    except Exception:
        return None

# ============================================================
# Render DOCX (inclusivo)
# ============================================================
def renderizar_adecuacion(data_json: Dict[str, Any], logo_bytes: Optional[bytes], activar_img: bool, model_img_id: Optional[str]) -> bytes:
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Verdana"
    style.font.size = Pt(11)

    header = doc.add_table(rows=1, cols=2)
    if logo_bytes:
        try:
            header.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(logo_bytes), width=Inches(0.85))
        except Exception:
            pass

    info = header.rows[0].cells[1].paragraphs[0]
    info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    al = data_json.get("alumno", {})
    info.add_run(f"ALUMNO: {al.get('nombre','')}\nGRUPO: {al.get('grupo','')} | APOYO: {al.get('diagnostico','')}").bold = True

    for item in data_json.get("documento", []):
        if not isinstance(item, dict):
            continue

        enun = str(item.get("enunciado_original", "")).strip()
        pista = str(item.get("pista", "")).strip()
        visual = normalize_visual(item.get("visual", {}))

        p_orig = doc.add_paragraph(enun)
        p_orig.paragraph_format.line_spacing = 1.5

        p_pista = doc.add_paragraph()
        run = p_pista.add_run(f"💡 {pista}")
        run.font.color.rgb = RGBColor(0, 128, 0)
        run.italic = True
        run.font.size = Pt(10)

        if activar_img and model_img_id and visual.get("habilitado"):
            pv = str(visual.get("prompt", "")).strip()
            if pv:
                if not pv.startswith(IMAGE_PROMPT_PREFIX):
                    pv = IMAGE_PROMPT_PREFIX + pv
                img_data = generar_imagen_ia(model_img_id, pv)
                if img_data:
                    pic = doc.add_paragraph()
                    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    try:
                        pic.add_run().add_picture(img_data, width=Inches(2.5))
                    except Exception:
                        pass

        doc.add_paragraph("")

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# ============================================================
# UI + Proceso
# ============================================================
def main():
    st.title("Motor Pedagógico Determinista v14.6")
    st.caption("Super robusto: escaneo forzado de modelos antes de UI, JSON estricto con reparación, ZIP blindado.")

    # Cargar planilla
    try:
        df = pd.read_csv(URL_PLANILLA)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Error cargando planilla: {e}")
        return

    # Column mapping por posición (mantiene tu esquema histórico)
    grado_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    alumno_col = df.columns[2] if len(df.columns) > 2 else df.columns[0]
    grupo_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]
    diag_col = df.columns[4] if len(df.columns) > 4 else df.columns[0]

    with st.sidebar:
        st.header("Arranque (consultar primero)")
        st.write(f"Boot: {BOOT.get('boot_time')}")
        st.write(f"Modelo texto (auto): {BOOT.get('text_model')}")
        if BOOT.get("image_model"):
            st.write(f"Modelo imagen (auto): {BOOT.get('image_model')}")
            st.caption(BOOT.get("image_reason", ""))
        else:
            st.write("Modelo imagen (auto): N/A")
            st.caption(BOOT.get("image_reason", ""))

        st.divider()
        st.header("Selección")
        grado = st.selectbox("Grado", df[grado_col].dropna().unique())
        df_f = df[df[grado_col] == grado].copy()

        alcance = st.radio("Adecuar para:", ["Todo el grado", "Seleccionar alumnos"], horizontal=True)
        alumnos_lista = df_f[alumno_col].dropna().unique().tolist()
        if alcance == "Seleccionar alumnos":
            seleccion = st.multiselect("Alumnos", alumnos_lista)
            if seleccion:
                alumnos_final = df_f[df_f[alumno_col].isin(seleccion)].copy()
            else:
                alumnos_final = df_f.iloc[0:0].copy()
                st.info("Seleccioná al menos 1 alumno para procesar.")
        else:
            alumnos_final = df_f

        st.divider()
        st.header("Salida / Assets")
        activar_img_user = st.checkbox("Generar imágenes IA (si disponible)", value=True)
        activar_img = activar_img_user and (BOOT.get("image_model") is not None)
        if activar_img_user and not activar_img:
            st.warning("Imágenes desactivadas: no se detectó un modelo de imagen funcional en tu cuenta.")

        logo = st.file_uploader("Logo", type=["png", "jpg", "jpeg"])
        l_bytes = logo.read() if logo else None

        st.caption(f"Pydantic: {'ON' if PYDANTIC_AVAILABLE else 'OFF'}")

    archivo = st.file_uploader("Subir Examen Base (DOCX)", type=["docx"])
    if not archivo:
        return

    if st.button("🚀 INICIAR PROCESAMIENTO"):
        if len(alumnos_final) == 0:
            st.error("No hay alumnos para procesar (selección vacía).")
            return

        txt_base = extraer_texto_docx(archivo)
        if not txt_base.strip():
            st.error("No se pudo extraer texto del DOCX (vacío).")
            return

        exam_hash = hash_text(txt_base)
        model_text = BOOT["text_model"]
        model_img = BOOT.get("image_model") if activar_img else None

        zip_io = io.BytesIO()
        logs: List[str] = []
        errors: List[str] = []
        ok_count = 0

        logs.append(f"Inicio: {now_str()}")
        logs.append(f"Modelo texto: {model_text}")
        logs.append(f"Modelo imagen: {model_img if model_img else 'N/A'}")
        logs.append(f"Imágenes habilitadas: {bool(model_img)}")
        logs.append(f"Grado: {grado}")
        logs.append(f"Total alumnos: {len(alumnos_final)}")
        logs.append(f"Hash examen: {exam_hash}")
        logs.append(f"Pydantic: {'ON' if PYDANTIC_AVAILABLE else 'OFF'}")
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

                status.info(f"Procesando: {n} ({i}/{total})")

                try:
                    data = request_json(n, d, g, txt_base, model_text, exam_hash)
                    docx_bytes = renderizar_adecuacion(data, l_bytes, activar_img=bool(model_img), model_img_id=model_img)
                    zf.writestr(f"Adecuacion_{safe_filename(n)}.docx", docx_bytes)
                    ok_count += 1
                except Exception as e:
                    msg = f"{n} :: {type(e).__name__} :: {e}"
                    errors.append(msg)
                    zf.writestr(f"ERROR_{safe_filename(n)}.txt", msg)

                prog.progress(i / total)

            resumen = []
            resumen.append("RESUMEN")
            resumen.append(f"Fin: {now_str()}")
            resumen.append(f"Procesados: {total}")
            resumen.append(f"OK: {ok_count}")
            resumen.append(f"Errores: {len(errors)}")
            if errors:
                resumen.append("")
                resumen.append("ERRORES (primeros 200):")
                resumen.extend([f"- {e}" for e in errors[:200]])
                if len(errors) > 200:
                    resumen.append(f"... truncado ({len(errors)} errores totales)")
            zf.writestr("_RESUMEN.txt", "\n".join(resumen))

        st.success(f"Lote finalizado. OK: {ok_count} | Errores: {len(errors)}")
        st.download_button("📥 Descargar Adecuaciones (ZIP)", zip_io.getvalue(), f"adecuaciones_v14_6.zip", mime="application/zip")

if __name__ == "__main__":
    main()
