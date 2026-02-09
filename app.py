import streamlit as st
import google.generativeai as genai
from docx import Document
import io

# 1. Configuración de página e interfaz
st.set_page_config(page_title="Motor de Adecuación Pedagógica", layout="centered")
st.title("Adaptación Automática de Contenidos")
st.write("Carga un examen en .docx y selecciona la dificultad del alumno.")

# 2. Configuración de Gemini
# Los secretos se cargan desde la nube de Streamlit una vez publicado
try:
    genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])
    model = genai.GenerativeModel('gemini-1.5-flash')
except:
    st.error("Falta la configuración de la API Key en los secretos.")

# 3. Selector de diagnósticos
opciones_diagnostico = [
    "Dislexia",
    "Discalculia",
    "Disgrafía",
    "TDAH",
    "Dificultad General"
]
diagnostico = st.selectbox("Seleccione el diagnóstico del alumno:", opciones_diagnostico)

# 4. Carga de archivo
uploaded_file = st.file_uploader("Subir examen original (.docx)", type="docx")

def leer_docx(file):
    doc = Document(file)
    return "\n".join([para.text for para in doc.paragraphs])

if uploaded_file is not None:
    texto_original = leer_docx(uploaded_file)
    
    if st.button("Generar Adecuación"):
        with st.spinner("Gemini está procesando la adecuación..."):
            
            # System Prompt con tus instrucciones específicas
            prompt_sistema = f"""
            Eres un experto en psicopedagogía de primaria. Tu tarea es recibir un examen y devolverlo 
            re-escrito con adecuaciones para un alumno con {diagnostico}.
            
            REGLAS SEGÚN DIAGNÓSTICO:
            - Si es Dislexia: Usa frases cortas, interlineado amplio (simulado con espacios), resalta verbos en negrita, simplifica vocabulario técnico.
            - Si es Discalculia: Desglosa problemas en listas de pasos, sugiere apoyos visuales, usa cuadrículas.
            - Si es Disgrafía: Transforma desarrollo en opción múltiple o completar huecos.
            - Si es TDAH: Una consigna por oración. Fragmenta la tarea. Elimina decoraciones.
            - Si es Dificultad General: Jerarquiza con negritas, incluye un ejemplo resuelto (Ejercicio 0).

            TEXTO ORIGINAL A ADAPTAR:
            {texto_original}
            
            Devuelve el contenido listo para copiar y pegar, manteniendo la estructura de examen.
            """

            try:
                response = model.generate_content(prompt_sistema)
                st.subheader(f"Examen adecuado para: {diagnostico}")
                st.markdown(response.text)
                
                # Opción para copiar/descargar el texto
                st.download_button("Descargar resultado como TXT", response.text, file_name="examen_adecuado.txt")
            except Exception as e:
                st.error(f"Error al procesar: {e}")
