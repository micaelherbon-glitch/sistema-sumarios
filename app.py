import streamlit as st
import os
import json
import datetime
import io
import re
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image
import google.generativeai as genai

# ═════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE PÁGINA
# ═════════════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Sistema de Sumarios", page_icon="🛡️", layout="wide")

APP_ORG = "Comisaría Tres de Febrero — Seccional 1ra."
MODEL_NAME = "gemini-2.5-flash"
PLANTILLA_PATH = "plantilla.docx"

# ═════════════════════════════════════════════════════════════════════════════
# ESTADO DE LA SESIÓN (Para mantener datos entre pestañas)
# ═════════════════════════════════════════════════════════════════════════════
if "datos_extraidos" not in st.session_state:
    st.session_state.datos_extraidos = {}
if "sintesis" not in st.session_state:
    st.session_state.sintesis = ""
if "denuncia_texto" not in st.session_state:
    st.session_state.denuncia_texto = ""
if "visu_texto" not in st.session_state:
    st.session_state.visu_texto = ""

# ═════════════════════════════════════════════════════════════════════════════
# FUNCIONES DE IA
# ═════════════════════════════════════════════════════════════════════════════
def configurar_ia():
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        st.error("No se encontró la API Key en los secretos de Streamlit.")
        st.stop()
    genai.configure(api_key=api_key)
    return genai.GenerativeModel(MODEL_NAME)

def extraer_datos(texto, modelo):
    PROMPT_EXTRACCION = """Actuá como redactor judicial policial de la Provincia de Buenos Aires (Argentina).
    Analizá la siguiente denuncia o acta y devolvé ÚNICAMENTE un JSON válido, sin texto adicional...
    (Mismo prompt exacto de tu código original adaptado para resumir aquí por brevedad. Asegúrate de pegar tu prompt completo aquí)
    """
    resp = modelo.generate_content(PROMPT_EXTRACCION + texto)
    t = resp.text.strip()
    t = re.sub(r"^```(?:json)?", "", t, flags=re.IGNORECASE).strip()
    t = re.sub(r"```$", "", t).strip()
    return json.loads(t)

# ═════════════════════════════════════════════════════════════════════════════
# INTERFAZ WEB
# ═════════════════════════════════════════════════════════════════════════════
st.title("🛡️ SISTEMA DE SUMARIOS")
st.subheader(APP_ORG)

modelo_ia = configurar_ia()

tab1, tab2, tab3, tab4 = st.tabs(["1. Denuncia", "2. Datos", "3. Visu y Fotos", "4. Generar Documento"])

# --- PESTAÑA 1: DENUNCIA ---
with tab1:
    st.markdown("### Captura de Denuncia")
    denuncia_input = st.text_area("Pegue aquí el texto completo de la denuncia o acta oficial", height=250)
    
    if st.button("Extraer datos con IA"):
        if denuncia_input:
            with st.spinner("Procesando con IA..."):
                try:
                    datos = extraer_datos(denuncia_input, modelo_ia)
                    st.session_state.datos_extraidos = datos
                    st.session_state.sintesis = datos.get("sintesis_breve_comunicado", "")
                    st.session_state.denuncia_texto = denuncia_input
                    st.success("Datos extraídos correctamente. Revisa la pestaña 'Datos'.")
                except Exception as e:
                    st.error(f"Error al procesar: {e}")
        else:
            st.warning("Ingrese el texto de la denuncia.")
            
    st.text_area("Síntesis generada (Editable)", value=st.session_state.sintesis, height=150, key="sintesis_edit")

# --- PESTAÑA 2: DATOS ---
with tab2:
    st.markdown("### Revisión de Datos Extraídos")
    d = st.session_state.datos_extraidos.get("datos_extract", {})
    
    col1, col2 = st.columns(2)
    with col1:
        caratula = st.text_input("Carátula", value=d.get("caratula", ""))
        ipp = st.text_input("N° IPP", value=d.get("ipp_medio", ""))
        denunciante = st.text_input("Denunciante", value=d.get("denunciante_nombre", ""))
    with col2:
        ufi = st.text_input("UFI N°", value=d.get("ufi", ""))
        lugar = st.text_input("Lugar del hecho", value=d.get("lugar_hecho_corto", ""))
        victima = st.text_input("Víctima", value=d.get("victima_nombre", ""))

    st.markdown("**(Agrega el resto de los campos necesarios usando `st.text_input` o `st.text_area` mapeando a tu diccionario)**")

# --- PESTAÑA 3: VISU Y FOTOS ---
with tab3:
    st.markdown("### Inspección de Visu y Fotografías")
    visu_input = st.text_area("Descripción de Visu", value=st.session_state.visu_texto, height=150)
    fotos_subidas = st.file_uploader("Subir fotografías (JPG, PNG)", accept_multiple_files=True, type=['jpg', 'png', 'jpeg'])
    
    if fotos_subidas:
        st.write(f"{len(fotos_subidas)} fotografías cargadas.")

# --- PESTAÑA 4: GENERAR ---
with tab4:
    st.markdown("### Generar Documento Word")
    if st.button("Generar Documento"):
        if not os.path.exists(PLANTILLA_PATH):
            st.error("No se encontró plantilla.docx en el servidor.")
        else:
            with st.spinner("Generando documento..."):
                # Aquí insertas la lógica de docxtpl (doc = DocxTemplate(PLANTILLA_PATH))
                # doc.render(contexto)
                # ...
                
                # En lugar de guardar en disco, guardar en memoria para descarga
                doc = DocxTemplate(PLANTILLA_PATH)
                # doc.render(ctx) <- Reemplazar ctx por tu diccionario de variables
                
                bio = io.BytesIO()
                doc.save(bio)
                
                st.download_button(
                    label="📥 Descargar Sumario Word",
                    data=bio.getvalue(),
                    file_name=f"SUMARIO_{caratula}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )