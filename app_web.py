import io
import os
import re
import zipfile
import openpyxl
import pandas as pd
from pypdf import PdfReader
import streamlit as st

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA (PRIMERA LÍNEA)
# ==========================================
icono_app = "sobresalir.png" if os.path.exists("sobresalir.png") else "🏠"

st.set_page_config(
    page_title="Auditor de Planillas CEV",
    page_icon=icono_app,
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==========================================
# 2. ESTILOS VISUALES Y CORRECCIÓN DE TEXTOS
# ==========================================
st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');

    /* Tipografía global controlada (evita romper íconos nativos) */
    html, body, .stMarkdown, p, label, h1, h2, h3, h4 {
        font-family: 'Montserrat', sans-serif !important;
    }

    /* Encabezados */
    h1 {
        color: #0b3c61 !important;
        font-weight: 800 !important;
        letter-spacing: -0.5px;
        font-size: 2.1rem !important;
        margin: 0 !important;
        padding-top: 5px !important;
    }

    h3 {
        color: #1e3a8a !important;
        font-weight: 600 !important;
        font-size: 1.25rem !important;
        margin-top: 0.5rem !important;
    }

    /* Contenedor Dropzone de subida */
    [data-testid="stFileUploaderDropzone"] {
        background-color: #ffffff !important;
        border: 2px dashed #00529B !important;
        border-radius: 12px !important;
        padding: 20px 24px !important;
        box-shadow: 0 4px 12px rgba(0, 82, 155, 0.05);
    }

    /* 1. Corrección del botón: elimina 'uploadpload' y coloca 'Subir' */
    [data-testid="stFileUploaderDropzone"] button {
        font-size: 0 !important;
        position: relative !important;
        padding: 8px 24px !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploaderDropzone"] button * {
        display: none !important;
    }
    [data-testid="stFileUploaderDropzone"] button::after {
        content: "Subir";
        font-size: 14px !important;
        font-weight: 600 !important;
        font-family: 'Montserrat', sans-serif !important;
        color: #1e293b !important;
        display: inline-block !important;
        visibility: visible !important;
    }

    /* 2. Reemplazo del texto de límite a 'Peso Máx. 2 GB por Archivo comprimido' */
    [data-testid="stFileUploaderDropzone"] small {
