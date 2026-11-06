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
if os.path.exists("sobresalir.svg"):
    icono_app = "sobresalir.svg"
elif os.path.exists("sobresalir.png"):
    icono_app = "sobresalir.png"
else:
    icono_app = "🏠"

st.set_page_config(
    page_title="Auditor de Planillas CEV",
    page_icon=icono_app,
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==========================================
# 2. ESTILOS VISUALES Y CONTROL DE TEXTOS
# ==========================================
st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');

    html, body, .stMarkdown, p, label, h1, h2, h3, h4 {
        font-family: 'Montserrat', sans-serif !important;
    }

    /* Ocultar barra superior y menú de Streamlit */
    header[data-testid="stHeader"] {
        display: none !important;
    }

    /* Ocultar indicador de ejecución ("Stop") */
    [data-testid="stStatusWidget"] {
        display: none !important;
        visibility: hidden !important;
    }

    /* Ocultar pie de página e insignias flotantes */
    footer,
    [class*="viewerBadge"],
    [data-testid="stManageAppButton"],
    #manage-app-button,
    button[id*="manage-app"] {
        display: none !important;
        visibility: hidden !important;
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

    /* Caja contenedora del Uploader */
    [data-
