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

    html, body, .stMarkdown, p, label, h1, h2, h3, h4 {
        font-family: 'Montserrat', sans-serif !important;
    }

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

    [data-testid="stFileUploaderDropzone"] {
        background-color: #ffffff !important;
        border: 2px dashed #00529B !important;
        border-radius: 12px !important;
        padding: 20px 24px !important;
        box-shadow: 0 4px 12px rgba(0, 82, 155, 0.05);
    }

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
        content: 'Subir' !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        font-family: 'Montserrat', sans-serif !important;
        color: #1e293b !important;
        display: inline-block !important;
        visibility: visible !important;
    }

    [data-testid="stFileUploaderDropzone"] small {
        font-size: 0 !important;
        line-height: 1 !important;
        visibility: hidden !important;
    }
    [data-testid="stFileUploaderDropzone"] small::after {
        content: 'Peso Máx. 2 GB por Archivo comprimido' !important;
        font-size: 13px !important;
        font-weight: 500 !important;
        font-family: 'Montserrat', sans-serif !important;
        color: #64748b !important;
        display: inline-block !important;
        visibility: visible !important;
        margin-left: 4px !important;
    }

    .stButton > button {
        background: linear-gradient(135deg, #00529B 0%, #003366 100%) !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 10px 24px !important;
        box-shadow: 0 4px 10px rgba(0, 51, 102, 0.25) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 14px rgba(0, 51, 102, 0.35) !important;
    }

    [data-testid="stMetric"] {
        background-color: #ffffff;
        border-radius: 10px;
        padding: 16px 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }
</style>
""",
    unsafe_allow_html=True,
)


# ==========================================
# 3. FUNCIONES DE PROCESAMIENTO Y AUDITORÍA
# ==========================================
def normalizar_rol(rol_crudo):
    """Extrae y estandariza el formato MANZANA-PREDIO (ej: 1514-57)."""
    if not rol_crudo:
        return None
    texto = str(rol_crudo).strip()

    # Coincidencia directa de patrón de ROL chileno
    match = re.search(r"\b(\d{1,6})\s*[-–]\s*(\d{1,5})\b", texto)
    if match:
        return f"{int(match.group(1))}-{int(match.group(2))}"

    # Limpieza de caracteres residuales
    limpio = re.sub(r"[^\d\-–]", "", texto).replace("–", "-")
    partes = limpio.split("-")
    if len(partes) == 2 and partes[0].isdigit() and partes[1].isdigit():
        return f"{int(partes[0])}-{int(partes[1])}"
    return None


def extraer_roles_sii(pdf_bytes):
    """Extrae los roles válidos del certificado SII en formato PDF."""
    roles_encontrados = set()
    try:
        lector = PdfReader(io.BytesIO(pdf_bytes))
        texto_completo = ""
        for pagina in lector.pages:
            texto_pagina = pagina.extract_text()
            if texto_pagina:
                texto_completo += " " + texto_pagina

        coincidencias = re.findall(r"\b\d{1,6}\s*[-–]\s*\d{1,5}\b", texto_completo)
        for c in coincidencias:
            rol_norm = normalizar_rol(c)
            if rol_norm:
                roles_encontrados.add(rol_norm)
    except Exception as e:
        st.warning(f"Aviso al leer PDF del SII: {e}")
    return roles_encontrados


def auditar_planilla_excel(excel_bytes, nombre_archivo):
    """Extrae ROL, comuna y dirección revisando prioritariamente hojas CEV."""
    datos = {
        "archivo": nombre_archivo,
        "rol": None,
        "comuna": "No detectada",
        "direccion": "No detectada",
        "errores": [],
    }

    try:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True, read_only=True)

        # Priorizar CEV-CEVE, PBTD y hojas de datos generales
        hojas_ordenadas = []
        for nombre in wb.sheetnames:
            nombre_u = nombre.upper()
            if any(k in nombre_u for k in ["CEV", "PBTD", "GENERAL", "PROYECTO", "FICHA", "DATOS"]):
                hojas_ordenadas.insert(0, wb[nombre])
            else:
                hojas_ordenadas.append(wb[nombre])

        # Búsqueda a través de las hojas ordenadas
        for hoja in hojas_ordenadas:
            if datos["rol"] and datos["comuna"] != "No detectada" and datos["direccion"] != "No detectada":
                break

            for fila in hoja.iter_rows(max_row=80, values_only=True):
                for i, celda in enumerate(fila):
                    if celda is None:
                        continue
                    texto_celda = str(celda).strip().upper()

                    # 1. Detección de ROL
                    if not datos["rol"] and "ROL" in texto_celda:
                        # Revisar si el valor está en la misma celda
                        rol_en_celda = normalizar_rol(texto_celda)
                        if rol_en_celda:
                            datos["rol"] = rol_en_celda
                        else:
                            # Revisar celdas adyacentes a la derecha
                            for j in range(i + 1, min(i + 6, len(fila))):
                                if fila[j] is not None and str(fila[j]).strip():
                                    posible_rol = normalizar_rol(fila[j])
                                    if posible_rol:
                                        datos["rol"] = posible_rol
                                        break

                    # 2. Detección de COMUNA
                    if datos["comuna"] == "No detectada" and "COMUNA" in texto_celda:
                        for j in range(i + 1, min(i + 5, len(fila))):
                            val = fila[j]
                            if val is not None and str(val).strip() and "COMUNA" not in str(val).upper():
                                datos["comuna"] = str(val).strip()
                                break

                    # 3. Detección de DIRECCIÓN
                    if datos["direccion"] == "No detectada" and any(
                        x in texto_celda for x in ["DIRECCION", "DIRECCIÓN", "CALLE", "UBICACION"]
                    ):
                        for j in range(i + 1, min(i + 5, len(fila))):
                            val = fila[j]
                            if val is not None and str(val).strip():
                                val_u = str(val).upper()
                                if not any(x in val_u for x in ["DIRECCION", "DIRECCIÓN", "CALLE"]):
                                    datos["direccion"] = str(val).strip()
                                    break

        wb.close()

        if not datos["rol"]:
            datos["errores"].append("No se encontró celda con ROL asignado.")

    except Exception as e:
        datos["errores"].append(f"Archivo ilegible o dañado ({str(e)}).")

    return datos


def generar_excel_reporte(df_resultados):
    """Genera el reporte consolidado formateado en memoria."""
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df_resultados.to_excel(writer, index=False, sheet_name="Auditoria_CEV")
        hoja = writer.sheets["Auditoria_CEV"]

        for col in hoja.columns:
            longitud_maxima = max(len(str(celda.value or "")) for celda in col)
            nombre_columna = col[0].column_letter
            hoja.column_dimensions[nombre_columna].width = max(longitud_maxima + 4, 15)

    return salida.getvalue()


# ==========================================
# 4. INTERFAZ DE USUARIO
# ==========================================
if os.path.exists("sobresalir.png"):
    col_logo, col_titulo = st.columns([0.06, 0.94])
    with col_logo:
        st.image("sobresalir.png", width=55)
    with col_titulo:
        st.title("Auditor de Planillas - Plataforma Evaluadores CEV")
else:
    st.title("🏠 Auditor de Planillas - Plataforma Evaluadores CEV")

st.markdown("### Control previo para carga masiva en el Sistema de Calificación Energética de Viviendas")
st.markdown(
    "Herramienta de control de calidad, correlatividad de ROL, consistencia de datos generales "
    "y validación automática contra el Certificado de Asignación de Roles del SII."
)

archivo_zip = st.file_uploader(
    "📁 Sube aquí el paquete del proyecto (.ZIP) con las planillas oficiales y el Certificado SII",
    type=["zip"],
    help="El archivo comprimido puede contener carpetas, subcarpetas, archivos .xlsx, .xlsm y el PDF del SII.",
)

if archivo_zip is not None:
    if st.button("🚀 Auditar Lote de Planillas CEV"):
        with st.spinner("Procesando paquete comprimido y cruzando información..."):
            roles_sii = set()
            planillas_datos = []
            archivos_excel = []
            archivo_pdf_sii = None

            with zipfile.ZipFile(archivo_zip, "r") as z:
                for nombre_interno in z.namelist():
                    # Ignorar metadatos de macOS, carpetas vacías y archivos temporales de Excel
                    if "__MACOSX" in nombre_interno or nombre_interno.endswith("/") or "/.~" in nombre_interno or nombre_interno.startswith("~"):
                        continue

                    nombre_lower = nombre_interno.lower()

                    # Ignorar reportes de auditoría descargados previamente
                    if "reporte_auditoria" in nombre_lower:
                        continue

                    if nombre_lower.endswith(".pdf") and any(
                        palabra in nombre_lower for palabra in ["sii", "rol", "certificado"]
                    ):
                        archivo_pdf_sii = z.read(nombre_interno)
                    elif nombre_lower.endswith((".xlsx", ".xlsm")):
                        archivos_excel.append((nombre_interno, z.read(nombre_interno)))

                if archivo_pdf_sii:
                    roles_sii = extraer_roles_sii(archivo_pdf_sii)

                for nombre_archivo, contenido_bytes in archivos_excel:
                    datos = auditar_planilla_excel(contenido_bytes, nombre_archivo)
                    planillas_datos.append(datos)

            if not planillas_datos:
                st.error("No se encontraron planillas Excel (.xlsx / .xlsm) válidas dentro del paquete ZIP subido.")
            else:
                conteo_roles = {}
                for p in planillas_datos:
                    if p["rol"]:
                        conteo_roles[p["rol"]] = conteo_roles.get(p["rol"], 0) + 1

                filas_resumen = []
                total_aprobadas = 0
                total_observadas = 0

                for p in planillas_datos:
                    errores = list(p["errores"])

                    if p["rol"]:
                        if conteo_roles[p["rol"]] > 1:
                            errores.append(f"ROL duplicado en lote ({conteo_roles[p['rol']]} planillas con ROL {p['rol']}).")
                        if roles_sii and p["rol"] not in roles_sii:
                            errores.append(f"ROL {p['rol']} no figura en el Certificado SII.")

                    estado = "❌ Observado" if errores else "✅ Listo para CEV"
                    if errores:
                        total_observadas += 1
                    else:
                        total_aprobadas += 1

                    filas_resumen.append(
                        {
                            "Archivo": p["archivo"],
                            "ROL Detectado": p["rol"] if p["rol"] else "No detectado",
                            "Comuna": p["comuna"],
                            "Dirección": p["direccion"],
                            "Estado": estado,
                            "Detalle Observaciones": " | ".join(errores) if errores else "Sin inconsistencias detectadas",
                        }
                    )

                df_resultados = pd.DataFrame(filas_resumen)

                st.write("")
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Planillas Analizadas", len(planillas_datos))
                col2.metric("Aprobadas para Carga", total_aprobadas)
                col3.metric("Con Observaciones", total_observadas)

                st.write("")
                excel_descargable = generar_excel_reporte(df_resultados)
                st.download_button(
                    label="📥 Descargar Reporte de Auditoría (.xlsx)",
                    data=excel_descargable,
                    file_name="Reporte_Auditoria_CEV.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.write("")
                st.dataframe(df_resultados, use_container_width=True)
