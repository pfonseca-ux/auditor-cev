import io
import re
import zipfile
import openpyxl
import pandas as pd
from pypdf import PdfReader
import streamlit as st

# ==========================================
# CONFIGURACIÓN GENERAL Y ESTILOS VISUALES
# ==========================================
st.set_page_config(
    page_title="Auditor de Planillas CEV",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');

    /* Tipografía global */
    html, body, [class*="css"], .stMarkdown, p, span, label, div {
        font-family: 'Montserrat', sans-serif !important;
    }

    /* Encabezados y títulos */
    h1 {
        color: #0b3c61 !important;
        font-weight: 800 !important;
        letter-spacing: -0.5px;
        font-size: 2.2rem !important;
        margin-bottom: 0.2rem !important;
    }

    h3 {
        color: #1e3a8a !important;
        font-weight: 600 !important;
        font-size: 1.25rem !important;
        margin-top: 0 !important;
    }

    /* Recuadro de subida de archivos (Drag & Drop) */
    [data-testid="stFileUploader"] {
        background-color: #ffffff;
        border: 2px dashed #00529B;
        border-radius: 12px;
        padding: 24px;
        box-shadow: 0 4px 12px rgba(0, 82, 155, 0.06);
        transition: all 0.3s ease;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #002d54;
        box-shadow: 0 6px 16px rgba(0, 82, 155, 0.12);
    }

    /* Botones de acción principales */
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

    /* Tarjetas de métricas */
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
# FUNCIONES DE AUDITORÍA Y EXTRACCIÓN
# ==========================================
def normalizar_rol(rol_crudo):
    """Limpia caracteres y estandariza formato MANZANA-PREDIO."""
    if not rol_crudo:
        return None
    limpio = re.sub(r"[^\d\-–]", "", str(rol_crudo).strip())
    limpio = limpio.replace("–", "-")
    partes = limpio.split("-")
    if len(partes) == 2 and partes[0].isdigit() and partes[1].isdigit():
        return f"{int(partes[0])}-{int(partes[1])}"
    return limpio if limpio else None


def extraer_roles_sii(pdf_bytes):
    """Extrae todos los roles válidos del certificado SII en PDF."""
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
    """Extrae ROL y datos generales de la planilla CEV."""
    datos = {
        "archivo": nombre_archivo,
        "rol": None,
        "comuna": "No detectada",
        "direccion": "No detectada",
        "errores": [],
    }

    try:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True, read_only=True)
        hoja_evaluacion = None

        # Priorizar hojas típicas de la CEV
        nombres_objetivo = ["GENERAL", "DATOS GENERALES", "PROYECTO", "FICHA", "EVALUACION"]
        for nombre in wb.sheetnames:
            if any(objetivo in nombre.upper() for objetivo in nombres_objetivo):
                hoja_evaluacion = wb[nombre]
                break

        if hoja_evaluacion is None:
            hoja_evaluacion = wb.active

        # Búsqueda de campos clave dentro de las primeras filas/columnas
        limite_filas = 50
        for fila in hoja_evaluacion.iter_rows(max_row=limite_filas, values_only=True):
            for i, celda in enumerate(fila):
                if celda is None:
                    continue
                texto_celda = str(celda).strip().upper()

                # Detección de etiqueta ROL
                if "ROL" in texto_celda and not datos["rol"]:
                    for j in range(i + 1, min(i + 4, len(fila))):
                        posible_rol = normalizar_rol(fila[j])
                        if posible_rol:
                            datos["rol"] = posible_rol
                            break

                # Detección de COMUNA
                if "COMUNA" in texto_celda and datos["comuna"] == "No detectada":
                    for j in range(i + 1, min(i + 3, len(fila))):
                        if fila[j] and str(fila[j]).strip():
                            datos["comuna"] = str(fila[j]).strip()
                            break

                # Detección de DIRECCIÓN
                if any(x in texto_celda for x in ["DIRECCION", "DIRECCIÓN", "CALLE"]) and datos["direccion"] == "No detectada":
                    for j in range(i + 1, min(i + 3, len(fila))):
                        if fila[j] and str(fila[j]).strip():
                            datos["direccion"] = str(fila[j]).strip()
                            break

        wb.close()

        if not datos["rol"]:
            datos["errores"].append("No se encontró celda con ROL asignado.")

    except Exception as e:
        datos["errores"].append(f"Archivo ilegible o dañado ({str(e)}).")

    return datos


def generar_excel_reporte(df_resultados):
    """Genera el reporte consolidado en memoria para descarga."""
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
# INTERFAZ DE USUARIO
# ==========================================
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
                    if "__MACOSX" in nombre_interno or nombre_interno.endswith("/"):
                        continue
                    nombre_lower = nombre_interno.lower()
                    if nombre_lower.endswith(".pdf") and ("sii" in nombre_lower or "rol" in nombre_lower or "certificado" in nombre_lower):
                        archivo_pdf_sii = z.read(nombre_interno)
                    elif nombre_lower.endswith((".xlsx", ".xlsm")):
                        archivos_excel.append((nombre_interno, z.read(nombre_interno)))

                if archivo_pdf_sii:
                    roles_sii = extraer_roles_sii(archivo_pdf_sii)

                for nombre_archivo, contenido_bytes in archivos_excel:
                    datos = auditar_planilla_excel(contenido_bytes, nombre_archivo)
                    planillas_datos.append(datos)

            if not planillas_datos:
                st.error("No se encontraron planillas Excel (.xlsx / .xlsm) dentro del paquete ZIP subido.")
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

                    # Validación de correlatividad y unicidad de ROL
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
