import datetime
import io
import re
import shutil
import tempfile
import unicodedata
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Auditor de Planillas CEV",
    layout="wide",
    page_icon="🏠"
)

st.title("🏠 Auditor de Planillas - Plataforma Evaluadores CEV")
st.subheader("Control previo para carga masiva en el Sistema de Calificación Energética de Viviendas")
st.markdown(
    "Herramienta de control de calidad, unicidad de ROL, consistencia de datos generales "
    "y validación automática contra el Certificado de Asignación de Roles del SII."
)

# --- REGLAS Y FUNCIONES DE AUDITORÍA CEV ---
def normalizar(texto: str) -> str:
    if not texto: return ""
    texto_str = str(texto).strip()
    texto_nfd = unicodedata.normalize("NFD", texto_str)
    sin_tildes = "".join(c for c in texto_nfd if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", sin_tildes).strip().lower()

MATCHERS = {
    "tipo_calificacion": ["tipo de calificacion", "tipo calificacion"],
    "region": ["region"],
    "comuna": ["comuna"],
    "zona_termica": ["zona termica proyecto", "zona termica"],
    "dormitorios": ["dormitorios de la vivienda", "dormitorios"],
    "identificacion": ["identificacion de la vivienda", "identificacion de la"],
    "proyecto": ["nombre del proyecto"],
    "direccion": ["direccion de la vivienda", "direccion"],
    "tipo_vivienda": ["tipo de vivienda"],
    "rol_vivienda": ["rol vivienda", "rol de la vivienda"],
    "evaluador": ["evaluador energetico", "evaluador"],
    "rut_evaluador": ["rut evaluador", "rol registro de evaluadores"],
    "solicitado_por": ["solicitado por"],
    "rut_mandante": ["rut mandante"],
}

CAMPOS_COMUNES_PROYECTO = [
    ("proyecto", "Nombre del Proyecto"),
    ("direccion", "Dirección"),
    ("comuna", "Comuna"),
    ("region", "Región"),
    ("zona_termica", "Zona Térmica"),
    ("tipo_calificacion", "Tipo de Calificación"),
    ("evaluador", "Evaluador Energético"),
    ("rut_evaluador", "RUT Evaluador"),
    ("solicitado_por", "Solicitado por"),
    ("rut_mandante", "RUT Mandante"),
]

ERRORES_EXCEL = {"#VALOR!", "#¡VALOR!", "#REF!", "#N/A", "#¿NOMBRE?", "#DIV/0!", "#NAME?", "#N/D"}

def extraer_roles_desde_pdf_sii(carpeta: Path):
    pdfs = list(carpeta.rglob("*.pdf"))
    if not pdfs: return {}, None
    roles_map, pdf_usado = {}, None
    for pdf_path in pdfs:
        try:
            reader = pypdf.PdfReader(str(pdf_path))
            txt_norm = normalizar(reader.pages[0].extract_text()) if len(reader.pages) > 0 else ""
            if any(t in txt_norm for t in ["roles de avaluo", "asignacion de roles", "f 2893", "f2893"]):
                pdf_usado = pdf_path.name
                for page in reader.pages:
                    for linea in page.extract_text().splitlines():
                        m_rol = re.search(r"\b(\d{4,6}-\d{4,6})\b", linea)
                        if m_rol:
                            rol = m_rol.group(1)
                            m_u = re.search(r"(?:DEP|DPTO|CASA)[-\s]*([0-9]+[A-Za-z]?)", linea, re.IGNORECASE)
                            m_e = re.search(r"(?:EDIF|TORRE|BLOQUE)[-\s]*([0-9A-Za-z]+)", linea, re.IGNORECASE)
                            if m_u:
                                u_num = m_u.group(1).upper()
                                e_letra = m_e.group(1).upper() if m_e else "GENERAL"
                                roles_map[(e_letra, u_num)] = rol
                                roles_map[("GENERAL", u_num)] = rol
                break
        except Exception:
            continue
    return roles_map, pdf_usado

def descomponer_unidad(texto: str, ruta_archivo: Path = None):
    texto = texto or ""
    m_num = re.search(r"(?:depto\.?|departamento|casa|dpto\.?)\s*([0-9]+[a-zA-Z]?)", texto, re.IGNORECASE)
    if not m_num:
        m_num = re.search(r"\b([0-9]{2,4}[a-zA-Z]?)\b", texto)
    m_torre = re.search(r"(?:torre|edificio|bloque|condominio)\s*([a-zA-Z0-9]+)", texto, re.IGNORECASE)
    
    num = m_num.group(1).upper() if m_num else None
    torre = m_torre.group(1).upper() if m_torre else None
    if not torre and ruta_archivo:
        for parte in ruta_archivo.parts:
            m_t_ruta = re.search(r"(?:torre|edificio|bloque)\s*([a-zA-Z0-9]+)", parte, re.IGNORECASE)
            if m_t_ruta:
                torre = m_t_ruta.group(1).upper()
                break
    return num, (torre or "GENERAL")

def extraer_datos_completos(ruta_excel: Path):
    try:
        wb = openpyxl.load_workbook(ruta_excel, read_only=True, data_only=True)
    except Exception as e:
        return None, f"No se pudo abrir el archivo: {e}"

    nombres = {normalizar(s): s for s in wb.sheetnames}
    hoja = nombres.get("cev-ceve") or nombres.get("cev ceve") or next((s for s in wb.sheetnames if "CEV" in s.upper()), None)
    if not hoja:
        wb.close()
        return None, "Falta pestaña oficial CEV-CEVE"

    ws = wb[hoja]
    filas = list(ws.iter_rows(max_row=40, values_only=True))
    datos = {}
    for row in filas:
        row_str = [str(c).strip() if c is not None else "" for c in row]
        row_norm = [normalizar(c) for c in row_str]
        for key, tokens in MATCHERS.items():
            if key in datos: continue
            for idx, c_norm in enumerate(row_norm):
                if any(t in c_norm for t in tokens):
                    for next_idx in range(idx + 1, len(row_str)):
                        if row_str[next_idx] not in ("", "None"):
                            datos[key] = row_str[next_idx]
                            break
                    break
    wb.close()
    return datos, None

# --- PANEL DE CONTROL ---
archivo_zip = st.file_uploader(
    "📂 Sube aquí el paquete del proyecto (.ZIP) con las planillas oficiales y el Certificado SII",
    type=["zip"]
)

if archivo_zip:
    if st.button("🚀 Auditar Lote de Planillas CEV", type="primary", use_container_width=True):
        temp_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(archivo_zip) as z:
            z.extractall(temp_dir)

        carpeta_raiz = Path(temp_dir)
        roles_sii, pdf_usado = extraer_roles_desde_pdf_sii(carpeta_raiz)
        archivos = [
            f for f in sorted(list(carpeta_raiz.rglob("*.xlsm")) + list(carpeta_raiz.rglob("*.xlsx")))
            if not f.name.startswith("~$") and not f.name.lower().startswith("reporte_auditoria")
        ]

        if not archivos:
            st.error("⚠️ No se encontraron planillas (.xlsx o .xlsm) dentro del archivo ZIP.")
            shutil.rmtree(temp_dir, ignore_errors=True)
        else:
            if pdf_usado:
                st.success(f"📄 Certificado SII detectado y procesado: **{pdf_usado}** ({len(roles_sii)} unidades mapeadas)")
            else:
                st.warning("ℹ️ No se detectó certificado SII en PDF dentro del lote. Se omite el cruce oficial de roles.")

            datos_por_archivo = {}
            errores_por_archivo = defaultdict(list)
            tipologia_errores_por_grupo = defaultdict(lambda: defaultdict(int))
            valores_proyecto = {clave: Counter() for clave, _ in CAMPOS_COMUNES_PROYECTO}
            roles_vistos = defaultdict(list)
            ident_vistas = defaultdict(list)

            progreso = st.progress(0, text="Verificando planillas CEV...")
            total = len(archivos)

            for i, arch in enumerate(archivos):
                progreso.progress((i + 1) / total, text=f"Auditando {arch.name}...")
                r_rel = str(arch.relative_to(carpeta_raiz))
                grupo_nombre = str(arch.parent.relative_to(carpeta_raiz)) if arch.parent != carpeta_raiz else "RAÍZ"

                datos, err = extraer_datos_completos(arch)
                if err:
                    errores_por_archivo[r_rel].append(err)
                    tipologia_errores_por_grupo[grupo_nombre]["Error de lectura Excel"] += 1
                    continue

                datos_por_archivo[r_rel] = datos
                for k, _ in CAMPOS_COMUNES_PROYECTO:
                    val = datos.get(k, "").strip()
                    if val and val.upper() not in ERRORES_EXCEL:
                        valores_proyecto[k][val] += 1

                ident = datos.get("identificacion", "").strip()
                num_c, torre_c = descomponer_unidad(ident, arch)

                if not ident:
                    errores_por_archivo[r_rel].append("Identificación en celda VACÍA")
                    tipologia_errores_por_grupo[grupo_nombre]["Identificación VACÍA"] += 1
                elif ident.upper() in ERRORES_EXCEL:
                    errores_por_archivo[r_rel].append(f"Fórmula rota ({ident}) en Identificación")
                    tipologia_errores_por_grupo[grupo_nombre]["Fórmula rota Identificación"] += 1
                else:
                    ident_vistas[normalizar(ident)].append(r_rel)
                    num_a, torre_a = descomponer_unidad(arch.stem, arch)
                    if num_a and num_c and num_a != num_c:
                        errores_por_archivo[r_rel].append(f"DESFASE: Nombre archivo indica '{num_a}' pero celda dice '{num_c}'")
                        tipologia_errores_por_grupo[grupo_nombre]["Desfase número Unidad"] += 1
                    if torre_a != "GENERAL" and torre_c != "GENERAL" and torre_a != torre_c:
                        errores_por_archivo[r_rel].append(f"DESFASE TORRE: Archivo indica '{torre_a}' pero celda dice '{torre_c}'")
                        tipologia_errores_por_grupo[grupo_nombre]["Desfase Torre"] += 1

                dorm = datos.get("dormitorios", "").strip()
                if not dorm or dorm in ("0", ""):
                    errores_por_archivo[r_rel].append(f"Dormitorios no definidos o en cero ('{dorm}')")
                    tipologia_errores_por_grupo[grupo_nombre]["Dormitorios no definidos"] += 1

                rol = datos.get("rol_vivienda", "").strip()
                if not rol:
                    errores_por_archivo[r_rel].append("Campo 'Rol vivienda' en BLANCO")
                    tipologia_errores_por_grupo[grupo_nombre]["Rol vivienda en BLANCO"] += 1
                elif rol.upper() in ERRORES_EXCEL:
                    errores_por_archivo[r_rel].append(f"Fórmula rota ({rol}) en Rol vivienda")
                    tipologia_errores_por_grupo[grupo_nombre]["Fórmula rota Rol vivienda"] += 1
                else:
                    roles_vistos[rol].append((r_rel, grupo_nombre))
                    if roles_sii and num_c:
                        rol_esp = roles_sii.get((torre_c, num_c)) or roles_sii.get(("GENERAL", num_c))
                        if rol_esp and rol_esp != rol:
                            errores_por_archivo[r_rel].append(f"ERROR ROL SII: Celda tiene '{rol}', pero según Asignación SII corresponde '{rol_esp}'")
                            tipologia_errores_por_grupo[grupo_nombre]["Rol no coincide con SII"] += 1

            # Discrepancias en datos maestros
            valores_maestros = {k: v.most_common(1)[0][0] for k, v in valores_proyecto.items() if v}
            for r_rel, d in datos_por_archivo.items():
                g_nom = str((carpeta_raiz / r_rel).parent.relative_to(carpeta_raiz)) if (carpeta_raiz / r_rel).parent != carpeta_raiz else "RAÍZ"
                for clave, etiqueta in CAMPOS_COMUNES_PROYECTO:
                    v_esp = valores_maestros.get(clave)
                    v_act = d.get(clave, "").strip()
                    if v_esp and not v_act:
                        errores_por_archivo[r_rel].append(f"Campo '{etiqueta}' en BLANCO")
                        tipologia_errores_por_grupo[g_nom][f"Falta {etiqueta}"] += 1
                    elif v_esp and v_act and normalizar(v_act) != normalizar(v_esp):
                        errores_por_archivo[r_rel].append(f"Discrepancia en '{etiqueta}': figura '{v_act}' (proyecto registra '{v_esp}')")
                        tipologia_errores_por_grupo[g_nom][f"Discrepancia en {etiqueta}"] += 1

            # Duplicados cruzados
            for id_dup, rutas in ident_vistas.items():
                if len(rutas) > 1:
                    for r in rutas:
                        g_nom = str((carpeta_raiz / r).parent.relative_to(carpeta_raiz)) if (carpeta_raiz / r).parent != carpeta_raiz else "RAÍZ"
                        errores_por_archivo[r].append(f"IDENTIFICACIÓN DUPLICADA: '{id_dup.upper()}' en {len(rutas)} planillas")
                        tipologia_errores_por_grupo[g_nom]["Identificación repetida"] += 1

            for rol_dup, tuplas in roles_vistos.items():
                if len(tuplas) > 1:
                    for r, g in tuplas:
                        errores_por_archivo[r].append(f"ROL REPETIDO: Rol '{rol_dup}' en {len(tuplas)} planillas")
                        tipologia_errores_por_grupo[g]["Rol repetido"] += 1

            progreso.empty()

            # --- RESULTADOS VISUALES ---
            total_obs = len(errores_por_archivo)
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Planillas Analizadas", total)
            c2.metric("Planillas Aprobadas para Subir", total - total_obs)
            c3.metric("Planillas con Observaciones", total_obs)

            # Generar Reporte Excel
            wb_rep = openpyxl.Workbook()
            ws_res = wb_rep.active
            ws_res.title = "RESUMEN"
            ws_res.append(["Edificio / Subsector", "Tipo de Observación", "Cantidad"])
            
            fill_head = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            font_head = Font(color="FFFFFF", bold=True)
            for c in range(1, 4):
                ws_res.cell(1, c).fill = fill_head
                ws_res.cell(1, c).font = font_head
                ws_res.cell(1, c).alignment = Alignment(horizontal="center")

            r_idx = 2
            for grp, errs in sorted(tipologia_errores_por_grupo.items()):
                for t_err, cant in sorted(errs.items(), key=lambda x: x[1], reverse=True):
                    ws_res.cell(row=r_idx, column=1, value=grp)
                    ws_res.cell(row=r_idx, column=2, value=t_err)
                    ws_res.cell(row=r_idx, column=3, value=cant)
                    r_idx += 1

            ws_det = wb_rep.create_sheet(title="OBSERVACIONES")
            ws_det.append(["Archivo", "Identificación", "Rol", "Detalle de Errores"])
            for c in range(1, 5):
                ws_det.cell(1, c).fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
                ws_det.cell(1, c).font = font_head
                ws_det.cell(1, c).alignment = Alignment(horizontal="center")

            filas_tabla = []
            for a in archivos:
                r_rel = str(a.relative_to(carpeta_raiz))
                fallas = errores_por_archivo.get(r_rel, [])
                d = datos_por_archivo.get(r_rel, {})
                obs_txt = " | ".join(fallas) if fallas else "Sin errores detectados"
                if fallas:
                    ws_det.append([a.name, d.get("identificacion", ""), d.get("rol_vivienda", ""), obs_txt])
                filas_tabla.append({
                    "Archivo": a.name,
                    "Identificación": d.get("identificacion", "S/I"),
                    "Rol": d.get("rol_vivienda", "S/R"),
                    "Estado": "❌ Observado" if fallas else "✅ Listo para CEV",
                    "Observaciones": obs_txt
                })

            buffer = io.BytesIO()
            wb_rep.save(buffer)
            buffer.seek(0)

            st.download_button(
                label="📥 Descargar Reporte de Auditoría (.xlsx)",
                data=buffer,
                file_name=f"Auditoria_CEV_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary"
            )

            st.subheader("📋 Estado Detallado por Planilla")
            st.dataframe(pd.DataFrame(filas_tabla), use_container_width=True)

            shutil.rmtree(temp_dir, ignore_errors=True)