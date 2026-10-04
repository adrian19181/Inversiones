import datetime
import io
import os
import socket
import subprocess
import sys

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

# ---------------------------------------------------------
# AUTO-LANZADOR AUTOMÁTICO EN WINDOWS (DOBLE CLICK)
# ---------------------------------------------------------
if (
    __name__ == "__main__"
    and not os.environ.get("STREAMLIT_RUNNING")
    and not st.runtime.exists()
):
    os.environ["STREAMLIT_RUNNING"] = "true"
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))
        IP = s.getsockname()[0]
    except Exception:
        IP = "127.0.0.1"
    finally:
        s.close()

    os.system("cls" if os.name == "nt" else "clear")
    print("=" * 60)
    print("  DASHBOARD: TODAS MIS INVERSIONES")
    print("=" * 60)
    print("\n  PARA VER EN TU CELULAR O NAVEGADOR, ABRE ESTA DIRECCIÓN:\n")
    print(f"  👉  http://{IP}:8501  👈\n")
    print("=" * 60)

    script_path = os.path.abspath(__file__)
    subprocess.run([
        sys.executable,
        "-m",
        "streamlit",
        "run",
        script_path,
        "--server.address=0.0.0.0",
    ])
    sys.exit()

# ---------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Todas mis Inversiones",
    page_icon="📈",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------
# SCRIPT INVISIBLE: CIERRE AUTOMÁTICO EN MÓVIL (POPOVER)
# ---------------------------------------------------------
components.html(
    """
    <script>
    const doc = window.parent.document;
    doc.addEventListener('change', function(e) {
        if (e.target.closest('div[data-testid="stPopoverBody"]') || e.target.closest('div[data-baseweb="popover"]')) {
            setTimeout(function() {
                doc.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true, cancelable: true }));
            }, 120);
        }
    });
    </script>
    """,
    height=0,
    width=0,
)

# ---------------------------------------------------------
# ESTILOS CSS GLOBALES
# ---------------------------------------------------------
st.markdown(
    """
    <style>
        /* FONDO GENERAL Y CONTENEDOR MÓVIL */
        html, body, .stApp, [data-testid="stAppViewContainer"] {
            background-color: #0E1117 !important;
            color: #FAFAFA !important;
        }
        [data-testid="stHeader"] { background-color: rgba(0, 0, 0, 0) !important; }
        .block-container { 
            padding: 1.0rem 0.4rem 2rem 0.4rem !important; 
            max-width: 740px !important; 
        }

        /* TÍTULOS Y TEXTOS */
        h1, h2, h3, h4, h5, h6, [data-testid="stMarkdownContainer"] h3 {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
        }

        /* SELECTBOX / DROPDOWN */
        div[data-testid="stSelectbox"] label p, div[data-testid="stRadio"] label p {
            color: #00E676 !important;
            font-size: 0.95rem !important;
            font-weight: 800 !important;
        }

        /* BOTÓN REFRESCAR */
        div[data-testid="stButton"] > button {
            background-color: #1E222B !important;
            border: 1.5px solid #107C41 !important;
            border-radius: 8px !important;
            padding: 0.4rem 0.6rem !important;
            width: 100% !important; 
        }
        div[data-testid="stButton"] > button * {
            color: #FFFFFF !important;
            -webkit-text-fill-color: #FFFFFF !important;
            font-weight: 700 !important;
        }
        div[data-testid="stButton"] > button:hover {
            background-color: #107C41 !important;
            border-color: #107C41 !important;
        }
        /* AJUSTES MÓVIL */
        @media (max-width: 480px) {
            .block-container { padding: 0.5rem 0.2rem 1.5rem 0.2rem !important; }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# CONEXIÓN A GOOGLE DRIVE Y EXTRACCIÓN DINÁMICA
# ---------------------------------------------------------
DRIVE_FILE_ID = "1QA04CUCoI2hTaLjTXjCT37gqZMBE8EzW"

@st.cache_data(ttl=300)
def load_and_clean_data():
    session = requests.Session()
    urls_to_try = [
        f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx",
        f"https://drive.google.com/uc?export=download&id={DRIVE_FILE_ID}",
    ]

    excel_bytes = None
    for url in urls_to_try:
        try:
            res = session.get(url, timeout=30)
            for key, val in res.cookies.items():
                if key.startswith("download_warning"):
                    res = session.get(
                        f"{url}&confirm={val}" if "export?format=xlsx" not in url else url,
                        timeout=30,
                    )
                    break
            if res.content.startswith(b"PK\x03\x04"):
                excel_bytes = res.content
                break
        except Exception:
            continue

    if not excel_bytes:
        raise ValueError("No se pudo descargar el archivo de Excel. Revisa los permisos de Drive.")

    df_raw = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="INVERSIONES", engine="openpyxl")
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    # Filtrar solo transacciones válidas en la columna L (Tipo Renta Fija)
    df_clean = df_raw.dropna(subset=["Tipo Renta Fija"]).copy()
    
    df_clean["Tipo Renta Fija"] = df_clean["Tipo Renta Fija"].astype(str).str.strip()
    df_clean["Entidad"] = df_clean["Entidad"].astype(str).str.strip()

    num_cols = [
        "CAPITAL",
        "TASA",
        "RETENCION IMP, 2%",
        "VALOR GANADO CERTIFICADO",
        "DIAS",
        "INTERESE NETO",
        "NETO A RECIBIR",
    ]
    for col in num_cols:
        if col in df_clean.columns:
            df_clean[col] = pd.to_numeric(df_clean[col], errors="coerce").fillna(0.0)

    df_clean["FECHA INICIO"] = pd.to_datetime(df_clean["FECHA INICIO"], errors="coerce")
    df_clean["FECHA FIN"] = pd.to_datetime(df_clean["FECHA FIN"], errors="coerce")

    return df_clean

try:
    with st.spinner("Procesando datos en tiempo real..."):
        df_datos = load_and_clean_data()
except Exception as e:
    st.error(f"Error al conectar con Google Drive: {e}")
    st.stop()

# ---------------------------------------------------------
# HELPER: FORMATO DE FECHA DD/MM/YYYY
# ---------------------------------------------------------
def formato_fecha_estandar(fecha):
    if pd.isna(fecha): return "-"
    return fecha.strftime('%d/%m/%Y')

# ---------------------------------------------------------
# CABECERA & BOTÓN REFRESCAR
# ---------------------------------------------------------
header_col1, header_col2 = st.columns([2.5, 1.5], vertical_alignment="center")
with header_col1:
    st.markdown(
        "<h3 style='margin: 0; color: #FFFFFF !important;'>📈 Todas mis Inversiones</h3>",
        unsafe_allow_html=True,
    )
with header_col2:
    if st.button("🔄 Refrescar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.markdown("---")

# ---------------------------------------------------------
# FILTRO GENERAL PARA TABLA DE RESULTADOS CONSOLIDADOS
# ---------------------------------------------------------
opciones_filtro = [
    "Todas las Inversiones",
    "Plazos Fijos",
    "Ahorros Programados",
    "Austro Futuro",
    "Ahorros Vista Jardín Azuayo",
]

categoria_seleccionada = st.selectbox(
    "📊 Selecciona el Producto de Inversión (Consolidado):",
    opciones_filtro,
    index=0,
)

# ---------------------------------------------------------
# CÁLCULOS DINÁMICOS TABLA 1
# ---------------------------------------------------------
if categoria_seleccionada == "Todas las Inversiones":
    df_filtro = df_datos.copy()
elif categoria_seleccionada == "Plazos Fijos":
    df_filtro = df_datos[df_datos["Tipo Renta Fija"] == "Plazo Fijo"]
elif categoria_seleccionada == "Ahorros Programados":
    df_filtro = df_datos[df_datos["Tipo Renta Fija"] == "Ahorro Programado"]
elif categoria_seleccionada == "Austro Futuro":
    df_filtro = df_datos[df_datos["Tipo Renta Fija"] == "AustroFuturo"]
elif categoria_seleccionada == "Ahorros Vista Jardín Azuayo":
    df_filtro = df_datos[(df_datos["Tipo Renta Fija"] == "Ahorros Vista") & (df_datos["Entidad"] == "Jardin Azuayo")]
else:
    df_filtro = df_datos.copy()

min_fecha_inicio = df_filtro["FECHA INICIO"].min()
max_fecha_fin = df_filtro["FECHA FIN"].max()

if categoria_seleccionada in ["Plazos Fijos", "Ahorros Programados"]:
    dias_unicos = set()
    for _, row in df_filtro.iterrows():
        inicio = row["FECHA INICIO"]
        fin = row["FECHA FIN"]
        if pd.notna(inicio) and pd.notna(fin):
            actual = inicio
            while actual < fin:
                dias_unicos.add(actual)
                actual += datetime.timedelta(days=1)
    dias = len(dias_unicos)
else:
    if pd.notna(min_fecha_inicio) and pd.notna(max_fecha_fin):
        dias = (max_fecha_fin - min_fecha_inicio).days
    else:
        dias = 0

meses = dias / 30.0
anios = dias / 365.0

tot_neto = df_filtro["INTERESE NETO"].sum() if "INTERESE NETO" in df_filtro.columns else 0.0
tot_retencion = df_filtro["RETENCION IMP, 2%"].sum() if "RETENCION IMP, 2%" in df_filtro.columns else 0.0
tot_certificado = df_filtro["VALOR GANADO CERTIFICADO"].sum() if "VALOR GANADO CERTIFICADO" in df_filtro.columns else 0.0

num_cap_pond = (df_filtro["CAPITAL"] * df_filtro["DIAS"]).sum()
den_cap_pond = df_filtro["DIAS"].sum()
cap_invertido_ponderado = (num_cap_pond / den_cap_pond) if den_cap_pond > 0 else 0.0

num_pond = (df_filtro["CAPITAL"] * df_filtro["DIAS"] * df_filtro["TASA"]).sum()
den_pond = (df_filtro["CAPITAL"] * df_filtro["DIAS"]).sum()
tasa_pond = (num_pond / den_pond) if den_pond > 0 else 0.0

fecha_ini_str = formato_fecha_estandar(min_fecha_inicio)
fecha_fin_str = formato_fecha_estandar(max_fecha_fin)
tasa_str = f"{tasa_pond * 100:,.2f}%" if tasa_pond <= 1.0 else f"{tasa_pond:,.2f}%"

df_kpis = pd.DataFrame([
    {"Métrica": "Fecha Inicio", "Resultado": fecha_ini_str},
    {"Métrica": "Fecha Fin", "Resultado": fecha_fin_str},
    {"Métrica": "# Días", "Resultado": f"{int(dias):,}"},
    {"Métrica": "# Meses", "Resultado": f"{meses:,.2f}"},
    {"Métrica": "# Años", "Resultado": f"{anios:,.2f}"},
    {"Métrica": "Capital Invertido Ponderado", "Resultado": f"${cap_invertido_ponderado:,.2f}"},
    {"Métrica": "Intereses Netos Totales", "Resultado": f"${tot_neto:,.2f}"},
    {"Métrica": "Retención Impuestos", "Resultado": f"${tot_retencion:,.2f}"},
    {"Métrica": "Ganado Certificado", "Resultado": f"${tot_certificado:,.2f}"},
    {"Métrica": "Tasa Ponderada", "Resultado": tasa_str},
])

# ---------------------------------------------------------
# RENDERIZADO TABLA 1 (RESUMEN KPIS - 1 COLUMNA FIJA)
# ---------------------------------------------------------
def render_kpi_table_html(df):
    css = """<style>
.kpi-tbl-wrapper {
    max-height: 520px; width: 100%; overflow-x: auto; overflow-y: auto;
    border: 1px solid #475569; border-radius: 8px; background-color: #0F172A;
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.5); margin-top: 8px; margin-bottom: 16px;
    padding: 0px !important; display: block; position: relative;
}
.kpi-tbl-sticky {
    width: 100%; border-collapse: collapse !important; border-spacing: 0 !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    font-size: 0.88rem; margin: 0; table-layout: auto !important;
}
.kpi-tbl-sticky th {
    position: sticky !important; top: 0 !important; z-index: 20 !important;
    background-color: #1E293B !important; color: #FFFFFF !important;
    border: 1px solid #475569 !important; padding: 10px 8px !important;
    text-align: center !important; font-weight: 800 !important; white-space: nowrap !important;
}
.kpi-tbl-sticky th:first-child {
    position: sticky !important; top: 0 !important; left: 0 !important; z-index: 50 !important;
    width: 125px !important; min-width: 125px !important; max-width: 125px !important;
    text-align: center !important; box-shadow: 2px 0 5px rgba(0,0,0,0.4) !important;
    white-space: normal !important; word-wrap: break-word !important;
}
.kpi-tbl-sticky td {
    position: static !important; z-index: auto !important; padding: 9px 10px !important;
    border: 1px solid #334155 !important; vertical-align: middle !important;
    text-align: center !important; color: #000000 !important; white-space: nowrap !important;
}
.kpi-tbl-sticky .lbl-sticky-col {
    position: sticky !important; left: 0 !important; z-index: 30 !important;
    font-weight: 800 !important; text-align: center !important;
    width: 125px !important; min-width: 125px !important; max-width: 125px !important;
    background-color: #FACC15 !important; color: #000000 !important;
    border: 1px solid #334155 !important; box-shadow: 3px 0 6px rgba(0,0,0,0.4) !important;
    background-clip: padding-box !important; white-space: normal !important; word-wrap: break-word !important;
}
.row-excel-green td { background-color: #ECFDF5 !important; color: #000000 !important; font-weight: 700 !important; }
.row-excel-green .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-excel-yellow td { background-color: #FEF3C7 !important; color: #000000 !important; font-weight: 700 !important; }
.row-excel-yellow .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-excel-white td { background-color: #F8FAFC !important; color: #000000 !important; font-weight: 700 !important; }
.row-excel-white .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-excel-highlight td { background-color: #D1FAE5 !important; color: #000000 !important; font-weight: 800 !important; }
.row-excel-highlight .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }
</style>"""

    html = f'{css}<div class="kpi-tbl-wrapper"><table class="kpi-tbl-sticky">'
    html += '<thead><tr>'
    html += '<th style="width: 125px !important;">Métrica</th>'
    html += '<th>Resultado</th>'
    html += '</tr></thead><tbody>'

    row_classes = ["row-excel-green", "row-excel-yellow", "row-excel-white"]

    for idx, row in df.iterrows():
        metrica = str(row["Métrica"])
        res = str(row["Resultado"])
        
        is_highlight = "Intereses Netos" in metrica or "Capital Invertido Ponderado" in metrica
        if is_highlight:
            row_cls = "row-excel-highlight"
        else:
            row_cls = row_classes[idx % len(row_classes)]

        html += f'<tr class="{row_cls}">'
        html += f'<td class="lbl-sticky-col">{metrica}</td>'
        html += f'<td style="font-size: 1.05rem; font-weight: 800;">{res}</td>'
        html += '</tr>'

    html += '</tbody></table></div>'
    return html.replace("\n", " ")

st.markdown(
    f"<h4 style='color: #38BDF8 !important; margin-bottom: 4px;'>📌 Resultados Consolidados: {categoria_seleccionada}</h4>",
    unsafe_allow_html=True,
)
st.markdown(render_kpi_table_html(df_kpis), unsafe_allow_html=True)


# ---------------------------------------------------------
# SECCIÓN 2: TABLA DE DETALLES (2 COLUMNAS FIJAS)
# ---------------------------------------------------------
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    "<h4 style='color: #38BDF8 !important; margin-bottom: 4px;'>📝 Resumen Todas Las Inversiones</h4>",
    unsafe_allow_html=True,
)

# Filtro Horizontal de Leyendas
tipos_disponibles = ["Todos"] + sorted(list(df_datos["Tipo Renta Fija"].dropna().unique()))
filtro_detalles = st.radio(
    "Filtrar detalles por:",
    tipos_disponibles,
    horizontal=True,
    index=0
)

if filtro_detalles != "Todos":
    df_detalles = df_datos[df_datos["Tipo Renta Fija"] == filtro_detalles].copy()
else:
    df_detalles = df_datos.copy()

def render_detalles_inversiones_html(df):
    css = """<style>
.det-tbl-wrapper {
    max-height: 520px; width: 100%; overflow-x: auto; overflow-y: auto;
    border: 1px solid #475569; border-radius: 8px; background-color: #0F172A;
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.5); margin-top: 8px; margin-bottom: 24px;
    padding: 0px !important; display: block; position: relative;
}
.det-tbl-sticky {
    width: 100%; border-collapse: collapse !important; border-spacing: 0 !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    font-size: 0.85rem; margin: 0; table-layout: auto !important;
}
/* Encabezados Fijos (Fila 1) */
.det-tbl-sticky th {
    position: sticky !important; top: 0 !important; z-index: 20 !important;
    background-color: #1E293B !important; color: #FFFFFF !important;
    border: 1px solid #475569 !important; padding: 6px 8px !important;
    text-align: center !important; font-weight: 800 !important; white-space: nowrap !important;
}
/* Encabezado Columna L (Fija 1) - Ancho más ajustado (90px) */
.det-tbl-sticky th:nth-child(1) {
    position: sticky !important; top: 0 !important; left: 0 !important; z-index: 50 !important;
    width: 90px !important; min-width: 90px !important; max-width: 90px !important;
    text-align: center !important; white-space: normal !important; word-wrap: break-word !important;
}
/* Encabezado Columna K (Fija 2 - Con Sombra) */
.det-tbl-sticky th:nth-child(2) {
    position: sticky !important; top: 0 !important; left: 90px !important; z-index: 50 !important;
    width: 90px !important; min-width: 90px !important; max-width: 90px !important;
    text-align: center !important; box-shadow: 2px 0 5px rgba(0,0,0,0.4) !important;
    white-space: normal !important; word-wrap: break-word !important;
}

/* Celdas Móviles Normales */
.det-tbl-sticky td {
    position: static !important; z-index: auto !important; padding: 6px 8px !important;
    border: 1px solid #334155 !important; vertical-align: middle !important;
    text-align: center !important; color: #000000 !important; white-space: nowrap !important;
}

/* Celda Columna L (Fija 1) */
.det-tbl-sticky .sticky-col-1 {
    position: sticky !important; left: 0 !important; z-index: 30 !important; font-weight: 800 !important;
    width: 90px !important; min-width: 90px !important; max-width: 90px !important;
    background-color: #FACC15 !important; color: #000000 !important; border: 1px solid #334155 !important;
    background-clip: padding-box !important; white-space: normal !important; word-wrap: break-word !important;
}

/* Celda Columna K (Fija 2 - Con Sombra) */
.det-tbl-sticky .sticky-col-2 {
    position: sticky !important; left: 90px !important; z-index: 30 !important; font-weight: 800 !important;
    width: 90px !important; min-width: 90px !important; max-width: 90px !important;
    background-color: #FACC15 !important; color: #000000 !important; border: 1px solid #334155 !important;
    box-shadow: 3px 0 6px rgba(0,0,0,0.4) !important; background-clip: padding-box !important;
    white-space: normal !important; word-wrap: break-word !important;
}

/* Clases alternadas (Mismo Estilo) */
.row-det-green td { background-color: #ECFDF5 !important; font-weight: 700 !important; }
.row-det-green .sticky-col-1, .row-det-green .sticky-col-2 { background-color: #FACC15 !important; }

.row-det-yellow td { background-color: #FEF3C7 !important; font-weight: 700 !important; }
.row-det-yellow .sticky-col-1, .row-det-yellow .sticky-col-2 { background-color: #FACC15 !important; }

.row-det-white td { background-color: #F8FAFC !important; font-weight: 700 !important; }
.row-det-white .sticky-col-1, .row-det-white .sticky-col-2 { background-color: #FACC15 !important; }
</style>"""

    html = f'{css}<div class="det-tbl-wrapper"><table class="det-tbl-sticky">'
    html += '<thead><tr>'
    
    # Titulos Fijos en DOS LÍNEAS (Col L y K)
    html += '<th style="width: 90px !important;">Tipo<br>Renta</th>'
    html += '<th style="width: 90px !important;">Entidad<br>Financiera</th>'
    
    # Titulos Moviles en DOS LÍNEAS (Col A, B, G, H, I, M, F)
    html += '<th>Capital<br>Invertido</th>'
    html += '<th>Tasa<br>Nominal</th>'
    html += '<th>Plazo<br>(Días)</th>'
    html += '<th>Fecha<br>Inicio</th>'
    html += '<th>Fecha<br>Fin</th>'
    html += '<th>Interés<br>Neto</th>'
    html += '<th>Neto a<br>Recibir</th>'
    html += '</tr></thead><tbody>'

    row_classes = ["row-det-green", "row-det-yellow", "row-det-white"]

    for idx, row in df.iterrows():
        # Extracción
        tipo = str(row.get("Tipo Renta Fija", "-"))
        entidad = str(row.get("Entidad", "-"))
        cap = row.get("CAPITAL", 0.0)
        tasa = row.get("TASA", 0.0)
        dias = row.get("DIAS", 0.0)
        f_ini = row.get("FECHA INICIO", pd.NaT)
        f_fin = row.get("FECHA FIN", pd.NaT)
        i_neto = row.get("INTERESE NETO", 0.0)
        n_recibir = row.get("NETO A RECIBIR", 0.0)
        
        # Formateo multilínea de Col 1 y 2
        tipo_html = tipo.replace(" ", "<br>")
        if tipo_html == "AustroFuturo": tipo_html = "Austro<br>Futuro"
        entidad_html = entidad.replace(" ", "<br>")
        
        # Formateo Movil
        cap_str = f"${float(cap):,.2f}"
        tasa_f = float(tasa)
        tasa_str = f"{tasa_f*100:,.2f}%" if tasa_f <= 1.0 else f"{tasa_f:,.2f}%"
        dias_str = f"{int(float(dias)):,}"
        f_ini_str = formato_fecha_estandar(f_ini)
        f_fin_str = formato_fecha_estandar(f_fin)
        i_neto_str = f"${float(i_neto):,.2f}"
        n_recibir_str = f"${float(n_recibir):,.2f}"
        
        r_cls = row_classes[idx % len(row_classes)]

        html += f'<tr class="{r_cls}">'
        # Celdas Fijas con texto distribuido en 2 lineas
        html += f'<td class="sticky-col-1">{tipo_html}</td>'
        html += f'<td class="sticky-col-2">{entidad_html}</td>'
        
        # Celdas Moviles (A, B, G, H, I, M, F)
        html += f'<td>{cap_str}</td>'
        html += f'<td>{tasa_str}</td>'
        html += f'<td>{dias_str}</td>'
        html += f'<td>{f_ini_str}</td>'
        html += f'<td>{f_fin_str}</td>'
        html += f'<td>{i_neto_str}</td>'
        html += f'<td>{n_recibir_str}</td>'
        html += '</tr>'

    html += '</tbody></table></div>'
    return html.replace("\n", " ")

# Imprimir la segunda tabla 
st.markdown(render_detalles_inversiones_html(df_detalles), unsafe_allow_html=True)