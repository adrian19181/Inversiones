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
import plotly.express as px

# ---------------------------------------------------------
# AUTO-LANZADOR PARA EJECUCIÓN CON DOBLE CLIC EN LOCAL
# ---------------------------------------------------------
is_running_in_streamlit = (
    st.runtime.exists()
    or os.environ.get("STREAMLIT_RUNNING") == "true"
    or os.environ.get("STREAMLIT_SERVER_PORT") is not None
    or any("streamlit" in arg.lower() for arg in sys.argv)
)

if __name__ == "__main__" and not is_running_in_streamlit:
    os.environ["STREAMLIT_RUNNING"] = "true"
    
    # Obtener la IP local de la computadora para acceso desde celular
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
    print("\n  Iniciando servidor Streamlit...\n")
    print(f"  👉  Navegador local:  http://localhost:8501")
    print(f"  👉  Desde el celular: http://{IP}:8501\n")
    print("=" * 60)

    script_path = os.path.abspath(__file__)
    
    # Lanzar Streamlit
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

        /* SELECTBOX / DROPDOWN Y PESTAÑAS */
        div[data-testid="stSelectbox"] label p, div[data-testid="stRadio"] label p {
            color: #00E676 !important;
            font-size: 0.95rem !important;
            font-weight: 800 !important;
        }

        /* ESTILO PESTAÑAS (ST.TABS) */
        button[data-baseweb="tab"] {
            background-color: #1E293B !important;
            color: #94A3B8 !important;
            border-radius: 6px 6px 0 0 !important;
            padding: 8px 12px !important;
            font-weight: 700 !important;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            background-color: #0F172A !important;
            color: #38BDF8 !important;
            border-bottom: 2px solid #38BDF8 !important;
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
# RENDERIZADO TABLA 1 (RESUMEN KPIS - FILAS ANGOSTAS Y COMPACTAS)
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
    font-size: 0.85rem; margin: 0; table-layout: auto !important;
}
.kpi-tbl-sticky th {
    position: sticky !important; top: 0 !important; z-index: 20 !important;
    background-color: #1E293B !important; color: #FFFFFF !important;
    border: 1px solid #475569 !important; padding: 5px 6px !important;
    text-align: center !important; font-weight: 800 !important; white-space: nowrap !important;
}
.kpi-tbl-sticky th:first-child {
    position: sticky !important; top: 0 !important; left: 0 !important; z-index: 50 !important;
    width: 115px !important; min-width: 115px !important; max-width: 115px !important;
    text-align: center !important; box-shadow: 2px 0 5px rgba(0,0,0,0.4) !important;
    white-space: normal !important; word-wrap: break-word !important;
}
.kpi-tbl-sticky td {
    position: static !important; z-index: auto !important; padding: 4px 6px !important;
    border: 1px solid #334155 !important; vertical-align: middle !important;
    text-align: center !important; color: #000000 !important; white-space: nowrap !important;
    line-height: 1.15 !important;
}
.kpi-tbl-sticky .lbl-sticky-col {
    position: sticky !important; left: 0 !important; z-index: 30 !important;
    font-weight: 800 !important; text-align: center !important;
    width: 115px !important; min-width: 115px !important; max-width: 115px !important;
    background-color: #FACC15 !important; color: #000000 !important;
    border: 1px solid #334155 !important; box-shadow: 3px 0 6px rgba(0,0,0,0.4) !important;
    background-clip: padding-box !important; white-space: normal !important; word-wrap: break-word !important;
    padding: 4px 6px !important; line-height: 1.15 !important;
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
    html += '<th style="width: 115px !important;">Métrica</th>'
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
        html += f'<td style="font-size: 0.95rem; font-weight: 800;">{res}</td>'
        html += '</tr>'

    html += '</tbody></table></div>'
    return html.replace("\n", " ")

st.markdown(
    f"<h4 style='color: #38BDF8 !important; margin-bottom: 4px;'>📌 Resultados Consolidados: {categoria_seleccionada}</h4>",
    unsafe_allow_html=True,
)
st.markdown(render_kpi_table_html(df_kpis), unsafe_allow_html=True)


# ---------------------------------------------------------
# SECCIÓN SEGUNDA: TABLAS INFORMATIVAS (PESTAÑAS DINÁMICAS)
# ---------------------------------------------------------
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    "<h4 style='color: #38BDF8 !important; margin-bottom: 4px;'>📊 Tablas Informativas</h4>",
    unsafe_allow_html=True,
)

# HELPER: FORMATO DE TÍTULOS EN DOS LÍNEAS
def dos_lineas_titulo(titulo):
    reemplazos = {
        "Año (FECHA FIN)": "Año<br>(FECHA FIN)",
        "Ahorro Programado": "Ahorro<br>Programado",
        "Ahorros Vista": "Ahorros<br>Vista",
        "AustroFuturo": "Austro<br>Futuro",
        "Plazo Fijo": "Plazo<br>Fijo",
        "Total general": "Total<br>general",
        "Etiquetas de fila": "Etiquetas<br>de fila",
        "Suma de INTERESE NETO": "Suma de<br>INTERESE NETO",
        "Cuenta de Tipo Renta Fija": "Cuenta de<br>Tipo Renta Fija"
    }
    if titulo in reemplazos:
        return reemplazos[titulo]
    if " " in titulo and "<br>" not in titulo:
        partes = titulo.split(" ", 1)
        return f"{partes[0]}<br>{partes[1]}"
    return titulo

# TABLA DINÁMICA 1: AÑO vs TIPO RENTA FIJA
def render_pivot_interes_neto_html(df_data):
    df_temp = df_data.copy()
    df_temp["Año FIN"] = df_temp["FECHA FIN"].dt.year
    df_valid = df_temp.dropna(subset=["Año FIN", "Tipo Renta Fija"]).copy()
    df_valid["Año FIN"] = df_valid["Año FIN"].astype(int).astype(str)

    pivot = pd.pivot_table(
        df_valid,
        index="Año FIN",
        columns="Tipo Renta Fija",
        values="INTERESE NETO",
        aggfunc="sum",
        fill_value=0.0,
        margins=True,
        margins_name="Total general"
    )

    cols = [c for c in pivot.columns if c != "Total general"]
    if "Total general" in pivot.columns:
        cols.append("Total general")

    pivot = pivot[cols]

    css = """<style>
.pvt-tbl-wrapper {
    max-height: 520px; width: 100%; overflow-x: auto; overflow-y: auto;
    border: 1px solid #475569; border-radius: 8px; background-color: #0F172A;
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.5); margin-top: 8px; margin-bottom: 24px;
    padding: 0px !important; display: block; position: relative;
}
.pvt-tbl-sticky {
    width: 100%; border-collapse: collapse !important; border-spacing: 0 !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    font-size: 0.85rem; margin: 0; table-layout: auto !important;
}
/* Encabezados Fijos (Fila 1) - Centrado y Ancho Inteligente */
.pvt-tbl-sticky th {
    position: sticky !important; top: 0 !important; z-index: 20 !important;
    background-color: #1E293B !important; color: #FFFFFF !important;
    border: 1px solid #475569 !important; padding: 6px 10px !important;
    text-align: center !important; font-weight: 800 !important; white-space: nowrap !important;
    line-height: 1.2 !important;
}
/* Esquina Superior Izquierda (Fija 2D) */
.pvt-tbl-sticky th:first-child {
    position: sticky !important; top: 0 !important; left: 0 !important; z-index: 50 !important;
    text-align: center !important; box-shadow: 2px 0 5px rgba(0,0,0,0.4) !important;
    white-space: nowrap !important; background-color: #1E293B !important;
}

/* Celdas de datos normales - Centrados y Ancho Inteligente */
.pvt-tbl-sticky td {
    position: static !important; z-index: auto !important; padding: 6px 10px !important;
    border: 1px solid #334155 !important; vertical-align: middle !important;
    text-align: center !important; color: #000000 !important; white-space: nowrap !important;
}

/* Columna 1 Fija (Años) - Centrado */
.pvt-tbl-sticky .lbl-sticky-col {
    position: sticky !important; left: 0 !important; z-index: 30 !important; font-weight: 800 !important;
    text-align: center !important; border: 1px solid #334155 !important;
    box-shadow: 3px 0 6px rgba(0,0,0,0.4) !important; background-clip: padding-box !important;
    white-space: nowrap !important;
}

/* Filas normales alternadas */
.row-pvt-green td { background-color: #ECFDF5 !important; color: #000000 !important; font-weight: 600 !important; }
.row-pvt-green .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-pvt-yellow td { background-color: #FEF3C7 !important; color: #000000 !important; font-weight: 600 !important; }
.row-pvt-yellow .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-pvt-white td { background-color: #F8FAFC !important; color: #000000 !important; font-weight: 600 !important; }
.row-pvt-white .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

/* Fila de Total General */
.row-pvt-total td { background-color: #BAE6FD !important; color: #000000 !important; font-weight: 900 !important; }
.row-pvt-total .lbl-sticky-col { background-color: #38BDF8 !important; color: #000000 !important; font-weight: 900 !important; }
</style>"""

    html = f'{css}<div class="pvt-tbl-wrapper"><table class="pvt-tbl-sticky">'
    html += '<thead><tr>'
    html += f'<th>{dos_lineas_titulo("Año (FECHA FIN)")}</th>'
    for col in cols:
        html += f'<th>{dos_lineas_titulo(col)}</th>'
    html += '</tr></thead><tbody>'

    row_classes = ["row-pvt-green", "row-pvt-yellow", "row-pvt-white"]

    for idx, (year, row_data) in enumerate(pivot.iterrows()):
        is_total = (year == "Total general")
        row_cls = "row-pvt-total" if is_total else row_classes[idx % len(row_classes)]

        year_str = "Total<br>general" if is_total else str(year)

        html += f'<tr class="{row_cls}">'
        html += f'<td class="lbl-sticky-col">{year_str}</td>'

        for col in cols:
            val = row_data[col]
            val_str = "-" if (val == 0.0 or pd.isna(val)) else f"${val:,.2f}"
            
            if col == "Total general" or is_total:
                html += f'<td style="font-weight: 800;">{val_str}</td>'
            else:
                html += f'<td>{val_str}</td>'

        html += '</tr>'

    html += '</tbody></table></div>'
    return html.replace("\n", " ")


# TABLA DINÁMICA 2: ENTIDAD vs INTERES NETO
def render_pivot_entidad_html(df_data):
    df_temp = df_data.copy()
    df_valid = df_temp.dropna(subset=["Entidad"]).copy()

    pivot = pd.pivot_table(
        df_valid,
        index="Entidad",
        values="INTERESE NETO",
        aggfunc="sum",
        fill_value=0.0,
        margins=True,
        margins_name="Total general"
    ).reset_index()

    css = """<style>
.pvt-tbl-wrapper {
    max-height: 520px; width: 100%; overflow-x: auto; overflow-y: auto;
    border: 1px solid #475569; border-radius: 8px; background-color: #0F172A;
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.5); margin-top: 8px; margin-bottom: 24px;
    padding: 0px !important; display: block; position: relative;
}
.pvt-tbl-sticky {
    width: 100%; border-collapse: collapse !important; border-spacing: 0 !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    font-size: 0.85rem; margin: 0; table-layout: auto !important;
}
/* Encabezados Fijos (Fila 1) - Centrado y Ancho Inteligente */
.pvt-tbl-sticky th {
    position: sticky !important; top: 0 !important; z-index: 20 !important;
    background-color: #1E293B !important; color: #FFFFFF !important;
    border: 1px solid #475569 !important; padding: 6px 10px !important;
    text-align: center !important; font-weight: 800 !important; white-space: nowrap !important;
    line-height: 1.2 !important;
}
/* Esquina Superior Izquierda (Fija 2D) */
.pvt-tbl-sticky th:first-child {
    position: sticky !important; top: 0 !important; left: 0 !important; z-index: 50 !important;
    text-align: center !important; box-shadow: 2px 0 5px rgba(0,0,0,0.4) !important;
    white-space: nowrap !important; background-color: #1E293B !important;
}

/* Celdas de datos normales - Centrados */
.pvt-tbl-sticky td {
    position: static !important; z-index: auto !important; padding: 6px 10px !important;
    border: 1px solid #334155 !important; vertical-align: middle !important;
    text-align: center !important; color: #000000 !important; white-space: nowrap !important;
}

/* Columna 1 Fija (Entidades) - Centrado */
.pvt-tbl-sticky .lbl-sticky-col {
    position: sticky !important; left: 0 !important; z-index: 30 !important; font-weight: 800 !important;
    text-align: center !important; border: 1px solid #334155 !important;
    box-shadow: 3px 0 6px rgba(0,0,0,0.4) !important; background-clip: padding-box !important;
    white-space: nowrap !important;
}

.row-pvt-green td { background-color: #ECFDF5 !important; color: #000000 !important; font-weight: 600 !important; }
.row-pvt-green .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-pvt-yellow td { background-color: #FEF3C7 !important; color: #000000 !important; font-weight: 600 !important; }
.row-pvt-yellow .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-pvt-white td { background-color: #F8FAFC !important; color: #000000 !important; font-weight: 600 !important; }
.row-pvt-white .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-pvt-total td { background-color: #BAE6FD !important; color: #000000 !important; font-weight: 900 !important; }
.row-pvt-total .lbl-sticky-col { background-color: #38BDF8 !important; color: #000000 !important; font-weight: 900 !important; }
</style>"""

    html = f'{css}<div class="pvt-tbl-wrapper"><table class="pvt-tbl-sticky">'
    html += '<thead><tr>'
    html += f'<th>{dos_lineas_titulo("Etiquetas de fila")}</th>'
    html += f'<th>{dos_lineas_titulo("Suma de INTERESE NETO")}</th>'
    html += '</tr></thead><tbody>'

    row_classes = ["row-pvt-green", "row-pvt-yellow", "row-pvt-white"]

    for idx, row in pivot.iterrows():
        entidad = str(row["Entidad"])
        val = row["INTERESE NETO"]
        is_total = (entidad == "Total general")
        row_cls = "row-pvt-total" if is_total else row_classes[idx % len(row_classes)]

        entidad_str = "Total<br>general" if is_total else entidad.replace(" ", "<br>")
        val_str = "-" if (val == 0.0 or pd.isna(val)) else f"${val:,.2f}"

        html += f'<tr class="{row_cls}">'
        html += f'<td class="lbl-sticky-col">{entidad_str}</td>'
        html += f'<td style="font-weight: 800;">{val_str}</td>'
        html += '</tr>'

    html += '</tbody></table></div>'
    return html.replace("\n", " ")


# TABLA DINÁMICA 3: TIPO RENTA FIJA vs SUMA INTERÉS NETO & CUENTA
def render_pivot_tipo_renta_html(df_data):
    df_temp = df_data.copy()
    df_valid = df_temp.dropna(subset=["Tipo Renta Fija"]).copy()

    pivot = df_valid.groupby("Tipo Renta Fija").agg(
        Suma_Interes=("INTERESE NETO", "sum"),
        Cuenta_Tipo=("Tipo Renta Fija", "count")
    ).reset_index()

    total_sum = pivot["Suma_Interes"].sum()
    total_count = pivot["Cuenta_Tipo"].sum()

    css = """<style>
.pvt-tbl-wrapper {
    max-height: 520px; width: 100%; overflow-x: auto; overflow-y: auto;
    border: 1px solid #475569; border-radius: 8px; background-color: #0F172A;
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.5); margin-top: 8px; margin-bottom: 24px;
    padding: 0px !important; display: block; position: relative;
}
.pvt-tbl-sticky {
    width: 100%; border-collapse: collapse !important; border-spacing: 0 !important;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    font-size: 0.85rem; margin: 0; table-layout: auto !important;
}
.pvt-tbl-sticky th {
    position: sticky !important; top: 0 !important; z-index: 20 !important;
    background-color: #1E293B !important; color: #FFFFFF !important;
    border: 1px solid #475569 !important; padding: 6px 10px !important;
    text-align: center !important; font-weight: 800 !important; white-space: nowrap !important;
    line-height: 1.2 !important;
}
.pvt-tbl-sticky th:first-child {
    position: sticky !important; top: 0 !important; left: 0 !important; z-index: 50 !important;
    text-align: center !important; box-shadow: 2px 0 5px rgba(0,0,0,0.4) !important;
    white-space: nowrap !important; background-color: #1E293B !important;
}
.pvt-tbl-sticky td {
    position: static !important; z-index: auto !important; padding: 6px 10px !important;
    border: 1px solid #334155 !important; vertical-align: middle !important;
    text-align: center !important; color: #000000 !important; white-space: nowrap !important;
}
.pvt-tbl-sticky .lbl-sticky-col {
    position: sticky !important; left: 0 !important; z-index: 30 !important; font-weight: 800 !important;
    text-align: center !important; border: 1px solid #334155 !important;
    box-shadow: 3px 0 6px rgba(0,0,0,0.4) !important; background-clip: padding-box !important;
    white-space: nowrap !important;
}
.row-pvt-green td { background-color: #ECFDF5 !important; color: #000000 !important; font-weight: 600 !important; }
.row-pvt-green .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-pvt-yellow td { background-color: #FEF3C7 !important; color: #000000 !important; font-weight: 600 !important; }
.row-pvt-yellow .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-pvt-white td { background-color: #F8FAFC !important; color: #000000 !important; font-weight: 600 !important; }
.row-pvt-white .lbl-sticky-col { background-color: #FACC15 !important; color: #000000 !important; }

.row-pvt-total td { background-color: #BAE6FD !important; color: #000000 !important; font-weight: 900 !important; }
.row-pvt-total .lbl-sticky-col { background-color: #38BDF8 !important; color: #000000 !important; font-weight: 900 !important; }
</style>"""

    html = f'{css}<div class="pvt-tbl-wrapper"><table class="pvt-tbl-sticky">'
    html += '<thead><tr>'
    html += f'<th>{dos_lineas_titulo("Etiquetas de fila")}</th>'
    html += f'<th>{dos_lineas_titulo("Suma de INTERESE NETO")}</th>'
    html += f'<th>{dos_lineas_titulo("Cuenta de Tipo Renta Fija")}</th>'
    html += '</tr></thead><tbody>'

    row_classes = ["row-pvt-green", "row-pvt-yellow", "row-pvt-white"]

    for idx, row in pivot.iterrows():
        tipo = str(row["Tipo Renta Fija"])
        val_sum = row["Suma_Interes"]
        val_cnt = int(row["Cuenta_Tipo"])
        
        tipo_str = tipo.replace(" ", "<br>")
        if tipo_str == "AustroFuturo": tipo_str = "Austro<br>Futuro"

        val_sum_str = "-" if (val_sum == 0.0 or pd.isna(val_sum)) else f"${val_sum:,.2f}"

        r_cls = row_classes[idx % len(row_classes)]
        html += f'<tr class="{r_cls}">'
        html += f'<td class="lbl-sticky-col">{tipo_str}</td>'
        html += f'<td style="font-weight: 800;">{val_sum_str}</td>'
        html += f'<td style="font-weight: 800;">{val_cnt}</td>'
        html += '</tr>'

    # Fila Total General
    tot_sum_str = f"${total_sum:,.2f}"
    html += f'<tr class="row-pvt-total">'
    html += f'<td class="lbl-sticky-col">Total<br>general</td>'
    html += f'<td style="font-weight: 900;">{tot_sum_str}</td>'
    html += f'<td style="font-weight: 900;">{total_count}</td>'
    html += '</tr>'

    html += '</tbody></table></div>'
    return html.replace("\n", " ")


tab_int_ano, tab_int_entidad, tab_int_tipo = st.tabs([
    "💰 Interés Neto por Año y Tipo",
    "🏦 Interés Neto por Entidad",
    "📑 Resumen por Tipo de Renta"
])

with tab_int_ano:
    st.markdown(render_pivot_interes_neto_html(df_datos), unsafe_allow_html=True)

with tab_int_entidad:
    st.markdown(render_pivot_entidad_html(df_datos), unsafe_allow_html=True)

with tab_int_tipo:
    st.markdown(render_pivot_tipo_renta_html(df_datos), unsafe_allow_html=True)


# ---------------------------------------------------------
# SECCIÓN TERCERA: GRÁFICOS DINÁMICOS
# ---------------------------------------------------------
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    "<h4 style='color: #38BDF8 !important; margin-bottom: 4px;'>📈 Gráficos Dinámicos</h4>",
    unsafe_allow_html=True,
)

tab_chart_entidad, tab_chart_tipo, tab_chart_mas = st.tabs([
    "🥧 Interés Neto por Entidad",
    "📊 Interés Neto por Tipo de Renta",
    "➕ Más Gráficos"
])

with tab_chart_entidad:
    df_chart_entidad = df_datos.dropna(subset=["Entidad"]).groupby("Entidad", as_index=False)["INTERESE NETO"].sum()
    
    # Formatear el nombre de cada entidad en dos líneas (<br>)
    df_chart_entidad["Entidad_Leyenda"] = df_chart_entidad["Entidad"].apply(lambda x: str(x).strip().replace(" ", "<br>"))
    
    colors_distintivos = ["#0284C7", "#38BDF8", "#A855F7", "#EC4899", "#06B6D4", "#818CF8"]
    
    fig_entidad = px.pie(
        df_chart_entidad,
        names="Entidad_Leyenda",
        values="INTERESE NETO",
        title="Total",
        color_discrete_sequence=colors_distintivos
    )
    
    fig_entidad.update_traces(
        textposition='inside',
        textinfo='percent',
        textfont_size=14,
        textfont_color='white',
        marker=dict(line=dict(color='#0F172A', width=2))
    )
    
    fig_entidad.update_layout(
        template="plotly_dark",
        paper_bgcolor="#0F172A",
        plot_bgcolor="#0F172A",
        font=dict(color="#FFFFFF", size=13),
        title_x=0.5,
        title_xanchor="center",
        title_font=dict(size=20, color="#FFFFFF"),
        legend_title_text="",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.15,
            xanchor="center",
            x=0.5,
            font=dict(size=16)
        ),
        margin=dict(l=10, r=10, t=40, b=80),
        height=450
    )
    
    # Bloquear zoom/pan en táctil y ocultar modebar
    st.plotly_chart(
        fig_entidad,
        use_container_width=True,
        config={
            'displayModeBar': False,
            'scrollZoom': False,
            'doubleClick': False
        }
    )

with tab_chart_tipo:
    df_chart_tipo = (
        df_datos.dropna(subset=["Tipo Renta Fija"])
        .groupby("Tipo Renta Fija", as_index=False)["INTERESE NETO"]
        .sum()
        .sort_values(by="INTERESE NETO", ascending=True)
    )
    
    # Formatear etiquetas en 2 líneas
    df_chart_tipo["Tipo_Label"] = df_chart_tipo["Tipo Renta Fija"].apply(
        lambda x: "Austro<br>Futuro" if str(x).strip() == "AustroFuturo" else str(x).strip().replace(" ", "<br>")
    )
    df_chart_tipo["Texto_Monto"] = df_chart_tipo["INTERESE NETO"].apply(lambda x: f"${x:,.2f}")

    fig_tipo = px.bar(
        df_chart_tipo,
        x="INTERESE NETO",
        y="Tipo_Label",
        orientation="h",
        text="Texto_Monto",
        color="Tipo_Label",
        color_discrete_sequence=["#38BDF8", "#0284C7", "#A855F7", "#EC4899", "#06B6D4"],
        title="SUMA DE INTERESE NETO"
    )

    # Posicionamiento inteligente del texto: dentro si cabe, a la derecha si la barra es corta
    fig_tipo.update_traces(
        textposition='auto',
        textfont_size=13,
        textfont_color='white',
        insidetextanchor='end',
        marker=dict(line=dict(color='#0F172A', width=1))
    )

    fig_tipo.update_layout(
        template="plotly_dark",
        paper_bgcolor="#0F172A",
        plot_bgcolor="#0F172A",
        font=dict(color="#FFFFFF", size=13),
        title_x=0.5,
        title_xanchor="center",
        title_font=dict(size=20, color="#FFFFFF"),
        showlegend=False,
        # fixedrange=True para evitar zoom/pan al tocar en pantallas móviles
        xaxis=dict(title="", showgrid=True, gridcolor="#334155", tickprefix="$", fixedrange=True),
        yaxis=dict(title="", fixedrange=True),
        margin=dict(l=10, r=60, t=50, b=30),
        height=380
    )

    # Bloquear interacciones de zoom en móvil y ocultar barra de botones flotante
    st.plotly_chart(
        fig_tipo,
        use_container_width=True,
        config={
            'displayModeBar': False,
            'scrollZoom': False,
            'doubleClick': False
        }
    )

with tab_chart_mas:
    st.info("Pestaña disponible para agregar más gráficos dinámicos.")


# ---------------------------------------------------------
# SECCIÓN DETALLES: TABLA DE DETALLES (2 COLUMNAS FIJAS)
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