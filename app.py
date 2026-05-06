
from pathlib import Path
import hmac
import json
import math

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import folium
from folium.plugins import MarkerCluster, HeatMap, Fullscreen, MiniMap
from streamlit_folium import st_folium

st.cache_data.clear()

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Dashboard territorial-electoral Kennedy",
    page_icon="🗳️",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_DIR = Path("data")
CONSOLIDADO = DATA_DIR / "kennedy_mira_consolidado.xlsx"
UPZ_GEOJSON = DATA_DIR / "upz_kennedy.geojson"
KENNEDY_CENTER = [4.6260, -74.1570]

COLOR_TEXT = "#0F172A"
COLOR_MUTED = "#334155"
COLOR_BORDER = "#E2E8F0"
COLOR_RED = "#DC2626"
COLOR_GREEN = "#16A34A"
COLOR_BLUE = "#2563EB"
COLOR_ORANGE = "#F97316"
COLOR_PURPLE = "#7C3AED"


# ============================================================
# ESTILOS
# ============================================================

st.markdown(
    """
    <style>
    .stApp {
        background: #F8FAFC;
        color: #0F172A;
    }

    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1500px;
    }

    h1, h2, h3, h4, h5, h6 {
        color: #0F172A !important;
        font-weight: 800 !important;
        letter-spacing: -0.02em;
    }

    p, span, div, label {
        color: #1E293B;
    }

    .subtitle {
        color: #334155;
        font-size: 1.02rem;
        margin-top: -0.6rem;
        margin-bottom: 1.2rem;
    }

    .section-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 1.2rem 1.35rem;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05);
        margin-bottom: 1rem;
    }

    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 1.1rem 1.2rem;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
        min-height: 122px;
    }

    .metric-label {
        color: #475569;
        font-size: 0.9rem;
        font-weight: 700;
        margin-bottom: 0.35rem;
    }

    .metric-value {
        color: #0F172A;
        font-size: 2.15rem;
        font-weight: 850;
        line-height: 1.1;
    }

    .metric-delta-positive {
        display: inline-block;
        background: #DCFCE7;
        color: #166534;
        border-radius: 999px;
        padding: 0.18rem 0.55rem;
        font-weight: 800;
        font-size: 0.85rem;
        margin-top: 0.5rem;
    }

    .metric-delta-negative {
        display: inline-block;
        background: #FEE2E2;
        color: #991B1B;
        border-radius: 999px;
        padding: 0.18rem 0.55rem;
        font-weight: 800;
        font-size: 0.85rem;
        margin-top: 0.5rem;
    }

    [data-testid="stTabs"] button {
        color: #334155 !important;
        font-weight: 700 !important;
    }

    [data-testid="stTabs"] button[aria-selected="true"] {
        color: #DC2626 !important;
        border-bottom-color: #DC2626 !important;
    }

    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }

    section[data-testid="stSidebar"] * {
        color: #0F172A !important;
    }

    div[data-testid="stDataFrame"] {
        background: #FFFFFF;
        border-radius: 12px;
    }

    div[data-testid="stExpander"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
    }

    div[data-testid="stExpander"] * {
        color: #1E293B !important;
    }

    .note-box {
        background: #EFF6FF;
        border-left: 4px solid #2563EB;
        border-radius: 12px;
        padding: 0.9rem 1rem;
        margin: 0.7rem 0;
        color: #1E3A8A;
    }

    .warning-box {
        background: #FEF3C7;
        border-left: 4px solid #F59E0B;
        border-radius: 12px;
        padding: 0.9rem 1rem;
        margin: 0.7rem 0;
        color: #78350F;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# AUTENTICACIÓN OPCIONAL
# ============================================================

def check_password():
    """Si existen secrets [auth], exige login. Si no existen, deja entrar."""
    try:
        auth_cfg = st.secrets.get("auth", None)
    except Exception:
        auth_cfg = None

    if not auth_cfg:
        return True

    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if st.session_state["autenticado"]:
        return True

    st.title("Acceso restringido")
    st.caption("Dashboard territorial-electoral Kennedy MIRA")

    usuario_correcto = str(auth_cfg.get("usuario", ""))
    password_correcto = str(auth_cfg.get("password", ""))

    with st.form("login_form"):
        usuario = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        ingresar = st.form_submit_button("Ingresar")

    if ingresar:
        if hmac.compare_digest(usuario, usuario_correcto) and hmac.compare_digest(password, password_correcto):
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")

    return False


if not check_password():
    st.stop()


# ============================================================
# FUNCIONES
# ============================================================

@st.cache_data
def cargar_datos(path: Path):
    if not path.exists():
        return None

    xls = pd.ExcelFile(path)

    def read(sheet):
        if sheet in xls.sheet_names:
            return pd.read_excel(path, sheet_name=sheet)
        return pd.DataFrame()

    data = {
        "resumen_general": read("resumen_general"),
        "puestos": read("puestos_votacion"),
        "actividades": read("actividades_campana"),
        "mesas": read("mesas_trabajo"),
        "iglesias": read("iglesias"),
        "resumen_iglesia": read("resumen_iglesia"),
        "resumen_puesto": read("resumen_puesto"),
        "resumen_barrio": read("resumen_barrio"),
        "matriz": read("matriz_priorizacion"),
        "informe": read("informe_ejecutivo"),
        "control": read("control_calidad"),
    }

    for key in ["puestos", "actividades", "mesas", "iglesias", "resumen_iglesia", "matriz"]:
        if key in data and not data[key].empty:
            for col in data[key].columns:
                if "FECHA" in str(col).upper():
                    data[key][col] = pd.to_datetime(data[key][col], errors="coerce")

    return data


def fmt_number(value, decimals=0):
    if pd.isna(value):
        return "N.D."
    return f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(value):
    if pd.isna(value):
        return "N.D."
    return f"{value:.1%}".replace(".", ",")


def get_indicador(df, indicador, default=np.nan):
    if df is None or df.empty:
        return default
    row = df[df["INDICADOR"].astype(str).str.strip().eq(indicador)]
    if row.empty:
        return default
    return row.iloc[0]["VALOR"]


def metric_card(label, value, delta=None, positive=True):
    delta_html = ""
    if delta is not None:
        cls = "metric-delta-positive" if positive else "metric-delta-negative"
        arrow = "↑" if positive else "↓"
        delta_html = f'<div class="{cls}">{arrow} {delta}</div>'
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def aplicar_filtros(puestos, actividades, mesas, filtros):
    iglesias_sel = filtros.get("iglesias", [])
    prioridad_sel = filtros.get("prioridad", [])
    puestos_f = puestos.copy()
    acts_f = actividades.copy()
    mesas_f = mesas.copy()

    if iglesias_sel:
        puestos_f = puestos_f[puestos_f["IGLESIA"].isin(iglesias_sel)]
        if "IGLESIA" in acts_f.columns:
            acts_f = acts_f[acts_f["IGLESIA"].isin(iglesias_sel)]
        if "IGLESIA" in mesas_f.columns:
            mesas_f = mesas_f[mesas_f["IGLESIA"].isin(iglesias_sel)]

    if prioridad_sel and "PRIORIDAD" in puestos_f.columns:
        puestos_f = puestos_f[puestos_f["PRIORIDAD"].isin(prioridad_sel)]

    if "ESTRATEGIA" in acts_f.columns:
        valid_estrategias = ["LIBERTAD RELIGIOSA", "POLITICO COMUNITARIA", "POLÍTICO COMUNITARIA"]
        acts_f = acts_f[acts_f["ESTRATEGIA"].str.strip().str.upper().isin(valid_estrategias)]

    return puestos_f, acts_f, mesas_f


def crear_mapa(puestos, iglesias, actividades, mesas):
    m = folium.Map(location=KENNEDY_CENTER, zoom_start=13, tiles="CartoDB positron", control_scale=True)
    Fullscreen(position="topleft").add_to(m)
    MiniMap(toggle_display=True, position="bottomleft").add_to(m)

    if UPZ_GEOJSON.exists():
        try:
            with open(UPZ_GEOJSON, "r", encoding="utf-8") as f:
                gj = json.load(f)
            folium.GeoJson(
                gj,
                name="UPZ Kennedy",
                style_function=lambda feature: {
                    "fillColor": "#EFF6FF",
                    "color": "#2563EB",
                    "weight": 1.2,
                    "fillOpacity": 0.08,
                },
                tooltip=folium.GeoJsonTooltip(fields=[], aliases=[]),
            ).add_to(m)
        except Exception:
            pass

    # Heatmap
    valid_heat = puestos.dropna(subset=["LATITUD", "LONGITUD", "VOTOS_2026"]).copy()
    heat_data = valid_heat[["LATITUD", "LONGITUD", "VOTOS_2026"]].values.tolist()
    if heat_data:
        heat_layer = folium.FeatureGroup(name="Calor votos 2026", show=False)
        HeatMap(heat_data, radius=28, blur=18, min_opacity=0.25).add_to(heat_layer)
        heat_layer.add_to(m)

    # Puestos
    puestos_layer = folium.FeatureGroup(name="Puestos de votación", show=True)
    cluster = MarkerCluster(name="Cluster puestos").add_to(puestos_layer)
    for _, r in puestos.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        var = r.get("VARIACION_ABSOLUTA", np.nan)
        color = "green" if pd.notna(var) and var > 0 else "red" if pd.notna(var) and var < 0 else "gray"
        popup = f"""
        <div style="font-family:Arial; width:330px;">
        <h4 style="margin-bottom:6px;">{r.get('PUESTO','')}</h4>
        <b>Iglesia:</b> {r.get('IGLESIA','')}<br>
        <b>Barrio:</b> {r.get('BARRIO','') or 'Sin dato'}<br>
        <b>UPZ:</b> {r.get('UPZ','') or 'Sin dato'}<br>
        <b>Votos 2026:</b> {fmt_number(r.get('VOTOS_2026'),1)}<br>
        <b>Votos 2023:</b> {fmt_number(r.get('VOTOS_2023'),1)}<br>
        <b>Variación:</b> {fmt_number(r.get('VARIACION_ABSOLUTA'),1)} ({fmt_pct(r.get('VARIACION_PORCENTUAL'))})<br>
        <b>Prioridad:</b> {r.get('PRIORIDAD','')}<br>
        <b>Acción:</b> {r.get('ACCION_RECOMENDADA','')}<br>
        </div>
        """
        radius = max(5, min(17, float(r.get("VOTOS_2026", 0) or 0) / 14))
        folium.CircleMarker(
            location=[r["LATITUD"], r["LONGITUD"]],
            radius=radius,
            popup=folium.Popup(popup, max_width=380),
            tooltip=f"{r.get('PUESTO','')} | {r.get('IGLESIA','')} | {fmt_number(r.get('VOTOS_2026'),1)}",
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.72,
            weight=1.2,
        ).add_to(cluster)
    puestos_layer.add_to(m)

    # Churches
    iglesia_layer = folium.FeatureGroup(name="Iglesias / templos", show=True)
    for _, r in iglesias.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        url = r.get("URL", "")
        link = f'<br><a href="{url}" target="_blank">Ver ubicación IDMJI</a>' if isinstance(url, str) and url else ""
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            popup=folium.Popup(
                f"<b>{r.get('IGLESIA','')}</b><br>Lat: {r['LATITUD']}<br>Lon: {r['LONGITUD']}{link}",
                max_width=280,
            ),
            tooltip=f"Iglesia: {r.get('IGLESIA','')}",
            icon=folium.Icon(color="purple", icon="home", prefix="fa"),
        ).add_to(iglesia_layer)
    iglesia_layer.add_to(m)

    # Activities
    acts_layer = folium.FeatureGroup(name="Actividades de campaña", show=False)
    for _, r in actividades.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        folium.CircleMarker(
            location=[r["LATITUD"], r["LONGITUD"]],
            radius=4,
            color="blue",
            fill=True,
            fill_opacity=0.6,
            tooltip=f"{r.get('TIPO_ACTIVIDAD','')} | {r.get('IGLESIA','')}",
            popup=folium.Popup(
                f"<b>Actividad:</b> {r.get('TIPO_ACTIVIDAD','')}<br><b>Iglesia:</b> {r.get('IGLESIA','')}<br><b>Barrio:</b> {r.get('BARRIO','')}",
                max_width=280,
            ),
        ).add_to(acts_layer)
    acts_layer.add_to(m)

    # Mesas
    mesas_layer = folium.FeatureGroup(name="Mesas de trabajo", show=False)
    for _, r in mesas.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            tooltip=f"Mesa | {r.get('IGLESIA','')} | {r.get('BARRIO','')}",
            popup=folium.Popup(
                f"<b>Mesa:</b> {r.get('TEMA','')}<br><b>Iglesia:</b> {r.get('IGLESIA','')}<br><b>Barrio:</b> {r.get('BARRIO','')}<br><b>Estado:</b> {r.get('ESTADO','')}",
                max_width=320,
            ),
            icon=folium.Icon(color="orange", icon="info-sign"),
        ).add_to(mesas_layer)
    mesas_layer.add_to(m)

    legend_html = """
    <div style="position: fixed; bottom: 35px; right: 35px; z-index:9999; background:white; padding:12px 14px; border:1px solid #CBD5E1; border-radius:10px; box-shadow:0 3px 12px rgba(0,0,0,.12); font-size:13px;">
    <b>Lectura del mapa</b><br>
    <span style="color:green;">●</span> Puesto con crecimiento<br>
    <span style="color:red;">●</span> Puesto con caída<br>
    <span style="color:gray;">●</span> Sin comparación<br>
    <span style="color:purple;">⬟</span> Iglesia / templo<br>
    <span style="color:orange;">⬟</span> Mesa de trabajo<br>
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))
    folium.LayerControl(collapsed=False).add_to(m)
    return m


def download_excel_link():
    with open(CONSOLIDADO, "rb") as f:
        return f.read()


# ============================================================
# CARGA
# ============================================================

data = cargar_datos(CONSOLIDADO)

if data is None:
    st.error("No se encontró `data/kennedy_mira_consolidado.xlsx`. Sube el Excel consolidado a la carpeta `data/` del repositorio.")
    st.stop()

resumen_general = data["resumen_general"]
puestos = data["puestos"]
actividades = data["actividades"]
mesas = data["mesas"]
iglesias = data["iglesias"]
resumen_iglesia = data["resumen_iglesia"]
resumen_puesto = data["resumen_puesto"]
resumen_barrio = data["resumen_barrio"]
matriz = data["matriz"]
informe = data["informe"]
control = data.get("control", pd.DataFrame())

# Filtrar iglesias oficiales permitidas
IGLESIAS_OFICIALES_PERMITIDAS = ["CLASS ROMA", "KENNEDY CENTRAL", "PATIO BONITO", "CARVAJAL", "VALLADOLID"]

def filtrar_iglesias(df):
    if df is not None and not df.empty and "IGLESIA" in df.columns:
        return df[df["IGLESIA"].isin(IGLESIAS_OFICIALES_PERMITIDAS)].copy()
    return df

# Agregar los registros descartados a control de calidad
def agregar_a_control(df, nombre_origen):
    global control
    if df is not None and not df.empty and "IGLESIA" in df.columns:
        descartados = df[~df["IGLESIA"].isin(IGLESIAS_OFICIALES_PERMITIDAS)].copy()
        if not descartados.empty:
            descartados["ORIGEN_ERROR"] = f"Iglesia no oficial en {nombre_origen}"
            if control is None or control.empty:
                control = pd.DataFrame(columns=descartados.columns)
            # Añadir columnas faltantes a control
            for col in descartados.columns:
                if col not in control.columns:
                    control[col] = pd.Series(dtype=descartados[col].dtype)
            control = pd.concat([control, descartados], ignore_index=True)

agregar_a_control(puestos, "puestos")
agregar_a_control(actividades, "actividades")
agregar_a_control(mesas, "mesas")
agregar_a_control(iglesias, "iglesias")
agregar_a_control(resumen_iglesia, "resumen_iglesia")

puestos = filtrar_iglesias(puestos)
actividades = filtrar_iglesias(actividades)
mesas = filtrar_iglesias(mesas)
iglesias = filtrar_iglesias(iglesias)
resumen_iglesia = filtrar_iglesias(resumen_iglesia)

# Ensure numerics
for df in [puestos, resumen_iglesia, resumen_puesto, resumen_barrio, matriz]:
    if not df.empty:
        for col in ["VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "LATITUD", "LONGITUD"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("Configuración del análisis")
    st.markdown("Fuente única: `kennedy_mira_consolidado.xlsx`")

    iglesias_oficiales = IGLESIAS_OFICIALES_PERMITIDAS
    default_iglesias = [i for i in iglesias_oficiales if i in puestos["IGLESIA"].unique()]
    selected_iglesias = st.multiselect("Iglesias / templos", iglesias_oficiales, default=default_iglesias)

    prioridades = sorted(puestos["PRIORIDAD"].dropna().unique().tolist()) if "PRIORIDAD" in puestos.columns else []
    selected_prioridades = st.multiselect("Prioridad", prioridades, default=prioridades)

    st.divider()
    st.caption("Seguridad")
    if st.session_state.get("autenticado"):
        if st.button("Cerrar sesión"):
            st.session_state["autenticado"] = False
            st.rerun()

    st.divider()
    if UPZ_GEOJSON.exists():
        st.success("Capa UPZ detectada.")


puestos_f, actividades_f, mesas_f = aplicar_filtros(
    puestos,
    actividades,
    mesas,
    {"iglesias": selected_iglesias, "prioridad": selected_prioridades},
)

resumen_iglesia_f = resumen_iglesia[resumen_iglesia["IGLESIA"].isin(selected_iglesias)].copy()


# ============================================================
# ENCABEZADO
# ============================================================

st.title("Dashboard territorial-electoral Kennedy")
st.markdown(
    '<div class="subtitle">Campaña Congreso 2026 · Partido MIRA · Votación, gestión, iglesias, puestos, barrios, UPZ y priorización territorial</div>',
    unsafe_allow_html=True,
)

# Métricas oficiales
total_2026 = get_indicador(resumen_general, "Total Kennedy votos promedio 2026")
total_2023 = get_indicador(resumen_general, "Total Kennedy votos promedio 2023")
var_abs = get_indicador(resumen_general, "Variación absoluta Kennedy")
var_pct = get_indicador(resumen_general, "Variación porcentual Kennedy")
puestos_total = get_indicador(resumen_general, "Puestos totales analizados")

if "ESTRATEGIA" in actividades_f.columns:
    actividades_total = len(actividades_f)
else:
    actividades_total = get_indicador(resumen_general, "Actividades de campaña consolidadas")

mesas_total = get_indicador(resumen_general, "Mesas incorporadas reporte Bogotá Tecnología Kennedy")
if pd.isna(mesas_total):
    mesas_total = get_indicador(resumen_general, "Mesas de trabajo consolidadas")
iglesias_total = get_indicador(resumen_general, "Iglesias oficiales")
jal_total = get_indicador(resumen_general, "JAL 2023 Kennedy")
concejo_total = get_indicador(resumen_general, "Concejo 2023 Kennedy")
camara_total = get_indicador(resumen_general, "Cámara 2026 Kennedy")
senado_total = get_indicador(resumen_general, "Senado 2026 Kennedy")
mesas_2026_reporte = get_indicador(resumen_general, "Mesas 2026 reporte localidad")
testigos_2023_reporte = get_indicador(resumen_general, "Testigos 2023 reporte localidad")

c1, c2, c3, c4 = st.columns(4)
with c1:
    metric_card("Total general Kennedy 2026", fmt_number(total_2026, 1))
with c2:
    metric_card("Total general Kennedy 2023", fmt_number(total_2023, 1))
with c3:
    metric_card("Variación electoral", fmt_number(var_abs, 1), fmt_number(abs(var_abs), 1), positive=var_abs >= 0)
with c4:
    metric_card("Variación porcentual", fmt_pct(var_pct), fmt_pct(abs(var_pct)), positive=var_pct >= 0)

c5, c6, c7, c8 = st.columns(4)
with c5:
    metric_card("Puestos analizados", fmt_number(puestos_total, 0))
with c6:
    metric_card("Actividades de campaña", fmt_number(actividades_total, 0))
with c7:
    metric_card("Mesas de trabajo", fmt_number(mesas_total, 0))
with c8:
    metric_card("Iglesias oficiales", fmt_number(iglesias_total, 0))

c9, c10, c11, c12 = st.columns(4)
with c9:
    metric_card("JAL 2023", fmt_number(jal_total, 1))
with c10:
    metric_card("Concejo 2023", fmt_number(concejo_total, 1))
with c11:
    metric_card("Cámara 2026", fmt_number(camara_total, 1))
with c12:
    metric_card("Senado 2026", fmt_number(senado_total, 1))


# ============================================================
# TABS
# ============================================================

tab_resumen, tab_mapa, tab_iglesia, tab_puesto, tab_barrio, tab_prioridad, tab_export = st.tabs(
    [
        "Resumen ejecutivo",
        "Mapa territorial",
        "Análisis por iglesia",
        "Análisis por puesto",
        "Barrio / UPZ",
        "Priorización",
        "Exportables",
    ]
)

with tab_resumen:
    st.subheader("Portada y lectura ejecutiva")
    st.markdown(
        """
        <div class="section-card">
        Este tablero consolida la votación promedio 2026, la base comparativa 2023, actividades de campaña, mesas de trabajo e iglesias responsables.
        La lectura está orientada a identificar <b>puestos de recuperación</b>, <b>puestos de consolidación</b> y <b>brechas entre presencia territorial y resultado electoral</b>.
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not informe.empty:
        for _, r in informe.iterrows():
            with st.expander(str(r["SECCION"]), expanded=(r["SECCION"] == "Resumen general")):
                st.write(r["TEXTO"])

    st.markdown("### Hallazgos automáticos")
    col1, col2 = st.columns(2)
    with col1:
        top_pos = puestos.sort_values("VARIACION_ABSOLUTA", ascending=False).head(10)
        fig = px.bar(
            top_pos.sort_values("VARIACION_ABSOLUTA"),
            x="VARIACION_ABSOLUTA",
            y="PUESTO",
            orientation="h",
            color="IGLESIA",
            title="Top 10 puestos con mayor crecimiento",
        )
        fig.update_layout(height=420, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        top_neg = puestos.sort_values("VARIACION_ABSOLUTA", ascending=True).head(10)
        fig = px.bar(
            top_neg.sort_values("VARIACION_ABSOLUTA", ascending=False),
            x="VARIACION_ABSOLUTA",
            y="PUESTO",
            orientation="h",
            color="IGLESIA",
            title="Top 10 puestos con mayor caída",
        )
        fig.update_layout(height=420, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Desglose de actividades de campaña")
    if not actividades_f.empty and "TIPO_ACTIVIDAD" in actividades_f.columns:
        acts_counts = actividades_f["TIPO_ACTIVIDAD"].fillna("NO ESPECIFICADO").value_counts().reset_index()
        acts_counts.columns = ["Tipo de Actividad", "Cantidad"]
        fig_acts = px.bar(
            acts_counts.sort_values("Cantidad", ascending=True),
            x="Cantidad",
            y="Tipo de Actividad",
            orientation="h",
            title="Distribución por tipo",
        )
        fig_acts.update_layout(height=350, showlegend=False, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
        
        ca1, ca2 = st.columns([1, 2])
        with ca1:
            st.dataframe(acts_counts, use_container_width=True, hide_index=True)
        with ca2:
            st.plotly_chart(fig_acts, use_container_width=True)
    else:
        st.info("No hay actividades para desglosar con los filtros actuales.")

    st.markdown("### Desglose electoral por corporación")
    jal_23 = puestos_f["JAL_2023"].sum() if "JAL_2023" in puestos_f.columns else 0
    concejo_23 = puestos_f["MIRA_CONCEJO_2023"].sum() if "MIRA_CONCEJO_2023" in puestos_f.columns else 0
    camara_26 = puestos_f["CAMARA_2026"].sum() if "CAMARA_2026" in puestos_f.columns else 0
    senado_26 = puestos_f["SENADO_2026"].sum() if "SENADO_2026" in puestos_f.columns else 0
    
    elec_data = pd.DataFrame(
        {
            "Corporación": ["JAL", "Concejo", "Cámara", "Senado"],
            "Votos": [jal_23, concejo_23, camara_26, senado_26],
            "Bloque de análisis": ["JAL / Concejo 2023", "JAL / Concejo 2023", "Cámara / Senado 2026", "Cámara / Senado 2026"],
        }
    )
    
    ce1, ce2 = st.columns([1, 2])
    with ce1:
        st.markdown("**JAL / Concejo 2023**")
        st.dataframe(elec_data[elec_data["Bloque de análisis"].eq("JAL / Concejo 2023")], hide_index=True, use_container_width=True)
        st.markdown("**Cámara / Senado 2026**")
        st.dataframe(elec_data[elec_data["Bloque de análisis"].eq("Cámara / Senado 2026")], hide_index=True, use_container_width=True)
        st.markdown("**Variables operativas del nuevo reporte**")
        st.dataframe(
            pd.DataFrame(
                {
                    "Indicador": ["Mesas 2026", "Testigos 2023"],
                    "Valor": [mesas_2026_reporte, testigos_2023_reporte],
                }
            ),
            hide_index=True,
            use_container_width=True,
        )
    with ce2:
        fig_elec = px.bar(
            elec_data, 
            x="Corporación", 
            y="Votos", 
            color="Bloque de análisis",
            text_auto=".1f",
            title="Comparación separada por corporación"
        )
        fig_elec.update_layout(height=350, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
        st.plotly_chart(fig_elec, use_container_width=True)

with tab_mapa:
    st.subheader("Mapa interactivo territorial")
    st.markdown('<div class="note-box">Active o desactive capas para comparar puestos de votación, iglesias, mesas de trabajo, actividades y calor electoral 2026.</div>', unsafe_allow_html=True)
    mapa = crear_mapa(puestos_f, iglesias, actividades_f, mesas_f)
    st_folium(mapa, width=None, height=720)

with tab_iglesia:
    st.subheader("Análisis por iglesia")
    cols_show = [
        "IGLESIA", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL",
        "PUESTOS", "ACTIVIDADES_CAMPANA", "MESAS_TRABAJO", "PUESTO_MAYOR_VOTACION",
        "PUESTO_MAYOR_CAIDA", "PUESTO_MAYOR_CRECIMIENTO"
    ]
    st.dataframe(resumen_iglesia_f[cols_show], use_container_width=True, hide_index=True)

    col1, col2 = st.columns(2)
    with col1:
        fig = px.bar(
            resumen_iglesia_f.sort_values("VOTOS_2026"),
            x="VOTOS_2026",
            y="IGLESIA",
            orientation="h",
            title="Votos 2026 por iglesia",
            color="IGLESIA",
        )
        fig.update_layout(height=420, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        fig = go.Figure()
        ri = resumen_iglesia_f.sort_values("VOTOS_2026", ascending=False)
        fig.add_trace(go.Bar(x=ri["IGLESIA"], y=ri["VOTOS_2023"], name="2023"))
        fig.add_trace(go.Bar(x=ri["IGLESIA"], y=ri["VOTOS_2026"], name="2026"))
        fig.update_layout(
            title="Comparación 2023 vs 2026 por iglesia",
            barmode="group",
            height=420,
            paper_bgcolor="white",
            plot_bgcolor="white",
            font_color=COLOR_TEXT,
        )
        st.plotly_chart(fig, use_container_width=True)

    for iglesia in iglesias["IGLESIA"].tolist():
        sub = puestos[puestos["IGLESIA"].eq(iglesia)].copy()
        res = resumen_iglesia[resumen_iglesia["IGLESIA"].eq(iglesia)]
        with st.expander(f"Lectura estratégica: {iglesia}", expanded=False):
            if not res.empty:
                st.write(res.iloc[0].get("LECTURA_ESTRATEGICA", ""))
                st.write("**Recomendación:**", res.iloc[0].get("RECOMENDACION", ""))
            if sub.empty:
                st.info("No hay puestos asignados en la matriz electoral. Se mantiene como iglesia oficial para análisis territorial.")
            else:
                st.dataframe(
                    sub[["PUESTO", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "PRIORIDAD", "ACCION_RECOMENDADA"]]
                    .sort_values("VOTOS_2026", ascending=False),
                    use_container_width=True,
                    hide_index=True,
                )

with tab_puesto:
    st.subheader("Análisis por puesto de votación")

    col1, col2 = st.columns(2)
    with col1:
        top_growth = puestos_f.sort_values("VARIACION_ABSOLUTA", ascending=False).head(15)
        fig = px.bar(
            top_growth.sort_values("VARIACION_ABSOLUTA"),
            x="VARIACION_ABSOLUTA",
            y="PUESTO",
            orientation="h",
            color="IGLESIA",
            title="Puestos con mayor crecimiento",
        )
        fig.update_layout(height=560, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        top_drop = puestos_f.sort_values("VARIACION_ABSOLUTA", ascending=True).head(15)
        fig = px.bar(
            top_drop.sort_values("VARIACION_ABSOLUTA", ascending=False),
            x="VARIACION_ABSOLUTA",
            y="PUESTO",
            orientation="h",
            color="IGLESIA",
            title="Puestos con mayor caída",
        )
        fig.update_layout(height=560, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Matriz analítica por puesto")
    cols_puesto = [
        "PUESTO", "IGLESIA", "MIRA_CONCEJO_2023", "JAL_2023", "CAMARA_2026", "SENADO_2026",
        "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL",
        "MESAS_2026_REPORTE", "TESTIGOS_2023_REPORTE", "VOTOS_AFINIDAD_E11_2023",
        "TEMPLO_REPORTE", "TIENE_MESA_TRABAJO", "PRIORIDAD", "ACCION_RECOMENDADA"
    ]
    cols_puesto = [c for c in cols_puesto if c in puestos_f.columns]
    st.dataframe(
        puestos_f[cols_puesto].sort_values(["PRIORIDAD", "VOTOS_2026"], ascending=[True, False]),
        use_container_width=True,
        hide_index=True,
    )

with tab_barrio:
    st.subheader("Barrio / UPZ")
    if resumen_barrio.empty:
        st.info("No hay resumen por barrio o UPZ disponible.")
    else:
        st.dataframe(resumen_barrio, use_container_width=True, hide_index=True)

    if "UPZ" in puestos.columns and puestos["UPZ"].notna().any() and puestos["UPZ"].astype(str).str.len().gt(0).any():
        upz_summary = puestos.groupby("UPZ", dropna=False).agg(
            VOTOS_2026=("VOTOS_2026", "sum"),
            VOTOS_2023=("VOTOS_2023", "sum"),
            PUESTOS=("PUESTO", "nunique"),
        ).reset_index()
        upz_summary["VARIACION_ABSOLUTA"] = upz_summary["VOTOS_2026"] - upz_summary["VOTOS_2023"]
        fig = px.bar(
            upz_summary.sort_values("VOTOS_2026"),
            x="VOTOS_2026",
            y="UPZ",
            orientation="h",
            title="Votos 2026 por UPZ",
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.markdown(
            '<div class="warning-box">La base consolidada aún no tiene UPZ asignada. Para activar análisis espacial por UPZ, agregue <b>data/upz_kennedy.geojson</b> o complete la columna UPZ en el Excel consolidado.</div>',
            unsafe_allow_html=True,
        )

with tab_prioridad:
    st.subheader("Matriz de priorización territorial")
    st.dataframe(matriz, use_container_width=True, hide_index=True)

    prioridad_count = matriz.groupby("NIVEL_PRIORIDAD").size().reset_index(name="PUESTOS")
    fig = px.bar(
        prioridad_count,
        x="NIVEL_PRIORIDAD",
        y="PUESTOS",
        title="Distribución de puestos por prioridad",
        color="NIVEL_PRIORIDAD",
    )
    fig.update_layout(height=380, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

with tab_export:
    st.subheader("Exportables")
    st.markdown("Descargue la base maestra consolidada o tablas específicas para anexos del informe.")

    if CONSOLIDADO.exists():
        st.download_button(
            "Descargar Excel maestro consolidado",
            data=download_excel_link(),
            file_name="kennedy_mira_consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    exportables = {
        "puestos_votacion.csv": puestos,
        "resumen_iglesia.csv": resumen_iglesia,
        "matriz_priorizacion.csv": matriz,
        "actividades_campana.csv": actividades,
        "mesas_trabajo.csv": mesas,
        "control_calidad.csv": control,
    }

    for fname, df in exportables.items():
        st.download_button(
            f"Descargar {fname}",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name=fname,
            mime="text/csv",
        )

    with st.expander("Control de calidad de datos"):
        st.dataframe(control, use_container_width=True, hide_index=True)
