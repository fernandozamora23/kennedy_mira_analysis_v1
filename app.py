
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
import hmac
import html
from io import BytesIO
import json
import math

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import folium
from folium.plugins import HeatMap, Fullscreen, MiniMap
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
LOCALIDADES_GEOJSON = DATA_DIR / "localidades_bogota.geojson"
KENNEDY_CENTER = [4.6260, -74.1570]

COLOR_TEXT = "#0F172A"
COLOR_MUTED = "#334155"
COLOR_BORDER = "#E2E8F0"
COLOR_RED = "#DC2626"
COLOR_GREEN = "#16A34A"
COLOR_BLUE = "#2563EB"
COLOR_ORANGE = "#F97316"
COLOR_PURPLE = "#7C3AED"
TEMPLOS_OFICIALES = ["CLASS ROMA", "KENNEDY CENTRAL", "PATIO BONITO", "CARVAJAL", "VALLADOLID"]
COLORES_TEMPLOS = {
    "CLASS ROMA": "#7C3AED",
    "KENNEDY CENTRAL": "#2563EB",
    "PATIO BONITO": "#16A34A",
    "CARVAJAL": "#F97316",
    "VALLADOLID": "#DC2626",
}
DIST_COLS_TEMPLOS = {
    "CLASS ROMA": "DIST_CLASS_ROMA_KM",
    "KENNEDY CENTRAL": "DIST_KENNEDY_CENTRAL_KM",
    "PATIO BONITO": "DIST_PATIO_BONITO_KM",
    "CARVAJAL": "DIST_CARVAJAL_KM",
    "VALLADOLID": "DIST_VALLADOLID_KM",
}


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
        border-radius: 12px;
        padding: 0.85rem 1rem;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        min-height: 96px;
    }

    .metric-label {
        color: #475569;
        font-size: 0.82rem;
        font-weight: 700;
        margin-bottom: 0.35rem;
    }

    .metric-value {
        color: #0F172A;
        font-size: 1.62rem;
        font-weight: 850;
        line-height: 1.1;
    }

    .summary-ribbon {
        background: linear-gradient(90deg, #FFFFFF 0%, #F8FAFC 100%);
        border: 1px solid #E2E8F0;
        border-left: 5px solid #2563EB;
        border-radius: 16px;
        padding: 1rem 1.15rem;
        margin: 0.8rem 0 1.1rem 0;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.045);
    }

    .summary-ribbon b {
        color: #0F172A;
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
        "asignacion": read("asignacion_puestos"),
        "resumen_asignacion": read("resumen_asignacion_templos"),
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
    number = Decimal(str(value)).quantize(Decimal("1") if decimals == 0 else Decimal(f"1.{'0' * decimals}"), rounding=ROUND_HALF_UP)
    return f"{number:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(value):
    if pd.isna(value):
        return "N.D."
    return f"{value:.1%}".replace(".", ",")


def safe_html(value, default="Sin dato"):
    if value is None or pd.isna(value):
        return default
    text = str(value).strip()
    return html.escape(text) if text else default


def cargar_geojson(path: Path):
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def geojson_bounds(geojson_obj):
    coords = []

    def collect(value):
        if isinstance(value, list) and len(value) >= 2 and all(isinstance(x, (int, float)) for x in value[:2]):
            lon, lat = value[:2]
            if -75 < lon < -73 and 3 < lat < 6:
                coords.append((lat, lon))
        elif isinstance(value, list):
            for item in value:
                collect(item)

    for feature in (geojson_obj or {}).get("features", []):
        collect(feature.get("geometry", {}).get("coordinates", []))

    if not coords:
        return None
    lats = [c[0] for c in coords]
    lons = [c[1] for c in coords]
    return [[min(lats), min(lons)], [max(lats), max(lons)]]


def haversine_km(lat1, lon1, lat2, lon2):
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
        return np.nan
    radio_tierra_km = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [float(lat1), float(lon1), float(lat2), float(lon2)])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 2 * radio_tierra_km * math.asin(math.sqrt(a))


def calcular_distancias_a_templos(puestos_df, iglesias_df):
    templos = iglesias_df[iglesias_df["IGLESIA"].isin(TEMPLOS_OFICIALES)].dropna(subset=["LATITUD", "LONGITUD"]).copy()
    rows = []
    for _, r in puestos_df.iterrows():
        row = {
            "PUESTO_ID": r.get("PUESTO_ID"),
            "PUESTO": r.get("PUESTO"),
            "DIRECCION": r.get("DIRECCION"),
            "BARRIO": r.get("BARRIO"),
            "UPZ": r.get("UPZ"),
            "LATITUD": r.get("LATITUD"),
            "LONGITUD": r.get("LONGITUD"),
            "IGLESIA_ACTUAL": r.get("IGLESIA"),
            "VOTOS_2026": r.get("VOTOS_2026"),
            "VOTOS_2023": r.get("VOTOS_2023"),
            "VARIACION_ABSOLUTA": r.get("VARIACION_ABSOLUTA"),
            "VARIACION_PORCENTUAL": r.get("VARIACION_PORCENTUAL"),
            "PRIORIDAD": r.get("PRIORIDAD"),
        }
        lat = pd.to_numeric(r.get("LATITUD"), errors="coerce")
        lon = pd.to_numeric(r.get("LONGITUD"), errors="coerce")
        distancias = {}
        for _, templo in templos.iterrows():
            distancia = haversine_km(lat, lon, templo["LATITUD"], templo["LONGITUD"])
            distancias[templo["IGLESIA"]] = distancia
            row[DIST_COLS_TEMPLOS[templo["IGLESIA"]]] = distancia
        if pd.isna(lat) or pd.isna(lon):
            iglesia_actual = row.get("IGLESIA_ACTUAL")
            row["TEMPLO_MAS_CERCANO"] = "SIN COORDENADAS"
            row["DISTANCIA_MINIMA_KM"] = np.nan
            row["TEMPLO_ASIGNADO_PROPUESTO"] = iglesia_actual if iglesia_actual in TEMPLOS_OFICIALES else "PENDIENTE"
            row["OBSERVACION_ASIGNACION"] = "Sin coordenadas válidas; revisar manualmente."
        else:
            templo_cercano = min(distancias, key=lambda t: distancias[t] if pd.notna(distancias[t]) else np.inf)
            row["TEMPLO_MAS_CERCANO"] = templo_cercano
            row["DISTANCIA_MINIMA_KM"] = distancias[templo_cercano]
            row["TEMPLO_ASIGNADO_PROPUESTO"] = templo_cercano
            row["OBSERVACION_ASIGNACION"] = "Propuesta por cercanía geográfica."
        rows.append(row)
    return pd.DataFrame(rows)


def aplicar_ajustes_asignacion(asignacion_df):
    df = asignacion_df.copy()
    ajustes = st.session_state.get("ajustes_asignacion", {})
    df["TEMPLO_ASIGNADO_FINAL"] = df["PUESTO"].map(ajustes).fillna(df["TEMPLO_ASIGNADO_PROPUESTO"])
    df["DISTANCIA_ASIGNADA_KM"] = df.apply(
        lambda r: r.get(DIST_COLS_TEMPLOS.get(r.get("TEMPLO_ASIGNADO_FINAL"), ""), np.nan),
        axis=1,
    )
    return df


def crear_resumen_asignacion(asignacion_df):
    rows = []
    for templo in TEMPLOS_OFICIALES:
        sub = asignacion_df[asignacion_df["TEMPLO_ASIGNADO_FINAL"].eq(templo)].copy()
        v26 = pd.to_numeric(sub.get("VOTOS_2026", pd.Series(dtype=float)), errors="coerce").sum()
        v23 = pd.to_numeric(sub.get("VOTOS_2023", pd.Series(dtype=float)), errors="coerce").sum()
        var_abs = v26 - v23
        rows.append({
            "TEMPLO": templo,
            "PUESTOS_ASIGNADOS": int(len(sub)),
            "VOTOS_2026_ASIGNADOS": int(round(v26)),
            "VOTOS_2023_ASIGNADOS": int(round(v23)),
            "VARIACION_ABSOLUTA": int(round(var_abs)),
            "VARIACION_PORCENTUAL": var_abs / v23 if v23 else np.nan,
            "DISTANCIA_PROMEDIO_KM": pd.to_numeric(sub.get("DISTANCIA_ASIGNADA_KM", pd.Series(dtype=float)), errors="coerce").mean(),
            "DISTANCIA_MAXIMA_KM": pd.to_numeric(sub.get("DISTANCIA_ASIGNADA_KM", pd.Series(dtype=float)), errors="coerce").max(),
            "PUESTOS_PRIORIDAD_ALTA": int(sub.get("PRIORIDAD", pd.Series(dtype=str)).astype(str).str.upper().eq("ALTA").sum()),
            "BARRIOS_CUBIERTOS": int(sub.get("BARRIO", pd.Series(dtype=str)).replace("", np.nan).dropna().nunique()),
            "UPZ_CUBIERTAS": int(sub.get("UPZ", pd.Series(dtype=str)).replace("", np.nan).dropna().nunique()),
        })
    return pd.DataFrame(rows)


def crear_tabla_puestos_por_templo(asignacion_df):
    grupos = {
        templo: asignacion_df[asignacion_df["TEMPLO_ASIGNADO_FINAL"].eq(templo)]["PUESTO"].sort_values().tolist()
        for templo in TEMPLOS_OFICIALES
    }
    max_len = max([len(v) for v in grupos.values()] + [0])
    return pd.DataFrame({templo: valores + [""] * (max_len - len(valores)) for templo, valores in grupos.items()})


def crear_mapa_asignacion(asignacion_df, iglesias_df):
    m = folium.Map(location=KENNEDY_CENTER, zoom_start=13, tiles="CartoDB positron", control_scale=True)
    Fullscreen(position="topleft").add_to(m)
    MiniMap(toggle_display=True, position="bottomleft").add_to(m)

    templos = iglesias_df[iglesias_df["IGLESIA"].isin(TEMPLOS_OFICIALES)].dropna(subset=["LATITUD", "LONGITUD"]).copy()
    templo_coords = {r["IGLESIA"]: (r["LATITUD"], r["LONGITUD"]) for _, r in templos.iterrows()}
    templos_layer = folium.FeatureGroup(name="Templos oficiales", show=True)
    for _, r in templos.iterrows():
        color = COLORES_TEMPLOS.get(r["IGLESIA"], "#334155")
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            tooltip=f"Templo: {r['IGLESIA']}",
            popup=folium.Popup(f"<b>{safe_html(r['IGLESIA'])}</b><br>Lat: {r['LATITUD']}<br>Lon: {r['LONGITUD']}", max_width=260),
            icon=folium.Icon(color="purple", icon="home", prefix="fa"),
        ).add_to(templos_layer)
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            icon=folium.DivIcon(html=f'<div style="background:white;border:1px solid {color};border-radius:8px;padding:3px 7px;color:{color};font-size:11px;font-weight:800;white-space:nowrap;">{safe_html(r["IGLESIA"])}</div>'),
        ).add_to(templos_layer)
    templos_layer.add_to(m)

    puestos_layer = folium.FeatureGroup(name="Puestos por templo asignado", show=True)
    lineas_layer = folium.FeatureGroup(name="Líneas puesto-templo", show=True)
    for _, r in asignacion_df.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        templo = r.get("TEMPLO_ASIGNADO_FINAL")
        color = COLORES_TEMPLOS.get(templo, "#64748B")
        distancia = pd.to_numeric(r.get("DISTANCIA_ASIGNADA_KM"), errors="coerce")
        popup = f"""
        <div style="font-family:Arial; width:330px;">
        <h4 style="margin-bottom:6px;">{safe_html(r.get('PUESTO'))}</h4>
        <b>Dirección:</b> {safe_html(r.get('DIRECCION'))}<br>
        <b>Barrio:</b> {safe_html(r.get('BARRIO'))}<br>
        <b>UPZ:</b> {safe_html(r.get('UPZ'))}<br>
        <b>Iglesia actual:</b> {safe_html(r.get('IGLESIA_ACTUAL'))}<br>
        <b>Templo más cercano:</b> {safe_html(r.get('TEMPLO_MAS_CERCANO'))}<br>
        <b>Templo asignado:</b> {safe_html(templo)}<br>
        <b>Distancia mínima:</b> {fmt_number(distancia, 2)} km<br>
        <b>Votos 2026:</b> {fmt_number(r.get('VOTOS_2026'), 0)}<br>
        <b>Prioridad:</b> {safe_html(r.get('PRIORIDAD'))}
        </div>
        """
        folium.CircleMarker(
            location=[r["LATITUD"], r["LONGITUD"]],
            radius=6,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.78,
            weight=1.3,
            tooltip=f"{r.get('PUESTO')} | {safe_html(r.get('BARRIO'))} | {safe_html(templo)} | {fmt_number(distancia, 2)} km | {fmt_number(r.get('VOTOS_2026'), 0)} votos",
            popup=folium.Popup(popup, max_width=380),
        ).add_to(puestos_layer)
        if templo in templo_coords:
            folium.PolyLine(
                locations=[[r["LATITUD"], r["LONGITUD"]], list(templo_coords[templo])],
                color=color,
                weight=1,
                opacity=0.28,
            ).add_to(lineas_layer)
    lineas_layer.add_to(m)
    puestos_layer.add_to(m)

    legend_items = "".join(
        f'<div><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:{color};margin-right:6px;"></span>{templo}</div>'
        for templo, color in COLORES_TEMPLOS.items()
    )
    m.get_root().html.add_child(folium.Element(f"""
    <div style="position: fixed; bottom: 35px; right: 35px; z-index:9999; background:white; padding:12px 14px; border:1px solid #CBD5E1; border-radius:10px; box-shadow:0 3px 12px rgba(0,0,0,.12); font-size:13px;">
    <b>Asignación territorial</b>{legend_items}
    </div>
    """))
    folium.LayerControl(collapsed=False).add_to(m)
    return m


def exportar_asignacion_excel(asignacion_df, resumen_df, tabla_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        asignacion_df.to_excel(writer, sheet_name="asignacion_detallada", index=False)
        resumen_df.to_excel(writer, sheet_name="resumen_por_templo", index=False)
        tabla_df.to_excel(writer, sheet_name="tabla_puestos_por_templo", index=False)
    return output.getvalue()


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

    localidades_gj = cargar_geojson(LOCALIDADES_GEOJSON)
    if localidades_gj:
        def style_localidad(feature):
            nombre = str((feature.get("properties") or {}).get("LocNombre", "")).strip().upper()
            is_kennedy = nombre == "KENNEDY"
            return {
                "fillColor": "#F8FAFC" if is_kennedy else "#94A3B8",
                "color": "#0F172A" if is_kennedy else "#CBD5E1",
                "weight": 2.0 if is_kennedy else 0.7,
                "fillOpacity": 0.02 if is_kennedy else 0.14,
            }

        folium.GeoJson(
            localidades_gj,
            name="Contexto localidades Bogotá",
            style_function=style_localidad,
            tooltip=folium.GeoJsonTooltip(fields=["LocNombre"], aliases=["Localidad:"], sticky=True),
            show=False,
        ).add_to(m)

    upz_gj = cargar_geojson(UPZ_GEOJSON)
    if upz_gj:
        try:
            upz_fields = [f for f in ["UPZ", "NOMBRE", "CODIGO", "AREA_HA"] if any(f in (x.get("properties") or {}) for x in upz_gj.get("features", []))]
            folium.GeoJson(
                upz_gj,
                name="UPZ Kennedy",
                style_function=lambda feature: {
                    "fillColor": "#DBEAFE",
                    "color": "#2563EB",
                    "weight": 1.1,
                    "fillOpacity": 0.13,
                },
                tooltip=folium.GeoJsonTooltip(
                    fields=upz_fields,
                    aliases=["UPZ:", "Nombre:", "Código:", "Área ha:"][: len(upz_fields)],
                    sticky=True,
                ) if upz_fields else None,
            ).add_to(m)
            bounds = geojson_bounds(upz_gj)
            if bounds:
                m.fit_bounds(bounds)
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
    puestos_layer = folium.FeatureGroup(name="Puestos de votación fijos", show=True)
    for _, r in puestos.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        var = r.get("VARIACION_ABSOLUTA", np.nan)
        color = "green" if pd.notna(var) and var > 0 else "red" if pd.notna(var) and var < 0 else "gray"
        puesto = safe_html(r.get("PUESTO", ""))
        iglesia = safe_html(r.get("IGLESIA", ""))
        barrio = safe_html(r.get("BARRIO", ""))
        upz = safe_html(r.get("UPZ", ""))
        accion = safe_html(r.get("ACCION_RECOMENDADA", ""))
        popup = f"""
        <div style="font-family:Arial; width:330px;">
        <h4 style="margin-bottom:6px;">{puesto}</h4>
        <b>Iglesia:</b> {iglesia}<br>
        <b>Barrio:</b> {barrio}<br>
        <b>UPZ:</b> {upz}<br>
        <b>Votos 2026:</b> {fmt_number(r.get('VOTOS_2026'),0)}<br>
        <b>Votos 2023:</b> {fmt_number(r.get('VOTOS_2023'),0)}<br>
        <b>Variación:</b> {fmt_number(r.get('VARIACION_ABSOLUTA'),0)} ({fmt_pct(r.get('VARIACION_PORCENTUAL'))})<br>
        <b>Actividades:</b> {fmt_number(r.get('ACTIVIDADES_CAMPANA'),0)}<br>
        <b>Mesas:</b> {fmt_number(r.get('MESAS_TRABAJO_BARRIO'),0)}<br>
        <b>Prioridad:</b> {safe_html(r.get('PRIORIDAD',''))}<br>
        <b>Acción:</b> {accion}<br>
        </div>
        """
        radius = max(5, min(17, float(r.get("VOTOS_2026", 0) or 0) / 14))
        folium.CircleMarker(
            location=[r["LATITUD"], r["LONGITUD"]],
            radius=radius,
            popup=folium.Popup(popup, max_width=380),
            tooltip=f"{r.get('PUESTO','')} | {r.get('IGLESIA','')} | {fmt_number(r.get('VOTOS_2026'),0)}",
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.72,
            weight=1.2,
        ).add_to(puestos_layer)
    puestos_layer.add_to(m)

    # Churches
    iglesia_layer = folium.FeatureGroup(name="Iglesias / templos", show=True)
    for _, r in iglesias.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        url = r.get("URL", "")
        link = f'<br><a href="{url}" target="_blank">Ver ubicación IDMJI</a>' if isinstance(url, str) and url else ""
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            popup=folium.Popup(
                f"<b>{safe_html(r.get('IGLESIA',''))}</b><br>Lat: {r['LATITUD']}<br>Lon: {r['LONGITUD']}{link}",
                max_width=280,
            ),
            tooltip=f"Iglesia: {r.get('IGLESIA','')}",
            icon=folium.Icon(color="purple", icon="home", prefix="fa"),
        ).add_to(iglesia_layer)
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            icon=folium.DivIcon(
                html=f"""
                <div style="transform:translate(18px,-8px);background:#FFFFFF;border:1px solid #DDD6FE;
                border-radius:8px;padding:3px 7px;color:#4C1D95;font-size:11px;font-weight:800;
                box-shadow:0 1px 5px rgba(15,23,42,.18);white-space:nowrap;">
                {safe_html(r.get('IGLESIA',''))}
                </div>
                """
            ),
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
                f"""
                <div style="font-family:Arial; width:280px;">
                <b>Actividad:</b> {safe_html(r.get('TIPO_ACTIVIDAD',''))}<br>
                <b>Iglesia:</b> {safe_html(r.get('IGLESIA',''))}<br>
                <b>Barrio:</b> {safe_html(r.get('BARRIO',''))}<br>
                <b>Líder:</b> {safe_html(r.get('LIDER',''))}<br>
                <b>Dirección:</b> {safe_html(r.get('DIRECCION',''))}<br>
                <b>Observaciones:</b> {safe_html(r.get('OBSERVACIONES',''))}
                </div>
                """,
                max_width=340,
            ),
        ).add_to(acts_layer)
    acts_layer.add_to(m)

    # Mesas
    mesas_layer = folium.FeatureGroup(name="Mesas de trabajo", show=True)
    for _, r in mesas.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            tooltip=f"Mesa | {r.get('IGLESIA','')} | {r.get('BARRIO','')}",
            popup=folium.Popup(
                f"""
                <div style="font-family:Arial; width:300px;">
                <b>Mesa:</b> {safe_html(r.get('TEMA',''))}<br>
                <b>Iglesia:</b> {safe_html(r.get('IGLESIA',''))}<br>
                <b>Barrio:</b> {safe_html(r.get('BARRIO',''))}<br>
                <b>Líder:</b> {safe_html(r.get('LIDER',''))}<br>
                <b>Concejal:</b> {safe_html(r.get('CONCEJAL',''))}<br>
                <b>Beneficiarios:</b> {fmt_number(r.get('BENEFICIARIOS'),0)}<br>
                <b>Compromisos:</b> {fmt_number(r.get('COMPROMISOS_TOTAL'),0)}<br>
                <b>Estado:</b> {safe_html(r.get('ESTADO',''))}<br>
                <b>Dirección:</b> {safe_html(r.get('DIRECCION',''))}<br>
                <b>Observaciones:</b> {safe_html(r.get('OBSERVACIONES',''))}
                </div>
                """,
                max_width=360,
            ),
            icon=folium.Icon(color="orange", icon="info-sign"),
        ).add_to(mesas_layer)
    mesas_layer.add_to(m)

    upz_legend = '<span style="color:#2563EB;">■</span> UPZ Kennedy<br>' if upz_gj else ""
    legend_html = f"""
    <div style="position: fixed; bottom: 35px; right: 35px; z-index:9999; background:white; padding:12px 14px; border:1px solid #CBD5E1; border-radius:10px; box-shadow:0 3px 12px rgba(0,0,0,.12); font-size:13px;">
    <b>Lectura del mapa</b><br>
    <span style="color:green;">●</span> Puesto con crecimiento<br>
    <span style="color:red;">●</span> Puesto con caída<br>
    <span style="color:gray;">●</span> Sin comparación<br>
    <span style="color:purple;">⬟</span> Iglesia / templo<br>
    <span style="color:blue;">●</span> Actividad de campaña<br>
    <span style="color:orange;">⬟</span> Mesa de trabajo<br>
    <span style="color:#94A3B8;">■</span> Otras localidades<br>
    {upz_legend}
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
asignacion = data.get("asignacion", pd.DataFrame())
resumen_asignacion = data.get("resumen_asignacion", pd.DataFrame())

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
if asignacion.empty:
    asignacion = calcular_distancias_a_templos(puestos, iglesias)

# Ensure numerics
for df in [puestos, resumen_iglesia, resumen_puesto, resumen_barrio, matriz, asignacion, resumen_asignacion]:
    if not df.empty:
        for col in ["VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "LATITUD", "LONGITUD", "DISTANCIA_MINIMA_KM"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        for col in DIST_COLS_TEMPLOS.values():
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        for col in [
            "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "JAL_2023", "MIRA_CONCEJO_2023",
            "CAMARA_2026", "SENADO_2026", "CENSO_2023", "BENEFICIARIOS", "MESAS_2026_REPORTE",
            "TESTIGOS_2023_REPORTE", "VOTOS_AFINIDAD_E11_2023", "VOTOS_MIRA_2023_PROP_LISTA",
            "ACTIVIDADES_CAMPANA", "ACTIVIDADES_CAMPANA_IGLESIA", "MESAS_TRABAJO_BARRIO",
            "MESAS_TRABAJO", "PUESTOS",
        ]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").round().astype("Int64")


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("Configuración del análisis")
    st.markdown("Fuente única: `kennedy_mira_consolidado.xlsx`")

    iglesias_oficiales = IGLESIAS_OFICIALES_PERMITIDAS
    default_iglesias = iglesias_oficiales
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
        st.success("Capa UPZ Kennedy detectada.")
    if LOCALIDADES_GEOJSON.exists():
        st.success("Capa de localidades Bogotá detectada.")


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

if "FUENTE" in actividades_f.columns:
    actividades_oficiales_total = int(actividades_f["FUENTE"].eq("AGENDA GENERAL CON CANDIDATOS").sum())
else:
    actividades_oficiales_total = get_indicador(resumen_general, "Actividades oficiales agenda general")

if "TIPO_ACTIVIDAD" in actividades_f.columns:
    volanteos_total = int(actividades_f["TIPO_ACTIVIDAD"].eq("VOLANTEO").sum())
else:
    volanteos_total = get_indicador(resumen_general, "Volanteos confirmados Kennedy")

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
    metric_card("Total general Kennedy 2026", fmt_number(total_2026, 0))
with c2:
    metric_card("Total general Kennedy 2023", fmt_number(total_2023, 0))
with c3:
    metric_card("Variación electoral", fmt_number(var_abs, 0), fmt_number(abs(var_abs), 0), positive=var_abs >= 0)
with c4:
    metric_card("Variación porcentual", fmt_pct(var_pct), fmt_pct(abs(var_pct)), positive=var_pct >= 0)

c5, c6, c7, c8 = st.columns(4)
with c5:
    metric_card("Puestos analizados", fmt_number(puestos_total, 0))
with c6:
    metric_card("Actividades oficiales", fmt_number(actividades_oficiales_total, 0))
with c7:
    metric_card("Volanteos confirmados", fmt_number(volanteos_total, 0))
with c8:
    metric_card("Mesas de trabajo", fmt_number(mesas_total, 0))

c9, c10, c11, c12, c13 = st.columns(5)
with c9:
    metric_card("Iglesias oficiales", fmt_number(iglesias_total, 0))
with c10:
    metric_card("JAL 2023", fmt_number(jal_total, 0))
with c11:
    metric_card("Concejo 2023", fmt_number(concejo_total, 0))
with c12:
    metric_card("Cámara 2026", fmt_number(camara_total, 0))
with c13:
    metric_card("Senado 2026", fmt_number(senado_total, 0))

st.markdown(
    f"""
    <div class="summary-ribbon">
    <b>Lectura ejecutiva:</b> el dashboard separa la comparación <b>JAL / Concejo 2023</b> frente a
    <b>Cámara / Senado 2026</b>, mantiene solo las cinco iglesias oficiales y cruza puestos fijos,
    <b>{fmt_number(actividades_oficiales_total, 0)} actividades oficiales</b>,
    <b>{fmt_number(volanteos_total, 0)} volanteos confirmados</b> y mesas de trabajo con prioridad territorial.
    La capa geográfica usa localidades de Bogotá como contexto y UPZ Kennedy solo cuando exista una capa oficial.
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# TABS
# ============================================================

tab_resumen, tab_mapa, tab_asignacion, tab_iglesia, tab_puesto, tab_barrio, tab_prioridad, tab_export = st.tabs(
    [
        "Resumen ejecutivo",
        "Mapa territorial",
        "Asignación de puestos",
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
            text_auto=".0f",
            title="Comparación separada por corporación"
        )
        fig_elec.update_layout(height=350, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
        st.plotly_chart(fig_elec, use_container_width=True)

with tab_mapa:
    st.subheader("Mapa interactivo territorial")
    st.markdown(
        '<div class="note-box">Los puestos de votación se muestran como puntos fijos para lectura territorial estable. Active o desactive capas para comparar iglesias, mesas, actividades, calor electoral y variación.</div>',
        unsafe_allow_html=True,
    )
    mapa = crear_mapa(puestos_f, iglesias, actividades_f, mesas_f)
    st_folium(mapa, width=None, height=720)

with tab_asignacion:
    st.subheader("Asignación territorial de puestos")
    st.markdown(
        """
        <div class="section-card">
        Esta sección propone una distribución preliminar de los 123 puestos de votación entre los 5 templos,
        usando como criterio base la cercanía geográfica. La asignación puede ajustarse manualmente para
        discusión política, logística y territorial.
        </div>
        """,
        unsafe_allow_html=True,
    )

    asignacion_base = asignacion.copy()
    asignacion_final = aplicar_ajustes_asignacion(asignacion_base)
    resumen_final = crear_resumen_asignacion(asignacion_final)
    tabla_templos = crear_tabla_puestos_por_templo(asignacion_final)

    puestos_con_coord = asignacion_final.dropna(subset=["LATITUD", "LONGITUD"]).shape[0]
    puestos_sin_coord = len(asignacion_final) - puestos_con_coord
    templo_mayor = resumen_final.sort_values("PUESTOS_ASIGNADOS", ascending=False).iloc[0]
    distancia_prom = pd.to_numeric(asignacion_final["DISTANCIA_ASIGNADA_KM"], errors="coerce").mean()
    distancia_max = pd.to_numeric(asignacion_final["DISTANCIA_ASIGNADA_KM"], errors="coerce").max()

    a1, a2, a3, a4, a5, a6 = st.columns(6)
    with a1:
        metric_card("Total puestos", fmt_number(len(asignacion_final), 0))
    with a2:
        metric_card("Con coordenadas", fmt_number(puestos_con_coord, 0))
    with a3:
        metric_card("Sin coordenadas", fmt_number(puestos_sin_coord, 0))
    with a4:
        metric_card("Mayor carga", safe_html(templo_mayor["TEMPLO"]))
    with a5:
        metric_card("Distancia promedio", f"{fmt_number(distancia_prom, 2)} km")
    with a6:
        metric_card("Distancia máxima", f"{fmt_number(distancia_max, 2)} km")

    mapa_asignacion = crear_mapa_asignacion(asignacion_final, iglesias)
    st_folium(mapa_asignacion, width=None, height=720)

    st.markdown("### Ajuste temporal de asignación")
    if "ajustes_asignacion" not in st.session_state:
        st.session_state["ajustes_asignacion"] = {}

    lista_puestos = asignacion_final["PUESTO"].dropna().sort_values().tolist()
    puesto_sel = st.selectbox("Selecciona un puesto de votación", lista_puestos)
    puesto_row = asignacion_final[asignacion_final["PUESTO"].eq(puesto_sel)].iloc[0]
    templo_actual = puesto_row.get("TEMPLO_ASIGNADO_FINAL")
    index_templo = TEMPLOS_OFICIALES.index(templo_actual) if templo_actual in TEMPLOS_OFICIALES else 0

    info_cols = [
        ("Barrio", puesto_row.get("BARRIO")),
        ("Dirección", puesto_row.get("DIRECCION")),
        ("UPZ", puesto_row.get("UPZ")),
        ("Iglesia actual", puesto_row.get("IGLESIA_ACTUAL")),
        ("Templo más cercano", puesto_row.get("TEMPLO_MAS_CERCANO")),
        ("Votos 2026", fmt_number(puesto_row.get("VOTOS_2026"), 0)),
        ("Prioridad", puesto_row.get("PRIORIDAD")),
    ]
    st.dataframe(pd.DataFrame(info_cols, columns=["Campo", "Valor"]), hide_index=True, use_container_width=True)

    dist_info = pd.DataFrame(
        [{"Templo": templo, "Distancia km": puesto_row.get(col)} for templo, col in DIST_COLS_TEMPLOS.items()]
    )
    st.dataframe(dist_info, hide_index=True, use_container_width=True)

    templo_nuevo = st.selectbox("Asignar este puesto al templo", TEMPLOS_OFICIALES, index=index_templo)
    col_guardar, col_limpiar = st.columns([1, 3])
    with col_guardar:
        if st.button("Guardar ajuste temporal"):
            st.session_state["ajustes_asignacion"][puesto_sel] = templo_nuevo
            st.success(f"Ajuste temporal guardado: {puesto_sel} → {templo_nuevo}")
            st.rerun()
    with col_limpiar:
        if st.button("Limpiar ajustes temporales"):
            st.session_state["ajustes_asignacion"] = {}
            st.rerun()

    st.markdown("### Resumen por templo")
    st.dataframe(resumen_final, hide_index=True, use_container_width=True)

    st.markdown("### Puestos asignados por templo")
    st.dataframe(tabla_templos, hide_index=True, use_container_width=True)

    with st.expander("Lectura automática de la asignación", expanded=True):
        templo_menor = resumen_final.sort_values("PUESTOS_ASIGNADOS", ascending=True).iloc[0]
        puestos_lejanos = int(pd.to_numeric(asignacion_final["DISTANCIA_ASIGNADA_KM"], errors="coerce").gt(3).sum())
        alta_por_templo = resumen_final.sort_values("PUESTOS_PRIORIDAD_ALTA", ascending=False).iloc[0]
        st.write(f"El templo con mayor carga territorial propuesta es {templo_mayor['TEMPLO']}, con {fmt_number(templo_mayor['PUESTOS_ASIGNADOS'], 0)} puestos asignados.")
        st.write(f"El templo con menor carga territorial propuesta es {templo_menor['TEMPLO']}, con {fmt_number(templo_menor['PUESTOS_ASIGNADOS'], 0)} puestos asignados.")
        st.write(f"La distancia promedio entre puestos y templos asignados es de {fmt_number(distancia_prom, 2)} km.")
        st.write(f"Hay {fmt_number(puestos_lejanos, 0)} puestos a más de 3 km del templo asignado, que requieren revisión logística.")
        st.write(f"Hay {fmt_number(alta_por_templo['PUESTOS_PRIORIDAD_ALTA'], 0)} puestos de prioridad alta asignados a {alta_por_templo['TEMPLO']}.")

    with st.expander("Advertencias metodológicas"):
        st.write(
            "La asignación automática se basa únicamente en distancia geográfica entre el puesto de votación y el templo. "
            "Esta propuesta debe ser revisada con criterios adicionales: liderazgo comunitario, histórico de votación, "
            "capacidad operativa del templo, rutas de transporte, UPZ, barrios priorizados, presencia de mesas de trabajo "
            "y conocimiento territorial de los equipos."
        )

    csv_asignacion = asignacion_final.to_csv(index=False).encode("utf-8-sig")
    excel_asignacion = exportar_asignacion_excel(asignacion_final, resumen_final, tabla_templos)
    d1, d2 = st.columns(2)
    with d1:
        st.download_button("Descargar CSV asignación final", csv_asignacion, "asignacion_puestos_final.csv", "text/csv")
    with d2:
        st.download_button(
            "Descargar Excel asignación final",
            excel_asignacion,
            "asignacion_puestos_final.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

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
