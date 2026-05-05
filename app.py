from io import BytesIO
from pathlib import Path

import folium
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from folium.plugins import Fullscreen, HeatMap, MarkerCluster, MiniMap
from streamlit_folium import st_folium

try:
    import geopandas as gpd
except ImportError:
    gpd = None


st.set_page_config(page_title="Dashboard territorial-electoral Kennedy", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
ARCHIVO_CONSOLIDADO = BASE_DIR / "data" / "kennedy_mira_consolidado.xlsx"
ARCHIVO_UPZ = BASE_DIR / "data" / "upz_kennedy.geojson"
IGLESIAS_OFICIALES = ["CLASS ROMA", "KENNEDY CENTRAL", "PATIO BONITO", "CARVAJAL", "VALLADOLID"]

COLORES = {
    "fondo": "#F8FAFC",
    "tarjeta": "#FFFFFF",
    "texto": "#0F172A",
    "secundario": "#334155",
    "borde": "#E2E8F0",
    "rojo": "#DC2626",
    "verde": "#16A34A",
    "azul": "#2563EB",
    "naranja": "#F97316",
    "morado": "#7C3AED",
    "gris": "#64748B",
}

st.markdown(
    f"""
    <style>
    .stApp {{ background: {COLORES['fondo']}; color: {COLORES['texto']}; }}
    .block-container {{ padding-top: 1.7rem; padding-bottom: 2.5rem; }}
    h1, h2, h3, h4, p, label, span {{ color: {COLORES['texto']}; }}
    div[data-testid="stMetric"] {{
        background: {COLORES['tarjeta']};
        border: 1px solid {COLORES['borde']};
        border-radius: 10px;
        padding: 15px 16px;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06);
    }}
    div[data-testid="stMetricLabel"] p {{ color: {COLORES['secundario']}; font-weight: 650; }}
    div[data-testid="stMetricValue"] {{ color: {COLORES['texto']}; }}
    .hero {{
        background: {COLORES['tarjeta']};
        border: 1px solid {COLORES['borde']};
        border-radius: 12px;
        padding: 22px 24px;
        margin-bottom: 18px;
    }}
    .hero-title {{ font-size: 2rem; font-weight: 780; color: {COLORES['texto']}; margin-bottom: 4px; }}
    .hero-subtitle {{ color: {COLORES['secundario']}; font-size: 1rem; }}
    .panel {{
        background: {COLORES['tarjeta']};
        border: 1px solid {COLORES['borde']};
        border-radius: 10px;
        padding: 16px 18px;
        margin: 8px 0 14px 0;
    }}
    .muted {{ color: {COLORES['secundario']}; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def formato_numero(valor):
    if pd.isna(valor):
        return "N/D"
    return f"{valor:,.0f}".replace(",", ".")


def formato_pct(valor):
    if pd.isna(valor):
        return "N/D"
    return f"{valor:.1%}"


def variacion_color(valor):
    if pd.isna(valor):
        return COLORES["gris"]
    if valor > 0:
        return COLORES["verde"]
    if valor < 0:
        return COLORES["rojo"]
    return COLORES["gris"]


def csv_bytes(df):
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def excel_bytes(hojas):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nombre, df in hojas.items():
            df.to_excel(writer, sheet_name=nombre[:31], index=False)
    return output.getvalue()


def leer_hoja(nombre):
    return pd.read_excel(ARCHIVO_CONSOLIDADO, sheet_name=nombre, engine="openpyxl")


@st.cache_data(show_spinner=False)
def cargar_consolidado(timestamp):
    del timestamp
    hojas = {
        "puestos_votacion": leer_hoja("puestos_votacion"),
        "actividades_campana": leer_hoja("actividades_campana"),
        "mesas_trabajo": leer_hoja("mesas_trabajo"),
        "iglesias": leer_hoja("iglesias"),
        "resumen_iglesia": leer_hoja("resumen_iglesia"),
        "resumen_puesto": leer_hoja("resumen_puesto"),
        "resumen_barrio": leer_hoja("resumen_barrio"),
        "matriz_priorizacion": leer_hoja("matriz_priorizacion"),
        "informe_ejecutivo": leer_hoja("informe_ejecutivo"),
    }
    return hojas


def validar_consolidado():
    if not ARCHIVO_CONSOLIDADO.exists():
        st.error("No se encontro data/kennedy_mira_consolidado.xlsx.")
        st.info("Ejecuta `python consolidar_datos.py` para generar la base maestra antes de abrir el dashboard.")
        return False
    requeridas = [
        "puestos_votacion",
        "actividades_campana",
        "mesas_trabajo",
        "iglesias",
        "resumen_iglesia",
        "resumen_puesto",
        "resumen_barrio",
        "matriz_priorizacion",
        "informe_ejecutivo",
    ]
    try:
        hojas = pd.ExcelFile(ARCHIVO_CONSOLIDADO, engine="openpyxl").sheet_names
    except Exception as exc:
        st.error(f"No se pudo abrir el consolidado: {exc}")
        return False
    faltantes = [hoja for hoja in requeridas if hoja not in hojas]
    if faltantes:
        st.error("El consolidado no tiene todas las hojas requeridas.")
        st.warning(", ".join(faltantes))
        return False
    return True


def aplicar_filtros(puestos, actividades, mesas, matriz):
    st.sidebar.header("Filtros de analisis")
    iglesia_sel = st.sidebar.multiselect("Iglesia oficial", IGLESIAS_OFICIALES, default=IGLESIAS_OFICIALES)
    barrio_sel = st.sidebar.multiselect("Barrio", sorted(puestos["BARRIO"].dropna().astype(str).unique()))
    upz_sel = st.sidebar.multiselect("UPZ", sorted(puestos["UPZ"].dropna().astype(str).unique()))
    prioridad_sel = st.sidebar.multiselect("Prioridad", ["ALTA", "MEDIA", "BAJA"])
    puesto_sel = st.sidebar.multiselect("Puesto de votacion", sorted(puestos["PUESTO"].dropna().astype(str).unique()))

    puestos_f = puestos[puestos["IGLESIA"].isin(iglesia_sel)].copy()
    actividades_f = actividades[actividades["IGLESIA"].isin(iglesia_sel)].copy()
    mesas_f = mesas[mesas["IGLESIA"].isin(iglesia_sel)].copy()
    matriz_f = matriz[matriz["IGLESIA"].isin(iglesia_sel)].copy()

    if barrio_sel:
        puestos_f = puestos_f[puestos_f["BARRIO"].isin(barrio_sel)]
        actividades_f = actividades_f[actividades_f["BARRIO"].isin(barrio_sel)]
        mesas_f = mesas_f[mesas_f["BARRIO"].isin(barrio_sel)]
        matriz_f = matriz_f[matriz_f["BARRIO"].isin(barrio_sel)]
    if upz_sel:
        puestos_f = puestos_f[puestos_f["UPZ"].isin(upz_sel)]
        matriz_f = matriz_f[matriz_f["UPZ"].isin(upz_sel)]
    if prioridad_sel:
        puestos_f = puestos_f[puestos_f["PRIORIDAD"].isin(prioridad_sel)]
        matriz_f = matriz_f[matriz_f["NIVEL_PRIORIDAD"].isin(prioridad_sel)]
    if puesto_sel:
        puestos_f = puestos_f[puestos_f["PUESTO"].isin(puesto_sel)]
        matriz_f = matriz_f[matriz_f["PUESTO"].isin(puesto_sel)]
    return puestos_f, actividades_f, mesas_f, matriz_f


def recalcular_resumen_iglesia(puestos, actividades, mesas, resumen_base):
    base = pd.DataFrame({"IGLESIA": IGLESIAS_OFICIALES})
    if puestos.empty:
        resumen = base.copy()
        for col in ["VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "PUESTOS", "BARRIOS", "ACTIVIDADES_CAMPANA", "MESAS_TRABAJO"]:
            resumen[col] = 0
        resumen["VARIACION_PORCENTUAL"] = np.nan
        return resumen
    resumen = puestos.groupby("IGLESIA", as_index=False).agg(
        VOTOS_2026=("VOTOS_2026", "sum"),
        VOTOS_2023=("VOTOS_2023", "sum"),
        PUESTOS=("PUESTO_ID", "nunique"),
        BARRIOS=("BARRIO", "nunique"),
    )
    resumen = base.merge(resumen, on="IGLESIA", how="left").fillna({"VOTOS_2026": 0, "VOTOS_2023": 0, "PUESTOS": 0, "BARRIOS": 0})
    resumen["VARIACION_ABSOLUTA"] = resumen["VOTOS_2026"] - resumen["VOTOS_2023"]
    resumen["VARIACION_PORCENTUAL"] = np.where(resumen["VOTOS_2023"].gt(0), resumen["VARIACION_ABSOLUTA"] / resumen["VOTOS_2023"], np.nan)
    resumen["ACTIVIDADES_CAMPANA"] = resumen["IGLESIA"].map(actividades.groupby("IGLESIA").size()).fillna(0).astype(int)
    resumen["MESAS_TRABAJO"] = resumen["IGLESIA"].map(mesas.groupby("IGLESIA").size()).fillna(0).astype(int)
    cols_extra = ["PUESTO_MAYOR_VOTACION", "PUESTO_MAYOR_CAIDA", "PUESTO_MAYOR_CRECIMIENTO", "LECTURA_ESTRATEGICA", "RECOMENDACION"]
    resumen = resumen.merge(resumen_base[["IGLESIA"] + cols_extra], on="IGLESIA", how="left")
    return resumen


def recalcular_resumen_barrio(puestos, actividades, mesas):
    if puestos.empty:
        return pd.DataFrame()
    resumen = puestos.groupby(["BARRIO", "IGLESIA", "UPZ"], as_index=False).agg(
        VOTOS_2026=("VOTOS_2026", "sum"),
        VOTOS_2023=("VOTOS_2023", "sum"),
        PUESTOS=("PUESTO_ID", "nunique"),
    )
    resumen["VARIACION_ABSOLUTA"] = resumen["VOTOS_2026"] - resumen["VOTOS_2023"]
    resumen["VARIACION_PORCENTUAL"] = np.where(resumen["VOTOS_2023"].gt(0), resumen["VARIACION_ABSOLUTA"] / resumen["VOTOS_2023"], np.nan)
    resumen["ACTIVIDADES_CAMPANA"] = resumen["BARRIO"].map(actividades.groupby("BARRIO").size()).fillna(0).astype(int)
    resumen["MESAS_TRABAJO"] = resumen["BARRIO"].map(mesas.groupby("BARRIO").size()).fillna(0).astype(int)
    resumen["PRIORIDAD"] = np.where(
        (resumen["VARIACION_ABSOLUTA"] < 0) & (resumen["VOTOS_2026"] >= resumen["VOTOS_2026"].quantile(0.6)),
        "ALTA",
        np.where(resumen["VOTOS_2026"] >= resumen["VOTOS_2026"].quantile(0.35), "MEDIA", "BAJA"),
    )
    return resumen.sort_values("VOTOS_2026", ascending=False)


def agregar_leyenda(mapa):
    leyenda = f"""
    <div style="position: fixed; bottom: 36px; left: 36px; z-index: 9999; background: #FFFFFF;
        border: 1px solid {COLORES['borde']}; border-radius: 8px; padding: 12px 14px; font-size: 12px;
        color: {COLORES['texto']}; box-shadow: 0 2px 8px rgba(15, 23, 42, 0.18);">
        <b>Mapa territorial</b><br>
        <span style="color:{COLORES['verde']};">●</span> Puesto con crecimiento<br>
        <span style="color:{COLORES['rojo']};">●</span> Puesto con caida<br>
        <span style="color:{COLORES['gris']};">●</span> Sin comparacion<br>
        <span style="color:{COLORES['morado']};">◆</span> Iglesia oficial<br>
        <span style="color:{COLORES['azul']};">●</span> Actividad de campana<br>
        <span style="color:{COLORES['naranja']};">●</span> Mesa de trabajo
    </div>
    """
    mapa.get_root().html.add_child(folium.Element(leyenda))


def crear_mapa(puestos, iglesias, actividades, mesas):
    mapa = folium.Map(location=[4.628, -74.16], zoom_start=13, tiles="CartoDB positron", control_scale=True)
    Fullscreen(position="topleft").add_to(mapa)
    MiniMap(toggle_display=True, minimized=True).add_to(mapa)

    if ARCHIVO_UPZ.exists() and gpd is not None:
        try:
            upz = gpd.read_file(ARCHIVO_UPZ).to_crs("EPSG:4326")
            folium.GeoJson(
                upz,
                name="Poligonos UPZ",
                style_function=lambda _: {"fillColor": COLORES["azul"], "color": "#1E3A8A", "weight": 1, "fillOpacity": 0.08},
            ).add_to(mapa)
        except Exception as exc:
            st.warning(f"No se pudo cargar data/upz_kennedy.geojson: {exc}")

    cluster = MarkerCluster(name="Puestos de votacion").add_to(mapa)
    variacion = folium.FeatureGroup(name="Variacion electoral", show=True).add_to(mapa)
    for _, row in puestos.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        color = variacion_color(row["VARIACION_ABSOLUTA"])
        popup = f"""
        <div style="font-family: Arial; font-size: 13px; min-width: 285px; color: #0F172A;">
            <h4 style="margin:0 0 8px 0;">{row['PUESTO']}</h4>
            <b>Iglesia:</b> {row['IGLESIA']}<br>
            <b>Barrio:</b> {row['BARRIO']}<br>
            <b>UPZ:</b> {row['UPZ']}<br>
            <b>Votos 2026:</b> {formato_numero(row['VOTOS_2026'])}<br>
            <b>Votos 2023:</b> {formato_numero(row['VOTOS_2023'])}<br>
            <b>Variacion absoluta:</b> {formato_numero(row['VARIACION_ABSOLUTA'])}<br>
            <b>Variacion porcentual:</b> {formato_pct(row['VARIACION_PORCENTUAL'])}<br>
            <b>Actividades de campana:</b> {row['ACTIVIDADES_CAMPANA']}<br>
            <b>Mesas de trabajo:</b> {row['MESAS_TRABAJO_BARRIO']}<br>
            <b>Prioridad:</b> {row['PRIORIDAD']}<br>
            <b>Accion:</b> {row['ACCION_RECOMENDADA']}
        </div>
        """
        radio = max(5, min(18, float(row["VOTOS_2026"] or 0) / 18))
        folium.CircleMarker(
            location=[row["LATITUD"], row["LONGITUD"]],
            radius=radio,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.72,
            weight=2,
            tooltip=f"{row['PUESTO']} · {row['IGLESIA']}",
            popup=folium.Popup(popup, max_width=390),
        ).add_to(cluster)
        folium.CircleMarker(
            location=[row["LATITUD"], row["LONGITUD"]],
            radius=radio + 4,
            color=color,
            fill=False,
            weight=3,
            tooltip=f"Variacion electoral: {formato_numero(row['VARIACION_ABSOLUTA'])}",
        ).add_to(variacion)

    iglesias_layer = folium.FeatureGroup(name="Iglesias oficiales", show=True).add_to(mapa)
    for _, row in iglesias.iterrows():
        link = f"<br><a href='{row['URL']}' target='_blank'>Ver direccion</a>" if isinstance(row.get("URL"), str) and row["URL"] else ""
        html = f"<b>{row['IGLESIA']}</b><br>Lat: {row['LATITUD']}<br>Lon: {row['LONGITUD']}{link}"
        folium.Marker(
            location=[row["LATITUD"], row["LONGITUD"]],
            tooltip=row["IGLESIA"],
            popup=folium.Popup(html, max_width=280),
            icon=folium.Icon(color="purple", icon="star", prefix="fa"),
        ).add_to(iglesias_layer)
        folium.map.Marker(
            [row["LATITUD"], row["LONGITUD"]],
            icon=folium.DivIcon(
                html=f"""<div style="font-size:11px;font-weight:700;color:{COLORES['morado']};background:white;border:1px solid {COLORES['borde']};border-radius:4px;padding:2px 5px;white-space:nowrap;">{row['IGLESIA']}</div>"""
            ),
        ).add_to(iglesias_layer)

    actividades_layer = folium.FeatureGroup(name="Actividades de campana", show=False).add_to(mapa)
    for _, row in actividades.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        folium.CircleMarker(
            location=[row["LATITUD"], row["LONGITUD"]],
            radius=5,
            color=COLORES["azul"],
            fill=True,
            fill_color=COLORES["azul"],
            tooltip=str(row.get("TIPO_ACTIVIDAD", "Actividad")),
            popup=folium.Popup(f"<b>{row.get('TIPO_ACTIVIDAD', '')}</b><br>{row.get('IGLESIA', '')}<br>{row.get('BARRIO', '')}<br>{row.get('OBSERVACIONES', '')}", max_width=320),
        ).add_to(actividades_layer)

    mesas_layer = folium.FeatureGroup(name="Mesas de trabajo", show=False).add_to(mapa)
    for _, row in mesas.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        folium.CircleMarker(
            location=[row["LATITUD"], row["LONGITUD"]],
            radius=6,
            color=COLORES["naranja"],
            fill=True,
            fill_color=COLORES["naranja"],
            tooltip=str(row.get("TEMA", "Mesa de trabajo")),
            popup=folium.Popup(f"<b>{row.get('TEMA', '')}</b><br>{row.get('IGLESIA', '')}<br>{row.get('BARRIO', '')}<br>{row.get('ESTADO', '')}", max_width=320),
        ).add_to(mesas_layer)

    heat_data = puestos.dropna(subset=["LATITUD", "LONGITUD"])[["LATITUD", "LONGITUD", "VOTOS_2026"]].values.tolist()
    if heat_data:
        HeatMap(heat_data, name="Mapa de calor votos 2026", radius=24, blur=18, show=False).add_to(mapa)

    agregar_leyenda(mapa)
    folium.LayerControl(collapsed=False).add_to(mapa)
    return mapa


st.markdown(
    """
    <div class="hero">
      <div class="hero-title">Dashboard territorial-electoral Kennedy</div>
      <div class="hero-subtitle">Base maestra consolidada · Campaña Congreso 2026 · Partido MIRA · Variación electoral, presencia comunitaria y priorización de intervención</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if not validar_consolidado():
    st.stop()

datos = cargar_consolidado(ARCHIVO_CONSOLIDADO.stat().st_mtime)
puestos_all = datos["puestos_votacion"]
actividades_all = datos["actividades_campana"]
mesas_all = datos["mesas_trabajo"]
iglesias = datos["iglesias"]
resumen_iglesia_base = datos["resumen_iglesia"]
matriz_all = datos["matriz_priorizacion"]
informe = datos["informe_ejecutivo"]

puestos_qc = puestos_all[~puestos_all["IGLESIA"].isin(IGLESIAS_OFICIALES)].copy()
actividades_qc = actividades_all[~actividades_all["IGLESIA"].isin(IGLESIAS_OFICIALES)].copy()
mesas_qc = mesas_all[~mesas_all["IGLESIA"].isin(IGLESIAS_OFICIALES)].copy()

puestos, actividades, mesas, matriz = aplicar_filtros(puestos_all, actividades_all, mesas_all, matriz_all)
resumen_iglesia = recalcular_resumen_iglesia(puestos, actividades, mesas, resumen_iglesia_base)
resumen_barrio = recalcular_resumen_barrio(puestos, actividades, mesas)

votos_2026 = puestos["VOTOS_2026"].sum()
votos_2023 = puestos["VOTOS_2023"].sum()
var_abs = votos_2026 - votos_2023
var_pct = var_abs / votos_2023 if votos_2023 else np.nan

c1, c2, c3, c4 = st.columns(4)
c1.metric("Votos 2026", formato_numero(votos_2026))
c2.metric("Votos 2023", formato_numero(votos_2023))
c3.metric("Variacion electoral", formato_numero(var_abs), delta=formato_numero(var_abs))
c4.metric("Variacion porcentual", formato_pct(var_pct), delta=formato_pct(var_pct))
c5, c6, c7, c8 = st.columns(4)
c5.metric("Puestos analizados", formato_numero(puestos["PUESTO_ID"].nunique()))
c6.metric("Actividades de campana", formato_numero(len(actividades)))
c7.metric("Mesas de trabajo", formato_numero(len(mesas)))
c8.metric("Iglesias oficiales", formato_numero(len(IGLESIAS_OFICIALES)))

tabs = st.tabs(
    [
        "Resumen ejecutivo",
        "Mapa territorial",
        "Analisis por iglesia",
        "Analisis por puesto",
        "Barrio / UPZ",
        "Priorizacion",
        "Exportables",
    ]
)

with tabs[0]:
    st.subheader("Resumen ejecutivo")
    for _, row in informe.iterrows():
        with st.expander(row["SECCION"], expanded=row["SECCION"] in ["Resumen general", "Hallazgos principales"]):
            st.write(row["TEXTO"])

    col_a, col_b = st.columns([1.1, 1])
    with col_a:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.write("Lectura tecnica")
        st.write(
            "El analisis principal excluye registros sin clasificar para mantener consistencia institucional. "
            "La priorizacion combina concentracion territorial, variacion electoral, eficiencia territorial de campana y presencia comunitaria."
        )
        st.markdown("</div>", unsafe_allow_html=True)
        if not matriz.empty:
            st.dataframe(
                matriz.sort_values(["NIVEL_PRIORIDAD", "VOTOS_2026"], ascending=[True, False]).head(12),
                use_container_width=True,
            )
    with col_b:
        fig = px.bar(
            resumen_iglesia.sort_values("VOTOS_2026"),
            x="VOTOS_2026",
            y="IGLESIA",
            orientation="h",
            title="Concentracion territorial de votos 2026 por iglesia",
            color_discrete_sequence=[COLORES["azul"]],
        )
        fig.update_layout(height=420, font_color=COLORES["texto"], paper_bgcolor=COLORES["tarjeta"], plot_bgcolor=COLORES["tarjeta"])
        st.plotly_chart(fig, use_container_width=True)

    with st.expander("Control de calidad: registros sin clasificar"):
        st.write("Estos registros no entran al analisis principal hasta que se pueda asignar una iglesia oficial.")
        q1, q2, q3 = st.columns(3)
        q1.metric("Puestos sin clasificar", formato_numero(len(puestos_qc)))
        q2.metric("Actividades sin clasificar", formato_numero(len(actividades_qc)))
        q3.metric("Mesas sin clasificar", formato_numero(len(mesas_qc)))
        if not puestos_qc.empty:
            st.dataframe(puestos_qc[["PUESTO", "IGLESIA", "BARRIO", "UPZ", "VOTOS_2026", "VOTOS_2023"]], use_container_width=True)

with tabs[1]:
    st.subheader("Mapa territorial")
    st_folium(crear_mapa(puestos, iglesias, actividades, mesas), width=None, height=720)
    with st.expander("Como leer el mapa"):
        st.write(
            "Los puestos aparecen en verde cuando tienen crecimiento, rojo cuando presentan caida y gris cuando no hay comparacion. "
            "Las iglesias oficiales tienen icono destacado y nombre visible. Las capas permiten evaluar presencia comunitaria, actividades de campana y mapa de calor electoral."
        )

with tabs[2]:
    st.subheader("Analisis por iglesia")
    st.dataframe(resumen_iglesia.sort_values("VOTOS_2026", ascending=False), use_container_width=True)
    for iglesia in IGLESIAS_OFICIALES:
        datos_i = puestos[puestos["IGLESIA"].eq(iglesia)]
        resumen_i = resumen_iglesia[resumen_iglesia["IGLESIA"].eq(iglesia)]
        with st.expander(iglesia, expanded=not datos_i.empty):
            if datos_i.empty:
                st.write("Sin puestos oficiales asociados en la base consolidada filtrada.")
                continue
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Votos 2026", formato_numero(datos_i["VOTOS_2026"].sum()))
            k2.metric("Votos 2023", formato_numero(datos_i["VOTOS_2023"].sum()))
            k3.metric("Variacion", formato_numero(datos_i["VARIACION_ABSOLUTA"].sum()))
            k4.metric("Puestos", formato_numero(datos_i["PUESTO_ID"].nunique()))
            if not resumen_i.empty:
                st.write(resumen_i.iloc[0].get("LECTURA_ESTRATEGICA", ""))
                st.info(resumen_i.iloc[0].get("RECOMENDACION", ""))
            fig = px.bar(
                datos_i.sort_values("VARIACION_ABSOLUTA"),
                x="VARIACION_ABSOLUTA",
                y="PUESTO",
                orientation="h",
                title=f"Variacion electoral por puesto · {iglesia}",
                color="VARIACION_ABSOLUTA",
                color_continuous_scale=[COLORES["rojo"], "#E5E7EB", COLORES["verde"]],
            )
            fig.update_layout(height=430, coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(datos_i[["PUESTO", "BARRIO", "UPZ", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "PRIORIDAD", "ACCION_RECOMENDADA"]], use_container_width=True)

with tabs[3]:
    st.subheader("Analisis por puesto")
    col1, col2 = st.columns(2)
    with col1:
        top_crece = puestos.sort_values("VARIACION_ABSOLUTA", ascending=False).head(10)
        fig = px.bar(top_crece.sort_values("VARIACION_ABSOLUTA"), x="VARIACION_ABSOLUTA", y="PUESTO", orientation="h", title="Top 10 puestos con mayor crecimiento", color_discrete_sequence=[COLORES["verde"]])
        fig.update_layout(height=430)
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        top_cae = puestos.sort_values("VARIACION_ABSOLUTA", ascending=True).head(10)
        fig = px.bar(top_cae.sort_values("VARIACION_ABSOLUTA", ascending=False), x="VARIACION_ABSOLUTA", y="PUESTO", orientation="h", title="Top 10 puestos con mayor caida", color_discrete_sequence=[COLORES["rojo"]])
        fig.update_layout(height=430)
        st.plotly_chart(fig, use_container_width=True)

    fig = px.scatter(
        puestos,
        x="VOTOS_2026",
        y="VARIACION_ABSOLUTA",
        color="PRIORIDAD",
        size=np.maximum(puestos["ACTIVIDADES_CAMPANA"], 1),
        hover_name="PUESTO",
        hover_data=["IGLESIA", "BARRIO", "UPZ", "MESAS_TRABAJO_BARRIO"],
        title="Votos 2026 vs variacion absoluta",
        color_discrete_map={"ALTA": COLORES["rojo"], "MEDIA": COLORES["naranja"], "BAJA": COLORES["verde"]},
    )
    fig.update_layout(height=520)
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(puestos.sort_values("VOTOS_2026", ascending=False), use_container_width=True)

with tabs[4]:
    st.subheader("Analisis por barrio / UPZ")
    st.dataframe(resumen_barrio, use_container_width=True)
    if not resumen_barrio.empty:
        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(
                resumen_barrio.head(15).sort_values("VOTOS_2026"),
                x="VOTOS_2026",
                y="BARRIO",
                orientation="h",
                title="Concentracion territorial por barrio",
                color_discrete_sequence=[COLORES["azul"]],
            )
            fig.update_layout(height=460)
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            matriz_barrio = puestos.pivot_table(index="BARRIO", columns="IGLESIA", values="VOTOS_2026", aggfunc="sum", fill_value=0)
            if matriz_barrio.shape[0] > 1 and matriz_barrio.shape[1] > 1:
                fig = px.imshow(matriz_barrio, text_auto=True, aspect="auto", title="Matriz barrio / iglesia", color_continuous_scale="Blues")
                fig.update_layout(height=460)
                st.plotly_chart(fig, use_container_width=True)

with tabs[5]:
    st.subheader("Priorizacion de intervencion")
    st.dataframe(matriz.sort_values(["NIVEL_PRIORIDAD", "VOTOS_2026"], ascending=[True, False]), use_container_width=True)
    col1, col2 = st.columns(2)
    with col1:
        conteo = matriz["NIVEL_PRIORIDAD"].value_counts().rename_axis("NIVEL_PRIORIDAD").reset_index(name="PUESTOS")
        fig = px.bar(conteo, x="NIVEL_PRIORIDAD", y="PUESTOS", title="Puestos por nivel de prioridad", color="NIVEL_PRIORIDAD", color_discrete_map={"ALTA": COLORES["rojo"], "MEDIA": COLORES["naranja"], "BAJA": COLORES["verde"]})
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.write("Criterio analitico")
        st.write(
            "ALTA: caida con votacion relevante, alta votacion con baja actividad, alta votacion sin mesas, o brecha entre gestion y resultado electoral. "
            "MEDIA: oportunidad de crecimiento o consolidacion. BAJA: baja votacion o informacion insuficiente."
        )
        st.markdown("</div>", unsafe_allow_html=True)

with tabs[6]:
    st.subheader("Exportables")
    hojas_export = {
        "puestos_votacion": puestos,
        "actividades_campana": actividades,
        "mesas_trabajo": mesas,
        "iglesias": iglesias,
        "resumen_iglesia": resumen_iglesia,
        "resumen_barrio": resumen_barrio,
        "matriz_priorizacion": matriz,
        "informe_ejecutivo": informe,
    }
    cols = st.columns(3)
    for idx, (nombre, df) in enumerate(hojas_export.items()):
        cols[idx % 3].download_button(
            f"Descargar {nombre}.csv",
            data=csv_bytes(df),
            file_name=f"{nombre}.csv",
            mime="text/csv",
        )
    st.download_button(
        "Descargar Excel filtrado",
        data=excel_bytes(hojas_export),
        file_name="kennedy_mira_dashboard_filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        "Descargar base maestra consolidada",
        data=ARCHIVO_CONSOLIDADO.read_bytes(),
        file_name="kennedy_mira_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
