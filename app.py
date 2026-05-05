# app.py
# Dashboard geopolítico Kennedy - Campaña Congreso 2026 / Concejo-JAL 2023 / Mesas de trabajo
# Ejecutar:
#   pip install -r requirements.txt
#   streamlit run app.py

from pathlib import Path
import re
import unicodedata
import numpy as np
import pandas as pd
import streamlit as st
import folium
from streamlit_folium import st_folium
from folium.plugins import MarkerCluster, HeatMap

try:
    import geopandas as gpd
except ImportError:
    gpd = None

st.set_page_config(page_title="Informe Kennedy MIRA 2026", layout="wide")

# =============================
# 1. CONFIGURACIÓN DE ARCHIVOS
# =============================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
PROJECT_PARENT = BASE_DIR.parent


def _normalized_filename(value):
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("utf-8")
    return re.sub(r"\s+", " ", value).strip().upper()


def resolve_data_file(*candidate_names):
    """Busca un archivo por nombre en data/ y en la carpeta padre del proyecto."""
    search_dirs = [DATA_DIR, PROJECT_PARENT, BASE_DIR]
    normalized_candidates = {_normalized_filename(name) for name in candidate_names}

    for folder in search_dirs:
        for name in candidate_names:
            candidate = folder / name
            if candidate.exists():
                return candidate

    for folder in search_dirs:
        if not folder.exists():
            continue
        for candidate in folder.iterdir():
            if _normalized_filename(candidate.name) in normalized_candidates:
                return candidate

    return DATA_DIR / candidate_names[0]


ARCHIVO_CAMPANA = resolve_data_file(
    "CAMPAÑA CONGRESO 2026 KENNEDY (1).xlsx",
    "CAMPAÑA CONGRESO 2026 KENNEDY.xlsx",
)
ARCHIVO_GESTION = resolve_data_file(
    "Copia de Gestión Edil Lorena Garzón - 17 de febrero, 17_07.xlsx",
    "GESTION_EDIL_LORENA.xlsx",
)
ARCHIVO_VOTACION = resolve_data_file(
    "VOTACIÓN 2026.xlsx",
    "VOTACION_2026.xlsx",
)

# GeoJSON opcional: descargue desde IDECA/Datos Abiertos Bogotá una capa de UPZ o UPL
# y guárdela como data/kennedy_upz.geojson o data/kennedy_upl.geojson.
ARCHIVO_GEOMETRIA = resolve_data_file("kennedy_upz.geojson", "kennedy_upl.geojson")

# Coordenadas aproximadas de templos/sedes. Ajustar con datos reales del equipo.
TEMPLOS_COORDS = {
    "KENNEDY CENTRAL": (4.6275, -74.1530),
    "CARVAJAL": (4.6120, -74.1390),
    "PATIO BONITO": (4.6385, -74.1710),
    "KENNEDY CLASS": (4.6110, -74.1730),
    "CLASS": (4.6110, -74.1730),
    "CLASS ROMA": (4.6110, -74.1730),
}

# =============================
# 2. FUNCIONES DE LIMPIEZA
# =============================
def normalizar_texto(x):
    if pd.isna(x):
        return np.nan
    x = str(x).strip().upper()
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("utf-8")
    x = re.sub(r"\s+", " ", x)
    return x


def limpiar_columnas(df):
    df = df.copy()
    df.columns = [normalizar_texto(c).replace(" ", "_") for c in df.columns]
    return df


def extraer_lat_lon(valor):
    """Convierte '4.64, -74.17' en lat/lon numéricos."""
    if pd.isna(valor):
        return pd.Series([np.nan, np.nan])
    texto = str(valor).replace(";", ",")
    nums = re.findall(r"-?\d+(?:[.,]\d+)?", texto)
    if len(nums) >= 2:
        try:
            return pd.Series([float(nums[0].replace(",", ".")), float(nums[1].replace(",", "."))])
        except ValueError:
            return pd.Series([np.nan, np.nan])
    return pd.Series([np.nan, np.nan])


def asignar_temporada(fecha):
    if pd.isna(fecha):
        return "SIN FECHA"
    if fecha < pd.Timestamp("2026-03-08"):
        return "ANTES DE ELECCIÓN"
    if fecha == pd.Timestamp("2026-03-08"):
        return "DÍA ELECCIÓN"
    return "DESPUÉS DE ELECCIÓN"


def leer_excel_seguro(path, sheet_name):
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")

# =============================
# 3. CARGA DE DATOS
# =============================
@st.cache_data(show_spinner=False)
def cargar_datos():
    # Votación 2026 / comparativo 2023
    puestos = limpiar_columnas(leer_excel_seguro(ARCHIVO_VOTACION, "Hoja 5"))
    detalle_puestos = limpiar_columnas(leer_excel_seguro(ARCHIVO_VOTACION, "Hoja 3"))

    # Agendas de campaña
    agenda_general = limpiar_columnas(leer_excel_seguro(ARCHIVO_CAMPANA, "AGENDA GENERAL CON CANDIDATOS"))
    agenda_paralela = limpiar_columnas(leer_excel_seguro(ARCHIVO_CAMPANA, "AGENDA PARALELA"))
    mesas_campana = limpiar_columnas(leer_excel_seguro(ARCHIVO_CAMPANA, "Mesas"))

    agenda = pd.concat([agenda_general, agenda_paralela], ignore_index=True)

    # Gestión edil / mesas, si existe la hoja
    try:
        gestion = limpiar_columnas(leer_excel_seguro(ARCHIVO_GESTION, "SEGUIMIENTO MESAS DE TRABAJO"))
    except Exception:
        gestion = pd.DataFrame()

    # Coordenadas puestos
    if "COORDENADAS" in detalle_puestos.columns:
        detalle_puestos[["LAT", "LON"]] = detalle_puestos["COORDENADAS"].apply(extraer_lat_lon)

    # Estandarización nombres de puesto e iglesia
    for df in [puestos, detalle_puestos]:
        if "PUESTO_DE_VOTACION" in df.columns:
            df["PUESTO_KEY"] = df["PUESTO_DE_VOTACION"].apply(normalizar_texto)
        if "PUESTO" in df.columns:
            df["PUESTO_KEY"] = df["PUESTO"].apply(normalizar_texto)
        if "IGLESIA_RESPONSABLE" in df.columns:
            df["IGLESIA_KEY"] = df["IGLESIA_RESPONSABLE"].apply(normalizar_texto)

    # Cruce puestos con coordenadas
    if "PUESTO_KEY" in puestos.columns and "PUESTO_KEY" in detalle_puestos.columns:
        cols_geo = [c for c in ["PUESTO_KEY", "DIRECCION", "COORDENADAS", "LAT", "LON", "IGLESIA_KEY"] if c in detalle_puestos.columns]
        puestos = puestos.merge(detalle_puestos[cols_geo].drop_duplicates("PUESTO_KEY"), on="PUESTO_KEY", how="left", suffixes=("", "_GEO"))
        if "IGLESIA_KEY_GEO" in puestos.columns:
            puestos["IGLESIA_KEY"] = puestos.get("IGLESIA_KEY", puestos["IGLESIA_KEY_GEO"]).fillna(puestos["IGLESIA_KEY_GEO"])

    # Fechas campaña
    for col in ["FECHA_CAMPANA", "FECHA"]:
        if col in agenda.columns:
            agenda[col] = pd.to_datetime(agenda[col], errors="coerce")
    fecha_col = "FECHA_CAMPANA" if "FECHA_CAMPANA" in agenda.columns else "FECHA"
    if fecha_col in agenda.columns:
        agenda["TEMPORADA"] = agenda[fecha_col].apply(asignar_temporada)

    if "SEDE" in agenda.columns:
        agenda["IGLESIA_KEY"] = agenda["SEDE"].apply(normalizar_texto)

    # Coordenadas agenda si existen
    coord_cols = [c for c in agenda.columns if "COORDEN" in c]
    if coord_cols:
        agenda[["LAT", "LON"]] = agenda[coord_cols[0]].apply(extraer_lat_lon)
    else:
        agenda["LAT"] = np.nan
        agenda["LON"] = np.nan

    # Cuando no hay coordenadas de reunión, usar coordenada del templo como referencia general.
    for iglesia, (lat, lon) in TEMPLOS_COORDS.items():
        mask = agenda["IGLESIA_KEY"].eq(iglesia) & agenda["LAT"].isna()
        agenda.loc[mask, "LAT"] = lat
        agenda.loc[mask, "LON"] = lon

    return puestos, detalle_puestos, agenda, mesas_campana, gestion

puestos, detalle_puestos, agenda, mesas_campana, gestion = cargar_datos()

# =============================
# 4. FILTROS
# =============================
st.title("Informe territorial y electoral Kennedy — Partido MIRA")
st.caption("Congreso 2026 · Concejo/JAL 2023 · Campaña territorial · Mesas de trabajo")

iglesias = sorted([x for x in puestos.get("IGLESIA_KEY", pd.Series(dtype=str)).dropna().unique()])
iglesia_sel = st.sidebar.multiselect("Iglesia / sede responsable", iglesias, default=iglesias)

resultado_sel = st.sidebar.multiselect(
    "Resultado variación promedio",
    sorted([x for x in puestos.get("RESULTADO_VARIACION_PROMEDIO", pd.Series(dtype=str)).dropna().unique()]),
)

puestos_f = puestos.copy()
if iglesia_sel and "IGLESIA_KEY" in puestos_f.columns:
    puestos_f = puestos_f[puestos_f["IGLESIA_KEY"].isin(iglesia_sel)]
if resultado_sel and "RESULTADO_VARIACION_PROMEDIO" in puestos_f.columns:
    puestos_f = puestos_f[puestos_f["RESULTADO_VARIACION_PROMEDIO"].isin(resultado_sel)]

# =============================
# 5. INDICADORES
# =============================
def colnum(df, col):
    return pd.to_numeric(df[col], errors="coerce") if col in df.columns else pd.Series(dtype=float)

k1, k2, k3, k4 = st.columns(4)
k1.metric("Puestos analizados", f"{puestos_f['PUESTO_KEY'].nunique():,.0f}".replace(",", "."))
k2.metric("Promedio 2026", f"{colnum(puestos_f, 'PROMEDIO_2026').sum():,.0f}".replace(",", "."))
k3.metric("Promedio 2023", f"{colnum(puestos_f, 'PROMEDIO_2023').sum():,.0f}".replace(",", "."))
var_total = (colnum(puestos_f, 'PROMEDIO_2026').sum() / colnum(puestos_f, 'PROMEDIO_2023').sum() - 1) if colnum(puestos_f, 'PROMEDIO_2023').sum() else np.nan
k4.metric("Variación agregada", f"{var_total:.1%}" if pd.notna(var_total) else "N/D")

# =============================
# 6. TABLAS DE ANÁLISIS
# =============================
st.subheader("1. Lectura por iglesia responsable")
if "IGLESIA_KEY" in puestos_f.columns:
    resumen_iglesia = puestos_f.groupby("IGLESIA_KEY", dropna=False).agg(
        puestos=("PUESTO_KEY", "nunique"),
        camara_2026=("CAMARA_2026", "sum"),
        senado_2026=("SENADO_2026", "sum"),
        promedio_2026=("PROMEDIO_2026", "sum"),
        promedio_2023=("PROMEDIO_2023", "sum"),
    ).reset_index()
    resumen_iglesia["variacion"] = resumen_iglesia["promedio_2026"] / resumen_iglesia["promedio_2023"] - 1
    st.dataframe(resumen_iglesia.sort_values("promedio_2026", ascending=False), use_container_width=True)

st.subheader("2. Puestos con mayor crecimiento y mayor caída")
col_a, col_b = st.columns(2)
if "VARIACION_PROMEDIO" in puestos_f.columns:
    mejores = puestos_f.sort_values("VARIACION_PROMEDIO", ascending=False).head(15)
    criticos = puestos_f.sort_values("VARIACION_PROMEDIO", ascending=True).head(15)
    col_a.dataframe(mejores[[c for c in ["PUESTO_DE_VOTACION", "IGLESIA_KEY", "PROMEDIO_2023", "PROMEDIO_2026", "VARIACION_PROMEDIO", "SE_HA_HECHO_MESA_DE_TRABAJO"] if c in mejores.columns]], use_container_width=True)
    col_b.dataframe(criticos[[c for c in ["PUESTO_DE_VOTACION", "IGLESIA_KEY", "PROMEDIO_2023", "PROMEDIO_2026", "VARIACION_PROMEDIO", "SE_HA_HECHO_MESA_DE_TRABAJO"] if c in criticos.columns]], use_container_width=True)

# =============================
# 7. MAPA INTERACTIVO
# =============================
st.subheader("3. Mapa interactivo territorial")

m = folium.Map(location=[4.625, -74.160], zoom_start=13, tiles="CartoDB positron")

# Capa UPZ/UPL opcional
if ARCHIVO_GEOMETRIA.exists():
    if gpd is None:
        st.warning("Para cargar la capa UPZ/UPL opcional instala geopandas.")
    else:
        try:
            gdf = gpd.read_file(ARCHIVO_GEOMETRIA)
            # Si trae toda Bogotá, filtrar Kennedy por campos frecuentes.
            for campo in ["LocNombre", "LOCALIDAD", "locnombre", "Nombre_Localidad"]:
                if campo in gdf.columns:
                    gdf = gdf[gdf[campo].astype(str).str.upper().str.contains("KENNEDY", na=False)]
                    break
            folium.GeoJson(
                gdf,
                name="UPZ / UPL Kennedy",
                style_function=lambda x: {"fillColor": "#2E7D32", "color": "#1B5E20", "weight": 1, "fillOpacity": 0.08},
                tooltip=folium.GeoJsonTooltip(fields=[c for c in gdf.columns if c.lower() in ["nombre", "upz", "upl", "nom_upz", "nom_upl"]][:1]) if len(gdf.columns) else None,
            ).add_to(m)
        except Exception as e:
            st.warning(f"No se pudo cargar la geometría UPZ/UPL: {e}")

# Puestos de votación
cluster = MarkerCluster(name="Puestos de votación").add_to(m)
for _, r in puestos_f.dropna(subset=["LAT", "LON"]).iterrows():
    puesto = r.get("PUESTO_DE_VOTACION", r.get("PUESTO_KEY", "Puesto"))
    iglesia = r.get("IGLESIA_KEY", "N/D")
    p2023 = r.get("PROMEDIO_2023", np.nan)
    p2026 = r.get("PROMEDIO_2026", np.nan)
    var = r.get("VARIACION_PROMEDIO", np.nan)
    mesa = r.get("SE_HA_HECHO_MESA_DE_TRABAJO", "N/D")
    var_texto = f"{var:.1%}" if pd.notna(var) else "N/D"
    html = f"""
    <b>{puesto}</b><br>
    Iglesia: {iglesia}<br>
    Promedio 2023: {p2023}<br>
    Promedio 2026: {p2026}<br>
    Variación: {var_texto}<br>
    Mesa de trabajo: {mesa}
    """
    folium.CircleMarker(
        location=[r["LAT"], r["LON"]],
        radius=max(4, min(16, float(r.get("PROMEDIO_2026", 0) or 0) / 20)),
        popup=folium.Popup(html, max_width=330),
        tooltip=str(puesto),
        fill=True,
        fill_opacity=0.75,
        weight=1,
    ).add_to(cluster)

# Heatmap electoral
heat_data = puestos_f.dropna(subset=["LAT", "LON"])[["LAT", "LON", "PROMEDIO_2026"]].values.tolist() if "PROMEDIO_2026" in puestos_f.columns else []
if heat_data:
    HeatMap(heat_data, name="Calor electoral 2026", radius=22, blur=18).add_to(m)

# Actividades de campaña / reuniones
agenda_f = agenda.copy()
if iglesia_sel and "IGLESIA_KEY" in agenda_f.columns:
    agenda_f = agenda_f[agenda_f["IGLESIA_KEY"].isin(iglesia_sel)]

camp_layer = folium.FeatureGroup(name="Actividades campaña / reuniones", show=False).add_to(m)
for _, r in agenda_f.dropna(subset=["LAT", "LON"]).iterrows():
    actividad = r.get("ACTIVIDAD", "Actividad")
    detalle = r.get("DETALLE_DE_LA_ACTIVIDAD", "")
    fecha = r.get("FECHA_CAMPANA", r.get("FECHA", ""))
    sede = r.get("SEDE", r.get("IGLESIA_KEY", ""))
    folium.Marker(
        location=[r["LAT"], r["LON"]],
        tooltip=f"{actividad} - {sede}",
        popup=folium.Popup(f"<b>{actividad}</b><br>{sede}<br>{fecha}<br>{detalle}", max_width=350),
        icon=folium.Icon(icon="users", prefix="fa"),
    ).add_to(camp_layer)

folium.LayerControl().add_to(m)
st_folium(m, width=None, height=650)

# =============================
# 8. MATRIZ DE PRIORIZACIÓN
# =============================
st.subheader("4. Matriz de priorización territorial")
prior = puestos_f.copy()
if {"PROMEDIO_2026", "VARIACION_PROMEDIO"}.issubset(prior.columns):
    prior["PUNTAJE_PRIORIDAD"] = (
        prior["PROMEDIO_2026"].rank(pct=True) * 0.45 +
        (-prior["VARIACION_PROMEDIO"]).rank(pct=True) * 0.35 +
        prior.get("SE_HA_HECHO_MESA_DE_TRABAJO", pd.Series("NO", index=prior.index)).astype(str).str.upper().map({"SI": 0.2, "SÍ": 0.2}).fillna(0)
    )
    prior["CATEGORIA"] = pd.cut(prior["PUNTAJE_PRIORIDAD"], bins=[-1, .33, .66, 1.1], labels=["Baja", "Media", "Alta"])
    cols = [c for c in ["PUESTO_DE_VOTACION", "IGLESIA_KEY", "PROMEDIO_2023", "PROMEDIO_2026", "VARIACION_PROMEDIO", "SE_HA_HECHO_MESA_DE_TRABAJO", "PUNTAJE_PRIORIDAD", "CATEGORIA"] if c in prior.columns]
    st.dataframe(prior.sort_values("PUNTAJE_PRIORIDAD", ascending=False)[cols].head(50), use_container_width=True)

# =============================
# 9. EXPORTABLES
# =============================
st.subheader("5. Exportables para informe")
if st.button("Generar archivos CSV de salida"):
    out = Path("outputs")
    out.mkdir(exist_ok=True)
    puestos_f.to_csv(out / "puestos_analizados.csv", index=False, encoding="utf-8-sig")
    if "resumen_iglesia" in locals():
        resumen_iglesia.to_csv(out / "resumen_por_iglesia.csv", index=False, encoding="utf-8-sig")
    if "prior" in locals():
        prior.to_csv(out / "matriz_priorizacion.csv", index=False, encoding="utf-8-sig")
    st.success("Archivos generados en la carpeta outputs/.")
