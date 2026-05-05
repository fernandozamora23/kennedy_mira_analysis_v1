# app.py
# Dashboard territorial-electoral Kennedy - Campaña Congreso 2026 / Concejo-JAL 2023

from io import BytesIO
from pathlib import Path
import re
import unicodedata

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


st.set_page_config(page_title="Dashboard Kennedy MIRA 2026", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
ARCHIVO_CAMPANA = Path("data/CAMPAÑA CONGRESO 2026 KENNEDY (1).xlsx")
ARCHIVO_GESTION = Path("data/Copia de Gestión Edil Lorena Garzón - 17 de febrero, 17_07.xlsx")
ARCHIVO_VOTACION = Path("data/VOTACIÓN 2026.xlsx")
ARCHIVO_UPZ = Path("data/upz_kennedy.geojson")

COLORES = {
    "azul": "#1F77B4",
    "verde": "#2E7D32",
    "rojo": "#C62828",
    "naranja": "#F57C00",
    "morado": "#6A1B9A",
    "gris": "#6B7280",
    "fondo": "#F6F8FB",
}

IGLESIAS = [
    {
        "IGLESIA": "Class Roma",
        "LATITUD": 4.614359775316158,
        "LONGITUD": -74.17619195767098,
        "URL": "https://direcciones.idmji.org/es/iglesia/359/",
        "TIPO": "Iglesia / templo",
    },
    {
        "IGLESIA": "Patio Bonito",
        "LATITUD": 4.646035122997863,
        "LONGITUD": -74.17300841534194,
        "URL": "https://direcciones.idmji.org/es/iglesia/301/",
        "TIPO": "Iglesia / templo",
    },
    {
        "IGLESIA": "Kennedy",
        "LATITUD": 4.6217386978458155,
        "LONGITUD": -74.16501499477366,
        "URL": "",
        "TIPO": "Iglesia / templo",
    },
    {
        "IGLESIA": "Carvajal",
        "LATITUD": 4.616343469904612,
        "LONGITUD": -74.1404329155982,
        "URL": "",
        "TIPO": "Iglesia / templo",
    },
    {
        "IGLESIA": "Valladolid",
        "LATITUD": 4.647817860855581,
        "LONGITUD": -74.14806885512174,
        "URL": "",
        "TIPO": "Iglesia / templo",
    },
]

iglesias_df = pd.DataFrame(IGLESIAS)


st.markdown(
    """
    <style>
    .stApp { background: #F6F8FB; }
    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 16px 18px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
    }
    div[data-testid="stMetricLabel"] { color: #475569; }
    .block-container { padding-top: 1.8rem; padding-bottom: 2.5rem; }
    .section-card {
        background: white;
        border: 1px solid #E5E7EB;
        border-radius: 10px;
        padding: 18px 20px;
        margin: 8px 0 18px 0;
    }
    .hero-title { font-size: 2.2rem; font-weight: 760; color: #0F172A; margin-bottom: 0.25rem; }
    .hero-subtitle { color: #475569; font-size: 1rem; margin-bottom: 1.2rem; }
    .small-muted { color: #64748B; font-size: 0.9rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def normalizar_texto(valor):
    if pd.isna(valor):
        return np.nan
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("utf-8")
    texto = re.sub(r"\s+", " ", texto)
    return texto


def limpiar_columnas(df):
    df = df.copy()
    df.columns = [normalizar_texto(c).replace(" ", "_") for c in df.columns]
    return df


def ruta_repo(path):
    return BASE_DIR / path


def archivo_existe(path):
    return ruta_repo(path).exists()


@st.cache_data(show_spinner=False)
def hojas_excel(path_str):
    path = ruta_repo(Path(path_str))
    if not path.exists():
        return []
    try:
        return pd.ExcelFile(path, engine="openpyxl").sheet_names
    except Exception:
        return []


def hoja_preferida(hojas, preferidas):
    if not hojas:
        return None
    preferidas_norm = [normalizar_texto(x) for x in preferidas]
    for hoja in hojas:
        if normalizar_texto(hoja) in preferidas_norm:
            return hoja
    return hojas[0]


def selector_hoja(etiqueta, archivo, preferidas, key):
    hojas = hojas_excel(str(archivo))
    if not hojas:
        return None
    default = hoja_preferida(hojas, preferidas)
    indice = hojas.index(default) if default in hojas else 0
    return st.sidebar.selectbox(etiqueta, hojas, index=indice, key=key)


@st.cache_data(show_spinner=False)
def leer_excel_seguro(path_str, sheet_name):
    path = ruta_repo(Path(path_str))
    if not path.exists() or not sheet_name:
        return pd.DataFrame()
    try:
        return limpiar_columnas(pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl"))
    except Exception as exc:
        st.warning(f"No se pudo leer {path.name} / {sheet_name}: {exc}")
        return pd.DataFrame()


def extraer_lat_lon(valor):
    if pd.isna(valor):
        return pd.Series([np.nan, np.nan])
    texto = str(valor).replace(";", ",")
    nums = re.findall(r"-?\d+(?:[.,]\d+)?", texto)
    if len(nums) < 2:
        return pd.Series([np.nan, np.nan])
    try:
        return pd.Series([float(nums[0].replace(",", ".")), float(nums[1].replace(",", "."))])
    except ValueError:
        return pd.Series([np.nan, np.nan])


def serie_vacia(index, value=np.nan):
    return pd.Series(value, index=index)


def primera_columna(df, candidatas, contiene=None):
    for col in candidatas:
        if col in df.columns:
            return col
    if contiene:
        for col in df.columns:
            if all(fragmento in col for fragmento in contiene):
                return col
    return None


def asegurar_columna(df, nombre, candidatas, default=np.nan):
    origen = primera_columna(df, candidatas)
    df[nombre] = df[origen] if origen else default
    return df


def preparar_texto_territorial(df):
    df = df.copy()
    for col in ["IGLESIA_KEY", "BARRIO_KEY", "UPZ_KEY", "PUESTO_KEY", "TIPO_ACTIVIDAD_KEY"]:
        if col not in df.columns:
            base = col.replace("_KEY", "")
            df[col] = df[base].apply(normalizar_texto) if base in df.columns else np.nan
    return df


def asignar_temporada(fecha):
    if pd.isna(fecha):
        return "Sin fecha"
    if fecha < pd.Timestamp("2026-03-08"):
        return "Antes de elección"
    if fecha == pd.Timestamp("2026-03-08"):
        return "Día elección"
    return "Después de elección"


def asignar_upz_por_geometria(puntos_df, upz_geojson):
    if gpd is None or puntos_df.empty or not ruta_repo(upz_geojson).exists():
        return pd.Series(np.nan, index=puntos_df.index)
    if not {"LAT", "LON"}.issubset(puntos_df.columns):
        return pd.Series(np.nan, index=puntos_df.index)
    try:
        puntos_geo = puntos_df.dropna(subset=["LAT", "LON"]).copy()
        if puntos_geo.empty:
            return pd.Series(np.nan, index=puntos_df.index)
        puntos_gdf = gpd.GeoDataFrame(
            puntos_geo,
            geometry=gpd.points_from_xy(puntos_geo["LON"], puntos_geo["LAT"]),
            crs="EPSG:4326",
        )
        upz_gdf = gpd.read_file(ruta_repo(upz_geojson)).to_crs("EPSG:4326")
        nombre_col = primera_columna(
            upz_gdf,
            ["UPZ", "NOMBRE", "NOM_UPZ", "NOMBRE_UPZ", "UPL", "NOM_UPL"],
            contiene=["UPZ"],
        )
        if not nombre_col:
            nombre_col = upz_gdf.columns[0]
        joined = gpd.sjoin(puntos_gdf, upz_gdf[[nombre_col, "geometry"]], how="left", predicate="within")
        resultado = pd.Series(np.nan, index=puntos_df.index)
        resultado.loc[joined.index] = joined[nombre_col].values
        return resultado
    except Exception as exc:
        st.warning(f"No se pudo hacer cruce espacial con UPZ: {exc}")
        return pd.Series(np.nan, index=puntos_df.index)


def coord_iglesia(nombre):
    if pd.isna(nombre):
        return (np.nan, np.nan)
    nombre_norm = normalizar_texto(nombre)
    aliases = {
        "CLASS": "CLASS ROMA",
        "KENNEDY CLASS": "CLASS ROMA",
        "KENNEDY CENTRAL": "KENNEDY",
    }
    nombre_norm = aliases.get(nombre_norm, nombre_norm)
    for item in IGLESIAS:
        if normalizar_texto(item["IGLESIA"]) == nombre_norm:
            return (item["LATITUD"], item["LONGITUD"])
    return (np.nan, np.nan)


def preparar_puestos(puestos, detalle):
    puestos = puestos.copy()
    detalle = detalle.copy()
    asegurar_columna(puestos, "PUESTO", ["PUESTO_DE_VOTACION", "PUESTO"])
    asegurar_columna(puestos, "IGLESIA", ["IGLESIA_RESPONSABLE", "IGLESIA"])
    asegurar_columna(puestos, "BARRIO", ["BARRIO", "BARRIO_"])
    asegurar_columna(puestos, "UPZ", ["UPZ", "ZONA"])

    for col in ["PROMEDIO_2026", "PROMEDIO_2023", "CAMARA_2026", "SENADO_2026"]:
        if col in puestos.columns:
            puestos[col] = pd.to_numeric(puestos[col], errors="coerce")
    if "PROMEDIO_2026" not in puestos.columns:
        puestos["PROMEDIO_2026"] = serie_vacia(puestos.index, 0)
    if "PROMEDIO_2023" not in puestos.columns:
        puestos["PROMEDIO_2023"] = serie_vacia(puestos.index, 0)

    puestos["PUESTO_KEY"] = puestos["PUESTO"].apply(normalizar_texto)
    puestos["IGLESIA_KEY"] = puestos["IGLESIA"].apply(normalizar_texto)

    if "PUESTO" in detalle.columns:
        detalle["PUESTO_KEY"] = detalle["PUESTO"].apply(normalizar_texto)
    if "PUESTO_DE_VOTACION" in detalle.columns:
        detalle["PUESTO_KEY"] = detalle["PUESTO_DE_VOTACION"].apply(normalizar_texto)
    if "COORDENADAS" in detalle.columns:
        detalle[["LAT", "LON"]] = detalle["COORDENADAS"].apply(extraer_lat_lon)
    asegurar_columna(detalle, "DIRECCION", ["DIRECCION", "DIRECCION_"])
    asegurar_columna(detalle, "IGLESIA_DETALLE", ["IGLESIA_RESPONSABLE", "IGLESIA"])
    asegurar_columna(detalle, "BARRIO_DETALLE", ["BARRIO", "BARRIO_"])
    asegurar_columna(detalle, "UPZ_DETALLE", ["UPZ", "ZONA"])

    if "PUESTO_KEY" in detalle.columns:
        cols = [c for c in ["PUESTO_KEY", "DIRECCION", "COORDENADAS", "LAT", "LON", "IGLESIA_DETALLE", "BARRIO_DETALLE", "UPZ_DETALLE"] if c in detalle.columns]
        puestos = puestos.merge(detalle[cols].drop_duplicates("PUESTO_KEY"), on="PUESTO_KEY", how="left")

    for col_base, col_detalle in [("IGLESIA", "IGLESIA_DETALLE"), ("BARRIO", "BARRIO_DETALLE"), ("UPZ", "UPZ_DETALLE")]:
        if col_detalle in puestos.columns:
            puestos[col_base] = puestos[col_base].fillna(puestos[col_detalle])

    puestos["BARRIO"] = puestos["BARRIO"].fillna("Sin barrio")
    puestos["UPZ"] = puestos["UPZ"].fillna("Sin UPZ")
    puestos["VARIACION_ABS"] = puestos["PROMEDIO_2026"] - puestos["PROMEDIO_2023"]
    puestos["VARIACION_PCT"] = np.where(
        puestos["PROMEDIO_2023"].fillna(0).ne(0),
        puestos["VARIACION_ABS"] / puestos["PROMEDIO_2023"],
        np.nan,
    )
    if "LAT" not in puestos.columns:
        puestos["LAT"] = np.nan
    if "LON" not in puestos.columns:
        puestos["LON"] = np.nan
    puestos = preparar_texto_territorial(puestos)
    return puestos


def preparar_agenda(agenda_general, agenda_paralela):
    agenda = pd.concat([agenda_general, agenda_paralela], ignore_index=True)
    if agenda.empty:
        return agenda
    asegurar_columna(agenda, "IGLESIA", ["SEDE", "IGLESIA"])
    asegurar_columna(agenda, "BARRIO", ["BARRIO", "BARRIO_"])
    asegurar_columna(agenda, "UPZ", ["UPZ", "ZONA"])
    asegurar_columna(agenda, "TIPO_ACTIVIDAD", ["ACTIVIDAD", "TIPO_DE_ACTIVIDAD"])
    asegurar_columna(agenda, "DETALLE", ["DETALLE_DE_LA_ACTIVIDAD", "DESCRIPCION_Y/O_ACTIVIDAD", "DESCRIPCION"])
    asegurar_columna(agenda, "FECHA", ["FECHA_CAMPANA", "FECHA", "FECHA_DE_INICIO"])

    agenda["FECHA"] = pd.to_datetime(agenda["FECHA"], errors="coerce")
    agenda["TEMPORADA"] = agenda["FECHA"].apply(asignar_temporada)

    coord_col = primera_columna(agenda, ["COORDENADAS", "COORDENADAS_DE_LA_REUNION/_GESTION"], contiene=["COORDEN"])
    if coord_col:
        agenda[["LAT", "LON"]] = agenda[coord_col].apply(extraer_lat_lon)
    else:
        agenda["LAT"] = np.nan
        agenda["LON"] = np.nan

    for idx, row in agenda[agenda["LAT"].isna()].iterrows():
        lat, lon = coord_iglesia(row.get("IGLESIA"))
        agenda.loc[idx, ["LAT", "LON"]] = [lat, lon]

    agenda["BARRIO"] = agenda["BARRIO"].fillna("Sin barrio")
    agenda["UPZ"] = agenda["UPZ"].fillna("Sin UPZ")
    agenda = preparar_texto_territorial(agenda)
    return agenda


def preparar_mesas(mesas_campana, gestion):
    mesas = pd.concat([mesas_campana, gestion], ignore_index=True)
    if mesas.empty:
        return mesas
    asegurar_columna(mesas, "IGLESIA", ["IGLESIA", "IGLESIA_RESPONSABLE"])
    asegurar_columna(mesas, "BARRIO", ["BARRIO", "BARRIO_"])
    asegurar_columna(mesas, "UPZ", ["UPZ", "ZONA"])
    asegurar_columna(mesas, "TEMA", ["TEMAS", "TEMA", "OBJETIVO_DE_LA_REUNION", "NOMBRE_GESTION"])
    asegurar_columna(mesas, "ESTADO", ["ESTADO", "ESTADO_DE_SEGUIMIENTO", "ESTADO_ACTUAL_DEL_TRAMITE"])
    asegurar_columna(mesas, "RESPONSABLE", ["RESPONSABLE", "ENCARGADO", "APOYO_ASESOR"])
    asegurar_columna(mesas, "FECHA", ["FECHA", "FECHA_DE_INICIO", "MES"])
    coord_col = primera_columna(mesas, ["COORDENADAS", "GEOREFERENCIACION"], contiene=["COORDEN"])
    if coord_col:
        mesas[["LAT", "LON"]] = mesas[coord_col].apply(extraer_lat_lon)
    else:
        mesas["LAT"] = np.nan
        mesas["LON"] = np.nan
    for idx, row in mesas[mesas["LAT"].isna()].iterrows():
        lat, lon = coord_iglesia(row.get("IGLESIA"))
        mesas.loc[idx, ["LAT", "LON"]] = [lat, lon]
    mesas["BARRIO"] = mesas["BARRIO"].fillna("Sin barrio")
    mesas["UPZ"] = mesas["UPZ"].fillna("Sin UPZ")
    mesas = preparar_texto_territorial(mesas)
    return mesas


def agregar_conteos_territoriales(puestos, agenda, mesas):
    puestos = puestos.copy()
    for llave, nombre, origen in [
        ("IGLESIA_KEY", "ACTIVIDADES_IGLESIA", agenda),
        ("BARRIO_KEY", "ACTIVIDADES_BARRIO", agenda),
        ("UPZ_KEY", "ACTIVIDADES_UPZ", agenda),
    ]:
        if llave in origen.columns:
            conteo = origen.groupby(llave).size()
            puestos[nombre] = puestos[llave].map(conteo).fillna(0).astype(int)
        else:
            puestos[nombre] = 0
    for llave, nombre, origen in [
        ("IGLESIA_KEY", "MESAS_IGLESIA", mesas),
        ("BARRIO_KEY", "MESAS_BARRIO", mesas),
        ("UPZ_KEY", "MESAS_UPZ", mesas),
    ]:
        if llave in origen.columns:
            conteo = origen.groupby(llave).size()
            puestos[nombre] = puestos[llave].map(conteo).fillna(0).astype(int)
        else:
            puestos[nombre] = 0
    puestos["ACTIVIDADES_CERCANAS"] = puestos[["ACTIVIDADES_IGLESIA", "ACTIVIDADES_BARRIO", "ACTIVIDADES_UPZ"]].max(axis=1)
    puestos["MESAS_CERCANAS"] = puestos[["MESAS_IGLESIA", "MESAS_BARRIO", "MESAS_UPZ"]].max(axis=1)
    return puestos


def construir_prioridad_puestos(puestos):
    df = puestos.copy()
    votos = df["PROMEDIO_2026"].fillna(0)
    var = df["VARIACION_PCT"]
    umbral_alto = votos.quantile(0.65) if len(votos) else 0
    umbral_medio = votos.quantile(0.35) if len(votos) else 0

    razones = []
    acciones = []
    prioridad = []
    for _, row in df.iterrows():
        r = []
        a = []
        p = "Baja"
        votos_altos = row["PROMEDIO_2026"] >= umbral_alto
        votos_medios = row["PROMEDIO_2026"] >= umbral_medio
        caida_fuerte = pd.notna(row["VARIACION_PCT"]) and row["VARIACION_PCT"] <= -0.08
        crecimiento = pd.notna(row["VARIACION_PCT"]) and row["VARIACION_PCT"] > 0.05
        sin_actividad = row.get("ACTIVIDADES_CERCANAS", 0) == 0
        sin_mesa = row.get("MESAS_CERCANAS", 0) == 0

        if caida_fuerte and votos_altos:
            p = "Alta"
            r.append("caida fuerte con votacion relevante")
            a.append("programar visita territorial y seguimiento semanal")
        if votos_altos and sin_actividad:
            p = "Alta"
            r.append("votos altos sin actividades de campana registradas")
            a.append("activar agenda de contacto con lideres y referidos")
        if votos_altos and sin_mesa:
            p = "Alta"
            r.append("votos altos sin mesas de gestion registradas")
            a.append("abrir mesa de trabajo con responsable territorial")
        if p != "Alta" and votos_medios:
            p = "Media"
            r.append("votacion media con oportunidad de consolidacion")
            a.append("mantener presencia quincenal y validar lideres")
        if p != "Alta" and crecimiento:
            p = "Media"
            r.append("crecimiento moderado que debe consolidarse")
            a.append("replicar tacticas del puesto en barrios cercanos")
        if not r:
            r.append("bajo volumen electoral o datos insuficientes")
            a.append("monitorear y actualizar informacion")

        prioridad.append(p)
        razones.append("; ".join(r))
        acciones.append("; ".join(dict.fromkeys(a)))

    df["PRIORIDAD"] = prioridad
    df["RAZON_PRIORIDAD"] = razones
    df["ACCION_RECOMENDADA"] = acciones
    df["RESPONSABLE_SUGERIDO"] = df["IGLESIA"].fillna("Equipo territorial")
    df["TEMPORALIDAD_SUGERIDA"] = df["PRIORIDAD"].map({"Alta": "7 dias", "Media": "15 dias", "Baja": "30 dias"})
    return df


def resumen_territorial(df, llave, nombre_col, agenda, mesas):
    if df.empty or llave not in df.columns:
        return pd.DataFrame()
    resumen = df.groupby([llave, nombre_col], dropna=False).agg(
        votos_2026=("PROMEDIO_2026", "sum"),
        votos_2023=("PROMEDIO_2023", "sum"),
        puestos=("PUESTO_KEY", "nunique"),
        prioridad_alta=("PRIORIDAD", lambda x: (x == "Alta").sum()),
    ).reset_index()
    resumen["variacion_absoluta"] = resumen["votos_2026"] - resumen["votos_2023"]
    resumen["variacion_porcentual"] = np.where(resumen["votos_2023"].ne(0), resumen["variacion_absoluta"] / resumen["votos_2023"], np.nan)

    for origen, col in [(agenda, "actividades"), (mesas, "mesas")]:
        if llave in origen.columns and not origen.empty:
            conteo = origen.groupby(llave).size().rename(col).reset_index()
            resumen = resumen.merge(conteo, on=llave, how="left")
        else:
            resumen[col] = 0
    resumen[["actividades", "mesas"]] = resumen[["actividades", "mesas"]].fillna(0).astype(int)
    resumen["prioridad_territorial"] = np.select(
        [
            (resumen["prioridad_alta"] > 0) | ((resumen["votos_2026"] >= resumen["votos_2026"].quantile(0.65)) & (resumen["actividades"] == 0)),
            resumen["votos_2026"] >= resumen["votos_2026"].quantile(0.35),
        ],
        ["Alta", "Media"],
        default="Baja",
    )
    return resumen.sort_values("votos_2026", ascending=False)


def hallazgos(puestos, resumen_iglesia, resumen_upz):
    items = []
    if not resumen_iglesia.empty:
        mayor = resumen_iglesia.sort_values("votos_2026", ascending=False).iloc[0]
        items.append(f"Mayor votacion por iglesia: {mayor['IGLESIA']} con {mayor['votos_2026']:,.0f} votos 2026.".replace(",", "."))
        caida = resumen_iglesia.sort_values("variacion_absoluta", ascending=True).iloc[0]
        items.append(f"Mayor caida por iglesia: {caida['IGLESIA']} con {caida['variacion_absoluta']:,.0f} votos de variacion absoluta.".replace(",", "."))
    if not puestos.empty:
        crece = puestos.sort_values("VARIACION_ABS", ascending=False).iloc[0]
        baja = puestos.sort_values("VARIACION_ABS", ascending=True).iloc[0]
        items.append(f"Puesto con mayor crecimiento: {crece['PUESTO']} ({crece['VARIACION_ABS']:,.0f}).".replace(",", "."))
        items.append(f"Puesto con mayor caida: {baja['PUESTO']} ({baja['VARIACION_ABS']:,.0f}).".replace(",", "."))
        sin_mesas = puestos[(puestos["PROMEDIO_2026"] >= puestos["PROMEDIO_2026"].quantile(0.65)) & (puestos["MESAS_CERCANAS"] == 0)]
        if not sin_mesas.empty:
            items.append(f"{len(sin_mesas)} puestos de alta votacion no tienen mesas de gestion cercanas registradas.")
    if not resumen_upz.empty:
        upz_alta = resumen_upz[resumen_upz["prioridad_territorial"] == "Alta"]
        if not upz_alta.empty:
            items.append("UPZ prioritarias: " + ", ".join(upz_alta["UPZ"].head(5).astype(str)) + ".")
    return items


def formato_numero(valor):
    if pd.isna(valor):
        return "N/D"
    return f"{valor:,.0f}".replace(",", ".")


def formato_pct(valor):
    if pd.isna(valor):
        return "N/D"
    return f"{valor:.1%}"


def to_csv_bytes(df):
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def construir_excel(descargas):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nombre, df in descargas.items():
            df.to_excel(writer, sheet_name=nombre[:31], index=False)
    return output.getvalue()


def color_variacion(valor):
    if pd.isna(valor):
        return COLORES["gris"]
    return COLORES["verde"] if valor >= 0 else COLORES["rojo"]


def agregar_leyenda(mapa):
    leyenda = """
    <div style="position: fixed; bottom: 44px; left: 44px; z-index: 9999; background: white;
        border: 1px solid #CBD5E1; border-radius: 8px; padding: 12px 14px; font-size: 12px;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.18);">
        <b>Lectura territorial</b><br>
        <span style="color:#2E7D32;">●</span> Variacion positiva<br>
        <span style="color:#C62828;">●</span> Variacion negativa<br>
        <span style="color:#6A1B9A;">●</span> Iglesia / templo<br>
        <span style="color:#F57C00;">●</span> Gestion / mesas<br>
        <span style="color:#1F77B4;">●</span> Campana
    </div>
    """
    mapa.get_root().html.add_child(folium.Element(leyenda))


@st.cache_data(show_spinner=False)
def cargar_datos(hojas):
    puestos_raw = leer_excel_seguro(str(ARCHIVO_VOTACION), hojas["puestos"])
    detalle_raw = leer_excel_seguro(str(ARCHIVO_VOTACION), hojas["detalle"])
    agenda_general = leer_excel_seguro(str(ARCHIVO_CAMPANA), hojas["agenda_general"])
    agenda_paralela = leer_excel_seguro(str(ARCHIVO_CAMPANA), hojas["agenda_paralela"])
    mesas_campana = leer_excel_seguro(str(ARCHIVO_CAMPANA), hojas["mesas_campana"])
    gestion = leer_excel_seguro(str(ARCHIVO_GESTION), hojas["gestion"])

    puestos = preparar_puestos(puestos_raw, detalle_raw)
    agenda = preparar_agenda(agenda_general, agenda_paralela)
    mesas = preparar_mesas(mesas_campana, gestion)

    if ARCHIVO_UPZ.exists():
        for df in [puestos, agenda, mesas]:
            if "UPZ" not in df.columns or df["UPZ"].eq("Sin UPZ").all():
                upz_geom = asignar_upz_por_geometria(df, ARCHIVO_UPZ)
                df["UPZ"] = df["UPZ"].mask(df["UPZ"].eq("Sin UPZ") & upz_geom.notna(), upz_geom)
                df["UPZ_KEY"] = df["UPZ"].apply(normalizar_texto)

    puestos = agregar_conteos_territoriales(puestos, agenda, mesas)
    puestos = construir_prioridad_puestos(puestos)
    return puestos, agenda, mesas


def validar_archivos():
    faltantes = [archivo for archivo in [ARCHIVO_CAMPANA, ARCHIVO_GESTION, ARCHIVO_VOTACION] if not archivo_existe(archivo)]
    if not faltantes:
        return True
    st.error("No se encontraron todos los archivos requeridos en la carpeta data/ del repositorio.")
    for archivo in faltantes:
        st.warning(f"Falta: {archivo}")
    st.info("Agrega los archivos Excel a data/ y vuelve a ejecutar la app.")
    return False


st.markdown(
    """
    <div class="hero-title">Dashboard territorial-electoral Kennedy</div>
    <div class="hero-subtitle">Campaña Congreso 2026 · Partido MIRA · Votación, gestión, iglesias, barrios, UPZ y priorización territorial</div>
    """,
    unsafe_allow_html=True,
)

if not validar_archivos():
    st.stop()

st.sidebar.header("Configuración de datos")
hojas = {
    "puestos": selector_hoja("Votación - resumen por puesto", ARCHIVO_VOTACION, ["Hoja 5"], "hoja_puestos"),
    "detalle": selector_hoja("Votación - detalle y coordenadas", ARCHIVO_VOTACION, ["Hoja 3"], "hoja_detalle"),
    "agenda_general": selector_hoja("Campaña - agenda general", ARCHIVO_CAMPANA, ["AGENDA GENERAL CON CANDIDATOS"], "hoja_agenda_general"),
    "agenda_paralela": selector_hoja("Campaña - agenda paralela", ARCHIVO_CAMPANA, ["AGENDA PARALELA"], "hoja_agenda_paralela"),
    "mesas_campana": selector_hoja("Campaña - mesas", ARCHIVO_CAMPANA, ["Mesas"], "hoja_mesas_campana"),
    "gestion": selector_hoja("Gestión - mesas de trabajo", ARCHIVO_GESTION, ["SEGUIMIENTO MESAS DE TRABAJO"], "hoja_gestion"),
}

if any(v is None for v in hojas.values()):
    st.error("No fue posible detectar las hojas necesarias en los archivos Excel.")
    st.stop()

with st.spinner("Preparando datos territoriales y electorales..."):
    puestos, agenda, mesas = cargar_datos(hojas)

if not ARCHIVO_UPZ.exists():
    st.sidebar.warning("No se encontro data/upz_kennedy.geojson. El dashboard usara UPZ si existe como columna en los Excel.")
elif gpd is None:
    st.sidebar.warning("GeoJSON UPZ detectado, pero geopandas no esta instalado. Se omite el cruce espacial.")

st.sidebar.header("Filtros territoriales")
iglesias_opciones = sorted([x for x in puestos["IGLESIA_KEY"].dropna().unique()])
barrios_opciones = sorted([x for x in puestos["BARRIO_KEY"].dropna().unique()])
upz_opciones = sorted([x for x in puestos["UPZ_KEY"].dropna().unique()])
puestos_opciones = sorted([x for x in puestos["PUESTO_KEY"].dropna().unique()])
tipo_actividad_opciones = sorted([x for x in agenda.get("TIPO_ACTIVIDAD_KEY", pd.Series(dtype=str)).dropna().unique()])
prioridad_opciones = ["Alta", "Media", "Baja"]

iglesia_sel = st.sidebar.multiselect("Iglesia / templo", iglesias_opciones, default=iglesias_opciones)
barrio_sel = st.sidebar.multiselect("Barrio", barrios_opciones)
upz_sel = st.sidebar.multiselect("UPZ", upz_opciones)
tipo_actividad_sel = st.sidebar.multiselect("Tipo de actividad", tipo_actividad_opciones)
prioridad_sel = st.sidebar.multiselect("Prioridad", prioridad_opciones)
puesto_sel = st.sidebar.multiselect("Puesto de votacion", puestos_opciones)

puestos_f = puestos.copy()
agenda_f = agenda.copy()
mesas_f = mesas.copy()
if iglesia_sel:
    puestos_f = puestos_f[puestos_f["IGLESIA_KEY"].isin(iglesia_sel)]
    agenda_f = agenda_f[agenda_f["IGLESIA_KEY"].isin(iglesia_sel)] if "IGLESIA_KEY" in agenda_f.columns else agenda_f
    mesas_f = mesas_f[mesas_f["IGLESIA_KEY"].isin(iglesia_sel)] if "IGLESIA_KEY" in mesas_f.columns else mesas_f
if barrio_sel:
    puestos_f = puestos_f[puestos_f["BARRIO_KEY"].isin(barrio_sel)]
    agenda_f = agenda_f[agenda_f["BARRIO_KEY"].isin(barrio_sel)] if "BARRIO_KEY" in agenda_f.columns else agenda_f
    mesas_f = mesas_f[mesas_f["BARRIO_KEY"].isin(barrio_sel)] if "BARRIO_KEY" in mesas_f.columns else mesas_f
if upz_sel:
    puestos_f = puestos_f[puestos_f["UPZ_KEY"].isin(upz_sel)]
    agenda_f = agenda_f[agenda_f["UPZ_KEY"].isin(upz_sel)] if "UPZ_KEY" in agenda_f.columns else agenda_f
    mesas_f = mesas_f[mesas_f["UPZ_KEY"].isin(upz_sel)] if "UPZ_KEY" in mesas_f.columns else mesas_f
if tipo_actividad_sel and "TIPO_ACTIVIDAD_KEY" in agenda_f.columns:
    agenda_f = agenda_f[agenda_f["TIPO_ACTIVIDAD_KEY"].isin(tipo_actividad_sel)]
if prioridad_sel:
    puestos_f = puestos_f[puestos_f["PRIORIDAD"].isin(prioridad_sel)]
if puesto_sel:
    puestos_f = puestos_f[puestos_f["PUESTO_KEY"].isin(puesto_sel)]

resumen_iglesia = resumen_territorial(puestos_f, "IGLESIA_KEY", "IGLESIA", agenda_f, mesas_f)
resumen_barrio = resumen_territorial(puestos_f, "BARRIO_KEY", "BARRIO", agenda_f, mesas_f)
resumen_upz = resumen_territorial(puestos_f, "UPZ_KEY", "UPZ", agenda_f, mesas_f)
matriz_priorizacion = puestos_f.sort_values(["PRIORIDAD", "PROMEDIO_2026"], ascending=[True, False])

votos_2026 = puestos_f["PROMEDIO_2026"].sum()
votos_2023 = puestos_f["PROMEDIO_2023"].sum()
var_abs = votos_2026 - votos_2023
var_pct = var_abs / votos_2023 if votos_2023 else np.nan

k1, k2, k3, k4 = st.columns(4)
k1.metric("Votos 2026", formato_numero(votos_2026))
k2.metric("Votos 2023", formato_numero(votos_2023))
k3.metric("Variacion absoluta", formato_numero(var_abs), delta=formato_numero(var_abs))
k4.metric("Variacion porcentual", formato_pct(var_pct), delta=formato_pct(var_pct))
k5, k6, k7, k8 = st.columns(4)
k5.metric("Puestos analizados", formato_numero(puestos_f["PUESTO_KEY"].nunique()))
k6.metric("Actividades de campana", formato_numero(len(agenda_f)))
k7.metric("Mesas de trabajo", formato_numero(len(mesas_f)))
k8.metric("Iglesias analizadas", formato_numero(puestos_f["IGLESIA_KEY"].nunique()))

hallazgos_auto = hallazgos(puestos_f, resumen_iglesia, resumen_upz)

tab_resumen, tab_mapa, tab_iglesias, tab_graficas, tab_priorizacion, tab_informe, tab_exportables = st.tabs(
    ["Resumen ejecutivo", "Mapa territorial", "Iglesias / templos", "Graficas", "Priorizacion", "Informe ejecutivo", "Exportables"]
)

with tab_resumen:
    st.subheader("Portada y lectura ejecutiva")
    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.write("Este tablero cruza votacion 2026, base comparativa 2023, actividades de campana, mesas de trabajo, iglesias responsables, puestos de votacion y variables territoriales disponibles.")
        st.write("La lectura prioriza puntos donde hay alto potencial electoral, caidas relevantes o ausencia de presencia territorial registrada.")
        st.markdown("</div>", unsafe_allow_html=True)
        with st.expander("Hallazgos automaticos", expanded=True):
            for item in hallazgos_auto:
                st.write(f"- {item}")
        with st.expander("Recomendacion estrategica", expanded=True):
            altas = matriz_priorizacion[matriz_priorizacion["PRIORIDAD"] == "Alta"].head(10)
            if altas.empty:
                st.write("Mantener seguimiento territorial y completar informacion de barrios, UPZ y mesas.")
            else:
                st.write("Priorizar visitas, reuniones y mesas de gestion en estos puestos:")
                st.dataframe(altas[["PUESTO", "IGLESIA", "BARRIO", "UPZ", "PROMEDIO_2026", "VARIACION_ABS", "RAZON_PRIORIDAD", "ACCION_RECOMENDADA"]], use_container_width=True)
    with c2:
        if not resumen_iglesia.empty:
            fig = px.bar(
                resumen_iglesia.sort_values("votos_2026"),
                x="votos_2026",
                y="IGLESIA",
                orientation="h",
                title="Votos 2026 por iglesia responsable",
                color_discrete_sequence=[COLORES["azul"]],
            )
            fig.update_layout(height=420, margin=dict(l=10, r=10, t=55, b=10))
            st.plotly_chart(fig, use_container_width=True)

with tab_mapa:
    st.subheader("Mapa interactivo territorial")
    mapa = folium.Map(location=[4.628, -74.160], zoom_start=13, tiles="CartoDB positron", control_scale=True)
    Fullscreen(position="topleft").add_to(mapa)
    MiniMap(toggle_display=True, minimized=True).add_to(mapa)

    if ARCHIVO_UPZ.exists() and gpd is not None:
        try:
            gdf = gpd.read_file(ruta_repo(ARCHIVO_UPZ)).to_crs("EPSG:4326")
            folium.GeoJson(
                gdf,
                name="UPZ Kennedy",
                style_function=lambda _: {"fillColor": "#64748B", "color": "#334155", "weight": 1, "fillOpacity": 0.08},
            ).add_to(mapa)
        except Exception as exc:
            st.warning(f"No se pudo cargar la capa UPZ: {exc}")

    puestos_cluster = MarkerCluster(name="Puestos de votacion").add_to(mapa)
    variacion_layer = folium.FeatureGroup(name="Variacion electoral", show=False).add_to(mapa)
    for _, row in puestos_f.dropna(subset=["LAT", "LON"]).iterrows():
        votos_26 = row.get("PROMEDIO_2026", np.nan)
        votos_23 = row.get("PROMEDIO_2023", np.nan)
        popup = f"""
        <div style="font-family: Arial; font-size: 13px; min-width: 260px;">
            <h4 style="margin: 0 0 8px 0;">{row.get('PUESTO', 'Puesto')}</h4>
            <b>Barrio:</b> {row.get('BARRIO', 'N/D')}<br>
            <b>UPZ:</b> {row.get('UPZ', 'N/D')}<br>
            <b>Iglesia responsable:</b> {row.get('IGLESIA', 'N/D')}<br>
            <b>Votos 2026:</b> {formato_numero(votos_26)}<br>
            <b>Votos 2023:</b> {formato_numero(votos_23)}<br>
            <b>Variacion abs.:</b> {formato_numero(row.get('VARIACION_ABS'))}<br>
            <b>Variacion %:</b> {formato_pct(row.get('VARIACION_PCT'))}<br>
            <b>Actividades cercanas:</b> {row.get('ACTIVIDADES_CERCANAS', 0)}<br>
            <b>Mesas cercanas:</b> {row.get('MESAS_CERCANAS', 0)}<br>
            <b>Prioridad:</b> {row.get('PRIORIDAD', 'N/D')}
        </div>
        """
        radio = max(5, min(18, float(votos_26 or 0) / 18))
        folium.CircleMarker(
            location=[row["LAT"], row["LON"]],
            radius=radio,
            color=COLORES["azul"],
            fill=True,
            fill_color=COLORES["azul"],
            fill_opacity=0.72,
            weight=1,
            tooltip=f"{row.get('PUESTO', 'Puesto')} · {formato_numero(votos_26)} votos",
            popup=folium.Popup(popup, max_width=360),
        ).add_to(puestos_cluster)
        folium.CircleMarker(
            location=[row["LAT"], row["LON"]],
            radius=radio + 2,
            color=color_variacion(row.get("VARIACION_ABS")),
            fill=False,
            weight=3,
            tooltip=f"Variacion: {formato_numero(row.get('VARIACION_ABS'))}",
        ).add_to(variacion_layer)

    iglesias_layer = folium.FeatureGroup(name="Iglesias", show=True).add_to(mapa)
    for _, row in iglesias_df.iterrows():
        link = f"<br><a href='{row['URL']}' target='_blank'>Ver direccion</a>" if row["URL"] else ""
        popup = f"<b>{row['IGLESIA']}</b><br>{row['TIPO']}<br>Lat: {row['LATITUD']}<br>Lon: {row['LONGITUD']}{link}"
        folium.Marker(
            location=[row["LATITUD"], row["LONGITUD"]],
            tooltip=f"Iglesia: {row['IGLESIA']}",
            popup=folium.Popup(popup, max_width=280),
            icon=folium.Icon(color="purple", icon="home", prefix="fa"),
        ).add_to(iglesias_layer)

    mesas_layer = folium.FeatureGroup(name="Mesas de trabajo", show=False).add_to(mapa)
    for _, row in mesas_f.dropna(subset=["LAT", "LON"]).iterrows():
        popup = f"<b>{row.get('TEMA', 'Mesa de trabajo')}</b><br>{row.get('BARRIO', '')}<br>{row.get('IGLESIA', '')}<br>{row.get('ESTADO', '')}"
        folium.Marker(
            location=[row["LAT"], row["LON"]],
            tooltip=str(row.get("TEMA", "Mesa de trabajo")),
            popup=folium.Popup(popup, max_width=320),
            icon=folium.Icon(color="orange", icon="briefcase", prefix="fa"),
        ).add_to(mesas_layer)

    agenda_layer = folium.FeatureGroup(name="Actividades de campana", show=False).add_to(mapa)
    for _, row in agenda_f.dropna(subset=["LAT", "LON"]).iterrows():
        popup = f"<b>{row.get('TIPO_ACTIVIDAD', 'Actividad')}</b><br>{row.get('IGLESIA', '')}<br>{row.get('BARRIO', '')}<br>{row.get('DETALLE', '')}"
        folium.Marker(
            location=[row["LAT"], row["LON"]],
            tooltip=str(row.get("TIPO_ACTIVIDAD", "Actividad")),
            popup=folium.Popup(popup, max_width=320),
            icon=folium.Icon(color="blue", icon="users", prefix="fa"),
        ).add_to(agenda_layer)

    heat_data = puestos_f.dropna(subset=["LAT", "LON"])[["LAT", "LON", "PROMEDIO_2026"]].values.tolist()
    if heat_data:
        HeatMap(heat_data, name="Calor votos 2026", radius=24, blur=18).add_to(mapa)

    agregar_leyenda(mapa)
    folium.LayerControl(collapsed=False).add_to(mapa)
    st_folium(mapa, width=None, height=720)
    with st.expander("Como leer este mapa"):
        st.write("El tamano de los puntos de votacion representa votos 2026. El contorno verde indica crecimiento; rojo indica caida. Las capas se pueden activar o desactivar desde el control del mapa.")

with tab_iglesias:
    st.subheader("Analisis por iglesia / templo")
    for iglesia in ["Class Roma", "Patio Bonito", "Kennedy", "Carvajal", "Valladolid"]:
        iglesia_key = normalizar_texto(iglesia)
        df_i = puestos_f[puestos_f["IGLESIA_KEY"].eq(iglesia_key)]
        agenda_i = agenda_f[agenda_f["IGLESIA_KEY"].eq(iglesia_key)] if "IGLESIA_KEY" in agenda_f.columns else pd.DataFrame()
        mesas_i = mesas_f[mesas_f["IGLESIA_KEY"].eq(iglesia_key)] if "IGLESIA_KEY" in mesas_f.columns else pd.DataFrame()
        with st.expander(iglesia, expanded=not df_i.empty):
            if df_i.empty:
                st.write("No hay puestos asociados en los datos filtrados.")
                continue
            a, b, c, d = st.columns(4)
            a.metric("Votos 2026", formato_numero(df_i["PROMEDIO_2026"].sum()))
            b.metric("Votos 2023", formato_numero(df_i["PROMEDIO_2023"].sum()))
            d_abs = df_i["VARIACION_ABS"].sum()
            c.metric("Variacion", formato_numero(d_abs), delta=formato_numero(d_abs))
            d.metric("Puestos", formato_numero(df_i["PUESTO_KEY"].nunique()))
            st.write(
                f"Barrios asociados: {', '.join(df_i['BARRIO'].dropna().astype(str).unique()[:8]) or 'N/D'} · "
                f"UPZ asociadas: {', '.join(df_i['UPZ'].dropna().astype(str).unique()[:8]) or 'N/D'} · "
                f"Actividades: {len(agenda_i)} · Mesas: {len(mesas_i)}"
            )
            col1, col2 = st.columns([1, 1])
            with col1:
                st.dataframe(
                    df_i[["PUESTO", "BARRIO", "UPZ", "PROMEDIO_2026", "PROMEDIO_2023", "VARIACION_ABS", "VARIACION_PCT", "PRIORIDAD"]].sort_values("PROMEDIO_2026", ascending=False),
                    use_container_width=True,
                )
            with col2:
                fig = px.bar(
                    df_i.sort_values("VARIACION_ABS").tail(12),
                    x="VARIACION_ABS",
                    y="PUESTO",
                    orientation="h",
                    title=f"Variacion absoluta por puesto · {iglesia}",
                    color="VARIACION_ABS",
                    color_continuous_scale=["#C62828", "#E5E7EB", "#2E7D32"],
                )
                fig.update_layout(height=420, coloraxis_showscale=False)
                st.plotly_chart(fig, use_container_width=True)
            with st.expander("Lectura estrategica"):
                mejor = df_i.sort_values("VARIACION_ABS", ascending=False).iloc[0]
                peor = df_i.sort_values("VARIACION_ABS", ascending=True).iloc[0]
                st.write(
                    f"{iglesia} suma {formato_numero(df_i['PROMEDIO_2026'].sum())} votos 2026. "
                    f"El puesto que mas crece es {mejor['PUESTO']} y el punto mas critico es {peor['PUESTO']}. "
                    "La recomendacion es concentrar gestion donde hay alta votacion y baja presencia registrada."
                )

with tab_graficas:
    st.subheader("Graficas profesionales")
    if not resumen_iglesia.empty:
        col1, col2 = st.columns(2)
        with col1:
            fig = px.bar(resumen_iglesia.sort_values("votos_2026"), x="votos_2026", y="IGLESIA", orientation="h", title="Votos 2026 por iglesia", color_discrete_sequence=[COLORES["azul"]])
            fig.update_layout(height=430)
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            fig = px.pie(resumen_iglesia, names="IGLESIA", values="votos_2026", hole=0.55, title="Participacion por iglesia", color_discrete_sequence=px.colors.qualitative.Set2)
            fig.update_layout(height=430)
            st.plotly_chart(fig, use_container_width=True)
    col3, col4 = st.columns(2)
    with col3:
        top_crece = puestos_f.sort_values("VARIACION_ABS", ascending=False).head(10)
        fig = px.bar(top_crece.sort_values("VARIACION_ABS"), x="VARIACION_ABS", y="PUESTO", orientation="h", title="Top 10 crecimiento", color_discrete_sequence=[COLORES["verde"]])
        fig.update_layout(height=430)
        st.plotly_chart(fig, use_container_width=True)
    with col4:
        top_cae = puestos_f.sort_values("VARIACION_ABS", ascending=True).head(10)
        fig = px.bar(top_cae.sort_values("VARIACION_ABS", ascending=False), x="VARIACION_ABS", y="PUESTO", orientation="h", title="Top 10 caida", color_discrete_sequence=[COLORES["rojo"]])
        fig.update_layout(height=430)
        st.plotly_chart(fig, use_container_width=True)

    puestos_plot = puestos_f.copy()
    puestos_plot["ACTIVIDADES_CERCANAS"] = puestos_plot["ACTIVIDADES_CERCANAS"].clip(lower=1)
    fig = px.scatter(
        puestos_plot,
        x="PROMEDIO_2026",
        y="VARIACION_PCT",
        size="ACTIVIDADES_CERCANAS",
        color="PRIORIDAD",
        hover_name="PUESTO",
        hover_data=["IGLESIA", "BARRIO", "UPZ", "MESAS_CERCANAS"],
        title="Votos 2026 vs variacion porcentual",
        color_discrete_map={"Alta": COLORES["rojo"], "Media": COLORES["naranja"], "Baja": COLORES["verde"]},
    )
    fig.update_layout(height=520)
    st.plotly_chart(fig, use_container_width=True)

    matriz = puestos_f.pivot_table(index="IGLESIA", columns="UPZ", values="PROMEDIO_2026", aggfunc="sum", fill_value=0)
    if not matriz.empty:
        fig = px.imshow(matriz, text_auto=True, aspect="auto", title="Matriz iglesia vs UPZ por votos 2026", color_continuous_scale="Blues")
        fig.update_layout(height=520)
        st.plotly_chart(fig, use_container_width=True)
    with st.expander("Como leer estos graficos"):
        st.write("Las barras muestran concentracion y cambios; el scatter cruza volumen electoral con desempeno; la matriz permite ver donde se superponen iglesias y UPZ.")

with tab_priorizacion:
    st.subheader("Matriz de priorizacion")
    st.dataframe(
        matriz_priorizacion[[
            "PRIORIDAD",
            "PUESTO",
            "IGLESIA",
            "BARRIO",
            "UPZ",
            "PROMEDIO_2026",
            "PROMEDIO_2023",
            "VARIACION_ABS",
            "VARIACION_PCT",
            "RAZON_PRIORIDAD",
            "ACCION_RECOMENDADA",
            "RESPONSABLE_SUGERIDO",
            "TEMPORALIDAD_SUGERIDA",
        ]],
        use_container_width=True,
    )
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Resumen por UPZ")
        st.dataframe(resumen_upz, use_container_width=True)
    with c2:
        st.subheader("Resumen por barrio")
        st.dataframe(resumen_barrio.head(30), use_container_width=True)
    with st.expander("Recomendacion estrategica"):
        st.write("Prioridad alta combina caida fuerte, votacion relevante y ausencia de presencia territorial registrada. Prioridad media senala oportunidades de consolidacion.")

with tab_informe:
    st.subheader("Informe ejecutivo automatico")
    informe = [
        f"El tablero consolida {formato_numero(puestos_f['PUESTO_KEY'].nunique())} puestos de votacion, {formato_numero(len(agenda_f))} actividades de campana y {formato_numero(len(mesas_f))} mesas o gestiones.",
        f"El acumulado filtrado muestra {formato_numero(votos_2026)} votos 2026 frente a {formato_numero(votos_2023)} votos 2023, con variacion de {formato_numero(var_abs)} votos ({formato_pct(var_pct)}).",
    ]
    if hallazgos_auto:
        informe.append("Hallazgos principales: " + " ".join(hallazgos_auto))
    altas = matriz_priorizacion[matriz_priorizacion["PRIORIDAD"] == "Alta"].head(5)
    if not altas.empty:
        informe.append("Puestos criticos para agenda territorial: " + ", ".join(altas["PUESTO"].astype(str)) + ".")
    if not resumen_upz.empty:
        informe.append("Analisis por UPZ: priorizar " + ", ".join(resumen_upz.head(5)["UPZ"].astype(str)) + " por concentracion electoral y senales de presencia territorial.")
    informe.append("Recomendacion: combinar visitas a puestos de alta votacion, mesas de gestion en zonas sin cobertura y seguimiento quincenal a barrios con actividad sin crecimiento.")
    st.markdown("\n\n".join(informe))
    with st.expander("Agenda territorial sugerida"):
        st.write("Semana 1: puestos prioridad alta y barrios sin mesas. Semana 2: consolidacion de iglesias con crecimiento. Semana 3: UPZ con concentracion electoral y baja presencia territorial. Semana 4: revision de resultados y actualizacion de matriz.")

with tab_exportables:
    st.subheader("Exportables")
    resumen_puesto = matriz_priorizacion.copy()
    descargas = {
        "resumen_iglesia": resumen_iglesia,
        "resumen_puesto": resumen_puesto,
        "resumen_barrio": resumen_barrio,
        "resumen_upz": resumen_upz,
        "matriz_priorizacion": matriz_priorizacion,
    }
    cols = st.columns(3)
    for idx, (nombre, df) in enumerate(descargas.items()):
        cols[idx % 3].download_button(
            f"Descargar {nombre}.csv",
            data=to_csv_bytes(df),
            file_name=f"{nombre}.csv",
            mime="text/csv",
        )
    st.download_button(
        "Descargar Excel consolidado",
        data=construir_excel(descargas),
        file_name="dashboard_kennedy_mira_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
