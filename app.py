
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
import hmac
import html
from io import BytesIO
import json
import math
import sqlite3
from datetime import datetime, timezone

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import folium
from folium.plugins import HeatMap, Fullscreen, MiniMap
from streamlit_folium import st_folium

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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    .stApp {
        background: #F8FAFC;
        color: #0F172A;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
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
        background: rgba(255, 255, 255, 0.9);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(226, 232, 240, 0.8);
        border-radius: 16px;
        padding: 1.5rem 1.75rem;
        box-shadow: 0 8px 30px rgba(15, 23, 42, 0.04);
        margin-bottom: 1.5rem;
    }

    .metric-card {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(226, 232, 240, 0.9);
        border-radius: 16px;
        padding: 1.25rem 1.5rem;
        box-shadow: 0 4px 20px rgba(15, 23, 42, 0.06);
        min-height: 110px;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(15, 23, 42, 0.08);
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
        background: rgba(255, 255, 255, 0.9);
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(15, 23, 42, 0.03);
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
            st.session_state["usuario_actual"] = usuario
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")

    return False


if not check_password():
    st.stop()


AJUSTES_FILE = DATA_DIR / "ajustes_guardados.json"
AJUSTES_DB_FILE = DATA_DIR / "ajustes_territoriales.db"
GOOGLE_SHEETS_ACTUALES = "ajustes_actuales"
GOOGLE_SHEETS_HISTORIAL = "ajustes_historial"
AJUSTES_ACTUALES_COLUMNS = ["entidad", "entidad_id", "nombre_entidad", "templo_actual", "usuario", "motivo", "actualizado_en"]
AJUSTES_HISTORIAL_COLUMNS = ["id", "creado_en", "entidad", "entidad_id", "nombre_entidad", "templo_anterior", "templo_nuevo", "usuario", "motivo"]

def cargar_ajustes_guardados():
    if AJUSTES_FILE.exists():
        try:
            with open(AJUSTES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"ajustes_asignacion": {}, "ajustes_mesas": {}, "ajustes_actividades": {}}


def _ahora_utc_iso():
    return datetime.now(timezone.utc).isoformat()


def _session_key_to_entidad(session_key):
    mapping = {
        "ajustes_asignacion": "puesto",
        "ajustes_mesas": "mesa",
        "ajustes_actividades": "actividad",
    }
    if session_key not in mapping:
        raise ValueError(f"Session key no soportada: {session_key}")
    return mapping[session_key]


def _entidad_to_session_key(entidad):
    mapping = {
        "puesto": "ajustes_asignacion",
        "mesa": "ajustes_mesas",
        "actividad": "ajustes_actividades",
    }
    return mapping.get(entidad)


def _normalize_entity_id(entidad, entity_id):
    if entidad in {"mesa", "actividad"}:
        try:
            return int(entity_id)
        except Exception:
            return str(entity_id)
    return str(entity_id)


def get_db_connection():
    conn = sqlite3.connect(AJUSTES_DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def _empty_ajustes_actuales_df():
    return pd.DataFrame(columns=AJUSTES_ACTUALES_COLUMNS)


def _empty_historial_ajustes_df():
    return pd.DataFrame(columns=AJUSTES_HISTORIAL_COLUMNS)


def _normalize_sheet_df(df, columns):
    if df is None or df.empty:
        return pd.DataFrame(columns=columns)
    out = df.copy()
    for col in columns:
        if col not in out.columns:
            out[col] = ""
    return out[columns].fillna("")


def _google_sheets_config():
    try:
        sheets_cfg = st.secrets.get("google_sheets", {})
        service_account = st.secrets.get("gcp_service_account", {})
    except Exception:
        return None

    spreadsheet_id = str(sheets_cfg.get("spreadsheet_id", "")).strip() if sheets_cfg else ""
    if not spreadsheet_id or not service_account:
        return None

    service_account_info = dict(service_account)
    if "private_key" in service_account_info:
        service_account_info["private_key"] = str(service_account_info["private_key"]).replace("\\n", "\n")
    return spreadsheet_id, service_account_info


@st.cache_resource(show_spinner=False)
def _get_google_spreadsheet():
    cfg = _google_sheets_config()
    if not cfg:
        return None

    spreadsheet_id, service_account_info = cfg
    import gspread
    from google.oauth2.service_account import Credentials

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    client = gspread.authorize(credentials)
    return client.open_by_key(spreadsheet_id)


def _get_or_create_worksheet(spreadsheet, title, columns):
    try:
        worksheet = spreadsheet.worksheet(title)
    except Exception:
        worksheet = spreadsheet.add_worksheet(title=title, rows=1000, cols=max(len(columns), 8))

    values = worksheet.get_all_values()
    if not values:
        worksheet.append_row(columns, value_input_option="USER_ENTERED")
    elif values[0] != columns:
        for idx, col in enumerate(columns, start=1):
            worksheet.update_cell(1, idx, col)
    return worksheet


def _init_google_sheets_storage():
    spreadsheet = _get_google_spreadsheet()
    if spreadsheet is None:
        return False
    _get_or_create_worksheet(spreadsheet, GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS)
    _get_or_create_worksheet(spreadsheet, GOOGLE_SHEETS_HISTORIAL, AJUSTES_HISTORIAL_COLUMNS)
    return True


def _google_sheets_ready():
    if not _google_sheets_config():
        return False
    if "google_sheets_ready" not in st.session_state:
        try:
            st.session_state["google_sheets_ready"] = bool(_init_google_sheets_storage())
            st.session_state.pop("google_sheets_error", None)
        except Exception as exc:
            st.session_state["google_sheets_ready"] = False
            st.session_state["google_sheets_error"] = str(exc)
    return bool(st.session_state.get("google_sheets_ready"))


def persistence_backend_label():
    if _google_sheets_ready():
        return "Google Sheets online"
    if _google_sheets_config() and st.session_state.get("google_sheets_error"):
        return "SQLite local (Google Sheets no conectado)"
    return "SQLite local"


def _read_google_sheet_df(title, columns):
    spreadsheet = _get_google_spreadsheet()
    worksheet = _get_or_create_worksheet(spreadsheet, title, columns)
    records = worksheet.get_all_records(default_blank="")
    return _normalize_sheet_df(pd.DataFrame(records), columns)


def _rewrite_google_sheet_df(title, columns, df):
    spreadsheet = _get_google_spreadsheet()
    worksheet = _get_or_create_worksheet(spreadsheet, title, columns)
    clean_df = _normalize_sheet_df(df, columns).astype(str)
    rows = clean_df.values.tolist()
    worksheet.clear()
    worksheet.append_row(columns, value_input_option="USER_ENTERED")
    if rows:
        worksheet.append_rows(rows, value_input_option="USER_ENTERED")


def _append_google_sheet_row(title, columns, row):
    spreadsheet = _get_google_spreadsheet()
    worksheet = _get_or_create_worksheet(spreadsheet, title, columns)
    worksheet.append_row([str(row.get(col, "") or "") for col in columns], value_input_option="USER_ENTERED")


def _next_historial_id(historial_df):
    if historial_df.empty or "id" not in historial_df.columns:
        return 1
    current_max = pd.to_numeric(historial_df["id"], errors="coerce").max()
    if pd.isna(current_max):
        return 1
    return int(current_max) + 1


def _init_ajustes_sqlite():
    with get_db_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS ajustes_actuales (
                entidad TEXT NOT NULL,
                entidad_id TEXT NOT NULL,
                nombre_entidad TEXT,
                templo_actual TEXT NOT NULL,
                usuario TEXT,
                motivo TEXT,
                actualizado_en TEXT NOT NULL,
                PRIMARY KEY (entidad, entidad_id)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS ajustes_historial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entidad TEXT NOT NULL,
                entidad_id TEXT NOT NULL,
                nombre_entidad TEXT,
                templo_anterior TEXT,
                templo_nuevo TEXT,
                usuario TEXT,
                motivo TEXT,
                creado_en TEXT NOT NULL
            )
            """
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_historial_creado_en ON ajustes_historial(creado_en DESC)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_historial_entidad_id ON ajustes_historial(entidad, entidad_id)")


def init_ajustes_db():
    _init_ajustes_sqlite()
    if _google_sheets_config():
        _google_sheets_ready()


def _cargar_ajustes_desde_sqlite():
    ajustes = {"ajustes_asignacion": {}, "ajustes_mesas": {}, "ajustes_actividades": {}}
    with get_db_connection() as conn:
        rows = conn.execute("SELECT entidad, entidad_id, templo_actual FROM ajustes_actuales").fetchall()
    for row in rows:
        session_key = _entidad_to_session_key(row["entidad"])
        if not session_key:
            continue
        key = _normalize_entity_id(row["entidad"], row["entidad_id"])
        ajustes[session_key][key] = row["templo_actual"]
    return ajustes


def _cargar_ajustes_desde_google_sheets():
    ajustes = {"ajustes_asignacion": {}, "ajustes_mesas": {}, "ajustes_actividades": {}}
    rows = _read_google_sheet_df(GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS)
    for _, row in rows.iterrows():
        session_key = _entidad_to_session_key(row["entidad"])
        if not session_key:
            continue
        key = _normalize_entity_id(row["entidad"], row["entidad_id"])
        ajustes[session_key][key] = row["templo_actual"]
    return ajustes


def cargar_ajustes_desde_db():
    if _google_sheets_ready():
        try:
            return _cargar_ajustes_desde_google_sheets()
        except Exception as exc:
            st.session_state["google_sheets_error"] = str(exc)
    return _cargar_ajustes_desde_sqlite()


def _registrar_ajuste_sqlite(session_key, entity_id, nombre_entidad, templo_nuevo, motivo=""):
    entidad = _session_key_to_entidad(session_key)
    entidad_id = str(_normalize_entity_id(entidad, entity_id))
    usuario = st.session_state.get("usuario_actual", "usuario_dashboard")
    now = _ahora_utc_iso()

    with get_db_connection() as conn:
        previo = conn.execute(
            "SELECT templo_actual FROM ajustes_actuales WHERE entidad = ? AND entidad_id = ?",
            (entidad, entidad_id),
        ).fetchone()
        templo_anterior = previo["templo_actual"] if previo else None
        if templo_anterior == str(templo_nuevo):
            return False, templo_anterior

        conn.execute(
            """
            INSERT INTO ajustes_actuales(entidad, entidad_id, nombre_entidad, templo_actual, usuario, motivo, actualizado_en)
            VALUES(?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(entidad, entidad_id) DO UPDATE SET
                nombre_entidad=excluded.nombre_entidad,
                templo_actual=excluded.templo_actual,
                usuario=excluded.usuario,
                motivo=excluded.motivo,
                actualizado_en=excluded.actualizado_en
            """,
            (entidad, entidad_id, str(nombre_entidad or ""), str(templo_nuevo), usuario, str(motivo or ""), now),
        )
        conn.execute(
            """
            INSERT INTO ajustes_historial(entidad, entidad_id, nombre_entidad, templo_anterior, templo_nuevo, usuario, motivo, creado_en)
            VALUES(?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (entidad, entidad_id, str(nombre_entidad or ""), templo_anterior, str(templo_nuevo), usuario, str(motivo or ""), now),
        )
    return True, templo_anterior


def _registrar_ajuste_google_sheets(session_key, entity_id, nombre_entidad, templo_nuevo, motivo=""):
    entidad = _session_key_to_entidad(session_key)
    entidad_id = str(_normalize_entity_id(entidad, entity_id))
    usuario = st.session_state.get("usuario_actual", "usuario_dashboard")
    now = _ahora_utc_iso()

    actuales = _read_google_sheet_df(GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS)
    historial = _read_google_sheet_df(GOOGLE_SHEETS_HISTORIAL, AJUSTES_HISTORIAL_COLUMNS)
    match = actuales["entidad"].astype(str).eq(entidad) & actuales["entidad_id"].astype(str).eq(entidad_id)
    templo_anterior = actuales.loc[match, "templo_actual"].iloc[0] if match.any() else None
    if str(templo_anterior or "") == str(templo_nuevo):
        return False, templo_anterior

    nueva_fila = {
        "entidad": entidad,
        "entidad_id": entidad_id,
        "nombre_entidad": str(nombre_entidad or ""),
        "templo_actual": str(templo_nuevo),
        "usuario": usuario,
        "motivo": str(motivo or ""),
        "actualizado_en": now,
    }
    actuales = actuales.loc[~match].copy()
    actuales = pd.concat([actuales, pd.DataFrame([nueva_fila])], ignore_index=True)

    historial_row = {
        "id": _next_historial_id(historial),
        "creado_en": now,
        "entidad": entidad,
        "entidad_id": entidad_id,
        "nombre_entidad": str(nombre_entidad or ""),
        "templo_anterior": templo_anterior or "",
        "templo_nuevo": str(templo_nuevo),
        "usuario": usuario,
        "motivo": str(motivo or ""),
    }
    _rewrite_google_sheet_df(GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS, actuales)
    _append_google_sheet_row(GOOGLE_SHEETS_HISTORIAL, AJUSTES_HISTORIAL_COLUMNS, historial_row)
    return True, templo_anterior


def registrar_ajuste_en_db(session_key, entity_id, nombre_entidad, templo_nuevo, motivo=""):
    if _google_sheets_ready():
        try:
            changed, templo_anterior = _registrar_ajuste_google_sheets(session_key, entity_id, nombre_entidad, templo_nuevo, motivo)
            if changed:
                _registrar_ajuste_sqlite(session_key, entity_id, nombre_entidad, templo_nuevo, motivo)
            return changed, templo_anterior
        except Exception as exc:
            st.session_state["google_sheets_error"] = str(exc)
            st.warning("No se pudo guardar en Google Sheets. Se guardará una copia local en SQLite.")
    return _registrar_ajuste_sqlite(session_key, entity_id, nombre_entidad, templo_nuevo, motivo)


def _limpiar_ajustes_sqlite(session_key, motivo=""):
    entidad = _session_key_to_entidad(session_key)
    usuario = st.session_state.get("usuario_actual", "usuario_dashboard")
    now = _ahora_utc_iso()
    with get_db_connection() as conn:
        actuales = conn.execute(
            "SELECT entidad_id, nombre_entidad, templo_actual FROM ajustes_actuales WHERE entidad = ?",
            (entidad,),
        ).fetchall()
        for row in actuales:
            conn.execute(
                """
                INSERT INTO ajustes_historial(entidad, entidad_id, nombre_entidad, templo_anterior, templo_nuevo, usuario, motivo, creado_en)
                VALUES(?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (entidad, row["entidad_id"], row["nombre_entidad"], row["templo_actual"], None, usuario, str(motivo or "limpieza masiva"), now),
            )
        conn.execute("DELETE FROM ajustes_actuales WHERE entidad = ?", (entidad,))
    return len(actuales)


def _limpiar_ajustes_google_sheets(session_key, motivo=""):
    entidad = _session_key_to_entidad(session_key)
    usuario = st.session_state.get("usuario_actual", "usuario_dashboard")
    now = _ahora_utc_iso()
    actuales = _read_google_sheet_df(GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS)
    historial = _read_google_sheet_df(GOOGLE_SHEETS_HISTORIAL, AJUSTES_HISTORIAL_COLUMNS)
    match = actuales["entidad"].astype(str).eq(entidad)
    rows_to_clear = actuales.loc[match].copy()
    next_id = _next_historial_id(historial)

    for _, row in rows_to_clear.iterrows():
        _append_google_sheet_row(
            GOOGLE_SHEETS_HISTORIAL,
            AJUSTES_HISTORIAL_COLUMNS,
            {
                "id": next_id,
                "creado_en": now,
                "entidad": entidad,
                "entidad_id": row.get("entidad_id", ""),
                "nombre_entidad": row.get("nombre_entidad", ""),
                "templo_anterior": row.get("templo_actual", ""),
                "templo_nuevo": "",
                "usuario": usuario,
                "motivo": str(motivo or "limpieza masiva"),
            },
        )
        next_id += 1

    actuales = actuales.loc[~match].copy()
    _rewrite_google_sheet_df(GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS, actuales)
    return len(rows_to_clear)


def limpiar_ajustes_en_db(session_key, motivo=""):
    if _google_sheets_ready():
        try:
            total = _limpiar_ajustes_google_sheets(session_key, motivo)
            _limpiar_ajustes_sqlite(session_key, motivo)
            return total
        except Exception as exc:
            st.session_state["google_sheets_error"] = str(exc)
            st.warning("No se pudo limpiar Google Sheets. Se aplicará la limpieza local en SQLite.")
    return _limpiar_ajustes_sqlite(session_key, motivo)


def _obtener_historial_sqlite(limit=300):
    with get_db_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, creado_en, entidad, entidad_id, nombre_entidad, templo_anterior, templo_nuevo, usuario, motivo
            FROM ajustes_historial
            ORDER BY id DESC
            LIMIT ?
            """,
            (int(limit),),
        ).fetchall()
    if not rows:
        return _empty_historial_ajustes_df()
    return pd.DataFrame([dict(r) for r in rows])


def obtener_historial_ajustes(limit=300):
    if _google_sheets_ready():
        try:
            historial = _read_google_sheet_df(GOOGLE_SHEETS_HISTORIAL, AJUSTES_HISTORIAL_COLUMNS)
            if historial.empty:
                return _empty_historial_ajustes_df()
            historial["id_sort"] = pd.to_numeric(historial["id"], errors="coerce").fillna(0)
            return historial.sort_values("id_sort", ascending=False).drop(columns=["id_sort"]).head(int(limit)).reset_index(drop=True)
        except Exception as exc:
            st.session_state["google_sheets_error"] = str(exc)
    return _obtener_historial_sqlite(limit)


def _obtener_ajustes_actuales_sqlite(entidad=None):
    where = ""
    params = ()
    if entidad:
        where = "WHERE entidad = ?"
        params = (entidad,)
    with get_db_connection() as conn:
        rows = conn.execute(
            f"""
            SELECT entidad, entidad_id, nombre_entidad, templo_actual, usuario, motivo, actualizado_en
            FROM ajustes_actuales
            {where}
            ORDER BY actualizado_en DESC
            """,
            params,
        ).fetchall()
    if not rows:
        return _empty_ajustes_actuales_df()
    return pd.DataFrame([dict(r) for r in rows])


def obtener_ajustes_actuales_df(entidad=None):
    if _google_sheets_ready():
        try:
            actuales = _read_google_sheet_df(GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS)
            if entidad:
                actuales = actuales[actuales["entidad"].astype(str).eq(str(entidad))]
            if actuales.empty:
                return _empty_ajustes_actuales_df()
            return actuales.sort_values("actualizado_en", ascending=False).reset_index(drop=True)
        except Exception as exc:
            st.session_state["google_sheets_error"] = str(exc)
    return _obtener_ajustes_actuales_sqlite(entidad)


def migrar_ajustes_json_a_db(ajustes_json):
    for session_key, payload in (ajustes_json or {}).items():
        if session_key not in {"ajustes_asignacion", "ajustes_mesas", "ajustes_actividades"}:
            continue
        for raw_key, templo in (payload or {}).items():
            try:
                entity_key = raw_key if session_key == "ajustes_asignacion" else int(raw_key)
            except Exception:
                entity_key = raw_key
            registrar_ajuste_en_db(
                session_key=session_key,
                entity_id=entity_key,
                nombre_entidad=str(raw_key),
                templo_nuevo=templo,
                motivo="Migración inicial desde ajustes_guardados.json",
            )


def guardar_ajustes_guardados():
    datos = {
        "ajustes_asignacion": st.session_state.get("ajustes_asignacion", {}),
        "ajustes_mesas": st.session_state.get("ajustes_mesas", {}),
        "ajustes_actividades": st.session_state.get("ajustes_actividades", {})
    }
    with open(AJUSTES_FILE, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)

if "usuario_actual" not in st.session_state:
    st.session_state["usuario_actual"] = "usuario_dashboard"

if "ajustes_cargados" not in st.session_state:
    init_ajustes_db()
    ajustes_db = cargar_ajustes_desde_db()
    if any(bool(v) for v in ajustes_db.values()):
        st.session_state["ajustes_asignacion"] = ajustes_db.get("ajustes_asignacion", {})
        st.session_state["ajustes_mesas"] = ajustes_db.get("ajustes_mesas", {})
        st.session_state["ajustes_actividades"] = ajustes_db.get("ajustes_actividades", {})
    else:
        ajustes_disco = cargar_ajustes_guardados()
        st.session_state["ajustes_asignacion"] = ajustes_disco.get("ajustes_asignacion", {})
        st.session_state["ajustes_mesas"] = ajustes_disco.get("ajustes_mesas", {})
        st.session_state["ajustes_actividades"] = ajustes_disco.get("ajustes_actividades", {})
        if any(bool(v) for v in ajustes_disco.values()):
            migrar_ajustes_json_a_db(ajustes_disco)
    st.session_state["ajustes_cargados"] = True

# ============================================================
# FUNCIONES
# ============================================================

@st.cache_data
def cargar_datos(path: Path, mtime: float = 0):
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


@st.cache_data
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


def agregar_contorno_localidades(mapa):
    localidades_gj = cargar_geojson(LOCALIDADES_GEOJSON)
    if not localidades_gj:
        return False

    def style_localidad(feature):
        nombre = str((feature.get("properties") or {}).get("LocNombre", "")).strip().upper()
        is_kennedy = nombre == "KENNEDY"
        return {
            "fillColor": "#FFFFFF" if is_kennedy else "#94A3B8",
            "color": "#0F172A" if is_kennedy else "#64748B",
            "weight": 2.2 if is_kennedy else 0.9,
            "fillOpacity": 0.03 if is_kennedy else 0.08,
            "dashArray": None if is_kennedy else "4 4",
        }

    folium.GeoJson(
        localidades_gj,
        name="Contorno localidades Bogotá",
        style_function=style_localidad,
        tooltip=folium.GeoJsonTooltip(fields=["LocNombre"], aliases=["Localidad:"], sticky=True),
        show=True,
    ).add_to(mapa)
    return True


def haversine_km(lat1, lon1, lat2, lon2):
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
        return np.nan
    radio_tierra_km = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [float(lat1), float(lon1), float(lat2), float(lon2)])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 2 * radio_tierra_km * math.asin(math.sqrt(a))


@st.cache_data
def calcular_distancias_a_templos_v2(puestos_df, iglesias_df):
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
            iglesia_actual = row.get("IGLESIA_ACTUAL")
            iglesia_actual_valida = iglesia_actual if iglesia_actual in TEMPLOS_OFICIALES else "PENDIENTE"
            row["TEMPLO_MAS_CERCANO"] = templo_cercano
            row["DISTANCIA_MINIMA_KM"] = distancias[templo_cercano]
            row["TEMPLO_ASIGNADO_PROPUESTO"] = iglesia_actual_valida
            row["OBSERVACION_ASIGNACION"] = "Se conserva el templo asignado en el documento base."
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


def aplicar_ajustes_templo(df, session_key, id_col):
    if df is None or df.empty or id_col not in df.columns:
        return df
    out = df.copy()
    if "IGLESIA_ORIGINAL" not in out.columns:
        out["IGLESIA_ORIGINAL"] = out.get("IGLESIA", "")
    ajustes = st.session_state.get(session_key, {})
    if ajustes:
        out["IGLESIA"] = out[id_col].map(ajustes).fillna(out["IGLESIA"])
    out["TEMPLO_AJUSTADO"] = out[id_col].isin(ajustes.keys()) if ajustes else False
    return out


def crear_resumen_operativo_por_templo(actividades_df, mesas_df):
    rows = []
    for templo in TEMPLOS_OFICIALES:
        acts = actividades_df[actividades_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)] if not actividades_df.empty else pd.DataFrame()
        mesas_t = mesas_df[mesas_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)] if not mesas_df.empty else pd.DataFrame()
        rows.append({
            "TEMPLO": templo,
            "ACTIVIDADES": int(len(acts)),
            "VOLANTEOS": int(acts.get("TIPO_ACTIVIDAD", pd.Series(dtype=str)).astype(str).eq("VOLANTEO").sum()) if not acts.empty else 0,
            "MESAS_TRABAJO": int(len(mesas_t)),
            "BENEFICIARIOS_MESAS": int(pd.to_numeric(mesas_t.get("BENEFICIARIOS", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not mesas_t.empty else 0,
            "COMPROMISOS_MESAS": int(pd.to_numeric(mesas_t.get("COMPROMISOS_TOTAL", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not mesas_t.empty else 0,
            "AJUSTES_TEMPORALES": int(acts.get("TEMPLO_AJUSTADO", pd.Series(dtype=bool)).sum()) + int(mesas_t.get("TEMPLO_AJUSTADO", pd.Series(dtype=bool)).sum()),
        })
    return pd.DataFrame(rows)


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


def crear_resumen_asignacion_por_columna(asignacion_df, templo_col):
    if asignacion_df.empty or templo_col not in asignacion_df.columns:
        return pd.DataFrame()
    rows = []
    for templo in TEMPLOS_OFICIALES:
        sub = asignacion_df[asignacion_df[templo_col].eq(templo)].copy()
        v26 = pd.to_numeric(sub.get("VOTOS_2026", pd.Series(dtype=float)), errors="coerce").sum()
        v23 = pd.to_numeric(sub.get("VOTOS_2023", pd.Series(dtype=float)), errors="coerce").sum()
        var_abs = v26 - v23
        rows.append({
            "TEMPLO": templo,
            "PUESTOS": int(len(sub)),
            "VOTOS_2026": int(round(v26)),
            "VOTOS_2023": int(round(v23)),
            "VARIACION_ABSOLUTA": int(round(var_abs)),
            "VARIACION_PORCENTUAL": var_abs / v23 if v23 else np.nan,
            "PRIORIDAD_ALTA": int(sub.get("PRIORIDAD", pd.Series(dtype=str)).astype(str).str.upper().eq("ALTA").sum()),
        })
    return pd.DataFrame(rows)


def crear_tabla_puestos_por_templo(asignacion_df):
    grupos = {
        templo: asignacion_df[asignacion_df["TEMPLO_ASIGNADO_FINAL"].eq(templo)]["PUESTO"].sort_values().tolist()
        for templo in TEMPLOS_OFICIALES
    }
    max_len = max([len(v) for v in grupos.values()] + [0])
    return pd.DataFrame({templo: valores + [""] * (max_len - len(valores)) for templo, valores in grupos.items()})


def crear_icono_div(tipo, color, label):
    if tipo == "templo":
        return folium.DivIcon(
            icon_size=(32, 40),
            icon_anchor=(16, 40),
            html=f'''
            <div style="
                position:relative;
                width:32px;
                height:40px;
                filter:drop-shadow(0 4px 8px rgba(15,23,42,.35));
            ">
                <div style="
                    position:absolute;
                    left:4px;
                    top:0;
                    width:24px;
                    height:24px;
                    background:{color};
                    border:4px solid #FFFFFF;
                    border-radius:50% 50% 50% 0;
                    transform:rotate(-45deg);
                    box-sizing:border-box;
                "></div>
                <div style="
                    position:absolute;
                    left:11px;
                    top:7px;
                    width:10px;
                    height:10px;
                    border-radius:50%;
                    background:#FFFFFF;
                    box-shadow:0 0 0 2px rgba(255,255,255,.25);
                "></div>
            </div>
            '''
        )

    shape = "50%" if tipo == "mesa" else "6px"
    size = 20
    border = 2
    return folium.DivIcon(
        html=f'''
        <div style="
            width:{size}px;
            height:{size}px;
            border-radius:{shape};
            background:{color};
            border:{border}px solid #FFFFFF;
            box-shadow:0 3px 10px rgba(15,23,42,.30);
            color:#FFFFFF;
            font-family:'Inter', Arial, sans-serif;
            font-size:{14 if tipo == "templo" else 11}px;
            font-weight:900;
            display:flex;
            align-items:center;
            justify-content:center;
            line-height:1;
            box-sizing:border-box;
            transform:translate(-50%, -50%);
        ">{label}</div>
        '''
    )


def crear_etiqueta_templo(nombre, color, dx=30, dy=-18):
    return folium.DivIcon(
        html=f'''
        <div style="
            transform:translate({dx}px,{dy}px);
            background:rgba(255,255,255,.92);
            border:1.5px solid rgba(15,23,42,.16);
            border-left:6px solid {color};
            border-radius:7px;
            padding:7px 12px 7px 9px;
            color:#111827;
            font-family:'Inter', Arial, sans-serif;
            font-size:15px;
            font-weight:950;
            line-height:1;
            box-shadow:
                0 0 0 3px rgba(255,255,255,.86),
                0 6px 16px rgba(15,23,42,.22);
            white-space:nowrap;
            text-transform:uppercase;
            text-shadow:
                -1px -1px 0 #FFFFFF,
                1px -1px 0 #FFFFFF,
                -1px 1px 0 #FFFFFF,
                1px 1px 0 #FFFFFF,
                0 2px 0 #FFFFFF;
        ">
            {safe_html(nombre)}
        </div>
        '''
    )


def crear_heat_config(puestos_df):
    valid_heat = puestos_df.dropna(subset=["LATITUD", "LONGITUD", "VOTOS_2026"]).copy()
    if valid_heat.empty:
        return None

    votos = pd.to_numeric(valid_heat["VOTOS_2026"], errors="coerce").fillna(0)
    valid_heat["VOTOS_2026"] = votos
    q10, q40, q70, q90 = votos.quantile([0.10, 0.40, 0.70, 0.90]).tolist()
    vmax = max(float(votos.max()), 1.0)
    valid_heat["PESO_HEAT"] = votos.rank(pct=True, method="average").fillna(0).clip(0.08, 1.0)
    heat_data = valid_heat[["LATITUD", "LONGITUD", "PESO_HEAT"]].values.tolist()

    return {
        "data": heat_data,
        "q10": q10,
        "q40": q40,
        "q70": q70,
        "q90": q90,
        "vmax": vmax,
    }


def agregar_heatmap_electoral(mapa, puestos_df, show=True, name="Rango de calor electoral 2026"):
    heat_config = crear_heat_config(puestos_df)
    if not heat_config:
        return None

    heat_layer = folium.FeatureGroup(name=name, show=show)
    HeatMap(
        heat_config["data"],
        radius=25,
        blur=16,
        min_opacity=0.22,
        gradient={
            0.08: "#ECFEFF",
            0.28: "#7DD3FC",
            0.50: "#2563EB",
            0.72: "#F59E0B",
            0.90: "#EF4444",
            1.00: "#7F1D1D",
        },
    ).add_to(heat_layer)
    heat_layer.add_to(mapa)

    heat_legend_html = f"""
    <div style="position: fixed; top: 118px; left: 18px; z-index:9999; background: rgba(255, 255, 255, 0.96); backdrop-filter: blur(8px); padding:12px 14px; border:1px solid #CBD5E1; border-radius:8px; box-shadow:0 8px 22px rgba(15, 23, 42, 0.12); font-size:12px; width: 240px; font-family:'Inter', Arial, sans-serif;">
    <div style="color:#0F172A; font-weight:900; margin-bottom:7px;">Rango electoral 2026</div>
    <div style="background: linear-gradient(to right, #ECFEFF, #7DD3FC, #2563EB, #F59E0B, #EF4444, #7F1D1D); width: 100%; height: 13px; border-radius: 999px; margin: 8px 0;"></div>
    <div style="display:grid; grid-template-columns:repeat(4,1fr); gap:4px; color:#475569; font-size:10px; font-weight:700;">
        <span>P10<br>{fmt_number(heat_config['q10'],0)}</span>
        <span>P40<br>{fmt_number(heat_config['q40'],0)}</span>
        <span>P70<br>{fmt_number(heat_config['q70'],0)}</span>
        <span style="text-align:right;">P90<br>{fmt_number(heat_config['q90'],0)}</span>
    </div>
    <div style="color:#64748B; font-size:10px; margin-top:6px;">Max: {fmt_number(heat_config['vmax'],0)} votos. Escala ajustada al filtro actual.</div>
    </div>
    """
    mapa.get_root().html.add_child(folium.Element(heat_legend_html))
    return heat_config


def render_folium_map(mapa, height=760, key=None):
    st_folium(
        mapa,
        height=height,
        use_container_width=True,
        returned_objects=[],
        key=key,
    )


def crear_mapa_asignacion(asignacion_df, iglesias_df, layers_config=None):
    layers_config = layers_config or {}
    show_contorno = layers_config.get("contorno", True)
    show_heat = layers_config.get("heat", True)
    show_lineas = layers_config.get("lineas", True)
    show_puestos = layers_config.get("puestos", True)
    show_templos = layers_config.get("templos", True)

    m = folium.Map(location=KENNEDY_CENTER, zoom_start=13, tiles="CartoDB positron", control_scale=True)
    Fullscreen(position="topleft").add_to(m)
    MiniMap(toggle_display=True, position="bottomleft").add_to(m)
    if show_contorno:
        agregar_contorno_localidades(m)
    if show_heat:
        agregar_heatmap_electoral(m, asignacion_df, show=True, name="Rango de calor votos 2026")

    templos = iglesias_df[iglesias_df["IGLESIA"].isin(TEMPLOS_OFICIALES)].dropna(subset=["LATITUD", "LONGITUD"]).copy()
    templo_coords = {r["IGLESIA"]: (r["LATITUD"], r["LONGITUD"]) for _, r in templos.iterrows()}
    templos_layer = folium.FeatureGroup(name="Templos oficiales", show=True)
    for _, r in templos.iterrows():
        color = COLORES_TEMPLOS.get(r["IGLESIA"], "#334155")
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            tooltip=f"Templo oficial: {r['IGLESIA']}",
            popup=folium.Popup(
                f"<b>{safe_html(r['IGLESIA'])}</b><br>Lat: {r['LATITUD']}<br>Lon: {r['LONGITUD']}",
                max_width=260
            ),
            icon=crear_icono_div("templo", color, "T"),
        ).add_to(templos_layer)
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            icon=crear_etiqueta_templo(r["IGLESIA"], color, dx=34, dy=-20),
        ).add_to(templos_layer)

    puestos_layer = folium.FeatureGroup(name="Puestos por templo asignado", show=show_puestos)
    lineas_layer = folium.FeatureGroup(name="Líneas puesto-templo", show=show_lineas)
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
        
        if templo in templo_coords:
            folium.PolyLine(
                locations=[[r["LATITUD"], r["LONGITUD"]], list(templo_coords[templo])],
                color=color,
                weight=1.05,
                opacity=0.24,
                tooltip=f"{r.get('PUESTO')} → {templo} | {fmt_number(distancia, 2)} km",
            ).add_to(lineas_layer)
            
        folium.CircleMarker(
            location=[r["LATITUD"], r["LONGITUD"]],
            radius=5.2,
            color="#FFFFFF",
            fill=True,
            fill_color=color,
            fill_opacity=0.76,
            weight=1.2,
            tooltip=f"{r.get('PUESTO')} | {safe_html(r.get('BARRIO'))} | {safe_html(templo)} | {fmt_number(distancia, 2)} km | {fmt_number(r.get('VOTOS_2026'), 0)} votos",
            popup=folium.Popup(popup, max_width=380),
        ).add_to(puestos_layer)
        
    if show_lineas:
        lineas_layer.add_to(m)
    if show_puestos:
        puestos_layer.add_to(m)
    if show_templos:
        templos_layer.add_to(m)

    legend_items = "".join(
        f'''
        <div style="display:flex;align-items:center;gap:7px;margin:4px 0;">
            <span style="width:11px;height:11px;border-radius:50%;background:{color};border:2px solid #FFFFFF;box-shadow:0 0 0 1px {color};display:inline-block;"></span>
            <span>{templo}</span>
        </div>
        '''
        for templo, color in COLORES_TEMPLOS.items()
    )

    legend_html = f'''
    <div style="
        position: fixed;
        bottom: 35px;
        right: 35px;
        z-index:9999;
        background:white;
        padding:12px 14px;
        border:1px solid #CBD5E1;
        border-radius:8px;
        box-shadow:0 4px 14px rgba(15,23,42,.16);
        font-size:12px;
        color:#0F172A;
        min-width:210px;
        font-family: 'Inter', sans-serif;
    ">
        <div style="font-weight:900;margin-bottom:7px;">Asignación territorial</div>
        {legend_items}
        <div style="height:1px;background:#E2E8F0;margin:8px 0;"></div>
        <div style="display:flex;align-items:center;gap:7px;margin:4px 0;">
            <span style="width:24px;border-top:2px solid #64748B;display:inline-block;"></span>
            <span>Línea puesto-templo</span>
        </div>
        <div style="display:flex;align-items:center;gap:7px;margin:4px 0;">
            <span style="width:14px;height:14px;background:#0F172A;border:3px solid white;border-radius:50% 50% 50% 0;transform:rotate(-45deg);display:inline-block;box-shadow:0 0 0 1px #0F172A;"></span>
            <span>Templo oficial</span>
        </div>
    </div>
    '''
    m.get_root().html.add_child(folium.Element(legend_html))
    return m


def exportar_asignacion_excel(asignacion_df, resumen_df, tabla_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        asignacion_df.to_excel(writer, sheet_name="asignacion_detallada", index=False)
        resumen_df.to_excel(writer, sheet_name="resumen_por_templo", index=False)
        tabla_df.to_excel(writer, sheet_name="tabla_puestos_por_templo", index=False)
    return output.getvalue()


def exportar_asignacion_por_templo_excel(asignacion_df, actividades_df, mesas_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen_global = crear_resumen_asignacion(asignacion_df)
        resumen_global.to_excel(writer, sheet_name="resumen_global", index=False)
        informes = []
        for templo in TEMPLOS_OFICIALES:
            puestos_t = asignacion_df[asignacion_df["TEMPLO_ASIGNADO_FINAL"].eq(templo)].copy()
            acts_t = actividades_df[actividades_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not actividades_df.empty else pd.DataFrame()
            mesas_t = mesas_df[mesas_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not mesas_df.empty else pd.DataFrame()
            informes.append({
                "TEMPLO": templo,
                "PUESTOS": len(puestos_t),
                "VOTOS_2026": pd.to_numeric(puestos_t.get("VOTOS_2026", pd.Series(dtype=float)), errors="coerce").fillna(0).sum(),
                "ACTIVIDADES": len(acts_t),
                "MESAS": len(mesas_t),
                "LECTURA": generar_informe_templo_markdown(templo, puestos_t.rename(columns={"TEMPLO_ASIGNADO_FINAL": "IGLESIA"}), acts_t, mesas_t),
            })

            base_sheet = templo.replace(" ", "_")[:20]
            puestos_t.to_excel(writer, sheet_name=f"{base_sheet}_puestos"[:31], index=False)
            if not acts_t.empty:
                acts_t.to_excel(writer, sheet_name=f"{base_sheet}_acts"[:31], index=False)
            if not mesas_t.empty:
                mesas_t.to_excel(writer, sheet_name=f"{base_sheet}_mesas"[:31], index=False)
        pd.DataFrame(informes).to_excel(writer, sheet_name="lectura_por_templo", index=False)
    return output.getvalue()


def to_excel_bytes(df, sheet_name="Hoja1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()


def multi_sheet_excel_bytes(sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            clean_name = str(sheet_name)[:31]
            (df if df is not None else pd.DataFrame()).to_excel(writer, sheet_name=clean_name, index=False)
    return output.getvalue()


def generar_informe_territorial(asignacion_df, actividades_df, mesas_df):
    resumen_puestos = crear_resumen_asignacion(asignacion_df)
    resumen_operativo = crear_resumen_operativo_por_templo(actividades_df, mesas_df)
    total_puestos = len(asignacion_df)
    total_actividades = len(actividades_df)
    total_mesas = len(mesas_df)
    total_ajustes = len(st.session_state.get("ajustes_asignacion", {})) + len(st.session_state.get("ajustes_actividades", {})) + len(st.session_state.get("ajustes_mesas", {}))
    mayor_puestos = resumen_puestos.sort_values("PUESTOS_ASIGNADOS", ascending=False).iloc[0] if not resumen_puestos.empty else None
    mayor_operativo = resumen_operativo.assign(TOTAL_OPERATIVO=lambda d: d["ACTIVIDADES"] + d["MESAS_TRABAJO"]).sort_values("TOTAL_OPERATIVO", ascending=False).iloc[0] if not resumen_operativo.empty else None

    lineas = [
        "# Informe territorial de asignación y operación",
        "",
        "## Resumen general",
        f"- Puestos de votación analizados: {fmt_number(total_puestos, 0)}.",
        f"- Actividades de campaña y volanteos incluidos: {fmt_number(total_actividades, 0)}.",
        f"- Mesas de trabajo incluidas: {fmt_number(total_mesas, 0)}.",
        f"- Ajustes temporales aplicados en esta sesión: {fmt_number(total_ajustes, 0)}.",
    ]
    if mayor_puestos is not None:
        lineas.append(f"- Mayor carga de puestos propuesta: {mayor_puestos['TEMPLO']} con {fmt_number(mayor_puestos['PUESTOS_ASIGNADOS'], 0)} puestos.")
    if mayor_operativo is not None:
        lineas.append(f"- Mayor concentración operativa: {mayor_operativo['TEMPLO']} con {fmt_number(mayor_operativo['ACTIVIDADES'], 0)} actividades y {fmt_number(mayor_operativo['MESAS_TRABAJO'], 0)} mesas.")

    lineas.extend(["", "## Resumen por templo"])
    for _, r in resumen_operativo.iterrows():
        puestos_row = resumen_puestos[resumen_puestos["TEMPLO"].eq(r["TEMPLO"])]
        puestos_count = int(puestos_row.iloc[0]["PUESTOS_ASIGNADOS"]) if not puestos_row.empty else 0
        lineas.append(
            f"- {r['TEMPLO']}: {fmt_number(puestos_count, 0)} puestos, "
            f"{fmt_number(r['ACTIVIDADES'], 0)} actividades, {fmt_number(r['VOLANTEOS'], 0)} volanteos, "
            f"{fmt_number(r['MESAS_TRABAJO'], 0)} mesas, {fmt_number(r['BENEFICIARIOS_MESAS'], 0)} beneficiarios reportados."
        )

    lineas.extend([
        "",
        "## Lectura metodológica",
        "La asignación de puestos se basa en cercanía geográfica al templo; las actividades y mesas pueden reasignarse temporalmente para discusión territorial.",
        "La propuesta debe revisarse con liderazgo comunitario, capacidad operativa, rutas, barrios priorizados y conocimiento de los equipos locales.",
    ])
    return "\n".join(lineas)


def generar_informe_templo_markdown(templo, puestos_df, actividades_df, mesas_df, resumen_row=None):
    puestos_t = puestos_df[puestos_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not puestos_df.empty else pd.DataFrame()
    acts_t = actividades_df[actividades_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not actividades_df.empty else pd.DataFrame()
    mesas_t = mesas_df[mesas_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not mesas_df.empty else pd.DataFrame()

    votos_2026 = pd.to_numeric(puestos_t.get("VOTOS_2026", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    votos_2023 = pd.to_numeric(puestos_t.get("VOTOS_2023", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    variacion = votos_2026 - votos_2023
    variacion_pct = variacion / votos_2023 if votos_2023 else np.nan
    prioridad_alta = int(puestos_t.get("PRIORIDAD", pd.Series(dtype=str)).astype(str).str.upper().eq("ALTA").sum()) if not puestos_t.empty else 0
    puestos_caida = puestos_t.assign(_VAR=pd.to_numeric(puestos_t.get("VARIACION_ABSOLUTA", pd.Series(dtype=float)), errors="coerce")).sort_values("_VAR", ascending=True).head(5)
    puestos_crecimiento = puestos_t.assign(_VAR=pd.to_numeric(puestos_t.get("VARIACION_ABSOLUTA", pd.Series(dtype=float)), errors="coerce")).sort_values("_VAR", ascending=False).head(5)
    cobertura_operativa = len(acts_t) + len(mesas_t)

    if pd.notna(variacion_pct) and variacion_pct < -0.10:
        estado = "Critico: requiere plan de recuperación inmediato."
    elif pd.notna(variacion_pct) and variacion_pct < 0:
        estado = "Atención media: contiene pérdida y debe reforzarse."
    else:
        estado = "Fortaleza: sostener presencia y convertir crecimiento en estructura."

    lineas = [
        f"# Informe por templo: {templo}",
        "",
        "## Resumen ejecutivo",
        f"- Estado territorial: {estado}",
        f"- Puestos asignados: {fmt_number(len(puestos_t), 0)}.",
        f"- Votos 2026: {fmt_number(votos_2026, 0)} frente a {fmt_number(votos_2023, 0)} en 2023.",
        f"- Variación: {fmt_number(variacion, 0)} votos ({fmt_pct(variacion_pct)}).",
        f"- Puestos de prioridad alta: {fmt_number(prioridad_alta, 0)}.",
        f"- Actividades y mesas registradas: {fmt_number(cobertura_operativa, 0)} ({fmt_number(len(acts_t), 0)} actividades, {fmt_number(len(mesas_t), 0)} mesas).",
        "",
        "## Puestos a recuperar",
    ]
    if puestos_caida.empty:
        lineas.append("- Sin puestos con información suficiente de caída.")
    else:
        for _, r in puestos_caida.iterrows():
            lineas.append(f"- {r.get('PUESTO')}: {fmt_number(r.get('VARIACION_ABSOLUTA'), 0)} votos; prioridad {r.get('PRIORIDAD', 'N.D.')}.")

    lineas.extend(["", "## Puestos a consolidar"])
    if puestos_crecimiento.empty:
        lineas.append("- Sin puestos con información suficiente de crecimiento.")
    else:
        for _, r in puestos_crecimiento.iterrows():
            lineas.append(f"- {r.get('PUESTO')}: +{fmt_number(r.get('VARIACION_ABSOLUTA'), 0)} votos; mantener estructura territorial.")

    lineas.extend([
        "",
        "## Recomendación operativa",
        "- Concentrar revisión semanal en puestos de prioridad alta y variación negativa.",
        "- Cruzar mesas de trabajo con los puestos de mayor caudal electoral para cerrar brechas de presencia.",
        "- Validar la asignación de templo con liderazgo local antes de convertirla en decisión operativa final.",
    ])
    return "\n".join(lineas)


def generar_informe_ejecutivo_markdown(puestos_df, resumen_iglesia_df, matriz_df, actividades_df, mesas_df):
    def col_serie(df, col, default=""):
        if df is None or df.empty:
            return pd.Series(dtype=object)
        if col in df.columns:
            return df[col]
        return pd.Series(default, index=df.index)

    total_2026 = pd.to_numeric(col_serie(puestos_df, "VOTOS_2026", 0), errors="coerce").sum()
    total_2023 = pd.to_numeric(col_serie(puestos_df, "VOTOS_2023", 0), errors="coerce").sum()
    variacion = total_2026 - total_2023
    variacion_pct = variacion / total_2023 if total_2023 else np.nan
    altas = matriz_df[col_serie(matriz_df, "NIVEL_PRIORIDAD").astype(str).eq("ALTA")].copy() if not matriz_df.empty else pd.DataFrame()
    cambios = puestos_df[col_serie(puestos_df, "CAMBIO_PROPUESTO_TEMPLO").astype(str).eq("SI")].copy() if not puestos_df.empty else pd.DataFrame()
    top_iglesia = resumen_iglesia_df.sort_values("VOTOS_2026", ascending=False).iloc[0] if not resumen_iglesia_df.empty and "VOTOS_2026" in resumen_iglesia_df.columns else None
    mayor_caida = puestos_df.assign(_VAR=pd.to_numeric(col_serie(puestos_df, "VARIACION_ABSOLUTA", 0), errors="coerce")).sort_values("_VAR", ascending=True).head(5)
    mayor_crecimiento = puestos_df.assign(_VAR=pd.to_numeric(col_serie(puestos_df, "VARIACION_ABSOLUTA", 0), errors="coerce")).sort_values("_VAR", ascending=False).head(5)

    lineas = [
        "# Informe ejecutivo territorial-electoral Kennedy",
        "",
        "## 1. Resumen general",
        f"Kennedy registra {fmt_number(total_2026, 0)} votos 2026 frente a {fmt_number(total_2023, 0)} votos 2023, con una variación de {fmt_number(variacion, 0)} votos ({fmt_pct(variacion_pct)}).",
        f"El análisis integra {fmt_number(len(puestos_df), 0)} puestos de votación, {fmt_number(len(actividades_df), 0)} actividades de campaña y {fmt_number(len(mesas_df), 0)} mesas de trabajo.",
        "",
        "## 2. Hallazgos principales",
        f"- Puestos de prioridad alta: {fmt_number(len(altas), 0)}.",
        f"- Puestos con cambio operativo de templo sugerido: {fmt_number(len(cambios), 0)}.",
    ]
    if top_iglesia is not None:
        lineas.append(f"- Mayor concentración electoral 2026: {top_iglesia['IGLESIA']} con {fmt_number(top_iglesia['VOTOS_2026'], 0)} votos.")

    lineas.extend(["", "## 3. Puestos críticos de recuperación"])
    for _, r in mayor_caida.iterrows():
        lineas.append(f"- {r.get('PUESTO')}: {fmt_number(r.get('VARIACION_ABSOLUTA'), 0)} votos; responsable sugerido {r.get('TEMPLO_PROPUESTO', r.get('IGLESIA'))}.")

    lineas.extend(["", "## 4. Puestos de consolidación"])
    for _, r in mayor_crecimiento.iterrows():
        lineas.append(f"- {r.get('PUESTO')}: +{fmt_number(r.get('VARIACION_ABSOLUTA'), 0)} votos; mantener presencia comunitaria y testigos.")

    lineas.extend([
        "",
        "## 5. Recomendaciones estratégicas",
        "- Revisar primero puestos de alta votación con caída y baja presencia territorial.",
        "- Validar políticamente los cambios de templo sugeridos antes de convertirlos en decisión operativa.",
        "- Completar barrio y UPZ para producir una lectura de concentración territorial más fina.",
        "- Usar la asignación territorial como herramienta de discusión, no como decisión automática definitiva.",
        "",
        "## 6. Agenda sugerida 30/60/90",
        "- 0-30 días: revisión de puestos de prioridad alta, responsables y agenda territorial.",
        "- 31-60 días: mesas comunitarias en puestos de recuperación y zonas de oportunidad.",
        "- 61-90 días: consolidación de testigos, líderes y seguimiento por templo.",
    ])
    return "\n".join(lineas)


def get_indicador(df, indicador, default=np.nan):
    if df is None or df.empty:
        return default
    row = df[df["INDICADOR"].astype(str).str.strip().eq(indicador)]
    if row.empty:
        return default
    return row.iloc[0]["VALOR"]


def metric_card(label, value, delta=None, positive=True, icon="📍"):
    delta_html = ""
    if delta is not None:
        cls = "metric-delta-positive" if positive else "metric-delta-negative"
        arrow = "↑" if positive else "↓"
        delta_html = f'<div class="{cls}">{arrow} {delta}</div>'
    st.markdown(
        f"""
        <div class="metric-card">
            <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                <div>
                    <div class="metric-label">{label}</div>
                    <div class="metric-value">{value}</div>
                </div>
                <div style="font-size: 1.5rem; opacity: 0.5;">{icon}</div>
            </div>
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


def crear_mapa(puestos, iglesias, actividades, mesas, map_mode="Vista general", layers_config=None):
    layers_config = layers_config or {}
    m = folium.Map(location=KENNEDY_CENTER, zoom_start=13, tiles="CartoDB positron", control_scale=True)
    Fullscreen(position="topleft").add_to(m)
    MiniMap(toggle_display=True, position="bottomleft").add_to(m)
    show_contorno = layers_config.get("contorno", True)
    show_upz = layers_config.get("upz", False)
    show_heat_default = layers_config.get("heat", map_mode in {"Vista general", "Vista de calor"})
    show_puestos_default = layers_config.get("puestos", map_mode == "Vista electoral")
    show_acts_default = layers_config.get("actividades", map_mode == "Vista operativa")
    show_mesas_default = layers_config.get("mesas", map_mode == "Vista operativa")
    show_templos = layers_config.get("templos", True)
    localidades_gj = cargar_geojson(LOCALIDADES_GEOJSON)
    if show_contorno and localidades_gj:
        agregar_contorno_localidades(m)

    upz_gj = cargar_geojson(UPZ_GEOJSON)
    if show_upz and upz_gj:
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

    if show_heat_default:
        agregar_heatmap_electoral(m, puestos, show=True, name="Rango de calor votos 2026")

    # MIRA Logo
    mira_logo_html = """
    <div style="position: fixed; top: 15px; left: 60px; z-index:9999; background: rgba(255, 255, 255, 0.85); backdrop-filter: blur(4px); padding: 5px 12px; border-radius: 6px; border: 1px solid #E2E8F0; font-family: 'Inter', sans-serif; font-weight: 900; color: #1E3A8A; font-size: 16px; letter-spacing: 1px; font-style: italic; box-shadow: 0 2px 8px rgba(0,0,0,0.05);">
    MIRA
    </div>
    """
    m.get_root().html.add_child(folium.Element(mira_logo_html))

    # Puestos
    puestos_layer = folium.FeatureGroup(name="Puestos de votación fijos", show=show_puestos_default)
    for _, r in puestos.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        iglesia = r.get("IGLESIA", "")
        color_templo = COLORES_TEMPLOS.get(iglesia, "#64748B")
        puesto = safe_html(r.get("PUESTO", ""))
        iglesia = safe_html(r.get("IGLESIA", ""))
        barrio = safe_html(r.get("BARRIO", ""))
        upz = safe_html(r.get("UPZ", ""))
        accion = safe_html(r.get("ACCION_RECOMENDADA", ""))
        votos = float(r.get("VOTOS_2026", 0) or 0)
        radius = max(4, min(8, 4 + math.sqrt(votos) / 4))

        popup = f"""
        <div style="font-family:'Inter', sans-serif; width:340px; color:#0F172A;">
        <h4 style="margin-bottom:12px; font-weight:800; border-bottom: 1px solid #E2E8F0; padding-bottom:8px;">{puesto}</h4>
        <table style="width:100%; border-collapse: collapse; font-size: 12.5px;">
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600; color:#475569;">Iglesia</td><td style="padding:6px 8px; font-weight:700;">{iglesia}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600; color:#475569;">Barrio (UPZ)</td><td style="padding:6px 8px;">{barrio} ({upz})</td></tr>
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600; color:#475569;">Votos 2026</td><td style="padding:6px 8px; font-weight:800; color:#2563EB;">{fmt_number(r.get('VOTOS_2026'),0)}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600; color:#475569;">Variación</td><td style="padding:6px 8px;">{fmt_number(r.get('VARIACION_ABSOLUTA'),0)} ({fmt_pct(r.get('VARIACION_PORCENTUAL'))})</td></tr>
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600; color:#475569;">Actividades</td><td style="padding:6px 8px;">{fmt_number(r.get('ACTIVIDADES_CAMPANA'),0)}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600; color:#475569;">Mesas de trabajo</td><td style="padding:6px 8px;">{fmt_number(r.get('MESAS_TRABAJO_BARRIO'),0)}</td></tr>
            <tr style="background:#FEF2F2;"><td style="padding:6px 8px; font-weight:600; color:#991B1B;">Prioridad</td><td style="padding:6px 8px; font-weight:700; color:#991B1B;">{safe_html(r.get('PRIORIDAD',''))}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600; color:#475569;">Acción</td><td style="padding:6px 8px;">{accion}</td></tr>
        </table>
        </div>
        """
        
        folium.CircleMarker(
            location=[r["LATITUD"], r["LONGITUD"]],
            radius=radius,
            color="#FFFFFF",
            weight=1.2,
            fill=True,
            fill_color=color_templo,
            fill_opacity=0.72,
            popup=folium.Popup(popup, max_width=380),
            tooltip=f"{r.get('PUESTO','')} | {r.get('IGLESIA','')} | {fmt_number(r.get('VOTOS_2026'),0)} votos",
        ).add_to(puestos_layer)
    if show_puestos_default:
        puestos_layer.add_to(m)

    # Churches
    iglesia_layer = folium.FeatureGroup(name="Iglesias / templos", show=True)
    for _, r in iglesias.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        color = COLORES_TEMPLOS.get(r["IGLESIA"], "#334155")
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            tooltip=f"Templo oficial: {r.get('IGLESIA','')}",
            popup=folium.Popup(
                f"<b>{safe_html(r.get('IGLESIA',''))}</b><br>Lat: {r['LATITUD']}<br>Lon: {r['LONGITUD']}",
                max_width=280,
            ),
            icon=crear_icono_div("templo", color, "T"),
        ).add_to(iglesia_layer)

        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            icon=crear_etiqueta_templo(r.get("IGLESIA", ""), color, dx=30, dy=-19),
        ).add_to(iglesia_layer)
    if show_templos:
        iglesia_layer.add_to(m)

    # Activities
    acts_layer = folium.FeatureGroup(name="Actividades de campaña", show=show_acts_default)
    for _, r in actividades.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        popup_actividad = f"""
        <div style="font-family:'Inter', sans-serif; width:280px; color:#0F172A;">
        <h4 style="margin-bottom:12px; font-weight:800; border-bottom: 1px solid #E2E8F0; padding-bottom:8px;">Actividad: {safe_html(r.get('TIPO_ACTIVIDAD',''))}</h4>
        <table style="width:100%; border-collapse: collapse; font-size: 12px;">
            <tr style="background:#F8FAFC;"><td style="padding:4px 8px; font-weight:600;">Templo asignado</td><td style="padding:4px 8px;">{safe_html(r.get('IGLESIA',''))}</td></tr>
            <tr><td style="padding:4px 8px; font-weight:600;">Barrio</td><td style="padding:4px 8px;">{safe_html(r.get('BARRIO',''))}</td></tr>
            <tr style="background:#F8FAFC;"><td style="padding:4px 8px; font-weight:600;">Líder</td><td style="padding:4px 8px;">{safe_html(r.get('LIDER',''))}</td></tr>
            <tr><td style="padding:4px 8px; font-weight:600;">Dirección</td><td style="padding:4px 8px;">{safe_html(r.get('DIRECCION',''))}</td></tr>
        </table>
        </div>
        """
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            tooltip=f"{r.get('TIPO_ACTIVIDAD','')} | {r.get('IGLESIA','')}",
            popup=folium.Popup(popup_actividad, max_width=340),
            icon=crear_icono_div("actividad", "#2563EB", "A"),
        ).add_to(acts_layer)
    if show_acts_default:
        acts_layer.add_to(m)

    mesas_layer = folium.FeatureGroup(name="Mesas de trabajo", show=show_mesas_default)
    for _, r in mesas.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        popup_mesa = f"""
        <div style="font-family:'Inter', sans-serif; width:300px; color:#0F172A;">
        <h4 style="margin-bottom:12px; font-weight:800; border-bottom: 1px solid #E2E8F0; padding-bottom:8px;">Mesa: {safe_html(r.get('TEMA',''))}</h4>
        <table style="width:100%; border-collapse: collapse; font-size: 12.5px;">
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600;">Templo asignado</td><td style="padding:6px 8px; font-weight:700;">{safe_html(r.get('IGLESIA',''))}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600;">Barrio</td><td style="padding:6px 8px;">{safe_html(r.get('BARRIO',''))}</td></tr>
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600;">Líder</td><td style="padding:6px 8px;">{safe_html(r.get('LIDER',''))}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600;">Concejal</td><td style="padding:6px 8px;">{safe_html(r.get('CONCEJAL',''))}</td></tr>
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600;">Beneficiarios</td><td style="padding:6px 8px; font-weight:800; color:#2563EB;">{fmt_number(r.get('BENEFICIARIOS'),0)}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600;">Estado</td><td style="padding:6px 8px; font-weight:700; color:#991B1B;">{safe_html(r.get('ESTADO',''))}</td></tr>
        </table>
        </div>
        """
        folium.Marker(
            location=[r["LATITUD"], r["LONGITUD"]],
            tooltip=f"Mesa | {r.get('IGLESIA','')} | {r.get('BARRIO','')}",
            popup=folium.Popup(popup_mesa, max_width=360),
            icon=crear_icono_div("mesa", "#F97316", "M"),
        ).add_to(mesas_layer)
    if show_mesas_default:
        mesas_layer.add_to(m)

    legend_items = "".join(
        f'''
        <div style="display:flex;align-items:center;gap:7px;margin:4px 0;">
            <span style="width:10px;height:10px;border-radius:50%;background:{color};display:inline-block;"></span>
            <span>Puesto ({templo})</span>
        </div>
        '''
        for templo, color in COLORES_TEMPLOS.items()
    )

    legend_html = f'''
    <div style="
        position: fixed;
        bottom: 35px;
        right: 35px;
        z-index:9999;
        background:white;
        padding:10px 12px;
        border:1px solid #CBD5E1;
        border-radius:10px;
        box-shadow:0 4px 14px rgba(15,23,42,.16);
        font-size:11px;
        color:#0F172A;
        min-width:190px;
        max-width:190px;
        font-family: 'Inter', sans-serif;
    ">
        <div style="font-weight:900;margin-bottom:6px;">Lectura del mapa</div>
        {legend_items}
        <div style="height:1px;background:#E2E8F0;margin:6px 0;"></div>
        <div style="display:flex;align-items:center;gap:7px;margin:3px 0;">
            <span style="width:14px;height:14px;background:#0F172A;border:3px solid white;border-radius:50% 50% 50% 0;transform:rotate(-45deg);display:inline-block;box-shadow:0 0 0 1px #0F172A;"></span>
            <span>Templo oficial</span>
        </div>
        <div style="display:flex;align-items:center;gap:7px;margin:3px 0;">
            <span style="width:16px;height:16px;border-radius:5px;background:#2563EB;color:white;font-size:9px;font-weight:900;display:inline-flex;align-items:center;justify-content:center;border:2px solid white;box-shadow:0 0 0 1px #2563EB;">A</span>
            <span>Actividad de campaña</span>
        </div>
        <div style="display:flex;align-items:center;gap:7px;margin:3px 0;">
            <span style="width:16px;height:16px;border-radius:50%;background:#F97316;color:white;font-size:9px;font-weight:900;display:inline-flex;align-items:center;justify-content:center;border:2px solid white;box-shadow:0 0 0 1px #F97316;">M</span>
            <span>Mesa de trabajo</span>
        </div>
        <div style="display:flex;align-items:center;gap:7px;margin:3px 0;">
            <span style="width:10px;height:10px;background:#94A3B8;opacity:.5;display:inline-block;"></span>
            <span>Otras localidades</span>
        </div>
        <div style="height:1px;background:#E2E8F0;margin:6px 0;"></div>
    </div>
    '''

    m.get_root().html.add_child(folium.Element(legend_html))
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
actividades = aplicar_ajustes_templo(actividades, "ajustes_actividades", "ACTIVIDAD_ID")
mesas = aplicar_ajustes_templo(mesas, "ajustes_mesas", "MESA_ID")
puestos = aplicar_ajustes_templo(puestos, "ajustes_asignacion", "PUESTO")
if not matriz.empty:
    matriz = aplicar_ajustes_templo(matriz, "ajustes_asignacion", "PUESTO")

if (st.session_state.get("ajustes_asignacion") or st.session_state.get("ajustes_actividades") or st.session_state.get("ajustes_mesas")) and not resumen_iglesia.empty:
    agregado_puestos = puestos.groupby("IGLESIA").agg(
        VOTOS_2026=("VOTOS_2026", "sum"),
        VOTOS_2023=("VOTOS_2023", "sum"),
        PUESTOS=("PUESTO", "nunique")
    ).reset_index()
    agregado_puestos["VARIACION_ABSOLUTA"] = agregado_puestos["VOTOS_2026"] - agregado_puestos["VOTOS_2023"]
    agregado_puestos["VARIACION_PORCENTUAL"] = np.where(agregado_puestos["VOTOS_2023"] > 0, agregado_puestos["VARIACION_ABSOLUTA"] / agregado_puestos["VOTOS_2023"], 0)
    
    if not actividades.empty and "IGLESIA" in actividades.columns:
        agregado_acts = actividades.groupby("IGLESIA").size().reset_index(name="ACTIVIDADES_CAMPANA")
    else:
        agregado_acts = pd.DataFrame(columns=["IGLESIA", "ACTIVIDADES_CAMPANA"])
        
    if not mesas.empty and "IGLESIA" in mesas.columns:
        agregado_mesas = mesas.groupby("IGLESIA").size().reset_index(name="MESAS_TRABAJO")
    else:
        agregado_mesas = pd.DataFrame(columns=["IGLESIA", "MESAS_TRABAJO"])
        
    cols_to_update = ["VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "PUESTOS", "ACTIVIDADES_CAMPANA", "MESAS_TRABAJO"]
    cols_to_drop = [c for c in cols_to_update if c in resumen_iglesia.columns]
    resumen_iglesia = resumen_iglesia.drop(columns=cols_to_drop)
    
    resumen_iglesia = resumen_iglesia.merge(agregado_puestos, on="IGLESIA", how="left")
    resumen_iglesia = resumen_iglesia.merge(agregado_acts, on="IGLESIA", how="left")
    resumen_iglesia = resumen_iglesia.merge(agregado_mesas, on="IGLESIA", how="left")
    
    for col in cols_to_update:
        if col in resumen_iglesia.columns:
            resumen_iglesia[col] = resumen_iglesia[col].fillna(0)

asignacion = calcular_distancias_a_templos_v2(puestos, iglesias)


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
            "MESAS_TRABAJO", "PUESTOS", "PUNTAJE_PRIORIDAD",
        ]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").round().astype("Int64")


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("Configuración del análisis")
    st.markdown("Fuente única: `kennedy_mira_consolidado.xlsx`")
    st.caption(f"Base de cambios: {persistence_backend_label()}")
    if st.session_state.get("google_sheets_error"):
        st.warning("Google Sheets está configurado, pero no conectado. Revise credenciales o permisos del archivo.")

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
puestos_alta = get_indicador(resumen_general, "Puestos prioridad alta", default=np.nan)
if pd.isna(puestos_alta) and "PRIORIDAD" in puestos.columns:
    puestos_alta = int(puestos["PRIORIDAD"].astype(str).eq("ALTA").sum())
puestos_cambio_templo = get_indicador(resumen_general, "Puestos con cambio de templo sugerido", default=np.nan)
if pd.isna(puestos_cambio_templo) and "CAMBIO_PROPUESTO_TEMPLO" in puestos.columns:
    puestos_cambio_templo = int(puestos["CAMBIO_PROPUESTO_TEMPLO"].astype(str).eq("SI").sum())

c1, c2, c3, c4 = st.columns(4)

with c1:
    metric_card("Votos Kennedy 2026", fmt_number(total_2026, 0))

with c2:
    metric_card("Variación vs 2023", fmt_number(var_abs, 0), fmt_number(abs(var_abs), 0), positive=var_abs >= 0)

with c3:
    metric_card("Puestos analizados", fmt_number(puestos_total, 0))

with c4:
    metric_card("Puestos prioridad alta", fmt_number(puestos_alta, 0))

with st.expander("Ver indicadores complementarios", expanded=False):
    e1, e2, e3, e4 = st.columns(4)
    with e1:
        metric_card("Total Kennedy 2023", fmt_number(total_2023, 0))
    with e2:
        metric_card("Variación porcentual", fmt_pct(var_pct), fmt_pct(abs(var_pct)), positive=var_pct >= 0)
    with e3:
        metric_card("Iglesias oficiales", fmt_number(iglesias_total, 0))
    with e4:
        metric_card("Cambios sugeridos", fmt_number(puestos_cambio_templo, 0))

    e5, e6, e7, e8, e9 = st.columns(5)
    with e5:
        metric_card("Actividades oficiales", fmt_number(actividades_oficiales_total, 0))
    with e6:
        metric_card("Volanteos confirmados", fmt_number(volanteos_total, 0))
    with e7:
        metric_card("Mesas de trabajo", fmt_number(mesas_total, 0))
    with e8:
        metric_card("JAL / Concejo 2023", f"{fmt_number(jal_total, 0)} / {fmt_number(concejo_total, 0)}")
    with e9:
        metric_card("Cámara / Senado 2026", f"{fmt_number(camara_total, 0)} / {fmt_number(senado_total, 0)}")

st.markdown(
    f"""
    <div class="summary-ribbon">
    <b>Lectura ejecutiva:</b> Kennedy registra <b>{fmt_number(total_2026, 0)}</b> votos para 2026,
    con una variación de <b>{fmt_number(var_abs, 0)}</b> frente a 2023.
    El tablero permite priorizar puestos, revisar presencia territorial y ajustar la asignación operativa
    por templo para la estrategia electoral.
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
    informe_ejecutivo_md = generar_informe_ejecutivo_markdown(puestos, resumen_iglesia, matriz, actividades, mesas)

    foco_cols = [
        "NIVEL_PRIORIDAD", "PUNTAJE_PRIORIDAD", "ROL_ANALITICO", "PUESTO", "IGLESIA_HISTORICA_2026",
        "TEMPLO_PROPUESTO", "VOTOS_2026", "VARIACION_ABSOLUTA", "FACTORES_PRIORIDAD", "ACCION_RECOMENDADA",
    ]
    foco_cols = [c for c in foco_cols if c in matriz.columns]
    focos = matriz[foco_cols].sort_values(
        [c for c in ["NIVEL_PRIORIDAD", "PUNTAJE_PRIORIDAD", "VOTOS_2026"] if c in matriz.columns],
        ascending=[True, False, False][: len([c for c in ["NIVEL_PRIORIDAD", "PUNTAJE_PRIORIDAD", "VOTOS_2026"] if c in matriz.columns])],
    ).head(12)

    st.markdown("### Panel de decisión")
    dcol1, dcol2 = st.columns([2, 1])
    with dcol1:
        st.dataframe(focos, hide_index=True, width="stretch")
        st.download_button("Descargar panel en Excel", to_excel_bytes(focos, "Panel Decision"), "panel_decision.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_panel")
    with dcol2:
        st.markdown(
            """
            <div class="section-card">
            <b>Uso recomendado</b><br>
            1. Revisar prioridad alta.<br>
            2. Validar templo propuesto con liderazgo local.<br>
            3. Programar acción territorial y responsable.<br>
            4. Exportar informe para seguimiento.
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.download_button(
            "Descargar informe ejecutivo",
            informe_ejecutivo_md.encode("utf-8"),
            "informe_ejecutivo_kennedy.md",
            "text/markdown",
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
        st.plotly_chart(fig, width="stretch")
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
        st.plotly_chart(fig, width="stretch")

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
            st.dataframe(acts_counts, width="stretch", hide_index=True)
        with ca2:
            st.plotly_chart(fig_acts, width="stretch")
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
        st.dataframe(elec_data[elec_data["Bloque de análisis"].eq("JAL / Concejo 2023")], hide_index=True, width="stretch")
        st.markdown("**Cámara / Senado 2026**")
        st.dataframe(elec_data[elec_data["Bloque de análisis"].eq("Cámara / Senado 2026")], hide_index=True, width="stretch")
        st.markdown("**Variables operativas del nuevo reporte**")
        st.dataframe(
            pd.DataFrame(
                {
                    "Indicador": ["Mesas 2026", "Testigos 2023"],
                    "Valor": [mesas_2026_reporte, testigos_2023_reporte],
                }
            ),
            hide_index=True,
            width="stretch",
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
        st.plotly_chart(fig_elec, width="stretch")

with tab_mapa:
    st.subheader("Mapa interactivo territorial")
    st.markdown(
        '<div class="note-box">Este mapa permite revisar presencia territorial, mesas de trabajo y actividades de campaña por templo. Los ajustes de templo quedan guardados en base de datos con historial de cambios y no alteran el Excel maestro original.</div>',
        unsafe_allow_html=True,
    )
    map_mode = st.radio(
        "Modo de vista",
        ["Vista general", "Vista electoral", "Vista operativa", "Vista de calor"],
        horizontal=True,
    )
    opciones_templo = ["Todos los templos"] + list(IGLESIAS_OFICIALES_PERMITIDAS)
    filtro_templo = st.selectbox("Filtrar vista del mapa por templo", opciones_templo, key="filtro_mapa_templo")
    ventana_tiempo = st.selectbox(
        "Ventana temporal de actividades/mesas",
        ["Todo el histórico", "Últimos 30 días", "Últimos 60 días", "Últimos 90 días"],
        index=0,
    )

    puestos_mapa = puestos_f.copy()
    acts_mapa = actividades_f.copy()
    mesas_mapa = mesas_f.copy()
    iglesias_mapa = iglesias.copy()

    def _filtrar_por_ventana_temporal(df, ventana):
        if df is None or df.empty or ventana == "Todo el histórico":
            return df
        candidatos = [c for c in df.columns if "FECHA" in str(c).upper()]
        if not candidatos:
            return df
        fecha_col = candidatos[0]
        dias = int(ventana.split()[1])
        corte = pd.Timestamp.now() - pd.Timedelta(days=dias)
        fechas = pd.to_datetime(df[fecha_col], errors="coerce")
        mask = fechas.notna() & fechas.ge(corte)
        return df[mask].copy()

    acts_mapa = _filtrar_por_ventana_temporal(acts_mapa, ventana_tiempo)
    mesas_mapa = _filtrar_por_ventana_temporal(mesas_mapa, ventana_tiempo)

    if filtro_templo != "Todos los templos":
        puestos_mapa = puestos_mapa[puestos_mapa["IGLESIA"].eq(filtro_templo)]
        if "IGLESIA" in acts_mapa.columns:
            acts_mapa = acts_mapa[acts_mapa["IGLESIA"].eq(filtro_templo)]
        if "IGLESIA" in mesas_mapa.columns:
            mesas_mapa = mesas_mapa[mesas_mapa["IGLESIA"].eq(filtro_templo)]
        iglesias_mapa = iglesias_mapa[iglesias_mapa["IGLESIA"].eq(filtro_templo)]

    resumen_operativo_mapa = crear_resumen_operativo_por_templo(acts_mapa, mesas_mapa)
    map_a1, map_a2, map_a3, map_a4 = st.columns(4)
    with map_a1:
        metric_card("Puestos visibles", fmt_number(len(puestos_mapa), 0), icon="📍")
    with map_a2:
        metric_card("Actividades visibles", fmt_number(len(acts_mapa), 0), icon="📅")
    with map_a3:
        metric_card("Mesas visibles", fmt_number(len(mesas_mapa), 0), icon="👥")
    with map_a4:
        ajustes_operativos = len(st.session_state.get("ajustes_actividades", {})) + len(st.session_state.get("ajustes_mesas", {}))
        metric_card("Ajustes operativos", fmt_number(ajustes_operativos, 0), icon="⚙️")

    puestos_sin_coord = int(puestos_mapa[["LATITUD", "LONGITUD"]].isna().any(axis=1).sum()) if not puestos_mapa.empty else 0
    acts_sin_coord = int(acts_mapa[["LATITUD", "LONGITUD"]].isna().any(axis=1).sum()) if not acts_mapa.empty else 0
    mesas_sin_coord = int(mesas_mapa[["LATITUD", "LONGITUD"]].isna().any(axis=1).sum()) if not mesas_mapa.empty else 0

    q1, q2, q3 = st.columns(3)
    with q1:
        metric_card("Puestos sin coordenadas", fmt_number(puestos_sin_coord, 0), icon="🧭")
    with q2:
        metric_card("Actividades sin coordenadas", fmt_number(acts_sin_coord, 0), icon="📌")
    with q3:
        metric_card("Mesas sin coordenadas", fmt_number(mesas_sin_coord, 0), icon="📍")

    with st.expander("Capas del mapa territorial", expanded=False):
        layer_cols = st.columns(6)
        with layer_cols[0]:
            layer_contorno = st.checkbox("Contorno", value=True, key=f"territorial_contorno_{map_mode}")
        with layer_cols[1]:
            layer_templos = st.checkbox("Templos", value=True, key=f"territorial_templos_{map_mode}")
        with layer_cols[2]:
            layer_heat = st.checkbox("Rango electoral", value=map_mode in {"Vista general", "Vista de calor"}, key=f"territorial_heat_{map_mode}")
        with layer_cols[3]:
            layer_puestos = st.checkbox("Puestos", value=map_mode == "Vista electoral", key=f"territorial_puestos_{map_mode}")
        with layer_cols[4]:
            layer_actividades = st.checkbox("Actividades", value=map_mode == "Vista operativa", key=f"territorial_actividades_{map_mode}")
        with layer_cols[5]:
            layer_mesas = st.checkbox("Mesas", value=map_mode == "Vista operativa", key=f"territorial_mesas_{map_mode}")

    territorial_layers = {
        "contorno": layer_contorno,
        "templos": layer_templos,
        "heat": layer_heat,
        "puestos": layer_puestos,
        "actividades": layer_actividades,
        "mesas": layer_mesas,
    }
    mapa = crear_mapa(puestos_mapa, iglesias_mapa, acts_mapa, mesas_mapa, map_mode=map_mode, layers_config=territorial_layers)
    st.markdown("### Mapa territorial")
    st.markdown("<div style='font-size:14px; color:#475569; margin-bottom:10px;'>💡 <b>Vista inicial limpia:</b> el modo de vista controla si se muestran votos, calor, actividades o mesas. Las convenciones del mapa quedan integradas abajo.</div>", unsafe_allow_html=True)
    map_key = f"mapa_territorial_v3_{map_mode}_{filtro_templo}_{ventana_tiempo}_{layer_contorno}_{layer_templos}_{layer_heat}_{layer_puestos}_{layer_actividades}_{layer_mesas}".replace(" ", "_").replace("/", "_")
    with st.container(border=True):
        render_folium_map(mapa, height=790, key=map_key)

    with st.expander("Ajustar templo de una mesa de trabajo", expanded=False):
        st.caption("Ajuste definitivo para modificar la asignación. No modifica el Excel maestro directamente pero sí los reportes exportables.")
        if mesas.empty:
            st.info("No hay mesas disponibles para ajustar.")
        else:
            mesa_tmp = mesas.copy()
            mesa_tmp["LABEL"] = mesa_tmp.apply(
                lambda r: f"{int(r['MESA_ID'])} · {r.get('NOMBRE_GESTION', r.get('TEMA',''))} · {r.get('BARRIO','SIN BARRIO')} · {r.get('IGLESIA','')}",
                axis=1,
            )
            mesa_label = st.selectbox("Mesa de trabajo", mesa_tmp["LABEL"].tolist(), key="sel_mesa_ajuste_compacto")
            mesa_row = mesa_tmp[mesa_tmp["LABEL"].eq(mesa_label)].iloc[0]
            mesa_index = TEMPLOS_OFICIALES.index(mesa_row["IGLESIA"]) if mesa_row.get("IGLESIA") in TEMPLOS_OFICIALES else 0

            m1, m2 = st.columns([2, 1])
            with m1:
                st.dataframe(
                    pd.DataFrame(
                        [
                            ("Tema", mesa_row.get("TEMA")),
                            ("Barrio", mesa_row.get("BARRIO")),
                            ("Templo original", mesa_row.get("IGLESIA_ORIGINAL")),
                            ("Templo actual", mesa_row.get("IGLESIA")),
                            ("Líder", mesa_row.get("LIDER")),
                            ("Estado", mesa_row.get("ESTADO")),
                        ],
                        columns=["Campo", "Valor"],
                    ),
                    hide_index=True,
                    width="stretch",
                )
            with m2:
                mesa_templo = st.selectbox("Templo asignado", TEMPLOS_OFICIALES, index=mesa_index, key="templo_mesa_ajuste_compacto")
                if st.button("Guardar ajuste de mesa"):
                    st.session_state.setdefault("ajustes_mesas", {})
                    st.session_state.setdefault("ajustes_mesas", {})[mesa_row["MESA_ID"]] = mesa_templo
                    registrar_ajuste_en_db(
                        session_key="ajustes_mesas",
                        entity_id=mesa_row["MESA_ID"],
                        nombre_entidad=mesa_row.get("NOMBRE_GESTION", mesa_row.get("TEMA", "")),
                        templo_nuevo=mesa_templo,
                        motivo="Ajuste manual desde pestaña Mapa territorial",
                    )
                    guardar_ajustes_guardados()
                    st.success("Ajuste de mesa guardado.")
                    st.rerun()
                if st.button("Limpiar ajustes de mesas"):
                    total_limpiados = limpiar_ajustes_en_db("ajustes_mesas", motivo="Limpieza manual desde pestaña Mapa territorial")
                    st.session_state["ajustes_mesas"] = {}
                    guardar_ajustes_guardados()
                    st.info(f"Se limpiaron {fmt_number(total_limpiados, 0)} ajuste(s) de mesas en la base.")
                    st.rerun()

    st.markdown("### Resumen operativo por templo")
    st.dataframe(resumen_operativo_mapa, hide_index=True, width="stretch")
    st.download_button("Descargar tabla en Excel", to_excel_bytes(resumen_operativo_mapa, "Resumen Operativo"), "resumen_operativo.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_res_op")

    asignacion_reporte_mapa = aplicar_ajustes_asignacion(asignacion.copy())
    informe_mapa = generar_informe_territorial(asignacion_reporte_mapa, actividades, mesas)
    with st.expander("Informe territorial automático", expanded=True):
        st.markdown(informe_mapa)
        st.download_button(
            "Descargar informe territorial",
            informe_mapa.encode("utf-8"),
            "informe_territorial_operativo.md",
            "text/markdown",
        )

with tab_asignacion:
    st.subheader("Asignación territorial de puestos")
    st.markdown(
        """
        <div class="section-card">
        <b>Regla de interpretación.</b><br>
        La asignación vigente no es una recomendación automática del sistema. Es la última decisión territorial
        guardada por el equipo. El templo más cercano y la distancia son insumos técnicos para la discusión,
        pero no modifican la asignación por sí solos.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="note-box">
        La asignación parte del templo original. La cercanía y las líneas puesto-templo son criterios técnicos
        de referencia. Los cambios solo aplican cuando se guardan manualmente y quedan persistidos en la base de datos.
        </div>
        """,
        unsafe_allow_html=True,
    )

    asignacion_base = asignacion.copy()
    asignacion_final = aplicar_ajustes_asignacion(asignacion_base)
    ajustes_puestos = st.session_state.get("ajustes_asignacion", {})
    actuales_puestos_db = obtener_ajustes_actuales_df("puesto")
    historial_puestos_db = obtener_historial_ajustes(limit=5000)
    if not historial_puestos_db.empty:
        historial_puestos_db = historial_puestos_db[historial_puestos_db["entidad"].eq("puesto")].copy()

    asignacion_vista = asignacion_final.copy()
    asignacion_vista["ES_AJUSTADO"] = asignacion_vista["PUESTO"].isin(ajustes_puestos.keys())
    asignacion_vista["ESTADO_ASIGNACION"] = np.where(asignacion_vista["ES_AJUSTADO"], "Ajustado manualmente", "Original")
    if "MESAS_TRABAJO_BARRIO" in puestos.columns:
        mesas_por_puesto = puestos[["PUESTO", "MESAS_TRABAJO_BARRIO"]].drop_duplicates("PUESTO").copy()
        asignacion_vista = asignacion_vista.merge(mesas_por_puesto, on="PUESTO", how="left")
    else:
        asignacion_vista["MESAS_TRABAJO_BARRIO"] = 0

    votos_num = pd.to_numeric(asignacion_vista.get("VOTOS_2026", pd.Series(dtype=float)), errors="coerce").fillna(0)
    var_num = pd.to_numeric(asignacion_vista.get("VARIACION_ABSOLUTA", pd.Series(dtype=float)), errors="coerce").fillna(0)
    mesas_num = pd.to_numeric(asignacion_vista.get("MESAS_TRABAJO_BARRIO", pd.Series(dtype=float)), errors="coerce").fillna(0)
    dist_num = pd.to_numeric(asignacion_vista.get("DISTANCIA_ASIGNADA_KM", pd.Series(dtype=float)), errors="coerce")
    umbral_votos = votos_num.quantile(0.70) if not votos_num.empty else 0
    asignacion_vista["MESA_TRABAJO"] = np.where(mesas_num.gt(0), "SI", "NO")
    asignacion_vista["NIVEL_RIESGO"] = np.where(
        (votos_num >= umbral_votos) & (var_num < 0) & (mesas_num <= 0),
        "ROJA",
        np.where((var_num < 0) & (mesas_num <= 0), "AMARILLA", "VERDE"),
    )
    asignacion_vista["RECOMENDACION"] = np.select(
        [
            asignacion_vista["NIVEL_RIESGO"].eq("ROJA"),
            asignacion_vista["NIVEL_RIESGO"].eq("AMARILLA"),
            dist_num.gt(3),
        ],
        [
            "Priorizar recuperación territorial y mesa de trabajo.",
            "Revisar presencia y agenda comunitaria.",
            "Validar logística por distancia al templo.",
        ],
        default="Mantener seguimiento regular.",
    )

    st.markdown("### Indicadores principales")
    col_total, col_visibles, col_ajustes, col_criticos, col_valladolid = st.columns(5)

    st.markdown("### Filtros de análisis")
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        filtro_asignacion_templo = st.selectbox("Templo vigente", ["Todos los templos"] + TEMPLOS_OFICIALES, key="filtro_asignacion_templo")
    with f2:
        prioridades_asignacion = sorted(asignacion_vista.get("PRIORIDAD", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        filtro_prioridad_asignacion = st.selectbox("Prioridad", ["Todas"] + prioridades_asignacion, key="filtro_prioridad_asignacion")
    with f3:
        filtro_estado_asignacion = st.selectbox("Estado de asignación", ["Todos", "Original", "Ajustado manualmente"], key="filtro_estado_asignacion")
    with f4:
        filtro_criticos_asignacion = st.selectbox("Puestos críticos", ["Todos", "Solo críticos", "Excluir críticos"], key="filtro_criticos_asignacion")

    asignacion_filtrada = asignacion_vista.copy()
    if filtro_asignacion_templo != "Todos los templos":
        asignacion_filtrada = asignacion_filtrada[asignacion_filtrada["TEMPLO_ASIGNADO_FINAL"].eq(filtro_asignacion_templo)].copy()
    if filtro_prioridad_asignacion != "Todas":
        asignacion_filtrada = asignacion_filtrada[asignacion_filtrada.get("PRIORIDAD", pd.Series(dtype=str)).astype(str).eq(filtro_prioridad_asignacion)].copy()
    if filtro_estado_asignacion != "Todos":
        asignacion_filtrada = asignacion_filtrada[asignacion_filtrada["ESTADO_ASIGNACION"].eq(filtro_estado_asignacion)].copy()
    if filtro_criticos_asignacion == "Solo críticos":
        asignacion_filtrada = asignacion_filtrada[asignacion_filtrada["NIVEL_RIESGO"].eq("ROJA")].copy()
    elif filtro_criticos_asignacion == "Excluir críticos":
        asignacion_filtrada = asignacion_filtrada[~asignacion_filtrada["NIVEL_RIESGO"].eq("ROJA")].copy()

    puestos_visibles_asig = len(asignacion_filtrada)
    puestos_criticos_visibles = int(asignacion_filtrada["NIVEL_RIESGO"].eq("ROJA").sum()) if not asignacion_filtrada.empty else 0
    asignados_valladolid = int(asignacion_vista["TEMPLO_ASIGNADO_FINAL"].eq("VALLADOLID").sum()) if "TEMPLO_ASIGNADO_FINAL" in asignacion_vista.columns else 0
    with col_total:
        metric_card("Total puestos", fmt_number(len(asignacion_vista), 0))
    with col_visibles:
        metric_card("Puestos visibles", fmt_number(puestos_visibles_asig, 0))
    with col_ajustes:
        metric_card("Ajustes guardados", fmt_number(len(ajustes_puestos), 0))
    with col_criticos:
        metric_card("Puestos críticos", fmt_number(puestos_criticos_visibles, 0))
    with col_valladolid:
        metric_card("Asignados a Valladolid", fmt_number(asignados_valladolid, 0))

    st.markdown("### Mapa de asignación de puestos de votación")
    st.caption("Cada punto conserva el color del templo vigente; las líneas muestran la relación puesto-templo. Las convenciones ejecutivas del mapa aparecen integradas abajo.")
    with st.expander("Capas del mapa de asignación", expanded=False):
        asig_layer_cols = st.columns(5)
        with asig_layer_cols[0]:
            asig_layer_contorno = st.checkbox("Contorno", value=True, key="asignacion_layer_contorno")
        with asig_layer_cols[1]:
            asig_layer_heat = st.checkbox("Rango electoral", value=True, key="asignacion_layer_heat")
        with asig_layer_cols[2]:
            asig_layer_lineas = st.checkbox("Líneas", value=True, key="asignacion_layer_lineas")
        with asig_layer_cols[3]:
            asig_layer_puestos = st.checkbox("Puestos", value=True, key="asignacion_layer_puestos")
        with asig_layer_cols[4]:
            asig_layer_templos = st.checkbox("Templos", value=True, key="asignacion_layer_templos")

    asignacion_layers = {
        "contorno": asig_layer_contorno,
        "heat": asig_layer_heat,
        "lineas": asig_layer_lineas,
        "puestos": asig_layer_puestos,
        "templos": asig_layer_templos,
    }
    mapa_asignacion = crear_mapa_asignacion(asignacion_filtrada, iglesias, layers_config=asignacion_layers)
    asig_map_key = f"mapa_asignacion_v3_{filtro_asignacion_templo}_{filtro_prioridad_asignacion}_{filtro_estado_asignacion}_{filtro_criticos_asignacion}_{asig_layer_contorno}_{asig_layer_heat}_{asig_layer_lineas}_{asig_layer_puestos}_{asig_layer_templos}".replace(" ", "_").replace("/", "_")
    with st.container(border=True):
        render_folium_map(mapa_asignacion, height=790, key=asig_map_key)

    st.markdown("### Semáforo territorial")
    semaforo_cols = [
        "PUESTO", "TEMPLO_ASIGNADO_FINAL", "IGLESIA_ACTUAL", "VOTOS_2026", "VARIACION_ABSOLUTA",
        "MESA_TRABAJO", "DISTANCIA_ASIGNADA_KM", "NIVEL_RIESGO", "PRIORIDAD", "RECOMENDACION",
    ]
    semaforo_cols = [c for c in semaforo_cols if c in asignacion_filtrada.columns]
    semaforo_df = asignacion_filtrada[semaforo_cols].copy()
    semaforo_df = semaforo_df.rename(
        columns={
            "TEMPLO_ASIGNADO_FINAL": "TEMPLO VIGENTE",
            "IGLESIA_ACTUAL": "TEMPLO ORIGINAL",
            "VOTOS_2026": "VOTOS 2026",
            "VARIACION_ABSOLUTA": "VARIACION",
            "DISTANCIA_ASIGNADA_KM": "DISTANCIA AL TEMPLO",
            "NIVEL_RIESGO": "RIESGO",
        }
    )
    if "RIESGO" in semaforo_df.columns:
        orden_riesgo = {"ROJA": 0, "AMARILLA": 1, "VERDE": 2}
        semaforo_df["_ORDEN"] = semaforo_df["RIESGO"].map(orden_riesgo).fillna(3)
        semaforo_df = semaforo_df.sort_values(["_ORDEN", "VOTOS 2026"], ascending=[True, False]).drop(columns=["_ORDEN"])
    st.dataframe(semaforo_df, hide_index=True, width="stretch")

    st.markdown("### Resumen de decisiones")
    cambios_df = asignacion_vista[asignacion_vista["ES_AJUSTADO"]].copy()
    if not cambios_df.empty:
        cambios_df = cambios_df.merge(
            actuales_puestos_db.add_prefix("DB_"),
            left_on="PUESTO",
            right_on="DB_entidad_id",
            how="left",
        )
        cambios_show = pd.DataFrame(
            {
                "PUESTO": cambios_df["PUESTO"],
                "TEMPLO ORIGINAL": cambios_df.get("IGLESIA_ACTUAL", pd.Series("", index=cambios_df.index)),
                "TEMPLO VIGENTE": cambios_df["TEMPLO_ASIGNADO_FINAL"],
                "USUARIO": cambios_df.get("DB_usuario", pd.Series("", index=cambios_df.index)),
                "FECHA CAMBIO": cambios_df.get("DB_actualizado_en", pd.Series("", index=cambios_df.index)),
                "NOTA": cambios_df.get("DB_motivo", pd.Series("", index=cambios_df.index)),
                "VOTOS 2026": cambios_df.get("VOTOS_2026", pd.Series(dtype=float)),
            }
        )
    else:
        cambios_show = pd.DataFrame(columns=["PUESTO", "TEMPLO ORIGINAL", "TEMPLO VIGENTE", "USUARIO", "FECHA CAMBIO", "NOTA", "VOTOS 2026"])
    st.dataframe(cambios_show, hide_index=True, width="stretch")

    resumen_documental = crear_resumen_asignacion_por_columna(asignacion_vista, "IGLESIA_ACTUAL")
    resumen_final = crear_resumen_asignacion(asignacion_vista)
    impacto = resumen_final[["TEMPLO", "PUESTOS_ASIGNADOS", "VOTOS_2026_ASIGNADOS"]].merge(
        resumen_documental[["TEMPLO", "PUESTOS", "VOTOS_2026"]],
        on="TEMPLO",
        how="left",
        suffixes=("_VIGENTE", "_ORIGINAL"),
    )
    impacto["DELTA_PUESTOS"] = impacto["PUESTOS_ASIGNADOS"] - impacto["PUESTOS"].fillna(0)
    impacto["DELTA_VOTOS_2026"] = impacto["VOTOS_2026_ASIGNADOS"] - impacto["VOTOS_2026"].fillna(0)
    st.markdown("#### Impacto por templo")
    st.dataframe(impacto, hide_index=True, width="stretch")

    with st.expander("Ajustar templo de un puesto de votación", expanded=False):
        st.caption("El ajuste manual se registra como decisión territorial vigente y queda en historial.")
        if "ajustes_asignacion" not in st.session_state:
            st.session_state["ajustes_asignacion"] = {}
        lista_puestos = asignacion_filtrada["PUESTO"].dropna().sort_values().tolist()
        if not lista_puestos:
            lista_puestos = asignacion_vista["PUESTO"].dropna().sort_values().tolist()
        puesto_sel = st.selectbox("Puesto de votación", lista_puestos, key="puesto_ajuste_compacto")
        puesto_row = asignacion_vista[asignacion_vista["PUESTO"].eq(puesto_sel)].iloc[0]
        templo_actual = puesto_row.get("TEMPLO_ASIGNADO_FINAL")
        index_templo = TEMPLOS_OFICIALES.index(templo_actual) if templo_actual in TEMPLOS_OFICIALES else 0

        p1, p2 = st.columns([2, 1])
        with p1:
            st.dataframe(
                pd.DataFrame(
                    [
                        ("Templo documento base", puesto_row.get("IGLESIA_ACTUAL")),
                        ("Templo más cercano", puesto_row.get("TEMPLO_MAS_CERCANO")),
                        ("Templo propuesta actual", puesto_row.get("TEMPLO_ASIGNADO_PROPUESTO")),
                        ("Templo vigente", puesto_row.get("TEMPLO_ASIGNADO_FINAL")),
                        ("Votos 2026", fmt_number(puesto_row.get("VOTOS_2026"), 0)),
                        ("Prioridad", puesto_row.get("PRIORIDAD")),
                        ("Estado asignación", puesto_row.get("ESTADO_ASIGNACION")),
                    ],
                    columns=["Campo", "Valor"],
                ),
                hide_index=True,
                width="stretch",
            )
            historial_puesto = historial_puestos_db[historial_puestos_db["entidad_id"].astype(str).eq(str(puesto_sel))].copy() if not historial_puestos_db.empty else pd.DataFrame()
            if not historial_puesto.empty:
                st.markdown("Historial de cambios")
                st.dataframe(historial_puesto[["creado_en", "templo_anterior", "templo_nuevo", "usuario", "motivo"]], hide_index=True, width="stretch")
        with p2:
            templo_nuevo = st.selectbox("Templo final", TEMPLOS_OFICIALES, index=index_templo, key="templo_puesto_ajuste_compacto")
            nota_cambio = st.text_area("Justificación o nota del cambio", key="nota_cambio_puesto", height=110)
            if st.button("Guardar cambio definitivo"):
                st.session_state["ajustes_asignacion"][puesto_sel] = templo_nuevo
                registrar_ajuste_en_db(
                    session_key="ajustes_asignacion",
                    entity_id=puesto_sel,
                    nombre_entidad=puesto_sel,
                    templo_nuevo=templo_nuevo,
                    motivo=nota_cambio or "Ajuste manual desde pestaña Asignación de puestos",
                )
                guardar_ajustes_guardados()
                st.success(f"Ajuste guardado: {puesto_sel} -> {templo_nuevo}")
                st.rerun()
            if st.button("Limpiar ajustes de puestos"):
                total_limpiados = limpiar_ajustes_en_db("ajustes_asignacion", motivo="Limpieza manual desde pestaña Asignación de puestos")
                st.session_state["ajustes_asignacion"] = {}
                guardar_ajustes_guardados()
                st.info(f"Se limpiaron {fmt_number(total_limpiados, 0)} ajuste(s) de puestos en la base.")
                st.rerun()

    st.markdown("### Exportables compactos")
    tabla_templos = crear_tabla_puestos_por_templo(asignacion_vista)
    informe_general_asignacion = generar_informe_territorial(asignacion_vista, actividades, mesas)
    excel_asignacion = exportar_asignacion_excel(asignacion_vista, resumen_final, tabla_templos)
    excel_cambios = multi_sheet_excel_bytes(
        {
            "cambios_guardados": cambios_show,
            "historial": historial_puestos_db,
            "impacto_templo": impacto,
        }
    )
    excel_por_templo = exportar_asignacion_por_templo_excel(asignacion_vista, actividades, mesas)
    e1, e2, e3, e4 = st.columns(4)
    with e1:
        st.download_button("Descargar asignación consolidada XLSX", excel_asignacion, "asignacion_consolidada_vigente.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with e2:
        st.download_button("Descargar cambios guardados XLSX", excel_cambios, "cambios_guardados_asignacion.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with e3:
        st.download_button("Descargar informe por templo", excel_por_templo, "informe_por_templo_asignacion.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with e4:
        st.download_button("Descargar informe general", informe_general_asignacion.encode("utf-8"), "informe_general_asignacion.md", "text/markdown")

with tab_iglesia:
    st.subheader("Análisis por iglesia")
    cols_show = [
        "IGLESIA", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL",
        "PUESTOS", "ACTIVIDADES_CAMPANA", "MESAS_TRABAJO", "PUESTO_MAYOR_VOTACION",
        "PUESTO_MAYOR_CAIDA", "PUESTO_MAYOR_CRECIMIENTO"
    ]
    st.dataframe(resumen_iglesia_f[cols_show], width="stretch", hide_index=True)
    st.download_button("Descargar análisis por iglesia en Excel", to_excel_bytes(resumen_iglesia_f[cols_show], "Analisis Iglesia"), "analisis_iglesia.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_iglesia")

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
        st.plotly_chart(fig, width="stretch")
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
        st.plotly_chart(fig, width="stretch")

    # Averages for radar chart
    promedio_votos = resumen_iglesia_f["VOTOS_2026"].mean() if not resumen_iglesia_f.empty else 1
    promedio_crecimiento = resumen_iglesia_f["VARIACION_PORCENTUAL"].mean() if not resumen_iglesia_f.empty else 0.01
    promedio_actividades = resumen_iglesia_f["ACTIVIDADES_CAMPANA"].mean() if not resumen_iglesia_f.empty else 1
    promedio_mesas = resumen_iglesia_f["MESAS_TRABAJO"].mean() if not resumen_iglesia_f.empty else 1

    for iglesia in iglesias["IGLESIA"].tolist():
        sub_puestos = puestos_f[puestos_f["IGLESIA"].eq(iglesia)].copy()
        sub_acts = actividades_f[actividades_f["IGLESIA"].eq(iglesia)].copy()
        sub_mesas = mesas_f[mesas_f["IGLESIA"].eq(iglesia)].copy()
        res = resumen_iglesia_f[resumen_iglesia_f["IGLESIA"].eq(iglesia)]
        
        with st.expander(f"Ficha Estratégica Integral: {iglesia}", expanded=False):
            if not res.empty:
                r_iglesia = res.iloc[0]
                informe_templo_md = generar_informe_templo_markdown(iglesia, puestos_f, actividades_f, mesas_f, r_iglesia)
                
                # 1. KPIs
                k1, k2, k3, k4 = st.columns(4)
                votos_2026 = r_iglesia.get("VOTOS_2026", 0)
                votos_2023 = r_iglesia.get("VOTOS_2023", 0)
                var_abs = r_iglesia.get("VARIACION_ABSOLUTA", 0)
                var_pct = r_iglesia.get("VARIACION_PORCENTUAL", 0)
                acts_totales = r_iglesia.get("ACTIVIDADES_CAMPANA", 0) + r_iglesia.get("MESAS_TRABAJO", 0)
                roi = votos_2026 / acts_totales if acts_totales > 0 else 0
                
                # Semaforo de Riesgo
                if var_pct < -0.10:
                    semaforo = "Critico (>10% perdida)"
                elif var_pct < 0:
                    semaforo = "Medio (perdida)"
                else:
                    semaforo = "Fortaleza (crecimiento)"
                
                with k1:
                    metric_card("Votos 2026", fmt_number(votos_2026, 0), icon="📊")
                with k2:
                    metric_card("Retención", fmt_pct(var_pct), icon="📈")
                with k3:
                    metric_card("ROI Político", f"{fmt_number(roi, 1)} v/act", icon="⚡")
                with k4:
                    metric_card("Estado", semaforo, icon="🚦")

                st.download_button(
                    f"Descargar informe {iglesia}",
                    informe_templo_md.encode("utf-8"),
                    f"informe_templo_{iglesia.lower().replace(' ', '_')}.md",
                    "text/markdown",
                    key=f"dl_informe_templo_{iglesia}",
                )
                
                # 2. Charts and Maps
                col_chart, col_map = st.columns([1, 1.2])
                with col_chart:
                    # Radar Chart
                    fig_radar = go.Figure()
                    fig_radar.add_trace(go.Scatterpolar(
                        r=[
                            min(votos_2026 / max(1, promedio_votos), 2.5), 
                            min(max(var_pct + 1, 0) / max(0.01, promedio_crecimiento + 1), 2.5), 
                            min(r_iglesia.get("ACTIVIDADES_CAMPANA", 0) / max(1, promedio_actividades), 2.5), 
                            min(r_iglesia.get("MESAS_TRABAJO", 0) / max(1, promedio_mesas), 2.5),
                            min(roi / max(1, (promedio_votos / max(1, promedio_actividades + promedio_mesas))), 2.5)
                        ],
                        theta=['Caudal<br>Electoral', 'Retención<br>Electoral', 'Campaña<br>(Actividades)', 'Trabajo<br>Social (Mesas)', 'Eficiencia<br>(ROI)'],
                        fill='toself',
                        name=iglesia,
                        line_color='#1D4ED8',
                        fillcolor='rgba(29, 78, 216, 0.3)'
                    ))
                    fig_radar.update_layout(
                        polar=dict(radialaxis=dict(visible=False, range=[0, 2.5])),
                        showlegend=False,
                        title="Perfil Estratégico (vs Promedio Kennedy)",
                        margin=dict(t=50, b=20, l=40, r=40),
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)"
                    )
                    st.plotly_chart(fig_radar, use_container_width=True)

                with col_map:
                    # Create Sub-map
                    templo_coord = iglesias[iglesias["IGLESIA"].eq(iglesia)]
                    if not templo_coord.empty:
                        lat_t = templo_coord.iloc[0]["LATITUD"]
                        lon_t = templo_coord.iloc[0]["LONGITUD"]
                        sub_m = folium.Map(location=[lat_t, lon_t], zoom_start=14, tiles="cartodbpositron")
                        
                        # Add Temple
                        folium.Marker(
                            [lat_t, lon_t], 
                            icon=crear_icono_div("templo", COLORES_TEMPLOS.get(iglesia, "#1E3A8A"), "T"),
                            tooltip=iglesia
                        ).add_to(sub_m)
                        
                        # Add Puestos and lines
                        for _, r in sub_puestos.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
                            folium.CircleMarker(
                                [r["LATITUD"], r["LONGITUD"]],
                                radius=max(5, min(14, float(r.get("VOTOS_2026", 0))/15)),
                                color="#1E3A8A", fill=True, fill_color="#3B82F6", fill_opacity=0.7, weight=1,
                                tooltip=f"{r.get('PUESTO')} | {fmt_number(r.get('VOTOS_2026'),0)} votos"
                            ).add_to(sub_m)
                            folium.PolyLine([[lat_t, lon_t], [r["LATITUD"], r["LONGITUD"]]], color="#1E3A8A", weight=1.5, opacity=0.35, dash_array="4,6").add_to(sub_m)

                        # Add Mesas
                        for _, r in sub_mesas.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
                            folium.Marker(
                                [r["LATITUD"], r["LONGITUD"]],
                                icon=crear_icono_div("mesa", "#F97316", "M"),
                                tooltip=f"Mesa: {r.get('TEMA')}"
                            ).add_to(sub_m)
                            
                        st_folium(sub_m, height=420, use_container_width=True, key=f"submap_{iglesia}")
                    else:
                        st.info("Sin coordenadas del templo.")

                st.markdown("---")
                st.write(f"**Lectura estratégica:** {r_iglesia.get('LECTURA_ESTRATEGICA', '')}")
                st.write(f"**Recomendación táctica:** {r_iglesia.get('RECOMENDACION', '')}")
            
            if sub_puestos.empty:
                st.info("No hay puestos asignados en la matriz electoral. Se mantiene como iglesia oficial para análisis territorial.")
            else:
                st.dataframe(
                    sub_puestos[["PUESTO", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "PRIORIDAD", "ACCION_RECOMENDADA"]]
                    .sort_values("VOTOS_2026", ascending=False),
                    width="stretch",
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
        st.plotly_chart(fig, width="stretch")
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
        st.plotly_chart(fig, width="stretch")

    st.markdown("### Matriz analítica por puesto")
    cols_puesto = [
        "PUESTO", "IGLESIA", "IGLESIA_HISTORICA_2026", "TEMPLO_OPERATIVO_ACTUAL", "TEMPLO_PROPUESTO",
        "CAMBIO_PROPUESTO_TEMPLO", "ROL_ANALITICO", "MIRA_CONCEJO_2023", "JAL_2023", "CAMARA_2026", "SENADO_2026",
        "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL",
        "MESAS_2026_REPORTE", "TESTIGOS_2023_REPORTE", "VOTOS_AFINIDAD_E11_2023",
        "TEMPLO_REPORTE", "TIENE_MESA_TRABAJO", "PUNTAJE_PRIORIDAD", "FACTORES_PRIORIDAD",
        "PRIORIDAD", "ACCION_RECOMENDADA"
    ]
    cols_puesto = [c for c in cols_puesto if c in puestos_f.columns]
    df_puesto_show = puestos_f[cols_puesto].sort_values(["PRIORIDAD", "VOTOS_2026"], ascending=[True, False])
    st.dataframe(df_puesto_show, width="stretch", hide_index=True)
    st.download_button("Descargar matriz por puesto en Excel", to_excel_bytes(df_puesto_show, "Matriz Puesto"), "matriz_puesto.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_mat_puesto")

with tab_barrio:
    st.subheader("Barrio / UPZ")
    if resumen_barrio.empty:
        st.info("No hay resumen por barrio o UPZ disponible.")
    else:
        st.dataframe(resumen_barrio, width="stretch", hide_index=True)
        st.download_button("Descargar resumen por barrio en Excel", to_excel_bytes(resumen_barrio, "Resumen Barrio"), "resumen_barrio.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_res_barrio")

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
        st.plotly_chart(fig, width="stretch")
    else:
        st.markdown(
            '<div class="warning-box">La base consolidada aún no tiene UPZ asignada. Para activar análisis espacial por UPZ, agregue <b>data/upz_kennedy.geojson</b> o complete la columna UPZ en el Excel consolidado.</div>',
            unsafe_allow_html=True,
        )

with tab_prioridad:
    st.subheader("Matriz de priorización territorial")
    st.markdown(
        '<div class="note-box">La priorización combina concentración electoral, variación, presencia comunitaria, mesas de trabajo y distancia logística. El puntaje ayuda a ordenar intervención; no reemplaza la validación política en territorio.</div>',
        unsafe_allow_html=True,
    )
    matriz_show = matriz.copy()
    if "PUNTAJE_PRIORIDAD" in matriz_show.columns:
        matriz_show = matriz_show.sort_values(["NIVEL_PRIORIDAD", "PUNTAJE_PRIORIDAD", "VOTOS_2026"], ascending=[True, False, False])
    st.dataframe(matriz_show, width="stretch", hide_index=True)
    st.download_button("Descargar matriz de priorización en Excel", to_excel_bytes(matriz_show, "Priorizacion"), "matriz_priorizacion.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_mat_prior")

    prioridad_count = matriz.groupby("NIVEL_PRIORIDAD").size().reset_index(name="PUESTOS")
    pcol1, pcol2 = st.columns(2)
    with pcol1:
        fig = px.bar(
            prioridad_count,
            x="NIVEL_PRIORIDAD",
            y="PUESTOS",
            title="Distribución de puestos por prioridad",
            color="NIVEL_PRIORIDAD",
        )
        fig.update_layout(height=380, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, showlegend=False)
        st.plotly_chart(fig, width="stretch")
    with pcol2:
        if {"VOTOS_2026", "VARIACION_ABSOLUTA", "PUNTAJE_PRIORIDAD"}.issubset(matriz.columns):
            fig = px.scatter(
                matriz,
                x="VOTOS_2026",
                y="VARIACION_ABSOLUTA",
                size="PUNTAJE_PRIORIDAD",
                color="NIVEL_PRIORIDAD",
                hover_name="PUESTO",
                title="Concentración electoral vs variación",
            )
            fig.update_layout(height=380, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
            st.plotly_chart(fig, width="stretch")

    if "ROL_ANALITICO" in matriz.columns:
        rol_count = matriz.groupby(["ROL_ANALITICO", "NIVEL_PRIORIDAD"]).size().reset_index(name="PUESTOS")
        fig = px.bar(
            rol_count,
            x="ROL_ANALITICO",
            y="PUESTOS",
            color="NIVEL_PRIORIDAD",
            title="Puestos por rol analítico y prioridad",
        )
        fig.update_layout(height=390, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
        st.plotly_chart(fig, width="stretch")

with tab_export:
    st.subheader("Exportables")
    st.markdown("Descargue la base maestra consolidada o tablas específicas para anexos del informe.")
    informe_ejecutivo_export = generar_informe_ejecutivo_markdown(puestos, resumen_iglesia, matriz, actividades, mesas)
    historial_ajustes = obtener_historial_ajustes(limit=5000)

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
        "historial_ajustes_templo.csv": historial_ajustes,
    }

    for fname, df in exportables.items():
        st.download_button(
            f"Descargar {fname}",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name=fname,
            mime="text/csv",
        )

    st.download_button(
        "Descargar informe ejecutivo territorial",
        data=informe_ejecutivo_export.encode("utf-8"),
        file_name="informe_ejecutivo_territorial_kennedy.md",
        mime="text/markdown",
    )

    with st.expander("Control de calidad de datos"):
        st.dataframe(control, width="stretch", hide_index=True)

    with st.expander("Bitácora de ajustes de templo (base de datos)"):
        st.dataframe(historial_ajustes, width="stretch", hide_index=True)
