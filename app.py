
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
import hmac
import html
from io import BytesIO
import json
import math
import re
import sqlite3
import subprocess
import unicodedata
import urllib.request
from datetime import datetime, timezone

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import folium
from branca.element import MacroElement, Template
from folium.plugins import HeatMap, Fullscreen, MiniMap
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont
from streamlit_folium import folium_static, st_folium

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
TESTIGOS_RESUMEN_CSV = DATA_DIR / "testigos_resumen_2026.csv"
APOYOS_CIUDADANOS_CANDIDATES = [
    DATA_DIR / "Gestion apoyo ciudadano Kennedy.xlsx",
    Path("..") / "Gestion apoyo ciudadano Kennedy.xlsx",
]
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
COLOR_PURPLE = "#C026D3"
BASE_TILE_URL = "https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png"
BASE_TILE_ATTR = "&copy; OpenStreetMap contributors &copy; CARTO"
TEMPLOS_OFICIALES = ["CLASS ROMA", "KENNEDY CENTRAL", "PATIO BONITO", "CARVAJAL", "VALLADOLID"]
METODOLOGIAS_ELECTORALES = {
    "camara": {
        "label": "JAL 2023 vs Cámara 2026",
        "caption": "Lectura política específica de Cámara.",
        "base_col": "JAL_2023",
        "actual_col": "CAMARA_2026",
        "base_label": "JAL 2023",
        "actual_label": "Cámara 2026",
    },
    "promedio": {
        "label": "Promedio electoral 2023 vs 2026",
        "caption": "Lectura general balanceada.",
        "base_col": "VOTOS_2023_BASE",
        "actual_col": "VOTOS_2026_BASE",
        "base_label": "Promedio 2023",
        "actual_label": "Promedio 2026",
    },
    "senado": {
        "label": "Concejo 2023 vs Senado 2026",
        "caption": "Lectura política específica de Senado.",
        "base_col": "MIRA_CONCEJO_2023",
        "actual_col": "SENADO_2026",
        "base_label": "Concejo 2023",
        "actual_label": "Senado 2026",
    },
}
METODOLOGIA_DEFAULT = "camara"
COLORES_TEMPLOS = {
    "CLASS ROMA": "#C026D3",
    "KENNEDY CENTRAL": "#2563EB",
    "PATIO BONITO": "#16A34A",
    "CARVAJAL": "#F97316",
    "VALLADOLID": "#DC2626",
}
PLOTLY_TEMPLO_ORDERS = {
    "IGLESIA": TEMPLOS_OFICIALES,
    "TEMPLO": TEMPLOS_OFICIALES,
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

    .metric-icon {
        width: 2.35rem;
        height: 2.35rem;
        border-radius: 12px;
        background: linear-gradient(135deg, #EFF6FF 0%, #ECFDF5 100%);
        border: 1px solid rgba(148, 163, 184, 0.28);
        color: #2563EB;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        flex: 0 0 auto;
    }

    .metric-icon svg {
        width: 1.1rem;
        height: 1.1rem;
        stroke: currentColor;
        stroke-width: 2;
        fill: none;
        stroke-linecap: round;
        stroke-linejoin: round;
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

    .status-strip {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 0.85rem;
        margin: 0.25rem 0 1.35rem 0;
    }

    .status-item {
        background: linear-gradient(135deg, rgba(255,255,255,0.98) 0%, rgba(240,253,244,0.86) 100%);
        border: 1px solid rgba(187, 247, 208, 0.95);
        border-radius: 15px;
        padding: 0.9rem 1rem;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.055);
    }

    .status-item span {
        display: block;
        color: #475569;
        font-size: 0.76rem;
        font-weight: 800;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        margin-bottom: 0.25rem;
    }

    .status-item strong {
        color: #0F172A;
        font-size: 1.08rem;
        line-height: 1.2;
    }

    .layer-hint {
        background: #F0FDF4;
        border: 1px solid #BBF7D0;
        border-radius: 12px;
        color: #14532D;
        font-size: 0.9rem;
        font-weight: 650;
        padding: 0.72rem 0.9rem;
        margin: 0.45rem 0 0.7rem 0;
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
        border: 1px solid rgba(203, 213, 225, 0.82);
        border-radius: 14px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.045);
        overflow: hidden;
    }

    div[data-testid="stDataFrame"] [role="columnheader"] {
        background: #EFF6FF !important;
        color: #1E3A8A !important;
        font-weight: 850 !important;
        letter-spacing: 0.01em !important;
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

    @media (max-width: 900px) {
        .status-strip {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }
    }

    @media (max-width: 560px) {
        .status-strip {
            grid-template-columns: 1fr;
        }
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
AJUSTES_ACTUALES_COLUMNS = [
    "entidad", "entidad_id", "nombre_entidad", "templo_actual", "barrio_actual",
    "lider_actual", "usuario", "motivo", "actualizado_en",
]
AJUSTES_HISTORIAL_COLUMNS = [
    "id", "creado_en", "entidad", "entidad_id", "nombre_entidad",
    "templo_anterior", "templo_nuevo", "barrio_anterior", "barrio_nuevo",
    "lider_anterior", "lider_nuevo", "usuario", "motivo",
]

def cargar_ajustes_guardados():
    if AJUSTES_FILE.exists():
        try:
            with open(AJUSTES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "ajustes_asignacion": {},
        "ajustes_mesas": {},
        "ajustes_actividades": {},
        "ajustes_mesas_barrio": {},
        "ajustes_mesas_lider": {},
    }


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


def _normalize_spreadsheet_id(value):
    """Accept a raw Google Sheet ID or a full Sheets URL pasted in secrets."""
    raw = str(value or "").strip()
    if not raw:
        return ""
    url_match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", raw)
    if url_match:
        return url_match.group(1)
    return raw.split("/edit", 1)[0].split("?", 1)[0].split("#", 1)[0].strip()


def _google_sheets_config():
    try:
        sheets_cfg = st.secrets.get("google_sheets", {})
        service_account = st.secrets.get("gcp_service_account", {})
    except Exception:
        return None

    spreadsheet_id = _normalize_spreadsheet_id(sheets_cfg.get("spreadsheet_id", "")) if sheets_cfg else ""
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
        old_header = values[0]
        old_rows = values[1:]
        if old_rows:
            max_width = max([len(old_header)] + [len(row) for row in old_rows])
            safe_header = old_header + [f"__extra_{idx}" for idx in range(len(old_header), max_width)]
            safe_rows = [row + [""] * (max_width - len(row)) for row in old_rows]
            migrated = pd.DataFrame(safe_rows, columns=safe_header)
        else:
            migrated = pd.DataFrame(columns=old_header)
        migrated = _normalize_sheet_df(migrated, columns).astype(str)
        worksheet.clear()
        worksheet.append_row(columns, value_input_option="USER_ENTERED")
        if not migrated.empty:
            worksheet.append_rows(migrated.values.tolist(), value_input_option="USER_ENTERED")
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
        st.session_state["google_sheets_ready"] = False
        st.session_state["google_sheets_error"] = "Secrets de Google Sheets incompletos o no disponibles."
        return False
    if "google_sheets_ready" not in st.session_state or not st.session_state.get("google_sheets_ready"):
        try:
            st.session_state["google_sheets_ready"] = bool(_init_google_sheets_storage())
            if st.session_state["google_sheets_ready"]:
                st.session_state.pop("google_sheets_error", None)
        except Exception as exc:
            st.session_state["google_sheets_ready"] = False
            st.session_state["google_sheets_error"] = str(exc)
    elif st.session_state.get("google_sheets_ready"):
        st.session_state.pop("google_sheets_error", None)
    return bool(st.session_state.get("google_sheets_ready"))


def persistence_backend_label():
    if _google_sheets_ready():
        return "Google Sheets online"
    if _google_sheets_config() and st.session_state.get("google_sheets_error"):
        return "SQLite local (Google Sheets no conectado)"
    return "SQLite local"


def google_sheets_diagnostics():
    try:
        sheets_cfg = st.secrets.get("google_sheets", {})
        service_account = st.secrets.get("gcp_service_account", {})
    except Exception as exc:
        return {
            "Secrets leídos": "No",
            "Detalle": str(exc)[:240],
        }

    raw_spreadsheet_id = sheets_cfg.get("spreadsheet_id", "") if sheets_cfg else ""
    spreadsheet_id = _normalize_spreadsheet_id(raw_spreadsheet_id)
    service_email = str(service_account.get("client_email", "")) if service_account else ""
    has_private_key = bool(service_account.get("private_key")) if service_account else False
    masked_spreadsheet_id = (
        f"{spreadsheet_id[:8]}...{spreadsheet_id[-8:]}" if len(spreadsheet_id) > 16 else (spreadsheet_id or "No configurado")
    )

    return {
        "Secrets leídos": "Sí",
        "google_sheets": "Sí" if bool(sheets_cfg) else "No",
        "spreadsheet_id": masked_spreadsheet_id,
        "cuenta_servicio": service_email or "No configurada",
        "private_key": "Sí" if has_private_key else "No",
        "estado": "Conectado" if st.session_state.get("google_sheets_ready") else "No conectado",
        "último_error": st.session_state.get("google_sheets_error", "")[:300] or "Sin error registrado",
    }


def _read_google_sheet_df(title, columns):
    spreadsheet = _get_google_spreadsheet()
    worksheet = _get_or_create_worksheet(spreadsheet, title, columns)
    records = worksheet.get_all_records(default_blank="")
    df = _normalize_sheet_df(pd.DataFrame(records), columns)
    if title == GOOGLE_SHEETS_ACTUALES and not df.empty:
        usuario_parece_fecha = df["usuario"].astype(str).str.match(r"^\d{4}-\d{2}-\d{2}", na=False)
        fecha_vacia = df["actualizado_en"].astype(str).str.strip().eq("")
        filas_desplazadas = usuario_parece_fecha & fecha_vacia
        if filas_desplazadas.any():
            old_usuario = df.loc[filas_desplazadas, "barrio_actual"].copy()
            old_motivo = df.loc[filas_desplazadas, "lider_actual"].copy()
            old_fecha = df.loc[filas_desplazadas, "usuario"].copy()
            df.loc[filas_desplazadas, "usuario"] = old_usuario
            df.loc[filas_desplazadas, "motivo"] = old_motivo
            df.loc[filas_desplazadas, "actualizado_en"] = old_fecha
            df.loc[filas_desplazadas, ["barrio_actual", "lider_actual"]] = ""
    return df


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
                barrio_actual TEXT,
                lider_actual TEXT,
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
                barrio_anterior TEXT,
                barrio_nuevo TEXT,
                lider_anterior TEXT,
                lider_nuevo TEXT,
                usuario TEXT,
                motivo TEXT,
                creado_en TEXT NOT NULL
            )
            """
        )
        for table, columns in {
            "ajustes_actuales": {"barrio_actual": "TEXT", "lider_actual": "TEXT"},
            "ajustes_historial": {
                "barrio_anterior": "TEXT",
                "barrio_nuevo": "TEXT",
                "lider_anterior": "TEXT",
                "lider_nuevo": "TEXT",
            },
        }.items():
            existing_cols = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
            for col, col_type in columns.items():
                if col not in existing_cols:
                    conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_historial_creado_en ON ajustes_historial(creado_en DESC)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_historial_entidad_id ON ajustes_historial(entidad, entidad_id)")


def init_ajustes_db():
    _init_ajustes_sqlite()
    if _google_sheets_config():
        _google_sheets_ready()


def _cargar_ajustes_desde_sqlite():
    ajustes = cargar_ajustes_guardados()
    for key in ["ajustes_asignacion", "ajustes_mesas", "ajustes_actividades", "ajustes_mesas_barrio", "ajustes_mesas_lider"]:
        ajustes.setdefault(key, {})
    with get_db_connection() as conn:
        rows = conn.execute("SELECT entidad, entidad_id, templo_actual, barrio_actual, lider_actual FROM ajustes_actuales").fetchall()
    for row in rows:
        session_key = _entidad_to_session_key(row["entidad"])
        if not session_key:
            continue
        key = _normalize_entity_id(row["entidad"], row["entidad_id"])
        ajustes[session_key][key] = row["templo_actual"]
        if row["entidad"] == "mesa":
            if row["barrio_actual"]:
                ajustes["ajustes_mesas_barrio"][key] = row["barrio_actual"]
            if row["lider_actual"]:
                ajustes["ajustes_mesas_lider"][key] = row["lider_actual"]
    return ajustes


def _cargar_ajustes_desde_google_sheets():
    ajustes = cargar_ajustes_guardados()
    for key in ["ajustes_asignacion", "ajustes_mesas", "ajustes_actividades", "ajustes_mesas_barrio", "ajustes_mesas_lider"]:
        ajustes.setdefault(key, {})
    rows = _read_google_sheet_df(GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS)
    for _, row in rows.iterrows():
        session_key = _entidad_to_session_key(row["entidad"])
        if not session_key:
            continue
        key = _normalize_entity_id(row["entidad"], row["entidad_id"])
        ajustes[session_key][key] = row["templo_actual"]
        if str(row["entidad"]) == "mesa":
            if str(row.get("barrio_actual", "")).strip():
                ajustes["ajustes_mesas_barrio"][key] = row.get("barrio_actual", "")
            if str(row.get("lider_actual", "")).strip():
                ajustes["ajustes_mesas_lider"][key] = row.get("lider_actual", "")
    return ajustes


def cargar_ajustes_desde_db():
    if _google_sheets_ready():
        try:
            return _cargar_ajustes_desde_google_sheets()
        except Exception as exc:
            st.session_state["google_sheets_error"] = str(exc)
    return _cargar_ajustes_desde_sqlite()


def _registrar_ajuste_sqlite(session_key, entity_id, nombre_entidad, templo_nuevo, motivo="", barrio_nuevo=None, lider_nuevo=None):
    entidad = _session_key_to_entidad(session_key)
    entidad_id = str(_normalize_entity_id(entidad, entity_id))
    usuario = st.session_state.get("usuario_actual", "usuario_dashboard")
    now = _ahora_utc_iso()

    with get_db_connection() as conn:
        previo = conn.execute(
            "SELECT templo_actual, barrio_actual, lider_actual FROM ajustes_actuales WHERE entidad = ? AND entidad_id = ?",
            (entidad, entidad_id),
        ).fetchone()
        templo_anterior = previo["templo_actual"] if previo else None
        barrio_anterior = previo["barrio_actual"] if previo else ""
        lider_anterior = previo["lider_actual"] if previo else ""
        barrio_nuevo = str(barrio_anterior or "") if barrio_nuevo is None else str(barrio_nuevo)
        lider_nuevo = str(lider_anterior or "") if lider_nuevo is None else str(lider_nuevo)
        if (
            str(templo_anterior or "") == str(templo_nuevo or "")
            and str(barrio_anterior or "") == barrio_nuevo
            and str(lider_anterior or "") == lider_nuevo
        ):
            return False, templo_anterior

        conn.execute(
            """
            INSERT INTO ajustes_actuales(entidad, entidad_id, nombre_entidad, templo_actual, barrio_actual, lider_actual, usuario, motivo, actualizado_en)
            VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(entidad, entidad_id) DO UPDATE SET
                nombre_entidad=excluded.nombre_entidad,
                templo_actual=excluded.templo_actual,
                barrio_actual=excluded.barrio_actual,
                lider_actual=excluded.lider_actual,
                usuario=excluded.usuario,
                motivo=excluded.motivo,
                actualizado_en=excluded.actualizado_en
            """,
            (entidad, entidad_id, str(nombre_entidad or ""), str(templo_nuevo), barrio_nuevo, lider_nuevo, usuario, str(motivo or ""), now),
        )
        conn.execute(
            """
            INSERT INTO ajustes_historial(entidad, entidad_id, nombre_entidad, templo_anterior, templo_nuevo, barrio_anterior, barrio_nuevo, lider_anterior, lider_nuevo, usuario, motivo, creado_en)
            VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (entidad, entidad_id, str(nombre_entidad or ""), templo_anterior, str(templo_nuevo), barrio_anterior, barrio_nuevo, lider_anterior, lider_nuevo, usuario, str(motivo or ""), now),
        )
    return True, templo_anterior


def _registrar_ajuste_google_sheets(session_key, entity_id, nombre_entidad, templo_nuevo, motivo="", barrio_nuevo=None, lider_nuevo=None):
    entidad = _session_key_to_entidad(session_key)
    entidad_id = str(_normalize_entity_id(entidad, entity_id))
    usuario = st.session_state.get("usuario_actual", "usuario_dashboard")
    now = _ahora_utc_iso()

    actuales = _read_google_sheet_df(GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS)
    historial = _read_google_sheet_df(GOOGLE_SHEETS_HISTORIAL, AJUSTES_HISTORIAL_COLUMNS)
    match = actuales["entidad"].astype(str).eq(entidad) & actuales["entidad_id"].astype(str).eq(entidad_id)
    templo_anterior = actuales.loc[match, "templo_actual"].iloc[0] if match.any() else None
    barrio_anterior = actuales.loc[match, "barrio_actual"].iloc[0] if match.any() and "barrio_actual" in actuales.columns else ""
    lider_anterior = actuales.loc[match, "lider_actual"].iloc[0] if match.any() and "lider_actual" in actuales.columns else ""
    barrio_nuevo = str(barrio_anterior or "") if barrio_nuevo is None else str(barrio_nuevo)
    lider_nuevo = str(lider_anterior or "") if lider_nuevo is None else str(lider_nuevo)
    if (
        str(templo_anterior or "") == str(templo_nuevo or "")
        and str(barrio_anterior or "") == barrio_nuevo
        and str(lider_anterior or "") == lider_nuevo
    ):
        return False, templo_anterior

    nueva_fila = {
        "entidad": entidad,
        "entidad_id": entidad_id,
        "nombre_entidad": str(nombre_entidad or ""),
        "templo_actual": str(templo_nuevo),
        "barrio_actual": barrio_nuevo,
        "lider_actual": lider_nuevo,
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
        "barrio_anterior": barrio_anterior or "",
        "barrio_nuevo": barrio_nuevo,
        "lider_anterior": lider_anterior or "",
        "lider_nuevo": lider_nuevo,
        "usuario": usuario,
        "motivo": str(motivo or ""),
    }
    _rewrite_google_sheet_df(GOOGLE_SHEETS_ACTUALES, AJUSTES_ACTUALES_COLUMNS, actuales)
    _append_google_sheet_row(GOOGLE_SHEETS_HISTORIAL, AJUSTES_HISTORIAL_COLUMNS, historial_row)
    return True, templo_anterior


def registrar_ajuste_en_db(session_key, entity_id, nombre_entidad, templo_nuevo, motivo="", barrio_nuevo=None, lider_nuevo=None):
    if _google_sheets_ready():
        try:
            changed, templo_anterior = _registrar_ajuste_google_sheets(session_key, entity_id, nombre_entidad, templo_nuevo, motivo, barrio_nuevo, lider_nuevo)
            if changed:
                _registrar_ajuste_sqlite(session_key, entity_id, nombre_entidad, templo_nuevo, motivo, barrio_nuevo, lider_nuevo)
            return changed, templo_anterior
        except Exception as exc:
            st.session_state["google_sheets_error"] = str(exc)
            st.warning("No se pudo guardar en Google Sheets. Se guardará una copia local en SQLite.")
    return _registrar_ajuste_sqlite(session_key, entity_id, nombre_entidad, templo_nuevo, motivo, barrio_nuevo, lider_nuevo)


def _limpiar_ajustes_sqlite(session_key, motivo=""):
    entidad = _session_key_to_entidad(session_key)
    usuario = st.session_state.get("usuario_actual", "usuario_dashboard")
    now = _ahora_utc_iso()
    with get_db_connection() as conn:
        actuales = conn.execute(
            "SELECT entidad_id, nombre_entidad, templo_actual, barrio_actual, lider_actual FROM ajustes_actuales WHERE entidad = ?",
            (entidad,),
        ).fetchall()
        for row in actuales:
            conn.execute(
                """
                INSERT INTO ajustes_historial(entidad, entidad_id, nombre_entidad, templo_anterior, templo_nuevo, barrio_anterior, barrio_nuevo, lider_anterior, lider_nuevo, usuario, motivo, creado_en)
                VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (entidad, row["entidad_id"], row["nombre_entidad"], row["templo_actual"], None, row["barrio_actual"], "", row["lider_actual"], "", usuario, str(motivo or "limpieza masiva"), now),
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
                "barrio_anterior": row.get("barrio_actual", ""),
                "barrio_nuevo": "",
                "lider_anterior": row.get("lider_actual", ""),
                "lider_nuevo": "",
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
            SELECT id, creado_en, entidad, entidad_id, nombre_entidad, templo_anterior, templo_nuevo, barrio_anterior, barrio_nuevo, lider_anterior, lider_nuevo, usuario, motivo
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
            SELECT entidad, entidad_id, nombre_entidad, templo_actual, barrio_actual, lider_actual, usuario, motivo, actualizado_en
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
        "ajustes_actividades": st.session_state.get("ajustes_actividades", {}),
        "ajustes_mesas_barrio": st.session_state.get("ajustes_mesas_barrio", {}),
        "ajustes_mesas_lider": st.session_state.get("ajustes_mesas_lider", {}),
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
        st.session_state["ajustes_mesas_barrio"] = ajustes_db.get("ajustes_mesas_barrio", {})
        st.session_state["ajustes_mesas_lider"] = ajustes_db.get("ajustes_mesas_lider", {})
    else:
        ajustes_disco = cargar_ajustes_guardados()
        st.session_state["ajustes_asignacion"] = ajustes_disco.get("ajustes_asignacion", {})
        st.session_state["ajustes_mesas"] = ajustes_disco.get("ajustes_mesas", {})
        st.session_state["ajustes_actividades"] = ajustes_disco.get("ajustes_actividades", {})
        st.session_state["ajustes_mesas_barrio"] = ajustes_disco.get("ajustes_mesas_barrio", {})
        st.session_state["ajustes_mesas_lider"] = ajustes_disco.get("ajustes_mesas_lider", {})
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


@st.cache_data
def cargar_testigos_resumen(path: Path, mtime: float = 0):
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path)
    if "TEMPLO" not in df.columns:
        return pd.DataFrame()
    for col in [
        "TOTAL_TESTIGOS",
        "TESTIGOS_MESA_O_REMANENTE",
        "TESTIGOS_COMISION_ESCRUTADORA",
        "BENEFICIARIOS_MESAS_TRABAJO",
        "CANTIDAD_MESAS_TRABAJO_ASOCIADAS",
        "TESTIGOS_DOBLE_ROL",
        "LIDERES",
        "NO_LIDERES",
        "TESTIGOS_CON_REFERIDOS",
        "REFERIDOS_REGISTRADOS",
        "REFERIDOS_INACTIVOS",
        "REFERIDOS_ACTIVOS_ESTIMADOS",
    ]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    for col in [
        "PCT_MESA_O_REMANENTE",
        "PCT_COMISION",
        "PCT_BENEFICIARIOS_MESAS",
        "PCT_LIDERES",
        "PCT_CON_REFERIDOS",
        "REFERIDOS_POR_TESTIGO",
        "REFERIDOS_POR_LIDER",
    ]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def first_existing_path(candidates):
    for candidate in candidates:
        candidate = Path(candidate)
        if candidate.exists() and not candidate.name.startswith("~$"):
            return candidate
    return None


def normalizar_templo_apoyo(value):
    text = strip_accents(str(value or "")).upper().strip()
    text = re.sub(r"^\d+\s*[.-]\s*", "", text)
    if "KENNEDY" in text and "CENTRAL" in text:
        return "KENNEDY CENTRAL"
    if "PATIO" in text and "BONITO" in text:
        return "PATIO BONITO"
    if "CLASS" in text or "CLAS" in text:
        return "CLASS ROMA"
    if "CARVAJAL" in text:
        return "CARVAJAL"
    if "VALLADOLID" in text:
        return "VALLADOLID"
    return "SIN TEMPLO OFICIAL"


def limpiar_columna_apoyo(column):
    text = strip_accents(str(column)).upper().replace("\n", " ").replace("|", "").strip()
    text = re.sub(r"\s+", " ", text)
    mapping = {
        "CORPORACION": "CORPORACION",
        "AÑO": "ANIO",
        "ANO": "ANIO",
        "COD MES": "COD_MES",
        "MES INICIO": "MES_INICIO",
        "CODIGO SEGUIMIENTO": "CODIGO_SEGUIMIENTO",
        "TIPO DE ACTIVIDAD": "TIPO_ACTIVIDAD",
        "TEMA": "TEMA",
        "SUBTEMA": "SUBTEMA",
        "DESCRIPCION Y/O ACTIVIDAD": "DESCRIPCION",
        "BARRIO": "BARRIO",
        "LOCALIDAD": "LOCALIDAD",
        "IGLESIA": "IGLESIA_ORIGINAL",
        "ESTADO ACTUAL DEL TRAMITE": "ESTADO_TRAMITE",
        "RESULTADOS DE LA GESTION": "RESULTADO_GESTION",
        "NOMBRE DEL ASESOR ENCARGADO": "GESTOR",
    }
    return mapping.get(text, text.replace(" ", "_"))


@st.cache_data
def cargar_apoyos_ciudadanos(path: Path | None, mtime: float = 0):
    if path is None or not Path(path).exists():
        return pd.DataFrame()
    df = pd.read_excel(path)
    if df.empty:
        return pd.DataFrame()
    df = df.rename(columns={col: limpiar_columna_apoyo(col) for col in df.columns})
    for col in ["CORPORACION", "MES_INICIO", "TIPO_ACTIVIDAD", "TEMA", "SUBTEMA", "BARRIO", "LOCALIDAD", "IGLESIA_ORIGINAL", "ESTADO_TRAMITE", "RESULTADO_GESTION"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()
    if "CODIGO_SEGUIMIENTO" not in df.columns:
        df["CODIGO_SEGUIMIENTO"] = np.arange(1, len(df) + 1)
    df["ANIO"] = pd.to_numeric(df.get("ANIO", pd.Series(dtype=float)), errors="coerce").astype("Int64")
    df["COD_MES"] = pd.to_numeric(df.get("COD_MES", pd.Series(dtype=float)), errors="coerce").astype("Int64")
    df["TEMPLO"] = df["IGLESIA_ORIGINAL"].map(normalizar_templo_apoyo)
    df["TEMA"] = df["TEMA"].replace("", "Sin tema")
    df["SUBTEMA"] = df["SUBTEMA"].replace("", "Sin subtema")
    df["BARRIO"] = df["BARRIO"].replace("", "Sin dato")
    df["ESTADO_TRAMITE"] = df["ESTADO_TRAMITE"].str.upper().replace({
        "EN PROCESO": "EN PROCESO",
        "FINALIZADO": "FINALIZADO",
        "PENDIENTE": "PENDIENTE",
    })
    df["RESULTADO_GESTION"] = df["RESULTADO_GESTION"].str.upper().replace("", "PENDIENTE")
    df["PERIODO"] = np.where(
        df["ANIO"].notna() & df["COD_MES"].notna(),
        df["ANIO"].astype(str).str.replace(".0", "", regex=False) + "-" + df["COD_MES"].astype(str).str.zfill(2),
        "",
    )
    keep_cols = [
        "CODIGO_SEGUIMIENTO", "CORPORACION", "ANIO", "COD_MES", "MES_INICIO",
        "TIPO_ACTIVIDAD", "TEMA", "SUBTEMA", "BARRIO", "LOCALIDAD", "TEMPLO",
        "IGLESIA_ORIGINAL", "ESTADO_TRAMITE", "RESULTADO_GESTION", "PERIODO",
    ]
    return df[[c for c in keep_cols if c in df.columns]].copy()


def fmt_number(value, decimals=0):
    if pd.isna(value):
        return "N.D."
    number = Decimal(str(value)).quantize(Decimal("1") if decimals == 0 else Decimal(f"1.{'0' * decimals}"), rounding=ROUND_HALF_UP)
    return f"{number:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(value):
    if pd.isna(value):
        return "N.D."
    return f"{value:.1%}".replace(".", ",")


def fmt_variacion(value, pct):
    return f"{fmt_number(value, 0)} ({fmt_pct(pct)})"


TABLE_COLUMN_LABELS = {
    "ACTIVIDAD_ID": "ID actividad",
    "ACTIVIDADES_CAMPANA": "Actividades campaña",
    "ACTIVIDADES_CAMPANA_IGLESIA": "Actividades campaña",
    "ANIO": "Año",
    "APOYOS": "Apoyos",
    "BARRIO": "Barrio",
    "BENEFICIARIOS": "Beneficiarios",
    "BENEFICIARIOS_EXTERNOS": "Beneficiarios externos",
    "BENEFICIARIOS_INTERNOS": "Beneficiarios internos",
    "BENEFICIARIOS_NO_INFOMIRA": "Beneficiarios externos",
    "BENEFICIARIOS_REFERIDOS": "Beneficiarios referidos",
    "CAMARA_2026": "Cámara 2026",
    "CAMBIO_PROPUESTO_TEMPLO": "Cambio de templo",
    "CENSO_2023": "Censo 2023",
    "CODIGO_SEGUIMIENTO": "Código seguimiento",
    "COD_MES": "Mes número",
    "CORPORACION": "Corporación",
    "creado_en": "Fecha",
    "DELTA_PUESTOS": "Delta puestos",
    "DELTA_VOTOS_2026": "Delta votos 2026",
    "DIRECCION": "Dirección",
    "DISTANCIA_ASIGNADA_KM": "Distancia asignada km",
    "DISTANCIA_MINIMA_KM": "Distancia mínima km",
    "entidad": "Entidad",
    "entidad_id": "ID entidad",
    "ESTADO": "Estado",
    "ESTADO_AJUSTE": "Estado ajuste",
    "ESTADO_ASIGNACION": "Estado asignación",
    "ESTADO_TRAMITE": "Estado del trámite",
    "FECHA_CAMBIO": "Fecha cambio",
    "IGLESIA": "Iglesia",
    "IGLESIA_ACTUAL": "Iglesia actual",
    "IGLESIA_HISTORICA_2026": "Iglesia histórica 2026",
    "IGLESIA_ORIGINAL": "Iglesia original",
    "JAL_2023": "JAL 2023",
    "LATITUD": "Latitud",
    "LIDER": "Líder",
    "LONGITUD": "Longitud",
    "MESA_ID": "ID mesa",
    "MES_INICIO": "Mes",
    "MESAS_2026_REPORTE": "Mesas 2026",
    "MESAS_TRABAJO": "Mesas de trabajo",
    "MESAS_TRABAJO_BARRIO": "Mesas de trabajo",
    "MIRA_CONCEJO_2023": "Concejo 2023",
    "motivo": "Nota",
    "NOMBRE_GESTION": "Nombre de la mesa",
    "PCT_BENEFICIARIOS_MESAS": "% beneficiarios mesas",
    "PCT_COMISION": "% comisión",
    "PCT_CON_REFERIDOS": "% con referidos",
    "PCT_LIDERES": "% líderes",
    "PCT_MESA_O_REMANENTE": "% mesa / remanente",
    "PERIODO": "Periodo",
    "PUESTO": "Puesto",
    "PUESTOS": "Puestos",
    "PUESTOS_ASIGNADOS": "Puestos asignados",
    "PUESTO_MAYOR_CAIDA": "Puesto mayor caída",
    "PUESTO_MAYOR_CRECIMIENTO": "Puesto mayor crecimiento",
    "PUESTO_MAYOR_VOTACION": "Puesto mayor votación",
    "REFERIDOS_POR_LIDER": "Referidos por líder",
    "REFERIDOS_POR_TESTIGO": "Referidos por testigo",
    "ROL_ANALITICO": "Rol analítico",
    "SENADO_2026": "Senado 2026",
    "SUBTEMA": "Subtema",
    "TASA_POSITIVA": "Tasa positiva",
    "TEMPLO": "Templo",
    "TEMA": "Tema",
    "TEMPLO_ASIGNADO_FINAL": "Templo vigente",
    "TEMPLO_ASIGNADO_PROPUESTO": "Templo propuesto",
    "TEMPLO_MAS_CERCANO": "Templo más cercano",
    "TEMPLO_OPERATIVO_ACTUAL": "Templo operativo actual",
    "TEMPLO_ORIGINAL": "Templo original",
    "TEMPLO_PROPUESTO": "Templo propuesto",
    "TEMPLO_REPORTE": "Templo reporte",
    "TEMPLO_VIGENTE": "Templo vigente",
    "templo_anterior": "Templo anterior",
    "templo_actual": "Templo actual",
    "templo_nuevo": "Templo nuevo",
    "TESTIGOS_2023_REPORTE": "Testigos 2023",
    "TOTAL_TESTIGOS": "Testigos electorales",
    "TESTIGOS_ELECTORALES": "Testigos electorales",
    "TESTIGOS_LIDERES": "Testigos líderes",
    "TESTIGOS_BENEFICIARIOS_MESAS": "Testigos beneficiarios mesas",
    "BENEFICIARIOS_MESAS_TRABAJO": "Testigos beneficiarios mesas",
    "TIENE_MESA_TRABAJO": "Tiene mesa de trabajo",
    "TIPO_ACTIVIDAD": "Tipo actividad",
    "RESULTADO_GESTION": "Resultado de la gestión",
    "UPZ": "UPZ",
    "usuario": "Usuario",
    "VARIACION_ABSOLUTA": "Variación",
    "VARIACION_PORCENTUAL": "Variación %",
    "VOTOS_2023": "Votos 2023",
    "VOTOS_2026": "Votos 2026",
    "VOTOS_2026_ASIGNADOS": "Votos 2026 asignados",
    "VOTOS_AFINIDAD_E11_2023": "Votos afinidad E11 2023",
}


def prettify_table_column(column):
    text = str(column).strip()
    if text in TABLE_COLUMN_LABELS:
        return TABLE_COLUMN_LABELS[text]
    if text.isupper() or "_" in text:
        text = text.replace("_", " ").strip().lower()
        replacements = {
            "id": "ID",
            "upz": "UPZ",
            "jal": "JAL",
            "mira": "MIRA",
            "roi": "ROI",
            "pct": "%",
            "km": "km",
        }
        words = [replacements.get(word, word.capitalize()) for word in text.split()]
        return " ".join(words)
    return text


def format_percent_cell(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, str):
        text = value.strip()
        return text if text else ""
    try:
        number = float(value)
    except Exception:
        return str(value)
    if abs(number) <= 1.5:
        return fmt_pct(number)
    return f"{number:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def format_table_number(value, column_name):
    if value is None or pd.isna(value):
        return ""
    try:
        number = float(value)
    except Exception:
        return value
    text = strip_accents(str(column_name)).lower()
    if "latitud" in text or "longitud" in text:
        return f"{number:.6f}".replace(".", ",")
    if "distancia" in text or "km" in text:
        return fmt_number(number, 2)
    if "por testigo" in text or "por lider" in text or "por líder" in text:
        return fmt_number(number, 1)
    decimals = 0 if float(number).is_integer() else 1
    return fmt_number(number, decimals)


def format_table_for_display(data):
    if not isinstance(data, pd.DataFrame):
        return data
    df = data.copy()
    for column in df.columns:
        col_key = strip_accents(str(column)).lower()
        is_percent = (
            "porcentual" in col_key
            or col_key.startswith("%")
            or col_key.startswith("pct")
            or col_key.endswith("_pct")
            or col_key.startswith("tasa")
            or " porcentaje" in col_key
        )
        if is_percent:
            df[column] = df[column].map(format_percent_cell)
        elif pd.api.types.is_numeric_dtype(df[column]):
            df[column] = df[column].map(lambda value, col=column: format_table_number(value, col))
        else:
            df[column] = df[column].fillna("").astype(str)
    return df.rename(columns={column: prettify_table_column(column) for column in df.columns})


if not hasattr(st, "_kennedy_original_dataframe"):
    st._kennedy_original_dataframe = st.dataframe


def kennedy_dataframe(data=None, *args, **kwargs):
    return st._kennedy_original_dataframe(format_table_for_display(data), *args, **kwargs)


st.dataframe = kennedy_dataframe


def safe_html(value, default="Sin dato"):
    if value is None or pd.isna(value):
        return default
    text = str(value).strip()
    return html.escape(text) if text else default


def strip_accents(text):
    return "".join(c for c in unicodedata.normalize("NFKD", str(text)) if not unicodedata.combining(c))


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
    if session_key == "ajustes_mesas":
        if "BARRIO_ORIGINAL" not in out.columns:
            out["BARRIO_ORIGINAL"] = out.get("BARRIO", "")
        if "LIDER_ORIGINAL" not in out.columns:
            out["LIDER_ORIGINAL"] = out.get("LIDER", "")
    ajustes = st.session_state.get(session_key, {})
    if ajustes:
        out["IGLESIA"] = out[id_col].map(ajustes).fillna(out["IGLESIA"])
    out["TEMPLO_AJUSTADO"] = out[id_col].isin(ajustes.keys()) if ajustes else False
    if session_key == "ajustes_mesas":
        ajustes_barrio = st.session_state.get("ajustes_mesas_barrio", {})
        ajustes_lider = st.session_state.get("ajustes_mesas_lider", {})
        if ajustes_barrio and "BARRIO" in out.columns:
            out["BARRIO"] = out[id_col].map(ajustes_barrio).fillna(out["BARRIO"])
        if ajustes_lider and "LIDER" in out.columns:
            out["LIDER"] = out[id_col].map(ajustes_lider).fillna(out["LIDER"])
        out["BARRIO_AJUSTADO"] = out[id_col].isin(ajustes_barrio.keys()) if ajustes_barrio else False
        out["LIDER_AJUSTADO"] = out[id_col].isin(ajustes_lider.keys()) if ajustes_lider else False
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


class MapControlHtml(MacroElement):
    def __init__(self, html_content, position="bottomright"):
        super().__init__()
        self._name = "MapControlHtml"
        self.html_content = html_content
        self.position = position
        self._template = Template(
            """
            {% macro script(this, kwargs) %}
            var {{ this.get_name() }} = L.control({position: '{{ this.position }}'});
            {{ this.get_name() }}.onAdd = function (map) {
                var div = L.DomUtil.create('div');
                div.innerHTML = {{ this.html_content|tojson }};
                L.DomEvent.disableClickPropagation(div);
                L.DomEvent.disableScrollPropagation(div);
                return div;
            };
            {{ this.get_name() }}.addTo({{ this._parent.get_name() }});
            {% endmacro %}
            """
        )


def add_map_control_html(mapa, html_content, position="bottomright"):
    mapa.add_child(MapControlHtml(html_content, position=position))


def crear_mapa_base(location=None, zoom_start=13, control_scale=True):
    mapa = folium.Map(
        location=location or KENNEDY_CENTER,
        zoom_start=zoom_start,
        tiles=None,
        control_scale=control_scale,
    )
    folium.TileLayer(
        tiles=BASE_TILE_URL,
        attr=BASE_TILE_ATTR,
        name="Base territorial",
        control=False,
        opacity=0.96,
    ).add_to(mapa)
    return mapa


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
    <div style="background: rgba(255, 255, 255, 0.96); backdrop-filter: blur(8px); padding:12px 14px; border:1px solid #CBD5E1; border-radius:8px; box-shadow:0 8px 22px rgba(15, 23, 42, 0.12); font-size:12px; width: 240px; font-family:'Inter', Arial, sans-serif;">
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
    add_map_control_html(mapa, heat_legend_html, position="bottomleft")
    return heat_config


def render_folium_map(mapa, height=760, key=None):
    folium_static(mapa, width=1600, height=height)


def crear_mapa_asignacion(asignacion_df, iglesias_df, layers_config=None):
    layers_config = layers_config or {}
    show_contorno = layers_config.get("contorno", True)
    show_heat = layers_config.get("heat", True)
    show_lineas = layers_config.get("lineas", True)
    show_puestos = layers_config.get("puestos", True)
    show_templos = layers_config.get("templos", True)

    m = crear_mapa_base(location=KENNEDY_CENTER, zoom_start=13, control_scale=True)
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
        <b>{safe_html(globals().get("metodologia_actual_label", "Votos 2026"))}:</b> {fmt_number(r.get('VOTOS_2026'), 0)}
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
    <div style="background:white;padding:12px 14px;border:1px solid #CBD5E1;border-radius:8px;box-shadow:0 4px 14px rgba(15,23,42,.16);font-size:12px;color:#0F172A;min-width:210px;font-family:'Inter', sans-serif;">
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
    add_map_control_html(m, legend_html, position="bottomright")
    return m


def exportar_asignacion_excel(asignacion_df, resumen_df, tabla_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        asignacion_df.to_excel(writer, sheet_name="asignacion_detallada", index=False)
        resumen_df.to_excel(writer, sheet_name="resumen_por_templo", index=False)
        tabla_df.to_excel(writer, sheet_name="tabla_puestos_por_templo", index=False)
    return output.getvalue()


def exportar_asignacion_por_templo_excel(asignacion_df, actividades_df, mesas_df, testigos_df=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen_global = crear_resumen_asignacion(asignacion_df)
        resumen_global.to_excel(writer, sheet_name="resumen_global", index=False)
        informes = []
        for templo in TEMPLOS_OFICIALES:
            puestos_t = asignacion_df[asignacion_df["TEMPLO_ASIGNADO_FINAL"].eq(templo)].copy()
            acts_t = actividades_df[actividades_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not actividades_df.empty else pd.DataFrame()
            mesas_t = mesas_df[mesas_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not mesas_df.empty else pd.DataFrame()
            testigos_total, testigos_lideres, testigos_benef_mesas = testigos_metricas_templo(testigos_df, templo)
            informes.append({
                "TEMPLO": templo,
                "PUESTOS": len(puestos_t),
                "VOTOS_2026": pd.to_numeric(puestos_t.get("VOTOS_2026", pd.Series(dtype=float)), errors="coerce").fillna(0).sum(),
                "ACTIVIDADES": len(acts_t),
                "MESAS": len(mesas_t),
                "TESTIGOS": testigos_total,
                "TESTIGOS_LIDERES": testigos_lideres,
                "TESTIGOS_BENEFICIARIOS_MESAS": testigos_benef_mesas,
                "LECTURA": generar_informe_templo_markdown(templo, puestos_t.rename(columns={"TEMPLO_ASIGNADO_FINAL": "IGLESIA"}), acts_t, mesas_t, testigos_df=testigos_df),
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


def ocultar_columnas_priorizacion(df):
    if df is None or df.empty:
        return df
    columnas_excluidas = {
        "PRIORIDAD", "NIVEL_PRIORIDAD", "PUNTAJE_PRIORIDAD", "RAZON_PRIORIDAD",
        "FACTORES_PRIORIDAD", "VARIABLE_CRITICA", "ACCION_RECOMENDADA",
        "TEMPORALIDAD", "RESPONSABLE_SUGERIDO",
    }
    return df.drop(columns=[c for c in columnas_excluidas if c in df.columns]).copy()


def multi_sheet_excel_bytes(sheets):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            clean_name = str(sheet_name)[:31]
            (df if df is not None else pd.DataFrame()).to_excel(writer, sheet_name=clean_name, index=False)
    return output.getvalue()


def generar_informe_territorial(asignacion_df, actividades_df, mesas_df, testigos_df=None):
    resumen_puestos = crear_resumen_asignacion(asignacion_df)
    resumen_operativo = crear_resumen_operativo_por_templo(actividades_df, mesas_df)
    resumen_testigos = preparar_testigos_por_templo(testigos_df, resumen_puestos["TEMPLO"].tolist() if "TEMPLO" in resumen_puestos.columns else TEMPLOS_OFICIALES)
    total_puestos = len(asignacion_df)
    total_actividades = len(actividades_df)
    total_mesas = len(mesas_df)
    total_testigos = sum_numeric(resumen_testigos, "TOTAL_TESTIGOS")
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
        f"- Testigos electorales por templo: {fmt_number(total_testigos, 0)}.",
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
        testigos_row = resumen_testigos[resumen_testigos["TEMPLO"].eq(r["TEMPLO"])]
        testigos_count = float(testigos_row.iloc[0].get("TOTAL_TESTIGOS", 0)) if not testigos_row.empty else 0
        lideres_count = float(testigos_row.iloc[0].get("LIDERES", 0)) if not testigos_row.empty else 0
        lineas.append(
            f"- {r['TEMPLO']}: {fmt_number(puestos_count, 0)} puestos, "
            f"{fmt_number(r['ACTIVIDADES'], 0)} actividades, {fmt_number(r['VOLANTEOS'], 0)} volanteos, "
            f"{fmt_number(r['MESAS_TRABAJO'], 0)} mesas, {fmt_number(r['BENEFICIARIOS_MESAS'], 0)} beneficiarios reportados, "
            f"{fmt_number(testigos_count, 0)} testigos ({fmt_number(lideres_count, 0)} líderes)."
        )

    lineas.extend([
        "",
        "## Lectura metodológica",
        "La asignación de puestos se basa en cercanía geográfica al templo; las actividades y mesas pueden reasignarse temporalmente para discusión territorial.",
        "La propuesta debe revisarse con liderazgo comunitario, capacidad operativa, rutas, barrios relevantes y conocimiento de los equipos locales.",
    ])
    return "\n".join(lineas)


def generar_informe_templo_markdown(templo, puestos_df, actividades_df, mesas_df, resumen_row=None, testigos_df=None):
    puestos_t = puestos_df[puestos_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not puestos_df.empty else pd.DataFrame()
    acts_t = actividades_df[actividades_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not actividades_df.empty else pd.DataFrame()
    mesas_t = mesas_df[mesas_df.get("IGLESIA", pd.Series(dtype=str)).eq(templo)].copy() if not mesas_df.empty else pd.DataFrame()

    votos_2026 = pd.to_numeric(puestos_t.get("VOTOS_2026", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    votos_2023 = pd.to_numeric(puestos_t.get("VOTOS_2023", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    variacion = votos_2026 - votos_2023
    variacion_pct = variacion / votos_2023 if votos_2023 else np.nan
    puestos_caida = puestos_t.assign(_VAR=pd.to_numeric(puestos_t.get("VARIACION_ABSOLUTA", pd.Series(dtype=float)), errors="coerce")).sort_values("_VAR", ascending=True).head(5)
    puestos_crecimiento = puestos_t.assign(_VAR=pd.to_numeric(puestos_t.get("VARIACION_ABSOLUTA", pd.Series(dtype=float)), errors="coerce")).sort_values("_VAR", ascending=False).head(5)
    cobertura_operativa = len(acts_t) + len(mesas_t)
    testigos_total, testigos_lideres, testigos_benef_mesas = testigos_metricas_templo(testigos_df, templo)

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
        f"- {globals().get('metodologia_actual_label', 'Votos 2026')}: {fmt_number(votos_2026, 0)} frente a {fmt_number(votos_2023, 0)} en {globals().get('metodologia_base_label', 'Votos 2023')}.",
        f"- Variación: {fmt_variacion(variacion, variacion_pct)}.",
        f"- Actividades y mesas registradas: {fmt_number(cobertura_operativa, 0)} ({fmt_number(len(acts_t), 0)} actividades, {fmt_number(len(mesas_t), 0)} mesas).",
        f"- Testigos electorales: {fmt_number(testigos_total, 0)}; líderes: {fmt_number(testigos_lideres, 0)}; beneficiarios de mesas: {fmt_number(testigos_benef_mesas, 0)}.",
        "",
        "## Puestos a recuperar",
    ]
    if puestos_caida.empty:
        lineas.append("- Sin puestos con información suficiente de caída.")
    else:
        for _, r in puestos_caida.iterrows():
            lineas.append(f"- {r.get('PUESTO')}: {fmt_number(r.get('VARIACION_ABSOLUTA'), 0)} votos.")

    lineas.extend(["", "## Puestos a consolidar"])
    if puestos_crecimiento.empty:
        lineas.append("- Sin puestos con información suficiente de crecimiento.")
    else:
        for _, r in puestos_crecimiento.iterrows():
            lineas.append(f"- {r.get('PUESTO')}: +{fmt_number(r.get('VARIACION_ABSOLUTA'), 0)} votos; mantener estructura territorial.")

    lineas.extend([
        "",
        "## Recomendación operativa",
        "- Concentrar revisión semanal en puestos con variación negativa.",
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
    cambios = puestos_df[col_serie(puestos_df, "CAMBIO_PROPUESTO_TEMPLO").astype(str).eq("SI")].copy() if not puestos_df.empty else pd.DataFrame()
    top_iglesia = resumen_iglesia_df.sort_values("VOTOS_2026", ascending=False).iloc[0] if not resumen_iglesia_df.empty and "VOTOS_2026" in resumen_iglesia_df.columns else None
    mayor_caida = puestos_df.assign(_VAR=pd.to_numeric(col_serie(puestos_df, "VARIACION_ABSOLUTA", 0), errors="coerce")).sort_values("_VAR", ascending=True).head(5)
    mayor_crecimiento = puestos_df.assign(_VAR=pd.to_numeric(col_serie(puestos_df, "VARIACION_ABSOLUTA", 0), errors="coerce")).sort_values("_VAR", ascending=False).head(5)

    lineas = [
        "# Informe ejecutivo territorial-electoral Kennedy",
        "",
        "## 1. Resumen general",
        f"Kennedy registra {fmt_number(total_2026, 0)} votos 2026 frente a {fmt_number(total_2023, 0)} votos 2023, con una variación de {fmt_variacion(variacion, variacion_pct)}.",
        f"El análisis integra {fmt_number(len(puestos_df), 0)} puestos de votación, {fmt_number(len(actividades_df), 0)} actividades de campaña y {fmt_number(len(mesas_df), 0)} mesas de trabajo.",
        "",
        "## 2. Hallazgos principales",
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
        "- 0-30 días: revisión de puestos con mayor variación negativa, responsables y agenda territorial.",
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


ICON_ALIASES = {
    "📍": "map-pin",
    "📅": "calendar",
    "👥": "users",
    "💾": "save",
    "🧭": "compass",
    "📌": "map-pin",
    "📊": "bar-chart",
    "📈": "trend-up",
    "⚡": "activity",
    "🚦": "status",
    "🗳️": "vote",
    "⭐": "star",
    "🔗": "link",
    "📋": "clipboard",
}


def infer_metric_icon(label, icon=None):
    if icon and icon not in {"auto", "📍"}:
        return ICON_ALIASES.get(icon, icon)
    text = strip_accents(str(label)).lower()
    rules = [
        ("testigos", "clipboard"),
        ("beneficiarios", "users"),
        ("votos", "vote"),
        ("variacion", "trend-up"),
        ("puestos", "map"),
        ("iglesias", "building"),
        ("templos", "building"),
        ("cambios", "shuffle"),
        ("ajustes", "save"),
        ("actividades", "calendar"),
        ("volanteos", "file-text"),
        ("mesas", "users"),
        ("jal", "scale"),
        ("concejo", "scale"),
        ("camara", "landmark"),
        ("senado", "landmark"),
        ("coordenadas", "compass"),
        ("lideres", "star"),
        ("referidos", "link"),
        ("estado", "status"),
        ("roi", "activity"),
    ]
    for needle, value in rules:
        if needle in text:
            return value
    return "bar-chart"


def metric_icon_svg(icon_name):
    paths = {
        "activity": '<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>',
        "bar-chart": '<line x1="12" y1="20" x2="12" y2="10"></line><line x1="18" y1="20" x2="18" y2="4"></line><line x1="6" y1="20" x2="6" y2="16"></line>',
        "building": '<rect x="4" y="3" width="16" height="18" rx="2"></rect><path d="M9 7h1"></path><path d="M14 7h1"></path><path d="M9 12h1"></path><path d="M14 12h1"></path><path d="M10 21v-4h4v4"></path>',
        "calendar": '<rect x="3" y="4" width="18" height="18" rx="2"></rect><path d="M16 2v4"></path><path d="M8 2v4"></path><path d="M3 10h18"></path>',
        "clipboard": '<rect x="8" y="2" width="8" height="4" rx="1"></rect><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"></path>',
        "compass": '<circle cx="12" cy="12" r="10"></circle><path d="M16 8l-2.5 5.5L8 16l2.5-5.5L16 8z"></path>',
        "file-text": '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><path d="M14 2v6h6"></path><path d="M8 13h8"></path><path d="M8 17h6"></path>',
        "landmark": '<path d="M3 10h18"></path><path d="M5 10v8"></path><path d="M9 10v8"></path><path d="M15 10v8"></path><path d="M19 10v8"></path><path d="M4 18h16"></path><path d="M12 3l8 5H4l8-5z"></path>',
        "link": '<path d="M10 13a5 5 0 0 0 7.1 0l2-2a5 5 0 0 0-7.1-7.1l-1.1 1.1"></path><path d="M14 11a5 5 0 0 0-7.1 0l-2 2A5 5 0 0 0 12 20.1l1.1-1.1"></path>',
        "map": '<path d="M9 18l-6 3V6l6-3 6 3 6-3v15l-6 3-6-3z"></path><path d="M9 3v15"></path><path d="M15 6v15"></path>',
        "map-pin": '<path d="M12 21s7-4.4 7-11a7 7 0 1 0-14 0c0 6.6 7 11 7 11z"></path><circle cx="12" cy="10" r="2.5"></circle>',
        "save": '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path><path d="M17 21v-8H7v8"></path><path d="M7 3v5h8"></path>',
        "scale": '<path d="M12 3v18"></path><path d="M5 6h14"></path><path d="M6 6l-3 7h6L6 6z"></path><path d="M18 6l-3 7h6l-3-7z"></path>',
        "shuffle": '<path d="M16 3h5v5"></path><path d="M4 20l17-17"></path><path d="M21 16v5h-5"></path><path d="M15 15l6 6"></path><path d="M4 4l5 5"></path>',
        "star": '<path d="M12 3l2.7 5.5 6.1.9-4.4 4.3 1 6-5.4-2.9-5.4 2.9 1-6-4.4-4.3 6.1-.9L12 3z"></path>',
        "status": '<path d="M4 4h16v6H4z"></path><path d="M4 14h16v6H4z"></path><path d="M8 7h.01"></path><path d="M8 17h.01"></path>',
        "trend-up": '<path d="M3 17l6-6 4 4 8-8"></path><path d="M14 7h7v7"></path>',
        "users": '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M22 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path>',
        "vote": '<path d="M9 12l2 2 4-4"></path><path d="M5 7h14"></path><path d="M5 7l2-4h10l2 4"></path><rect x="3" y="7" width="18" height="14" rx="2"></rect>',
    }
    return f'<div class="metric-icon" aria-hidden="true"><svg viewBox="0 0 24 24">{paths.get(icon_name, paths["bar-chart"])}</svg></div>'


def metric_card(label, value, delta=None, positive=True, icon="auto"):
    delta_html = ""
    if delta is not None:
        cls = "metric-delta-positive" if positive else "metric-delta-negative"
        arrow = "↑" if positive else "↓"
        delta_html = f'<div class="{cls}">{arrow} {delta}</div>'
    icon_html = metric_icon_svg(infer_metric_icon(label, icon))
    st.markdown(
        f"""
        <div class="metric-card">
            <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                <div>
                    <div class="metric-label">{label}</div>
                    <div class="metric-value">{value}</div>
                </div>
                {icon_html}
            </div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def sum_numeric(df, col):
    if df is None or df.empty or col not in df.columns:
        return 0
    return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()


def calcular_beneficiarios_mesas(df):
    total = sum_numeric(df, "BENEFICIARIOS")
    if "BENEFICIARIOS_INTERNOS" in df.columns:
        internos = sum_numeric(df, "BENEFICIARIOS_INTERNOS")
    elif "BENEFICIARIOS_NO_INFOMIRA" in df.columns:
        internos = max(total - sum_numeric(df, "BENEFICIARIOS_NO_INFOMIRA"), 0)
    else:
        internos = sum_numeric(df, "BENEFICIARIOS_REFERIDOS")

    if "BENEFICIARIOS_EXTERNOS" in df.columns:
        externos = sum_numeric(df, "BENEFICIARIOS_EXTERNOS")
    elif "BENEFICIARIOS_NO_INFOMIRA" in df.columns:
        externos = sum_numeric(df, "BENEFICIARIOS_NO_INFOMIRA")
    else:
        externos = max(total - internos, 0)
    return total, internos, externos


def aplicar_metodologia_electoral(df, metodologia_key):
    if df is None or df.empty:
        return df
    cfg = METODOLOGIAS_ELECTORALES.get(metodologia_key, METODOLOGIAS_ELECTORALES[METODOLOGIA_DEFAULT])
    result = df.copy()
    if "VOTOS_2023_BASE" not in result.columns:
        result["VOTOS_2023_BASE"] = pd.to_numeric(result.get("VOTOS_2023", 0), errors="coerce").fillna(0)
    if "VOTOS_2026_BASE" not in result.columns:
        result["VOTOS_2026_BASE"] = pd.to_numeric(result.get("VOTOS_2026", 0), errors="coerce").fillna(0)

    base = pd.to_numeric(result.get(cfg["base_col"], pd.Series(0, index=result.index)), errors="coerce").fillna(0)
    actual = pd.to_numeric(result.get(cfg["actual_col"], pd.Series(0, index=result.index)), errors="coerce").fillna(0)
    result["VOTOS_2023"] = base
    result["VOTOS_2026"] = actual
    result["VARIACION_ABSOLUTA"] = actual - base
    result["VARIACION_PORCENTUAL"] = np.where(base.gt(0), result["VARIACION_ABSOLUTA"] / base, 0)
    result["METODOLOGIA_ELECTORAL"] = cfg["label"]
    result["BASE_ELECTORAL_2023"] = cfg["base_label"]
    result["BASE_ELECTORAL_2026"] = cfg["actual_label"]
    return result


def resumen_metodologia_electoral(metodologia_key):
    cfg = METODOLOGIAS_ELECTORALES.get(metodologia_key, METODOLOGIAS_ELECTORALES[METODOLOGIA_DEFAULT])
    return cfg["label"], cfg["caption"], cfg["base_label"], cfg["actual_label"]


TESTIGOS_COLUMNS = [
    "TOTAL_TESTIGOS",
    "TESTIGOS_MESA_O_REMANENTE",
    "TESTIGOS_COMISION_ESCRUTADORA",
    "BENEFICIARIOS_MESAS_TRABAJO",
    "CANTIDAD_MESAS_TRABAJO_ASOCIADAS",
    "TESTIGOS_DOBLE_ROL",
    "LIDERES",
    "NO_LIDERES",
    "TESTIGOS_CON_REFERIDOS",
    "REFERIDOS_REGISTRADOS",
    "REFERIDOS_INACTIVOS",
    "REFERIDOS_ACTIVOS_ESTIMADOS",
]


def preparar_testigos_por_templo(testigos_df, templos=None):
    base_cols = ["TEMPLO"] + TESTIGOS_COLUMNS
    if templos is None:
        templos = TEMPLOS_OFICIALES

    if testigos_df is None or testigos_df.empty or "TEMPLO" not in testigos_df.columns:
        df = pd.DataFrame({"TEMPLO": templos})
    else:
        df = testigos_df.copy()
        df["TEMPLO"] = df["TEMPLO"].fillna("").astype(str).str.strip().str.upper()
        df = df.groupby("TEMPLO", as_index=False).sum(numeric_only=True)
        df = pd.DataFrame({"TEMPLO": templos}).merge(df, on="TEMPLO", how="left")

    for col in TESTIGOS_COLUMNS:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["TESTIGOS_ELECTORALES"] = df["TOTAL_TESTIGOS"]
    df["TESTIGOS_LIDERES"] = df["LIDERES"]
    df["TESTIGOS_BENEFICIARIOS_MESAS"] = df["BENEFICIARIOS_MESAS_TRABAJO"]
    return df[base_cols + ["TESTIGOS_ELECTORALES", "TESTIGOS_LIDERES", "TESTIGOS_BENEFICIARIOS_MESAS"]]


def testigos_metricas_templo(testigos_df, templo):
    resumen = preparar_testigos_por_templo(testigos_df, [templo])
    row = resumen.iloc[0] if not resumen.empty else pd.Series(dtype=object)
    total = float(row.get("TOTAL_TESTIGOS", 0) or 0)
    lideres = float(row.get("LIDERES", 0) or 0)
    benef_mesas = float(row.get("BENEFICIARIOS_MESAS_TRABAJO", 0) or 0)
    return total, lideres, benef_mesas


def sumar_testigos_metricas(testigos_df, templos=None):
    resumen = preparar_testigos_por_templo(testigos_df, templos)
    return {
        "TOTAL_TESTIGOS": float(resumen["TOTAL_TESTIGOS"].sum()),
        "LIDERES": float(resumen["LIDERES"].sum()),
        "BENEFICIARIOS_MESAS_TRABAJO": float(resumen["BENEFICIARIOS_MESAS_TRABAJO"].sum()),
    }


def pdf_font(size=24, bold=False):
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


PDF_W, PDF_H = 1240, 1754
PDF_MARGIN = 60
PDF_PRINT_SCALE = 8  # Aumentado para mayor calidad de impresión
PDF_EXPORT_DPI = 600
PDF_EXPORT_QUALITY = 100
PDF_NAVY = "#0B1F3A"
PDF_BLUE = "#2563EB"
PDF_TEAL = "#14B8A6"
PDF_ORANGE = "#F97316"
PDF_PURPLE = "#7C3AED"
PDF_GREEN = "#16A34A"
PDF_TEXT = "#0F172A"
PDF_MUTED = "#64748B"
PDF_BORDER = "#D8E2EF"
PDF_PANEL = "#FFFFFF"
PDF_BG = "#F6F8FC"


def text_width(draw, text, font):
    box = draw.textbbox((0, 0), str(text), font=font)
    return box[2] - box[0]


def line_height(font):
    box = font.getbbox("Ag")
    return max(box[3] - box[1], 14)


def split_long_word(draw, word, font, max_width):
    chunks, current = [], ""
    for ch in str(word):
        candidate = current + ch
        if text_width(draw, candidate, font) <= max_width or not current:
            current = candidate
        else:
            chunks.append(current)
            current = ch
    if current:
        chunks.append(current)
    return chunks


def wrap_text(draw, text, font, max_width, max_lines=None):
    words = str("" if pd.isna(text) else text).split()
    lines, line = [], ""
    for word in words:
        candidate = f"{line} {word}".strip()
        if text_width(draw, candidate, font) <= max_width:
            line = candidate
            continue
        if line:
            lines.append(line)
        if text_width(draw, word, font) <= max_width:
            line = word
        else:
            split_words = split_long_word(draw, word, font, max_width)
            lines.extend(split_words[:-1])
            line = split_words[-1] if split_words else ""
        if max_lines and len(lines) >= max_lines:
            return lines[:max_lines]
    if line:
        lines.append(line)
    lines = lines or [""]
    if max_lines:
        return lines[:max_lines]
    return lines


def fmt_pdf_number(value):
    try:
        if pd.isna(value):
            return ""
        return fmt_number(value, 0)
    except Exception:
        return str(value)


def normalize_pdf_value(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        return fmt_pdf_number(value)
    return str(value)


def draw_page_chrome(draw, title, subtitle=None, eyebrow="Informe territorial por templo", page_w=PDF_W):
    draw.rectangle([0, 0, page_w, 150], fill=PDF_NAVY)
    draw.rectangle([0, 150, page_w, 164], fill=PDF_TEAL)
    draw.text((PDF_MARGIN, 36), eyebrow.upper(), fill="#93C5FD", font=pdf_font(14, True))
    draw.text((PDF_MARGIN, 66), title, fill="#FFFFFF", font=pdf_font(34, True))
    if subtitle:
        draw.text((PDF_MARGIN, 114), subtitle, fill="#D9E8FF", font=pdf_font(17))


def stamp_pdf_footer(page, page_no, total_pages):
    draw = ImageDraw.Draw(page)
    page_w, page_h = page.size
    y = page_h - 58
    draw.line([PDF_MARGIN, y - 18, page_w - PDF_MARGIN, y - 18], fill="#D8E2EF", width=1)
    draw.text((PDF_MARGIN, y), "Dashboard territorial-electoral Kennedy", fill=PDF_MUTED, font=pdf_font(12, True))
    label = f"Página {page_no} de {total_pages}"
    draw.text((page_w - PDF_MARGIN - text_width(draw, label, pdf_font(12, True)), y), label, fill=PDF_MUTED, font=pdf_font(12, True))


def new_pdf_page(title, subtitle=None, eyebrow="Informe territorial por templo", size=(PDF_W, PDF_H)):
    page = Image.new("RGB", size, PDF_BG)
    draw = ImageDraw.Draw(page)
    draw_page_chrome(draw, title, subtitle, eyebrow, page_w=size[0])
    return page, draw


def draw_section_title(draw, x, y, title, accent=PDF_BLUE, subtitle=None):
    draw.rounded_rectangle([x, y + 4, x + 8, y + 40], radius=4, fill=accent)
    draw.text((x + 22, y), title, fill=PDF_TEXT, font=pdf_font(25, True))
    if subtitle:
        draw.text((x + 22, y + 36), subtitle, fill=PDF_MUTED, font=pdf_font(13))


def draw_pdf_card(draw, xy, title, value, width=245, height=92, accent=PDF_BLUE):
    x, y = xy
    draw.rounded_rectangle([x + 3, y + 5, x + width + 3, y + height + 5], radius=16, fill="#E6EDF7")
    draw.rounded_rectangle([x, y, x + width, y + height], radius=16, fill=PDF_PANEL, outline=PDF_BORDER, width=1)
    draw.rounded_rectangle([x, y, x + 8, y + height], radius=5, fill=accent)
    draw.text((x + 22, y + 17), str(title), fill="#42526B", font=pdf_font(16, True))
    value_font = pdf_font(28, True)
    if text_width(draw, str(value), value_font) > width - 44:
        value_font = pdf_font(25, True)
    value_lines = wrap_text(draw, value, value_font, width - 44, max_lines=2)
    if len(value_lines) > 1 and height < 100:
        value_font = pdf_font(23, True)
        value_lines = wrap_text(draw, value, value_font, width - 44, max_lines=2)
    yy = y + 47 if len(value_lines) == 1 else y + 42
    for line in value_lines:
        draw.text((x + 22, yy), line, fill=PDF_TEXT, font=value_font)
        yy += line_height(value_font) + 6


def draw_pdf_table_header(draw, x, y, widths, headers, accent=PDF_BLUE):
    header_h = 50
    cursor = x
    for header, width in zip(headers, widths):
        draw.rectangle([cursor, y, cursor + width, y + header_h], fill=accent, outline=accent)
        header_font = pdf_font(12, True) if text_width(draw, header, pdf_font(14, True)) > width - 16 else pdf_font(14, True)
        header_lines = wrap_text(draw, header, header_font, width - 16, max_lines=2)
        yy = y + 8 if len(header_lines) > 1 else y + 15
        for line in header_lines:
            draw.text((cursor + 8, yy), str(line), fill="#FFFFFF", font=header_font)
            yy += line_height(header_font) + 1
        cursor += width
    return y + header_h


def draw_paginated_pdf_table(pages, title, df, columns, widths, accent=PDF_BLUE, subtitle=None, empty_text="Sin registros.", body_font_size=13, max_lines_per_cell=3):
    df = df.copy() if df is not None else pd.DataFrame()
    body_font = pdf_font(body_font_size)
    row_gap = 0
    table_x = PDF_MARGIN
    table_w = sum(widths)
    max_y = PDF_H - 92
    page_no = 1
    page, draw = new_pdf_page(title, subtitle, eyebrow="Detalle territorial")
    pages.append(page)
    draw_section_title(draw, table_x, 198, title, accent, subtitle)
    y = 265 if subtitle else 245
    draw.rounded_rectangle([table_x, y, table_x + table_w, y + 46], radius=10, fill=accent)
    y = draw_pdf_table_header(draw, table_x, y, widths, columns, accent)

    if df.empty:
        draw.rounded_rectangle([table_x, y, table_x + table_w, y + 78], radius=8, fill="#FFFFFF", outline=PDF_BORDER)
        draw.text((table_x + 18, y + 26), empty_text, fill=PDF_MUTED, font=pdf_font(14, True))
        return

    for row_idx, (_, row) in enumerate(df.iterrows()):
        line_sets = []
        for col, width in zip(columns, widths):
            lines = wrap_text(draw, normalize_pdf_value(row.get(col, "")), body_font, width - 20, max_lines=max_lines_per_cell)
            line_sets.append(lines)
        row_h = max(42, max(len(lines) for lines in line_sets) * (line_height(body_font) + 3) + 18)
        if y + row_h > max_y:
            page_no += 1
            page, draw = new_pdf_page(title, f"Continuación {page_no}", eyebrow="Detalle territorial")
            pages.append(page)
            draw_section_title(draw, table_x, 198, title, accent, f"Continuación {page_no}")
            y = 265
            draw.rounded_rectangle([table_x, y, table_x + table_w, y + 46], radius=10, fill=accent)
            y = draw_pdf_table_header(draw, table_x, y, widths, columns, accent)
        fill = "#FFFFFF" if row_idx % 2 == 0 else "#F8FBFF"
        cursor = table_x
        for col_idx, (col, width, lines) in enumerate(zip(columns, widths, line_sets)):
            draw.rectangle([cursor, y, cursor + width, y + row_h], fill=fill, outline="#E6EDF5")
            yy = y + 9
            for line in lines:
                color = PDF_TEXT if col_idx == 0 else "#334155"
                draw.text((cursor + 10, yy), line, fill=color, font=body_font)
                yy += line_height(body_font) + 3
            cursor += width
        y += row_h + row_gap


def collect_temple_map_points(iglesia, templo_row, puestos_df, mesas_df):
    points = []
    if templo_row is not None and pd.notna(templo_row.get("LATITUD")) and pd.notna(templo_row.get("LONGITUD")):
        points.append(("templo", float(templo_row.get("LATITUD")), float(templo_row.get("LONGITUD")), iglesia))
    if puestos_df is not None and {"LATITUD", "LONGITUD"}.issubset(puestos_df.columns):
        coords_puestos = puestos_df.copy()
        coords_puestos["LATITUD"] = pd.to_numeric(coords_puestos["LATITUD"], errors="coerce")
        coords_puestos["LONGITUD"] = pd.to_numeric(coords_puestos["LONGITUD"], errors="coerce")
        for _, r in coords_puestos.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
            points.append(("puesto", float(r["LATITUD"]), float(r["LONGITUD"]), r.get("PUESTO", "")))
    if mesas_df is not None and {"LATITUD", "LONGITUD"}.issubset(mesas_df.columns):
        coords_mesas = mesas_df.copy()
        coords_mesas["LATITUD"] = pd.to_numeric(coords_mesas["LATITUD"], errors="coerce")
        coords_mesas["LONGITUD"] = pd.to_numeric(coords_mesas["LONGITUD"], errors="coerce")
        for _, r in coords_mesas.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
            points.append(("mesa", float(r["LATITUD"]), float(r["LONGITUD"]), r.get("NOMBRE_GESTION", r.get("TEMA", ""))))
    return points


def latlon_to_world_pixel(lat, lon, zoom):
    lat = max(min(float(lat), 85.05112878), -85.05112878)
    lon = float(lon)
    scale = 256 * (2 ** zoom)
    sin_lat = math.sin(math.radians(lat))
    x = (lon + 180.0) / 360.0 * scale
    y = (0.5 - math.log((1 + sin_lat) / (1 - sin_lat)) / (4 * math.pi)) * scale
    return x, y


@st.cache_data(show_spinner=False, ttl=86400)
def fetch_cartocdn_tile(zoom, x, y):
    local_tile = DATA_DIR / "map_tiles" / "voyager" / str(zoom) / str(x) / f"{y}.png"
    if local_tile.exists():
        return local_tile.read_bytes()
    url = BASE_TILE_URL.format(s="a", z=zoom, x=x, y=y, r="")
    request = urllib.request.Request(url, headers={"User-Agent": "KennedyMiraDashboard/1.0"})
    try:
        with urllib.request.urlopen(request, timeout=5) as response:
            return response.read()
    except Exception:
        result = subprocess.run(
            ["curl", "-L", "--fail", "--silent", "--max-time", "6", url],
            capture_output=True,
            check=True,
        )
        return result.stdout


def choose_tile_zoom(points, width, height):
    if not points:
        return 14
    for zoom in range(15, 11, -1):
        coords = [latlon_to_world_pixel(lat, lon, zoom) for _, lat, lon, _ in points]
        xs = [p[0] for p in coords]
        ys = [p[1] for p in coords]
        if max(xs) - min(xs) <= width * 1.45 and max(ys) - min(ys) <= height * 1.45:
            return zoom
    return 12


def render_tile_map_image(points, width, height):
    if not points:
        return None, None
    preferred_zoom = choose_tile_zoom(points, width, height)
    for zoom in range(preferred_zoom, 11, -1):
        coords = [latlon_to_world_pixel(lat, lon, zoom) for _, lat, lon, _ in points]
        xs = [p[0] for p in coords]
        ys = [p[1] for p in coords]
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)
        bbox_w = max(max_x - min_x, 96)
        bbox_h = max(max_y - min_y, 96)
        aspect = width / max(height, 1)
        view_w = bbox_w / 0.86
        view_h = bbox_h / 0.82
        if view_w / max(view_h, 1) < aspect:
            view_w = view_h * aspect
        else:
            view_h = view_w / aspect

        center_x = (min_x + max_x) / 2
        center_y = (min_y + max_y) / 2
        left = center_x - view_w / 2
        top = center_y - view_h / 2
        right = left + view_w
        bottom = top + view_h
        first_tile_x = math.floor(left / 256)
        first_tile_y = math.floor(top / 256)
        last_tile_x = math.floor(right / 256)
        last_tile_y = math.floor(bottom / 256)
        tile_count = 2 ** zoom
        stitched = Image.new(
            "RGB",
            ((last_tile_x - first_tile_x + 1) * 256, (last_tile_y - first_tile_y + 1) * 256),
            "#F1F5F9",
        )
        fetched = 0

        for tx in range(first_tile_x, last_tile_x + 1):
            for ty in range(first_tile_y, last_tile_y + 1):
                if ty < 0 or ty >= tile_count:
                    continue
                try:
                    tile_bytes = fetch_cartocdn_tile(zoom, tx % tile_count, ty)
                    tile = Image.open(BytesIO(tile_bytes)).convert("RGB")
                    if tile.size != (256, 256):
                        tile = tile.resize((256, 256), Image.Resampling.LANCZOS)
                    fetched += 1
                except Exception:
                    tile = Image.new("RGB", (256, 256), "#EDF2F7")
                stitched.paste(tile, ((tx - first_tile_x) * 256, (ty - first_tile_y) * 256))

        if fetched == 0:
            continue

        crop_left = int(left - first_tile_x * 256)
        crop_top = int(top - first_tile_y * 256)
        cropped = stitched.crop((crop_left, crop_top, crop_left + int(view_w), crop_top + int(view_h)))
        map_img = cropped.resize((width, height), Image.Resampling.LANCZOS).convert("RGBA")
        scale_x = width / view_w
        scale_y = height / view_h

        def project(lat, lon, current_zoom=zoom, crop_left_world=left, crop_top_world=top):
            px, py = latlon_to_world_pixel(lat, lon, current_zoom)
            return (px - crop_left_world) * scale_x, (py - crop_top_world) * scale_y

        return map_img, project

    return None, None


def draw_real_temple_map(page, draw, x, y, w, h, iglesia, templo_row, puestos_df, mesas_df):
    points = collect_temple_map_points(iglesia, templo_row, puestos_df, mesas_df)
    if not points:
        return False
    map_img, project = render_tile_map_image(points, int(w), int(h))
    if map_img is None or project is None:
        return False

    overlay = Image.new("RGBA", map_img.size, (255, 255, 255, 0))
    odraw = ImageDraw.Draw(overlay)
    temple_xy = None
    for kind, lat, lon, _ in points:
        if kind == "templo":
            temple_xy = project(lat, lon)
            break

    if temple_xy:
        for kind, lat, lon, _ in points:
            if kind in {"puesto", "mesa"}:
                px, py = project(lat, lon)
                color = (37, 99, 235, 80) if kind == "puesto" else (249, 115, 22, 95)
                odraw.line([temple_xy[0], temple_xy[1], px, py], fill=color, width=3)

    label_font = pdf_font(17, True)
    small_font = pdf_font(12, True)
    for kind, lat, lon, name in points:
        px, py = project(lat, lon)
        if kind == "templo":
            r = 16
            odraw.ellipse([px - r - 3, py - r - 3, px + r + 3, py + r + 3], fill=(255, 255, 255, 235))
            odraw.ellipse([px - r, py - r, px + r, py + r], fill=(124, 58, 237, 245), outline=(255, 255, 255, 255), width=4)
            odraw.text((px + 22, py - 12), str(iglesia), fill=(15, 23, 42, 255), font=label_font)
        elif kind == "mesa":
            r = 13
            odraw.ellipse([px - r - 2, py - r - 2, px + r + 2, py + r + 2], fill=(255, 255, 255, 235))
            odraw.ellipse([px - r, py - r, px + r, py + r], fill=(249, 115, 22, 245), outline=(255, 255, 255, 255), width=3)
            odraw.text((px - 5, py - 7), "M", fill=(255, 255, 255, 255), font=small_font)
        else:
            r = 8
            odraw.ellipse([px - r - 2, py - r - 2, px + r + 2, py + r + 2], fill=(255, 255, 255, 230))
            odraw.ellipse([px - r, py - r, px + r, py + r], fill=(37, 99, 235, 225), outline=(255, 255, 255, 255), width=2)

    map_img = Image.alpha_composite(map_img, overlay).convert("RGB")
    page.paste(map_img, (int(x), int(y)))
    draw.rounded_rectangle([x, y, x + w, y + h], radius=18, outline="#CBD5E1", width=2)
    draw.rectangle([x + 16, y + h - 50, x + 420, y + h - 12], fill="#FFFFFF", outline="#E2E8F0")
    legend_items = [("Puestos", PDF_BLUE), ("Mesas", PDF_ORANGE), ("Templo", PDF_PURPLE)]
    lx = x + 30
    ly = y + h - 39
    for label, color in legend_items:
        draw.ellipse([lx, ly + 3, lx + 14, ly + 17], fill=color, outline="#FFFFFF", width=2)
        draw.text((lx + 22, ly), label, fill="#334155", font=pdf_font(12, True))
        lx += 120
    count_label = f"{len(puestos_df) if puestos_df is not None else 0} puestos | {len(mesas_df) if mesas_df is not None else 0} mesas"
    count_font = pdf_font(12, True)
    draw.rectangle([x + w - 188, y + h - 50, x + w - 16, y + h - 12], fill="#FFFFFF", outline="#E2E8F0")
    draw.text((x + w - 102 - text_width(draw, count_label, count_font) / 2, ly), count_label, fill="#334155", font=count_font)
    draw.text((x + w - 380, y + h - 29), "© OpenStreetMap contributors © CARTO", fill="#334155", font=pdf_font(10, True))
    return True


def draw_static_temple_map(draw, x, y, w, h, iglesia, templo_row, puestos_df, mesas_df):
    draw.rounded_rectangle([x + 4, y + 6, x + w + 4, y + h + 6], radius=18, fill="#E2EAF5")
    draw.rounded_rectangle([x, y, x + w, y + h], radius=18, fill="#F2F7FC", outline=PDF_BORDER, width=1)
    for i in range(1, 7):
        gx = x + i * w / 7
        gy = y + i * h / 7
        draw.line([gx, y + 18, gx, y + h - 18], fill="#DDE7F2", width=1)
        draw.line([x + 18, gy, x + w - 18, gy], fill="#DDE7F2", width=1)
    for offset, color, width_line in [(0, "#D5E2F0", 5), (45, "#E8EFF7", 3), (90, "#D5E2F0", 4)]:
        draw.line([x + 40, y + h - 170 - offset, x + w * 0.38, y + h * 0.55 - offset, x + w - 55, y + 135 + offset], fill=color, width=width_line)
        draw.line([x + 95 + offset, y + 70, x + w * 0.52, y + h * 0.50, x + w - 160 + offset, y + h - 90], fill=color, width=max(2, width_line - 1))

    points = collect_temple_map_points(iglesia, templo_row, puestos_df, mesas_df)

    if not points:
        draw.text((x + 28, y + 28), "Sin coordenadas para mapa", fill=PDF_MUTED, font=pdf_font(16, True))
        return

    lats = [p[1] for p in points]
    lons = [p[2] for p in points]
    min_lat, max_lat = min(lats), max(lats)
    min_lon, max_lon = min(lons), max(lons)
    lat_pad = max((max_lat - min_lat) * 0.15, 0.005)
    lon_pad = max((max_lon - min_lon) * 0.15, 0.005)
    min_lat, max_lat = min_lat - lat_pad, max_lat + lat_pad
    min_lon, max_lon = min_lon - lon_pad, max_lon + lon_pad

    def project(lat, lon):
        px = x + 54 + (lon - min_lon) / max(max_lon - min_lon, 1e-6) * (w - 108)
        py = y + h - 58 - (lat - min_lat) / max(max_lat - min_lat, 1e-6) * (h - 126)
        return px, py

    temple_xy = None
    for kind, lat, lon, name in points:
        px, py = project(lat, lon)
        if kind == "templo":
            temple_xy = (px, py)
            break
    if temple_xy:
        for kind, lat, lon, name in points:
            if kind in {"puesto", "mesa"}:
                px, py = project(lat, lon)
                draw.line([temple_xy[0], temple_xy[1], px, py], fill="#C5D3E4", width=2 if kind == "mesa" else 1)

    for kind, lat, lon, name in points:
        px, py = project(lat, lon)
        if kind == "templo":
            draw.ellipse([px - 15, py - 15, px + 15, py + 15], fill=PDF_PURPLE, outline="#FFFFFF", width=4)
            draw.text((px + 18, py - 10), str(iglesia), fill=PDF_TEXT, font=pdf_font(13, True))
        elif kind == "mesa":
            draw.ellipse([px - 8, py - 8, px + 8, py + 8], fill=PDF_ORANGE, outline="#FFFFFF", width=3)
        else:
            draw.ellipse([px - 6, py - 6, px + 6, py + 6], fill=PDF_BLUE, outline="#FFFFFF", width=2)
    legend_y = y + h - 43
    legend_items = [("Puestos", PDF_BLUE), ("Mesas", PDF_ORANGE), ("Templo", PDF_PURPLE)]
    lx = x + 22
    for label, color in legend_items:
        draw.ellipse([lx, legend_y + 4, lx + 14, legend_y + 18], fill=color, outline="#FFFFFF", width=2)
        draw.text((lx + 22, legend_y), label, fill="#334155", font=pdf_font(13, True))
        lx += 140
    count_label = f"{len(puestos_df) if puestos_df is not None else 0} puestos | {len(mesas_df) if mesas_df is not None else 0} mesas"
    draw.text((x + w - 22 - text_width(draw, count_label, pdf_font(13, True)), legend_y), count_label, fill=PDF_MUTED, font=pdf_font(13, True))


def collect_general_map_points(puestos_df, iglesias_df, actividades_df=None, mesas_df=None, mode="puestos"):
    points = []
    if iglesias_df is not None and not iglesias_df.empty and {"LATITUD", "LONGITUD"}.issubset(iglesias_df.columns):
        templos = iglesias_df[iglesias_df["IGLESIA"].isin(TEMPLOS_OFICIALES)].copy()
        templos["LATITUD"] = pd.to_numeric(templos["LATITUD"], errors="coerce")
        templos["LONGITUD"] = pd.to_numeric(templos["LONGITUD"], errors="coerce")
        for _, r in templos.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
            points.append(("templo", float(r["LATITUD"]), float(r["LONGITUD"]), r.get("IGLESIA", "")))
    if mode in {"rango", "puestos"} and puestos_df is not None and {"LATITUD", "LONGITUD"}.issubset(puestos_df.columns):
        coords = puestos_df.copy()
        coords["LATITUD"] = pd.to_numeric(coords["LATITUD"], errors="coerce")
        coords["LONGITUD"] = pd.to_numeric(coords["LONGITUD"], errors="coerce")
        for _, r in coords.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
            points.append(("puesto", float(r["LATITUD"]), float(r["LONGITUD"]), r.get("PUESTO", "")))
    if mode == "actividades" and actividades_df is not None and {"LATITUD", "LONGITUD"}.issubset(actividades_df.columns):
        coords = actividades_df.copy()
        coords["LATITUD"] = pd.to_numeric(coords["LATITUD"], errors="coerce")
        coords["LONGITUD"] = pd.to_numeric(coords["LONGITUD"], errors="coerce")
        for _, r in coords.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
            points.append(("actividad", float(r["LATITUD"]), float(r["LONGITUD"]), r.get("TIPO_ACTIVIDAD", "")))
    if mode == "mesas" and mesas_df is not None and {"LATITUD", "LONGITUD"}.issubset(mesas_df.columns):
        coords = mesas_df.copy()
        coords["LATITUD"] = pd.to_numeric(coords["LATITUD"], errors="coerce")
        coords["LONGITUD"] = pd.to_numeric(coords["LONGITUD"], errors="coerce")
        for _, r in coords.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
            points.append(("mesa", float(r["LATITUD"]), float(r["LONGITUD"]), r.get("NOMBRE_GESTION", r.get("TEMA", ""))))
    return points


def _geojson_coordinate_lines(geojson_obj):
    lines = []

    def parse_coords(coords):
        if not isinstance(coords, list):
            return
        if coords and isinstance(coords[0], (int, float)):
            return
        if coords and isinstance(coords[0], list) and coords[0] and isinstance(coords[0][0], (int, float)):
            line = []
            for coord in coords:
                if len(coord) >= 2:
                    lon, lat = coord[:2]
                    if -75 < lon < -73 and 3 < lat < 6:
                        line.append((lat, lon))
            if len(line) >= 2:
                lines.append(line)
            return
        for item in coords:
            parse_coords(item)

    for feature in (geojson_obj or {}).get("features", []):
        parse_coords((feature.get("geometry") or {}).get("coordinates", []))
    return lines


def draw_pdf_contours(odraw, project):
    geojson_obj = cargar_geojson(LOCALIDADES_GEOJSON)
    if not geojson_obj:
        return
    for line in _geojson_coordinate_lines(geojson_obj):
        pts = [project(lat, lon) for lat, lon in line]
        if len(pts) >= 2:
            odraw.line(pts, fill=(15, 23, 42, 130), width=3, joint="curve")


def vote_range_color(value):
    try:
        votes = float(value)
    except Exception:
        votes = 0
    if votes >= 150:
        return (220, 38, 38, 235)
    if votes >= 100:
        return (249, 115, 22, 235)
    if votes >= 60:
        return (37, 99, 235, 230)
    return (20, 184, 166, 230)


def draw_general_map(page, draw, x, y, w, h, mode, puestos_df, iglesias_df, actividades_df, mesas_df):
    points = collect_general_map_points(puestos_df, iglesias_df, actividades_df, mesas_df, mode=mode)
    if not points:
        draw.rounded_rectangle([x, y, x + w, y + h], radius=18, fill="#F2F7FC", outline=PDF_BORDER)
        draw.text((x + 28, y + 28), "Sin coordenadas disponibles para este mapa.", fill=PDF_MUTED, font=pdf_font(17, True))
        return False

    map_img, project = render_tile_map_image(points, int(w), int(h))
    if map_img is None or project is None:
        draw.rounded_rectangle([x, y, x + w, y + h], radius=18, fill="#F2F7FC", outline=PDF_BORDER)
        draw.text((x + 28, y + 28), "No se pudo cargar el mapa base; revise la conexión o caché de teselas.", fill=PDF_MUTED, font=pdf_font(17, True))
        return False

    overlay = Image.new("RGBA", map_img.size, (255, 255, 255, 0))
    odraw = ImageDraw.Draw(overlay)
    draw_pdf_contours(odraw, project)

    temple_coords = {}
    if iglesias_df is not None and not iglesias_df.empty:
        templos = iglesias_df[iglesias_df["IGLESIA"].isin(TEMPLOS_OFICIALES)].dropna(subset=["LATITUD", "LONGITUD"]).copy()
        for _, r in templos.iterrows():
            temple_coords[r.get("IGLESIA")] = project(float(r["LATITUD"]), float(r["LONGITUD"]))

    if mode in {"puestos", "rango"} and puestos_df is not None and not puestos_df.empty:
        for _, r in puestos_df.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
            templo = r.get("TEMPLO_ASIGNADO_FINAL", r.get("IGLESIA"))
            if templo in temple_coords:
                px, py = project(float(r["LATITUD"]), float(r["LONGITUD"]))
                line_color = COLORES_TEMPLOS.get(templo, "#64748B")
                rgb = tuple(int(line_color.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
                odraw.line([temple_coords[templo][0], temple_coords[templo][1], px, py], fill=rgb + (70,), width=2)

    font_marker = pdf_font(11, True)
    for _, r in (puestos_df if puestos_df is not None else pd.DataFrame()).dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        if mode not in {"puestos", "rango"}:
            continue
        px, py = project(float(r["LATITUD"]), float(r["LONGITUD"]))
        if mode == "rango":
            fill = vote_range_color(r.get("VOTOS_2026"))
            radius = max(7, min(18, float(pd.to_numeric(r.get("VOTOS_2026"), errors="coerce") or 0) / 11))
        else:
            color = COLORES_TEMPLOS.get(r.get("TEMPLO_ASIGNADO_FINAL", r.get("IGLESIA")), PDF_BLUE)
            fill = tuple(int(color.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)) + (235,)
            radius = 8
        odraw.ellipse([px - radius - 2, py - radius - 2, px + radius + 2, py + radius + 2], fill=(255, 255, 255, 235))
        odraw.ellipse([px - radius, py - radius, px + radius, py + radius], fill=fill, outline=(255, 255, 255, 255), width=2)

    if mode == "actividades" and actividades_df is not None and not actividades_df.empty:
        for _, r in actividades_df.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
            px, py = project(float(r["LATITUD"]), float(r["LONGITUD"]))
            odraw.rounded_rectangle([px - 10, py - 10, px + 10, py + 10], radius=4, fill=(124, 58, 237, 235), outline=(255, 255, 255, 255), width=2)
            odraw.text((px - 4, py - 7), "A", fill=(255, 255, 255, 255), font=font_marker)

    if mode == "mesas" and mesas_df is not None and not mesas_df.empty:
        for _, r in mesas_df.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
            px, py = project(float(r["LATITUD"]), float(r["LONGITUD"]))
            odraw.ellipse([px - 12, py - 12, px + 12, py + 12], fill=(249, 115, 22, 235), outline=(255, 255, 255, 255), width=3)
            odraw.text((px - 5, py - 7), "M", fill=(255, 255, 255, 255), font=font_marker)

    label_font = pdf_font(13, True)
    for name, xy in temple_coords.items():
        color = COLORES_TEMPLOS.get(name, PDF_PURPLE)
        fill = tuple(int(color.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)) + (245,)
        px, py = xy
        odraw.ellipse([px - 15, py - 15, px + 15, py + 15], fill=(255, 255, 255, 245))
        odraw.ellipse([px - 12, py - 12, px + 12, py + 12], fill=fill, outline=(255, 255, 255, 255), width=3)
        odraw.text((px + 16, py - 10), str(name), fill=(15, 23, 42, 245), font=label_font)

    page.paste(Image.alpha_composite(map_img, overlay).convert("RGB"), (int(x), int(y)))
    draw.rounded_rectangle([x, y, x + w, y + h], radius=18, outline="#CBD5E1", width=2)
    return True


def draw_map_legend_panel(draw, x, y, w, title, body, items):
    draw.rounded_rectangle([x, y, x + w, y + 118], radius=14, fill="#F8FEFF", outline="#BDEFEA")
    draw.text((x + 22, y + 18), title, fill=PDF_TEXT, font=pdf_font(17, True))
    yy = y + 46
    for line in wrap_text(draw, body, pdf_font(12), w - 44, max_lines=2):
        draw.text((x + 22, yy), line, fill="#475569", font=pdf_font(12))
        yy += 17
    lx = x + 22
    ly = y + 92
    for label, color in items:
        draw.ellipse([lx, ly, lx + 12, ly + 12], fill=color, outline="#FFFFFF", width=2)
        draw.text((lx + 18, ly - 2), label, fill="#334155", font=pdf_font(11, True))
        lx += max(118, text_width(draw, label, pdf_font(11, True)) + 42)


def build_general_pdf_map_page(title, mode, puestos_df, iglesias_df, actividades_df, mesas_df, legend_title, legend_body, legend_items):
    page, draw = new_pdf_page(
        title,
        "Mapa de lectura territorial para Kennedy con capas activas de contorno, templos y operación.",
        eyebrow="Informe electoral general",
        size=(PDF_H, PDF_W),
    )
    page_w, _ = page.size
    draw_section_title(draw, PDF_MARGIN, 202, title, PDF_TEAL)
    map_x, map_y = PDF_MARGIN, 270
    map_w, map_h = page_w - PDF_MARGIN * 2, 760
    draw_general_map(page, draw, map_x, map_y, map_w, map_h, mode, puestos_df, iglesias_df, actividades_df, mesas_df)
    draw_map_legend_panel(draw, PDF_MARGIN, 1055, page_w - PDF_MARGIN * 2, legend_title, legend_body, legend_items)
    return page


def draw_pdf_horizontal_bars(draw, x, y, w, h, data, label_col, value_col, title, color=PDF_BLUE):
    draw.rounded_rectangle([x, y, x + w, y + h], radius=16, fill="#FFFFFF", outline=PDF_BORDER)
    draw.text((x + 22, y + 18), title, fill=PDF_TEXT, font=pdf_font(18, True))
    if data is None or data.empty:
        draw.text((x + 22, y + 58), "Sin datos disponibles.", fill=PDF_MUTED, font=pdf_font(13, True))
        return
    df = data.copy().head(8)
    values = pd.to_numeric(df[value_col], errors="coerce").fillna(0)
    max_value = max(float(values.abs().max()), 1)
    bar_x = x + 220
    bar_w = w - 300
    row_h = max(28, (h - 66) / max(len(df), 1))
    for idx, (_, row) in enumerate(df.iterrows()):
        yy = y + 58 + idx * row_h
        label = str(row.get(label_col, ""))[:26]
        val = float(pd.to_numeric(row.get(value_col), errors="coerce") or 0)
        draw.text((x + 22, yy + 4), label, fill="#334155", font=pdf_font(12, True))
        draw.rounded_rectangle([bar_x, yy + 7, bar_x + bar_w, yy + 22], radius=7, fill="#E8EEF7")
        fill_w = max(3, abs(val) / max_value * bar_w)
        bar_color = PDF_ORANGE if val < 0 else color
        draw.rounded_rectangle([bar_x, yy + 7, bar_x + fill_w, yy + 22], radius=7, fill=bar_color)
        value_label = fmt_number(val, 0)
        draw.text((bar_x + bar_w + 12, yy + 2), value_label, fill=PDF_TEXT, font=pdf_font(12, True))


def build_general_pdf_summary_page(metricas, resumen_templo, top_crecimiento, top_caida):
    page, draw = new_pdf_page(
        "Informe electoral general Kennedy",
        "Resumen ejecutivo, mapas territoriales, graficas y tablas de seguimiento.",
        eyebrow="Informe electoral general",
    )
    draw.rounded_rectangle([PDF_MARGIN, 210, PDF_W - PDF_MARGIN, 342], radius=18, fill="#FFFFFF", outline=PDF_BORDER)
    draw.text((PDF_MARGIN + 28, 236), "Resumen ejecutivo Kennedy", fill=PDF_TEXT, font=pdf_font(25, True))
    draw.text(
        (PDF_MARGIN + 28, 280),
        "El informe consolida votación, operación territorial, mesas de trabajo, actividades, beneficiarios y asignación vigente por templo.",
        fill="#334155",
        font=pdf_font(15),
    )
    card_w, card_h = 255, 92
    start_y = 382
    for idx, (label, value) in enumerate(metricas[:12]):
        row, col = divmod(idx, 4)
        draw_pdf_card(draw, (PDF_MARGIN + col * (card_w + 28), start_y + row * 112), label, value, width=card_w, height=card_h, accent=metric_accent(label))

    draw_section_title(draw, PDF_MARGIN, 742, "Graficas ejecutivas", PDF_BLUE)
    chart_top = top_crecimiento[["PUESTO", "VARIACION_ABSOLUTA"]].rename(columns={"PUESTO": "Puesto", "VARIACION_ABSOLUTA": "Variación"})
    chart_drop = top_caida[["PUESTO", "VARIACION_ABSOLUTA"]].rename(columns={"PUESTO": "Puesto", "VARIACION_ABSOLUTA": "Variación"})
    draw_pdf_horizontal_bars(draw, PDF_MARGIN, 805, 530, 300, chart_top, "Puesto", "Variación", "Mayor crecimiento", PDF_GREEN)
    draw_pdf_horizontal_bars(draw, PDF_MARGIN + 570, 805, 530, 300, chart_drop, "Puesto", "Variación", "Mayor caída", PDF_BLUE)

    draw_section_title(draw, PDF_MARGIN, 1160, "Resumen por templo", PDF_TEAL)
    display = resumen_templo[[
        "TEMPLO", "PUESTOS_ASIGNADOS", "VOTOS_2026_ASIGNADOS", "VARIACION_ABSOLUTA",
        "MESAS_TRABAJO", "BENEFICIARIOS_MESAS", "TOTAL_TESTIGOS", "LIDERES",
        "BENEFICIARIOS_MESAS_TRABAJO",
    ]].copy()
    display = format_table_for_display(display)
    headers = list(display.columns)
    widths = [170, 110, 130, 95, 110, 140, 120, 100, 125]
    y = draw_pdf_table_header(draw, PDF_MARGIN, 1225, widths, headers, PDF_TEAL)
    for row_idx, (_, row) in enumerate(display.iterrows()):
        fill = "#FFFFFF" if row_idx % 2 == 0 else "#F8FBFF"
        cursor = PDF_MARGIN
        for col, width in zip(headers, widths):
            draw.rectangle([cursor, y, cursor + width, y + 50], fill=fill, outline="#E6EDF5")
            draw.text((cursor + 10, y + 17), str(row.get(col, "")), fill=PDF_TEXT, font=pdf_font(11, True if col == "Templo" else False))
            cursor += width
        y += 50
    return page


def draw_compact_pdf_table(draw, x, y, widths, headers, rows, accent=PDF_TEAL, row_h=48, font_size=10):
    y = draw_pdf_table_header(draw, x, y, widths, headers, accent)
    for row_idx, row in enumerate(rows):
        fill = "#FFFFFF" if row_idx % 2 == 0 else "#F8FBFF"
        cursor = x
        for col, width in zip(headers, widths):
            draw.rectangle([cursor, y, cursor + width, y + row_h], fill=fill, outline="#E6EDF5")
            value = str(row.get(col, ""))
            font = pdf_font(font_size, col in {"#", "Templo"})
            max_chars = max(5, int(width / max(font_size * 0.52, 1)))
            if len(value) > max_chars:
                value = value[: max_chars - 1] + "…"
            draw.text((cursor + 9, y + 15), value, fill=PDF_TEXT, font=font)
            cursor += width
        y += row_h
    return y


def prepare_top10_general_table(df):
    if df is None or df.empty:
        return []
    actual_label = globals().get("metodologia_actual_label", "Votos 2026")
    base_label = globals().get("metodologia_base_label", "Votos 2023")
    top = df.copy().head(10).reset_index(drop=True)
    rows = []
    for idx, (_, row) in enumerate(top.iterrows(), start=1):
        rows.append({
            "#": idx,
            "Puesto": row.get("PUESTO", ""),
            "Templo": row.get("IGLESIA", row.get("TEMPLO_ASIGNADO_FINAL", "")),
            actual_label: fmt_number(row.get("VOTOS_2026", 0), 0),
            base_label: fmt_number(row.get("VOTOS_2023", 0), 0),
            "Var.": fmt_variacion(row.get("VARIACION_ABSOLUTA", 0), row.get("VARIACION_PORCENTUAL", 0)),
        })
    return rows


def draw_pdf_mini_title(draw, x, y, title, accent):
    draw.rounded_rectangle([x, y + 4, x + 7, y + 31], radius=4, fill=accent)
    draw.text((x + 18, y), title, fill=PDF_TEXT, font=pdf_font(18, True))


def build_general_pdf_detail_page(resumen_templo, top_crecimiento, top_caida):
    page, draw = new_pdf_page(
        "Tablas territoriales consolidadas",
        "Resumen por templo y ranking de puestos con mayor crecimiento y mayor caida.",
        eyebrow="Detalle territorial",
        size=(PDF_H, PDF_W),
    )
    page_w, _ = page.size

    draw_section_title(draw, PDF_MARGIN, 196, "Tabla general por templo", PDF_TEAL)
    draw.text((PDF_MARGIN, 235), "Asignacion vigente, votos, operacion territorial, testigos y beneficiarios.", fill=PDF_MUTED, font=pdf_font(15))

    resumen_pdf = resumen_templo[[
        "TEMPLO", "PUESTOS_ASIGNADOS", "VOTOS_2026_ASIGNADOS", "VOTOS_2023_ASIGNADOS",
        "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "ACTIVIDADES", "VOLANTEOS",
        "MESAS_TRABAJO", "BENEFICIARIOS_MESAS", "TOTAL_TESTIGOS", "LIDERES",
        "BENEFICIARIOS_MESAS_TRABAJO",
    ]].copy()
    resumen_pdf = format_table_for_display(resumen_pdf)
    headers = list(resumen_pdf.columns)
    widths = [160, 115, 145, 145, 115, 105, 120, 100, 120, 140, 125, 100, 144]
    y = draw_compact_pdf_table(
        draw,
        PDF_MARGIN,
        276,
        widths,
        headers,
        [row.to_dict() for _, row in resumen_pdf.iterrows()],
        accent=PDF_TEAL,
        row_h=54,
        font_size=11,
    )

    second_y = max(y + 50, 630)
    table_gap = 28
    mini_w = (page_w - PDF_MARGIN * 2 - table_gap) / 2
    left_x = PDF_MARGIN
    right_x = PDF_MARGIN + mini_w + table_gap
    headers = ["#", "Puesto", "Templo", globals().get("metodologia_actual_label", "Votos 2026"), globals().get("metodologia_base_label", "Votos 2023"), "Var."]
    widths = [38, 300, 120, 100, 100, 125]

    draw_pdf_mini_title(draw, left_x, second_y, "Top 10 puestos con mayor crecimiento", PDF_GREEN)
    draw_compact_pdf_table(
        draw,
        left_x,
        second_y + 42,
        widths,
        headers,
        prepare_top10_general_table(top_crecimiento),
        accent=PDF_GREEN,
        row_h=39,
        font_size=11,
    )

    draw_pdf_mini_title(draw, right_x, second_y, "Top 10 puestos con mayor caída", PDF_ORANGE)
    draw_compact_pdf_table(
        draw,
        right_x,
        second_y + 42,
        widths,
        headers,
        prepare_top10_general_table(top_caida),
        accent=PDF_ORANGE,
        row_h=39,
        font_size=11,
    )
    return page


def generar_resumen_general_kennedy(asignacion_df, actividades_df, mesas_df, testigos_df=None):
    resumen_asig = crear_resumen_asignacion(asignacion_df)
    resumen_oper = crear_resumen_operativo_por_templo(actividades_df, mesas_df)
    resumen = resumen_asig.merge(resumen_oper, on="TEMPLO", how="left")
    testigos_templo = preparar_testigos_por_templo(testigos_df, resumen["TEMPLO"].tolist() if "TEMPLO" in resumen.columns else TEMPLOS_OFICIALES)
    resumen = resumen.merge(testigos_templo, on="TEMPLO", how="left")
    for col in ["ACTIVIDADES", "VOLANTEOS", "MESAS_TRABAJO", "BENEFICIARIOS_MESAS", "COMPROMISOS_MESAS", "AJUSTES_TEMPORALES"]:
        if col in resumen.columns:
            resumen[col] = pd.to_numeric(resumen[col], errors="coerce").fillna(0).astype(int)
    for col in ["TOTAL_TESTIGOS", "LIDERES", "BENEFICIARIOS_MESAS_TRABAJO", "TESTIGOS_ELECTORALES", "TESTIGOS_LIDERES", "TESTIGOS_BENEFICIARIOS_MESAS"]:
        if col in resumen.columns:
            resumen[col] = pd.to_numeric(resumen[col], errors="coerce").fillna(0).astype(int)
    return resumen


def generar_metricas_generales_pdf(asignacion_df, actividades_df, mesas_df, resumen_templo, testigos_df=None):
    actual_label = globals().get("metodologia_actual_label", "Votos 2026")
    base_label = globals().get("metodologia_base_label", "Votos 2023")
    votos_2026 = sum_numeric(asignacion_df, "VOTOS_2026")
    votos_2023 = sum_numeric(asignacion_df, "VOTOS_2023")
    var_abs = votos_2026 - votos_2023
    var_pct = var_abs / votos_2023 if votos_2023 else np.nan
    benef_total, benef_internos, benef_externos = calcular_beneficiarios_mesas(mesas_df)
    testigos = sumar_testigos_metricas(testigos_df, resumen_templo["TEMPLO"].tolist() if "TEMPLO" in resumen_templo.columns else TEMPLOS_OFICIALES)
    return [
        (actual_label, fmt_number(votos_2026, 0)),
        (base_label, fmt_number(votos_2023, 0)),
        ("Variación", fmt_variacion(var_abs, var_pct)),
        ("Puestos", fmt_number(len(asignacion_df), 0)),
        ("Actividades", fmt_number(len(actividades_df), 0)),
        ("Mesas de trabajo", fmt_number(len(mesas_df), 0)),
        ("Beneficiarios mesas", fmt_number(benef_total, 0)),
        ("Beneficiarios internos", fmt_number(benef_internos, 0)),
        ("Beneficiarios externos", fmt_number(benef_externos, 0)),
        ("Testigos electorales", fmt_number(testigos["TOTAL_TESTIGOS"], 0)),
        ("Testigos líderes", fmt_number(testigos["LIDERES"], 0)),
        ("Testigos benef. mesas", fmt_number(testigos["BENEFICIARIOS_MESAS_TRABAJO"], 0)),
        ("Templos activos", fmt_number(resumen_templo["TEMPLO"].nunique(), 0)),
        ("Ajustes puestos", fmt_number(len(st.session_state.get("ajustes_asignacion", {})), 0)),
        ("Ajustes operativos", fmt_number(len(st.session_state.get("ajustes_mesas", {})) + len(st.session_state.get("ajustes_actividades", {})), 0)),
    ]


def generar_pdf_electoral_general(asignacion_df, puestos_df, iglesias_df, actividades_df, mesas_df, testigos_df=None):
    resumen_templo = generar_resumen_general_kennedy(asignacion_df, actividades_df, mesas_df, testigos_df)
    ranking = asignacion_df.copy()
    ranking["VARIACION_ABSOLUTA"] = pd.to_numeric(ranking.get("VARIACION_ABSOLUTA", 0), errors="coerce").fillna(0)
    if "IGLESIA" not in ranking.columns:
        ranking["IGLESIA"] = ranking.get("TEMPLO_ASIGNADO_FINAL", ranking.get("IGLESIA_ACTUAL", ""))
    top_crecimiento = ranking.sort_values("VARIACION_ABSOLUTA", ascending=False).head(10)
    top_caida = ranking.sort_values("VARIACION_ABSOLUTA", ascending=True).head(10)
    metricas = generar_metricas_generales_pdf(asignacion_df, actividades_df, mesas_df, resumen_templo, testigos_df)
    pages = [build_general_pdf_summary_page(metricas, resumen_templo, top_crecimiento, top_caida)]
    map_specs = [
        (
            "Mapa 1 - Rango electoral",
            "rango",
            "Lectura del rango electoral",
            "Los colores y tamanos de los puntos ordenan los puestos segun la metodologia electoral seleccionada. Rojo/naranja indica mayor caudal; azul/turquesa menor caudal relativo.",
            [("Alto", "#DC2626"), ("Medio alto", "#F97316"), ("Medio", PDF_BLUE), ("Bajo", PDF_TEAL), ("Templo", PDF_PURPLE)],
        ),
        (
            "Mapa 2 - Puestos de votacion",
            "puestos",
            "Lectura de puestos de votacion",
            "Cada punto representa un puesto de votacion y el color corresponde al templo vigente. Las lineas muestran relacion operativa puesto-templo.",
            [("Puestos", PDF_BLUE), ("Templos", PDF_PURPLE), ("Linea asignacion", "#64748B")],
        ),
        (
            "Mapa 3 - Actividades de campana",
            "actividades",
            "Lectura de actividades",
            "Los marcadores morados muestran actividades de campana con coordenadas disponibles, junto con templos y contorno territorial.",
            [("Actividades", PDF_PURPLE), ("Templos", PDF_GREEN), ("Contorno", "#0F172A")],
        ),
        (
            "Mapa 4 - Mesas de trabajo",
            "mesas",
            "Lectura de mesas de trabajo",
            "Los marcadores naranjas muestran mesas de trabajo activas con coordenadas, para ubicar gestion social frente a los templos.",
            [("Mesas", PDF_ORANGE), ("Templos", PDF_PURPLE), ("Contorno", "#0F172A")],
        ),
    ]
    for title, mode, legend_title, body, items in map_specs:
        pages.append(build_general_pdf_map_page(title, mode, asignacion_df, iglesias_df, actividades_df, mesas_df, legend_title, body, items))

    pages.append(build_general_pdf_detail_page(resumen_templo, top_crecimiento, top_caida))

    total_pages = len(pages)
    for idx, page in enumerate(pages, start=1):
        stamp_pdf_footer(page, idx, total_pages)
    pages = prepare_print_pdf_pages(pages)
    output = BytesIO()
    pages[0].save(
        output,
        format="PDF",
        save_all=True,
        append_images=pages[1:],
        resolution=PDF_EXPORT_DPI,
        quality=PDF_EXPORT_QUALITY,
        subsampling=0,
    )
    return output.getvalue()


def metric_accent(label):
    key = strip_accents(str(label)).lower()
    if "testigo" in key or "lider" in key:
        return PDF_PURPLE
    if "beneficiario" in key:
        return PDF_GREEN
    if "voto" in key or "camara" in key or "senado" in key or "jal" in key or "concejo" in key:
        return PDF_BLUE
    if "variacion" in key:
        return PDF_ORANGE
    if "mesa" in key:
        return PDF_TEAL
    return PDF_PURPLE


def build_pdf_summary_page(iglesia, metricas, puestos_df, mesas_df, templo_row):
    page, draw = new_pdf_page(
        f"Informe territorial: {iglesia}",
        "Mapa operativo, indicadores electorales y detalle de mesas y puestos asignados.",
    )
    draw.rounded_rectangle([PDF_MARGIN, 210, PDF_W - PDF_MARGIN, 360], radius=18, fill="#FFFFFF", outline=PDF_BORDER)
    draw.text((PDF_MARGIN + 28, 238), "Resumen ejecutivo", fill=PDF_TEXT, font=pdf_font(25, True))
    draw.text(
        (PDF_MARGIN + 28, 282),
        "La ficha consolida la asignación vigente del templo, incluyendo ajustes guardados, mesas de trabajo, beneficiarios y votación histórica.",
        fill="#334155",
        font=pdf_font(16),
    )

    metric_rows = max(1, math.ceil(len(metricas) / 4))
    card_w = 255
    card_h = 90 if metric_rows > 3 else 104
    row_step = 108 if metric_rows > 3 else 128
    start_y = 400 if metric_rows > 3 else 410
    for idx, (label, value) in enumerate(metricas):
        row, col = divmod(idx, 4)
        draw_pdf_card(
            draw,
            (PDF_MARGIN + col * (card_w + 28), start_y + row * row_step),
            label,
            value,
            width=card_w,
            height=card_h,
            accent=metric_accent(label),
        )
    map_title_y = start_y + metric_rows * row_step + 28
    draw_section_title(draw, PDF_MARGIN, map_title_y, "Mapa territorial del templo", PDF_TEAL)
    map_x, map_y = PDF_MARGIN, map_title_y + 52
    map_w = PDF_W - PDF_MARGIN * 2
    map_h = min(650, PDF_H - map_y - 190)
    rendered = draw_real_temple_map(page, draw, map_x, map_y, map_w, map_h, iglesia, templo_row, puestos_df, mesas_df)
    if not rendered:
        draw_static_temple_map(draw, map_x, map_y, map_w, map_h, iglesia, templo_row, puestos_df, mesas_df)
    return page


def build_pdf_map_page(iglesia, puestos_df, mesas_df, templo_row):
    page, draw = new_pdf_page(
        f"Mapa territorial: {iglesia}",
        "Azul: puestos de votación. Naranja: mesas de trabajo. Morado: templo asignado.",
        eyebrow="Territorio operativo",
        size=(PDF_H, PDF_W),
    )
    page_w, page_h = page.size
    draw_section_title(draw, PDF_MARGIN, 210, "Mapa territorial completo", PDF_TEAL)
    map_x, map_y = PDF_MARGIN, 280
    map_w, map_h = page_w - PDF_MARGIN * 2, 720
    rendered = draw_real_temple_map(page, draw, map_x, map_y, map_w, map_h, iglesia, templo_row, puestos_df, mesas_df)
    if not rendered:
        draw_static_temple_map(draw, map_x, map_y, map_w, map_h, iglesia, templo_row, puestos_df, mesas_df)
    return page


def prepare_mesas_pdf(mesas_df):
    if mesas_df is None or mesas_df.empty:
        return pd.DataFrame(columns=["Nombre de la mesa", "Beneficiarios", "Líder"])
    df = mesas_df.copy()
    df["Nombre de la mesa"] = df.get("NOMBRE_GESTION", df.get("TEMA", "")).fillna("").astype(str)
    df["Beneficiarios"] = pd.to_numeric(df.get("BENEFICIARIOS", 0), errors="coerce").fillna(0).map(lambda v: fmt_number(v, 0))
    df["Líder"] = df.get("LIDER", "").fillna("").astype(str)
    return df[["Nombre de la mesa", "Beneficiarios", "Líder"]].sort_values("Nombre de la mesa")


def prepare_puestos_pdf(puestos_df):
    actual_label = globals().get("metodologia_actual_label", "Votos 2026")
    base_label = globals().get("metodologia_base_label", "Votos 2023")
    if puestos_df is None or puestos_df.empty:
        return pd.DataFrame(columns=["Puesto", actual_label, base_label, "Variación"])
    df = puestos_df.copy()
    df["Puesto"] = df.get("PUESTO", "").fillna("").astype(str)
    df[actual_label] = pd.to_numeric(df.get("VOTOS_2026", 0), errors="coerce").fillna(0)
    df[base_label] = pd.to_numeric(df.get("VOTOS_2023", 0), errors="coerce").fillna(0)
    df["Variación"] = pd.to_numeric(df.get("VARIACION_ABSOLUTA", df[actual_label] - df[base_label]), errors="coerce").fillna(0)
    df = df.sort_values(actual_label, ascending=False)
    for col in [actual_label, base_label, "Variación"]:
        df[col] = df[col].map(lambda v: fmt_number(v, 0))
    return df[["Puesto", actual_label, base_label, "Variación"]]


def prepare_print_pdf_pages(pages):
    print_pages = []
    for page in pages:
        page = page.convert("RGB")
        high_res = page.resize(
            (page.width * PDF_PRINT_SCALE, page.height * PDF_PRINT_SCALE),
            Image.Resampling.LANCZOS,
        )
        high_res = ImageEnhance.Contrast(high_res).enhance(1.04)
        high_res = ImageEnhance.Sharpness(high_res).enhance(1.18)
        high_res = high_res.filter(ImageFilter.UnsharpMask(radius=0.55, percent=165, threshold=2))
        print_pages.append(high_res)
    return print_pages


def generar_pdf_templo(iglesia, metricas, puestos_df, mesas_df, templo_row):
    pages = [build_pdf_summary_page(iglesia, metricas, puestos_df, mesas_df, templo_row)]

    mesas_pdf = prepare_mesas_pdf(mesas_df)
    draw_paginated_pdf_table(
        pages,
        "Mesas de trabajo asignadas",
        mesas_pdf,
        ["Nombre de la mesa", "Beneficiarios", "Líder"],
        [700, 155, 245],
        accent=PDF_ORANGE,
        subtitle=f"{len(mesas_pdf)} registros asignados a {iglesia}",
    )

    puestos_pdf = prepare_puestos_pdf(puestos_df)
    draw_paginated_pdf_table(
        pages,
        "Puestos de votación asignados",
        puestos_pdf,
        list(puestos_pdf.columns),
        [570, 175, 175, 180],
        accent=PDF_BLUE,
        subtitle=f"{len(puestos_pdf)} puestos asignados a {iglesia}",
    )

    total_pages = len(pages)
    for idx, page in enumerate(pages, start=1):
        stamp_pdf_footer(page, idx, total_pages)
    pages = prepare_print_pdf_pages(pages)
    output = BytesIO()
    pages[0].save(
        output,
        format="PDF",
        save_all=True,
        append_images=pages[1:],
        resolution=PDF_EXPORT_DPI,
        quality=PDF_EXPORT_QUALITY,
        subsampling=0,
    )
    return output.getvalue()


def status_strip(items):
    html_items = "".join(
        f'<div class="status-item"><span>{safe_html(label)}</span><strong>{safe_html(value)}</strong></div>'
        for label, value in items
    )
    st.markdown(f'<div class="status-strip">{html_items}</div>', unsafe_allow_html=True)


def aplicar_filtros(puestos, actividades, mesas, filtros):
    iglesias_sel = filtros.get("iglesias", [])
    puestos_f = puestos.copy()
    acts_f = actividades.copy()
    mesas_f = mesas.copy()

    if iglesias_sel:
        puestos_f = puestos_f[puestos_f["IGLESIA"].isin(iglesias_sel)]
        if "IGLESIA" in acts_f.columns:
            acts_f = acts_f[acts_f["IGLESIA"].isin(iglesias_sel)]
        if "IGLESIA" in mesas_f.columns:
            mesas_f = mesas_f[mesas_f["IGLESIA"].isin(iglesias_sel)]

    if "ESTRATEGIA" in acts_f.columns:
        valid_estrategias = ["LIBERTAD RELIGIOSA", "POLITICO COMUNITARIA", "POLÍTICO COMUNITARIA"]
        acts_f = acts_f[acts_f["ESTRATEGIA"].str.strip().str.upper().isin(valid_estrategias)]

    return puestos_f, acts_f, mesas_f


def crear_mapa(puestos, iglesias, actividades, mesas, map_mode="Vista general", layers_config=None):
    layers_config = layers_config or {}
    m = crear_mapa_base(location=KENNEDY_CENTER, zoom_start=13, control_scale=True)
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
        votos = float(r.get("VOTOS_2026", 0) or 0)
        radius = max(4, min(8, 4 + math.sqrt(votos) / 4))

        popup = f"""
        <div style="font-family:'Inter', sans-serif; width:340px; color:#0F172A;">
        <h4 style="margin-bottom:12px; font-weight:800; border-bottom: 1px solid #E2E8F0; padding-bottom:8px;">{puesto}</h4>
        <table style="width:100%; border-collapse: collapse; font-size: 12.5px;">
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600; color:#475569;">Iglesia</td><td style="padding:6px 8px; font-weight:700;">{iglesia}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600; color:#475569;">Barrio (UPZ)</td><td style="padding:6px 8px;">{barrio} ({upz})</td></tr>
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600; color:#475569;">{safe_html(globals().get("metodologia_actual_label", "Votos 2026"))}</td><td style="padding:6px 8px; font-weight:800; color:#2563EB;">{fmt_number(r.get('VOTOS_2026'),0)}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600; color:#475569;">Variación</td><td style="padding:6px 8px;">{fmt_variacion(r.get('VARIACION_ABSOLUTA'), r.get('VARIACION_PORCENTUAL'))}</td></tr>
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600; color:#475569;">Actividades</td><td style="padding:6px 8px;">{fmt_number(r.get('ACTIVIDADES_CAMPANA'),0)}</td></tr>
            <tr><td style="padding:6px 8px; font-weight:600; color:#475569;">Mesas de trabajo</td><td style="padding:6px 8px;">{fmt_number(r.get('MESAS_TRABAJO_BARRIO'),0)}</td></tr>
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
        nombre_mesa = r.get("NOMBRE_GESTION") or r.get("TEMA") or "Mesa sin nombre"
        popup_mesa = f"""
        <div style="font-family:'Inter', sans-serif; width:360px; color:#0F172A;">
        <h4 style="margin-bottom:12px; font-weight:800; border-bottom: 1px solid #E2E8F0; padding-bottom:8px;">Mesa: {safe_html(nombre_mesa)}</h4>
        <table style="width:100%; border-collapse: collapse; font-size: 12.5px;">
            <tr style="background:#F8FAFC;"><td style="padding:6px 8px; font-weight:600;">Tema</td><td style="padding:6px 8px;">{safe_html(r.get('TEMA',''))}</td></tr>
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
            tooltip=f"{nombre_mesa} | {r.get('IGLESIA','')} | {r.get('BARRIO','')}",
            popup=folium.Popup(popup_mesa, max_width=430),
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
    <div style="background:white;padding:10px 12px;border:1px solid #CBD5E1;border-radius:10px;box-shadow:0 4px 14px rgba(15,23,42,.16);font-size:11px;color:#0F172A;min-width:190px;max-width:190px;font-family:'Inter', sans-serif;">
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

    add_map_control_html(m, legend_html, position="bottomright")
    return m


@st.cache_data(show_spinner=False, max_entries=20)
def cached_crear_mapa_html(puestos, iglesias, actividades, mesas, map_mode="Vista general", layers_config_tuple=None):
    layers_config = dict(layers_config_tuple) if layers_config_tuple else None
    m = crear_mapa(puestos, iglesias, actividades, mesas, map_mode, layers_config)
    html = m.get_root().render()
    del m
    import gc
    gc.collect()
    return html

@st.cache_data(show_spinner=False, max_entries=20)
def cached_crear_mapa_asignacion_html(asignacion_df, iglesias_df, layers_config_tuple=None):
    layers_config = dict(layers_config_tuple) if layers_config_tuple else None
    m = crear_mapa_asignacion(asignacion_df, iglesias_df, layers_config)
    html = m.get_root().render()
    del m
    import gc
    gc.collect()
    return html

@st.cache_data(show_spinner=False, max_entries=20)
def cached_submap_html(iglesia, lat_t, lon_t, sub_puestos, sub_mesas):
    sub_m = crear_mapa_base(location=[lat_t, lon_t], zoom_start=14, control_scale=False)
    import folium
    folium.Marker(
        [lat_t, lon_t],
        icon=crear_icono_div("templo", COLORES_TEMPLOS.get(iglesia, "#1E3A8A"), "T"),
        tooltip=iglesia,
    ).add_to(sub_m)
    for _, r in sub_puestos.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        folium.CircleMarker(
            [r["LATITUD"], r["LONGITUD"]],
            radius=max(5, min(14, float(r.get("VOTOS_2026", 0)) / 15)),
            color="#1E3A8A",
            fill=True,
            fill_color="#3B82F6",
            fill_opacity=0.7,
            weight=1,
            tooltip=f"{r.get('PUESTO')} | {fmt_number(r.get('VOTOS_2026'), 0)} votos",
        ).add_to(sub_m)
        folium.PolyLine([[lat_t, lon_t], [r["LATITUD"], r["LONGITUD"]]], color="#1E3A8A", weight=1.4, opacity=0.28, dash_array="4,6").add_to(sub_m)
    for _, r in sub_mesas.dropna(subset=["LATITUD", "LONGITUD"]).iterrows():
        nombre_mesa = r.get("NOMBRE_GESTION") or r.get("TEMA") or "Mesa sin nombre"
        folium.Marker(
            [r["LATITUD"], r["LONGITUD"]],
            icon=crear_icono_div("mesa", "#F97316", "M"),
            tooltip=f"{nombre_mesa} | {r.get('BARRIO', '')}",
        ).add_to(sub_m)
    html = sub_m.get_root().render()
    del sub_m
    import gc
    gc.collect()
    return html


def download_excel_link():
    with open(CONSOLIDADO, "rb") as f:
        return f.read()


# ============================================================
# CARGA
# ============================================================

data = cargar_datos(CONSOLIDADO, CONSOLIDADO.stat().st_mtime if CONSOLIDADO.exists() else 0)

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
matriz = data["matriz"]
informe = data["informe"]
control = data.get("control", pd.DataFrame())
asignacion = data.get("asignacion", pd.DataFrame())
resumen_asignacion = data.get("resumen_asignacion", pd.DataFrame())
testigos_resumen = cargar_testigos_resumen(
    TESTIGOS_RESUMEN_CSV,
    TESTIGOS_RESUMEN_CSV.stat().st_mtime if TESTIGOS_RESUMEN_CSV.exists() else 0,
)
APOYOS_CIUDADANOS_FILE = first_existing_path(APOYOS_CIUDADANOS_CANDIDATES)
apoyos_ciudadanos = cargar_apoyos_ciudadanos(
    APOYOS_CIUDADANOS_FILE,
    APOYOS_CIUDADANOS_FILE.stat().st_mtime if APOYOS_CIUDADANOS_FILE and APOYOS_CIUDADANOS_FILE.exists() else 0,
)

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

# ============================================================
# SIDEBAR PRINCIPAL
# ============================================================

with st.sidebar:
    st.header("Configuración del análisis")
    st.markdown("Fuente única: `kennedy_mira_consolidado.xlsx`")
    backend_label = persistence_backend_label()
    st.caption(f"Base de cambios: {backend_label}")
    if st.session_state.get("google_sheets_ready"):
        st.success("Google Sheets conectado. Los cambios se guardan online.")
    elif st.session_state.get("google_sheets_error"):
        st.warning("Google Sheets está configurado, pero no conectado. Revise credenciales o permisos del archivo.")
        with st.expander("Diagnóstico Google Sheets", expanded=False):
            for key, value in google_sheets_diagnostics().items():
                st.caption(f"**{key}:** {value}")

    metodologia_labels = [cfg["label"] for cfg in METODOLOGIAS_ELECTORALES.values()]
    metodologia_keys = list(METODOLOGIAS_ELECTORALES.keys())
    selected_label = st.selectbox(
        "Metodología electoral",
        metodologia_labels,
        index=metodologia_keys.index(METODOLOGIA_DEFAULT),
        help="Esta selección recalcula votos, variación, mapas, tablas y PDFs en toda la página.",
    )
    metodologia_electoral = metodologia_keys[metodologia_labels.index(selected_label)]
    metodologia_label, metodologia_caption, metodologia_base_label, metodologia_actual_label = resumen_metodologia_electoral(metodologia_electoral)
    TABLE_COLUMN_LABELS["VOTOS_2023"] = f"{metodologia_base_label} lectura"
    TABLE_COLUMN_LABELS["VOTOS_2026"] = f"{metodologia_actual_label} lectura"
    TABLE_COLUMN_LABELS["VOTOS_2023_ASIGNADOS"] = f"{metodologia_base_label} asignado"
    TABLE_COLUMN_LABELS["VOTOS_2026_ASIGNADOS"] = f"{metodologia_actual_label} asignado"
    st.caption(metodologia_caption)

    iglesias_oficiales = IGLESIAS_OFICIALES_PERMITIDAS
    default_iglesias = iglesias_oficiales
    selected_iglesias = st.multiselect("Iglesias / templos", iglesias_oficiales, default=default_iglesias)

actividades = aplicar_ajustes_templo(actividades, "ajustes_actividades", "ACTIVIDAD_ID")
mesas = aplicar_ajustes_templo(mesas, "ajustes_mesas", "MESA_ID")
puestos = aplicar_ajustes_templo(puestos, "ajustes_asignacion", "PUESTO")
if not matriz.empty:
    matriz = aplicar_ajustes_templo(matriz, "ajustes_asignacion", "PUESTO")

puestos = aplicar_metodologia_electoral(puestos, metodologia_electoral)
resumen_puesto = aplicar_metodologia_electoral(resumen_puesto, metodologia_electoral)
if not matriz.empty:
    matriz = aplicar_metodologia_electoral(matriz, metodologia_electoral)

if not resumen_iglesia.empty:
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

    extremos_rows = []
    for iglesia_nombre, sub in puestos.groupby("IGLESIA"):
        sub = sub.copy()
        sub["VOTOS_2026"] = pd.to_numeric(sub.get("VOTOS_2026", 0), errors="coerce").fillna(0)
        sub["VARIACION_ABSOLUTA"] = pd.to_numeric(sub.get("VARIACION_ABSOLUTA", 0), errors="coerce").fillna(0)
        mayor_votacion = sub.sort_values("VOTOS_2026", ascending=False).iloc[0]["PUESTO"] if not sub.empty else "SIN PUESTOS ASIGNADOS"
        mayor_caida = sub.sort_values("VARIACION_ABSOLUTA", ascending=True).iloc[0]["PUESTO"] if not sub.empty else "SIN PUESTOS ASIGNADOS"
        mayor_crecimiento = sub.sort_values("VARIACION_ABSOLUTA", ascending=False).iloc[0]["PUESTO"] if not sub.empty else "SIN PUESTOS ASIGNADOS"
        extremos_rows.append({
            "IGLESIA": iglesia_nombre,
            "PUESTO_MAYOR_VOTACION": mayor_votacion,
            "PUESTO_MAYOR_CAIDA": mayor_caida,
            "PUESTO_MAYOR_CRECIMIENTO": mayor_crecimiento,
        })
    agregado_extremos = pd.DataFrame(extremos_rows)
        
    cols_to_update = [
        "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL",
        "PUESTOS", "ACTIVIDADES_CAMPANA", "MESAS_TRABAJO", "PUESTO_MAYOR_VOTACION",
        "PUESTO_MAYOR_CAIDA", "PUESTO_MAYOR_CRECIMIENTO",
    ]
    cols_to_drop = [c for c in cols_to_update if c in resumen_iglesia.columns]
    resumen_iglesia = resumen_iglesia.drop(columns=cols_to_drop)
    
    resumen_iglesia = resumen_iglesia.merge(agregado_puestos, on="IGLESIA", how="left")
    resumen_iglesia = resumen_iglesia.merge(agregado_acts, on="IGLESIA", how="left")
    resumen_iglesia = resumen_iglesia.merge(agregado_mesas, on="IGLESIA", how="left")
    resumen_iglesia = resumen_iglesia.merge(agregado_extremos, on="IGLESIA", how="left")
    
    for col in ["VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "PUESTOS", "ACTIVIDADES_CAMPANA", "MESAS_TRABAJO"]:
        if col in resumen_iglesia.columns:
            resumen_iglesia[col] = resumen_iglesia[col].fillna(0)

testigos_por_templo = preparar_testigos_por_templo(testigos_resumen, IGLESIAS_OFICIALES_PERMITIDAS)
if not resumen_iglesia.empty and "IGLESIA" in resumen_iglesia.columns:
    testigos_merge = testigos_por_templo[
        ["TEMPLO", "TOTAL_TESTIGOS", "LIDERES", "BENEFICIARIOS_MESAS_TRABAJO"]
    ].rename(columns={"TEMPLO": "IGLESIA"})
    resumen_iglesia = resumen_iglesia.drop(columns=[c for c in ["TOTAL_TESTIGOS", "LIDERES", "BENEFICIARIOS_MESAS_TRABAJO"] if c in resumen_iglesia.columns])
    resumen_iglesia = resumen_iglesia.merge(testigos_merge, on="IGLESIA", how="left")
    for col in ["TOTAL_TESTIGOS", "LIDERES", "BENEFICIARIOS_MESAS_TRABAJO"]:
        resumen_iglesia[col] = pd.to_numeric(resumen_iglesia[col], errors="coerce").fillna(0).astype(int)

asignacion = calcular_distancias_a_templos_v2(puestos, iglesias)


# Ensure numerics
for df in [puestos, resumen_iglesia, resumen_puesto, matriz, asignacion, resumen_asignacion]:
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
            "BENEFICIARIOS_REFERIDOS", "BENEFICIARIOS_NO_INFOMIRA", "BENEFICIARIOS_INTERNOS", "BENEFICIARIOS_EXTERNOS",
            "TESTIGOS_2023_REPORTE", "VOTOS_AFINIDAD_E11_2023", "VOTOS_MIRA_2023_PROP_LISTA",
            "ACTIVIDADES_CAMPANA", "ACTIVIDADES_CAMPANA_IGLESIA", "MESAS_TRABAJO_BARRIO",
            "MESAS_TRABAJO", "PUESTOS", "PUNTAJE_PRIORIDAD", "TOTAL_TESTIGOS", "LIDERES",
            "BENEFICIARIOS_MESAS_TRABAJO",
        ]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").round().astype("Int64")


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
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
    {"iglesias": selected_iglesias},
)

resumen_iglesia_f = resumen_iglesia[resumen_iglesia["IGLESIA"].isin(selected_iglesias)].copy()


# ============================================================
# ENCABEZADO
# ============================================================

st.title("Dashboard territorial-electoral Kennedy")
st.markdown(
    '<div class="subtitle">Campaña Congreso 2026 · Partido MIRA · Votación, gestión, iglesias, puestos, barrios y UPZ</div>',
    unsafe_allow_html=True,
)

# Métricas oficiales según metodología seleccionada
total_2026 = sum_numeric(puestos, "VOTOS_2026")
total_2023 = sum_numeric(puestos, "VOTOS_2023")
var_abs = total_2026 - total_2023
var_pct = var_abs / total_2023 if total_2023 else 0
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
puestos_cambio_templo = get_indicador(resumen_general, "Puestos con cambio de templo sugerido", default=np.nan)
if pd.isna(puestos_cambio_templo) and "CAMBIO_PROPUESTO_TEMPLO" in puestos.columns:
    puestos_cambio_templo = int(puestos["CAMBIO_PROPUESTO_TEMPLO"].astype(str).eq("SI").sum())
beneficiarios_mesas_total = pd.to_numeric(mesas.get("BENEFICIARIOS", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
if "BENEFICIARIOS_INTERNOS" in mesas.columns:
    beneficiarios_mesas_internos = pd.to_numeric(mesas["BENEFICIARIOS_INTERNOS"], errors="coerce").fillna(0).sum()
elif "BENEFICIARIOS_NO_INFOMIRA" in mesas.columns:
    beneficiarios_mesas_internos = max(beneficiarios_mesas_total - pd.to_numeric(mesas["BENEFICIARIOS_NO_INFOMIRA"], errors="coerce").fillna(0).sum(), 0)
else:
    beneficiarios_mesas_internos = pd.to_numeric(mesas.get("BENEFICIARIOS_REFERIDOS", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
if "BENEFICIARIOS_EXTERNOS" in mesas.columns:
    beneficiarios_mesas_externos = pd.to_numeric(mesas["BENEFICIARIOS_EXTERNOS"], errors="coerce").fillna(0).sum()
elif "BENEFICIARIOS_NO_INFOMIRA" in mesas.columns:
    beneficiarios_mesas_externos = pd.to_numeric(mesas["BENEFICIARIOS_NO_INFOMIRA"], errors="coerce").fillna(0).sum()
else:
    beneficiarios_mesas_externos = max(beneficiarios_mesas_total - beneficiarios_mesas_internos, 0)
testigos_global = sumar_testigos_metricas(testigos_resumen, IGLESIAS_OFICIALES_PERMITIDAS)

c1, c2, c3, c4 = st.columns(4)

with c1:
    metric_card(metodologia_actual_label, fmt_number(total_2026, 0))

with c2:
    metric_card(f"Variación vs {metodologia_base_label}", fmt_variacion(var_abs, var_pct), fmt_variacion(abs(var_abs), abs(var_pct)), positive=var_abs >= 0)

with c3:
    metric_card("Puestos analizados", fmt_number(puestos_total, 0))

with c4:
    metric_card("Iglesias oficiales", fmt_number(iglesias_total, 0))

with st.expander("Ver indicadores complementarios", expanded=False):
    e1, e2, e3, e4 = st.columns(4)
    with e1:
        metric_card(metodologia_base_label, fmt_number(total_2023, 0))
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

    b1, b2, b3 = st.columns(3)
    with b1:
        metric_card("Beneficiarios mesas", fmt_number(beneficiarios_mesas_total, 0))
    with b2:
        metric_card("Beneficiarios internos", fmt_number(beneficiarios_mesas_internos, 0))
    with b3:
        metric_card("Beneficiarios externos", fmt_number(beneficiarios_mesas_externos, 0))

    t1, t2, t3 = st.columns(3)
    with t1:
        metric_card("Testigos electorales", fmt_number(testigos_global["TOTAL_TESTIGOS"], 0))
    with t2:
        metric_card("Testigos líderes", fmt_number(testigos_global["LIDERES"], 0))
    with t3:
        metric_card("Testigos benef. mesas", fmt_number(testigos_global["BENEFICIARIOS_MESAS_TRABAJO"], 0))

st.markdown(
    f"""
    <div class="summary-ribbon">
    <b>Lectura ejecutiva:</b> bajo la metodología <b>{metodologia_label}</b>, Kennedy registra
    <b>{fmt_number(total_2026, 0)}</b> en {metodologia_actual_label},
    con una variación de <b>{fmt_number(var_abs, 0)}</b> frente a {metodologia_base_label}.
    El tablero permite revisar puestos, analizar presencia territorial y ajustar la asignación operativa
    por templo para la estrategia electoral.
    </div>
    """,
    unsafe_allow_html=True,
)

ajustes_total_global = (
    len(st.session_state.get("ajustes_asignacion", {}))
    + len(st.session_state.get("ajustes_mesas", {}))
    + len(st.session_state.get("ajustes_actividades", {}))
)
status_strip(
    [
        ("Base de cambios", persistence_backend_label()),
        ("Ajustes guardados", fmt_number(ajustes_total_global, 0)),
        ("Templos activos", f"{fmt_number(len(selected_iglesias), 0)} de {fmt_number(len(IGLESIAS_OFICIALES_PERMITIDAS), 0)}"),
    ]
)


# ============================================================
# TABS
# ============================================================

tab_resumen, tab_mapa, tab_asignacion, tab_mesas, tab_apoyos, tab_iglesia, tab_puesto, tab_export = st.tabs(
    [
        "Resumen ejecutivo",
        "Mapa territorial",
        "Asignación de puestos",
        "Mesas de trabajo",
        "Apoyos ciudadanos",
        "Análisis por iglesia",
        "Análisis por puesto",
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
    ranking_variacion = puestos.copy()
    ranking_variacion["VARIACION_ABSOLUTA"] = pd.to_numeric(ranking_variacion["VARIACION_ABSOLUTA"], errors="coerce")
    ranking_variacion = ranking_variacion.dropna(subset=["VARIACION_ABSOLUTA", "PUESTO"])
    with col1:
        top_pos = ranking_variacion.sort_values("VARIACION_ABSOLUTA", ascending=False).head(10)
        top_pos_plot = top_pos.sort_values("VARIACION_ABSOLUTA", ascending=True)
        fig = px.bar(
            top_pos_plot,
            x="VARIACION_ABSOLUTA",
            y="PUESTO",
            orientation="h",
            color="IGLESIA",
            color_discrete_map=COLORES_TEMPLOS,
            category_orders=PLOTLY_TEMPLO_ORDERS,
            title="Top 10 puestos con mayor crecimiento",
            text="VARIACION_ABSOLUTA",
        )
        fig.update_traces(texttemplate="%{text:.0f}", textposition="outside", cliponaxis=False)
        fig.update_yaxes(categoryorder="array", categoryarray=top_pos_plot["PUESTO"].tolist())
        fig.update_layout(height=420, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, xaxis_title="VARIACION_ABSOLUTA", yaxis_title="PUESTO")
        st.plotly_chart(fig, width="stretch")
    with col2:
        top_neg = ranking_variacion.sort_values("VARIACION_ABSOLUTA", ascending=True).head(10)
        top_neg_plot = top_neg.sort_values("VARIACION_ABSOLUTA", ascending=False)
        fig = px.bar(
            top_neg_plot,
            x="VARIACION_ABSOLUTA",
            y="PUESTO",
            orientation="h",
            color="IGLESIA",
            color_discrete_map=COLORES_TEMPLOS,
            category_orders=PLOTLY_TEMPLO_ORDERS,
            title="Top 10 puestos con mayor caída",
            text="VARIACION_ABSOLUTA",
        )
        fig.update_traces(texttemplate="%{text:.0f}", textposition="outside", cliponaxis=False)
        fig.update_yaxes(categoryorder="array", categoryarray=top_neg_plot["PUESTO"].tolist())
        fig.update_layout(height=420, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, xaxis_title="VARIACION_ABSOLUTA", yaxis_title="PUESTO")
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

    st.markdown('<div class="layer-hint">Capas del mapa: active o apague rangos, puntos y operación sin usar el control desplegable interno de Leaflet.</div>', unsafe_allow_html=True)
    with st.expander("Capas del mapa territorial", expanded=True):
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
    html_map = cached_crear_mapa_html(puestos_mapa, iglesias_mapa, acts_mapa, mesas_mapa, map_mode, tuple(territorial_layers.items()))
    st.markdown("### Mapa territorial")
    st.markdown("<div style='font-size:14px; color:#475569; margin-bottom:10px;'>💡 <b>Vista inicial limpia:</b> el modo de vista controla si se muestran votos, calor, actividades o mesas. Las convenciones del mapa quedan integradas abajo.</div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.components.v1.html(html_map, height=790)

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
                mesa_barrio = st.text_input("Barrio", value=str(mesa_row.get("BARRIO", "") or ""), key="barrio_mesa_ajuste_compacto")
                mesa_lider = st.text_input("Líder", value=str(mesa_row.get("LIDER", "") or ""), key="lider_mesa_ajuste_compacto")
                if st.button("Guardar ajuste de mesa"):
                    st.session_state.setdefault("ajustes_mesas", {})
                    st.session_state.setdefault("ajustes_mesas_barrio", {})
                    st.session_state.setdefault("ajustes_mesas_lider", {})
                    mesa_barrio = mesa_barrio.strip()
                    mesa_lider = mesa_lider.strip()
                    st.session_state["ajustes_mesas"][mesa_row["MESA_ID"]] = mesa_templo
                    st.session_state["ajustes_mesas_barrio"][mesa_row["MESA_ID"]] = mesa_barrio
                    st.session_state["ajustes_mesas_lider"][mesa_row["MESA_ID"]] = mesa_lider
                    registrar_ajuste_en_db(
                        session_key="ajustes_mesas",
                        entity_id=mesa_row["MESA_ID"],
                        nombre_entidad=mesa_row.get("NOMBRE_GESTION", mesa_row.get("TEMA", "")),
                        templo_nuevo=mesa_templo,
                        barrio_nuevo=mesa_barrio,
                        lider_nuevo=mesa_lider,
                        motivo="Ajuste manual desde pestaña Mapa territorial",
                    )
                    guardar_ajustes_guardados()
                    st.success("Ajuste de mesa guardado.")
                    st.rerun()
                if st.button("Limpiar ajustes de mesas"):
                    total_limpiados = limpiar_ajustes_en_db("ajustes_mesas", motivo="Limpieza manual desde pestaña Mapa territorial")
                    st.session_state["ajustes_mesas"] = {}
                    st.session_state["ajustes_mesas_barrio"] = {}
                    st.session_state["ajustes_mesas_lider"] = {}
                    guardar_ajustes_guardados()
                    st.info(f"Se limpiaron {fmt_number(total_limpiados, 0)} ajuste(s) de mesas en la base.")
                    st.rerun()

    st.markdown("### Resumen operativo por templo")
    st.dataframe(resumen_operativo_mapa, hide_index=True, width="stretch")
    st.download_button("Descargar tabla en Excel", to_excel_bytes(resumen_operativo_mapa, "Resumen Operativo"), "resumen_operativo.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_res_op")

    asignacion_reporte_mapa = aplicar_ajustes_asignacion(asignacion.copy())
    informe_mapa = generar_informe_territorial(asignacion_reporte_mapa, actividades, mesas, testigos_resumen)
    with st.expander("Informe territorial automático", expanded=True):
        resumen_general_pdf = generar_resumen_general_kennedy(asignacion_reporte_mapa, actividades, mesas, testigos_resumen)
        metricas_general_pdf = generar_metricas_generales_pdf(asignacion_reporte_mapa, actividades, mesas, resumen_general_pdf, testigos_resumen)
        st.markdown("### Informe electoral general Kennedy")
        st.markdown(
            """
            <div class="section-card">
            Este informe consolida la lectura general de Kennedy con resumen ejecutivo, mapas territoriales por capa, gráficas comparativas y tablas listas para seguimiento.
            </div>
            """,
            unsafe_allow_html=True,
        )

        for start in range(0, min(len(metricas_general_pdf), 12), 4):
            metric_cols = st.columns(4)
            for col, (label, value) in zip(metric_cols, metricas_general_pdf[start:start + 4]):
                with col:
                    metric_card(label, value)

        chart_t1, chart_t2 = st.columns(2)
        resumen_chart = resumen_general_pdf.copy()
        with chart_t1:
            fig_votos_templo = px.bar(
                resumen_chart.sort_values("VOTOS_2026_ASIGNADOS", ascending=True),
                x="VOTOS_2026_ASIGNADOS",
                y="TEMPLO",
                orientation="h",
                color="TEMPLO",
                color_discrete_map=COLORES_TEMPLOS,
                category_orders=PLOTLY_TEMPLO_ORDERS,
                title=f"{metodologia_actual_label} por templo vigente",
                text="VOTOS_2026_ASIGNADOS",
            )
            fig_votos_templo.update_traces(texttemplate="%{text:.0f}", textposition="outside", cliponaxis=False)
            fig_votos_templo.update_layout(height=380, showlegend=False, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, xaxis_title=metodologia_actual_label, yaxis_title="")
            st.plotly_chart(fig_votos_templo, width="stretch")
        with chart_t2:
            operativo_long = resumen_chart[["TEMPLO", "ACTIVIDADES", "MESAS_TRABAJO", "BENEFICIARIOS_MESAS", "TOTAL_TESTIGOS"]].melt("TEMPLO", var_name="Indicador", value_name="Cantidad")
            operativo_long["Indicador"] = operativo_long["Indicador"].map({
                "ACTIVIDADES": "Actividades",
                "MESAS_TRABAJO": "Mesas",
                "BENEFICIARIOS_MESAS": "Beneficiarios",
                "TOTAL_TESTIGOS": "Testigos",
            })
            fig_operativo = px.bar(
                operativo_long,
                x="TEMPLO",
                y="Cantidad",
                color="Indicador",
                barmode="group",
                title="Operación territorial por templo",
                text_auto=".0f",
            )
            fig_operativo.update_layout(height=380, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, xaxis_title="", yaxis_title="Cantidad")
            st.plotly_chart(fig_operativo, width="stretch")

        tabla_general, tabla_growth, tabla_drop = st.tabs(["Resumen por templo", "Mayor crecimiento", "Mayor caída"])
        ranking_informe = asignacion_reporte_mapa.copy()
        ranking_informe["VARIACION_ABSOLUTA"] = pd.to_numeric(ranking_informe.get("VARIACION_ABSOLUTA", 0), errors="coerce").fillna(0)
        if "IGLESIA" not in ranking_informe.columns:
            ranking_informe["IGLESIA"] = ranking_informe.get("TEMPLO_ASIGNADO_FINAL", ranking_informe.get("IGLESIA_ACTUAL", ""))
        with tabla_general:
            st.dataframe(
                resumen_general_pdf[[
                    "TEMPLO", "PUESTOS_ASIGNADOS", "VOTOS_2026_ASIGNADOS", "VOTOS_2023_ASIGNADOS",
                    "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "ACTIVIDADES", "VOLANTEOS",
                    "MESAS_TRABAJO", "BENEFICIARIOS_MESAS", "TOTAL_TESTIGOS", "LIDERES",
                    "BENEFICIARIOS_MESAS_TRABAJO",
                ]],
                hide_index=True,
                width="stretch",
            )
        with tabla_growth:
            st.dataframe(
                ranking_informe.sort_values("VARIACION_ABSOLUTA", ascending=False).head(10)[[
                    "PUESTO", "IGLESIA", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL",
                ]],
                hide_index=True,
                width="stretch",
            )
        with tabla_drop:
            st.dataframe(
                ranking_informe.sort_values("VARIACION_ABSOLUTA", ascending=True).head(10)[[
                    "PUESTO", "IGLESIA", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL",
                ]],
                hide_index=True,
                width="stretch",
            )

        pdf_general = generar_pdf_electoral_general(asignacion_reporte_mapa, puestos, iglesias, actividades, mesas, testigos_resumen)
        st.download_button(
            "Descargar informe electoral PDF",
            pdf_general,
            "informe_electoral_general_kennedy.pdf",
            "application/pdf",
            key="dl_informe_electoral_general_pdf",
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
            "Revisar recuperación territorial y mesa de trabajo.",
            "Revisar presencia y agenda comunitaria.",
            "Validar logística por distancia al templo.",
        ],
        default="Mantener seguimiento regular.",
    )

    st.markdown("### Indicadores principales")
    col_total, col_visibles, col_ajustes, col_criticos, col_valladolid = st.columns(5)

    st.markdown("### Filtros de análisis")
    filtro_asignacion_templo = st.selectbox("Templo vigente", ["Todos los templos"] + TEMPLOS_OFICIALES, key="filtro_asignacion_templo")

    asignacion_filtrada = asignacion_vista.copy()
    if filtro_asignacion_templo != "Todos los templos":
        asignacion_filtrada = asignacion_filtrada[asignacion_filtrada["TEMPLO_ASIGNADO_FINAL"].eq(filtro_asignacion_templo)].copy()

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
    st.markdown('<div class="layer-hint">Capas del mapa: use estos controles para ajustar la lectura sin ocultar las convenciones ejecutivas.</div>', unsafe_allow_html=True)
    with st.expander("Capas del mapa de asignación", expanded=True):
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
    html_map = cached_crear_mapa_asignacion_html(asignacion_filtrada, iglesias, tuple(asignacion_layers.items()))
    with st.container(border=True):
        st.components.v1.html(html_map, height=790)

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
                        (metodologia_actual_label, fmt_number(puesto_row.get("VOTOS_2026"), 0)),
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

    resumen_documental = crear_resumen_asignacion_por_columna(asignacion_vista, "IGLESIA_ACTUAL")
    resumen_final = crear_resumen_asignacion(asignacion_vista)
    impacto = resumen_final[["TEMPLO", "PUESTOS_ASIGNADOS", "VOTOS_2026_ASIGNADOS"]].merge(
        resumen_documental[["TEMPLO", "PUESTOS", "VOTOS_2026"]],
        on="TEMPLO",
        how="left",
        suffixes=("_VIGENTE", "_ORIGINAL"),
    )
    impacto["PUESTOS"] = impacto["PUESTOS"].fillna(0)
    impacto["VOTOS_2026"] = impacto["VOTOS_2026"].fillna(0)
    impacto["DELTA_PUESTOS"] = impacto["PUESTOS_ASIGNADOS"] - impacto["PUESTOS"]
    impacto["DELTA_VOTOS_2026"] = impacto["VOTOS_2026_ASIGNADOS"] - impacto["VOTOS_2026"]
    impacto_total = pd.DataFrame(
        [
            {
                "TEMPLO": "TOTAL KENNEDY",
                "PUESTOS_ASIGNADOS": impacto["PUESTOS_ASIGNADOS"].sum(),
                "VOTOS_2026_ASIGNADOS": impacto["VOTOS_2026_ASIGNADOS"].sum(),
                "PUESTOS": impacto["PUESTOS"].sum(),
                "VOTOS_2026": impacto["VOTOS_2026"].sum(),
                "DELTA_PUESTOS": impacto["DELTA_PUESTOS"].sum(),
                "DELTA_VOTOS_2026": impacto["DELTA_VOTOS_2026"].sum(),
            }
        ]
    )
    impacto = pd.concat([impacto_total, impacto], ignore_index=True)
    st.markdown("### Impacto por templo")
    st.caption("Primero se muestra el total Kennedy y debajo el detalle por templo, comparando asignación vigente contra el documento base.")
    st.dataframe(impacto, hide_index=True, width="stretch")

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

    st.markdown("### Exportables compactos")
    tabla_templos = crear_tabla_puestos_por_templo(asignacion_vista)
    informe_general_asignacion = generar_informe_territorial(asignacion_vista, actividades, mesas, testigos_resumen)
    excel_asignacion = exportar_asignacion_excel(asignacion_vista, resumen_final, tabla_templos)
    excel_cambios = multi_sheet_excel_bytes(
        {
            "cambios_guardados": cambios_show,
            "historial": historial_puestos_db,
            "impacto_templo": impacto,
        }
    )
    excel_por_templo = exportar_asignacion_por_templo_excel(asignacion_vista, actividades, mesas, testigos_resumen)
    e1, e2, e3, e4 = st.columns(4)
    with e1:
        st.download_button("Descargar asignación consolidada XLSX", excel_asignacion, "asignacion_consolidada_vigente.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with e2:
        st.download_button("Descargar cambios guardados XLSX", excel_cambios, "cambios_guardados_asignacion.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with e3:
        st.download_button("Descargar informe por templo", excel_por_templo, "informe_por_templo_asignacion.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with e4:
        st.download_button("Descargar informe general", informe_general_asignacion.encode("utf-8"), "informe_general_asignacion.md", "text/markdown")

with tab_mesas:
    st.subheader("Mesas de trabajo")
    st.markdown(
        """
        <div class="section-card">
        <b>Flujo de decisión operativa.</b><br>
        Esta vista sirve para revisar las mesas de trabajo, validar su templo vigente, guardar ajustes manuales
        y descargar un reporte de cambios. Cada ajuste queda persistido en la base activa y en el historial.
        </div>
        """,
        unsafe_allow_html=True,
    )

    mesas_vista = mesas.copy()
    if "MESA_ID" not in mesas_vista.columns:
        st.warning("La base de mesas no tiene identificador MESA_ID; no se pueden guardar ajustes de forma segura.")
    else:
        actuales_mesas_db = obtener_ajustes_actuales_df("mesa")
        historial_mesas_db = obtener_historial_ajustes(limit=5000)
        if not historial_mesas_db.empty:
            historial_mesas_db = historial_mesas_db[historial_mesas_db["entidad"].eq("mesa")].copy()

        mesas_vista["MESA_ID_TXT"] = mesas_vista["MESA_ID"].astype(str)
        mesas_vista["TEMPLO_VIGENTE"] = mesas_vista.get("IGLESIA", pd.Series("", index=mesas_vista.index)).fillna("")
        mesas_vista["TEMPLO_ORIGINAL"] = mesas_vista.get("IGLESIA_ORIGINAL", mesas_vista["TEMPLO_VIGENTE"]).fillna("")
        mesas_vista["BARRIO_ORIGINAL"] = mesas_vista.get("BARRIO_ORIGINAL", mesas_vista.get("BARRIO", pd.Series("", index=mesas_vista.index))).fillna("")
        mesas_vista["LIDER_ORIGINAL"] = mesas_vista.get("LIDER_ORIGINAL", mesas_vista.get("LIDER", pd.Series("", index=mesas_vista.index))).fillna("")
        ajuste_mesa_any = (
            mesas_vista.get("TEMPLO_AJUSTADO", pd.Series(False, index=mesas_vista.index)).fillna(False).astype(bool)
            | mesas_vista.get("BARRIO_AJUSTADO", pd.Series(False, index=mesas_vista.index)).fillna(False).astype(bool)
            | mesas_vista.get("LIDER_AJUSTADO", pd.Series(False, index=mesas_vista.index)).fillna(False).astype(bool)
        )
        mesas_vista["ESTADO_AJUSTE"] = np.where(ajuste_mesa_any, "Ajustada manualmente", "Original")

        st.markdown("### Indicadores de mesas")
        mt1, mt2, mt3, mt4 = st.columns(4)
        mesas_sin_coord_total = int(mesas_vista[["LATITUD", "LONGITUD"]].isna().any(axis=1).sum()) if {"LATITUD", "LONGITUD"}.issubset(mesas_vista.columns) else 0
        with mt1:
            metric_card("Total mesas", fmt_number(len(mesas_vista), 0), icon="👥")
        with mt2:
            mesas_ajustadas_ids = set(st.session_state.get("ajustes_mesas", {}).keys())
            mesas_ajustadas_ids |= set(st.session_state.get("ajustes_mesas_barrio", {}).keys())
            mesas_ajustadas_ids |= set(st.session_state.get("ajustes_mesas_lider", {}).keys())
            metric_card("Ajustes guardados", fmt_number(len(mesas_ajustadas_ids), 0), icon="💾")
        with mt3:
            metric_card("Mesas sin coordenadas", fmt_number(mesas_sin_coord_total, 0), icon="🧭")
        with mt4:
            metric_card("Templos con mesas", fmt_number(mesas_vista["TEMPLO_VIGENTE"].nunique(), 0), icon="📍")

        st.markdown("### Filtros de revisión")
        mf1, mf2, mf3 = st.columns([1, 1, 1.4])
        with mf1:
            filtro_mesa_templo = st.selectbox("Templo vigente", ["Todos los templos"] + TEMPLOS_OFICIALES, key="mesas_tab_templo")
        with mf2:
            filtro_mesa_estado = st.selectbox("Estado de ajuste", ["Todos", "Original", "Ajustada manualmente"], key="mesas_tab_estado")
        with mf3:
            filtro_mesa_texto = st.text_input("Buscar por tema, barrio, líder o nombre", key="mesas_tab_busqueda")

        mesas_filtradas = mesas_vista.copy()
        if filtro_mesa_templo != "Todos los templos":
            mesas_filtradas = mesas_filtradas[mesas_filtradas["TEMPLO_VIGENTE"].eq(filtro_mesa_templo)].copy()
        if filtro_mesa_estado != "Todos":
            mesas_filtradas = mesas_filtradas[mesas_filtradas["ESTADO_AJUSTE"].eq(filtro_mesa_estado)].copy()
        if filtro_mesa_texto.strip():
            texto = filtro_mesa_texto.strip().upper()
            columnas_busqueda = [c for c in ["NOMBRE_GESTION", "TEMA", "BARRIO", "LIDER", "ESTADO"] if c in mesas_filtradas.columns]
            if columnas_busqueda:
                mask = mesas_filtradas[columnas_busqueda].fillna("").astype(str).agg(" ".join, axis=1).str.upper().str.contains(texto, regex=False)
                mesas_filtradas = mesas_filtradas[mask].copy()

        st.markdown("### Mapa de mesas filtradas")
        if {"LATITUD", "LONGITUD"}.issubset(mesas_filtradas.columns):
            mesas_mapa = mesas_filtradas.copy()
            mesas_mapa["IGLESIA"] = mesas_mapa["TEMPLO_VIGENTE"]
            mesas_mapa_con_coord = mesas_mapa.dropna(subset=["LATITUD", "LONGITUD"])
            if mesas_mapa_con_coord.empty:
                st.info("No hay mesas con coordenadas para los filtros actuales.")
            else:
                layer_cfg = {
                        "contorno": True,
                        "upz": False,
                        "heat": False,
                        "puestos": False,
                        "actividades": False,
                        "mesas": True,
                        "templos": True,
                    }
                html_map = cached_crear_mapa_html(
                    puestos.head(0),
                    iglesias,
                    actividades.head(0),
                    mesas_mapa,
                    "Vista operativa",
                    tuple(layer_cfg.items()),
                )
                with st.container(border=True):
                    st.components.v1.html(html_map, height=620)
        else:
            st.info("La base de mesas no tiene columnas LATITUD y LONGITUD para mostrar el mapa.")

        st.markdown("### Mesa operativa filtrada")
        mesas_cols = [
            "MESA_ID", "NOMBRE_GESTION", "TEMA", "BARRIO", "TEMPLO_ORIGINAL", "TEMPLO_VIGENTE",
            "ESTADO_AJUSTE", "LIDER", "ESTADO", "BENEFICIARIOS",
        ]
        mesas_cols = [c for c in mesas_cols if c in mesas_filtradas.columns]
        st.dataframe(mesas_filtradas[mesas_cols], hide_index=True, width="stretch")

        st.markdown("### Ajuste manual de mesa")
        if mesas_filtradas.empty:
            st.info("No hay mesas con los filtros actuales.")
        else:
            mesa_selector_df = mesas_filtradas.copy()
            mesa_selector_df["LABEL"] = mesa_selector_df.apply(
                lambda r: f"{r.get('MESA_ID')} · {r.get('NOMBRE_GESTION', r.get('TEMA', 'SIN NOMBRE'))} · {r.get('BARRIO', 'SIN BARRIO')} · {r.get('TEMPLO_VIGENTE', '')}",
                axis=1,
            )
            mesa_label = st.selectbox("Mesa de trabajo", mesa_selector_df["LABEL"].tolist(), key="mesas_tab_selector")
            mesa_row = mesa_selector_df[mesa_selector_df["LABEL"].eq(mesa_label)].iloc[0]
            mesa_id_raw = mesa_row.get("MESA_ID")
            try:
                mesa_id_key = int(mesa_id_raw)
            except Exception:
                mesa_id_key = str(mesa_id_raw)
            templo_actual_mesa = mesa_row.get("TEMPLO_VIGENTE")
            mesa_index = TEMPLOS_OFICIALES.index(templo_actual_mesa) if templo_actual_mesa in TEMPLOS_OFICIALES else 0

            ma1, ma2 = st.columns([2, 1])
            with ma1:
                st.dataframe(
                    pd.DataFrame(
                        [
                            ("Mesa ID", mesa_row.get("MESA_ID")),
                            ("Nombre", mesa_row.get("NOMBRE_GESTION", "")),
                            ("Tema", mesa_row.get("TEMA", "")),
                            ("Barrio original", mesa_row.get("BARRIO_ORIGINAL", "")),
                            ("Barrio vigente", mesa_row.get("BARRIO", "")),
                            ("Templo original", mesa_row.get("TEMPLO_ORIGINAL", "")),
                            ("Templo vigente", mesa_row.get("TEMPLO_VIGENTE", "")),
                            ("Líder original", mesa_row.get("LIDER_ORIGINAL", "")),
                            ("Líder vigente", mesa_row.get("LIDER", "")),
                            ("Estado", mesa_row.get("ESTADO", "")),
                        ],
                        columns=["Campo", "Valor"],
                    ),
                    hide_index=True,
                    width="stretch",
                )
                historial_mesa = historial_mesas_db[historial_mesas_db["entidad_id"].astype(str).eq(str(mesa_id_key))].copy() if not historial_mesas_db.empty else pd.DataFrame()
                if not historial_mesa.empty:
                    st.markdown("Historial de esta mesa")
                    historial_cols_mesa = [
                        "creado_en", "templo_anterior", "templo_nuevo", "barrio_anterior", "barrio_nuevo",
                        "lider_anterior", "lider_nuevo", "usuario", "motivo",
                    ]
                    historial_cols_mesa = [c for c in historial_cols_mesa if c in historial_mesa.columns]
                    st.dataframe(historial_mesa[historial_cols_mesa], hide_index=True, width="stretch")
            with ma2:
                templo_mesa_nuevo = st.selectbox("Templo vigente final", TEMPLOS_OFICIALES, index=mesa_index, key="mesas_tab_templo_final")
                barrio_mesa_nuevo = st.text_input("Barrio vigente", value=str(mesa_row.get("BARRIO", "") or ""), key="mesas_tab_barrio_final")
                lider_mesa_nuevo = st.text_input("Líder vigente", value=str(mesa_row.get("LIDER", "") or ""), key="mesas_tab_lider_final")
                nota_mesa = st.text_area("Justificación o nota", key="mesas_tab_nota", height=120)
                if st.button("Guardar cambio definitivo de mesa", key="mesas_tab_guardar"):
                    st.session_state.setdefault("ajustes_mesas", {})
                    st.session_state.setdefault("ajustes_mesas_barrio", {})
                    st.session_state.setdefault("ajustes_mesas_lider", {})
                    barrio_mesa_nuevo = barrio_mesa_nuevo.strip()
                    lider_mesa_nuevo = lider_mesa_nuevo.strip()
                    st.session_state["ajustes_mesas"][mesa_id_key] = templo_mesa_nuevo
                    st.session_state["ajustes_mesas_barrio"][mesa_id_key] = barrio_mesa_nuevo
                    st.session_state["ajustes_mesas_lider"][mesa_id_key] = lider_mesa_nuevo
                    registrar_ajuste_en_db(
                        session_key="ajustes_mesas",
                        entity_id=mesa_id_key,
                        nombre_entidad=mesa_row.get("NOMBRE_GESTION", mesa_row.get("TEMA", "")),
                        templo_nuevo=templo_mesa_nuevo,
                        barrio_nuevo=barrio_mesa_nuevo,
                        lider_nuevo=lider_mesa_nuevo,
                        motivo=nota_mesa or "Ajuste manual desde pestaña Mesas de trabajo",
                    )
                    guardar_ajustes_guardados()
                    st.success(f"Mesa {mesa_id_key} guardada en {templo_mesa_nuevo}, barrio {barrio_mesa_nuevo or 'sin dato'} y líder {lider_mesa_nuevo or 'sin dato'}.")
                    st.rerun()
                if st.button("Limpiar ajustes de mesas", key="mesas_tab_limpiar"):
                    total_limpiados = limpiar_ajustes_en_db("ajustes_mesas", motivo="Limpieza manual desde pestaña Mesas de trabajo")
                    st.session_state["ajustes_mesas"] = {}
                    st.session_state["ajustes_mesas_barrio"] = {}
                    st.session_state["ajustes_mesas_lider"] = {}
                    guardar_ajustes_guardados()
                    st.info(f"Se limpiaron {fmt_number(total_limpiados, 0)} ajuste(s) de mesas en la base.")
                    st.rerun()

        st.markdown("### Resumen de cambios de mesas")
        if not actuales_mesas_db.empty:
            cambios_mesas_show = actuales_mesas_db[actuales_mesas_db["entidad"].eq("mesa")].copy()
        else:
            cambios_mesas_show = pd.DataFrame(columns=AJUSTES_ACTUALES_COLUMNS)
        st.dataframe(cambios_mesas_show, hide_index=True, width="stretch")

        excel_mesas = multi_sheet_excel_bytes(
            {
                "mesas_vigentes": mesas_vista.drop(columns=[c for c in ["MESA_ID_TXT"] if c in mesas_vista.columns]),
                "mesas_filtradas": mesas_filtradas.drop(columns=[c for c in ["MESA_ID_TXT"] if c in mesas_filtradas.columns]),
                "cambios_guardados": cambios_mesas_show,
                "historial_mesas": historial_mesas_db,
            }
        )
        st.download_button(
            "Descargar reporte de mesas XLSX",
            excel_mesas,
            "reporte_mesas_trabajo_vigente.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_reporte_mesas_tab",
        )

with tab_apoyos:
    st.subheader("Apoyos ciudadanos")
    if apoyos_ciudadanos.empty:
        st.info("No se encontró el archivo `Gestion apoyo ciudadano Kennedy.xlsx`. Déjelo en la carpeta principal de Kennedy o en `data/` para activar esta sección.")
    else:
        fuente_apoyos = APOYOS_CIUDADANOS_FILE.name if APOYOS_CIUDADANOS_FILE else "Gestion apoyo ciudadano Kennedy.xlsx"
        st.markdown(
            f"""
            <div class="section-card">
            Esta vista analiza los apoyos ciudadanos registrados en <b>{fuente_apoyos}</b>.
            Se omite el nombre de quien gestionó cada caso y se concentra la lectura en templo, temas, estado del trámite y resultado de la gestión.
            </div>
            """,
            unsafe_allow_html=True,
        )

        templos_apoyo = [t for t in TEMPLOS_OFICIALES if t in set(apoyos_ciudadanos["TEMPLO"])]
        if "SIN TEMPLO OFICIAL" in set(apoyos_ciudadanos["TEMPLO"]):
            templos_apoyo.append("SIN TEMPLO OFICIAL")
        filtro_apoyo_templo = st.selectbox("Filtrar apoyos por templo", ["Todos los templos"] + templos_apoyo, key="filtro_apoyos_templo")

        apoyos_view = apoyos_ciudadanos.copy()
        if filtro_apoyo_templo != "Todos los templos":
            apoyos_view = apoyos_view[apoyos_view["TEMPLO"].eq(filtro_apoyo_templo)].copy()

        total_apoyos = len(apoyos_view)
        finalizados = int(apoyos_view["ESTADO_TRAMITE"].eq("FINALIZADO").sum())
        en_proceso = int(apoyos_view["ESTADO_TRAMITE"].isin(["EN PROCESO", "PENDIENTE"]).sum())
        positivos = int(apoyos_view["RESULTADO_GESTION"].eq("POSITIVO").sum())
        pendientes_resultado = int(apoyos_view["RESULTADO_GESTION"].eq("PENDIENTE").sum())
        temas_total = int(apoyos_view["TEMA"].nunique())
        barrios_total = int(apoyos_view["BARRIO"].replace("Sin dato", np.nan).dropna().nunique())

        a1, a2, a3, a4 = st.columns(4)
        with a1:
            metric_card("Apoyos registrados", fmt_number(total_apoyos, 0), icon="users")
        with a2:
            metric_card("Finalizados", fmt_number(finalizados, 0), icon="clipboard")
        with a3:
            metric_card("En proceso / pendientes", fmt_number(en_proceso, 0), icon="activity")
        with a4:
            metric_card("Resultados positivos", fmt_number(positivos, 0), icon="trend-up")

        a5, a6, a7, a8 = st.columns(4)
        with a5:
            metric_card("Temas atendidos", fmt_number(temas_total, 0), icon="file-text")
        with a6:
            metric_card("Barrios reportados", fmt_number(barrios_total, 0), icon="map")
        with a7:
            metric_card("Pendientes de resultado", fmt_number(pendientes_resultado, 0), icon="status")
        with a8:
            tasa_positiva = positivos / total_apoyos if total_apoyos else 0
            metric_card("Tasa positiva", fmt_pct(tasa_positiva), icon="star")

        if apoyos_view.empty:
            st.info("No hay apoyos ciudadanos para el filtro seleccionado.")
        else:
            g1, g2 = st.columns(2)
            with g1:
                apoyos_templo = apoyos_view.groupby("TEMPLO", dropna=False).size().reset_index(name="APOYOS")
                fig_templo = px.bar(
                    apoyos_templo.sort_values("APOYOS", ascending=True),
                    x="APOYOS",
                    y="TEMPLO",
                    orientation="h",
                    color="TEMPLO",
                    color_discrete_map=COLORES_TEMPLOS,
                    category_orders=PLOTLY_TEMPLO_ORDERS,
                    title="Apoyos ciudadanos por templo",
                    text="APOYOS",
                )
                fig_templo.update_traces(texttemplate="%{text:.0f}", textposition="outside", cliponaxis=False)
                fig_templo.update_layout(height=390, showlegend=False, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, xaxis_title="Apoyos", yaxis_title="")
                st.plotly_chart(fig_templo, width="stretch")

            with g2:
                apoyos_tema = apoyos_view.groupby("TEMA", dropna=False).size().reset_index(name="APOYOS").sort_values("APOYOS", ascending=False).head(12)
                fig_tema = px.bar(
                    apoyos_tema.sort_values("APOYOS", ascending=True),
                    x="APOYOS",
                    y="TEMA",
                    orientation="h",
                    title="Temas con mayor demanda ciudadana",
                    color="APOYOS",
                    color_continuous_scale=["#DBEAFE", "#2563EB"],
                    text="APOYOS",
                )
                fig_tema.update_traces(texttemplate="%{text:.0f}", textposition="outside", cliponaxis=False)
                fig_tema.update_layout(height=390, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, xaxis_title="Apoyos", yaxis_title="", coloraxis_showscale=False)
                st.plotly_chart(fig_tema, width="stretch")

            g3, g4 = st.columns(2)
            with g3:
                estado_count = apoyos_view.groupby("ESTADO_TRAMITE", dropna=False).size().reset_index(name="APOYOS")
                fig_estado = px.pie(
                    estado_count,
                    names="ESTADO_TRAMITE",
                    values="APOYOS",
                    hole=0.48,
                    title="Estado actual del trámite",
                    color="ESTADO_TRAMITE",
                    color_discrete_map={"FINALIZADO": COLOR_GREEN, "EN PROCESO": COLOR_BLUE, "PENDIENTE": COLOR_ORANGE},
                )
                fig_estado.update_layout(height=360, paper_bgcolor="white", font_color=COLOR_TEXT)
                st.plotly_chart(fig_estado, width="stretch")

            with g4:
                resultado_count = apoyos_view.groupby("RESULTADO_GESTION", dropna=False).size().reset_index(name="APOYOS")
                fig_resultado = px.bar(
                    resultado_count.sort_values("APOYOS", ascending=False),
                    x="RESULTADO_GESTION",
                    y="APOYOS",
                    color="RESULTADO_GESTION",
                    title="Resultado de la gestión",
                    text="APOYOS",
                    color_discrete_map={"POSITIVO": COLOR_GREEN, "PENDIENTE": COLOR_ORANGE, "NEGATIVO": COLOR_RED},
                )
                fig_resultado.update_traces(texttemplate="%{text:.0f}", textposition="outside", cliponaxis=False)
                fig_resultado.update_layout(height=360, showlegend=False, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, xaxis_title="", yaxis_title="Apoyos")
                st.plotly_chart(fig_resultado, width="stretch")

            if apoyos_view["PERIODO"].replace("", np.nan).dropna().any():
                periodo_count = apoyos_view[apoyos_view["PERIODO"].ne("")].groupby(["PERIODO", "TEMPLO"], dropna=False).size().reset_index(name="APOYOS")
                fig_periodo = px.line(
                    periodo_count.sort_values("PERIODO"),
                    x="PERIODO",
                    y="APOYOS",
                    color="TEMPLO",
                    markers=True,
                    color_discrete_map=COLORES_TEMPLOS,
                    category_orders=PLOTLY_TEMPLO_ORDERS,
                    title="Evolución mensual de apoyos ciudadanos",
                )
                fig_periodo.update_layout(height=390, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, xaxis_title="", yaxis_title="Apoyos")
                st.plotly_chart(fig_periodo, width="stretch")

            resumen_apoyo_templo = apoyos_view.groupby("TEMPLO", dropna=False).agg(
                APOYOS=("CODIGO_SEGUIMIENTO", "count"),
                TEMAS=("TEMA", "nunique"),
                BARRIOS=("BARRIO", lambda s: s.replace("Sin dato", np.nan).dropna().nunique()),
                FINALIZADOS=("ESTADO_TRAMITE", lambda s: s.eq("FINALIZADO").sum()),
                EN_PROCESO=("ESTADO_TRAMITE", lambda s: s.isin(["EN PROCESO", "PENDIENTE"]).sum()),
                POSITIVOS=("RESULTADO_GESTION", lambda s: s.eq("POSITIVO").sum()),
                PENDIENTES_RESULTADO=("RESULTADO_GESTION", lambda s: s.eq("PENDIENTE").sum()),
            ).reset_index().sort_values("APOYOS", ascending=False)
            resumen_apoyo_templo["TASA_POSITIVA"] = np.where(
                resumen_apoyo_templo["APOYOS"].gt(0),
                resumen_apoyo_templo["POSITIVOS"] / resumen_apoyo_templo["APOYOS"],
                0,
            )

            st.markdown("### Resumen por templo")
            st.dataframe(resumen_apoyo_templo, hide_index=True, width="stretch")

            temas_por_templo = apoyos_view.groupby(["TEMPLO", "TEMA"], dropna=False).agg(
                APOYOS=("CODIGO_SEGUIMIENTO", "count"),
                FINALIZADOS=("ESTADO_TRAMITE", lambda s: s.eq("FINALIZADO").sum()),
                POSITIVOS=("RESULTADO_GESTION", lambda s: s.eq("POSITIVO").sum()),
                PENDIENTES=("RESULTADO_GESTION", lambda s: s.eq("PENDIENTE").sum()),
            ).reset_index().sort_values(["APOYOS", "TEMPLO"], ascending=[False, True])

            st.markdown("### Temas atendidos por templo")
            st.dataframe(temas_por_templo, hide_index=True, width="stretch")

            detalle_cols = [
                "CODIGO_SEGUIMIENTO", "ANIO", "MES_INICIO", "TIPO_ACTIVIDAD", "TEMA",
                "SUBTEMA", "BARRIO", "TEMPLO", "ESTADO_TRAMITE", "RESULTADO_GESTION",
            ]
            detalle_cols = [c for c in detalle_cols if c in apoyos_view.columns]
            st.markdown("### Detalle operativo sin gestor")
            detalle_apoyos = apoyos_view.sort_values(["ANIO", "COD_MES"], ascending=[False, False])
            st.dataframe(detalle_apoyos[detalle_cols], hide_index=True, width="stretch")

            excel_apoyos = multi_sheet_excel_bytes(
                {
                    "resumen_templo": resumen_apoyo_templo,
                    "temas_por_templo": temas_por_templo,
                    "detalle_sin_gestor": detalle_apoyos[detalle_cols],
                }
            )
            st.download_button(
                "Descargar análisis de apoyos XLSX",
                excel_apoyos,
                "analisis_apoyos_ciudadanos_kennedy.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_apoyos_ciudadanos",
            )

with tab_iglesia:
    st.subheader("Análisis por iglesia")
    cols_show = [
        "IGLESIA", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL",
        "PUESTOS", "ACTIVIDADES_CAMPANA", "MESAS_TRABAJO", "TOTAL_TESTIGOS", "LIDERES",
        "BENEFICIARIOS_MESAS_TRABAJO", "PUESTO_MAYOR_VOTACION",
        "PUESTO_MAYOR_CAIDA", "PUESTO_MAYOR_CRECIMIENTO"
    ]
    cols_show = [c for c in cols_show if c in resumen_iglesia_f.columns]
    st.dataframe(resumen_iglesia_f[cols_show], width="stretch", hide_index=True)
    st.download_button("Descargar análisis por iglesia en Excel", to_excel_bytes(resumen_iglesia_f[cols_show], "Analisis Iglesia"), "analisis_iglesia.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_iglesia")

    for iglesia in iglesias["IGLESIA"].tolist():
        sub_puestos = puestos_f[puestos_f["IGLESIA"].eq(iglesia)].copy()
        sub_acts = actividades_f[actividades_f["IGLESIA"].eq(iglesia)].copy()
        sub_mesas = mesas_f[mesas_f["IGLESIA"].eq(iglesia)].copy()
        res = resumen_iglesia_f[resumen_iglesia_f["IGLESIA"].eq(iglesia)]
        
        with st.expander(f"Ficha territorial: {iglesia}", expanded=False):
            r_iglesia = res.iloc[0] if not res.empty else pd.Series(dtype=object)
            votos_2026 = sum_numeric(sub_puestos, "VOTOS_2026")
            votos_2023 = sum_numeric(sub_puestos, "VOTOS_2023")
            var_abs = votos_2026 - votos_2023
            var_pct = var_abs / votos_2023 if votos_2023 else np.nan
            benef_total, benef_internos, benef_externos = calcular_beneficiarios_mesas(sub_mesas)
            testigos_total, testigos_lideres, testigos_benef_mesas = testigos_metricas_templo(testigos_resumen, iglesia)

            metricas_templo = [
                ("Mesas de trabajo", fmt_number(len(sub_mesas), 0)),
                ("Actividades campaña", fmt_number(len(sub_acts), 0)),
                ("Beneficiarios mesas", fmt_number(benef_total, 0)),
                ("Beneficiarios internos", fmt_number(benef_internos, 0)),
                ("Beneficiarios externos", fmt_number(benef_externos, 0)),
                (metodologia_actual_label, fmt_number(votos_2026, 0)),
                (metodologia_base_label, fmt_number(votos_2023, 0)),
                ("Variación", fmt_variacion(var_abs, var_pct)),
                ("Cámara 2026", fmt_number(sum_numeric(sub_puestos, "CAMARA_2026"), 0)),
                ("Senado 2026", fmt_number(sum_numeric(sub_puestos, "SENADO_2026"), 0)),
                ("JAL 2023", fmt_number(sum_numeric(sub_puestos, "JAL_2023"), 0)),
                ("Concejo 2023", fmt_number(sum_numeric(sub_puestos, "MIRA_CONCEJO_2023"), 0)),
                ("Testigos electorales", fmt_number(testigos_total, 0)),
                ("Testigos líderes", fmt_number(testigos_lideres, 0)),
                ("Testigos benef. mesas", fmt_number(testigos_benef_mesas, 0)),
            ]

            for start in range(0, len(metricas_templo), 4):
                cols_metricas = st.columns(4)
                for col, (label, value) in zip(cols_metricas, metricas_templo[start:start + 4]):
                    with col:
                        metric_card(label, value)

            templo_coord = iglesias[iglesias["IGLESIA"].eq(iglesia)]
            templo_row = templo_coord.iloc[0] if not templo_coord.empty else None
            if templo_row is not None:
                lat_t = templo_row["LATITUD"]
                lon_t = templo_row["LONGITUD"]
                html_map = cached_submap_html(iglesia, lat_t, lon_t, sub_puestos, sub_mesas)
                st.markdown("### Mapa territorial del templo")
                st.components.v1.html(html_map, height=460)
            else:
                st.info("Sin coordenadas del templo.")

            st.markdown("### Mesas de trabajo asignadas a este templo")
            mesas_cols_templo = [
                "MESA_ID", "NOMBRE_GESTION", "TEMA", "BARRIO", "BENEFICIARIOS",
                "BENEFICIARIOS_REFERIDOS", "BENEFICIARIOS_NO_INFOMIRA", "LIDER", "ESTADO",
            ]
            mesas_cols_templo = [c for c in mesas_cols_templo if c in sub_mesas.columns]
            if sub_mesas.empty:
                st.info("No hay mesas asignadas a este templo con la asignación vigente.")
            else:
                st.dataframe(sub_mesas[mesas_cols_templo].sort_values("NOMBRE_GESTION" if "NOMBRE_GESTION" in sub_mesas.columns else "MESA_ID"), hide_index=True, width="stretch")

            st.markdown("### Puestos de votación asignados a este templo")
            puestos_cols_templo = [
                "PUESTO", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL",
                "CAMARA_2026", "SENADO_2026", "JAL_2023", "MIRA_CONCEJO_2023",
            ]
            puestos_cols_templo = [c for c in puestos_cols_templo if c in sub_puestos.columns]
            if sub_puestos.empty:
                st.info("No hay puestos asignados a este templo.")
            else:
                st.dataframe(sub_puestos[puestos_cols_templo].sort_values("VOTOS_2026", ascending=False), width="stretch", hide_index=True)

            cache_key = f"pdf_bytes_{iglesia}"
            if cache_key not in st.session_state:
                if st.button(f"Preparar informe PDF de {iglesia}", key=f"btn_prep_{iglesia}"):
                    with st.spinner(f"Generando PDF de alta calidad para {iglesia}... (puede tardar unos segundos)"):
                        st.session_state[cache_key] = generar_pdf_templo(iglesia, metricas_templo, sub_puestos, sub_mesas, templo_row)
                        if hasattr(st, "rerun"):
                            st.rerun()
                        else:
                            st.experimental_rerun()
            else:
                st.success("PDF generado y listo para descargar.")
                st.download_button(
                    f"Descargar informe PDF {iglesia}",
                    st.session_state[cache_key],
                    f"informe_templo_{iglesia.lower().replace(' ', '_')}.pdf",
                    "application/pdf",
                    key=f"dl_pdf_templo_{iglesia}",
                )

with tab_puesto:
    st.subheader("Análisis por puesto de votación")

    st.markdown("### Testigos electorales Congreso 2026")
    if testigos_resumen.empty:
        st.info("No se encontró el resumen agregado de testigos. Agregue `data/testigos_resumen_2026.csv` para activar este análisis.")
    else:
        testigos_base_total = pd.DataFrame({"TEMPLO": TEMPLOS_OFICIALES}).merge(testigos_resumen, on="TEMPLO", how="left").fillna(0)

        puestos_por_templo = puestos[puestos["IGLESIA"].isin(TEMPLOS_OFICIALES)].groupby("IGLESIA", dropna=False).agg(
            PUESTOS=("PUESTO", "nunique"),
            VOTOS_2026=("VOTOS_2026", "sum"),
        ).reset_index().rename(columns={"IGLESIA": "TEMPLO"})
        testigos_analisis_total = testigos_base_total.merge(puestos_por_templo, on="TEMPLO", how="left").fillna(0)
        testigos_numeric_cols = [
            "TOTAL_TESTIGOS",
            "TESTIGOS_MESA_O_REMANENTE",
            "TESTIGOS_COMISION_ESCRUTADORA",
            "BENEFICIARIOS_MESAS_TRABAJO",
            "CANTIDAD_MESAS_TRABAJO_ASOCIADAS",
            "TESTIGOS_DOBLE_ROL",
            "LIDERES",
            "NO_LIDERES",
            "TESTIGOS_CON_REFERIDOS",
            "REFERIDOS_REGISTRADOS",
            "REFERIDOS_INACTIVOS",
            "REFERIDOS_ACTIVOS_ESTIMADOS",
            "PUESTOS",
            "VOTOS_2026",
        ]
        for col in testigos_numeric_cols:
            if col not in testigos_analisis_total.columns:
                testigos_analisis_total[col] = 0
            testigos_analisis_total[col] = pd.to_numeric(testigos_analisis_total[col], errors="coerce").fillna(0)

        total_testigos_kennedy = float(testigos_analisis_total["TOTAL_TESTIGOS"].sum())
        testigos_analisis_total["% DEL TOTAL KENNEDY"] = np.where(
            total_testigos_kennedy > 0,
            testigos_analisis_total["TOTAL_TESTIGOS"] / total_testigos_kennedy,
            0,
        )
        testigos_analisis_total["% MESA / REMANENTE"] = np.where(
            testigos_analisis_total["TOTAL_TESTIGOS"].gt(0),
            testigos_analisis_total["TESTIGOS_MESA_O_REMANENTE"] / testigos_analisis_total["TOTAL_TESTIGOS"],
            0,
        )
        testigos_analisis_total["% COMISION"] = np.where(
            testigos_analisis_total["TOTAL_TESTIGOS"].gt(0),
            testigos_analisis_total["TESTIGOS_COMISION_ESCRUTADORA"] / testigos_analisis_total["TOTAL_TESTIGOS"],
            0,
        )
        testigos_analisis_total["% BENEFICIARIOS MESAS"] = np.where(
            testigos_analisis_total["TOTAL_TESTIGOS"].gt(0),
            testigos_analisis_total["BENEFICIARIOS_MESAS_TRABAJO"] / testigos_analisis_total["TOTAL_TESTIGOS"],
            0,
        )
        testigos_analisis_total["% LIDERES"] = np.where(
            testigos_analisis_total["TOTAL_TESTIGOS"].gt(0),
            testigos_analisis_total["LIDERES"] / testigos_analisis_total["TOTAL_TESTIGOS"],
            0,
        )
        testigos_analisis_total["% CON REFERIDOS"] = np.where(
            testigos_analisis_total["TOTAL_TESTIGOS"].gt(0),
            testigos_analisis_total["TESTIGOS_CON_REFERIDOS"] / testigos_analisis_total["TOTAL_TESTIGOS"],
            0,
        )
        testigos_analisis_total["TESTIGOS_POR_PUESTO"] = np.where(
            testigos_analisis_total["PUESTOS"].gt(0),
            testigos_analisis_total["TOTAL_TESTIGOS"] / testigos_analisis_total["PUESTOS"],
            0,
        )
        testigos_analisis_total["VOTOS_2026_POR_TESTIGO"] = np.where(
            testigos_analisis_total["TOTAL_TESTIGOS"].gt(0),
            testigos_analisis_total["VOTOS_2026"] / testigos_analisis_total["TOTAL_TESTIGOS"],
            0,
        )
        testigos_analisis_total["REFERIDOS POR TESTIGO"] = np.where(
            testigos_analisis_total["TOTAL_TESTIGOS"].gt(0),
            testigos_analisis_total["REFERIDOS_REGISTRADOS"] / testigos_analisis_total["TOTAL_TESTIGOS"],
            0,
        )
        testigos_analisis_total["REFERIDOS POR LIDER"] = np.where(
            testigos_analisis_total["LIDERES"].gt(0),
            testigos_analisis_total["REFERIDOS_REGISTRADOS"] / testigos_analisis_total["LIDERES"],
            0,
        )

        total_row = {
            "TEMPLO": "TOTAL KENNEDY",
            "TOTAL_TESTIGOS": testigos_analisis_total["TOTAL_TESTIGOS"].sum(),
            "TESTIGOS_MESA_O_REMANENTE": testigos_analisis_total["TESTIGOS_MESA_O_REMANENTE"].sum(),
            "TESTIGOS_COMISION_ESCRUTADORA": testigos_analisis_total["TESTIGOS_COMISION_ESCRUTADORA"].sum(),
            "BENEFICIARIOS_MESAS_TRABAJO": testigos_analisis_total["BENEFICIARIOS_MESAS_TRABAJO"].sum(),
            "CANTIDAD_MESAS_TRABAJO_ASOCIADAS": testigos_analisis_total["CANTIDAD_MESAS_TRABAJO_ASOCIADAS"].sum(),
            "TESTIGOS_DOBLE_ROL": testigos_analisis_total["TESTIGOS_DOBLE_ROL"].sum(),
            "LIDERES": testigos_analisis_total["LIDERES"].sum(),
            "NO_LIDERES": testigos_analisis_total["NO_LIDERES"].sum(),
            "TESTIGOS_CON_REFERIDOS": testigos_analisis_total["TESTIGOS_CON_REFERIDOS"].sum(),
            "REFERIDOS_REGISTRADOS": testigos_analisis_total["REFERIDOS_REGISTRADOS"].sum(),
            "REFERIDOS_INACTIVOS": testigos_analisis_total["REFERIDOS_INACTIVOS"].sum(),
            "REFERIDOS_ACTIVOS_ESTIMADOS": testigos_analisis_total["REFERIDOS_ACTIVOS_ESTIMADOS"].sum(),
            "PUESTOS": testigos_analisis_total["PUESTOS"].sum(),
            "VOTOS_2026": testigos_analisis_total["VOTOS_2026"].sum(),
        }
        total_row["% DEL TOTAL KENNEDY"] = 1 if total_row["TOTAL_TESTIGOS"] else 0
        total_row["% MESA / REMANENTE"] = total_row["TESTIGOS_MESA_O_REMANENTE"] / total_row["TOTAL_TESTIGOS"] if total_row["TOTAL_TESTIGOS"] else 0
        total_row["% COMISION"] = total_row["TESTIGOS_COMISION_ESCRUTADORA"] / total_row["TOTAL_TESTIGOS"] if total_row["TOTAL_TESTIGOS"] else 0
        total_row["% BENEFICIARIOS MESAS"] = total_row["BENEFICIARIOS_MESAS_TRABAJO"] / total_row["TOTAL_TESTIGOS"] if total_row["TOTAL_TESTIGOS"] else 0
        total_row["% LIDERES"] = total_row["LIDERES"] / total_row["TOTAL_TESTIGOS"] if total_row["TOTAL_TESTIGOS"] else 0
        total_row["% CON REFERIDOS"] = total_row["TESTIGOS_CON_REFERIDOS"] / total_row["TOTAL_TESTIGOS"] if total_row["TOTAL_TESTIGOS"] else 0
        total_row["TESTIGOS_POR_PUESTO"] = total_row["TOTAL_TESTIGOS"] / total_row["PUESTOS"] if total_row["PUESTOS"] else 0
        total_row["VOTOS_2026_POR_TESTIGO"] = total_row["VOTOS_2026"] / total_row["TOTAL_TESTIGOS"] if total_row["TOTAL_TESTIGOS"] else 0
        total_row["REFERIDOS POR TESTIGO"] = total_row["REFERIDOS_REGISTRADOS"] / total_row["TOTAL_TESTIGOS"] if total_row["TOTAL_TESTIGOS"] else 0
        total_row["REFERIDOS POR LIDER"] = total_row["REFERIDOS_REGISTRADOS"] / total_row["LIDERES"] if total_row["LIDERES"] else 0

        testigos_analisis = testigos_analisis_total[testigos_analisis_total["TEMPLO"].isin(selected_iglesias)].copy()
        testigos_analisis["TESTIGOS_POR_PUESTO"] = np.where(
            testigos_analisis["PUESTOS"].gt(0),
            testigos_analisis["TOTAL_TESTIGOS"] / testigos_analisis["PUESTOS"],
            0,
        )
        testigos_analisis["VOTOS_2026_POR_TESTIGO"] = np.where(
            testigos_analisis["TOTAL_TESTIGOS"].gt(0),
            testigos_analisis["VOTOS_2026"] / testigos_analisis["TOTAL_TESTIGOS"],
            0,
        )
        testigos_analisis["REFERIDOS POR TESTIGO"] = np.where(
            testigos_analisis["TOTAL_TESTIGOS"].gt(0),
            testigos_analisis["REFERIDOS_REGISTRADOS"] / testigos_analisis["TOTAL_TESTIGOS"],
            0,
        )
        testigos_analisis["REFERIDOS POR LIDER"] = np.where(
            testigos_analisis["LIDERES"].gt(0),
            testigos_analisis["REFERIDOS_REGISTRADOS"] / testigos_analisis["LIDERES"],
            0,
        )

        total_testigos = int(total_row["TOTAL_TESTIGOS"])
        total_mesa_rem = int(total_row["TESTIGOS_MESA_O_REMANENTE"])
        total_lideres = int(total_row["LIDERES"])
        total_con_referidos = int(total_row["TESTIGOS_CON_REFERIDOS"])
        total_referidos = int(total_row["REFERIDOS_REGISTRADOS"])

        tg1, tg2, tg3, tg4, tg5 = st.columns(5)
        with tg1:
            metric_card("Total Kennedy", fmt_number(total_testigos, 0), icon="👥")
        with tg2:
            metric_card("Mesa / remanente", f"{fmt_number(total_mesa_rem, 0)} ({fmt_pct(total_row['% MESA / REMANENTE'])})", icon="🗳️")
        with tg3:
            metric_card("Líderes", f"{fmt_number(total_lideres, 0)} ({fmt_pct(total_row['% LIDERES'])})", icon="⭐")
        with tg4:
            metric_card("Con referidos", f"{fmt_number(total_con_referidos, 0)} ({fmt_pct(total_row['% CON REFERIDOS'])})", icon="🔗")
        with tg5:
            metric_card("Referidos registrados", fmt_number(total_referidos, 0), icon="📋")

        st.markdown(
            '<div class="note-box"><b>Cómo leer este bloque:</b> el total Kennedy aparece primero para dar contexto general. “Líderes” muestra cuántos testigos tienen rol de liderazgo; “Con referidos” indica cuántos registraron al menos una persona; y “Referidos por líder” ayuda a comparar capacidad de movilización entre templos. “Referidos activos estimados” descuenta los referidos marcados como inactivos en la base.</div>',
            unsafe_allow_html=True,
        )

        tc1, tc2 = st.columns([1.25, 1])
        with tc1:
            roles_df = testigos_analisis[
                [
                    "TEMPLO",
                    "TESTIGOS_MESA_O_REMANENTE",
                    "TESTIGOS_COMISION_ESCRUTADORA",
                    "BENEFICIARIOS_MESAS_TRABAJO",
                    "LIDERES",
                    "TESTIGOS_CON_REFERIDOS",
                ]
            ].melt("TEMPLO", var_name="ROL", value_name="PERSONAS")
            roles_df["ROL"] = roles_df["ROL"].replace(
                {
                    "TESTIGOS_MESA_O_REMANENTE": "Mesa / remanente",
                    "TESTIGOS_COMISION_ESCRUTADORA": "Comisión escrutadora",
                    "BENEFICIARIOS_MESAS_TRABAJO": "Beneficiario mesas",
                    "LIDERES": "Líderes",
                    "TESTIGOS_CON_REFERIDOS": "Con referidos",
                }
            )
            fig_testigos = px.bar(
                roles_df,
                x="TEMPLO",
                y="PERSONAS",
                color="ROL",
                barmode="group",
                title="Cobertura de testigos y beneficiarios por templo",
            )
            fig_testigos.update_layout(height=430, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT)
            st.plotly_chart(fig_testigos, width="stretch")
        with tc2:
            fig_ratio = px.bar(
                testigos_analisis.sort_values("REFERIDOS_REGISTRADOS", ascending=True),
                x="REFERIDOS_REGISTRADOS",
                y="TEMPLO",
                orientation="h",
                title="Referidos registrados por templo",
                color="TEMPLO",
                color_discrete_map=COLORES_TEMPLOS,
                category_orders=PLOTLY_TEMPLO_ORDERS,
            )
            fig_ratio.update_layout(height=430, paper_bgcolor="white", plot_bgcolor="white", font_color=COLOR_TEXT, showlegend=False)
            st.plotly_chart(fig_ratio, width="stretch")

        testigos_tabla = pd.concat([pd.DataFrame([total_row]), testigos_analisis], ignore_index=True)
        testigos_show = testigos_tabla.rename(
            columns={
                "TOTAL_TESTIGOS": "TOTAL TESTIGOS",
                "TESTIGOS_MESA_O_REMANENTE": "TESTIGOS MESA / REMANENTE",
                "TESTIGOS_COMISION_ESCRUTADORA": "COMISION ESCRUTADORA",
                "BENEFICIARIOS_MESAS_TRABAJO": "BENEFICIARIOS MESAS",
                "CANTIDAD_MESAS_TRABAJO_ASOCIADAS": "MESAS ASOCIADAS",
                "TESTIGOS_DOBLE_ROL": "DOBLE ROL",
                "LIDERES": "LIDERES",
                "NO_LIDERES": "NO LIDERES",
                "TESTIGOS_CON_REFERIDOS": "TESTIGOS CON REFERIDOS",
                "REFERIDOS_REGISTRADOS": "REFERIDOS REGISTRADOS",
                "REFERIDOS_INACTIVOS": "REFERIDOS INACTIVOS",
                "REFERIDOS_ACTIVOS_ESTIMADOS": "REFERIDOS ACTIVOS ESTIMADOS",
                "PUESTOS": "PUESTOS VISIBLES",
                "VOTOS_2026": "VOTOS 2026",
                "TESTIGOS_POR_PUESTO": "TESTIGOS POR PUESTO",
                "VOTOS_2026_POR_TESTIGO": "VOTOS 2026 POR TESTIGO",
                "% DEL TOTAL KENNEDY": "% DEL TOTAL KENNEDY",
                "% MESA / REMANENTE": "% MESA / REMANENTE",
                "% COMISION": "% COMISION",
                "% BENEFICIARIOS MESAS": "% BENEFICIARIOS MESAS",
                "% LIDERES": "% LIDERES",
                "% CON REFERIDOS": "% CON REFERIDOS",
                "REFERIDOS POR TESTIGO": "REFERIDOS POR TESTIGO",
                "REFERIDOS POR LIDER": "REFERIDOS POR LIDER",
            }
        )
        cols_testigos_show = [
            "TEMPLO",
            "TOTAL TESTIGOS",
            "% DEL TOTAL KENNEDY",
            "LIDERES",
            "% LIDERES",
            "NO LIDERES",
            "TESTIGOS CON REFERIDOS",
            "% CON REFERIDOS",
            "REFERIDOS REGISTRADOS",
            "REFERIDOS ACTIVOS ESTIMADOS",
            "REFERIDOS INACTIVOS",
            "REFERIDOS POR TESTIGO",
            "REFERIDOS POR LIDER",
            "TESTIGOS MESA / REMANENTE",
            "% MESA / REMANENTE",
            "COMISION ESCRUTADORA",
            "% COMISION",
            "BENEFICIARIOS MESAS",
            "% BENEFICIARIOS MESAS",
            "MESAS ASOCIADAS",
            "DOBLE ROL",
            "PUESTOS VISIBLES",
            "VOTOS 2026",
            "TESTIGOS POR PUESTO",
            "VOTOS 2026 POR TESTIGO",
        ]
        cols_testigos_show = [c for c in cols_testigos_show if c in testigos_show.columns]
        testigos_show_fmt = testigos_show[cols_testigos_show].copy()
        for col in ["% DEL TOTAL KENNEDY", "% LIDERES", "% CON REFERIDOS", "% MESA / REMANENTE", "% COMISION", "% BENEFICIARIOS MESAS"]:
            if col in testigos_show_fmt.columns:
                testigos_show_fmt[col] = testigos_show_fmt[col].map(fmt_pct)
        for col in ["REFERIDOS POR TESTIGO", "REFERIDOS POR LIDER", "TESTIGOS POR PUESTO", "VOTOS 2026 POR TESTIGO"]:
            if col in testigos_show_fmt.columns:
                testigos_show_fmt[col] = pd.to_numeric(testigos_show_fmt[col], errors="coerce").map(lambda v: fmt_number(v, 1))
        st.dataframe(testigos_show_fmt, hide_index=True, width="stretch")
        st.download_button(
            "Descargar análisis de testigos por templo",
            to_excel_bytes(testigos_show_fmt, "Testigos por templo"),
            "analisis_testigos_por_templo.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_testigos_templo",
        )

    col1, col2 = st.columns(2)
    with col1:
        top_growth = puestos_f.sort_values("VARIACION_ABSOLUTA", ascending=False).head(15)
        fig = px.bar(
            top_growth.sort_values("VARIACION_ABSOLUTA"),
            x="VARIACION_ABSOLUTA",
            y="PUESTO",
            orientation="h",
            color="IGLESIA",
            color_discrete_map=COLORES_TEMPLOS,
            category_orders=PLOTLY_TEMPLO_ORDERS,
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
            color_discrete_map=COLORES_TEMPLOS,
            category_orders=PLOTLY_TEMPLO_ORDERS,
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
        "TEMPLO_REPORTE", "TIENE_MESA_TRABAJO",
    ]
    cols_puesto = [c for c in cols_puesto if c in puestos_f.columns]
    df_puesto_show = puestos_f[cols_puesto].sort_values("VOTOS_2026", ascending=False)
    st.dataframe(df_puesto_show, width="stretch", hide_index=True)
    st.download_button("Descargar matriz por puesto en Excel", to_excel_bytes(df_puesto_show, "Matriz Puesto"), "matriz_puesto.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_mat_puesto")

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
        "puestos_votacion.csv": ocultar_columnas_priorizacion(puestos),
        "resumen_iglesia.csv": resumen_iglesia,
        "actividades_campana.csv": actividades,
        "mesas_trabajo.csv": mesas,
        "control_calidad.csv": control,
        "historial_ajustes_templo.csv": historial_ajustes,
    }
    if not apoyos_ciudadanos.empty:
        exportables["apoyos_ciudadanos_sin_gestor.csv"] = apoyos_ciudadanos

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
