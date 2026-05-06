
"""
consolidar_datos.py

Genera la base maestra `data/kennedy_mira_consolidado.xlsx` a partir de los tres
archivos originales de campaña, gestión y votación. El dashboard app.py lee
únicamente este Excel consolidado.

Ejecutar:
    python consolidar_datos.py
"""

from pathlib import Path
from difflib import SequenceMatcher
import re
import unicodedata

import numpy as np
import pandas as pd


DATA_DIR = Path("data")
ARCHIVO_CAMPANA = DATA_DIR / "CAMPAÑA CONGRESO 2026 KENNEDY (1).xlsx"
ARCHIVO_GESTION = DATA_DIR / "Copia de Gestión Edil Lorena Garzón - 17 de febrero, 17_07.xlsx"
ARCHIVO_VOTACION = DATA_DIR / "VOTACIÓN 2026.xlsx"
ARCHIVO_PUESTOS_LOCALIDAD = DATA_DIR / "Puestos Localidad de Kennedy 2026.xlsx"
SALIDA = DATA_DIR / "kennedy_mira_consolidado.xlsx"


def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))


def clean_text(s):
    if pd.isna(s):
        return ""
    s = strip_accents(str(s)).upper().strip()
    return re.sub(r"\s+", " ", s)


def key_text(s):
    s = clean_text(s).replace("Ñ", "N")
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    s = f" {s} "
    for a, b in {
        " I D E ": " IED ",
        " DIST ": " DISTRITAL ",
        " COL ": " COLEGIO ",
        " DIS ": " DISTRITAL ",
        " INTSMAR ": " INSTMAR ",
    }.items():
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s).strip()


def key_puesto_simple(s):
    s = key_text(s)
    s = re.sub(r"\b(COLEGIO|DISTRITAL|COL|DIS|IED|SEDE)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def normalizar_iglesia(v):
    s = clean_text(v)
    if not s or s == "NAN":
        return "SIN CLASIFICAR"
    if any(x in s for x in ["CLASS", "CLAS", "ROMA", "KENNEDY CLASS"]):
        return "CLASS ROMA"
    if "PATIO" in s or "BELLAVISTA" in s:
        return "PATIO BONITO"
    if "CARVAJAL" in s or "EDEN" in s:
        return "CARVAJAL"
    if "VALLADOLID" in s:
        return "VALLADOLID"
    if "KENNEDY" in s or "CENTRAL" in s or "CASABLANCA" in s:
        return "KENNEDY CENTRAL"
    return "SIN CLASIFICAR"


def parse_coord(v):
    if pd.isna(v):
        return np.nan, np.nan
    vals = []
    for n in re.findall(r"-?\d+(?:[.,]\d+)?", str(v)):
        try:
            vals.append(float(n.replace(",", ".")))
        except Exception:
            pass
    for i in range(len(vals) - 1):
        lat, lon = vals[i], vals[i + 1]
        if 3.5 <= lat <= 5.5 and -75 <= lon <= -73:
            return lat, lon
    return np.nan, np.nan


def best_match(key, candidates):
    best, score = None, 0
    for c in candidates:
        s = SequenceMatcher(None, key, c).ratio()
        if s > score:
            best, score = c, s
    return best, score


def normalize_cols(df):
    out = df.copy()
    out.columns = [clean_text(c) for c in out.columns]
    return out


def cargar_puestos_localidad():
    if not ARCHIVO_PUESTOS_LOCALIDAD.exists():
        return pd.DataFrame()
    df = normalize_cols(pd.read_excel(ARCHIVO_PUESTOS_LOCALIDAD, sheet_name=0))
    if "PUESTO 2025" not in df.columns:
        return pd.DataFrame()
    df = df[df["PUESTO 2025"].notna()].copy()
    df["PUESTO_KEY"] = df["PUESTO 2025"].map(key_text)
    df["PUESTO_KEY_SIMPLE"] = df["PUESTO 2025"].map(key_puesto_simple)
    return df


def enriquecer_con_puestos_localidad(puestos, detalle):
    if detalle.empty:
        for col in [
            "CODIGO_PUESTO_2026",
            "LLAVE_PUESTO_2026",
            "DIRECCION_2026_REPORTE",
            "MESAS_2026_REPORTE",
            "TESTIGOS_2023_REPORTE",
            "VOTOS_MIRA_2023_PROP_LISTA",
            "VOTOS_AFINIDAD_E11_2023",
            "MENOS_DE_1_KM",
            "TEMPLO_REPORTE",
            "IGLESIA_REPORTE_2026",
            "MATCH_PUESTOS_LOCALIDAD_2026",
        ]:
            puestos[col] = np.nan
        return puestos

    by_key = {r["PUESTO_KEY"]: r for _, r in detalle.iterrows()}
    by_simple = {r["PUESTO_KEY_SIMPLE"]: r for _, r in detalle.iterrows() if r["PUESTO_KEY_SIMPLE"]}
    keys = list(by_key.keys())
    simple_keys = list(by_simple.keys())

    enriched = []
    for _, row in puestos.iterrows():
        key = key_text(row["PUESTO"])
        simple = key_puesto_simple(row["PUESTO"])
        match, score, metodo = None, 0, "SIN MATCH"
        if key in by_key:
            match, score, metodo = by_key[key], 1, "EXACTO"
        elif simple in by_simple:
            match, score, metodo = by_simple[simple], 1, "EXACTO_SIMPLE"
        else:
            contained = [k for k in keys if len(k) >= 18 and (k in key or key in k)]
            if contained:
                b = max(contained, key=len)
                match, score, metodo = by_key[b], 0.99, "CONTENIDO"
            else:
                b, s = best_match(key, keys)
                if s >= 0.82:
                    match, score, metodo = by_key[b], s, "FUZZY"
                else:
                    contained_simple = [k for k in simple_keys if len(k) >= 12 and (k in simple or simple in k)]
                    if contained_simple:
                        b = max(contained_simple, key=len)
                        match, score, metodo = by_simple[b], 0.99, "CONTENIDO_SIMPLE"
                    else:
                        b, s = best_match(simple, simple_keys)
                        if s >= 0.86:
                            match, score, metodo = by_simple[b], s, "FUZZY_SIMPLE"

        record = row.to_dict()
        if match is not None:
            record.update({
                "CODIGO_PUESTO_2026": match.get("CODIGO DE PUESTO", ""),
                "LLAVE_PUESTO_2026": match.get("LLAVE", ""),
                "DIRECCION_2026_REPORTE": match.get("DIRECCIONES 2026", ""),
                "MESAS_2026_REPORTE": pd.to_numeric(match.get("MESAS 2026"), errors="coerce"),
                "TESTIGOS_2023_REPORTE": pd.to_numeric(match.get("TESTIGOS 2023"), errors="coerce"),
                "VOTOS_MIRA_2023_PROP_LISTA": pd.to_numeric(match.get("VOTOS MIRA 2023 + PROPORCION A LISTA"), errors="coerce"),
                "VOTOS_AFINIDAD_E11_2023": pd.to_numeric(match.get("SUMA DE VOTOS_AFINIDAD E-11 2023"), errors="coerce"),
                "MENOS_DE_1_KM": "SI" if clean_text(match.get("MENOS DE 1 KILOMETRO")) == "X" else "NO",
                "TEMPLO_REPORTE": clean_text(match.get("TEMPLO")),
                "IGLESIA_REPORTE_2026": normalizar_iglesia(match.get("TEMPLO")),
                "MATCH_PUESTOS_LOCALIDAD_2026": f"{metodo}:{score:.3f}",
            })
            if not record.get("DIRECCION") and match.get("DIRECCIONES 2026"):
                record["DIRECCION"] = match.get("DIRECCIONES 2026")
        else:
            record.update({
                "CODIGO_PUESTO_2026": "",
                "LLAVE_PUESTO_2026": "",
                "DIRECCION_2026_REPORTE": "",
                "MESAS_2026_REPORTE": np.nan,
                "TESTIGOS_2023_REPORTE": np.nan,
                "VOTOS_MIRA_2023_PROP_LISTA": np.nan,
                "VOTOS_AFINIDAD_E11_2023": np.nan,
                "MENOS_DE_1_KM": "",
                "TEMPLO_REPORTE": "",
                "IGLESIA_REPORTE_2026": "SIN CLASIFICAR",
                "MATCH_PUESTOS_LOCALIDAD_2026": "SIN MATCH",
            })
        enriched.append(record)
    return pd.DataFrame(enriched)


def cargar_votacion():
    v5 = pd.read_excel(ARCHIVO_VOTACION, sheet_name="Hoja 5")
    v3 = pd.read_excel(ARCHIVO_VOTACION, sheet_name="Hoja 3")
    v7 = pd.read_excel(ARCHIVO_VOTACION, sheet_name="Hoja 7")

    puestos = v5.iloc[1:].copy()
    puestos = puestos[puestos["PUESTO DE VOTACIÓN"].notna()].copy()
    puestos["PUESTO_KEY"] = puestos["PUESTO DE VOTACIÓN"].map(key_text)

    v3 = v3[v3["PUESTO"].notna()].copy()
    v3 = v3[~v3["PUESTO"].map(key_text).eq("TOTAL")].copy()
    v3["PUESTO_KEY_SRC"] = v3["PUESTO"].map(key_text)
    h3_map = {r["PUESTO_KEY_SRC"]: r for _, r in v3.iterrows()}
    h3_keys = list(h3_map.keys())

    v7 = v7[v7["Puestos 2026"].notna()].copy()
    v7["PUESTO_KEY_SRC"] = v7["Puestos 2026"].map(key_text)
    h7_map = {r["PUESTO_KEY_SRC"]: r for _, r in v7.iterrows()}
    h7_keys = list(h7_map.keys())

    rows = []
    for _, r in puestos.iterrows():
        pk = r["PUESTO_KEY"]

        h3row, h7row, h3score, h7score = None, None, 0, 0

        if pk in h3_map:
            h3row, h3score = h3_map[pk], 1
        else:
            b, s = best_match(pk, h3_keys)
            if s >= 0.82:
                h3row, h3score = h3_map[b], s

        if pk in h7_map:
            h7row, h7score = h7_map[pk], 1
        else:
            b, s = best_match(pk, h7_keys)
            if s >= 0.82:
                h7row, h7score = h7_map[b], s

        church_raw = h7row.get("IGLESIA") if h7row is not None else None
        if normalizar_iglesia(church_raw) == "SIN CLASIFICAR" and h3row is not None:
            church_raw = h3row.get("IGLESIA RESPONSABLE")
        if normalizar_iglesia(church_raw) == "SIN CLASIFICAR":
            church_raw = r.get("IGLESIA RESPONSABLE")

        iglesia = normalizar_iglesia(church_raw)
        puesto_nombre = clean_text(r["PUESTO DE VOTACIÓN"])

        # Correcciones manuales validadas contra la matriz de puestos.
        if puesto_nombre == "ALMENAR":
            iglesia = "KENNEDY CENTRAL"
            h3row = {"DIRECCIÓN": "SALON COMUNAL ALMENAR", "COORDENADAS": "4.619939810553133, -74.17035065790459"}
        elif puesto_nombre == "CODEMA":
            iglesia = "PATIO BONITO"
            h3row = {"DIRECCIÓN": "COL. DIST. CODEMA", "COORDENADAS": "4.648603178621415, -74.16533739303424"}
        elif "CARVAJAL II SECTOR" in puesto_nombre:
            iglesia = "CARVAJAL"
            h3row = {"DIRECCIÓN": "CRA. 70B # 24A-40 SUR", "COORDENADAS": "4.61618790095059, -74.13594710434316"}

        direccion = h3row.get("DIRECCIÓN", "") if h3row is not None else ""
        coord = h3row.get("COORDENADAS", "") if h3row is not None else ""
        lat, lon = parse_coord(coord)

        votos2026 = pd.to_numeric(r.get("PROMEDIO 2026"), errors="coerce")
        votos2023 = pd.to_numeric(r.get("PROMEDIO 2023"), errors="coerce")
        var_abs = votos2026 - votos2023 if pd.notna(votos2026) and pd.notna(votos2023) else np.nan
        var_pct = var_abs / votos2023 if pd.notna(var_abs) and pd.notna(votos2023) and votos2023 else np.nan
        mesa = str(r.get("SE HA HECHO MESA DE TRABAJO", "")).strip().upper()

        rows.append({
            "PUESTO_ID": len(rows) + 1,
            "PUESTO": puesto_nombre,
            "DIRECCION": direccion,
            "BARRIO": "",
            "UPZ": "",
            "IGLESIA": iglesia,
            "LATITUD": lat,
            "LONGITUD": lon,
            "CENSO_2023": pd.to_numeric(r.get("CENSO 2023"), errors="coerce"),
            "MIRA_CONCEJO_2023": pd.to_numeric(r.get("MIRA CONCEJO 2023"), errors="coerce"),
            "JAL_2023": pd.to_numeric(r.get("JAL 2023"), errors="coerce"),
            "VOTOS_2023": votos2023,
            "CAMARA_2026": pd.to_numeric(r.get("CAMARA 2026"), errors="coerce"),
            "SENADO_2026": pd.to_numeric(r.get("SENADO 2026"), errors="coerce"),
            "VOTOS_2026": votos2026,
            "VARIACION_ABSOLUTA": var_abs,
            "VARIACION_PORCENTUAL": var_pct,
            "RESULTADO_VARIACION": r.get("Resultado variación promedio ", ""),
            "TIENE_MESA_TRABAJO": mesa,
            "CUALES_MESAS": r.get("CUÁL", ""),
            "BENEFICIARIOS": pd.to_numeric(r.get("BENEFICIARIOS"), errors="coerce"),
            "ACTIVIDADES_CAMPANA": 0,
            "MESAS_TRABAJO_BARRIO": 1 if mesa == "SI" else 0,
            "MATCH_PUESTO_COORDENADAS": round(h3score, 3),
            "MATCH_IGLESIA": round(h7score, 3),
        })

    puestos = pd.DataFrame(rows)
    puestos = enriquecer_con_puestos_localidad(puestos, cargar_puestos_localidad())
    return puestos, v5


def cargar_actividades():
    records = []
    for sh in ["AGENDA GENERAL CON CANDIDATOS", "AGENDA PARALELA", "Cronograma Kennedy enero"]:
        df = normalize_cols(pd.read_excel(ARCHIVO_CAMPANA, sheet_name=sh))
        for _, r in df.iterrows():
            if pd.isna(r.get("SEDE")) and pd.isna(r.get("ACTIVIDAD")):
                continue
            lat, lon = parse_coord(r.get("COORDENADAS", ""))
            records.append({
                "ACTIVIDAD_ID": len(records) + 1,
                "FUENTE": sh,
                "FECHA": pd.to_datetime(r.get("FECHA CAMPAÑA"), errors="coerce"),
                "IGLESIA": normalizar_iglesia(r.get("SEDE")),
                "BARRIO": clean_text(r.get("BARRIO")),
                "DIRECCION": r.get("DIRECCION", ""),
                "LATITUD": lat,
                "LONGITUD": lon,
                "TIPO_ACTIVIDAD": clean_text(r.get("ACTIVIDAD")),
                "LIDER": r.get("LIDER Y CELULAR", ""),
                "OBSERVACIONES": r.get("DETALLE DE LA ACTIVIDAD", ""),
            })
    return pd.DataFrame(records)


def cargar_mesas():
    records = []
    for sh in ["Mesas", "Hoja 27"]:
        try:
            df = normalize_cols(pd.read_excel(ARCHIVO_CAMPANA, sheet_name=sh))
        except Exception:
            continue
        for _, r in df.iterrows():
            if not r.dropna().shape[0]:
                continue
            lat, lon = parse_coord(r.get("UBICACION", ""))
            records.append({
                "MESA_ID": len(records) + 1,
                "FUENTE": f"campana:{sh}",
                "FECHA": pd.to_datetime(r.get("FECHA"), errors="coerce"),
                "IGLESIA": normalizar_iglesia(r.get("IGLESIA")),
                "BARRIO": clean_text(r.get("BARRIO")),
                "DIRECCION": r.get("UBICACION", ""),
                "LATITUD": lat,
                "LONGITUD": lon,
                "TEMA": r.get("TEMAS", ""),
                "ENTIDADES": "",
                "LIDER": r.get("SOLICITANTE", ""),
                "ESTADO": r.get("ESTADO", ""),
                "OBSERVACIONES": r.get("ACCIONES", ""),
            })

    for sh in ["SEGUIMIENTO MESAS DE TRABAJO", "ESTADO MESAS DE TRABAJO CORTE 2"]:
        try:
            df = normalize_cols(pd.read_excel(ARCHIVO_GESTION, sheet_name=sh))
        except Exception:
            continue
        for _, r in df.iterrows():
            records.append({
                "MESA_ID": len(records) + 1,
                "FUENTE": f"gestion:{sh}",
                "FECHA": pd.NaT,
                "IGLESIA": normalizar_iglesia(r.get("IGLESIA")),
                "BARRIO": clean_text(r.get("BARRIO")),
                "DIRECCION": "",
                "LATITUD": np.nan,
                "LONGITUD": np.nan,
                "TEMA": r.get("TEMA", ""),
                "ENTIDADES": r.get("ENTIDAD", ""),
                "LIDER": r.get("RESPONSABLE", "") or r.get("APOYO ASESOR", ""),
                "ESTADO": r.get("ESTADO DE SEGUIMIENTO", ""),
                "OBSERVACIONES": r.get("ACCION", ""),
            })

    mesas = pd.DataFrame(records)
    if not mesas.empty:
        mesas["DEDUP_KEY"] = mesas["BARRIO"].fillna("") + "|" + mesas["IGLESIA"].fillna("") + "|" + mesas["TEMA"].fillna("").astype(str).str[:50]
        mesas = mesas.drop_duplicates("DEDUP_KEY").drop(columns=["DEDUP_KEY"]).reset_index(drop=True)
        mesas["MESA_ID"] = range(1, len(mesas) + 1)
    return mesas


def classify_priority(row):
    v26 = row["VOTOS_2026"] if pd.notna(row["VOTOS_2026"]) else 0
    var = row["VARIACION_ABSOLUTA"] if pd.notna(row["VARIACION_ABSOLUTA"]) else 0
    mesas = row["MESAS_TRABAJO_BARRIO"] if pd.notna(row["MESAS_TRABAJO_BARRIO"]) else 0

    if v26 >= 100 and var < 0:
        return "ALTA", "Caída en puesto de alta votación", "Recuperar votación perdida con agenda territorial focalizada."
    if v26 >= 100 and mesas == 0:
        return "ALTA", "Alta votación sin mesa de trabajo registrada", "Programar mesa de trabajo o visita de seguimiento en el barrio."
    if var < -15:
        return "ALTA", "Caída electoral relevante", "Identificar líderes, causas de pérdida y plan de recuperación."
    if v26 >= 50 and var >= 0:
        return "MEDIA", "Crecimiento o base electoral media", "Consolidar mediante contacto con líderes y presencia comunitaria."
    if v26 >= 30:
        return "MEDIA", "Votación media con oportunidad", "Mantener seguimiento y evaluar actividad de bajo costo."
    return "BAJA", "Baja votación o datos insuficientes", "Monitorear; no priorizar en primera fase."


def main():
    DATA_DIR.mkdir(exist_ok=True)

    required = [ARCHIVO_CAMPANA, ARCHIVO_GESTION, ARCHIVO_VOTACION]
    missing = [str(p) for p in required if not p.exists()]
    if missing:
        raise FileNotFoundError("Faltan archivos originales en data/: " + ", ".join(missing))

    puestos, v5 = cargar_votacion()
    actividades = cargar_actividades()
    mesas = cargar_mesas()

    iglesias = pd.DataFrame([
        {"IGLESIA": "CLASS ROMA", "LATITUD": 4.614359775316158, "LONGITUD": -74.17619195767098, "URL": "https://direcciones.idmji.org/es/iglesia/359/"},
        {"IGLESIA": "PATIO BONITO", "LATITUD": 4.646035122997863, "LONGITUD": -74.17300841534194, "URL": "https://direcciones.idmji.org/es/iglesia/301/"},
        {"IGLESIA": "KENNEDY CENTRAL", "LATITUD": 4.6217386978458155, "LONGITUD": -74.16501499477366, "URL": ""},
        {"IGLESIA": "CARVAJAL", "LATITUD": 4.616343469904612, "LONGITUD": -74.1404329155982, "URL": ""},
        {"IGLESIA": "VALLADOLID", "LATITUD": 4.647817860855581, "LONGITUD": -74.14806885512174, "URL": ""},
    ])
    iglesias["TIPO"] = "Iglesia / templo"

    act_counts = actividades[actividades["IGLESIA"].isin(iglesias["IGLESIA"])].groupby("IGLESIA").size()
    mesa_counts = mesas[mesas["IGLESIA"].isin(iglesias["IGLESIA"])].groupby("IGLESIA").size()
    puestos["ACTIVIDADES_CAMPANA_IGLESIA"] = puestos["IGLESIA"].map(act_counts).fillna(0).astype(int)
    puestos["MESAS_TRABAJO_IGLESIA"] = puestos["IGLESIA"].map(mesa_counts).fillna(0).astype(int)

    pri = puestos.apply(classify_priority, axis=1)
    puestos["PRIORIDAD"] = [x[0] for x in pri]
    puestos["RAZON_PRIORIDAD"] = [x[1] for x in pri]
    puestos["ACCION_RECOMENDADA"] = [x[2] for x in pri]

    resumen_iglesia = []
    for iglesia in iglesias["IGLESIA"]:
        sub = puestos[puestos["IGLESIA"] == iglesia]
        v26, v23 = sub["VOTOS_2026"].sum(), sub["VOTOS_2023"].sum()
        va = v26 - v23
        vp = va / v23 if v23 else np.nan
        if sub.empty:
            maxv = maxdrop = maxgrowth = "SIN PUESTOS ASIGNADOS"
            lectura = "Iglesia oficial sin puestos asignados en la matriz electoral."
            recomendacion = "Definir si debe asumir puestos cercanos o mantener rol de apoyo territorial."
        else:
            maxv = sub.loc[sub["VOTOS_2026"].idxmax(), "PUESTO"]
            maxdrop = sub.loc[sub["VARIACION_ABSOLUTA"].idxmin(), "PUESTO"]
            maxgrowth = sub.loc[sub["VARIACION_ABSOLUTA"].idxmax(), "PUESTO"]
            lectura = f"{iglesia} concentra {v26:,.1f} votos promedio 2026 en {sub['PUESTO'].nunique()} puestos."
            recomendacion = "Consolidar puestos de alto rendimiento." if va >= 0 else "Priorizar puestos con mayor caída."
        resumen_iglesia.append({
            "IGLESIA": iglesia, "VOTOS_2026": v26, "VOTOS_2023": v23, "VARIACION_ABSOLUTA": va,
            "VARIACION_PORCENTUAL": vp, "PUESTOS": sub["PUESTO"].nunique(), "BARRIOS": sub["BARRIO"].replace("", np.nan).nunique(),
            "ACTIVIDADES_CAMPANA": int(act_counts.get(iglesia, 0)), "MESAS_TRABAJO": int(mesa_counts.get(iglesia, 0)),
            "PUESTO_MAYOR_VOTACION": maxv, "PUESTO_MAYOR_CAIDA": maxdrop, "PUESTO_MAYOR_CRECIMIENTO": maxgrowth,
            "LECTURA_ESTRATEGICA": lectura, "RECOMENDACION": recomendacion,
        })
    resumen_iglesia = pd.DataFrame(resumen_iglesia)

    resumen_puesto = puestos[["PUESTO", "IGLESIA", "BARRIO", "UPZ", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "ACTIVIDADES_CAMPANA_IGLESIA", "MESAS_TRABAJO_BARRIO", "PRIORIDAD", "ACCION_RECOMENDADA"]].copy()
    resumen_puesto = resumen_puesto.rename(columns={"ACTIVIDADES_CAMPANA_IGLESIA": "ACTIVIDADES_CAMPANA"})

    resumen_barrio = puestos.groupby(["BARRIO", "IGLESIA", "UPZ"], dropna=False).agg(
        VOTOS_2026=("VOTOS_2026", "sum"), VOTOS_2023=("VOTOS_2023", "sum"),
        PUESTOS=("PUESTO", "nunique"), ACTIVIDADES_CAMPANA=("ACTIVIDADES_CAMPANA_IGLESIA", "max"),
        MESAS_TRABAJO=("MESAS_TRABAJO_BARRIO", "sum")
    ).reset_index()
    resumen_barrio["VARIACION_ABSOLUTA"] = resumen_barrio["VOTOS_2026"] - resumen_barrio["VOTOS_2023"]
    resumen_barrio["VARIACION_PORCENTUAL"] = np.where(resumen_barrio["VOTOS_2023"] > 0, resumen_barrio["VARIACION_ABSOLUTA"] / resumen_barrio["VOTOS_2023"], np.nan)
    resumen_barrio["PRIORIDAD"] = np.where((resumen_barrio["VOTOS_2026"] >= 100) & (resumen_barrio["VARIACION_ABSOLUTA"] < 0), "ALTA", np.where(resumen_barrio["VOTOS_2026"] >= 50, "MEDIA", "BAJA"))
    resumen_barrio["ACCION_RECOMENDADA"] = np.where(resumen_barrio["PRIORIDAD"] == "ALTA", "Plan de recuperación barrial y mesa de seguimiento.", np.where(resumen_barrio["PRIORIDAD"] == "MEDIA", "Consolidar presencia territorial.", "Monitoreo."))

    matriz = puestos[["PUESTO", "IGLESIA", "BARRIO", "UPZ", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "PRIORIDAD", "RAZON_PRIORIDAD", "ACCION_RECOMENDADA"]].copy()
    matriz["NIVEL_PRIORIDAD"] = matriz["PRIORIDAD"]
    matriz["VARIABLE_CRITICA"] = matriz["RAZON_PRIORIDAD"]
    matriz["DIAGNOSTICO"] = np.where(matriz["VARIACION_ABSOLUTA"] < 0, "El puesto pierde votación frente a 2023.", "El puesto mantiene o aumenta votación.")
    matriz["TEMPORALIDAD"] = np.where(matriz["NIVEL_PRIORIDAD"] == "ALTA", "0-60 días", np.where(matriz["NIVEL_PRIORIDAD"] == "MEDIA", "60-120 días", "Seguimiento trimestral"))
    matriz["RESPONSABLE_SUGERIDO"] = matriz["IGLESIA"]
    matriz = matriz[["NIVEL_PRIORIDAD", "PUESTO", "IGLESIA", "BARRIO", "UPZ", "VOTOS_2026", "VOTOS_2023", "VARIACION_ABSOLUTA", "VARIACION_PORCENTUAL", "VARIABLE_CRITICA", "DIAGNOSTICO", "ACCION_RECOMENDADA", "TEMPORALIDAD", "RESPONSABLE_SUGERIDO"]]

    total = v5.iloc[0]
    total_2026, total_2023 = float(total["PROMEDIO 2026"]), float(total["PROMEDIO 2023"])
    total_jal_2023 = pd.to_numeric(puestos["JAL_2023"], errors="coerce").sum()
    total_concejo_2023 = pd.to_numeric(puestos["MIRA_CONCEJO_2023"], errors="coerce").sum()
    total_camara_2026 = pd.to_numeric(puestos["CAMARA_2026"], errors="coerce").sum()
    total_senado_2026 = pd.to_numeric(puestos["SENADO_2026"], errors="coerce").sum()
    mesas_2026_reporte = pd.to_numeric(puestos.get("MESAS_2026_REPORTE", pd.Series(dtype=float)), errors="coerce").sum()
    testigos_2023_reporte = pd.to_numeric(puestos.get("TESTIGOS_2023_REPORTE", pd.Series(dtype=float)), errors="coerce").sum()
    afinidad_2023_reporte = pd.to_numeric(puestos.get("VOTOS_AFINIDAD_E11_2023", pd.Series(dtype=float)), errors="coerce").sum()
    mira_prop_2023_reporte = pd.to_numeric(puestos.get("VOTOS_MIRA_2023_PROP_LISTA", pd.Series(dtype=float)), errors="coerce").sum()
    puestos_reporte_match = puestos.get("MATCH_PUESTOS_LOCALIDAD_2026", pd.Series("", index=puestos.index)).ne("SIN MATCH").sum()
    resumen_general = pd.DataFrame([
        {"INDICADOR": "Total Kennedy votos promedio 2026", "VALOR": total_2026, "NOTA": "Fila total oficial de Hoja 5."},
        {"INDICADOR": "Total Kennedy votos promedio 2023", "VALOR": total_2023, "NOTA": "Fila total oficial de Hoja 5."},
        {"INDICADOR": "Variación absoluta Kennedy", "VALOR": total_2026 - total_2023, "NOTA": "2026 - 2023."},
        {"INDICADOR": "Variación porcentual Kennedy", "VALOR": (total_2026 - total_2023) / total_2023, "NOTA": "Variación / 2023."},
        {"INDICADOR": "JAL 2023 Kennedy", "VALOR": total_jal_2023, "NOTA": "Suma por puestos de JAL 2023."},
        {"INDICADOR": "Concejo 2023 Kennedy", "VALOR": total_concejo_2023, "NOTA": "Suma por puestos de MIRA Concejo 2023."},
        {"INDICADOR": "Cámara 2026 Kennedy", "VALOR": total_camara_2026, "NOTA": "Suma por puestos de Cámara 2026."},
        {"INDICADOR": "Senado 2026 Kennedy", "VALOR": total_senado_2026, "NOTA": "Suma por puestos de Senado 2026."},
        {"INDICADOR": "Puestos totales analizados", "VALOR": len(puestos), "NOTA": "Excluye fila total KENNEDY."},
        {"INDICADOR": "Puestos con iglesia oficial asignada", "VALOR": int(puestos["IGLESIA"].isin(iglesias["IGLESIA"]).sum()), "NOTA": "Todos los puestos quedaron asignados a iglesias oficiales."},
        {"INDICADOR": "Iglesias oficiales", "VALOR": 5, "NOTA": "Class Roma, Kennedy Central, Patio Bonito, Carvajal y Valladolid."},
        {"INDICADOR": "Actividades de campaña consolidadas", "VALOR": len(actividades), "NOTA": "Agenda general, paralela y cronograma."},
        {"INDICADOR": "Mesas de trabajo consolidadas", "VALOR": len(mesas), "NOTA": "Bases de campaña y gestión."},
        {"INDICADOR": "Puestos cruzados con reporte localidad 2026", "VALOR": puestos_reporte_match, "NOTA": "Cruce contra Puestos Localidad de Kennedy 2026."},
        {"INDICADOR": "Mesas 2026 reporte localidad", "VALOR": mesas_2026_reporte, "NOTA": "Suma de Mesas 2026 del reporte de puestos."},
        {"INDICADOR": "Testigos 2023 reporte localidad", "VALOR": testigos_2023_reporte, "NOTA": "Suma de Testigos 2023 del reporte de puestos."},
        {"INDICADOR": "Votos MIRA 2023 proporción lista reporte", "VALOR": mira_prop_2023_reporte, "NOTA": "Suma del campo Votos MIRA 2023 + Proporción a lista."},
        {"INDICADOR": "Votos afinidad E-11 2023 reporte", "VALOR": afinidad_2023_reporte, "NOTA": "Suma del campo votos_afinidad E-11 2023."},
    ])

    informe = pd.DataFrame([
        {"SECCION": "Resumen general", "TEXTO": f"Kennedy registró {total_2026:,.1f} votos promedio en 2026 frente a {total_2023:,.1f} en 2023, para una variación de {total_2026-total_2023:,.1f} votos ({(total_2026-total_2023)/total_2023:.2%})."},
        {"SECCION": "Hallazgos principales", "TEXTO": "El tablero permite separar resultado general Kennedy y rendimiento por iglesia, evitando confundir totales filtrados con totales oficiales."},
        {"SECCION": "Lectura JAL y Concejo 2023", "TEXTO": f"La base 2023 separa JAL ({total_jal_2023:,.1f}) y Concejo ({total_concejo_2023:,.1f}), lo que permite distinguir comportamiento local y voto de corporación distrital."},
        {"SECCION": "Lectura Cámara y Senado 2026", "TEXTO": f"La base 2026 separa Cámara ({total_camara_2026:,.1f}) y Senado ({total_senado_2026:,.1f}), permitiendo evaluar rendimiento legislativo por puesto e iglesia."},
        {"SECCION": "Reporte de puestos localidad 2026", "TEXTO": f"Se cruzaron {puestos_reporte_match:,.0f} puestos con el archivo Puestos Localidad de Kennedy 2026, incorporando mesas 2026, testigos 2023, afinidad E-11 y templo reportado como variables operativas complementarias."},
        {"SECCION": "Recomendaciones estratégicas", "TEXTO": "Priorizar puestos de alta votación con caída, consolidar iglesias con saldo positivo y completar capa UPZ para análisis espacial."},
    ])

    control = pd.DataFrame([
        {"TIPO": "IGLESIA", "REGISTRO": "VALLADOLID", "OBSERVACION": "Iglesia oficial sin puestos asignados.", "ACCION_SUGERIDA": "Definir rol territorial."},
        {"TIPO": "UPZ", "REGISTRO": "data/upz_kennedy.geojson", "OBSERVACION": "No incluida en este paquete.", "ACCION_SUGERIDA": "Agregar GeoJSON oficial para análisis por UPZ."},
        {"TIPO": "CRUCE", "REGISTRO": "Puestos Localidad de Kennedy 2026", "OBSERVACION": f"{puestos_reporte_match} de {len(puestos)} puestos cruzados con el reporte complementario.", "ACCION_SUGERIDA": "Revisar puestos sin match antes de decisiones operativas."},
    ])

    with pd.ExcelWriter(SALIDA, engine="openpyxl") as writer:
        resumen_general.to_excel(writer, sheet_name="resumen_general", index=False)
        puestos.to_excel(writer, sheet_name="puestos_votacion", index=False)
        actividades.to_excel(writer, sheet_name="actividades_campana", index=False)
        mesas.to_excel(writer, sheet_name="mesas_trabajo", index=False)
        iglesias.to_excel(writer, sheet_name="iglesias", index=False)
        resumen_iglesia.to_excel(writer, sheet_name="resumen_iglesia", index=False)
        resumen_puesto.to_excel(writer, sheet_name="resumen_puesto", index=False)
        resumen_barrio.to_excel(writer, sheet_name="resumen_barrio", index=False)
        matriz.to_excel(writer, sheet_name="matriz_priorizacion", index=False)
        informe.to_excel(writer, sheet_name="informe_ejecutivo", index=False)
        control.to_excel(writer, sheet_name="control_calidad", index=False)

    print(f"Consolidado generado: {SALIDA}")


if __name__ == "__main__":
    main()
