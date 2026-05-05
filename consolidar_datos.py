from pathlib import Path
import re
import unicodedata

import numpy as np
import pandas as pd


DATA_DIR = Path(__file__).resolve().parent / "data"
ARCHIVO_CAMPANA = DATA_DIR / "CAMPAÑA CONGRESO 2026 KENNEDY (1).xlsx"
ARCHIVO_GESTION = DATA_DIR / "Copia de Gestión Edil Lorena Garzón - 17 de febrero, 17_07.xlsx"
ARCHIVO_VOTACION = DATA_DIR / "VOTACIÓN 2026.xlsx"
ARCHIVO_SALIDA = DATA_DIR / "kennedy_mira_consolidado.xlsx"

IGLESIAS_OFICIALES = ["CLASS ROMA", "KENNEDY CENTRAL", "PATIO BONITO", "CARVAJAL", "VALLADOLID"]
IGLESIAS_HISTORICAS_ANALISIS = ["CLASS ROMA", "KENNEDY CENTRAL", "PATIO BONITO", "CARVAJAL"]

IGLESIAS_BASE = pd.DataFrame(
    [
        {
            "IGLESIA": "CLASS ROMA",
            "LATITUD": 4.614359775316158,
            "LONGITUD": -74.17619195767098,
            "URL": "https://direcciones.idmji.org/es/iglesia/359/",
        },
        {
            "IGLESIA": "PATIO BONITO",
            "LATITUD": 4.646035122997863,
            "LONGITUD": -74.17300841534194,
            "URL": "https://direcciones.idmji.org/es/iglesia/301/",
        },
        {"IGLESIA": "KENNEDY CENTRAL", "LATITUD": 4.6217386978458155, "LONGITUD": -74.16501499477366, "URL": ""},
        {"IGLESIA": "CARVAJAL", "LATITUD": 4.616343469904612, "LONGITUD": -74.1404329155982, "URL": ""},
        {"IGLESIA": "VALLADOLID", "LATITUD": 4.647817860855581, "LONGITUD": -74.14806885512174, "URL": ""},
    ]
)


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("utf-8")
    texto = re.sub(r"[^A-Z0-9\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_puesto_match(valor):
    texto = normalizar_texto(valor)
    reemplazos = {
        "COLEGIO": "COL",
        "COL ": "COL ",
        "DISTRITAL": "DIST",
        "INSTITUCION EDUCATIVA": "IED",
    }
    for origen, destino in reemplazos.items():
        texto = re.sub(rf"\b{origen}\b", destino, texto)
    texto = re.sub(r"\bIED\b", "", texto)
    texto = re.sub(r"\bSEDE\b", "SEDE", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def normalizar_iglesia(valor):
    texto = normalizar_texto(valor)
    if not texto:
        return "SIN CLASIFICAR"
    if texto == "SIN CLASIFICAR":
        return "SIN CLASIFICAR"
    if "PATIO" in texto:
        return "PATIO BONITO"
    if "CARVAJAL" in texto:
        return "CARVAJAL"
    if "VALLADOLID" in texto:
        return "VALLADOLID"
    if "CLASS" in texto or re.search(r"\bCLAS\b", texto) or re.search(r"\bROMA\b", texto):
        return "CLASS ROMA"
    if "KENNEDY" in texto or "CENTRAL" in texto:
        return "KENNEDY CENTRAL"
    return "SIN CLASIFICAR"


def limpiar_columnas(df):
    df = df.copy()
    df.columns = [normalizar_texto(c).replace(" ", "_") for c in df.columns]
    return df


def leer_hoja(path, preferidas):
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo requerido: {path}")
    xl = pd.ExcelFile(path, engine="openpyxl")
    hojas_norm = {normalizar_texto(hoja): hoja for hoja in xl.sheet_names}
    for preferida in preferidas:
        hoja = hojas_norm.get(normalizar_texto(preferida))
        if hoja:
            return limpiar_columnas(pd.read_excel(path, sheet_name=hoja, engine="openpyxl"))
    return limpiar_columnas(pd.read_excel(path, sheet_name=xl.sheet_names[0], engine="openpyxl"))


def primera_columna(df, candidatas, contiene=None):
    for col in candidatas:
        if col in df.columns:
            return col
    if contiene:
        for col in df.columns:
            if all(fragmento in col for fragmento in contiene):
                return col
    return None


def valor_columna(df, candidatas, default=np.nan, contiene=None):
    col = primera_columna(df, candidatas, contiene=contiene)
    if col:
        return df[col]
    return pd.Series(default, index=df.index)


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


def coordenadas_por_iglesia(iglesia):
    fila = IGLESIAS_BASE[IGLESIAS_BASE["IGLESIA"].eq(iglesia)]
    if fila.empty:
        return pd.Series([np.nan, np.nan])
    return pd.Series([fila.iloc[0]["LATITUD"], fila.iloc[0]["LONGITUD"]])


def iglesia_mas_cercana(latitud, longitud):
    if pd.isna(latitud) or pd.isna(longitud):
        return "SIN CLASIFICAR"
    distancias = []
    for _, row in IGLESIAS_BASE.iterrows():
        distancia = (float(latitud) - row["LATITUD"]) ** 2 + (float(longitud) - row["LONGITUD"]) ** 2
        distancias.append((distancia, row["IGLESIA"]))
    return sorted(distancias, key=lambda x: x[0])[0][1]


def sugerir_iglesia_actual(iglesia_historica, barrio="", puesto="", latitud=np.nan, longitud=np.nan):
    iglesia_norm = normalizar_iglesia(iglesia_historica)

    contexto = " ".join([normalizar_texto(barrio), normalizar_texto(puesto)])
    if any(token in contexto for token in ["CASTILLA", "TABAKU", "TABACU", "TIBAKU"]):
        return "VALLADOLID"
    if "PATIO" in contexto:
        return "PATIO BONITO"
    if "CARVAJAL" in contexto or "ALQUERIA" in contexto:
        return "CARVAJAL"
    if "CLASS" in contexto or "ROMA" in contexto:
        return "CLASS ROMA"
    if "KENNEDY" in contexto or "BRITALIA" in contexto:
        return "KENNEDY CENTRAL"
    cercana = iglesia_mas_cercana(latitud, longitud)
    return cercana if cercana != "SIN CLASIFICAR" else iglesia_norm


def construir_marca_reasignacion(iglesia_historica, iglesia_sugerida):
    if iglesia_sugerida == "SIN CLASIFICAR" or iglesia_historica == "SIN CLASIFICAR":
        return "SIN SUGERENCIA"
    if iglesia_historica == iglesia_sugerida:
        return "MISMA IGLESIA"
    if iglesia_sugerida == "VALLADOLID":
        return "SUGERIDO VALLADOLID POR CERCANIA TERRITORIAL ACTUAL"
    return "SUGERENCIA TERRITORIAL ACTUAL"


def construir_puestos():
    puestos = leer_hoja(ARCHIVO_VOTACION, ["Hoja 5"])
    detalle = leer_hoja(ARCHIVO_VOTACION, ["Hoja 3"])

    puestos["PUESTO"] = valor_columna(puestos, ["PUESTO_DE_VOTACION", "PUESTO"]).astype(str).str.strip()
    puestos["PUESTO_ID"] = puestos["PUESTO"].apply(normalizar_texto)
    puestos["PUESTO_MATCH"] = puestos["PUESTO"].apply(normalizar_puesto_match)
    puestos["IGLESIA"] = valor_columna(puestos, ["IGLESIA_RESPONSABLE", "IGLESIA"]).apply(normalizar_iglesia)
    puestos["VOTOS_2026"] = pd.to_numeric(valor_columna(puestos, ["PROMEDIO_2026", "TOTAL_"]), errors="coerce").fillna(0)
    puestos["VOTOS_2023"] = pd.to_numeric(valor_columna(puestos, ["PROMEDIO_2023"]), errors="coerce").fillna(0)

    detalle["PUESTO"] = valor_columna(detalle, ["PUESTO_DE_VOTACION", "PUESTO"]).astype(str).str.strip()
    detalle["PUESTO_ID"] = detalle["PUESTO"].apply(normalizar_texto)
    detalle["PUESTO_MATCH"] = detalle["PUESTO"].apply(normalizar_puesto_match)
    detalle["DIRECCION"] = valor_columna(detalle, ["DIRECCION", "DIRECCION_"]).replace("", np.nan)
    detalle["BARRIO"] = valor_columna(detalle, ["BARRIO", "BARRIO_"]).replace("", np.nan)
    detalle["UPZ"] = valor_columna(detalle, ["UPZ", "ZONA"]).replace("", np.nan)
    detalle["IGLESIA_DETALLE"] = valor_columna(detalle, ["IGLESIA_RESPONSABLE", "IGLESIA"]).apply(normalizar_iglesia)
    if "COORDENADAS" in detalle.columns:
        detalle[["LATITUD", "LONGITUD"]] = detalle["COORDENADAS"].apply(extraer_lat_lon)
    else:
        detalle["LATITUD"] = np.nan
        detalle["LONGITUD"] = np.nan

    puestos = puestos.merge(
        detalle[["PUESTO_MATCH", "DIRECCION", "BARRIO", "UPZ", "LATITUD", "LONGITUD", "IGLESIA_DETALLE"]].drop_duplicates("PUESTO_MATCH"),
        on="PUESTO_MATCH",
        how="left",
    )
    puestos["IGLESIA"] = puestos["IGLESIA_DETALLE"].fillna(puestos["IGLESIA"]).apply(normalizar_iglesia)
    puestos["BARRIO"] = puestos["BARRIO"].fillna("SIN BARRIO")
    puestos["UPZ"] = puestos["UPZ"].fillna("SIN UPZ")
    puestos["VARIACION_ABSOLUTA"] = puestos["VOTOS_2026"] - puestos["VOTOS_2023"]
    puestos["VARIACION_PORCENTUAL"] = np.where(
        puestos["VOTOS_2023"].gt(0),
        puestos["VARIACION_ABSOLUTA"] / puestos["VOTOS_2023"],
        np.nan,
    )
    puestos["IGLESIA_ACTUAL_SUGERIDA"] = puestos.apply(
        lambda row: sugerir_iglesia_actual(row["IGLESIA"], row["BARRIO"], row["PUESTO"], row["LATITUD"], row["LONGITUD"]),
        axis=1,
    )
    puestos["MARCA_REASIGNACION"] = puestos.apply(
        lambda row: construir_marca_reasignacion(row["IGLESIA"], row["IGLESIA_ACTUAL_SUGERIDA"]),
        axis=1,
    )
    puestos = puestos[
        [
            "PUESTO_ID",
            "PUESTO",
            "DIRECCION",
            "BARRIO",
            "UPZ",
            "IGLESIA",
            "LATITUD",
            "LONGITUD",
            "VOTOS_2026",
            "VOTOS_2023",
            "VARIACION_ABSOLUTA",
            "VARIACION_PORCENTUAL",
            "IGLESIA_ACTUAL_SUGERIDA",
            "MARCA_REASIGNACION",
        ]
    ]
    return puestos.drop_duplicates("PUESTO_ID")


def construir_actividades():
    agenda_general = leer_hoja(ARCHIVO_CAMPANA, ["AGENDA GENERAL CON CANDIDATOS"])
    agenda_paralela = leer_hoja(ARCHIVO_CAMPANA, ["AGENDA PARALELA"])
    actividades = pd.concat([agenda_general, agenda_paralela], ignore_index=True)
    if actividades.empty:
        return pd.DataFrame(columns=["ACTIVIDAD_ID", "FECHA", "IGLESIA", "BARRIO", "DIRECCION", "LATITUD", "LONGITUD", "TIPO_ACTIVIDAD", "LIDER", "OBSERVACIONES"])

    salida = pd.DataFrame(index=actividades.index)
    salida["ACTIVIDAD_ID"] = [f"ACT-{i:04d}" for i in range(1, len(actividades) + 1)]
    salida["FECHA"] = pd.to_datetime(valor_columna(actividades, ["FECHA_CAMPANA", "FECHA"]), errors="coerce")
    salida["IGLESIA"] = valor_columna(actividades, ["SEDE", "IGLESIA"]).apply(normalizar_iglesia)
    salida["BARRIO"] = valor_columna(actividades, ["BARRIO", "BARRIO_"]).fillna("SIN BARRIO").replace("", "SIN BARRIO")
    salida["DIRECCION"] = valor_columna(actividades, ["DIRECCION", "DIRECCION_"])
    coord_col = primera_columna(actividades, ["COORDENADAS"], contiene=["COORDEN"])
    if coord_col:
        salida[["LATITUD", "LONGITUD"]] = actividades[coord_col].apply(extraer_lat_lon)
    else:
        salida["LATITUD"] = np.nan
        salida["LONGITUD"] = np.nan
    faltan_coord = salida["LATITUD"].isna()
    salida.loc[faltan_coord, ["LATITUD", "LONGITUD"]] = salida.loc[faltan_coord, "IGLESIA"].apply(coordenadas_por_iglesia).values
    salida["TIPO_ACTIVIDAD"] = valor_columna(actividades, ["ACTIVIDAD", "TIPO_DE_ACTIVIDAD"]).fillna("SIN TIPO")
    salida["LIDER"] = valor_columna(actividades, ["LIDER_Y_CELULAR", "LIDER"])
    salida["OBSERVACIONES"] = valor_columna(actividades, ["DETALLE_DE_LA_ACTIVIDAD", "LOGROS_JUSTIFICACION", "OBSERVACIONES"])
    salida["IGLESIA_ACTUAL_SUGERIDA"] = salida.apply(
        lambda row: sugerir_iglesia_actual(row["IGLESIA"], row["BARRIO"], "", row["LATITUD"], row["LONGITUD"]),
        axis=1,
    )
    salida["MARCA_REASIGNACION"] = salida.apply(
        lambda row: construir_marca_reasignacion(row["IGLESIA"], row["IGLESIA_ACTUAL_SUGERIDA"]),
        axis=1,
    )
    return salida


def construir_mesas():
    mesas_campana = leer_hoja(ARCHIVO_CAMPANA, ["Mesas"])
    mesas_gestion = leer_hoja(ARCHIVO_GESTION, ["SEGUIMIENTO MESAS DE TRABAJO"])
    mesas = pd.concat([mesas_campana, mesas_gestion], ignore_index=True)
    if mesas.empty:
        return pd.DataFrame(columns=["MESA_ID", "FECHA", "IGLESIA", "BARRIO", "DIRECCION", "LATITUD", "LONGITUD", "TEMA", "ENTIDADES", "LIDER", "ESTADO", "OBSERVACIONES"])

    salida = pd.DataFrame(index=mesas.index)
    salida["MESA_ID"] = [f"MESA-{i:04d}" for i in range(1, len(mesas) + 1)]
    salida["FECHA"] = pd.to_datetime(valor_columna(mesas, ["FECHA", "FECHA_DE_INICIO"]), errors="coerce")
    salida["IGLESIA"] = valor_columna(mesas, ["IGLESIA", "SEDE"]).apply(normalizar_iglesia)
    salida["BARRIO"] = valor_columna(mesas, ["BARRIO", "BARRIO_"]).fillna("SIN BARRIO").replace("", "SIN BARRIO")
    salida["DIRECCION"] = valor_columna(mesas, ["DIRECCION", "DIRECCION_DEL_BARRIO"])
    salida["LATITUD"] = np.nan
    salida["LONGITUD"] = np.nan
    coord_col = primera_columna(mesas, ["COORDENADAS", "GEOREFERENCIACION"], contiene=["COORDEN"])
    if coord_col:
        salida[["LATITUD", "LONGITUD"]] = mesas[coord_col].apply(extraer_lat_lon)
    faltan_coord = salida["LATITUD"].isna()
    salida.loc[faltan_coord, ["LATITUD", "LONGITUD"]] = salida.loc[faltan_coord, "IGLESIA"].apply(coordenadas_por_iglesia).values
    salida["TEMA"] = valor_columna(mesas, ["TEMA", "TEMAS", "OBJETIVO_DE_LA_REUNION", "NOMBRE_GESTION"]).fillna("SIN TEMA")
    salida["ENTIDADES"] = valor_columna(mesas, ["ENTIDAD", "ENTIDADES"])
    salida["LIDER"] = valor_columna(mesas, ["LIDER", "SOLICITANTE", "RESPONSABLE"])
    salida["ESTADO"] = valor_columna(mesas, ["ESTADO", "ESTADO_DE_SEGUIMIENTO"]).fillna("SIN ESTADO")
    salida["OBSERVACIONES"] = valor_columna(mesas, ["OBSERVACIONES", "CONCLUSIONES", "RESULTADOS_DE_LA_GESTION_PERCEPCION_DEL_CIUDADANO"])
    salida["IGLESIA_ACTUAL_SUGERIDA"] = salida.apply(
        lambda row: sugerir_iglesia_actual(row["IGLESIA"], row["BARRIO"], "", row["LATITUD"], row["LONGITUD"]),
        axis=1,
    )
    salida["MARCA_REASIGNACION"] = salida.apply(
        lambda row: construir_marca_reasignacion(row["IGLESIA"], row["IGLESIA_ACTUAL_SUGERIDA"]),
        axis=1,
    )
    return salida


def asignar_prioridad(puestos):
    df = puestos.copy()
    votos = df["VOTOS_2026"].fillna(0)
    umbral_alto = votos.quantile(0.65) if len(votos) else 0
    umbral_medio = votos.quantile(0.35) if len(votos) else 0
    prioridad, variable, diagnostico, accion, temporalidad = [], [], [], [], []

    for _, row in df.iterrows():
        alta_votacion = row["VOTOS_2026"] >= umbral_alto
        votacion_media = row["VOTOS_2026"] >= umbral_medio
        caida = row["VARIACION_ABSOLUTA"] < 0
        crecimiento = row["VARIACION_ABSOLUTA"] > 0
        pocas_actividades = row["ACTIVIDADES_CAMPANA"] <= 1
        sin_mesas = row["MESAS_TRABAJO_BARRIO"] == 0
        barrio_gestion_sin_mejora = row["MESAS_TRABAJO_BARRIO"] > 0 and row["VARIACION_ABSOLUTA"] <= 0

        if (caida and alta_votacion) or (alta_votacion and pocas_actividades) or (alta_votacion and sin_mesas) or barrio_gestion_sin_mejora:
            prioridad.append("ALTA")
            temporalidad.append("0-15 dias")
            if caida and alta_votacion:
                variable.append("Caida electoral con votacion relevante")
                diagnostico.append("Puesto de recuperacion: conserva volumen electoral, pero presenta perdida frente a 2023.")
                accion.append("Realizar visita territorial, validar liderazgos y activar seguimiento semanal.")
            elif alta_votacion and pocas_actividades:
                variable.append("Alta votacion con baja actividad de campana")
                diagnostico.append("Existe concentracion electoral sin presencia comunitaria suficiente registrada.")
                accion.append("Programar agenda de campana focalizada y contacto con lideres del entorno.")
            elif alta_votacion and sin_mesas:
                variable.append("Alta votacion sin mesas de trabajo")
                diagnostico.append("El puesto tiene peso electoral, pero no evidencia gestion comunitaria asociada al barrio.")
                accion.append("Abrir mesa de trabajo barrial y vincular responsables de iglesia.")
            else:
                variable.append("Gestion territorial sin mejora electoral")
                diagnostico.append("Hay presencia comunitaria, pero no se observa traduccion positiva en resultado electoral.")
                accion.append("Revisar calidad de la gestion, beneficiarios y conversion a compromiso electoral.")
        elif votacion_media or crecimiento:
            prioridad.append("MEDIA")
            temporalidad.append("15-30 dias")
            if crecimiento:
                variable.append("Crecimiento por consolidar")
                diagnostico.append("Puesto de consolidacion: muestra avance y requiere continuidad operativa.")
                accion.append("Mantener contacto territorial y replicar practicas en barrios cercanos.")
            else:
                variable.append("Oportunidad de crecimiento")
                diagnostico.append("Puesto con votacion media y margen de crecimiento por presencia territorial.")
                accion.append("Desarrollar actividades de bajo costo y seguimiento quincenal.")
        else:
            prioridad.append("BAJA")
            variable.append("Baja votacion o informacion insuficiente")
            diagnostico.append("El puesto requiere monitoreo, depuracion de datos y lectura territorial complementaria.")
            accion.append("Actualizar informacion y observar cambios antes de asignar recursos intensivos.")
            temporalidad.append("30-45 dias")

    df["PRIORIDAD"] = prioridad
    df["RAZON_PRIORIDAD"] = diagnostico
    df["ACCION_RECOMENDADA"] = accion
    df["NIVEL_PRIORIDAD"] = prioridad
    df["VARIABLE_CRITICA"] = variable
    df["DIAGNOSTICO"] = diagnostico
    df["TEMPORALIDAD"] = temporalidad
    df["RESPONSABLE_SUGERIDO"] = df["IGLESIA"].where(df["IGLESIA"].isin(IGLESIAS_OFICIALES), "Equipo territorial")
    return df


def construir_resumenes(puestos, actividades, mesas):
    puestos["ACTIVIDADES_CAMPANA"] = puestos["IGLESIA"].map(actividades.groupby("IGLESIA").size()).fillna(0).astype(int)
    puestos["MESAS_TRABAJO_BARRIO"] = puestos["BARRIO"].map(mesas.groupby("BARRIO").size()).fillna(0).astype(int)
    puestos = asignar_prioridad(puestos)

    oficiales = puestos[puestos["IGLESIA"].isin(IGLESIAS_HISTORICAS_ANALISIS)].copy()
    act_oficial = actividades[actividades["IGLESIA"].isin(IGLESIAS_HISTORICAS_ANALISIS)].copy()
    mesas_oficial = mesas[mesas["IGLESIA"].isin(IGLESIAS_HISTORICAS_ANALISIS)].copy()

    resumen_iglesia = oficiales.groupby("IGLESIA", as_index=False).agg(
        VOTOS_2026=("VOTOS_2026", "sum"),
        VOTOS_2023=("VOTOS_2023", "sum"),
        PUESTOS=("PUESTO_ID", "nunique"),
        BARRIOS=("BARRIO", "nunique"),
    )
    resumen_iglesia = pd.DataFrame({"IGLESIA": IGLESIAS_HISTORICAS_ANALISIS}).merge(resumen_iglesia, on="IGLESIA", how="left")
    resumen_iglesia[["VOTOS_2026", "VOTOS_2023", "PUESTOS", "BARRIOS"]] = resumen_iglesia[
        ["VOTOS_2026", "VOTOS_2023", "PUESTOS", "BARRIOS"]
    ].fillna(0)
    resumen_iglesia["VARIACION_ABSOLUTA"] = resumen_iglesia["VOTOS_2026"] - resumen_iglesia["VOTOS_2023"]
    resumen_iglesia["VARIACION_PORCENTUAL"] = np.where(
        resumen_iglesia["VOTOS_2023"].gt(0),
        resumen_iglesia["VARIACION_ABSOLUTA"] / resumen_iglesia["VOTOS_2023"],
        np.nan,
    )
    resumen_iglesia["ACTIVIDADES_CAMPANA"] = resumen_iglesia["IGLESIA"].map(act_oficial.groupby("IGLESIA").size()).fillna(0).astype(int)
    resumen_iglesia["MESAS_TRABAJO"] = resumen_iglesia["IGLESIA"].map(mesas_oficial.groupby("IGLESIA").size()).fillna(0).astype(int)

    for campo, asc in [("PUESTO_MAYOR_VOTACION", False), ("PUESTO_MAYOR_CAIDA", True), ("PUESTO_MAYOR_CRECIMIENTO", False)]:
        valores = {}
        for iglesia, grupo in oficiales.groupby("IGLESIA"):
            sort_col = "VOTOS_2026" if campo == "PUESTO_MAYOR_VOTACION" else "VARIACION_ABSOLUTA"
            valores[iglesia] = grupo.sort_values(sort_col, ascending=asc).iloc[0]["PUESTO"] if not grupo.empty else "SIN DATOS"
        resumen_iglesia[campo] = resumen_iglesia["IGLESIA"].map(valores)
        resumen_iglesia[campo] = resumen_iglesia[campo].fillna("SIN DATOS")

    resumen_iglesia["LECTURA_ESTRATEGICA"] = resumen_iglesia.apply(
        lambda r: (
            f"{r['IGLESIA']} concentra {int(r['VOTOS_2026'])} votos 2026 en {int(r['PUESTOS'])} puestos. "
            f"La variacion absoluta es {int(r['VARIACION_ABSOLUTA'])}, con presencia comunitaria registrada en "
            f"{int(r['ACTIVIDADES_CAMPANA'])} actividades y {int(r['MESAS_TRABAJO'])} mesas."
        ),
        axis=1,
    )
    resumen_iglesia["RECOMENDACION"] = resumen_iglesia.apply(
        lambda r: "Priorizar recuperacion electoral y gestion focalizada."
        if r["VARIACION_ABSOLUTA"] < 0
        else "Consolidar crecimiento y sostener eficiencia territorial de campana.",
        axis=1,
    )
    resumen_iglesia = resumen_iglesia[
        [
            "IGLESIA",
            "VOTOS_2026",
            "VOTOS_2023",
            "VARIACION_ABSOLUTA",
            "VARIACION_PORCENTUAL",
            "PUESTOS",
            "ACTIVIDADES_CAMPANA",
            "MESAS_TRABAJO",
            "BARRIOS",
            "PUESTO_MAYOR_VOTACION",
            "PUESTO_MAYOR_CAIDA",
            "PUESTO_MAYOR_CRECIMIENTO",
            "LECTURA_ESTRATEGICA",
            "RECOMENDACION",
        ]
    ]

    resumen_puesto = oficiales[
        [
            "PUESTO",
            "IGLESIA",
            "BARRIO",
            "UPZ",
            "VOTOS_2026",
            "VOTOS_2023",
            "VARIACION_ABSOLUTA",
            "VARIACION_PORCENTUAL",
            "IGLESIA_ACTUAL_SUGERIDA",
            "MARCA_REASIGNACION",
            "ACTIVIDADES_CAMPANA",
            "MESAS_TRABAJO_BARRIO",
            "PRIORIDAD",
            "ACCION_RECOMENDADA",
        ]
    ].copy()

    resumen_barrio = oficiales.groupby(["BARRIO", "IGLESIA", "UPZ"], as_index=False).agg(
        VOTOS_2026=("VOTOS_2026", "sum"),
        VOTOS_2023=("VOTOS_2023", "sum"),
        PUESTOS=("PUESTO_ID", "nunique"),
    )
    resumen_barrio["VARIACION_ABSOLUTA"] = resumen_barrio["VOTOS_2026"] - resumen_barrio["VOTOS_2023"]
    resumen_barrio["VARIACION_PORCENTUAL"] = np.where(
        resumen_barrio["VOTOS_2023"].gt(0),
        resumen_barrio["VARIACION_ABSOLUTA"] / resumen_barrio["VOTOS_2023"],
        np.nan,
    )
    resumen_barrio["ACTIVIDADES_CAMPANA"] = resumen_barrio["IGLESIA"].map(act_oficial.groupby("IGLESIA").size()).fillna(0).astype(int)
    resumen_barrio["MESAS_TRABAJO"] = resumen_barrio["BARRIO"].map(mesas_oficial.groupby("BARRIO").size()).fillna(0).astype(int)
    resumen_barrio["PRIORIDAD"] = np.where(
        (resumen_barrio["VARIACION_ABSOLUTA"] < 0) & (resumen_barrio["VOTOS_2026"] >= resumen_barrio["VOTOS_2026"].quantile(0.6)),
        "ALTA",
        np.where(resumen_barrio["VOTOS_2026"] >= resumen_barrio["VOTOS_2026"].quantile(0.35), "MEDIA", "BAJA"),
    )
    resumen_barrio["ACCION_RECOMENDADA"] = resumen_barrio["PRIORIDAD"].map(
        {
            "ALTA": "Intervenir con agenda barrial, mesa de gestion y seguimiento electoral.",
            "MEDIA": "Consolidar presencia comunitaria y monitorear conversion electoral.",
            "BAJA": "Actualizar informacion y mantener observacion.",
        }
    )
    resumen_barrio = resumen_barrio[
        [
            "BARRIO",
            "IGLESIA",
            "UPZ",
            "VOTOS_2026",
            "VOTOS_2023",
            "VARIACION_ABSOLUTA",
            "VARIACION_PORCENTUAL",
            "PUESTOS",
            "ACTIVIDADES_CAMPANA",
            "MESAS_TRABAJO",
            "PRIORIDAD",
            "ACCION_RECOMENDADA",
        ]
    ]

    matriz_priorizacion = oficiales[
        [
            "PUESTO_ID",
            "PUESTO",
            "IGLESIA",
            "BARRIO",
            "UPZ",
            "VOTOS_2026",
            "VOTOS_2023",
            "VARIACION_ABSOLUTA",
            "VARIACION_PORCENTUAL",
            "IGLESIA_ACTUAL_SUGERIDA",
            "MARCA_REASIGNACION",
            "ACTIVIDADES_CAMPANA",
            "MESAS_TRABAJO_BARRIO",
            "NIVEL_PRIORIDAD",
            "VARIABLE_CRITICA",
            "DIAGNOSTICO",
            "ACCION_RECOMENDADA",
            "TEMPORALIDAD",
            "RESPONSABLE_SUGERIDO",
        ]
    ].copy()

    puestos_votacion = puestos[
        [
            "PUESTO_ID",
            "PUESTO",
            "DIRECCION",
            "BARRIO",
            "UPZ",
            "IGLESIA",
            "LATITUD",
            "LONGITUD",
            "VOTOS_2026",
            "VOTOS_2023",
            "VARIACION_ABSOLUTA",
            "VARIACION_PORCENTUAL",
            "IGLESIA_ACTUAL_SUGERIDA",
            "MARCA_REASIGNACION",
            "ACTIVIDADES_CAMPANA",
            "MESAS_TRABAJO_BARRIO",
            "PRIORIDAD",
            "RAZON_PRIORIDAD",
            "ACCION_RECOMENDADA",
        ]
    ].copy()

    return puestos_votacion, resumen_iglesia, resumen_puesto, resumen_barrio, matriz_priorizacion


def construir_informe(resumen_iglesia, resumen_barrio, matriz_priorizacion):
    votos_2026 = resumen_iglesia["VOTOS_2026"].sum()
    votos_2023 = resumen_iglesia["VOTOS_2023"].sum()
    variacion = votos_2026 - votos_2023
    var_pct = variacion / votos_2023 if votos_2023 else np.nan
    top_iglesia = resumen_iglesia.sort_values("VOTOS_2026", ascending=False).iloc[0]
    criticos = matriz_priorizacion[matriz_priorizacion["NIVEL_PRIORIDAD"].eq("ALTA")].head(8)
    barrios = resumen_barrio.sort_values("VOTOS_2026", ascending=False).head(5)

    textos = [
        (
            "Resumen general",
            f"La base consolidada analiza la campana Congreso 2026 del Partido MIRA en Kennedy a partir de puestos de votacion, "
            f"actividades de campana, mesas de trabajo e iglesias oficiales. El total consolidado es de {votos_2026:,.0f} votos 2026 "
            f"frente a {votos_2023:,.0f} votos 2023, con una variacion absoluta de {variacion:,.0f} votos "
            f"({var_pct:.1%} si existe base comparable).",
        ),
        (
            "Hallazgos principales",
            f"La mayor concentracion territorial se observa en {top_iglesia['IGLESIA']}. La lectura combina rendimiento electoral, "
            "presencia comunitaria y brechas entre gestion y resultado electoral para distinguir puestos de recuperacion, consolidacion y oportunidad.",
        ),
        (
            "Lectura por iglesia",
            "El analisis electoral historico conserva las iglesias responsables registradas en la votacion: CLASS ROMA, KENNEDY CENTRAL, PATIO BONITO y CARVAJAL. "
            "VALLADOLID se mantiene como referencia territorial actual sugerida para zonas cercanas como Castilla, sin alterar la comparacion del momento electoral.",
        ),
        (
            "Lectura territorial",
            "Los barrios con mayor concentracion electoral requieren seguimiento diferenciado: "
            + ", ".join(barrios["BARRIO"].astype(str).tolist())
            + ". La presencia comunitaria debe evaluarse por su capacidad de convertirse en rendimiento electoral.",
        ),
        (
            "Puestos criticos",
            "Los puestos de mayor prioridad son: "
            + (", ".join(criticos["PUESTO"].astype(str).tolist()) if not criticos.empty else "no se identifican puestos de prioridad alta con la informacion disponible")
            + ".",
        ),
        (
            "Recomendaciones estrategicas",
            "Concentrar recursos en puestos con caida y votacion relevante, reforzar actividades donde existe alta votacion con baja presencia de campana, "
            "y abrir mesas de trabajo en barrios con concentracion electoral sin gestion comunitaria suficiente.",
        ),
        (
            "Agenda sugerida de seguimiento",
            "Semana 1: puestos de recuperacion y validacion de lideres. Semana 2: mesas de trabajo barriales. Semana 3: consolidacion de iglesias con crecimiento. "
            "Semana 4: revision de indicadores, actualizacion de base y ajuste de prioridades.",
        ),
    ]
    return pd.DataFrame(textos, columns=["SECCION", "TEXTO"])


def consolidar():
    DATA_DIR.mkdir(exist_ok=True)
    puestos = construir_puestos()
    actividades = construir_actividades()
    mesas = construir_mesas()
    puestos_votacion, resumen_iglesia, resumen_puesto, resumen_barrio, matriz_priorizacion = construir_resumenes(puestos, actividades, mesas)
    informe = construir_informe(resumen_iglesia, resumen_barrio, matriz_priorizacion)

    with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
        puestos_votacion.to_excel(writer, sheet_name="puestos_votacion", index=False)
        actividades.to_excel(writer, sheet_name="actividades_campana", index=False)
        mesas.to_excel(writer, sheet_name="mesas_trabajo", index=False)
        IGLESIAS_BASE.to_excel(writer, sheet_name="iglesias", index=False)
        resumen_iglesia.to_excel(writer, sheet_name="resumen_iglesia", index=False)
        resumen_puesto.to_excel(writer, sheet_name="resumen_puesto", index=False)
        resumen_barrio.to_excel(writer, sheet_name="resumen_barrio", index=False)
        matriz_priorizacion.to_excel(writer, sheet_name="matriz_priorizacion", index=False)
        informe.to_excel(writer, sheet_name="informe_ejecutivo", index=False)
    return ARCHIVO_SALIDA


if __name__ == "__main__":
    salida = consolidar()
    print(f"Archivo consolidado generado: {salida}")
