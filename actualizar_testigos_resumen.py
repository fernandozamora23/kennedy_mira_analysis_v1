from pathlib import Path

import numpy as np
import pandas as pd


SOURCE_DIR = Path(__file__).resolve().parent.parent / "Testigos"
OUTPUT_CSV = Path(__file__).resolve().parent / "data" / "testigos_resumen_2026.csv"

TEMPLO_MAP = {
    "KENNEDY": "KENNEDY CENTRAL",
    "KENNEDY CLASS ROMA": "CLASS ROMA",
}


def yes_mask(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().str.upper().isin({"SI", "SÍ", "YES", "TRUE", "1", "X"})


def numeric_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series(0, index=df.index)
    return pd.to_numeric(df[column], errors="coerce").fillna(0)


def boolean_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series(False, index=df.index)
    return yes_mask(df[column])


def normalize_templo(value: object, fallback: str) -> str:
    text = str(value).strip().upper() if pd.notna(value) else fallback
    return TEMPLO_MAP.get(text, text)


def build_summary(source_dir: Path = SOURCE_DIR) -> pd.DataFrame:
    rows = []
    for path in sorted(source_dir.glob("*.xlsx")):
        df = pd.read_excel(path)
        if df.empty:
            continue

        fallback = path.stem.upper()
        fallback = fallback.replace("KENNEDY CLASS ROMA", "CLASS ROMA").replace("KENNEDY", "KENNEDY CENTRAL")
        templo_raw = df["templo_infomira"].dropna().iloc[0] if "templo_infomira" in df.columns and df["templo_infomira"].notna().any() else fallback
        templo = normalize_templo(templo_raw, fallback)

        total = len(df)
        mesa = boolean_series(df, "testigo_mesa_o_remanente")
        comision = boolean_series(df, "testigo_comision_escrutadora")
        beneficiario = boolean_series(df, "beneficiario_mesas_trabajo")
        lider = boolean_series(df, "Es líder")
        referidos = numeric_series(df, "Cantidad de referidos")
        referidos_inactivos = numeric_series(df, "Cantidad de referidos inactivos")
        referidos_activos = (referidos - referidos_inactivos).clip(lower=0)
        cantidad_mesas = numeric_series(df, "cantidad_mesas_trabajo")

        row = {
            "TEMPLO": templo,
            "TOTAL_TESTIGOS": total,
            "TESTIGOS_MESA_O_REMANENTE": int(mesa.sum()),
            "TESTIGOS_COMISION_ESCRUTADORA": int(comision.sum()),
            "BENEFICIARIOS_MESAS_TRABAJO": int(beneficiario.sum()),
            "CANTIDAD_MESAS_TRABAJO_ASOCIADAS": int(cantidad_mesas.sum()),
            "TESTIGOS_DOBLE_ROL": int((mesa & comision).sum()),
            "LIDERES": int(lider.sum()),
            "NO_LIDERES": int((~lider).sum()),
            "TESTIGOS_CON_REFERIDOS": int((referidos > 0).sum()),
            "REFERIDOS_REGISTRADOS": int(referidos.sum()),
            "REFERIDOS_INACTIVOS": int(referidos_inactivos.sum()),
            "REFERIDOS_ACTIVOS_ESTIMADOS": int(referidos_activos.sum()),
        }
        rows.append(row)

    summary = pd.DataFrame(rows)
    if summary.empty:
        return summary

    summary = summary.groupby("TEMPLO", as_index=False).sum(numeric_only=True)
    ratios = {
        "PCT_MESA_O_REMANENTE": "TESTIGOS_MESA_O_REMANENTE",
        "PCT_COMISION": "TESTIGOS_COMISION_ESCRUTADORA",
        "PCT_BENEFICIARIOS_MESAS": "BENEFICIARIOS_MESAS_TRABAJO",
        "PCT_LIDERES": "LIDERES",
        "PCT_CON_REFERIDOS": "TESTIGOS_CON_REFERIDOS",
    }
    for pct_col, value_col in ratios.items():
        summary[pct_col] = np.where(summary["TOTAL_TESTIGOS"].gt(0), summary[value_col] / summary["TOTAL_TESTIGOS"], 0)

    summary["REFERIDOS_POR_TESTIGO"] = np.where(
        summary["TOTAL_TESTIGOS"].gt(0),
        summary["REFERIDOS_REGISTRADOS"] / summary["TOTAL_TESTIGOS"],
        0,
    )
    summary["REFERIDOS_POR_LIDER"] = np.where(
        summary["LIDERES"].gt(0),
        summary["REFERIDOS_REGISTRADOS"] / summary["LIDERES"],
        0,
    )

    return summary.sort_values("TEMPLO")


def main() -> None:
    summary = build_summary()
    if summary.empty:
        raise SystemExit(f"No se encontraron archivos .xlsx con datos en {SOURCE_DIR}")
    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    summary.to_csv(OUTPUT_CSV, index=False)
    print(f"Resumen actualizado: {OUTPUT_CSV}")
    print(summary[["TEMPLO", "TOTAL_TESTIGOS", "LIDERES", "TESTIGOS_CON_REFERIDOS", "REFERIDOS_REGISTRADOS"]].to_string(index=False))


if __name__ == "__main__":
    main()
