# Dashboard geopolítico Kennedy — MIRA 2026

## Objetivo
Construir un informe territorial y electoral con:
- votación Congreso 2026,
- comparación con Concejo/JAL 2023,
- reuniones de campaña,
- mesas de trabajo,
- puestos de votación con direcciones/coordenadas,
- lectura por iglesia/templo responsable,
- mapa interactivo por UPZ/UPL.

## Estructura esperada
Cree una carpeta `data/` y ponga allí los archivos con estos nombres:

```text
data/CAMPAÑA CONGRESO 2026 KENNEDY.xlsx
data/GESTION_EDIL_LORENA.xlsx
data/VOTACION_2026.xlsx
```

Opcional: descargue desde IDECA/Datos Abiertos Bogotá una capa GeoJSON de UPZ o UPL y guárdela como:

```text
data/kennedy_upz.geojson
```

## Ejecutar

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Productos del dashboard
1. Indicadores agregados.
2. Resumen por iglesia responsable.
3. Puestos con mayor crecimiento y mayor caída.
4. Mapa interactivo con puestos, actividades y heatmap.
5. Matriz de priorización territorial.
6. Exportables CSV para anexos del informe.

## Recomendación metodológica
Para un informe profesional, use el dashboard para producir:
- hallazgos electorales,
- hallazgos territoriales,
- brechas entre campaña, gestión y resultado electoral,
- priorización de puestos para seguimiento 2026–2027,
- plan de mesas de trabajo por iglesia y UPZ/UPL.
