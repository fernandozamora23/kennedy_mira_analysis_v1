# Prompt para Codex

Actúa como analista de datos electorales y territoriales para un equipo asesor del Concejo de Bogotá. Tengo tres archivos Excel:

1. `CAMPAÑA CONGRESO 2026 KENNEDY.xlsx`: agenda de campaña, reuniones, actividades territoriales y mesas.
2. `GESTION_EDIL_LORENA.xlsx`: seguimiento a mesas de trabajo y gestión comunitaria.
3. `VOTACION_2026.xlsx`: resultados de Congreso 2026, comparación con Concejo/JAL 2023, puestos de votación, direcciones, coordenadas e iglesia responsable.

Necesito que revises y mejores el archivo `app.py` para construir un dashboard en Streamlit que permita:

- limpiar columnas y textos automáticamente;
- leer las hojas principales de los tres Excel;
- cruzar puestos de votación con coordenadas;
- cruzar puestos con iglesia responsable;
- comparar promedio 2023 vs promedio 2026;
- calcular variación absoluta y porcentual;
- identificar puestos que subieron y bajaron;
- identificar si hubo mesa de trabajo asociada;
- generar una matriz de priorización territorial;
- mostrar un mapa interactivo de Kennedy con Folium;
- cargar una capa GeoJSON de UPZ o UPL de Kennedy;
- diferenciar puestos de votación, actividades de campaña y mesas de trabajo;
- permitir filtros por iglesia, resultado de variación, mesa de trabajo y UPZ/UPL;
- exportar tablas en CSV para anexos del informe.

El resultado debe ser profesional, reproducible y preparado para generar hallazgos del informe político-territorial.
