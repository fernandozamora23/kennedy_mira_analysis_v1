import streamlit as st
from app import cargar_datos, crear_mapa
from streamlit_folium import st_folium
from pathlib import Path

data = cargar_datos(Path("data/kennedy_mira_consolidado.xlsx"))
mapa = crear_mapa(data["puestos"], data["iglesias"], data["actividades"], data["mesas"])

st_folium(
    mapa,
    height=760,
    use_container_width=True,
    returned_objects=[],
)
