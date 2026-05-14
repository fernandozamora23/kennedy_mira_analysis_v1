import sys
from pathlib import Path
from app import cargar_datos, crear_mapa

data = cargar_datos(Path("data/kennedy_mira_consolidado.xlsx"))
try:
    mapa = crear_mapa(data["puestos"], data["iglesias"], data["actividades"], data["mesas"])
    print("Mapa creado correctamente.")
except Exception as e:
    import traceback
    traceback.print_exc()
