from app import *
m = crear_mapa_base()
html = m.get_root().render()
print("Map size without data:", len(html) / 1024, "KB")

puestos_f = puestos.copy()
iglesias_f = iglesias.copy()
acts_f = actividades.copy()
mesas_f = mesas.copy()
m2 = crear_mapa(puestos_f, iglesias_f, acts_f, mesas_f)
html2 = m2.get_root().render()
print("Map size WITH data:", len(html2) / 1024, "KB")
