import urllib.request
import json
import ssl

ssl._create_default_https_context = ssl._create_unverified_context

urls = [
    "https://raw.githubusercontent.com/ljgomezz/Bogota-GeoJSON/master/UPZ.geojson",
    "https://raw.githubusercontent.com/danielm-12/bogota_maps/master/data/upz.geojson",
    "https://raw.githubusercontent.com/jorge-castaneda/BogotaGeoJSON/master/upz.geojson",
    "https://raw.githubusercontent.com/finiterank/mapa-bogota/master/data/upz.geojson",
    "https://raw.githubusercontent.com/scadare/bogota_upz/master/upz.geojson"
]

data = None
for url in urls:
    try:
        print(f"Probando {url}")
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            data = json.loads(response.read().decode())
            print(f"¡Éxito con {url}!")
            break
    except Exception as e:
        print(f"Error: {e}")

if data:
    out_path = "data/upz_kennedy.geojson"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f)
    print(f"Guardado exitosamente en {out_path}")
else:
    print("No se pudo descargar de ninguna URL.")
