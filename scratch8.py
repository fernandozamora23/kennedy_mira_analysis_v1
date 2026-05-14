import folium
m = folium.Map(location=[45.5236, -122.6750])
html = m.get_root().render()
print(html[:100])
