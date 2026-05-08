# Dashboard territorial-electoral Kennedy MIRA

Este paquete reorganiza el proyecto en un flujo profesional:

```text
Excel originales y reporte Puestos Localidad de Kennedy 2026
        ↓
consolidar_datos.py
        ↓
data/kennedy_mira_consolidado.xlsx
        ↓
app.py
        ↓
dashboard territorial-electoral
```

## Archivos principales

- `app.py`: dashboard en Streamlit.
- `consolidar_datos.py`: script para regenerar el Excel maestro desde los archivos originales y el reporte complementario de puestos.
- `data/kennedy_mira_consolidado.xlsx`: base maestra consolidada que lee el dashboard.
- `data/Puestos Localidad de Kennedy 2026.xlsx`: reporte complementario con mesas 2026, testigos 2023, afinidad E-11, dirección 2026 y templo reportado por puesto.
- `requirements.txt`: dependencias.

## Cifras oficiales usadas en el tablero

El tablero separa el total general de Kennedy de los análisis filtrados por iglesia.

- Total Kennedy votos promedio 2026: 7.348
- Total Kennedy votos promedio 2023: 7.170,5
- Variación absoluta: +177,5
- Variación porcentual: +2,48%
- Puestos analizados: 123
- Iglesias oficiales: 5
- Cruce con `Puestos Localidad de Kennedy 2026`: 123 de 123 puestos.

Iglesias oficiales:

1. CLASS ROMA
2. KENNEDY CENTRAL
3. PATIO BONITO
4. CARVAJAL
5. VALLADOLID

Nota: Valladolid queda como iglesia oficial, aunque no tiene puestos electorales asignados en la matriz base.

## Separación electoral

El dashboard separa explícitamente dos bloques:

- **JAL / Concejo 2023**: usa las columnas `JAL_2023` y `MIRA_CONCEJO_2023`.
- **Cámara / Senado 2026**: usa las columnas `CAMARA_2026` y `SENADO_2026`.

El promedio 2023 y el promedio 2026 se conservan como indicadores generales para comparar variación electoral, pero el análisis por corporación se presenta por separado.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Actualizar en GitHub

Reemplaza los archivos del repositorio por los de este paquete y ejecuta:

```bash
git add app.py consolidar_datos.py requirements.txt README.md data/kennedy_mira_consolidado.xlsx
git commit -m "Consolida base maestra y redisenia dashboard territorial electoral"
git push
```

## UPZ

El dashboard está preparado para cargar una capa:

```text
data/upz_kennedy.geojson
```

Si agregas ese archivo, el mapa la dibuja como capa territorial. Si no existe, el tablero sigue funcionando y usa la columna `UPZ` si está diligenciada en el Excel maestro.

## Seguridad opcional

Si quieres proteger el dashboard con usuario y contraseña en Streamlit Cloud, configura Secrets así:

```toml
[auth]
usuario = "fernando"
password = "KennedyMira2026!"
```

Si no configuras `[auth]`, el tablero abre sin login.

## Base simple con Google Sheets

Para que los cambios guardados en la app online no se pierdan al reiniciar Streamlit Cloud, configura una hoja de Google Sheets como base de ajustes.

1. Crea un Google Sheet vacío.
2. Crea una cuenta de servicio en Google Cloud y descarga su JSON.
3. Comparte el Google Sheet con el `client_email` de la cuenta de servicio con permiso de editor.
4. En Streamlit Cloud, agrega estos Secrets:

```toml
[google_sheets]
spreadsheet_id = "ID_DEL_GOOGLE_SHEET"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "cuenta-servicio@proyecto.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

La app crea automáticamente dos pestañas en la hoja:

```text
ajustes_actuales
ajustes_historial
```

Si Google Sheets no está configurado, el dashboard usa SQLite local como respaldo.
