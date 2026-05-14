import toml
import gspread
from google.oauth2.service_account import Credentials

secrets = toml.load(".streamlit/secrets.toml")
creds_info = secrets["gcp_service_account"]
creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")

scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
credentials = Credentials.from_service_account_info(creds_info, scopes=scopes)
client = gspread.authorize(credentials)

try:
    spreadsheet = client.open_by_key(secrets["google_sheets"]["spreadsheet_id"])
    print("Opened spreadsheet:", spreadsheet.title)
    
    title = "ajustes_actuales"
    columns = ["entidad", "entidad_id", "nombre_entidad", "templo_actual", "usuario", "motivo", "actualizado_en"]
    
    try:
        worksheet = spreadsheet.worksheet(title)
        print("Worksheet exists!")
    except Exception as e:
        print("Worksheet doesn't exist, creating...", type(e))
        worksheet = spreadsheet.add_worksheet(title=title, rows=1000, cols=max(len(columns), 8))
        print("Created!")
        
    values = worksheet.get_all_values()
    print("Values count:", len(values))
except Exception as e:
    import traceback
    traceback.print_exc()
