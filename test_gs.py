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
    sheet = client.open_by_key(secrets["google_sheets"]["spreadsheet_id"])
    print("Success! Title:", sheet.title)
except Exception as e:
    import traceback
    traceback.print_exc()
