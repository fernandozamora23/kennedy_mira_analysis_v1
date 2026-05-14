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

id1 = "1OuQHS0L__n7_W7B7uHcZn-H9EdEfa3zL_aruZePhpZQ"
id2 = "1OuQHS0L__n7_W7B7uHcZn-H9EdEfa3zL_aruZePHpZQ"

print("Testing ID 1 (PhpZQ):")
try:
    sheet = client.open_by_key(id1)
    print("Success ID 1!", sheet.title)
except Exception as e:
    print("Failed ID 1:", e)

print("Testing ID 2 (PHpZQ):")
try:
    sheet = client.open_by_key(id2)
    print("Success ID 2!", sheet.title)
except Exception as e:
    print("Failed ID 2:", e)
