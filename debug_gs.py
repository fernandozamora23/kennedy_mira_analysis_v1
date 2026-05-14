import sys
sys.path.append(".")
from app import _init_google_sheets_storage

try:
    success = _init_google_sheets_storage()
    print("Success?", success)
except Exception as e:
    import traceback
    traceback.print_exc()
