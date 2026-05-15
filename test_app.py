from streamlit.testing.v1 import AppTest
import traceback
import sys

print("Running AppTest...")
try:
    at = AppTest.from_file("app.py", default_timeout=60)
    at.run()
    if at.exception:
        print("EXCEPTION DETECTED IN STREAMLIT:")
        for e in at.exception:
            print(traceback.format_exception(type(e), e, e.__traceback__))
        sys.exit(1)
    else:
        print("App ran successfully with no exceptions!")
except Exception as e:
    print("FATAL APPTEST ERROR:")
    traceback.print_exc()
    sys.exit(1)
