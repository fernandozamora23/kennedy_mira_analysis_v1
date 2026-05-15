import time
from streamlit.testing.v1 import AppTest
t0 = time.time()
at = AppTest.from_file("app.py")
print("Initialized in:", time.time() - t0)
t0 = time.time()
at.run(timeout=100)
print("Run completed in:", time.time() - t0)
