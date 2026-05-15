from streamlit.testing.v1 import AppTest
import time

print("Starting AppTest...")
t0 = time.time()
at = AppTest.from_file("app.py", default_timeout=60)
at.run()
t1 = time.time()
print(f"First run took: {t1 - t0:.2f}s")

t0 = time.time()
at.run()
t1 = time.time()
print(f"Second run (cached) took: {t1 - t0:.2f}s")
