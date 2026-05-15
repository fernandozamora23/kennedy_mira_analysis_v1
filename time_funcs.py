import time
import sys

class Profiler:
    def __init__(self):
        self.t0 = time.time()
        self.last = self.t0
    def tick(self, name):
        t = time.time()
        print(f"{name}: {t - self.last:.3f}s")
        self.last = t

p = Profiler()
import warnings
warnings.filterwarnings('ignore')

with open("app.py", "r") as f:
    code = f.read()

# We can just trace the execution or run AppTest.
from streamlit.testing.v1 import AppTest
at = AppTest.from_file("app.py", default_timeout=120)
p.tick("AppTest init")
at.run()
p.tick("AppTest run")
