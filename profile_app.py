import cProfile
import pstats
from streamlit.testing.v1 import AppTest

def run():
    at = AppTest.from_file("app.py")
    at.run(timeout=100)

cProfile.run("run()", "stats.prof")
p = pstats.Stats("stats.prof")
p.sort_stats("cumulative").print_stats(30)
