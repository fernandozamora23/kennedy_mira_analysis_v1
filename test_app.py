import sys
from streamlit.testing.v1 import AppTest

def test_app():
    at = AppTest.from_file("app.py")
    try:
        at.run(timeout=100)
        if at.exception:
            print("EXCEPTION FOUND:")
            print(at.exception[0])
            sys.exit(1)
        else:
            print("APP RAN SUCCESSFULLY")
    except Exception as e:
        print("RUNTIME ERROR:")
        print(e)
        sys.exit(1)

if __name__ == "__main__":
    test_app()
