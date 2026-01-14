import subprocess
import sys
import os
import webbrowser

def main():
    app_path = os.path.join(os.path.dirname(__file__), "app.py")
    subprocess.Popen([
        sys.executable,
        "-m",
        "streamlit",
        "run",
        app_path,
        "--server.headless=true",
        "--browser.serverAddress=localhost"
    ])
    # webbrowser.open("http://localhost:8501")

if __name__ == "__main__":
    main()


