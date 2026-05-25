import subprocess
import sys
import os
import webbrowser
import time
import threading

# ── Fix base directory for both script and frozen EXE ──
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(BASE_DIR)

IS_FROZEN = getattr(sys, 'frozen', False)

def install_requirements():
    """Only runs when NOT frozen (i.e. running as .py script)"""
    if IS_FROZEN:
        return  # EXE already has everything bundled — skip pip entirely

    required = ["fastapi", "uvicorn", "qrcode", "Pillow", "openpyxl", "python-multipart"]
    missing = []
    for pkg in required:
        import_name = pkg.replace("-", "_").lower()
        if import_name == "pillow":
            import_name = "PIL"
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pkg)

    if missing:
        print(f"Installing: {', '.join(missing)} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
        print("Packages installed successfully")
    else:
        print("All packages ready")

def open_browser():
    time.sleep(3)
    webbrowser.open("http://localhost:8000")

def run_server():
    print("\n" + "="*50)
    print("  FUSION RESTAURANT POS SYSTEM")
    print("="*50)
    print(f"Running from: {BASE_DIR}")
    print("Starting server...")
    print("URL:     http://localhost:8000")
    print("Excel:   ./exports/orders_export.xlsx")
    print("Photos:  ./static/uploads/")
    print("="*50)
    print("Browser opens in 3 seconds...\n")

    os.makedirs(os.path.join(BASE_DIR, "exports"), exist_ok=True)
    os.makedirs(os.path.join(BASE_DIR, "receipts"), exist_ok=True)
    os.makedirs(os.path.join(BASE_DIR, "static", "uploads"), exist_ok=True)

    threading.Thread(target=open_browser, daemon=True).start()

    if IS_FROZEN:
        import uvicorn
        uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=False)
    else:
        subprocess.run([
            sys.executable, "-m", "uvicorn", "main:app",
            "--host", "0.0.0.0", "--port", "8000", "--reload"
        ])

if __name__ == "__main__":
    install_requirements()
    run_server()