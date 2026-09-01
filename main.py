import sys
import os

# Add src to python import path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

def check_dependencies():
    missing = []
    try:
        import docx
    except ImportError:
        missing.append("python-docx")

    try:
        import PIL
    except ImportError:
        missing.append("Pillow")

    if missing:
        print(f"[!] Warning: The following packages are missing: {', '.join(missing)}")
        print("[!] Please run: pip install -r requirements.txt")

def main():
    check_dependencies()
    from src.gui import main as run_gui
    run_gui()

if __name__ == "__main__":
    main()
