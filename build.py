import subprocess
import sys
import os

def build_exe():
    # Ensure PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Build the EXE
    script_path = "v1.3.3 Rindex.py"
    
    cmd = [
        "pyinstaller",
        "--onefile",  # Single EXE file
        "--windowed",  # No console window
        "--add-data", "quotation_template.docx;.",
        "--add-data", "rindex.png;.",   # Include the template in the EXE bundle
        "--hidden-import=tkcalendar",  # Include tkcalendar dependencies
        "--hidden-import=docx",  # Include python-docx dependencies
        "--hidden-import=docx.enum.text",
        "--hidden-import=docx.shared",
        "--hidden-import=docx.table",
        "--name", "Rindex-PiGen",
        script_path
    ]
    
    subprocess.run(cmd, check=True)
    print("EXE built successfully")

if __name__ == "__main__":
    build_exe()