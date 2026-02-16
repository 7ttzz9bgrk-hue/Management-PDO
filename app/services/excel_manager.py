import os
import subprocess
import sys
from datetime import datetime


def open_excel_file(abs_path: str) -> dict:
    """Open an Excel file with the system default application."""
    if sys.platform == "win32":
        subprocess.Popen(["cmd", "/c", "start", "", abs_path])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", abs_path])
    else:
        subprocess.Popen(["xdg-open", abs_path])

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Opened Excel file: {os.path.basename(abs_path)}")
    return {"status": "opened", "file_path": abs_path}
