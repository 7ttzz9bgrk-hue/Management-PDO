import os
import shutil
import subprocess
import sys
from datetime import datetime


def open_excel_file(abs_path: str) -> dict:
    """Open an Excel file with the system default application."""
    if sys.platform == "win32":
        command = ["cmd", "/c", "start", "", abs_path]
    elif sys.platform == "darwin":
        command = ["open", abs_path]
    else:
        if shutil.which("xdg-open") is None:
            raise RuntimeError("xdg-open is not installed on this system")
        command = ["xdg-open", abs_path]

    subprocess.Popen(command)

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Opened Excel file: {os.path.basename(abs_path)}")
    return {"status": "opened", "file_path": abs_path}
