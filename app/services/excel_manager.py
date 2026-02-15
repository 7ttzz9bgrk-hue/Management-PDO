import os
import subprocess
import sys
from datetime import datetime

from app.state import excel_processes, excel_open_files


def open_excel_file(abs_path: str, original_path: str) -> dict:
    """Open an Excel file with the system default application."""
    if original_path in excel_open_files:
        return {"status": "already_open", "file_path": original_path}

    if sys.platform == "win32":
        proc = subprocess.Popen(["cmd", "/c", "start", "", abs_path])
    elif sys.platform == "darwin":
        proc = subprocess.Popen(["open", abs_path])
    else:
        proc = subprocess.Popen(["xdg-open", abs_path])

    excel_processes[abs_path] = proc
    excel_open_files[original_path] = abs_path
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Opened Excel file: {os.path.basename(abs_path)}")
    return {"status": "opened", "file_path": original_path}


def close_excel_file(abs_path: str, original_path: str) -> dict:
    """Close a previously opened Excel file."""
    proc = excel_processes.pop(abs_path, None)

    if proc and proc.poll() is None:
        try:
            if sys.platform == "win32":
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                    capture_output=True, timeout=5,
                )
            else:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()
        except Exception as e:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Process cleanup note: {str(e)}")

    removed = excel_open_files.pop(original_path, None)
    if removed:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Closed Excel file: {os.path.basename(abs_path)}")
        return {"status": "closed"}
    return {"status": "not_tracked", "message": "No tracked state for this file"}


def get_all_status() -> dict:
    """Return the open/close status of all tracked Excel files."""
    return {"processes": {path: "open" for path in excel_open_files}}
