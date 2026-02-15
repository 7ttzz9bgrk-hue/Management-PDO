import os
import subprocess
import sys
from datetime import datetime

from app.state import excel_processes


def open_excel_file(abs_path: str) -> dict:
    """Open an Excel file with the system default application."""
    if abs_path in excel_processes:
        proc = excel_processes[abs_path]
        if proc.poll() is None:
            return {"status": "already_open", "file_path": abs_path}
        else:
            del excel_processes[abs_path]

    if sys.platform == "win32":
        proc = subprocess.Popen(["cmd", "/c", "start", "", abs_path])
    elif sys.platform == "darwin":
        proc = subprocess.Popen(["open", abs_path])
    else:
        proc = subprocess.Popen(["xdg-open", abs_path])

    excel_processes[abs_path] = proc
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Opened Excel file: {os.path.basename(abs_path)}")
    return {"status": "opened", "file_path": abs_path}


def close_excel_file(abs_path: str) -> dict:
    """Close a previously opened Excel file."""
    proc = excel_processes.get(abs_path)

    if not proc:
        return {"status": "not_tracked", "message": "No tracked process for this file"}

    if proc.poll() is not None:
        del excel_processes[abs_path]
        return {"status": "already_closed"}

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
        excel_processes.pop(abs_path, None)
        return {"status": "closed", "message": f"Process cleanup attempted: {str(e)}"}

    del excel_processes[abs_path]
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Closed Excel file: {os.path.basename(abs_path)}")
    return {"status": "closed"}


def get_all_status() -> dict:
    """Return the open/close status of all tracked Excel files."""
    status = {}
    to_remove = []
    for path, proc in excel_processes.items():
        if proc.poll() is None:
            status[path] = "open"
        else:
            to_remove.append(path)
    for path in to_remove:
        del excel_processes[path]
    return {"processes": status}
