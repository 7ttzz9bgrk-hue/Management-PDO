import os
import time
from datetime import datetime

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from app.config import FILE_PATHS, DEBOUNCE_SECONDS
import app.state as state

observer = None


class ExcelFileHandler(FileSystemEventHandler):
    """Watches for changes to Excel files and triggers data reload."""

    def __init__(self, file_paths):
        self.file_paths = [os.path.abspath(fp) for fp in file_paths]
        self.last_reload = 0

    def on_modified(self, event):
        if event.is_directory:
            return

        modified_path = os.path.abspath(event.src_path)

        if not modified_path.lower().endswith(".xlsx"):
            return

        if os.path.basename(modified_path).startswith("~$"):
            return

        if modified_path not in self.file_paths:
            return

        with state.write_lock:
            if state.write_in_progress:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Ignoring change during write operation")
                return

        current_time = time.time()
        if current_time - self.last_reload > DEBOUNCE_SECONDS:
            self.last_reload = current_time
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Detected change in: {os.path.basename(modified_path)}")
            time.sleep(0.5)

            from app.services.data_loader import reload_data
            reload_data()


def start_file_watcher():
    """Start watching all Excel file directories for changes."""
    global observer

    event_handler = ExcelFileHandler(FILE_PATHS)
    observer = Observer()

    watched_dirs = set()
    for file_path in FILE_PATHS:
        abs_path = os.path.abspath(file_path)
        dir_path = os.path.dirname(abs_path) or "."
        if dir_path not in watched_dirs:
            watched_dirs.add(dir_path)
            observer.schedule(event_handler, dir_path, recursive=False)
            print(f"Watching directory: {dir_path}")

    observer.start()
    print("File watcher started - will auto-refresh when Excel files change")
