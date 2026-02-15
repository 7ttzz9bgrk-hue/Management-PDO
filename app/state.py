import threading

cached_data = {
    "all_sheets_data": {},
    "sheet_names": [],
    "last_updated": None,
}

data_version = 0

connected_clients = []

write_in_progress = False
write_lock = threading.Lock()

excel_processes = {}  # abs_path -> subprocess.Popen (for close attempt)
excel_open_files = {}  # original_path -> abs_path (state tracking)
