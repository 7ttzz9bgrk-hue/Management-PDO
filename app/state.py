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

