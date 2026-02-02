from fastapi import FastAPI, Request
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse, StreamingResponse
import pandas as pd
import asyncio
import os
import re
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import time

app = FastAPI()
templates = Jinja2Templates(directory="templates")

# ===== CONFIGURATION =====
FILE_PATHS = [
    "mockData.xlsx",
    # Add more file paths here:
    # "//server/shared/projects.xlsx",
    # "another.xlsx",
]

# ===== GLOBAL CACHE =====
cached_data = {
    "all_sheets_data": {},
    "sheet_names": [],
    "last_updated": None
}
data_version = 0  # Increments when data changes
connected_clients = []  # SSE clients waiting for updates

# ===== FILE WATCHER =====
class ExcelFileHandler(FileSystemEventHandler):
    """Watches for changes to Excel files and triggers data reload."""

    def __init__(self, file_paths):
        self.file_paths = [os.path.abspath(fp) for fp in file_paths]
        self.last_reload = 0
        self.debounce_seconds = 1  # Wait 1 second before reloading to handle multiple rapid changes

    def on_modified(self, event):
        if event.is_directory:
            return

        # Check if the modified file is one we're watching
        modified_path = os.path.abspath(event.src_path)

        # Excel creates temp files like ~$filename.xlsx, ignore those
        if os.path.basename(modified_path).startswith('~$'):
            return

        for watched_path in self.file_paths:
            if modified_path == watched_path:
                current_time = time.time()
                # Debounce: only reload if enough time has passed
                if current_time - self.last_reload > self.debounce_seconds:
                    self.last_reload = current_time
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] Detected change in: {os.path.basename(modified_path)}")
                    reload_data()
                break

def start_file_watcher():
    """Start watching all Excel file directories for changes."""
    global observer

    event_handler = ExcelFileHandler(FILE_PATHS)
    observer = Observer()

    # Watch the directories containing the Excel files
    watched_dirs = set()
    for file_path in FILE_PATHS:
        abs_path = os.path.abspath(file_path)
        dir_path = os.path.dirname(abs_path) or '.'
        if dir_path not in watched_dirs:
            watched_dirs.add(dir_path)
            observer.schedule(event_handler, dir_path, recursive=False)
            print(f"Watching directory: {dir_path}")

    observer.start()
    print("File watcher started - will auto-refresh when Excel files change")

def reload_data():
    """Reload data from Excel files and notify connected clients."""
    global data_version

    all_sheets_data, sheet_names = load_all_sheets_data()
    cached_data["all_sheets_data"] = all_sheets_data
    cached_data["sheet_names"] = sheet_names
    cached_data["last_updated"] = datetime.now().isoformat()
    data_version += 1

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Data reloaded (version {data_version})")

    # Notify all connected SSE clients
    notify_clients()

def load_all_sheets_data():
    all_data = {}
    valid_sheet_names = []
    
    for file_path in FILE_PATHS:
        try:
            with pd.ExcelFile(file_path) as excel_file:
                sheet_names = excel_file.sheet_names
        except Exception as e:
            print(f"Error opening file '{file_path}': {e}")
            continue
        
        for sheet_name in sheet_names:
            # Skip default Excel sheet names like "Sheet1", "Sheet2", etc.
            if re.match(r'^sheet\d+$', sheet_name.lower().strip()):
                print(f"Skipping default sheet name: '{sheet_name}'")
                continue

            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)

                if df.empty or len(df.columns) < 2:
                    print(f"Warning: Skipping empty sheet '{sheet_name}' in '{file_path}'")
                    continue
                
                # Rule 1: Stop at the first unnamed/undefined column
                valid_cols = []
                for col in df.columns:
                    col_str = str(col).strip()
                    if col_str.startswith('Unnamed') or col_str == '' or col_str == 'nan':
                        break
                    valid_cols.append(col)
                
                if len(valid_cols) < 2:
                    print(f"Warning: Skipping sheet '{sheet_name}' in '{file_path}' - not enough named columns")
                    continue
                
                df = df[valid_cols]
                
                # Rule 2: Stop at the first empty Task Name (first column) row
                task_name_col = df.columns[0]
                cut_index = None
                for i, value in enumerate(df[task_name_col]):
                    if pd.isna(value) or str(value).strip() == '':
                        cut_index = i
                        break
                
                if cut_index is not None:
                    df = df.iloc[:cut_index]
                
                if df.empty:
                    print(f"Warning: Skipping sheet '{sheet_name}' in '{file_path}' - no valid data")
                    continue
                
                # Initialize sheet if it doesn't exist
                if sheet_name not in all_data:
                    all_data[sheet_name] = {}
                    valid_sheet_names.append(sheet_name)
                
                for _, row in df.iterrows():
                    task_name = str(row[task_name_col])
                    
                    if task_name == 'nan' or not task_name.strip():
                        continue
                    
                    details = []
                    for col in df.columns[1:]:
                        value = row[col]
                        if pd.notna(value) and str(value).strip():
                            details.append(f"{col}: {value}")
                    
                    formatted_details = '\n'.join(details)
                    
                    # Merge: append to existing task or create new
                    if task_name not in all_data[sheet_name]:
                        all_data[sheet_name][task_name] = []
                    all_data[sheet_name][task_name].append(formatted_details)
                
                print(f"Loaded sheet '{sheet_name}' from '{file_path}'")
            
            except Exception as e:
                print(f"Error loading sheet '{sheet_name}' from '{file_path}': {e}")
                continue
    
    if not valid_sheet_names:
        print("Warning: No valid sheets found, creating default")
        all_data['Default'] = {
            'Sample Task': ['No data available\nPlease check your Excel file']
        }
        valid_sheet_names = ['Default']
    
    print(f"Total sheets loaded: {len(valid_sheet_names)}")
    return all_data, valid_sheet_names


# ===== SSE (Server-Sent Events) for real-time updates =====
def notify_clients():
    """Send update notification to all connected SSE clients."""
    global connected_clients
    # Mark all clients as needing an update
    for client in connected_clients:
        client["needs_update"] = True


async def event_generator():
    """Generate SSE events for a connected client."""
    global data_version

    client = {"needs_update": False, "last_version": data_version}
    connected_clients.append(client)

    try:
        while True:
            # Check if there's a new version
            if client["needs_update"] or client["last_version"] != data_version:
                client["needs_update"] = False
                client["last_version"] = data_version
                yield f"data: {data_version}\n\n"

            await asyncio.sleep(0.5)  # Check every 500ms
    finally:
        connected_clients.remove(client)


@app.get("/events")
async def sse_events():
    """SSE endpoint for real-time data update notifications."""
    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )


@app.get("/api/data")
async def get_data():
    """API endpoint to fetch the latest cached data."""
    return {
        "all_sheets_data": cached_data["all_sheets_data"],
        "sheet_names": cached_data["sheet_names"],
        "version": data_version,
        "last_updated": cached_data["last_updated"]
    }


@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    # Use cached data instead of loading fresh
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "all_sheets_data": cached_data["all_sheets_data"],
            "sheet_names": cached_data["sheet_names"],
            "data_version": data_version
        }
    )


@app.on_event("startup")
async def startup_event():
    """Load initial data and start file watcher on app startup."""
    # Load initial data
    reload_data()
    # Start file watcher in background thread
    threading.Thread(target=start_file_watcher, daemon=True).start()


@app.on_event("shutdown")
async def shutdown_event():
    """Stop file watcher on shutdown."""
    try:
        observer.stop()
        observer.join()
    except Exception:
        pass


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8889, reload=True)
