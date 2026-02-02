from fastapi import FastAPI, Request, HTTPException
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel
from typing import Dict, Any, Optional
import pandas as pd
import asyncio
import os
import re
import shutil
import tempfile
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import time
from openpyxl import load_workbook

app = FastAPI()
templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")

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

# ===== WRITE LOCK (prevents Watchdog reload during saves) =====
write_in_progress = False
write_lock = threading.Lock()

# ===== FILE WATCHER =====
class ExcelFileHandler(FileSystemEventHandler):
    """Watches for changes to Excel files and triggers data reload."""

    def __init__(self, file_paths):
        self.file_paths = [os.path.abspath(fp) for fp in file_paths]
        self.last_reload = 0
        self.debounce_seconds = 2  # Wait 2 seconds before reloading to handle multiple rapid changes

    def on_modified(self, event):
        global write_in_progress

        if event.is_directory:
            return

        # Check if the modified file is one we're watching
        modified_path = os.path.abspath(event.src_path)

        # Only watch .xlsx files
        if not modified_path.lower().endswith('.xlsx'):
            return

        # Excel creates temp files like ~$filename.xlsx, ignore those
        if os.path.basename(modified_path).startswith('~$'):
            return

        # Only react to files explicitly listed in FILE_PATHS
        if modified_path not in self.file_paths:
            return

        # Skip if we're currently writing to prevent reload loops
        with write_lock:
            if write_in_progress:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Ignoring change during write operation")
                return

        current_time = time.time()
        # Debounce: only reload if enough time has passed
        if current_time - self.last_reload > self.debounce_seconds:
            self.last_reload = current_time
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Detected change in: {os.path.basename(modified_path)}")
            reload_data()

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

def safe_read_excel(file_path: str, **kwargs):
    """
    Safely read an Excel file that might be open in Excel.
    Copies to a temp file first to avoid file lock issues.
    """
    # Create a temp file with the same extension
    _, ext = os.path.splitext(file_path)
    temp_fd, temp_path = tempfile.mkstemp(suffix=ext)
    os.close(temp_fd)

    try:
        # Copy the file to temp location (works even if Excel has it open)
        shutil.copy2(file_path, temp_path)
        # Read from the temp copy
        return pd.read_excel(temp_path, **kwargs)
    finally:
        # Clean up temp file
        try:
            os.unlink(temp_path)
        except Exception:
            pass


def safe_get_sheet_names(file_path: str) -> list:
    """
    Safely get sheet names from an Excel file that might be open in Excel.
    """
    _, ext = os.path.splitext(file_path)
    temp_fd, temp_path = tempfile.mkstemp(suffix=ext)
    os.close(temp_fd)

    try:
        shutil.copy2(file_path, temp_path)
        with pd.ExcelFile(temp_path) as excel_file:
            return excel_file.sheet_names
    finally:
        try:
            os.unlink(temp_path)
        except Exception:
            pass


def load_all_sheets_data():
    all_data = {}
    valid_sheet_names = []

    for file_path in FILE_PATHS:
        abs_file_path = os.path.abspath(file_path)  # Store absolute path for metadata
        try:
            sheet_names = safe_get_sheet_names(file_path)
        except Exception as e:
            print(f"Error opening file '{file_path}': {e}")
            continue

        for sheet_name in sheet_names:
            # Skip default Excel sheet names like "Sheet1", "Sheet2", etc.
            if re.match(r'^sheet\d+$', sheet_name.lower().strip()):
                print(f"Skipping default sheet name: '{sheet_name}'")
                continue

            try:
                df = safe_read_excel(file_path, sheet_name=sheet_name)

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

                # Store all column names for metadata
                all_columns = [str(col) for col in valid_cols]

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

                # Use enumerate to track row index for metadata
                for row_idx, (_, row) in enumerate(df.iterrows()):
                    task_name = str(row[task_name_col])

                    if task_name == 'nan' or not task_name.strip():
                        continue

                    # Build details string and raw values dict
                    details = []
                    raw_values = {}
                    for col in df.columns:
                        value = row[col]
                        # Store raw value (None if NaN)
                        raw_values[str(col)] = value if pd.notna(value) else None
                        # Add to details string (skip task name column and empty values)
                        if col != task_name_col and pd.notna(value) and str(value).strip():
                            details.append(f"{col}: {value}")

                    formatted_details = '\n'.join(details)

                    # Create enriched entry with metadata
                    entry = {
                        "details": formatted_details,
                        "metadata": {
                            "file_path": abs_file_path,
                            "sheet_name": sheet_name,
                            "row_index": row_idx,
                            "columns": all_columns,
                            "raw_values": raw_values,
                            "task_name": task_name  # For row validation during save
                        }
                    }

                    # Merge: append to existing task or create new
                    if task_name not in all_data[sheet_name]:
                        all_data[sheet_name][task_name] = []
                    all_data[sheet_name][task_name].append(entry)

                print(f"Loaded sheet '{sheet_name}' from '{file_path}'")

            except Exception as e:
                print(f"Error loading sheet '{sheet_name}' from '{file_path}': {e}")
                continue

    if not valid_sheet_names:
        print("Warning: No valid sheets found, creating default")
        all_data['Default'] = {
            'Sample Task': [{
                "details": "No data available\nPlease check your Excel file",
                "metadata": None
            }]
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


# ===== TASK UPDATE MODEL =====
class TaskUpdate(BaseModel):
    file_path: str
    sheet_name: str
    row_index: int
    task_name: str  # For validation that row hasn't shifted
    updates: Dict[str, Any]  # {column_name: new_value}
    new_columns: Optional[Dict[str, Any]] = None  # New columns to create


@app.post("/api/save-task")
async def save_task(update: TaskUpdate):
    """Save task changes back to the Excel file."""
    global write_in_progress

    try:
        # Validate file exists and is in allowed paths
        abs_path = os.path.abspath(update.file_path)
        allowed_paths = [os.path.abspath(fp) for fp in FILE_PATHS]

        if abs_path not in allowed_paths:
            raise HTTPException(status_code=403, detail="File not in allowed paths")

        if not os.path.exists(abs_path):
            raise HTTPException(status_code=404, detail="File not found")

        # Set write lock to prevent watchdog reload
        with write_lock:
            write_in_progress = True

        try:
            # Read entire Excel file (all sheets) using safe read to handle open files
            excel_data = safe_read_excel(abs_path, sheet_name=None, engine='openpyxl')

            if update.sheet_name not in excel_data:
                raise HTTPException(status_code=404, detail="Sheet not found")

            df = excel_data[update.sheet_name]

            # Validate row index
            if update.row_index < 0 or update.row_index >= len(df):
                raise HTTPException(status_code=400, detail="Invalid row index")

            # Validate task name hasn't shifted
            task_name_col = df.columns[0]
            current_task_name = str(df.iloc[update.row_index][task_name_col])
            if current_task_name != update.task_name:
                raise HTTPException(
                    status_code=409,
                    detail=f"Row position changed. Expected '{update.task_name}' but found '{current_task_name}'. Please refresh and try again."
                )

            # Apply updates to existing columns
            for col, value in update.updates.items():
                if col in df.columns:
                    # Convert empty string to None for proper Excel handling
                    df.at[update.row_index, col] = value if value != '' else None

            # Handle new columns
            if update.new_columns:
                for col, value in update.new_columns.items():
                    if col not in df.columns:
                        # Add new column with None for all rows
                        df[col] = None
                    # Set value for this row
                    df.at[update.row_index, col] = value if value != '' else None

            # Update the sheet in our data structure
            excel_data[update.sheet_name] = df

            # Read original column widths before overwriting (using temp copy to handle open files)
            original_col_widths = {}
            try:
                _, ext = os.path.splitext(abs_path)
                temp_fd, temp_path = tempfile.mkstemp(suffix=ext)
                os.close(temp_fd)
                shutil.copy2(abs_path, temp_path)
                temp_wb = load_workbook(temp_path)
                for sheet_name in temp_wb.sheetnames:
                    original_col_widths[sheet_name] = {}
                    ws = temp_wb[sheet_name]
                    for col_letter, dim in ws.column_dimensions.items():
                        if dim.width:
                            original_col_widths[sheet_name][col_letter] = dim.width
                temp_wb.close()
                os.unlink(temp_path)
            except Exception:
                pass  # If we can't read widths, continue without them

            # Write back to Excel file
            try:
                with pd.ExcelWriter(abs_path, engine='openpyxl', mode='w') as writer:
                    for sname, sheet_df in excel_data.items():
                        sheet_df.to_excel(writer, sheet_name=sname, index=False)

                    # Restore original column widths
                    for sname in writer.sheets:
                        ws = writer.sheets[sname]
                        if sname in original_col_widths:
                            for col_letter, width in original_col_widths[sname].items():
                                ws.column_dimensions[col_letter].width = width
            except PermissionError:
                raise HTTPException(
                    status_code=423,
                    detail="Cannot save: The Excel file is open in another program. Please close Excel and try again."
                )

            print(f"[{datetime.now().strftime('%H:%M:%S')}] Saved changes to {os.path.basename(abs_path)}, sheet '{update.sheet_name}', row {update.row_index}")

            # Small delay to let file system settle
            time.sleep(0.5)

        finally:
            # Release write lock
            with write_lock:
                write_in_progress = False

        # Trigger manual reload after write
        reload_data()

        return {"status": "success", "message": "Task updated successfully"}

    except HTTPException:
        raise
    except Exception as e:
        with write_lock:
            write_in_progress = False
        raise HTTPException(status_code=500, detail=f"Error saving task: {str(e)}")


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
    uvicorn.run("main:app", host="127.0.0.1", port=8889, reload=False)
