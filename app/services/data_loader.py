import os
import re
import time
from datetime import datetime, date

import pandas as pd

from app.config import FILE_PATHS, MAX_RELOAD_RETRIES, RELOAD_RETRY_DELAY, READ_RETRY_DELAY, READ_RETRY_ATTEMPTS
from app.services.excel_io import safe_read_excel, safe_get_sheet_names
import app.state as state


def load_all_sheets_data():
    """Load and parse all sheets from all configured Excel files."""
    all_data = {}
    valid_sheet_names = []

    for file_path in FILE_PATHS:
        abs_file_path = os.path.abspath(file_path)

        sheet_names = None
        for attempt in range(READ_RETRY_ATTEMPTS):
            try:
                sheet_names = safe_get_sheet_names(file_path)
                break
            except Exception as e:
                if attempt < READ_RETRY_ATTEMPTS - 1:
                    time.sleep(READ_RETRY_DELAY)
                else:
                    print(f"Error opening file '{file_path}' after {READ_RETRY_ATTEMPTS} attempts: {e}")

        if sheet_names is None:
            continue

        for sheet_name in sheet_names:
            if re.match(r"^sheet\d+$", sheet_name.lower().strip()):
                print(f"Skipping default sheet name: '{sheet_name}'")
                continue

            try:
                df = _read_sheet_with_retry(file_path, sheet_name)

                if df is None or df.empty or len(df.columns) < 2:
                    print(f"Warning: Skipping empty sheet '{sheet_name}' in '{file_path}'")
                    continue

                valid_cols = _get_valid_columns(df)
                if len(valid_cols) < 2:
                    print(f"Warning: Skipping sheet '{sheet_name}' in '{file_path}' - not enough named columns")
                    continue

                all_columns = [str(col) for col in valid_cols]
                df = df[valid_cols]
                df = _trim_to_first_empty_row(df)

                if df.empty:
                    print(f"Warning: Skipping sheet '{sheet_name}' in '{file_path}' - no valid data")
                    continue

                if sheet_name not in all_data:
                    all_data[sheet_name] = {}
                    valid_sheet_names.append(sheet_name)

                _parse_rows(df, sheet_name, abs_file_path, all_columns, all_data)
                print(f"Loaded sheet '{sheet_name}' from '{file_path}'")

            except Exception as e:
                print(f"Error loading sheet '{sheet_name}' from '{file_path}': {e}")
                continue

    if not valid_sheet_names:
        print("Warning: No valid sheets found, creating default")
        all_data["Default"] = {
            "Sample Task": [{
                "details": "No data available\nPlease check your Excel file",
                "metadata": None,
            }]
        }
        valid_sheet_names = ["Default"]

    print(f"Total sheets loaded: {len(valid_sheet_names)}")
    return all_data, valid_sheet_names


def reload_data():
    """Reload data from Excel files and notify connected clients."""
    for attempt in range(MAX_RELOAD_RETRIES):
        try:
            all_sheets_data, sheet_names = load_all_sheets_data()

            has_real_data = _validate_data(all_sheets_data)

            if not has_real_data and state.cached_data["sheet_names"] and "Default" not in state.cached_data["sheet_names"]:
                if attempt < MAX_RELOAD_RETRIES - 1:
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] Got empty data, retrying in {RELOAD_RETRY_DELAY}s... (attempt {attempt + 1}/{MAX_RELOAD_RETRIES})")
                    time.sleep(RELOAD_RETRY_DELAY)
                    continue
                else:
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] Warning: Could not read valid data after {MAX_RELOAD_RETRIES} attempts, keeping previous data")
                    return

            state.cached_data["all_sheets_data"] = all_sheets_data
            state.cached_data["sheet_names"] = sheet_names
            state.cached_data["last_updated"] = datetime.now().isoformat()
            state.data_version += 1

            print(f"[{datetime.now().strftime('%H:%M:%S')}] Data reloaded (version {state.data_version})")

            _notify_clients()
            return

        except Exception as e:
            if attempt < MAX_RELOAD_RETRIES - 1:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Error reloading data: {e}, retrying in {RELOAD_RETRY_DELAY}s... (attempt {attempt + 1}/{MAX_RELOAD_RETRIES})")
                time.sleep(RELOAD_RETRY_DELAY)
            else:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Error reloading data after {MAX_RELOAD_RETRIES} attempts: {e}")


# --- Private helpers ---

def _read_sheet_with_retry(file_path, sheet_name):
    for read_attempt in range(READ_RETRY_ATTEMPTS):
        try:
            return safe_read_excel(file_path, sheet_name=sheet_name)
        except Exception as read_err:
            if read_attempt < READ_RETRY_ATTEMPTS - 1:
                time.sleep(READ_RETRY_DELAY)
            else:
                raise read_err
    return None


def _get_valid_columns(df):
    valid_cols = []
    for col in df.columns:
        col_str = str(col).strip()
        if col_str.startswith("Unnamed") or col_str == "" or col_str == "nan":
            break
        valid_cols.append(col)
    return valid_cols


def _trim_to_first_empty_row(df):
    task_name_col = df.columns[0]
    cut_index = None
    for i, value in enumerate(df[task_name_col]):
        if pd.isna(value) or str(value).strip() == "":
            cut_index = i
            break
    if cut_index is not None:
        df = df.iloc[:cut_index]
    return df


def _parse_rows(df, sheet_name, abs_file_path, all_columns, all_data):
    task_name_col = df.columns[0]
    for row_idx, (_, row) in enumerate(df.iterrows()):
        task_name = str(row[task_name_col])
        if task_name == "nan" or not task_name.strip():
            continue

        details = []
        raw_values = {}
        for col in df.columns:
            value = row[col]
            if pd.isna(value):
                raw_values[str(col)] = None
            elif isinstance(value, (pd.Timestamp, datetime, date)):
                raw_values[str(col)] = value.isoformat()
            else:
                raw_values[str(col)] = value

            if col != task_name_col and pd.notna(value) and str(value).strip():
                details.append(f"{col}: {value}")

        entry = {
            "details": "\n".join(details),
            "metadata": {
                "file_path": abs_file_path,
                "sheet_name": sheet_name,
                "row_index": row_idx,
                "columns": all_columns,
                "raw_values": raw_values,
                "task_name": task_name,
            },
        }

        if task_name not in all_data[sheet_name]:
            all_data[sheet_name][task_name] = []
        all_data[sheet_name][task_name].append(entry)


def _validate_data(all_sheets_data):
    for sheet_name, tasks in all_sheets_data.items():
        if sheet_name != "Default":
            for task_name in tasks:
                if task_name != "Sample Task":
                    return True
    return False


def _notify_clients():
    for client in state.connected_clients:
        client["needs_update"] = True
