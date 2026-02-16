import io
import sys

import pandas as pd


def read_file_with_shared_access(file_path: str) -> bytes:
    """Read a file even if it's open in another program (like Excel)."""
    if sys.platform == "win32":
        import ctypes
        from ctypes import wintypes

        GENERIC_READ = 0x80000000
        FILE_SHARE_READ = 0x1
        FILE_SHARE_WRITE = 0x2
        FILE_SHARE_DELETE = 0x4
        OPEN_EXISTING = 3
        FILE_ATTRIBUTE_NORMAL = 0x80

        CreateFileW = ctypes.windll.kernel32.CreateFileW
        CreateFileW.argtypes = [
            wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
            wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE,
        ]
        CreateFileW.restype = wintypes.HANDLE

        ReadFile = ctypes.windll.kernel32.ReadFile
        GetFileSize = ctypes.windll.kernel32.GetFileSize
        CloseHandle = ctypes.windll.kernel32.CloseHandle

        handle = CreateFileW(
            file_path,
            GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            None,
        )

        if handle == -1:
            raise IOError(f"Cannot open file: {file_path}")

        try:
            file_size = GetFileSize(handle, None)
            buffer = ctypes.create_string_buffer(file_size)
            bytes_read = wintypes.DWORD()
            ReadFile(handle, buffer, file_size, ctypes.byref(bytes_read), None)
            return buffer.raw[: bytes_read.value]
        finally:
            CloseHandle(handle)
    else:
        with open(file_path, "rb") as f:
            return f.read()


def safe_read_excel(file_path: str, **kwargs) -> pd.DataFrame:
    """Safely read an Excel file that might be open in Excel."""
    file_bytes = read_file_with_shared_access(file_path)
    return pd.read_excel(io.BytesIO(file_bytes), **kwargs)


def safe_get_sheet_names(file_path: str) -> list:
    """Safely get sheet names from an Excel file that might be open."""
    file_bytes = read_file_with_shared_access(file_path)
    with pd.ExcelFile(io.BytesIO(file_bytes)) as excel_file:
        return excel_file.sheet_names
