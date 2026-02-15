import os

from fastapi import APIRouter, HTTPException

from app.config import FILE_PATHS
from app.models import ExcelFileRequest
from app.services.excel_manager import open_excel_file, close_excel_file, get_all_status

router = APIRouter()


@router.post("/open-excel")
async def open_excel(request: ExcelFileRequest):
    """Open an Excel file with the system default application."""
    abs_path = os.path.abspath(request.file_path)
    allowed_paths = [os.path.abspath(fp) for fp in FILE_PATHS]

    if abs_path not in allowed_paths:
        raise HTTPException(status_code=403, detail="File not in allowed paths")

    if not os.path.exists(abs_path):
        raise HTTPException(status_code=404, detail="File not found")

    try:
        return open_excel_file(abs_path, request.file_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to open file: {str(e)}")


@router.post("/close-excel")
async def close_excel(request: ExcelFileRequest):
    """Close a previously opened Excel file."""
    abs_path = os.path.abspath(request.file_path)
    return close_excel_file(abs_path, request.file_path)


@router.get("/excel-status")
async def excel_status():
    """Return the current open/close status of all tracked Excel files."""
    return get_all_status()
