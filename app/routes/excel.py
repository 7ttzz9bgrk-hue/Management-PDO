import os

from fastapi import APIRouter, HTTPException

from app.config import FILE_PATHS, ALLOWED_EXCEL_EXTENSIONS
from app.models import ExcelFileRequest
from app.services.excel_manager import open_excel_file
from app.services.path_guard import is_allowed_path, normalize_path

router = APIRouter()


@router.post("/open-excel")
async def open_excel(request: ExcelFileRequest):
    """Open an Excel file with the system default application."""
    abs_path = normalize_path(request.file_path)

    if not is_allowed_path(abs_path, FILE_PATHS):
        raise HTTPException(status_code=403, detail="File not in allowed paths")

    _, ext = os.path.splitext(abs_path)
    if ext.lower() not in ALLOWED_EXCEL_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Only .xlsx and .xlsm files are supported")

    if not os.path.isfile(abs_path):
        raise HTTPException(status_code=404, detail="File not found")

    try:
        return open_excel_file(abs_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to open file: {str(e)}")
