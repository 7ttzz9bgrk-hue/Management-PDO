import os

from fastapi import APIRouter, HTTPException

from app.config import FILE_PATHS
from app.models import ExcelFileRequest
from app.services.excel_manager import open_excel_file

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
        return open_excel_file(abs_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to open file: {str(e)}")
