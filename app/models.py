from pydantic import BaseModel
from typing import Dict, Any, Optional


class TaskUpdate(BaseModel):
    file_path: str
    sheet_name: str
    row_index: int
    task_name: str
    updates: Dict[str, Any]
    new_columns: Optional[Dict[str, Any]] = None


class ExcelFileRequest(BaseModel):
    file_path: str
