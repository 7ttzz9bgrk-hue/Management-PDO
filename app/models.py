from pydantic import BaseModel, Field
from typing import Dict, Any, Optional


class TaskUpdate(BaseModel):
    file_path: str = Field(min_length=1)
    sheet_name: str = Field(min_length=1)
    row_index: int = Field(ge=0)
    task_name: str = Field(min_length=1)
    updates: Dict[str, Any]
    new_columns: Optional[Dict[str, Any]] = None


class ExcelFileRequest(BaseModel):
    file_path: str = Field(min_length=1)
