from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

import app.state as state

templates = Jinja2Templates(directory="templates")

router = APIRouter()


@router.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "all_sheets_data": state.cached_data["all_sheets_data"],
            "sheet_names": state.cached_data["sheet_names"],
            "data_version": state.data_version,
        },
    )
