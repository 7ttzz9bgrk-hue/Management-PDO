from fastapi import FastAPI

from app.routes.pages import router as pages_router
from app.routes.data import router as data_router
from app.routes.excel import router as excel_router
from app.routes.events import router as events_router
from app.routes.health import router as health_router


def register_routes(app: FastAPI):
    app.include_router(pages_router)
    app.include_router(data_router, prefix="/api")
    app.include_router(excel_router, prefix="/api")
    app.include_router(events_router)
    app.include_router(health_router)
