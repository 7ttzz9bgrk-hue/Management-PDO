import logging
from pathlib import Path

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
import threading

from app.routes import register_routes
from app.services.data_loader import reload_data
from app.services.file_watcher import start_file_watcher

BASE_DIR = Path(__file__).resolve().parent.parent
logger = logging.getLogger(__name__)


def create_app() -> FastAPI:
    application = FastAPI(title="Management PDO")

    application.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")

    register_routes(application)

    @application.on_event("startup")
    async def startup_event():
        reload_data()
        threading.Thread(target=start_file_watcher, daemon=True).start()

    @application.on_event("shutdown")
    async def shutdown_event():
        from app.services.file_watcher import observer

        try:
            if observer is not None:
                observer.stop()
                observer.join(timeout=2)
        except Exception as exc:
            logger.warning("Failed to stop file watcher cleanly: %s", exc)

    return application
